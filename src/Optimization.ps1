    #region Optimization

    function Invoke-ImageOptimization {
    [CmdletBinding(SupportsShouldProcess)]
        param(
            [System.Collections.Generic.List[ImageUsage]]$usages,
            [string]$tempDir,
            [hashtable]$morphPairs
        )
        
        Write-Host "`n[PROC] Optimizing images..." -ForegroundColor Yellow
        
        # Group by physical image
        $imageGroups = @{}
        foreach ($usage in $usages) {
            if (-not $imageGroups.ContainsKey($usage.ImagePhysicalPath)) {
                $group = [ImageGroup]::new()
                $group.PhysicalPath = $usage.ImagePhysicalPath
                $group.IsSvgFallbackImage = $usage.IsSvgFallbackUsage
                # Read file size from disk - will be cropped file if crop was applied
                if (Test-Path -LiteralPath $usage.ImagePhysicalPath) {
                    $group.OriginalSizeBytes = (Get-Item -LiteralPath $usage.ImagePhysicalPath).Length
                } else {
                    $group.OriginalSizeBytes = $usage.BeforeSizeBytes
                }
                $imageGroups[$usage.ImagePhysicalPath] = $group
            }
            $imageGroups[$usage.ImagePhysicalPath].Usages.Add($usage)
        }
        
        $optimizedCount = 0
        $skippedCount = 0
        
        # Process groups in slide order (by minimum slide number of usages),
        # then by media filename as a tiebreaker for deterministic output.
        $sortedGroups = $imageGroups.Values | Sort-Object @(
            @{
                Expression = {
                    $slideUsages = $_.Usages | Where-Object { $_.ContextType -eq [ContextType]::Slide }
                    if ($slideUsages) {
                        ($slideUsages | Measure-Object -Property SlideNumber -Minimum).Minimum
                    } else {
                        9999  # No slide usages (master/layout only), sort to end
                    }
                }
            },
            @{
                Expression = { [System.IO.Path]::GetFileName($_.PhysicalPath) }
            }
        )
        
        # Determine if slide filters are active (used for scoping)
        $slideFiltersActive = ($IncludeSlides -and $IncludeSlides.Count -gt 0) -or ($ExcludeSlides -and $ExcludeSlides.Count -gt 0)
        
        # Pre-filter groups to only those in scope for accurate progress count
        $inScopeGroups = @($sortedGroups | Where-Object {
            $grp = $_
            $hasIncluded = $grp.Usages | Where-Object {
                if ($slideFiltersActive) {
                    $_.ContextType -eq [ContextType]::Slide -and (Test-SlideIncluded -SlideNumber $_.SlideNumber)
                } else {
                    $_.ContextType -ne [ContextType]::Slide -or (Test-SlideIncluded -SlideNumber $_.SlideNumber)
                }
            }
            $hasIncluded
        })
        
        $totalGroups = $inScopeGroups.Count
        $currentGroup = 0

        # Build a flat list of usages from in-scope groups and run the parallel probe
        # pre-pass. Probing only in-scope groups preserves work scoping.
        $inScopeUsages = [System.Collections.Generic.List[ImageUsage]]::new()
        foreach ($g in $inScopeGroups) {
            foreach ($u in $g.Usages) { $inScopeUsages.Add($u) }
        }
        Update-OptimizeProbesParallel -usages $inScopeUsages

        $scratchDir = Join-Path $tempDir '.opt-scratch'
        New-Item -ItemType Directory -Path $scratchDir -Force | Out-Null

        foreach ($group in $sortedGroups) {
            $mediaFile = Split-Path $group.PhysicalPath -Leaf
            
            # Check slide filter FIRST - skip groups where no usage is on an included slide
            # When slide filters are active, also skip masters/layouts (they affect all slides)
            $hasIncludedUsage = $group.Usages | Where-Object {
                if ($slideFiltersActive) {
                    # Only include slide usages that pass the filter
                    $_.ContextType -eq [ContextType]::Slide -and (Test-SlideIncluded -SlideNumber $_.SlideNumber)
                } else {
                    # No filter - include all (slides pass through, masters/layouts always included)
                    $_.ContextType -ne [ContextType]::Slide -or (Test-SlideIncluded -SlideNumber $_.SlideNumber)
                }
            }
            if (-not $hasIncludedUsage) {
                continue
            }
            
            # Now update progress (group is in scope)
            $currentGroup++
            Write-Progress -Activity "Optimizing images" -Status "$mediaFile ($currentGroup of $totalGroups)" `
                -PercentComplete (($currentGroup / [Math]::Max(1, $totalGroups)) * 100) -Id 3
            
            # Check if optimization forbidden (only after confirming group is in scope)
            $skipReason = Get-OptimizationSkipReason -group $group
            if ($skipReason) {
                foreach ($usage in $group.Usages) {
                    if ($usage.OptimizationStatus -eq 'Pending') {
                        $usage.OptimizationStatus = $skipReason.Status
                        $usage.WhyNotOptimized = $skipReason.Reason
                        $usage.ManualActionRequired = $skipReason.ManualAction
                        $usage.ManualActionHint = $skipReason.Hint
                    }
                }
                # Log first included usage as representative
                $firstUsage = $hasIncludedUsage | Select-Object -First 1
                Write-Host "   [SKIP] Skipped optimization for '$($firstUsage.ShapeName)' on $($firstUsage.Location): $($skipReason.Reason)" -ForegroundColor Gray
                $skippedCount += $group.Usages.Count
                continue
            }
            
            # Calculate target dimensions
            $maxDisplay = @{
                Width = ($group.Usages | Measure-Object -Property DisplayWidthPx -Maximum).Maximum
                Height = ($group.Usages | Measure-Object -Property DisplayHeightPx -Maximum).Maximum
            }
            $firstUsage = $hasIncludedUsage | Select-Object -First 1

            if ($PSCmdlet.ShouldProcess($mediaFile, "Optimize image")) {
                # Probe values were already cached by Update-OptimizeProbesParallel above.
                # Build a job or set skip statuses.
                $job = Get-OptimizationJob -group $group -maxDisplay $maxDisplay -scratchDir $scratchDir
                if (-not $job) {
                    if ($group.Usages[0].OptimizationStatus -match '^(Skipped|NoChange)') { $skippedCount += $group.Usages.Count }
                    continue
                }

                # Phase 3 (sequential here): run magick to the scratch file.
                & $script:MagickExe @($job.MagickArgs) 2>&1 | Out-Null
                $success = ($LASTEXITCODE -eq 0 -and (Test-Path -LiteralPath $job.ScratchPath))
                $afterSize = if ($success) { (Get-Item -LiteralPath $job.ScratchPath).Length } else { 0 }

                # Phase 4: commit (sequential, slide order).
                $result = Complete-OptimizationJob -job $job -group $group -scratchPath $job.ScratchPath -success $success -afterSize $afterSize -tempDir $tempDir
                if ($result.Success) { $optimizedCount++ } else { $skippedCount += $group.Usages.Count }
            }
        }

        if (Test-Path -LiteralPath $scratchDir) { Remove-Item -LiteralPath $scratchDir -Recurse -Force -ErrorAction SilentlyContinue }

        Write-Progress -Activity "Optimizing images" -Completed -Id 3
        Write-Host "[STATS] Optimization phase complete: $optimizedCount optimized, $skippedCount skipped" -ForegroundColor Green
    }

    function Get-OptimizationSkipReason {
        param([ImageGroup]$group)
        
        # Check file format - skip vector formats and other non-optimizable types
        $ext = [System.IO.Path]::GetExtension($group.PhysicalPath).ToLower()
        $supportedFormats = $script:Config.SUPPORTED_RASTER_FORMATS
        $convertibleFormats = $script:Config.CONVERTIBLE_FORMATS
        
        # Check for animated GIF (must skip - cropping/resizing breaks animation)
        if ($ext -eq '.gif' -and (Test-IsAnimatedGif -ImagePath $group.PhysicalPath)) {
            return @{
                Status = 'Skipped_AnimatedGif'
                Reason = 'Animated GIF - would break animation'
                ManualAction = $false
                Hint = ''
            }
        }
        
        if ($ext -notin $supportedFormats -and $ext -notin $convertibleFormats) {
            $fileName = Split-Path $group.PhysicalPath -Leaf
            $formatType = switch ($ext) {
                '.emf' { 'vector graphic (EMF)' }
                '.wmf' { 'vector graphic (WMF)' }
                '.svg' { 'vector graphic (SVG)' }
                default { "format ($ext)" }
            }
            
            return @{
                Status = 'Skipped_UnsupportedFormat'
                Reason = "Unsupported $formatType"
                ManualAction = $false
                Hint = ''
            }
        }
        
        # Source dimensions are front-loaded in parallel before optimization. Probe here
        # only as a fallback for any usage whose parallel probe failed (still 0).
        if ($group.Usages | Where-Object { $_.SourceWidthPx -eq 0 }) {
            try {
                $dims = Get-ImageDimensions -ImagePath $group.PhysicalPath

                foreach ($usage in $group.Usages) {
                    if ($usage.SourceWidthPx -eq 0) {
                        $usage.SourceWidthPx = $dims.Width
                        $usage.SourceHeightPx = $dims.Height
                    }
                }
            } catch {
                # Silently continue if dimension capture fails
            }
        }
        
        # Check SVG fallback
        if ($group.IsSvgFallbackImage) {
            return @{
                Status = 'Skipped_SvgFallback'
                Reason = 'PNG is auto-generated SVG fallback; PowerPoint regenerates'
                ManualAction = $false
                Hint = ''
            }
        }
        
        # Check for any already-set skip status (crop-related or Morph conflicts)
        $alreadySkipped = $group.Usages | Where-Object { 
            $_.OptimizationStatus -like 'Skipped_*' 
        } | Select-Object -First 1
        
        if ($alreadySkipped) {
            # Return the existing skip reason to prevent optimization and log it properly
            return @{
                Status = $alreadySkipped.OptimizationStatus
                Reason = $alreadySkipped.WhyNotOptimized
                ManualAction = $alreadySkipped.ManualActionRequired
                Hint = $alreadySkipped.ManualActionHint
            }
        }
        
        # Check if optimization is enabled for this context
        $hasSlide = $group.Usages | Where-Object { $_.ContextType -eq [ContextType]::Slide }
        $hasTemplate = $group.Usages | Where-Object { $_.ContextType -in @([ContextType]::Master, [ContextType]::Layout) }
        
        $contextTypes = ($group.Usages | Select-Object -ExpandProperty ContextType -Unique) -join ', '
        $fileName = Split-Path $group.PhysicalPath -Leaf
        Write-Verbose "  Checking skip reasons for $($fileName): contexts=[$contextTypes], hasSlide=$($null -ne $hasSlide), hasTemplate=$($null -ne $hasTemplate), OptimizeSlides=$script:OptimizeSlides, OptimizeMastersAndLayouts=$script:OptimizeMastersAndLayouts"
        
        if ($hasSlide -and -not $script:OptimizeSlides) {
            return @{
                Status = 'Skipped_SlidesFlagMissing'
                Reason = 'Used in slide; enable -OptimizeSlides'
                ManualAction = $true
                Hint = 'Run with -OptimizeSlides'
            }
        }
        
        if ($hasTemplate -and -not $script:OptimizeMastersAndLayouts) {
            return @{
                Status = 'Skipped_TemplatesFlagMissing'
                Reason = 'Used in master/layout; enable -OptimizeMastersAndLayouts'
                ManualAction = $true
                Hint = 'Run with -OptimizeMastersAndLayouts'
            }
        }
        
        # Check unresolved crops - determine appropriate hint based on context
        $hasUnresolvedCrop = @($group.Usages | Where-Object { $_.HasSrcRect -and -not $_.CropApplied })
        if ($hasUnresolvedCrop.Count -gt 0) {
            $firstUnresolved = $hasUnresolvedCrop[0]
            $hint = if ($firstUnresolved.ContextType -in @([ContextType]::Master, [ContextType]::Layout)) {
                'Run with -CropMastersAndLayouts'
            } else {
                'Run with -CropSlides'
            }
            return @{
                Status = 'Skipped_CropNotMaterialized'
                Reason = 'Crop present; enable cropping'
                ManualAction = $true
                Hint = $hint
            }
        }
        
        return $null
    }

    # Probe each unique media file's transparency percent (for .png and convertible formats)
    # or JPEG quality (for .jpg/.jpeg) in parallel, then apply cached values to all usages
    # of that file. Runspaces invoke only magick and return raw strings; parsing is sequential
    # in the parent. Uses $script:MagickExe. Tolerates an empty usage list (parallel block is
    # skipped). Progress bar uses Id 5 ("Analyzing images").
    function Update-OptimizeProbesParallel {
        param(
            [System.Collections.Generic.List[ImageUsage]]$usages
        )

        $magickExe = $script:MagickExe
        $convertibleFormats = $script:Config.CONVERTIBLE_FORMATS

        # Group by physical path and filter to only files that need probing.
        $uniqueGroups = @($usages | Group-Object -Property ImagePhysicalPath | Where-Object {
            $ext = [System.IO.Path]::GetExtension($_.Name).ToLower()
            $ext -in @('.png', '.jpg', '.jpeg') -or $ext -in $convertibleFormats
        })

        $rawResults = @()
        if ($uniqueGroups.Count -gt 0) {
            # Each runspace invokes only magick and returns a raw PSCustomObject.
            # The downstream ForEach-Object runs in the parent runspace, increments a
            # parent-local counter, and renders progress. No shared counter is needed.
            $total = $uniqueGroups.Count
            $rawResults = @($uniqueGroups | ForEach-Object {
                [PSCustomObject]@{ Path = $_.Name; Ext = [System.IO.Path]::GetExtension($_.Name).ToLower() }
            }) | ForEach-Object -Parallel {
                $path = $_.Path
                $ext  = $_.Ext
                if (-not (Test-Path -LiteralPath $path)) {
                    [PSCustomObject]@{ Path = $path; Kind = 'none'; Raw = $null }
                } elseif ($ext -eq '.png' -or $ext -in $using:convertibleFormats) {
                    $out = & $using:magickExe $path -alpha extract -format "%[fx:100*(1-mean)]" info: 2>&1
                    if ($LASTEXITCODE -ne 0) {
                        [PSCustomObject]@{ Path = $path; Kind = 'transparency'; Raw = $null }
                    } else {
                        $token = @($out) | Where-Object { $_ -notmatch '^(WARNING|System\.Management)' -and $_ -match '^\d+\.?\d*$' } | Select-Object -First 1
                        [PSCustomObject]@{ Path = $path; Kind = 'transparency'; Raw = $token }
                    }
                } elseif ($ext -in @('.jpg', '.jpeg')) {
                    $out = & $using:magickExe identify -format "%Q" $path 2>&1
                    if ($LASTEXITCODE -ne 0) {
                        [PSCustomObject]@{ Path = $path; Kind = 'quality'; Raw = $null }
                    } else {
                        if ($out -is [array]) { $out = $out[0] }
                        [PSCustomObject]@{ Path = $path; Kind = 'quality'; Raw = "$out" }
                    }
                } else {
                    [PSCustomObject]@{ Path = $path; Kind = 'none'; Raw = $null }
                }
            } -ThrottleLimit ([Environment]::ProcessorCount) | ForEach-Object -Begin { $count = 0 } -Process {
                $count++
                Write-Progress -Activity "Analyzing images" -Status "$count of $total" `
                    -PercentComplete (($count / $total) * 100) -Id 5
                $_
            }
            Write-Progress -Activity "Analyzing images" -Completed -Id 5
        }

        # Parse results sequentially and apply to all usages sharing each path.
        $probeMap = @{}
        foreach ($r in $rawResults) {
            if ($r.Kind -eq 'transparency') {
                $raw = $r.Raw
                if ($raw -and $raw -match '^\d+\.?\d*$') {
                    $probeMap[$r.Path] = [PSCustomObject]@{ Kind = 'transparency'; Value = [double]$raw }
                } else {
                    $probeMap[$r.Path] = [PSCustomObject]@{ Kind = 'transparency'; Value = 0.0 }
                }
            } elseif ($r.Kind -eq 'quality') {
                $raw = $r.Raw
                if ($raw -and $raw -match '^\d+$') {
                    $probeMap[$r.Path] = [PSCustomObject]@{ Kind = 'quality'; Value = [int]$raw }
                } else {
                    $probeMap[$r.Path] = [PSCustomObject]@{ Kind = 'quality'; Value = -1 }
                }
            }
        }

        foreach ($usage in $usages) {
            $p = $usage.ImagePhysicalPath
            if (-not $probeMap.ContainsKey($p)) { continue }
            $probe = $probeMap[$p]
            if ($probe.Kind -eq 'transparency') {
                $usage.EffectiveTransparencyPercent = $probe.Value
            } elseif ($probe.Kind -eq 'quality') {
                $usage.SourceJpegQuality = $probe.Value
            }
        }
    }

    function Get-ImageTransparency {
        param([string]$imagePath)
        
        $fileName = Split-Path $imagePath -Leaf
        
        # Use magick (not convert) to check if image has actual transparency
        # Returns percentage of pixels that are NOT fully opaque
        $result = & $script:MagickExe $imagePath -alpha extract -format "%[fx:100*(1-mean)]" info: 2>&1 | 
            Where-Object { $_ -notmatch '^(WARNING|System\.Management)' -and $_ -match '^\d+\.?\d*$' } |
            Select-Object -First 1
        
        if ($result -and $result -match '^\d+\.?\d*$') {
            $transp = [double]$result
            
            if (Test-VerboseMode) {
                $decision = if ($transp -le $TransparencyThresholdPercent) { "opaque (can convert to JPEG)" } else { "transparent (keep as PNG)" }
                Write-Verbose "  Transparency analysis for $($fileName): $($transp.ToString('F2'))% transparent pixels -> $decision"
            }
            
            return $transp
        }
        
        if (Test-VerboseMode) {
            Write-Verbose "  Transparency analysis for $($fileName): unable to determine, assuming opaque (0%)"
        }
        
        return 0.0
    }

    function Test-MinSavingsThreshold {
        param(
            [long]$beforeSize,
            [long]$afterSize
        )
        
        # Always reject if file grew larger (negative savings)
        if ($afterSize -ge $beforeSize) {
            return $false
        }
        
        if ($MinSavingsPercent -le 0) {
            return $true  # No minimum threshold (but still reject growth)
        }
        
        $savingsPercent = (($beforeSize - $afterSize) / $beforeSize) * 100
        return $savingsPercent -ge $MinSavingsPercent
    }

    function Get-OptimizationJob {
        param(
            [ImageGroup]$group,
            [hashtable]$maxDisplay,
            [string]$scratchDir
        )

        # Read cached source dimensions from first usage that has them, or first usage as fallback.
        $knownUsage = $group.Usages | Where-Object { $_.SourceWidthPx -gt 0 } | Select-Object -First 1
        if (-not $knownUsage) {
            $knownUsage = $group.Usages[0]
        }
        $srcWidth  = $knownUsage.SourceWidthPx
        $srcHeight = $knownUsage.SourceHeightPx

        # Compute desired display dimensions with headroom, clamped to source dims.
        $desiredWidth  = [Math]::Ceiling($maxDisplay.Width  * $HeadroomFactor)
        $desiredHeight = [Math]::Ceiling($maxDisplay.Height * $HeadroomFactor)

        # Background/fill images have no display size; keep source dimensions.
        if ($desiredWidth -eq 0 -or $desiredHeight -eq 0) {
            $desiredWidth  = $srcWidth
            $desiredHeight = $srcHeight
        }

        $targetWidth  = [Math]::Min($srcWidth,  $desiredWidth)
        $targetHeight = [Math]::Min($srcHeight, $desiredHeight)

        # Write source and target dims onto every usage so the CSV report has accurate values.
        # This mirrors the placement in Invoke-OptimizationOperation and runs regardless of which
        # branch (skip, convert, resize) follows.
        foreach ($usage in $group.Usages) {
            $usage.SourceWidthPx  = $srcWidth
            $usage.SourceHeightPx = $srcHeight
            $usage.TargetWidthPx  = $targetWidth
            $usage.TargetHeightPx = $targetHeight
        }

        $ext = [System.IO.Path]::GetExtension($group.PhysicalPath).ToLower()

        # JPEG: check if already at optimal quality and size.
        if ($ext -in @('.jpg', '.jpeg')) {
            $sourceQuality = $knownUsage.SourceJpegQuality
            if ($sourceQuality -gt 0 -and $sourceQuality -le $JpegQuality -and
                $srcWidth -eq $targetWidth -and $srcHeight -eq $targetHeight) {
                foreach ($usage in $group.Usages) {
                    $usage.OptimizationStatus  = 'Skipped_AlreadyOptimal'
                    $usage.WhyNotOptimized     = "Already at optimal quality (Q$sourceQuality <= Q$JpegQuality) and target size"
                    $usage.ManualActionRequired = $false
                }
                return $null
            }

            # Build ResizeJpeg job.
            $outExt     = $ext
            $scratchPath = Join-Path $scratchDir ("$([guid]::NewGuid())$outExt")
            $magickArgs = @(
                $group.PhysicalPath,
                '-filter', $script:Config.RESIZE_FILTER,
                '-resize', "${targetWidth}x${targetHeight}>",
                '-strip',
                '-quality', $JpegQuality,
                '-define', 'jpeg:optimize-coding=true',
                '-define', 'jpeg:dct-method=float',
                '-sampling-factor', $script:Config.JPEG_SAMPLING_FACTOR_DEFAULT,
                $scratchPath
            )
            $job = [OptimizeJob]::new()
            $job.GroupKey    = $group.PhysicalPath
            $job.SourcePath  = $group.PhysicalPath
            $job.ScratchPath = $scratchPath
            $job.MagickArgs  = $magickArgs
            $job.Operation   = 'OptimizeJpeg'
            $job.StatusName  = 'OptimizedJpeg'
            $job.BeforeSize  = $group.OriginalSizeBytes
            $job.NewExtension = ''
            return $job
        }

        # Convertible formats (BMP, TIFF, GIF): read cached transparency.
        if ($ext -in $script:Config.CONVERTIBLE_FORMATS) {
            $transp = $knownUsage.EffectiveTransparencyPercent
            if ($transp -gt $TransparencyThresholdPercent) {
                $outExt = '.png'
            } else {
                $outExt = '.jpeg'
            }

            $scratchPath = Join-Path $scratchDir ("$([guid]::NewGuid())$outExt")
            $magickArgs = @(
                $group.PhysicalPath,
                '-filter', $script:Config.RESIZE_FILTER,
                '-resize', "${targetWidth}x${targetHeight}>",
                '-strip'
            )
            if ($outExt -eq '.jpeg') {
                $magickArgs += @(
                    '-quality', $JpegQuality,
                    '-define', 'jpeg:optimize-coding=true',
                    '-define', 'jpeg:dct-method=float',
                    '-sampling-factor', $script:Config.JPEG_SAMPLING_FACTOR_DEFAULT
                )
                $statusName = 'ConvertedToJpeg'
            } else {
                $magickArgs += @(
                    '-define', "png:compression-level=$($script:Config.PNG_COMPRESSION_LEVEL)",
                    '-define', 'png:color-type=6'
                )
                $statusName = 'ConvertedToPng'
            }
            $magickArgs += $scratchPath

            $job = [OptimizeJob]::new()
            $job.GroupKey     = $group.PhysicalPath
            $job.SourcePath   = $group.PhysicalPath
            $job.ScratchPath  = $scratchPath
            $job.MagickArgs   = $magickArgs
            $job.Operation    = $statusName
            $job.StatusName   = $statusName
            $job.BeforeSize   = $group.OriginalSizeBytes
            $job.NewExtension = $outExt
            return $job
        }

        # Already at target size.
        if ($srcWidth -eq $targetWidth -and $srcHeight -eq $targetHeight) {
            if ($ext -eq '.png') {
                $transp = $knownUsage.EffectiveTransparencyPercent
                if ($transp -le $TransparencyThresholdPercent) {
                    # Opaque PNG at target size: convert to JPEG without resize.
                    $outExt      = '.jpeg'
                    $scratchPath = Join-Path $scratchDir ("$([guid]::NewGuid())$outExt")
                    $magickArgs  = @(
                        $group.PhysicalPath,
                        '-strip',
                        '-quality', $JpegQuality,
                        '-define', 'jpeg:optimize-coding=true',
                        '-define', 'jpeg:dct-method=float',
                        '-sampling-factor', $script:Config.JPEG_SAMPLING_FACTOR_DEFAULT,
                        $scratchPath
                    )
                    $job = [OptimizeJob]::new()
                    $job.GroupKey     = $group.PhysicalPath
                    $job.SourcePath   = $group.PhysicalPath
                    $job.ScratchPath  = $scratchPath
                    $job.MagickArgs   = $magickArgs
                    $job.Operation    = 'ConvertPngToJpeg'
                    $job.StatusName   = 'ConvertedPngToJpeg'
                    $job.BeforeSize   = $group.OriginalSizeBytes
                    $job.NewExtension = '.jpeg'
                    return $job
                }
            }

            # Already at target size and no format conversion needed.
            foreach ($usage in $group.Usages) {
                $usage.OptimizationStatus  = 'NoChangeNeeded'
                $usage.WhyNotOptimized     = 'Already at target size'
                $usage.ManualActionRequired = $false
            }
            return $null
        }

        # Needs resize.
        if ($ext -in @('.jpg', '.jpeg')) {
            # Already handled above; this branch is unreachable but kept for completeness.
            $outExt      = $ext
            $scratchPath = Join-Path $scratchDir ("$([guid]::NewGuid())$outExt")
            $magickArgs  = @(
                $group.PhysicalPath,
                '-filter', $script:Config.RESIZE_FILTER,
                '-resize', "${targetWidth}x${targetHeight}>",
                '-strip',
                '-quality', $JpegQuality,
                '-define', 'jpeg:optimize-coding=true',
                '-define', 'jpeg:dct-method=float',
                '-sampling-factor', $script:Config.JPEG_SAMPLING_FACTOR_DEFAULT,
                $scratchPath
            )
            $job = [OptimizeJob]::new()
            $job.GroupKey     = $group.PhysicalPath
            $job.SourcePath   = $group.PhysicalPath
            $job.ScratchPath  = $scratchPath
            $job.MagickArgs   = $magickArgs
            $job.Operation    = 'OptimizeJpeg'
            $job.StatusName   = 'OptimizedJpeg'
            $job.BeforeSize   = $group.OriginalSizeBytes
            $job.NewExtension = ''
            return $job
        } elseif ($ext -eq '.png') {
            $transp = $knownUsage.EffectiveTransparencyPercent
            if ($transp -le $TransparencyThresholdPercent) {
                # Opaque PNG needing resize: convert to JPEG with resize.
                $outExt      = '.jpeg'
                $scratchPath = Join-Path $scratchDir ("$([guid]::NewGuid())$outExt")
                $magickArgs  = @(
                    $group.PhysicalPath,
                    '-filter', $script:Config.RESIZE_FILTER,
                    '-resize', "${targetWidth}x${targetHeight}>",
                    '-strip',
                    '-quality', $JpegQuality,
                    '-define', 'jpeg:optimize-coding=true',
                    '-define', 'jpeg:dct-method=float',
                    '-sampling-factor', $script:Config.JPEG_SAMPLING_FACTOR_DEFAULT,
                    $scratchPath
                )
                $job = [OptimizeJob]::new()
                $job.GroupKey     = $group.PhysicalPath
                $job.SourcePath   = $group.PhysicalPath
                $job.ScratchPath  = $scratchPath
                $job.MagickArgs   = $magickArgs
                $job.Operation    = 'ConvertPngToJpeg'
                $job.StatusName   = 'ConvertedPngToJpeg'
                $job.BeforeSize   = $group.OriginalSizeBytes
                $job.NewExtension = '.jpeg'
                return $job
            } else {
                # Transparent PNG needing resize: optimize PNG in place.
                $outExt      = '.png'
                $scratchPath = Join-Path $scratchDir ("$([guid]::NewGuid())$outExt")
                $magickArgs  = @(
                    $group.PhysicalPath,
                    '-filter', $script:Config.RESIZE_FILTER,
                    '-resize', "${targetWidth}x${targetHeight}>",
                    '-strip',
                    '-define', "png:compression-level=$($script:Config.PNG_COMPRESSION_LEVEL)",
                    '-define', 'png:color-type=6',
                    $scratchPath
                )
                $job = [OptimizeJob]::new()
                $job.GroupKey     = $group.PhysicalPath
                $job.SourcePath   = $group.PhysicalPath
                $job.ScratchPath  = $scratchPath
                $job.MagickArgs   = $magickArgs
                $job.Operation    = 'OptimizePngAlpha'
                $job.StatusName   = 'OptimizedPngAlpha'
                $job.BeforeSize   = $group.OriginalSizeBytes
                $job.NewExtension = ''
                return $job
            }
        } else {
            # Unsupported format (should have been caught by Get-OptimizationSkipReason).
            foreach ($usage in $group.Usages) {
                $usage.OptimizationStatus  = 'Skipped_Other'
                $usage.WhyNotOptimized     = 'Unsupported format'
                $usage.ManualActionRequired = $false
            }
            return $null
        }
    }

    function Complete-OptimizationJob {
        param(
            [OptimizeJob]$job,
            [ImageGroup]$group,
            [string]$scratchPath,
            [bool]$success,
            [long]$afterSize,
            [string]$tempDir
        )

        # Magick failed: mark all usages and return.
        if (-not $success) {
            foreach ($usage in $group.Usages) {
                $usage.OptimizationStatus = 'Skipped_Other'
                $usage.WhyNotOptimized = 'Optimization failed'
                $usage.ManualActionRequired = $false
            }
            return @{ Success = $false }
        }

        $beforeSize = $job.BeforeSize

        # Savings check: reject if file did not shrink or did not clear the minimum threshold.
        if ($afterSize -ge $beforeSize -or -not (Test-MinSavingsThreshold -beforeSize $beforeSize -afterSize $afterSize)) {
            if (Test-Path -LiteralPath $scratchPath) {
                Remove-Item -LiteralPath $scratchPath -Force
            }
            $isConvert = $job.NewExtension -ne ''
            $reason = if ($afterSize -ge $beforeSize) {
                if ($isConvert) { 'Conversion yielded no savings' } else { 'No net savings achieved' }
            } else {
                $actualSavings = (($beforeSize - $afterSize) / $beforeSize) * 100
                "Savings $($actualSavings.ToString('F1'))% below $($MinSavingsPercent.ToString('F1'))% threshold"
            }
            foreach ($usage in $group.Usages) {
                $usage.OptimizationStatus = 'Skipped_NoNetSavings'
                $usage.WhyNotOptimized = $reason
                $usage.ManualActionRequired = $false
            }
            return @{ Success = $false }
        }

        $savedPercent = (($beforeSize - $afterSize) / $beforeSize) * 100

        if ($job.NewExtension -eq '') {
            # In-place operation: overwrite the original with the scratch file.
            Move-Item $scratchPath $group.PhysicalPath -Force
            $group.OptimizedSizeBytes = $afterSize
            foreach ($usage in $group.Usages) {
                $usage.AfterSizeBytes = $afterSize
                $usage.OptimizationStatus = $job.StatusName
            }
            $firstUsage = $group.Usages[0]
            $okVerb = switch ($job.Operation) {
                'OptimizeJpeg'    { 'Optimized JPEG' }
                'OptimizePngAlpha' { 'Optimized PNG with transparency' }
                default            { 'Optimized' }
            }
            Write-Host "   [OK] $okVerb for '$($firstUsage.ShapeName)' on $($firstUsage.Location) (saved $($savedPercent.ToString('F1'))%)" -ForegroundColor Green
        } else {
            # Convert operation: allocate a new media name, update rels, remove the original.
            $srcExtForMsg = [System.IO.Path]::GetExtension($group.PhysicalPath).ToLower()
            $out = New-UniqueMediaPath -basePath $group.PhysicalPath -suffix '' -newExtension $job.NewExtension
            Move-Item $scratchPath $out.Path -Force
            foreach ($usage in $group.Usages) {
                $null = Update-BlipRelationship -usage $usage -tempDir $tempDir -newMediaFileName $out.Name
                $usage.AfterSizeBytes = $afterSize
                $usage.OptimizationStatus = $job.StatusName
                $usage.OptimizedFile = $out.Name
            }
            Remove-Item -LiteralPath $group.PhysicalPath -Force
            $group.OptimizedSizeBytes = $afterSize
            $group.OptimizedPath = $out.Path
            $firstUsage = $group.Usages[0]
            $convVerb = if ($job.Operation -eq 'ConvertPngToJpeg') {
                'Converted PNG to JPEG'
            } else {
                "Converted $srcExtForMsg to $($job.NewExtension)"
            }
            Write-Host "   [CONV] $convVerb for '$($firstUsage.ShapeName)' on $($firstUsage.Location) (saved $($savedPercent.ToString('F1'))%)" -ForegroundColor Green
        }

        return @{ Success = $true }
    }

    #endregion
