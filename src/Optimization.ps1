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
                Write-Verbose "Optimizing $mediaFile [display: $($maxDisplay.Width)x$($maxDisplay.Height)px] from $($firstUsage.Location)"
                
                $result = Invoke-OptimizationOperation -group $group -maxDisplay $maxDisplay -tempDir $tempDir
                if ($result.Success) {
                    $optimizedCount++
                } else {
                    $skippedCount += $group.Usages.Count
                }
            }
        }
        
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

    function Invoke-OptimizationOperation {
        param(
            [ImageGroup]$group,
            [hashtable]$maxDisplay,
            [string]$tempDir
        )
        
        try {
            # Source dimensions are front-loaded in parallel; read the cached value and
            # fall back to a live probe only if it is still missing.
            $known = $group.Usages | Where-Object { $_.SourceWidthPx -gt 0 } | Select-Object -First 1
            if ($known) {
                $srcWidth = $known.SourceWidthPx
                $srcHeight = $known.SourceHeightPx
            } else {
                $dims = Get-ImageDimensions -ImagePath $group.PhysicalPath
                $srcWidth = $dims.Width
                $srcHeight = $dims.Height
            }
            
            # Calculate target size
            $desiredWidth = [Math]::Ceiling($maxDisplay.Width * $HeadroomFactor)
            $desiredHeight = [Math]::Ceiling($maxDisplay.Height * $HeadroomFactor)
            
            # For background/fill images with no display size, use source dimensions
            if ($desiredWidth -eq 0 -or $desiredHeight -eq 0) {
                $desiredWidth = $srcWidth
                $desiredHeight = $srcHeight
            }
            
            $targetWidth = [Math]::Min($srcWidth, $desiredWidth)
            $targetHeight = [Math]::Min($srcHeight, $desiredHeight)
            
            # Update usages
            foreach ($usage in $group.Usages) {
                $usage.SourceWidthPx = $srcWidth
                $usage.SourceHeightPx = $srcHeight
                $usage.TargetWidthPx = $targetWidth
                $usage.TargetHeightPx = $targetHeight
            }
            
            Write-Verbose "  Dimensions: source ${srcWidth}x${srcHeight}px -> target ${targetWidth}x${targetHeight}px [HeadroomFactor: $HeadroomFactor]"
            
            # Check JPEG quality for quality preservation
            $ext = [System.IO.Path]::GetExtension($group.PhysicalPath).ToLower()
            if ($ext -in @('.jpg', '.jpeg')) {
                $sourceQuality = Get-JpegQuality -ImagePath $group.PhysicalPath
                foreach ($usage in $group.Usages) {
                    $usage.SourceJpegQuality = $sourceQuality
                }
                
                # If source quality is at or below target and dimensions match, skip
                if ($sourceQuality -gt 0 -and $sourceQuality -le $JpegQuality -and 
                    $srcWidth -eq $targetWidth -and $srcHeight -eq $targetHeight) {
                    $firstUsage = $group.Usages[0]
                    foreach ($usage in $group.Usages) {
                        $usage.OptimizationStatus = 'Skipped_AlreadyOptimal'
                        $usage.WhyNotOptimized = "Already at optimal quality (Q$sourceQuality <= Q$JpegQuality) and target size"
                        $usage.ManualActionRequired = $false
                    }
                    Write-Host "   [SKIP] Skipped '$($firstUsage.ShapeName)' on $($firstUsage.Location): already at target size and quality (Q$sourceQuality)" -ForegroundColor Gray
                    return @{ Success = $false }
                }
            }
            
            # Handle convertible formats (BMP, TIFF, GIF)
            if ($ext -in $script:Config.CONVERTIBLE_FORMATS) {
                $result = ConvertTo-OptimizedFormat -group $group -targetWidth $targetWidth -targetHeight $targetHeight -tempDir $tempDir
                return $result
            }
            
            # Check if already at target
            if ($srcWidth -eq $targetWidth -and $srcHeight -eq $targetHeight) {
                # Check if PNG can be converted to JPEG
                if ($group.PhysicalPath -match '\.png$') {
                    $transp = Get-ImageTransparency -imagePath $group.PhysicalPath
                    foreach ($usage in $group.Usages) {
                        $usage.EffectiveTransparencyPercent = $transp
                    }
                    
                    if ($transp -le $TransparencyThresholdPercent) {
                        # Opaque - convert PNG to JPEG
                        $firstUsage = $group.Usages[0]
                        Write-Host "   [INFO] Effective transparency $($transp.ToString('F1'))% for '$($firstUsage.ShapeName)' on $($firstUsage.Location) - treating as opaque" -ForegroundColor Cyan
                        $result = ConvertTo-Jpeg -group $group -tempDir $tempDir -resize $false
                        return $result
                    }
                }
                
                $firstUsage = $group.Usages[0]
                foreach ($usage in $group.Usages) {
                    $usage.OptimizationStatus = 'NoChangeNeeded'
                    $usage.WhyNotOptimized = 'Already at target size'
                    $usage.ManualActionRequired = $false
                }
                Write-Host "   [SKIP] No change needed for '$($firstUsage.ShapeName)' on $($firstUsage.Location) (already at target)" -ForegroundColor Gray
                return @{ Success = $false }
            }
            
            # Determine optimization strategy
            if ($ext -eq '.jpg' -or $ext -eq '.jpeg') {
                $result = Optimize-Jpeg -group $group -targetWidth $targetWidth -targetHeight $targetHeight -tempDir $tempDir
            } elseif ($ext -eq '.png') {
                # Always check real transparency for PNGs (even if alpha channel exists)
                $transp = Get-ImageTransparency -imagePath $group.PhysicalPath
                foreach ($usage in $group.Usages) {
                    $usage.EffectiveTransparencyPercent = $transp
                }
                
                if ($transp -le $TransparencyThresholdPercent) {
                    # Opaque or negligible transparency - convert to JPEG
                    $firstUsage = $group.Usages[0]
                    Write-Host "   [INFO] Effective transparency $($transp.ToString('F1'))% for '$($firstUsage.ShapeName)' on $($firstUsage.Location) - treating as opaque" -ForegroundColor Cyan
                    $result = ConvertTo-Jpeg -group $group -tempDir $tempDir -resize $true
                } else {
                    # Has meaningful transparency - keep as PNG
                    $result = Optimize-PngWithAlpha -group $group -targetWidth $targetWidth -targetHeight $targetHeight -tempDir $tempDir
                }
            } else {
                foreach ($usage in $group.Usages) {
                    $usage.OptimizationStatus = 'Skipped_Other'
                    $usage.WhyNotOptimized = 'Unsupported format'
                }
                return @{ Success = $false }
            }
            
            return $result
            
        } catch {
            Write-Warning "Optimization failed for $($group.PhysicalPath): $_"
            foreach ($usage in $group.Usages) {
                $usage.OptimizationStatus = 'Skipped_Other'
                $usage.WhyNotOptimized = "Optimization failed: $_"
            }
            return @{ Success = $false }
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

    function Optimize-Jpeg {
        param(
            [ImageGroup]$group,
            [int]$targetWidth,
            [int]$targetHeight,
            [string]$tempDir
        )
        
        Write-Verbose "  Optimizing JPEG: target ${targetWidth}x${targetHeight}px, quality $JpegQuality"
        
        $outputPath = "$($group.PhysicalPath).tmp"
        
        $magickArgs = @(
            $group.PhysicalPath,
            '-filter', $script:Config.RESIZE_FILTER,
            '-resize', "${targetWidth}x${targetHeight}>",
            '-strip',
            '-quality', $JpegQuality,
            '-define', 'jpeg:optimize-coding=true',
            '-define', 'jpeg:dct-method=float',
            '-sampling-factor', $script:Config.JPEG_SAMPLING_FACTOR_DEFAULT,
            $outputPath
        )
        
        & $script:MagickExe $magickArgs 2>&1 | Out-Null
        
        if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $outputPath)) {
            throw "JPEG optimization failed"
        }
        
        $beforeSize = $group.OriginalSizeBytes
        $afterSize = (Get-Item $outputPath).Length
        
        # Check minimum savings threshold
        if ($afterSize -ge $beforeSize -or -not (Test-MinSavingsThreshold -beforeSize $beforeSize -afterSize $afterSize)) {
            Remove-Item -LiteralPath $outputPath -Force
            $firstUsage = $group.Usages[0]
            
            $reason = if ($afterSize -ge $beforeSize) {
                'No net savings achieved'
            } else {
                $actualSavings = (($beforeSize - $afterSize) / $beforeSize) * 100
                "Savings $($actualSavings.ToString('F1'))% below $($MinSavingsPercent.ToString('F1'))% threshold"
            }
            
            foreach ($usage in $group.Usages) {
                $usage.OptimizationStatus = 'Skipped_NoNetSavings'
                $usage.WhyNotOptimized = $reason
                $usage.ManualActionRequired = $false
            }
            
            if (Test-VerboseMode) {
                Write-Verbose "  JPEG optimization rejected: $(Format-ByteSize $beforeSize) -> $(Format-ByteSize $afterSize)"
            }
            
            Write-Host "   [INFO] $reason for '$($firstUsage.ShapeName)' on $($firstUsage.Location) - kept original" -ForegroundColor Cyan
            return @{ Success = $false }
        }
        
        # Replace original
        Move-Item $outputPath $group.PhysicalPath -Force
        $group.OptimizedSizeBytes = $afterSize
        
        $savedPercent = (($beforeSize - $afterSize) / $beforeSize) * 100
        $firstUsage = $group.Usages[0]
        
        foreach ($usage in $group.Usages) {
            $usage.AfterSizeBytes = $afterSize
            $usage.OptimizationStatus = 'OptimizedJpeg'
        }
        
        Write-Host "   [OK] Optimized JPEG for '$($firstUsage.ShapeName)' on $($firstUsage.Location) (saved $($savedPercent.ToString('F1'))%)" -ForegroundColor Green
        
        return @{ Success = $true }
    }

    function Optimize-PngWithAlpha {
        param(
            [ImageGroup]$group,
            [int]$targetWidth,
            [int]$targetHeight,
            [string]$tempDir
        )
        
        Write-Verbose "  Optimizing PNG with transparency: target ${targetWidth}x${targetHeight}px, compression level $($script:Config.PNG_COMPRESSION_LEVEL)"
        
        $outputPath = "$($group.PhysicalPath).tmp"
        
        $magickArgs = @(
            $group.PhysicalPath,
            '-filter', $script:Config.RESIZE_FILTER,
            '-resize', "${targetWidth}x${targetHeight}>",
            '-strip',
            '-define', "png:compression-level=$($script:Config.PNG_COMPRESSION_LEVEL)",
            '-define', 'png:color-type=6',
            $outputPath
        )
        
        & $script:MagickExe $magickArgs 2>&1 | Out-Null
        
        if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $outputPath)) {
            throw "PNG optimization failed"
        }
        
        $beforeSize = $group.OriginalSizeBytes
        $afterSize = (Get-Item $outputPath).Length
        
        # Check minimum savings threshold
        if ($afterSize -ge $beforeSize -or -not (Test-MinSavingsThreshold -beforeSize $beforeSize -afterSize $afterSize)) {
            Remove-Item -LiteralPath $outputPath -Force
            $firstUsage = $group.Usages[0]
            
            $reason = if ($afterSize -ge $beforeSize) {
                'No net savings achieved'
            } else {
                $actualSavings = (($beforeSize - $afterSize) / $beforeSize) * 100
                "Savings $($actualSavings.ToString('F1'))% below $($MinSavingsPercent.ToString('F1'))% threshold"
            }
            
            foreach ($usage in $group.Usages) {
                $usage.OptimizationStatus = 'Skipped_NoNetSavings'
                $usage.WhyNotOptimized = $reason
                $usage.ManualActionRequired = $false
            }
            
            if (Test-VerboseMode) {
                Write-Verbose "  PNG optimization rejected: $(Format-ByteSize $beforeSize) -> $(Format-ByteSize $afterSize)"
            }
            
            Write-Host "   [INFO] $reason for '$($firstUsage.ShapeName)' on $($firstUsage.Location) - kept original" -ForegroundColor Cyan
            return @{ Success = $false }
        }
        
        # Replace original
        Move-Item $outputPath $group.PhysicalPath -Force
        $group.OptimizedSizeBytes = $afterSize
        
        $savedPercent = (($beforeSize - $afterSize) / $beforeSize) * 100
        $firstUsage = $group.Usages[0]
        
        foreach ($usage in $group.Usages) {
            $usage.AfterSizeBytes = $afterSize
            $usage.OptimizationStatus = 'OptimizedPngAlpha'
        }
        
        Write-Host "   [OK] Optimized PNG with transparency for '$($firstUsage.ShapeName)' on $($firstUsage.Location) (saved $($savedPercent.ToString('F1'))%)" -ForegroundColor Green
        
        return @{ Success = $true }
    }

    function ConvertTo-Jpeg {
        param(
            [ImageGroup]$group,
            [string]$tempDir,
            [bool]$resize
        )
        
        $srcFile = Split-Path $group.PhysicalPath -Leaf
        Write-Verbose "  Converting PNG->JPEG: $srcFile [resize: $resize, quality: $JpegQuality]"
        
        # Generate unique JPEG filename
        $output = New-UniqueMediaPath -basePath $group.PhysicalPath -suffix '' -newExtension '.jpeg'
        $outputPath = $output.Path
        $outputName = $output.Name
        
        $magickArgs = @($group.PhysicalPath)
        
        if ($resize) {
            $usage = $group.Usages[0]
            $magickArgs += @(
                '-filter', $script:Config.RESIZE_FILTER,
                '-resize', "$($usage.TargetWidthPx)x$($usage.TargetHeightPx)>"
            )
        }
        
        $magickArgs += @(
            '-strip',
            '-quality', $JpegQuality,
            '-define', 'jpeg:optimize-coding=true',
            '-define', 'jpeg:dct-method=float',
            '-sampling-factor', $script:Config.JPEG_SAMPLING_FACTOR_DEFAULT,
            $outputPath
        )
        
        & $script:MagickExe $magickArgs 2>&1 | Out-Null
        
        if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $outputPath)) {
            throw "PNG to JPEG conversion failed"
        }
        
        $beforeSize = $group.OriginalSizeBytes
        $afterSize = (Get-Item $outputPath).Length
        
        # Check minimum savings threshold
        if ($afterSize -ge $beforeSize -or -not (Test-MinSavingsThreshold -beforeSize $beforeSize -afterSize $afterSize)) {
            Remove-Item -LiteralPath $outputPath -Force
            $firstUsage = $group.Usages[0]
            
            $reason = if ($afterSize -ge $beforeSize) {
                'Conversion yielded no savings'
            } else {
                $actualSavings = (($beforeSize - $afterSize) / $beforeSize) * 100
                "Savings $($actualSavings.ToString('F1'))% below $($MinSavingsPercent.ToString('F1'))% threshold"
            }
            
            foreach ($usage in $group.Usages) {
                $usage.OptimizationStatus = 'Skipped_NoNetSavings'
                $usage.WhyNotOptimized = $reason
                $usage.ManualActionRequired = $false
            }
            
            if (Test-VerboseMode) {
                Write-Verbose "  PNG->JPEG conversion rejected: $(Format-ByteSize $beforeSize) -> $(Format-ByteSize $afterSize)"
            }
            
            Write-Host "   [INFO] $reason for '$($firstUsage.ShapeName)' on $($firstUsage.Location) - kept original" -ForegroundColor Cyan
            return @{ Success = $false }
        }
        
        # Update all relationships pointing to this image
        Write-Verbose "  Updating $($group.Usages.Count) relationship(s) to point to new JPEG: $outputName"
        
        foreach ($usage in $group.Usages) {
            $null = Update-BlipRelationship -usage $usage -tempDir $tempDir -newMediaFileName $outputName
            
            $usage.AfterSizeBytes = $afterSize
            $usage.OptimizationStatus = 'ConvertedPngToJpeg'
            $usage.OptimizedFile = $outputName
        }
        
        # Remove original PNG
        Remove-Item -LiteralPath $group.PhysicalPath -Force
        $group.OptimizedSizeBytes = $afterSize
        $group.OptimizedPath = $outputPath
        
        $savedPercent = (($beforeSize - $afterSize) / $beforeSize) * 100
        $firstUsage = $group.Usages[0]
        Write-Host "   [CONV] Converted PNG to JPEG for '$($firstUsage.ShapeName)' on $($firstUsage.Location) (saved $($savedPercent.ToString('F1'))%)" -ForegroundColor Green
        
        return @{ Success = $true }
    }

    function ConvertTo-OptimizedFormat {
        param(
            [ImageGroup]$group,
            [int]$targetWidth,
            [int]$targetHeight,
            [string]$tempDir
        )
        
        $srcFile = Split-Path $group.PhysicalPath -Leaf
        $ext = [System.IO.Path]::GetExtension($group.PhysicalPath).ToLower()
        
        # Check if image has transparency (for TIFF and GIF)
        $hasTransparency = $false
        if ($ext -in @('.tif', '.tiff', '.gif')) {
            $transp = Get-ImageTransparency -imagePath $group.PhysicalPath
            $hasTransparency = $transp -gt $TransparencyThresholdPercent
            foreach ($usage in $group.Usages) {
                $usage.EffectiveTransparencyPercent = $transp
            }
            
            if ($hasTransparency) {
                $firstUsage = $group.Usages[0]
                Write-Host "   [INFO] Effective transparency $($transp.ToString('F1').Replace('.', ','))% for '$($firstUsage.ShapeName)' on $($firstUsage.Location) - keeping as PNG" -ForegroundColor Cyan
            } else {
                $firstUsage = $group.Usages[0]
                Write-Host "   [INFO] Effective transparency $($transp.ToString('F1').Replace('.', ','))% for '$($firstUsage.ShapeName)' on $($firstUsage.Location) - treating as opaque" -ForegroundColor Cyan
            }
        }
        
        # Determine output format
        $outputExt = if ($hasTransparency) { '.png' } else { '.jpeg' }
        
        Write-Verbose "  Converting $ext->$($outputExt): $srcFile [target: ${targetWidth}x${targetHeight}px]"
        
        # Generate unique filename
        $output = New-UniqueMediaPath -basePath $group.PhysicalPath -suffix '' -newExtension $outputExt
        $outputPath = $output.Path
        $outputName = $output.Name
        
        $magickArgs = @(
            $group.PhysicalPath,
            '-filter', $script:Config.RESIZE_FILTER,
            '-resize', "${targetWidth}x${targetHeight}>",
            '-strip'
        )
        
        if ($outputExt -eq '.jpeg') {
            $magickArgs += @(
                '-quality', $JpegQuality,
                '-define', 'jpeg:optimize-coding=true',
                '-define', 'jpeg:dct-method=float',
                '-sampling-factor', $script:Config.JPEG_SAMPLING_FACTOR_DEFAULT
            )
        } else {
            $magickArgs += @(
                '-define', "png:compression-level=$($script:Config.PNG_COMPRESSION_LEVEL)",
                '-define', 'png:color-type=6'
            )
        }
        
        $magickArgs += $outputPath
        
        & $script:MagickExe $magickArgs 2>&1 | Out-Null
        
        if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $outputPath)) {
            throw "$ext conversion failed"
        }
        
        $beforeSize = $group.OriginalSizeBytes
        $afterSize = (Get-Item $outputPath).Length
        
        # Check minimum savings threshold
        if (-not (Test-MinSavingsThreshold -beforeSize $beforeSize -afterSize $afterSize)) {
            Remove-Item -LiteralPath $outputPath -Force
            $firstUsage = $group.Usages[0]
            
            $reason = if ($afterSize -ge $beforeSize) {
                'Conversion yielded no savings'
            } else {
                $actualSavings = (($beforeSize - $afterSize) / $beforeSize) * 100
                "Savings $($actualSavings.ToString('F1'))% below $($MinSavingsPercent.ToString('F1'))% threshold"
            }
            
            foreach ($usage in $group.Usages) {
                $usage.OptimizationStatus = 'Skipped_NoNetSavings'
                $usage.WhyNotOptimized = $reason
                $usage.ManualActionRequired = $false
            }
            
            Write-Host "   [INFO] $reason for '$($firstUsage.ShapeName)' on $($firstUsage.Location) - kept original" -ForegroundColor Cyan
            return @{ Success = $false }
        }
        
        # Update all relationships
        Write-Verbose "  Updating $($group.Usages.Count) relationship(s) to point to new file: $outputName"
        
        $statusName = if ($outputExt -eq '.jpeg') { 'ConvertedToJpeg' } else { 'ConvertedToPng' }
        
        foreach ($usage in $group.Usages) {
            $null = Update-BlipRelationship -usage $usage -tempDir $tempDir -newMediaFileName $outputName
            
            $usage.AfterSizeBytes = $afterSize
            $usage.OptimizationStatus = $statusName
            $usage.OptimizedFile = $outputName
        }
        
        # Remove original
        Remove-Item -LiteralPath $group.PhysicalPath -Force
        $group.OptimizedSizeBytes = $afterSize
        $group.OptimizedPath = $outputPath
        
        $savedPercent = (($beforeSize - $afterSize) / $beforeSize) * 100
        $firstUsage = $group.Usages[0]
        Write-Host "   [CONV] Converted $ext to $outputExt for '$($firstUsage.ShapeName)' on $($firstUsage.Location) (saved $($savedPercent.ToString('F1'))%)" -ForegroundColor Green
        
        return @{ Success = $true }
    }

    #endregion
