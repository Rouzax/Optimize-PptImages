    #region Crop Normalization

    function Invoke-CropNormalization {
        param([ImageUsage]$usage)
        
        if (-not $usage.HasSrcRect) { return }
        
        $l = $usage.SrcRectLeft
        $t = $usage.SrcRectTop
        $r = $usage.SrcRectRight
        $b = $usage.SrcRectBottom
        
        $needsNormalization = $false
        $originalValues = @{}
        
        # Detect out-of-bounds
        if ($l -lt 0 -or $l -gt $script:Config.CROP_THOUSANDTHS_MAX -or
            $t -lt 0 -or $t -gt $script:Config.CROP_THOUSANDTHS_MAX -or
            $r -lt 0 -or $r -gt $script:Config.CROP_THOUSANDTHS_MAX -or
            $b -lt 0 -or $b -gt $script:Config.CROP_THOUSANDTHS_MAX) {
            $needsNormalization = $true
            $originalValues = @{ l=$l; t=$t; r=$r; b=$b }
        }
        
        if (-not $needsNormalization) { return }
        
        # Clamp to valid range
        $l = [Math]::Max(0, [Math]::Min($l, $script:Config.CROP_THOUSANDTHS_MAX))
        $t = [Math]::Max(0, [Math]::Min($t, $script:Config.CROP_THOUSANDTHS_MAX))
        $r = [Math]::Max(0, [Math]::Min($r, $script:Config.CROP_THOUSANDTHS_MAX))
        $b = [Math]::Max(0, [Math]::Min($b, $script:Config.CROP_THOUSANDTHS_MAX))
        
        # Compensate xfrm for normalized crops
        if ($usage.XfrmElement) {
            $ext = $usage.XfrmElement.SelectSingleNode('a:ext', $script:NsMgr)
            $off = $usage.XfrmElement.SelectSingleNode('a:off', $script:NsMgr)
            
            if ($ext -and $off) {
                $origCx = [long]$ext.GetAttribute('cx')
                $origCy = [long]$ext.GetAttribute('cy')
                $origX = [long]$off.GetAttribute('x')
                $origY = [long]$off.GetAttribute('y')
                
                # Calculate visible proportions (accounting for negative = extension)
                $origWidthProp = (100000.0 - $originalValues.l - $originalValues.r) / 100000.0
                $origHeightProp = (100000.0 - $originalValues.t - $originalValues.b) / 100000.0
                $newWidthProp = (100000.0 - $l - $r) / 100000.0
                $newHeightProp = (100000.0 - $t - $b) / 100000.0
                
                # Scale dimensions
                $newCx = [long]($origCx * ($newWidthProp / $origWidthProp))
                $newCy = [long]($origCy * ($newHeightProp / $origHeightProp))
                
                # Adjust position if negative crops were normalized
                $newX = $origX
                $newY = $origY
                
                if ($originalValues.l -lt 0 -and $l -eq 0) {
                    $newX = $origX + [long]($origCx * ([Math]::Abs($originalValues.l) / (100000.0 - $originalValues.l - $originalValues.r)))
                }
                if ($originalValues.t -lt 0 -and $t -eq 0) {
                    $newY = $origY + [long]($origCy * ([Math]::Abs($originalValues.t) / (100000.0 - $originalValues.t - $originalValues.b)))
                }
                
                # Clamp to minimums
                $newCx = [Math]::Max($script:Config.MIN_EMU_DIMENSION, $newCx)
                $newCy = [Math]::Max($script:Config.MIN_EMU_DIMENSION, $newCy)
                $newX = [Math]::Max(0, $newX)
                $newY = [Math]::Max(0, $newY)
                
                $ext.SetAttribute('cx', $newCx.ToString())
                $ext.SetAttribute('cy', $newCy.ToString())
                $off.SetAttribute('x', $newX.ToString())
                $off.SetAttribute('y', $newY.ToString())
            }
        }
        
        # Update srcRect in XML (sibling of blip within blipFill)
        $blipFill = $usage.BlipElement.ParentNode
        if ($blipFill) {
            $srcRect = $blipFill.SelectSingleNode('a:srcRect', $script:NsMgr)
            if ($srcRect) {
                if ($l -eq 0) { $srcRect.RemoveAttribute('l') } else { $srcRect.SetAttribute('l', ([int]$l).ToString()) }
                if ($t -eq 0) { $srcRect.RemoveAttribute('t') } else { $srcRect.SetAttribute('t', ([int]$t).ToString()) }
                if ($r -eq 0) { $srcRect.RemoveAttribute('r') } else { $srcRect.SetAttribute('r', ([int]$r).ToString()) }
                if ($b -eq 0) { $srcRect.RemoveAttribute('b') } else { $srcRect.SetAttribute('b', ([int]$b).ToString()) }
            }
        }
        
        # Update usage
        $usage.SrcRectLeft = $l
        $usage.SrcRectTop = $t
        $usage.SrcRectRight = $r
        $usage.SrcRectBottom = $b
        $usage.CropNormalized = $true
        
        # Recompute display size
        if ($usage.XfrmElement) {
            $ext = $usage.XfrmElement.SelectSingleNode('a:ext', $script:NsMgr)
            if ($ext) {
                $usage.DisplayWidthPx = ConvertFrom-EMU -emu ([long]$ext.GetAttribute('cx'))
                $usage.DisplayHeightPx = ConvertFrom-EMU -emu ([long]$ext.GetAttribute('cy'))
            }
        }
        
        # Build change description
        $changes = @()
        $details = @()
        foreach ($key in @('l', 't', 'r', 'b')) {
            $oldVal = $originalValues[$key]
            $newVal = Get-Variable -Name $key -ValueOnly
            if ($oldVal -ne $newVal) {
                $changes += "${key}:${oldVal}->${newVal}"
                
                # Build detailed explanation for verbose
                $side = switch ($key) {
                    'l' { 'left' }
                    't' { 'top' }
                    'r' { 'right' }
                    'b' { 'bottom' }
                }
                
                $reason = if ($oldVal -lt 0) {
                    "extended beyond image boundary (negative value)"
                } elseif ($oldVal -gt $script:Config.CROP_THOUSANDTHS_MAX) {
                    "exceeded maximum allowed value ($($script:Config.CROP_THOUSANDTHS_MAX))"
                } else {
                    "was out of valid range [0-$($script:Config.CROP_THOUSANDTHS_MAX)]"
                }
                
                $details += "  * $side edge: $oldVal -> $newVal (was $reason)"
            }
        }
        
        if ($changes.Count -gt 0) {
            Write-Host "[NOTE] Corrected illegal crop values for '$($usage.ShapeName)' on $($usage.Location) ($($changes -join ', '))" -ForegroundColor Cyan
            
            if (Test-VerboseMode) {
                Write-Verbose "Illegal crop correction details for '$($usage.ShapeName)':"
                foreach ($detail in $details) {
                    Write-Verbose $detail
                }
                
                # Show xfrm adjustments if made
                if ($usage.XfrmElement) {
                    $ext = $usage.XfrmElement.SelectSingleNode('a:ext', $script:NsMgr)
                    $off = $usage.XfrmElement.SelectSingleNode('a:off', $script:NsMgr)
                    if ($ext -and $off) {
                        Write-Verbose "  Adjusted shape transform (xfrm) to maintain visual appearance:"
                        Write-Verbose "    * Position: x=$($off.GetAttribute('x')) EMU, y=$($off.GetAttribute('y')) EMU"
                        Write-Verbose "    * Size: cx=$($ext.GetAttribute('cx')) EMU ($($usage.DisplayWidthPx)px), cy=$($ext.GetAttribute('cy')) EMU ($($usage.DisplayHeightPx)px)"
                    }
                }
                
                Write-Verbose "  Crop values are in thousandths (0-100000 range, where 100000 = 100%)"
                Write-Verbose "  Negative values indicate the image was extended beyond its boundaries"
            }
        }
    }

    #endregion

    #region Cropping

    function Invoke-ImageCropping {
        [CmdletBinding(SupportsShouldProcess)]
        param(
            [System.Collections.Generic.List[ImageUsage]]$usages,
            [string]$tempDir,
            [hashtable]$morphPairs
        )

        Write-Host "`n[CROP] Processing image crops..." -ForegroundColor Yellow

        $croppedCount = 0
        $skippedCount = 0

        # Determine if slide filters are active (used for scoping)
        $slideFiltersActive = ($IncludeSlides -and $IncludeSlides.Count -gt 0) -or ($ExcludeSlides -and $ExcludeSlides.Count -gt 0)

        # Get items with crops that are in scope for progress counter
        # (excludes items silently skipped due to slide filters, includes SVG fallbacks since they get skip messages)
        $itemsWithCrops = @($usages | Where-Object {
            $_.HasSrcRect -and
            # Check slide filter
            ($_.ContextType -ne [ContextType]::Slide -or (Test-SlideIncluded -SlideNumber $_.SlideNumber)) -and
            # When slide filters active, exclude masters/layouts
            (-not $slideFiltersActive -or $_.ContextType -eq [ContextType]::Slide)
        })
        $totalItems = $itemsWithCrops.Count
        $currentItem = 0

        $scratchDir = Join-Path $tempDir '.crop-scratch'
        New-Item -ItemType Directory -Path $scratchDir -Force | Out-Null

        # --- BUILD PASS ---
        # Iterate usages in the existing order. Silent slide-filter skips and the
        # ShouldProcess gate are honored here. For each usage with HasSrcRect that is
        # in scope, call Get-CropJob. Null results are tallied and messaged immediately.
        # Non-null jobs are collected for the parallel run. $jobs holds one leader per
        # unique CropKey; $usageJobPairs pairs every in-scope usage with its CropKey so
        # the commit pass can look up results and iterate in usage order.
        $jobs = [System.Collections.Generic.List[CropJob]]::new()
        $usageJobPairs = [System.Collections.Generic.List[PSCustomObject]]::new()
        $cropKeyToLeader = @{}

        foreach ($usage in $usages) {
            if (-not $usage.HasSrcRect) { continue }

            # Silent skip checks FIRST (before progress update)

            # Check slide filter (silently skip usages not in included slides)
            if ($usage.ContextType -eq [ContextType]::Slide -and -not (Test-SlideIncluded -SlideNumber $usage.SlideNumber)) {
                continue
            }

            # When slide filters are active, skip masters/layouts entirely
            # (they affect all slides, not just the filtered ones)
            if ($slideFiltersActive -and ($usage.ContextType -eq [ContextType]::Master -or $usage.ContextType -eq [ContextType]::Layout)) {
                continue
            }

            # Now update progress (item is in scope)
            $currentItem++
            Write-Progress -Activity "Processing crops" -Status "Preparing: $currentItem of $totalItems" `
                -PercentComplete (($currentItem / [Math]::Max(1, $totalItems)) * 100) -Id 2

            # Decide phase: Get-CropJob handles SVG/animated-GIF/flag/no-op/morph/validity decisions.
            # It sets usage.OptimizationStatus on skip/no-op and returns $null; returns a CropJob otherwise.
            $job = Get-CropJob -usage $usage -scratchDir $scratchDir

            if ($null -eq $job) {
                # Get-CropJob set the status. Emit console messages and tally based on what was decided.
                if ($usage.CropRemovedNoOp) {
                    # No-op crop removed: routine informational, route to verbose.
                    Write-Verbose "   [NOTE] Removed no-op crop for '$($usage.ShapeName)' on $($usage.Location) (full image)"
                    continue
                }

                switch ($usage.OptimizationStatus) {
                    'Skipped_SvgFallback' {
                        Write-Verbose "   [SKIP] Skipped crop for '$($usage.ShapeName)' on $($usage.Location): SVG fallback (PowerPoint regenerates)"
                        $skippedCount++
                    }
                    'Skipped_AnimatedGif' {
                        Write-Verbose "   [SKIP] Skipped crop for '$($usage.ShapeName)' on $($usage.Location): animated GIF (would break animation)"
                        $skippedCount++
                    }
                    'Skipped_CropNotMaterialized' {
                        if ($usage.ContextType -eq [ContextType]::Slide) {
                            Write-Host "   [SKIP] Skipped crop for '$($usage.ShapeName)' on $($usage.Location): enable -CropSlides" -ForegroundColor Gray
                        } else {
                            Write-Host "   [SKIP] Skipped crop for '$($usage.ShapeName)' on $($usage.Location): enable -CropMastersAndLayouts" -ForegroundColor Gray
                        }
                        $skippedCount++
                    }
                    'Skipped_MorphCropConflict' {
                        # Emit the [BLOCK] message and add 2 to match the old behavior (both pair members counted here).
                        $pair = $usage.MorphPair
                        $pairSlide = if ($pair) { $pair.SlideNumber } else { '?' }
                        Write-Host "   [BLOCK] Morph crop conflict for '$($usage.ShapeName)' across Slide $pairSlide <-> $($usage.SlideNumber) (skipping crop & optimization)" -ForegroundColor Red
                        $skippedCount += 2
                    }
                    'Skipped_CropInvalidOrUnsafe' {
                        Write-Host "[BLOCK] Skipped crop for '$($usage.ShapeName)' on $($usage.Location): invalid crop geometry" -ForegroundColor Red
                        $skippedCount++
                    }
                    default {
                        # Unknown skip status: count as skipped to be safe.
                        $skippedCount++
                    }
                }
                continue
            }

            # ShouldProcess gate: only collect the job when the operation is confirmed.
            if ($PSCmdlet.ShouldProcess("$($usage.ShapeName) on $($usage.Location)", "Crop image")) {
                if (-not $cropKeyToLeader.ContainsKey($job.CropKey)) {
                    $cropKeyToLeader[$job.CropKey] = $job
                    $jobs.Add($job)
                }
                $usageJobPairs.Add([PSCustomObject]@{ Usage = $usage; CropKey = $job.CropKey })
            }
        }

        # --- PARALLEL RUN ---
        # Execute all magick crop jobs in parallel; results keyed by Key.
        $jobResults = @{}
        if ($jobs.Count -gt 0) {
            $jobResults = Invoke-MagickJobsParallel -jobs $jobs -activity 'Processing crops' -progressId 2
        }

        # --- COMMIT PASS ---
        # Iterate usageJobPairs in the same order they were added (usage order).
        # Each unique CropKey is allocated a shared media file on first encounter;
        # subsequent usages with the same CropKey reuse that file.
        # A missing result key is treated as success=$false.
        $cropKeyToMedia = @{}
        foreach ($pair in $usageJobPairs) {
            $usage   = $pair.Usage
            $cropKey = $pair.CropKey
            $leader  = $cropKeyToLeader[$cropKey]
            $result  = $jobResults[$leader.Key]
            $ran       = $result -ne $null
            $success   = $ran -and $result.Success
            $afterSize = if ($ran) { $result.AfterSize } else { 0 }

            if (-not $success) {
                $usage.OptimizationStatus = 'Skipped_CropInvalidOrUnsafe'
                $usage.WhyNotOptimized    = 'Crop operation failed'
                Write-Warning "Crop failed for '$($usage.ShapeName)' on $($usage.Location)"
                $skippedCount++
                continue
            }

            if (-not $cropKeyToMedia.ContainsKey($cropKey)) {
                $out = New-UniqueMediaPath -basePath $leader.SourcePath -suffix '_cropped'
                Move-Item -LiteralPath $leader.ScratchPath $out.Path -Force
                $cropKeyToMedia[$cropKey] = $out
            }
            $shared = $cropKeyToMedia[$cropKey]

            Complete-CropUsage -usage $usage -sharedMediaName $shared.Name -sharedMediaPath $shared.Path `
                -beforeSize $leader.BeforeSize -afterSize $afterSize -tempDir $tempDir
            $croppedCount++
        }

        if (Test-Path -LiteralPath $scratchDir) {
            Remove-Item -LiteralPath $scratchDir -Recurse -Force -ErrorAction SilentlyContinue
        }

        Write-Progress -Activity "Processing crops" -Completed -Id 2
        Write-Host "[STATS] Cropping phase complete: $croppedCount cropped, $skippedCount skipped" -ForegroundColor Green
    }

    function Test-CropValidity {
        param([ImageUsage]$usage)
        
        $l = $usage.SrcRectLeft
        $t = $usage.SrcRectTop
        $r = $usage.SrcRectRight
        $b = $usage.SrcRectBottom
        
        # Check bounds
        if ($l -lt 0 -or $l -gt $script:Config.CROP_THOUSANDTHS_MAX) {
            if (Test-VerboseMode) {
                Write-Verbose "  Crop validation failed for '$($usage.ShapeName)': left=$l out of range [0-$($script:Config.CROP_THOUSANDTHS_MAX)]"
            }
            return $false
        }
        if ($t -lt 0 -or $t -gt $script:Config.CROP_THOUSANDTHS_MAX) {
            if (Test-VerboseMode) {
                Write-Verbose "  Crop validation failed for '$($usage.ShapeName)': top=$t out of range [0-$($script:Config.CROP_THOUSANDTHS_MAX)]"
            }
            return $false
        }
        if ($r -lt 0 -or $r -gt $script:Config.CROP_THOUSANDTHS_MAX) {
            if (Test-VerboseMode) {
                Write-Verbose "  Crop validation failed for '$($usage.ShapeName)': right=$r out of range [0-$($script:Config.CROP_THOUSANDTHS_MAX)]"
            }
            return $false
        }
        if ($b -lt 0 -or $b -gt $script:Config.CROP_THOUSANDTHS_MAX) {
            if (Test-VerboseMode) {
                Write-Verbose "  Crop validation failed for '$($usage.ShapeName)': bottom=$b out of range [0-$($script:Config.CROP_THOUSANDTHS_MAX)]"
            }
            return $false
        }
        
        # Check sum
        if (($l + $r) -ge $script:Config.CROP_THOUSANDTHS_MAX) {
            if (Test-VerboseMode) {
                Write-Verbose "  Crop validation failed for '$($usage.ShapeName)': left+right=$($l+$r) >= $($script:Config.CROP_THOUSANDTHS_MAX) (crops overlap horizontally)"
            }
            return $false
        }
        if (($t + $b) -ge $script:Config.CROP_THOUSANDTHS_MAX) {
            if (Test-VerboseMode) {
                Write-Verbose "  Crop validation failed for '$($usage.ShapeName)': top+bottom=$($t+$b) >= $($script:Config.CROP_THOUSANDTHS_MAX) (crops overlap vertically)"
            }
            return $false
        }
        
        # Check minimum dimensions
        $crop = Get-CropPercentage -left $l -top $t -right $r -bottom $b
        if ($crop.WidthPercent -lt $script:Config.MIN_CROPPED_DIMENSION_PERCENT) {
            if (Test-VerboseMode) {
                Write-Verbose "  Crop validation failed for '$($usage.ShapeName)': resulting width $($crop.WidthPercent.ToString('F1'))% < minimum $($script:Config.MIN_CROPPED_DIMENSION_PERCENT)%"
            }
            return $false
        }
        if ($crop.HeightPercent -lt $script:Config.MIN_CROPPED_DIMENSION_PERCENT) {
            if (Test-VerboseMode) {
                Write-Verbose "  Crop validation failed for '$($usage.ShapeName)': resulting height $($crop.HeightPercent.ToString('F1'))% < minimum $($script:Config.MIN_CROPPED_DIMENSION_PERCENT)%"
            }
            return $false
        }
        
        if (Test-VerboseMode) {
            Write-Verbose "  Crop validation passed for '$($usage.ShapeName)': l=$l, t=$t, r=$r, b=$b -> $($crop.WidthPercent.ToString('F1'))%x$($crop.HeightPercent.ToString('F1'))% of original"
        }
        
        return $true
    }

    function Get-CropJob {
        param(
            [ImageUsage]$usage,
            [string]$scratchDir
        )

        # SVG fallback: skip silently (PowerPoint regenerates these)
        if ($usage.IsSvgFallbackUsage) {
            $usage.OptimizationStatus = 'Skipped_SvgFallback'
            $usage.WhyNotOptimized = 'PNG is auto-generated SVG fallback; PowerPoint regenerates'
            $usage.ManualActionRequired = $false
            return $null
        }

        # Animated GIF: cropping would break the animation
        $ext = [System.IO.Path]::GetExtension($usage.ImagePhysicalPath).ToLower()
        if ($ext -eq '.gif' -and (Test-IsAnimatedGif -ImagePath $usage.ImagePhysicalPath)) {
            $usage.OptimizationStatus = 'Skipped_AnimatedGif'
            $usage.WhyNotOptimized = 'Animated GIF - would break animation'
            $usage.ManualActionRequired = $false
            return $null
        }

        # Crop-flag check: is cropping enabled for this context?
        if ($usage.ContextType -eq [ContextType]::Slide) {
            if (-not $script:CropSlides) {
                $usage.OptimizationStatus = 'Skipped_CropNotMaterialized'
                $usage.WhyNotOptimized = 'Slide cropping disabled; enable -CropSlides'
                $usage.ManualActionRequired = $true
                $usage.ManualActionHint = 'Run with -CropSlides'
                return $null
            }
        } elseif ($usage.ContextType -eq [ContextType]::Master -or $usage.ContextType -eq [ContextType]::Layout) {
            if (-not $script:CropMastersAndLayouts) {
                $usage.OptimizationStatus = 'Skipped_CropNotMaterialized'
                $usage.WhyNotOptimized = 'Master/Layout cropping disabled; enable -CropMastersAndLayouts'
                $usage.ManualActionRequired = $true
                $usage.ManualActionHint = 'Run with -CropMastersAndLayouts'
                return $null
            }
        }

        # No-op crop: remove the srcRect from XML and mark as cleaned up, no crop job needed
        if (Test-IsNoOpCrop -left $usage.SrcRectLeft -top $usage.SrcRectTop -right $usage.SrcRectRight -bottom $usage.SrcRectBottom) {
            $blipFill = $usage.BlipElement.ParentNode
            if ($blipFill) {
                $srcRect = $blipFill.SelectSingleNode('a:srcRect', $script:NsMgr)
                if ($srcRect) {
                    $null = $srcRect.ParentNode.RemoveChild($srcRect)
                }
            }
            $usage.HasSrcRect = $false
            $usage.CropRemovedNoOp = $true
            return $null
        }

        # Morph-pair conflict: meaningful crop on either side of a morph transition
        if ($usage.MorphPair) {
            $pair = $usage.MorphPair
            $usageHasMeaningful = -not (Test-IsNoOpCrop -left $usage.SrcRectLeft -top $usage.SrcRectTop -right $usage.SrcRectRight -bottom $usage.SrcRectBottom)
            $pairHasMeaningful = $pair.HasSrcRect -and -not (Test-IsNoOpCrop -left $pair.SrcRectLeft -top $pair.SrcRectTop -right $pair.SrcRectRight -bottom $pair.SrcRectBottom)

            if ($usageHasMeaningful -or $pairHasMeaningful) {
                $usage.OptimizationStatus = 'Skipped_MorphCropConflict'
                $pair.OptimizationStatus = 'Skipped_MorphCropConflict'
                $usage.WhyNotOptimized = 'Morph transition with crop conflict'
                $pair.WhyNotOptimized = 'Morph transition with crop conflict'
                $usage.ManualActionRequired = $true
                $pair.ManualActionRequired = $true
                $usage.ManualActionHint = 'Manually align crops or remove Morph'
                $pair.ManualActionHint = 'Manually align crops or remove Morph'
                return $null
            }
        }

        # Validate crop geometry
        if (-not (Test-CropValidity -usage $usage)) {
            $usage.OptimizationStatus = 'Skipped_CropInvalidOrUnsafe'
            $usage.WhyNotOptimized = 'Invalid crop geometry'
            $usage.ManualActionRequired = $true
            $usage.ManualActionHint = 'Review and fix crop values'
            return $null
        }

        # Resolve source dimensions from cache; fall back to a live identify only if not cached
        $srcWidth = $usage.SourceWidthPx
        $srcHeight = $usage.SourceHeightPx
        if ($srcWidth -eq 0) {
            $dims = Get-ImageDimensions -ImagePath $usage.ImagePhysicalPath
            $srcWidth = $dims.Width
            $srcHeight = $dims.Height
            $usage.SourceWidthPx = $srcWidth
            $usage.SourceHeightPx = $srcHeight
        }

        # Compute crop geometry in pixels from the srcRect percentages
        $leftPx   = [Math]::Round($srcWidth  * $usage.SrcRectLeft   / 100000.0)
        $topPx    = [Math]::Round($srcHeight * $usage.SrcRectTop    / 100000.0)
        $rightPx  = [Math]::Round($srcWidth  * $usage.SrcRectRight  / 100000.0)
        $bottomPx = [Math]::Round($srcHeight * $usage.SrcRectBottom / 100000.0)

        $cropWidth  = $srcWidth  - $leftPx - $rightPx
        $cropHeight = $srcHeight - $topPx  - $bottomPx

        # Clamp
        $cropWidth  = [Math]::Max(1, [Math]::Min($cropWidth,  $srcWidth))
        $cropHeight = [Math]::Max(1, [Math]::Min($cropHeight, $srcHeight))
        $leftPx     = [Math]::Max(0, [Math]::Min($leftPx,  $srcWidth  - 1))
        $topPx      = [Math]::Max(0, [Math]::Min($topPx,   $srcHeight - 1))

        $cropGeometry = "${cropWidth}x${cropHeight}+${leftPx}+${topPx}"
        $scratchPath  = Join-Path $scratchDir ("$([guid]::NewGuid())" + [System.IO.Path]::GetExtension($usage.ImagePhysicalPath))

        $job = [CropJob]::new()
        $job.Key         = "$($usage.PartPath)|$($usage.ShapeName)|$($usage.BlipRId)"
        $job.CropKey     = "$($usage.ImagePhysicalPath)|$cropGeometry"
        $job.SourcePath  = $usage.ImagePhysicalPath
        $job.ScratchPath = $scratchPath
        $job.MagickArgs  = @($usage.ImagePhysicalPath, '-crop', $cropGeometry, '+repage', $scratchPath)
        $job.BeforeSize  = $usage.BeforeSizeBytes
        $job.Usage       = $usage

        return $job
    }

    function Complete-CropUsage {
        param(
            [object]$usage,
            [string]$sharedMediaName,
            [string]$sharedMediaPath,
            [long]$beforeSize,
            [long]$afterSize,
            [string]$tempDir
        )

        $u = $usage
        $null = Update-BlipRelationship -usage $u -tempDir $tempDir -newMediaFileName $sharedMediaName

        $blipFill = $u.BlipElement.ParentNode
        if ($blipFill) {
            $srcRect = $blipFill.SelectSingleNode('a:srcRect', $script:NsMgr)
            if ($srcRect) { $null = $srcRect.ParentNode.RemoveChild($srcRect) }
            $stretch = $blipFill.SelectSingleNode('a:stretch', $script:NsMgr)
            if ($stretch) {
                $fillRect = $stretch.SelectSingleNode('a:fillRect', $script:NsMgr)
                if ($fillRect) {
                    $hasNegative = $false
                    foreach ($attr in @('l', 't', 'r', 'b')) {
                        $val = $fillRect.GetAttribute($attr)
                        if ($val -and [double]$val -lt 0) { $hasNegative = $true; break }
                    }
                    if ($hasNegative) {
                        $fillRect.RemoveAttribute('l'); $fillRect.RemoveAttribute('t')
                        $fillRect.RemoveAttribute('r'); $fillRect.RemoveAttribute('b')
                        Write-Verbose "    Reset fillRect attributes (crop materialized into image)"
                    }
                }
            }
        }

        $u.HasSrcRect        = $false
        $u.CropApplied       = $true
        $u.AfterSizeBytes    = $afterSize
        $u.OptimizedFile     = $sharedMediaName
        $u.ImagePhysicalPath = $sharedMediaPath

        $savedPercent = if ($beforeSize -gt 0) { (($beforeSize - $afterSize) / $beforeSize) * 100 } else { 0 }
        Write-Verbose "   [CROP] Cropped '$($u.ShapeName)' on $($u.Location) (saved $($savedPercent.ToString('F1'))%)"
    }

    #endregion
