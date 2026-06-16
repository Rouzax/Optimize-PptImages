    #region CSV Reporting

    # Build a reverse index of which slides use each layout and master, keyed by part
    # file name (e.g. 'slideLayout47.xml'). Slide numbers use the canonical deck order
    # from $slides (sldIdLst), so they match the report's SlideNumber column.
    # Returns @{ Layout = @{ name = @(numbers) }; Master = @{ name = @(numbers) } }.
    function Get-PartSlideUsage {
        param(
            [string]$tempDir,
            [System.Collections.Generic.List[SlideInfo]]$slides
        )

        $layoutToSlides = @{}
        $layoutToMaster = @{}
        $layoutRelType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout'
        $masterRelType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster'

        foreach ($slide in $slides) {
            $slideFile = Split-Path $slide.PartPath -Leaf
            $slideRelsPath = Join-Path $tempDir "ppt/slides/_rels/$slideFile.rels"
            if (-not (Test-Path -LiteralPath $slideRelsPath)) { continue }
            $relsDoc = Get-XmlDocument $slideRelsPath
            $layoutRel = $relsDoc.SelectSingleNode("//x:Relationship[@Type='$layoutRelType']", $script:NsMgr)
            if (-not $layoutRel) { continue }
            $layoutFile = Split-Path $layoutRel.GetAttribute('Target') -Leaf
            if (-not $layoutToSlides.ContainsKey($layoutFile)) {
                $layoutToSlides[$layoutFile] = [System.Collections.Generic.List[int]]::new()
            }
            $layoutToSlides[$layoutFile].Add($slide.Number)
        }

        # Map each layout to its master so master usage can be rolled up.
        $layoutRelsDir = Join-Path $tempDir 'ppt/slideLayouts/_rels'
        if (Test-Path -LiteralPath $layoutRelsDir) {
            foreach ($relsFile in (Get-ChildItem -Path $layoutRelsDir -Filter '*.rels')) {
                $relsDoc = Get-XmlDocument $relsFile.FullName
                $masterRel = $relsDoc.SelectSingleNode("//x:Relationship[@Type='$masterRelType']", $script:NsMgr)
                if (-not $masterRel) { continue }
                $layoutFile = $relsFile.Name -replace '\.rels$', ''
                $layoutToMaster[$layoutFile] = Split-Path $masterRel.GetAttribute('Target') -Leaf
            }
        }

        $masterToSlides = @{}
        foreach ($layoutFile in $layoutToSlides.Keys) {
            if (-not $layoutToMaster.ContainsKey($layoutFile)) { continue }
            $masterFile = $layoutToMaster[$layoutFile]
            if (-not $masterToSlides.ContainsKey($masterFile)) {
                $masterToSlides[$masterFile] = [System.Collections.Generic.List[int]]::new()
            }
            $masterToSlides[$masterFile].AddRange($layoutToSlides[$layoutFile])
        }

        return @{ Layout = $layoutToSlides; Master = $masterToSlides }
    }

    # Format a sorted, de-duplicated slide-number list, or '(none)' when empty.
    function Format-SlideList {
        param([System.Collections.Generic.List[int]]$numbers)
        if (-not $numbers -or $numbers.Count -eq 0) { return '(none)' }
        return (($numbers | Sort-Object -Unique) -join ', ')
    }

    # Console summary of think-cell content so users can find it and judge what is safe
    # to remove. Covers think-cell images (by shape name) and the tag/embedding parts
    # (by content). Reports total size, carrying layouts, and the slides that use them.
    function Write-ThinkCellSummary {
        param(
            [string]$tempDir,
            [System.Collections.Generic.List[ImageUsage]]$usages,
            [hashtable]$partSlideUsage
        )

        $tcUsages = @($usages | Where-Object { Test-ThinkCellName $_.ShapeName })

        # Detect think-cell tag/embedding parts by content.
        $tcParts = [System.Collections.Generic.List[string]]::new()
        $tagsDir = Join-Path $tempDir 'ppt/tags'
        if (Test-Path -LiteralPath $tagsDir) {
            foreach ($f in (Get-ChildItem -Path $tagsDir -Filter '*.xml' -File)) {
                if ([System.IO.File]::ReadAllText($f.FullName) -match 'THINKCELL|think[\s-]?cell') {
                    $tcParts.Add("tags/$($f.Name)")
                }
            }
        }
        $embDir = Join-Path $tempDir 'ppt/embeddings'
        if (Test-Path -LiteralPath $embDir) {
            foreach ($f in (Get-ChildItem -Path $embDir -File)) {
                $bytes = [System.IO.File]::ReadAllBytes($f.FullName)
                $text = [System.Text.Encoding]::ASCII.GetString($bytes)
                if ($text -match 'CSmartGrid|thinkcell') {
                    $tcParts.Add("embeddings/$($f.Name)")
                }
            }
        }

        if ($tcUsages.Count -eq 0 -and $tcParts.Count -eq 0) { return }

        # Total size: unique image files used by think-cell shapes + the parts. Use the
        # scanned BeforeSizeBytes so the footprint is reported even after cleanup deletes
        # the files.
        $totalBytes = 0L
        $countedFiles = [System.Collections.Generic.HashSet[string]]::new()
        foreach ($u in $tcUsages) {
            if ($u.OriginalFileName -and $countedFiles.Add($u.OriginalFileName)) {
                $totalBytes += $u.BeforeSizeBytes
            }
        }
        foreach ($p in $tcParts) {
            $pp = Join-Path $tempDir "ppt/$p"
            if (Test-Path -LiteralPath $pp) { $totalBytes += (Get-Item -LiteralPath $pp).Length }
        }
        $sizeMb = [math]::Round($totalBytes / 1MB, 2)

        # Group carrying layouts (from the pre-cleanup scan) and partition them by whether
        # they still exist in the final package, so the summary reflects what actually
        # happened rather than only what was detected.
        $layoutFiles = @($tcUsages |
            Where-Object { $_.ContextType -eq [ContextType]::Layout -and $_.PartPath } |
            ForEach-Object { Split-Path $_.PartPath -Leaf } |
            Sort-Object -Unique)
        $presentLayouts = [System.Collections.Generic.List[string]]::new()
        $removedLayouts = [System.Collections.Generic.List[string]]::new()
        foreach ($lf in $layoutFiles) {
            if (Test-Path -LiteralPath (Join-Path $tempDir "ppt/slideLayouts/$lf")) {
                $presentLayouts.Add($lf)
            } else {
                $removedLayouts.Add($lf)
            }
        }

        # "Removed" only when everything detected is gone: no surviving layouts and no
        # remaining add-in parts.
        $fullyRemoved = ($removedLayouts.Count -gt 0 -and $presentLayouts.Count -eq 0 -and $tcParts.Count -eq 0)
        if ($fullyRemoved) {
            Write-Host "`n[THINK-CELL] Removed think-cell content ($sizeMb MB)" -ForegroundColor Cyan
        } else {
            Write-Host "`n[THINK-CELL] Detected think-cell content ($sizeMb MB) - still in the file" -ForegroundColor Cyan
        }

        if ($removedLayouts.Count -gt 0) {
            $nums = ($removedLayouts | ForEach-Object { if ($_ -match '(\d+)\.xml$') { [int]$matches[1] } }) | Sort-Object -Unique
            Write-Host "             Removed unused layouts: $($nums -join ', ')" -ForegroundColor Cyan
        }
        if ($presentLayouts.Count -gt 0) {
            $allSlides = [System.Collections.Generic.List[int]]::new()
            foreach ($lf in $presentLayouts) {
                if ($partSlideUsage.Layout.ContainsKey($lf)) {
                    $allSlides.AddRange($partSlideUsage.Layout[$lf])
                }
            }
            $nums = ($presentLayouts | ForEach-Object { if ($_ -match '(\d+)\.xml$') { [int]$matches[1] } }) | Sort-Object -Unique
            Write-Host "             Carried by layouts: $($nums -join ', ')" -ForegroundColor Cyan
            Write-Host "             Used on slides: $(Format-SlideList $allSlides)" -ForegroundColor Cyan
            if ($allSlides.Count -eq 0 -and -not $script:CleanUnusedLayouts) {
                Write-Host "             These layouts are unused; run with -CleanUnusedLayouts to remove them." -ForegroundColor Cyan
            }
        }
        if ($tcParts.Count -gt 0) {
            Write-Host "             Add-in parts: $($tcParts -join ', ')" -ForegroundColor Cyan
        }
    }

    function Export-CsvReport {
        param(
            [System.Collections.Generic.List[ImageUsage]]$usages,
            [hashtable]$partSlideUsage
        )
        
        Write-Host "`n[CSV] Generating report..." -ForegroundColor Yellow
        
        # Create temp file for CSV (outside the extraction temp dir)
        $tempCsvPath = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "ppt-opt-$([System.Guid]::NewGuid().ToString('N').Substring(0,8)).csv")
        Write-Verbose "Temp CSV file: $tempCsvPath"
        
        if (-not $usages -or $usages.Count -eq 0) {
            Write-Host "No usages to report" -ForegroundColor Gray
            @() | Export-Csv -Path $tempCsvPath -NoTypeInformation -Encoding UTF8
            return $tempCsvPath
        }
        
        # Filter usages based on slide selection (same logic as processing phases)
        $slideFiltersActive = ($IncludeSlides -and $IncludeSlides.Count -gt 0) -or ($ExcludeSlides -and $ExcludeSlides.Count -gt 0)
        if ($slideFiltersActive) {
            $filteredUsages = @($usages | Where-Object {
                # Only include slide usages that pass the filter (exclude masters/layouts when filtering)
                $_.ContextType -eq [ContextType]::Slide -and (Test-SlideIncluded -SlideNumber $_.SlideNumber)
            })
            Write-Verbose "CSV filtered to $($filteredUsages.Count) usages (slide filter active)"
        } else {
            $filteredUsages = $usages
        }
        
        if ($filteredUsages.Count -eq 0) {
            Write-Host "No usages in scope to report" -ForegroundColor Gray
            @() | Export-Csv -Path $tempCsvPath -NoTypeInformation -Encoding UTF8
            return $tempCsvPath
        }
        
        # Sort by context type (Slide, Layout, Master), then slide number, then layout/master number
        $sorted = $filteredUsages | Sort-Object @(
            @{ Expression = { $_.ContextType }; Ascending = $true }
            @{ Expression = { $_.SlideNumber }; Ascending = $true }
            @{ Expression = { 
                if ($_.PartPath -match '(\d+)\.xml$') { 
                    [int]$matches[1] 
                } else { 
                    0 
                }
            }; Ascending = $true }
        )
        
        $rows = foreach ($usage in $sorted) {
            # Reverse-link: which slides use this layout/master asset. Slide rows already
            # carry their own SlideNumber, so leave UsedOnSlides blank there.
            $usedOnSlides = ''
            if ($partSlideUsage -and $usage.PartPath) {
                $partFile = Split-Path $usage.PartPath -Leaf
                if ($usage.ContextType -eq [ContextType]::Layout) {
                    $usedOnSlides = Format-SlideList $partSlideUsage.Layout[$partFile]
                } elseif ($usage.ContextType -eq [ContextType]::Master) {
                    $usedOnSlides = Format-SlideList $partSlideUsage.Master[$partFile]
                }
            }

            [PSCustomObject]@{
                Location = $usage.Location
                ContextType = $usage.ContextType.ToString()
                SlideNumber = $usage.SlideNumber
                SlideModified = if ($usage.CropApplied -or $usage.OptimizationStatus -match '^(Optimized|Converted)') { 'Yes' } else { 'No' }
                ShapeName = $usage.ShapeName
                IsThinkCell = if (Test-ThinkCellName $usage.ShapeName) { 'Yes' } else { 'No' }
                UsedOnSlides = $usedOnSlides
                MediaPart = if ($usage.OriginalFileName) { "ppt/media/$($usage.OriginalFileName)" } else { '' }
                SourceWidthPx = $usage.SourceWidthPx
                SourceHeightPx = $usage.SourceHeightPx
                DisplayWidthPx = $usage.DisplayWidthPx
                DisplayHeightPx = $usage.DisplayHeightPx
                TargetWidthPx = $usage.TargetWidthPx
                TargetHeightPx = $usage.TargetHeightPx
                OriginalFile = $usage.OriginalFileName
                OptimizedFile = if ($usage.OptimizedFile) { $usage.OptimizedFile } else { $usage.OriginalFileName }
                BeforeSizeBytes = $usage.BeforeSizeBytes
                AfterSizeBytes = if ($usage.AfterSizeBytes -gt 0) { $usage.AfterSizeBytes } else { $usage.BeforeSizeBytes }
                Savings = Format-Savings -before $usage.BeforeSizeBytes -after $(if ($usage.AfterSizeBytes -gt 0) { $usage.AfterSizeBytes } else { $usage.BeforeSizeBytes })
                HasSrcRect = $usage.HasSrcRect
                CropApplied = if ($usage.CropApplied) { 'Yes' } else { 'No' }
                CropNormalized = if ($usage.CropNormalized) { 'Yes' } else { 'No' }
                CropRemovedNoOp = if ($usage.CropRemovedNoOp) { 'Yes' } else { 'No' }
                SvgFallback = if ($usage.IsSvgFallbackUsage) { 'Yes' } else { 'No' }
                OptimizationStatus = $usage.OptimizationStatus
                WhyNotOptimized = $usage.WhyNotOptimized
                HeadroomFactor = $HeadroomFactor
                JpegQuality = $JpegQuality
                SourceJpegQuality = if ($usage.SourceJpegQuality -gt 0) { $usage.SourceJpegQuality } else { ' ' }
                TransparencyThresholdPercent = $TransparencyThresholdPercent
                EffectiveTransparencyPercent = $usage.EffectiveTransparencyPercent.ToString('F2')
                MinSavingsPercent = $MinSavingsPercent
                ManualActionRequired = if ($usage.ManualActionRequired) { 'Yes' } else { 'No' }
                ManualActionHint = $usage.ManualActionHint
            }
        }
        
        $rows | Export-Csv -Path $tempCsvPath -NoTypeInformation -Encoding UTF8
        Write-Host "[OK] CSV report generated" -ForegroundColor Green
        
        return $tempCsvPath
    }

    #endregion

    #region Output File Handling

    function Copy-OutputFilesToFinal {
        param(
            [string]$tempPptxPath,
            [string]$tempCsvPath,
            [string]$finalPptxPath,
            [string]$finalCsvPath
        )
        
        $result = @{
            PptxSuccess = $false
            CsvSuccess = $false
            PptxTempPath = $tempPptxPath
            CsvTempPath = $tempCsvPath
        }
        
        # Try to copy PPTX
        try {
            if (Test-FileInUse -Path $finalPptxPath) {
                Write-Warning "Cannot write to output file (in use): $finalPptxPath"
            } else {
                if (Test-Path -LiteralPath $finalPptxPath) {
                    Remove-Item -LiteralPath $finalPptxPath -Force
                }
                Copy-Item -LiteralPath $tempPptxPath -Destination $finalPptxPath -Force
                $result.PptxSuccess = $true
                Write-Host "[OUT] Output: $finalPptxPath" -ForegroundColor Green
            }
        } catch {
            Write-Warning "Failed to copy PPTX to final destination: $_"
        }
        
        # Try to copy CSV
        try {
            if (Test-FileInUse -Path $finalCsvPath) {
                Write-Warning "Cannot write to CSV file (in use): $finalCsvPath"
            } else {
                if (Test-Path -LiteralPath $finalCsvPath) {
                    Remove-Item -LiteralPath $finalCsvPath -Force
                }
                Copy-Item -LiteralPath $tempCsvPath -Destination $finalCsvPath -Force
                $result.CsvSuccess = $true
                Write-Host "[OUT] Report: $finalCsvPath" -ForegroundColor Green
            }
        } catch {
            Write-Warning "Failed to copy CSV to final destination: $_"
        }
        
        return $result
    }

    #endregion
