    #region Main Processing Function

    function Invoke-PptOptimization {
        param([string]$CurrentInputPath)
        
        $exitCode = 0
        
        try {
            Write-Host "[SCAN] Analyzing presentation structure..." -ForegroundColor Cyan
            Write-Host "   File: $CurrentInputPath" -ForegroundColor Gray
            
            # Initialize and validate
            $validationError = Initialize-Script -InputPath $CurrentInputPath -MagickPath $MagickPath
            if ($validationError) {
                Write-Host "`n[ERROR] $validationError" -ForegroundColor Red
                return [PSCustomObject]@{
                    InputPath = $CurrentInputPath
                    OutputPath = $null
                    ReportPath = $null
                    OriginalSizeBytes = 0
                    OptimizedSizeBytes = 0
                    SavingsBytes = 0
                    SavingsPercent = 0.0
                    ImagesCropped = 0
                    ImagesOptimized = 0
                    ImagesSkipped = 0
                    TotalImageUsages = 0
                    UniqueImages = 0
                    ExitCode = 1
                    Success = $false
                    ErrorType = 'Validation'
                    Error = $validationError
                }
            }
            
            Initialize-Namespaces
            
            # Create temp directory
            $tempDir = Join-Path ([System.IO.Path]::GetTempPath()) "ppt-optimize-$(New-Guid)"
            New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
            
            # Normalize to full path (avoid 8.3 filename issues)
            $tempDir = (Get-Item -LiteralPath $tempDir).FullName
            
            try {
                # Initialize temp output file tracking (for cleanup in finally)
                $tempPptxPath = $null
                $tempCsvPath = $null
                $copyFailed = $true  # Assume failure until copy succeeds
                
                # Extract (use .NET ZipFile for PS 5.1 compatibility - Expand-Archive only accepts .zip extension)
                Add-Type -Assembly 'System.IO.Compression.FileSystem' -ErrorAction SilentlyContinue
                [System.IO.Compression.ZipFile]::ExtractToDirectory($script:InputFile, $tempDir)
                
                # Scan (canonical slide order, slide dimensions, image usages)
                $scan = Get-PptImageScanData -tempDir $tempDir
                $slides = $scan.Slides
                $usages = $scan.Usages

                Write-Host "[STATS] Found $($usages.Count) image usages" -ForegroundColor Cyan
                $uniqueImages = if ($usages.Count -gt 0) { 
                    ($usages | Select-Object -Property ImagePhysicalPath -Unique).Count 
                } else { 
                    0 
                }
                Write-Host "[STATS] Across $uniqueImages unique images" -ForegroundColor Cyan
                
                if ($usages.Count -eq 0) {
                    Write-Host "`n[INFO] No images found to optimize" -ForegroundColor Yellow
                    
                    # Still create output file and empty CSV
                    Copy-Item $script:InputFile $script:OutputFile
                    
                    $emptyReport = @()
                    $emptyReport | Export-Csv -Path $script:CsvReport -NoTypeInformation -Encoding UTF8
                    
                    Write-Host "`n[DONE] Complete (no changes needed)" -ForegroundColor Green
                    Write-Host "Output: $script:OutputFile" -ForegroundColor Green
                    Write-Host "Report: $script:CsvReport" -ForegroundColor Green
                    
                    $exitCode = 2  # No changes needed
                    
                    # Return summary
                    return [PSCustomObject]@{
                        InputPath = $script:InputFile.ToString()
                        OutputPath = $script:OutputFile
                        ReportPath = $script:CsvReport
                        OriginalSizeBytes = (Get-Item $script:InputFile).Length
                        OptimizedSizeBytes = (Get-Item $script:InputFile).Length
                        SavingsBytes = 0
                        SavingsPercent = 0.0
                        ImagesCropped = 0
                        ImagesOptimized = 0
                        ImagesSkipped = 0
                        TotalImageUsages = 0
                        UniqueImages = 0
                        ExitCode = $exitCode
                        Success = $true
                        ErrorType = 'None'
                    }
                }
                
                # Find Morph pairs
                $morphPairs = Find-MorphPairs -usages $usages
                
                # Phase 1: Normalize crops (only if any cropping enabled)
                if ($script:CropSlides -or $script:CropMastersAndLayouts) {
                    $docsToSave = @{}
                    $slideFiltersActive = ($IncludeSlides -and $IncludeSlides.Count -gt 0) -or ($ExcludeSlides -and $ExcludeSlides.Count -gt 0)
                    
                    foreach ($usage in $usages) {
                        # Skip usages not in scope (same logic as cropping phase)
                        if ($usage.ContextType -eq [ContextType]::Slide -and -not (Test-SlideIncluded -SlideNumber $usage.SlideNumber)) {
                            continue
                        }
                        if ($slideFiltersActive -and ($usage.ContextType -eq [ContextType]::Master -or $usage.ContextType -eq [ContextType]::Layout)) {
                            continue
                        }
                        
                        Invoke-CropNormalization -usage $usage
                        if ($usage.CropNormalized -and $usage.SlideDocument) {
                            $slidePath = Join-Path $tempDir ($usage.PartPath)
                            $docsToSave[$slidePath] = $usage.SlideDocument
                        }
                    }
                    
                    # Save normalized XMLs
                    foreach ($path in $docsToSave.Keys) {
                        Save-XmlDocument -doc $docsToSave[$path] -path $path
                    }
                }
                
                # Phase 2: Cropping
                if ($script:CropSlides -or $script:CropMastersAndLayouts) {
                    Invoke-ImageCropping -usages $usages -tempDir $tempDir -morphPairs $morphPairs
                    
                    # Save XML changes
                    $docsToSave = @{}
                    foreach ($usage in $usages | Where-Object { $_.CropApplied -or $_.CropRemovedNoOp }) {
                        if ($usage.SlideDocument) {
                            $slidePath = Join-Path $tempDir ($usage.PartPath)
                            $docsToSave[$slidePath] = $usage.SlideDocument
                        }
                    }
                    foreach ($path in $docsToSave.Keys) {
                        Save-XmlDocument -doc $docsToSave[$path] -path $path
                    }
                }
                
                # Phase 3: Optimization
                Invoke-ImageOptimization -usages $usages -tempDir $tempDir -morphPairs $morphPairs
                
                # Save XML changes (format conversions update blip refs)
                $docsToSave = @{}
                foreach ($usage in $usages | Where-Object { $_.OptimizationStatus -match '^Converted' }) {
                    if ($usage.SlideDocument) {
                        $slidePath = Join-Path $tempDir ($usage.PartPath)
                        $docsToSave[$slidePath] = $usage.SlideDocument
                        $relativePath = $usage.PartPath
                        Write-Verbose "Marking document for save (format conversion): $relativePath"
                    } else {
                        Write-Warning "Converted usage on $($usage.Location) missing SlideDocument reference"
                    }
                }
                
                if ($docsToSave.Count -gt 0) {
                    Write-Host "[SAVE] Saving $($docsToSave.Count) updated document(s)..." -ForegroundColor Yellow
                    foreach ($path in $docsToSave.Keys) {
                        try {
                            $relativePath = $path -replace [regex]::Escape($tempDir + [System.IO.Path]::DirectorySeparatorChar), ''
                            Write-Verbose "Saving updated document: $relativePath"
                            Save-XmlDocument -doc $docsToSave[$path] -path $path
                        } catch {
                            Write-Error "Failed to save $path : $_"
                            throw
                        }
                    }
                }
                
                # Clear XML references to reduce memory usage (no longer needed after saves)
                foreach ($usage in $usages) {
                    $usage.SlideDocument = $null
                    $usage.BlipElement = $null
                    $usage.XfrmElement = $null
                }
                
                # Phase 3.5: Layout/Master cleanup (if requested)
                if ($script:CleanUnusedLayouts) {
                    Invoke-UnusedLayoutCleanup -tempDir $tempDir
                }

                # Phase 4: Cleanup
                Invoke-OrphanedPartCleanup -tempDir $tempDir
                Invoke-MediaCleanup -tempDir $tempDir
                Update-ContentTypes -tempDir $tempDir
                
                # Validate critical files before repacking
                $criticalFiles = @(
                    '[Content_Types].xml',
                    '_rels/.rels',
                    'ppt/presentation.xml',
                    'ppt/_rels/presentation.xml.rels'
                )
                
                Write-Verbose "Validating $($criticalFiles.Count) critical files before repack..."
                foreach ($file in $criticalFiles) {
                    $filePath = Join-Path $tempDir $file
                    if (-not (Test-Path -LiteralPath $filePath)) {
                        throw "Critical file missing before repack: $file"
                    }
                    Write-Verbose "  [v] $file"
                }
                Write-Verbose "Pre-repack validation passed - all critical files present"
                
                # Phase 5: Repack (to temp file)
                $tempPptxPath = Invoke-Repack -tempDir $tempDir

                # Optional: validate the repacked file with the Open XML SDK.
                if ($script:ValidateOutput) {
                    Test-OutputValidity -pptxPath $tempPptxPath
                }

                # Build the layout/master -> slides reverse index once (after cleanup, so it
                # reflects surviving layouts), then reuse for the report and the summary.
                $partSlideUsage = Get-PartSlideUsage -tempDir $tempDir -slides $slides
                Write-ThinkCellSummary -tempDir $tempDir -usages $usages -partSlideUsage $partSlideUsage

                # Generate report (to temp file)
                $tempCsvPath = Export-CsvReport -usages $usages -partSlideUsage $partSlideUsage
                
                # Final statistics (calculated from temp file)
                $originalSize = (Get-Item $script:InputFile).Length
                $optimizedSize = (Get-Item $tempPptxPath).Length
                $savedBytes = $originalSize - $optimizedSize
                $savedPercent = ($savedBytes / $originalSize) * 100
                
                $croppedCount = @($usages | Where-Object { $_.CropApplied } | Select-Object -Property OriginalFileName -Unique).Count
                $optimizedCount = @($usages | Where-Object { $_.OptimizationStatus -match '^(Optimized|Converted)' } | Select-Object -Property OriginalFileName -Unique).Count
                $skippedCount = @($usages | Where-Object { $_.OptimizationStatus -match '^Skipped' -or $_.OptimizationStatus -eq 'NoChangeNeeded' }).Count
                
                Write-Host "`n[DONE] Optimization Complete" -ForegroundColor Green
                Write-Host "[STATS] Images cropped: $croppedCount" -ForegroundColor Green
                Write-Host "[STATS] Images optimized: $optimizedCount" -ForegroundColor Green
                Write-Host "`n[SIZE] File size:" -ForegroundColor Green
                Write-Host "  Original: $((Format-ByteSize $originalSize))" -ForegroundColor Green
                Write-Host "  Optimized: $((Format-ByteSize $optimizedSize))" -ForegroundColor Green
                Write-Host "  Saved: $(Format-ByteSize $savedBytes) ($($savedPercent.ToString('F1'))%)" -ForegroundColor Green
                
                # Phase 6: Copy to final destinations
                Write-Host "`n[SAVE] Saving output files..." -ForegroundColor Yellow
                $copyResult = Copy-OutputFilesToFinal `
                    -tempPptxPath $tempPptxPath `
                    -tempCsvPath $tempCsvPath `
                    -finalPptxPath $script:OutputFile `
                    -finalCsvPath $script:CsvReport
                
                # Handle any copy failures
                $copyFailed = -not $copyResult.PptxSuccess -or -not $copyResult.CsvSuccess
                if ($copyFailed) {
                    Write-Host "`n[WARN]  Some output files could not be saved (files may be in use):" -ForegroundColor Yellow
                    if (-not $copyResult.PptxSuccess) {
                        Write-Host "  Optimized PPTX saved at: $tempPptxPath" -ForegroundColor Cyan
                        Write-Host "  -> Close the target file and copy manually, or re-run the script." -ForegroundColor Gray
                    }
                    if (-not $copyResult.CsvSuccess) {
                        Write-Host "  CSV report saved at: $tempCsvPath" -ForegroundColor Cyan
                        Write-Host "  -> Close the target file and copy manually, or re-run the script." -ForegroundColor Gray
                    }
                    $exitCode = 1
                } else {
                    $exitCode = 0
                }
                
                # Return summary object
                return [PSCustomObject]@{
                    InputPath = $script:InputFile.ToString()
                    OutputPath = if ($copyResult.PptxSuccess) { $script:OutputFile } else { $tempPptxPath }
                    ReportPath = if ($copyResult.CsvSuccess) { $script:CsvReport } else { $tempCsvPath }
                    OriginalSizeBytes = $originalSize
                    OptimizedSizeBytes = $optimizedSize
                    SavingsBytes = $savedBytes
                    SavingsPercent = [Math]::Round($savedPercent, 2)
                    ImagesCropped = $croppedCount
                    ImagesOptimized = $optimizedCount
                    ImagesSkipped = $skippedCount
                    TotalImageUsages = $usages.Count
                    UniqueImages = $uniqueImages
                    ExitCode = $exitCode
                    Success = (-not $copyFailed)
                    ErrorType = 'None'
                }
                
            } finally {
                # Cleanup extraction temp directory (always)
                if (Test-Path -LiteralPath $tempDir) {
                    Remove-Item -LiteralPath $tempDir -Recurse -Force -ErrorAction SilentlyContinue
                }
                
                # Cleanup temp output files only if copy succeeded
                if (-not $copyFailed) {
                    if ($tempPptxPath -and (Test-Path -LiteralPath $tempPptxPath)) {
                        Remove-Item -LiteralPath $tempPptxPath -Force -ErrorAction SilentlyContinue
                    }
                    if ($tempCsvPath -and (Test-Path -LiteralPath $tempCsvPath)) {
                        Remove-Item -LiteralPath $tempCsvPath -Force -ErrorAction SilentlyContinue
                    }
                } elseif ($tempPptxPath -or $tempCsvPath) {
                    # If we have temp files but copy failed (or error occurred), tell user where they are
                    $hasTempPptx = $tempPptxPath -and (Test-Path -LiteralPath $tempPptxPath)
                    $hasTempCsv = $tempCsvPath -and (Test-Path -LiteralPath $tempCsvPath)
                    if ($hasTempPptx -or $hasTempCsv) {
                        Write-Host "`n[SAVE] Temporary files preserved:" -ForegroundColor Yellow
                        if ($hasTempPptx) {
                            Write-Host "  PPTX: $tempPptxPath" -ForegroundColor Cyan
                        }
                        if ($hasTempCsv) {
                            Write-Host "  CSV:  $tempCsvPath" -ForegroundColor Cyan
                        }
                    }
                }
            }
            
        } catch {
            Write-Host "`n[ERROR] Unexpected error: $_" -ForegroundColor Red
            if (Test-VerboseMode) {
                Write-Host "`nStack trace:" -ForegroundColor DarkGray
                Write-Host $_.ScriptStackTrace -ForegroundColor DarkGray
            } else {
                Write-Host "   Run with -Verbose for detailed error information." -ForegroundColor Gray
            }
            $exitCode = 1
            
            return [PSCustomObject]@{
                InputPath = $CurrentInputPath
                OutputPath = $null
                ReportPath = $null
                OriginalSizeBytes = 0
                OptimizedSizeBytes = 0
                SavingsBytes = 0
                SavingsPercent = 0.0
                ImagesCropped = 0
                ImagesOptimized = 0
                ImagesSkipped = 0
                TotalImageUsages = 0
                UniqueImages = 0
                ExitCode = $exitCode
                Success = $false
                ErrorType = 'Processing'
                Error = $_.ToString()
            }
        }
    }

    #endregion
