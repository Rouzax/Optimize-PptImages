    #region Validation and Initialization

    function Initialize-Script {
        param(
            [string]$InputPath,
            [string]$MagickPath
        )

        # Returns: $null on success, or error message string on validation failure

        # Resolve paths
        $script:InputFile = Resolve-Path -LiteralPath $InputPath

        if (-not $OutputPath) {
            $dir = Split-Path $script:InputFile -Parent
            $base = [System.IO.Path]::GetFileNameWithoutExtension($script:InputFile)
            $ext = [System.IO.Path]::GetExtension($script:InputFile)
            $script:OutputFile = Join-Path $dir "$base.optimized$ext"
        } else {
            $script:OutputFile = $OutputPath
        }

        if (-not $CsvReportPath) {
            $dir = Split-Path $script:OutputFile -Parent
            $base = [System.IO.Path]::GetFileNameWithoutExtension($script:OutputFile)
            $script:CsvReport = Join-Path $dir "$base.opt-report.csv"
        } else {
            $script:CsvReport = $CsvReportPath
        }

        if (Test-VerboseMode) {
            Write-Verbose "Configuration:"
            Write-Verbose "  Input:  $script:InputFile"
            Write-Verbose "  Output: $script:OutputFile"
            Write-Verbose "  Report: $script:CsvReport"
            Write-Verbose "Operations enabled:"
            Write-Verbose "  OptimizeSlides: $script:OptimizeSlides"
            Write-Verbose "  OptimizeMastersAndLayouts: $script:OptimizeMastersAndLayouts"
            Write-Verbose "  CropSlides: $script:CropSlides"
            Write-Verbose "  CropMastersAndLayouts: $script:CropMastersAndLayouts"
            Write-Verbose "Parameters:"
            Write-Verbose "  HeadroomFactor: $HeadroomFactor"
            Write-Verbose "  JpegQuality: $JpegQuality"
            Write-Verbose "  TransparencyThreshold: $TransparencyThresholdPercent%"
            Write-Verbose "  MinSavingsPercent: $MinSavingsPercent%"
            if ($IncludeSlides) {
                Write-Verbose "  IncludeSlides: $($IncludeSlides -join ', ')"
            }
            if ($ExcludeSlides) {
                Write-Verbose "  ExcludeSlides: $($ExcludeSlides -join ', ')"
            }
        }

        # Warn about flags that are ignored when slide filters are active
        $slideFiltersActive = ($IncludeSlides -and $IncludeSlides.Count -gt 0) -or ($ExcludeSlides -and $ExcludeSlides.Count -gt 0)
        if ($slideFiltersActive) {
            $ignoredFlags = @()
            if ($script:CropMastersAndLayouts) { $ignoredFlags += '-CropMastersAndLayouts' }
            if ($script:OptimizeMastersAndLayouts) { $ignoredFlags += '-OptimizeMastersAndLayouts' }

            if ($ignoredFlags.Count -gt 0) {
                Write-Host "[WARN]  Note: $($ignoredFlags -join ', ') ignored when slide filters are active" -ForegroundColor DarkYellow
                Write-Host "   (Masters/layouts affect all slides, not just filtered selection)" -ForegroundColor DarkYellow
            }
        }

        # Check if output files are accessible (fail early)
        if (Test-FileInUse -Path $script:OutputFile) {
            return "Output file is in use: $($script:OutputFile)`n  -> Please close the file and try again."
        }
        if (Test-FileInUse -Path $script:CsvReport) {
            return "CSV report file is in use: $($script:CsvReport)`n  -> Please close the file (Excel?) and try again."
        }

        # Find ImageMagick
        $script:MagickExe = Find-ImageMagick -ProvidedPath $MagickPath
        if (-not $script:MagickExe) {
            return "ImageMagick not found.`n  -> Install ImageMagick and ensure 'magick' is in PATH, or use -MagickPath parameter."
        }

        Write-Verbose "ImageMagick: $script:MagickExe"

        # Return $null to indicate success
        return $null
    }

    #endregion
