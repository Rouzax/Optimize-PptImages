function Invoke-PptImageOptimization {
    [CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='Medium')]
param(
    [Parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
    [Alias('FullName', 'Path')]
    [ValidateScript({
        if (-not (Test-Path $_)) {
            throw "Input file not found: $_"
        }
        if ($_ -notmatch '\.(pptx|potx)$') {
            throw "Input file must be .pptx or .potx"
        }
        $true
    })]
    [string]$InputPath,

    [Parameter(Mandatory=$false)]
    [string]$OutputPath,

    [Parameter(Mandatory=$false)]
    [string]$MagickPath,

    [Parameter(Mandatory=$false)]
    [switch]$OptimizeSlides,

    [Parameter(Mandatory=$false)]
    [switch]$OptimizeMastersAndLayouts,

    [Parameter(Mandatory=$false)]
    [switch]$CropSlides,

    [Parameter(Mandatory=$false)]
    [switch]$CropMastersAndLayouts,

    [Parameter(Mandatory=$false)]
    [switch]$CleanUnusedLayouts,

    [Parameter(Mandatory=$false)]
    [switch]$ValidateOutput,

    [Parameter(Mandatory=$false)]
    [switch]$All,

    [Parameter(Mandatory=$false)]
    [ValidateRange(0.5, 4.0)]
    [double]$HeadroomFactor = 2.0,

    [Parameter(Mandatory=$false)]
    [ValidateRange(1, 100)]
    [int]$JpegQuality = 95,

    [Parameter(Mandatory=$false)]
    [ValidateRange(0.0, 100.0)]
    [double]$TransparencyThresholdPercent = 0.1,

    [Parameter(Mandatory=$false)]
    [ValidateRange(0.0, 50.0)]
    [double]$MinSavingsPercent = 0.0,

    [Parameter(Mandatory=$false)]
    [int[]]$IncludeSlides,

    [Parameter(Mandatory=$false)]
    [int[]]$ExcludeSlides,

    [Parameter(Mandatory=$false)]
    [string]$CsvReportPath,

    [Parameter(Mandatory=$false)]
    [switch]$PassThru,

    [Parameter(Mandatory=$false)]
    [switch]$Interactive
)

    begin {
        Set-StrictMode -Version Latest
        $ErrorActionPreference = 'Stop'

        # Assume failure until the end block sets the real exit code.
        # This prevents the wrapper from reading 0 if the function crashes before end{} completes.
        $script:LastFinalExitCode = 1

        # Guard: -Interactive only works for a single file
        if ($Interactive -and $MyInvocation.ExpectingInput) {
            throw "Interactive mode supports a single file only. Run without pipeline/batch input, or drop -Interactive."
        }

        # Track batch processing results
        $script:BatchResults = [System.Collections.Generic.List[PSCustomObject]]::new()
        $script:BatchMode = $false
        $script:FileCount = 0

        # Expand -All flag into individual flags
        if ($All) {
            $script:OptimizeSlides = $true
            $script:OptimizeMastersAndLayouts = $true
            $script:CropSlides = $true
            $script:CropMastersAndLayouts = $true
        } else {
            $script:OptimizeSlides = $OptimizeSlides.IsPresent
            $script:OptimizeMastersAndLayouts = $OptimizeMastersAndLayouts.IsPresent
            $script:CropSlides = $CropSlides.IsPresent
            $script:CropMastersAndLayouts = $CropMastersAndLayouts.IsPresent
        }
        $script:CleanUnusedLayouts = $CleanUnusedLayouts.IsPresent
        $script:ValidateOutput = $ValidateOutput.IsPresent
    }

    process {
    $script:FileCount++
    
    # Detect batch mode
    if ($script:FileCount -gt 1 -or $MyInvocation.ExpectingInput) {
        $script:BatchMode = $true
    }
    
    if ($script:BatchMode -and $script:FileCount -gt 1) {
        Write-Host ("`n" + ("=" * 60)) -ForegroundColor DarkGray
    }
    
    if ($Interactive) {
        if ($script:FileCount -gt 1 -or $MyInvocation.ExpectingInput) {
            throw "Interactive mode supports a single file only. Run without pipeline/batch input, or drop -Interactive."
        }

        $context = Get-PptImageScan -InputPath $InputPath -MagickPath $MagickPath
        try {
            $defaults = [PSCustomObject]@{
                OptimizeSlides               = $script:OptimizeSlides
                OptimizeMastersAndLayouts    = $script:OptimizeMastersAndLayouts
                CropSlides                   = $script:CropSlides
                CropMastersAndLayouts        = $script:CropMastersAndLayouts
                CleanUnusedLayouts           = $script:CleanUnusedLayouts
                HeadroomFactor               = $HeadroomFactor
                JpegQuality                  = $JpegQuality
                OutputPath                   = if ($OutputPath) { $OutputPath } else { '' }
                TransparencyThresholdPercent = $TransparencyThresholdPercent
                MinSavingsPercent            = $MinSavingsPercent
                IncludeSlides                = if ($IncludeSlides) { @($IncludeSlides) } else { @() }
                ExcludeSlides                = if ($ExcludeSlides) { @($ExcludeSlides) } else { @() }
            }

            $wiz = Show-PptInteractiveWizard -InputPath $InputPath -Findings $context.Findings -Defaults $defaults

            if ($wiz.Action -ne 'Run') {
                if ($wiz.Action -eq 'Copy') {
                    Write-Host ''
                    Write-Host (Format-PptCommandLine -InputPath $InputPath -Answers $wiz.Answers) -ForegroundColor Green
                }
                return
            }

            $a = $wiz.Answers
            $script:OptimizeSlides            = $a.OptimizeSlides
            $script:OptimizeMastersAndLayouts = $a.OptimizeMastersAndLayouts
            $script:CropSlides                = $a.CropSlides
            $script:CropMastersAndLayouts     = $a.CropMastersAndLayouts
            $script:CleanUnusedLayouts        = $a.CleanUnusedLayouts
            $HeadroomFactor               = $a.HeadroomFactor
            $JpegQuality                  = $a.JpegQuality
            $TransparencyThresholdPercent = $a.TransparencyThresholdPercent
            $MinSavingsPercent            = $a.MinSavingsPercent
            if ($a.OutputPath)    { $OutputPath = $a.OutputPath }
            if ($a.IncludeSlides -and $a.IncludeSlides.Count -gt 0) { $IncludeSlides = $a.IncludeSlides }
            if ($a.ExcludeSlides -and $a.ExcludeSlides.Count -gt 0) { $ExcludeSlides = $a.ExcludeSlides }

            Write-Host "`n[FILE] Processing file $($script:FileCount): $(Split-Path $InputPath -Leaf)" -ForegroundColor Cyan
            $result = Invoke-PptOptimization -CurrentInputPath $InputPath -PreparedScan $context
            $script:BatchResults.Add($result)
        } finally {
            if ($context.TempDir -and (Test-Path -LiteralPath $context.TempDir)) {
                Remove-Item -LiteralPath $context.TempDir -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
        return
    }

    Write-Host "`n[FILE] Processing file $($script:FileCount): $(Split-Path $InputPath -Leaf)" -ForegroundColor Cyan

    $result = Invoke-PptOptimization -CurrentInputPath $InputPath
    $script:BatchResults.Add($result)
    }

    end {
    # Output results
    if ($script:BatchMode -and $script:BatchResults.Count -gt 1) {
        Write-Host ("`n" + ("=" * 60)) -ForegroundColor DarkGray
        Write-Host "[STATS] BATCH SUMMARY" -ForegroundColor Cyan
        Write-Host ("=" * 60) -ForegroundColor DarkGray
        
        $totalOriginal = ($script:BatchResults | Measure-Object -Property OriginalSizeBytes -Sum).Sum
        $totalOptimized = ($script:BatchResults | Measure-Object -Property OptimizedSizeBytes -Sum).Sum
        $totalSaved = $totalOriginal - $totalOptimized
        $totalSavedPercent = if ($totalOriginal -gt 0) { ($totalSaved / $totalOriginal) * 100 } else { 0 }
        
        $successCount = @($script:BatchResults | Where-Object { $_.Success }).Count
        $failCount = @($script:BatchResults | Where-Object { -not $_.Success }).Count
        
        Write-Host "Files processed: $($script:BatchResults.Count)" -ForegroundColor Green
        Write-Host "  Successful: $successCount" -ForegroundColor Green
        if ($failCount -gt 0) {
            Write-Host "  Failed: $failCount" -ForegroundColor Red
        }
        Write-Host ""
        Write-Host "Total savings:" -ForegroundColor Green
        Write-Host "  Original: $(Format-ByteSize $totalOriginal)" -ForegroundColor Green
        Write-Host "  Optimized: $(Format-ByteSize $totalOptimized)" -ForegroundColor Green
        Write-Host "  Saved: $(Format-ByteSize $totalSaved) ($($totalSavedPercent.ToString('F1'))%)" -ForegroundColor Green
        
        # In batch mode, output results if -PassThru specified
        if ($PassThru) {
            $script:BatchResults
        }
    } else {
        # Single file mode - only output result if -PassThru specified
        if ($PassThru -and $script:BatchResults.Count -gt 0) {
            $result = $script:BatchResults[0]
            if ($result.ErrorType -ne 'Validation') {
                $result
            }
        }
    }
    
    # Set exit code based on results
    $finalExitCode = 0
    foreach ($result in $script:BatchResults) {
        if ($result.ExitCode -gt $finalExitCode) {
            $finalExitCode = $result.ExitCode
        }
    }
    
    # Store the final exit code so the script wrapper can propagate it.
    # Do not call exit here; calling exit inside a module function terminates
    # the caller's runspace, which breaks in-process test harnesses like Pester.
    # The thin Optimize-PptImages.ps1 wrapper reads this variable and calls exit.
    $script:LastFinalExitCode = $finalExitCode
    }
}
