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
    [switch]$PassThru
)

    begin {
        Set-StrictMode -Version Latest
        $ErrorActionPreference = 'Stop'

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
        if ($PassThru) {
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
    
    # Exit with appropriate code
    # 0 = Success with changes
    # 1 = Error occurred
    # 2 = Success but no changes needed
    if ($Host.Name -ne 'ConsoleHost' -or $script:BatchMode) {
        # Don't exit in ISE or batch mode - let caller handle
    } else {
        exit $finalExitCode
    }
    }
}
