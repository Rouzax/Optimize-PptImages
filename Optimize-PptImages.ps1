#Requires -Version 7.0
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
    Import-Module $PSScriptRoot/Optimize-PptImages.psd1 -Force
    $wrapped = Get-Command Invoke-PptImageOptimization
    $scriptCmd = { & $wrapped @PSBoundParameters }
    $steppablePipeline = $scriptCmd.GetSteppablePipeline($MyInvocation.CommandOrigin)
    $steppablePipeline.Begin($PSCmdlet)
}
process { $steppablePipeline.Process($_) }
end { $steppablePipeline.End() }
