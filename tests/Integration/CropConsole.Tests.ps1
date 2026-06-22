BeforeAll {
    Import-Module "$PSScriptRoot/../../Optimize-PptImages.psd1" -Force
    $script:fixture = Join-Path $PSScriptRoot '../fixtures/crop-sample.pptx'
}

Describe 'Crop console output routing' {

    It 'emits no routine [CROP]/[NOTE] lines to Write-Host on a non-verbose run' {
        InModuleScope Optimize-PptImages -Parameters @{ fixture = $script:fixture } {
            param($fixture)

            $script:OptimizeSlides             = $true
            $script:OptimizeMastersAndLayouts  = $true
            $script:CropSlides                 = $true
            $script:CropMastersAndLayouts      = $true
            $script:CleanUnusedLayouts         = $false
            $script:ValidateOutput             = $false

            # Outer-scope variables required by dynamic scoping under StrictMode.
            $MagickPath                   = $null
            $OutputPath                   = $null
            $CsvReportPath                = $null
            $IncludeSlides                = @()
            $ExcludeSlides                = @()
            $HeadroomFactor               = 2.0
            $JpegQuality                  = 95
            $TransparencyThresholdPercent = 0.1
            $MinSavingsPercent            = 0.0

            Mock Write-Host {}

            try {
                $result = Invoke-PptOptimization -CurrentInputPath $fixture
                $result.Success | Should -BeTrue

                # Routine per-usage crop success lines must NOT reach Write-Host on a normal run.
                Should -Invoke Write-Host -ParameterFilter {
                    $Object -match '^\s+\[CROP\] Cropped'
                } -Times 0 -Exactly -Because 'routine [CROP] success lines must be verbose-gated'

                # No-op crop removal lines must NOT reach Write-Host on a normal run.
                Should -Invoke Write-Host -ParameterFilter {
                    $Object -match '\[NOTE\] Removed no-op crop'
                } -Times 0 -Exactly -Because 'routine [NOTE] no-op lines must be verbose-gated'

                # The cropping [STATS] summary must always reach Write-Host.
                Should -Invoke Write-Host -ParameterFilter {
                    $Object -match '\[STATS\] Cropping phase complete'
                } -Times 1 -Exactly -Because '[STATS] cropping summary must always be visible'
            } finally {
                foreach ($p in @($result.OutputPath, $result.ReportPath)) {
                    if ($p -and (Test-Path -LiteralPath $p)) {
                        Remove-Item -LiteralPath $p -Force -ErrorAction SilentlyContinue
                    }
                }
            }
        }
    }

    It 'emits routine [CROP] lines to the verbose stream under -Verbose' {
        InModuleScope Optimize-PptImages -Parameters @{ fixture = $script:fixture } {
            param($fixture)

            $script:OptimizeSlides             = $true
            $script:OptimizeMastersAndLayouts  = $true
            $script:CropSlides                 = $true
            $script:CropMastersAndLayouts      = $true
            $script:CleanUnusedLayouts         = $false
            $script:ValidateOutput             = $false

            $MagickPath                   = $null
            $OutputPath                   = $null
            $CsvReportPath                = $null
            $IncludeSlides                = @()
            $ExcludeSlides                = @()
            $HeadroomFactor               = 2.0
            $JpegQuality                  = 95
            $TransparencyThresholdPercent = 0.1
            $MinSavingsPercent            = 0.0

            Mock Write-Verbose {}

            try {
                # Force VerbosePreference so Write-Verbose calls are honoured inside the module.
                $VerbosePreference = 'Continue'
                $result = Invoke-PptOptimization -CurrentInputPath $fixture
                $result.Success | Should -BeTrue

                # At least one routine crop success line must reach Write-Verbose.
                Should -Invoke Write-Verbose -ParameterFilter {
                    $Message -match '^\s+\[CROP\] Cropped'
                } -Times 1 -Because 'routine [CROP] lines must appear on the verbose stream'
            } finally {
                foreach ($p in @($result.OutputPath, $result.ReportPath)) {
                    if ($p -and (Test-Path -LiteralPath $p)) {
                        Remove-Item -LiteralPath $p -Force -ErrorAction SilentlyContinue
                    }
                }
            }
        }
    }
}
