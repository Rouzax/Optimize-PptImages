BeforeAll {
    Import-Module "$PSScriptRoot/../../Optimize-PptImages.psd1" -Force
    $script:fixture = Join-Path $PSScriptRoot '../fixtures/content-id-sample.pptx'
}

Describe 'Optimize console output routing' {

    It 'emits no routine [OK]/[CONV] lines to Write-Host on a non-verbose run' {
        InModuleScope Optimize-PptImages -Parameters @{ fixture = $script:fixture } {
            param($fixture)

            $script:OptimizeSlides             = $true
            $script:OptimizeMastersAndLayouts  = $true
            $script:CropSlides                 = $true
            $script:CropMastersAndLayouts      = $true
            $script:CleanUnusedLayouts         = $false
            $script:ValidateOutput             = $false

            # Outer-scope variables required by dynamic scoping under StrictMode.
            $MagickPath                 = $null
            $OutputPath                 = $null
            $CsvReportPath              = $null
            $IncludeSlides              = @()
            $ExcludeSlides              = @()
            $HeadroomFactor             = 2.0
            $JpegQuality                = 95
            $TransparencyThresholdPercent = 0.1
            $MinSavingsPercent          = 0.0

            Mock Write-Host {}

            $out = Join-Path ([System.IO.Path]::GetTempPath()) "console-test-$([guid]::NewGuid()).pptx"
            try {
                $result = Invoke-PptOptimization -CurrentInputPath $fixture
                $result.Success | Should -BeTrue

                # Routine per-image success lines must NOT reach Write-Host on a normal run.
                # Optimization [OK] lines always start with three leading spaces; this excludes
                # [OK] Repacked and [OK] CSV report lines which have no leading spaces.
                Should -Invoke Write-Host -ParameterFilter {
                    $Object -match '^\s+\[OK\]|^\s*\[CONV\]'
                } -Times 0 -Exactly

                # The optimization [STATS] summary must always reach Write-Host.
                Should -Invoke Write-Host -ParameterFilter {
                    $Object -match '\[STATS\] Optimization phase complete'
                } -Times 1 -Exactly -Because '[STATS] optimization summary must always be visible'
            } finally {
                foreach ($p in @($result.OutputPath, $result.ReportPath, $out)) {
                    if ($p -and (Test-Path -LiteralPath $p)) {
                        Remove-Item -LiteralPath $p -Force -ErrorAction SilentlyContinue
                    }
                }
            }
        }
    }

    It 'emits routine [OK]/[CONV] lines to the verbose stream under -Verbose' {
        InModuleScope Optimize-PptImages -Parameters @{ fixture = $script:fixture } {
            param($fixture)

            $script:OptimizeSlides             = $true
            $script:OptimizeMastersAndLayouts  = $true
            $script:CropSlides                 = $true
            $script:CropMastersAndLayouts      = $true
            $script:CleanUnusedLayouts         = $false
            $script:ValidateOutput             = $false

            $MagickPath                 = $null
            $OutputPath                 = $null
            $CsvReportPath              = $null
            $IncludeSlides              = @()
            $ExcludeSlides              = @()
            $HeadroomFactor             = 2.0
            $JpegQuality                = 95
            $TransparencyThresholdPercent = 0.1
            $MinSavingsPercent          = 0.0

            Mock Write-Verbose {}

            $out = Join-Path ([System.IO.Path]::GetTempPath()) "console-verbose-$([guid]::NewGuid()).pptx"
            try {
                # Force VerbosePreference so Write-Verbose calls are honoured inside the module.
                $VerbosePreference = 'Continue'
                $result = Invoke-PptOptimization -CurrentInputPath $fixture
                $result.Success | Should -BeTrue

                # At least one routine optimization success line must reach Write-Verbose.
                # The pattern anchors on leading spaces (3 spaces) to match optimization lines only.
                Should -Invoke Write-Verbose -ParameterFilter {
                    $Message -match '^\s+\[OK\]|^\s*\[CONV\]'
                } -Times 1 -Because 'routine lines must appear on the verbose stream'
            } finally {
                foreach ($p in @($result.OutputPath, $result.ReportPath, $out)) {
                    if ($p -and (Test-Path -LiteralPath $p)) {
                        Remove-Item -LiteralPath $p -Force -ErrorAction SilentlyContinue
                    }
                }
            }
        }
    }
}
