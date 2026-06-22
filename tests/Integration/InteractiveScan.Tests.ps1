BeforeAll {
    Import-Module "$PSScriptRoot/../../Optimize-PptImages.psd1" -Force
    $generator = Join-Path $PSScriptRoot '../fixtures/New-SampleDeck.ps1'
    $script:deck = Join-Path ([System.IO.Path]::GetTempPath()) "scan-core-$([guid]::NewGuid()).pptx"
    & $generator -OutputPath $script:deck | Out-Null
}

AfterAll {
    if (Test-Path -LiteralPath $script:deck) { Remove-Item -LiteralPath $script:deck -Force }
}

Describe 'Get-PptImageScanData' {
    It 'returns usages and slides for an extracted package' {
        InModuleScope Optimize-PptImages -Parameters @{ deck = $script:deck } {
            param($deck)
            Initialize-Namespaces
            $script:MagickExe = Find-ImageMagick -ProvidedPath $null
            $tempDir = Join-Path ([System.IO.Path]::GetTempPath()) "scan-core-x-$([guid]::NewGuid())"
            New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
            try {
                Add-Type -Assembly 'System.IO.Compression.FileSystem' -ErrorAction SilentlyContinue
                [System.IO.Compression.ZipFile]::ExtractToDirectory($deck, $tempDir)
                $result = Get-PptImageScanData -tempDir $tempDir
                $result.Slides.Count | Should -BeGreaterThan 0
                $result.Usages.Count | Should -BeGreaterThan 0
            } finally {
                Remove-Item -LiteralPath $tempDir -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
    }
}

Describe 'Get-PptImageScan' {
    It 'returns a live context with enriched source dimensions' {
        InModuleScope Optimize-PptImages -Parameters @{ deck = $script:deck } {
            param($deck)
            $ctx = Get-PptImageScan -InputPath $deck -MagickPath $null
            try {
                $ctx.Findings.Slides.UsageCount | Should -BeGreaterThan 0
                $ctx.TempDir | Should -Not -BeNullOrEmpty
                (Test-Path -LiteralPath $ctx.TempDir) | Should -BeTrue   # not cleaned: caller owns it
                ($ctx.Usages | Where-Object { $_.SourceWidthPx -gt 0 }).Count | Should -BeGreaterThan 0
            } finally {
                if ($ctx.TempDir -and (Test-Path -LiteralPath $ctx.TempDir)) {
                    Remove-Item -LiteralPath $ctx.TempDir -Recurse -Force -ErrorAction SilentlyContinue
                }
            }
        }
    }

    It 'throws a clear error for a missing file' {
        InModuleScope Optimize-PptImages {
            { Get-PptImageScan -InputPath 'does-not-exist.pptx' -MagickPath $null } |
                Should -Throw -ExpectedMessage '*not found*'
        }
    }
}

Describe 'Invoke-PptOptimization -PreparedScan reuse' {
    It 'optimizes using a handed-over scan without owning the temp dir' {
        InModuleScope Optimize-PptImages -Parameters @{ deck = $script:deck } {
            param($deck)
            $script:OptimizeSlides = $true
            $script:OptimizeMastersAndLayouts = $false
            $script:CropSlides = $false
            $script:CropMastersAndLayouts = $false
            $script:CleanUnusedLayouts = $false
            $script:ValidateOutput = $false

            # Outer-scope variables that Initialize-Script and optimization functions
            # read via dynamic scoping. Set defaults matching Invoke-PptImageOptimization.
            $MagickPath = $null
            $OutputPath = $null
            $CsvReportPath = $null
            $IncludeSlides = @()
            $ExcludeSlides = @()
            $HeadroomFactor = 2.0
            $JpegQuality = 95
            $TransparencyThresholdPercent = 0.1
            $MinSavingsPercent = 0.0

            $ctx = Get-PptImageScan -InputPath $deck -MagickPath $null
            try {
                $result = Invoke-PptOptimization -CurrentInputPath $deck -PreparedScan $ctx
                $result.Success | Should -BeTrue
                # Reuse contract: the function did not delete the wizard-owned temp dir.
                (Test-Path -LiteralPath $ctx.TempDir) | Should -BeTrue
            } finally {
                if ($ctx.TempDir -and (Test-Path -LiteralPath $ctx.TempDir)) {
                    Remove-Item -LiteralPath $ctx.TempDir -Recurse -Force -ErrorAction SilentlyContinue
                }
                if ($result -and $result.OutputPath -and (Test-Path -LiteralPath $result.OutputPath)) {
                    Remove-Item -LiteralPath $result.OutputPath -Force -ErrorAction SilentlyContinue
                }
                if ($result -and $result.ReportPath -and (Test-Path -LiteralPath $result.ReportPath)) {
                    Remove-Item -LiteralPath $result.ReportPath -Force -ErrorAction SilentlyContinue
                }
            }
        }
    }
}

Describe 'Non-interactive run front-loads source dimensions' {
    It 'populates SourceWidthPx in the CSV report for a normal optimize run' {
        InModuleScope Optimize-PptImages -Parameters @{ deck = $script:deck } {
            param($deck)
            $script:OptimizeSlides = $true
            $script:OptimizeMastersAndLayouts = $false
            $script:CropSlides = $false
            $script:CropMastersAndLayouts = $false
            $script:CleanUnusedLayouts = $false
            $script:ValidateOutput = $false

            # Outer-scope variables that Initialize-Script reads via dynamic scoping.
            $MagickPath = $null
            $OutputPath = $null
            $CsvReportPath = $null
            $IncludeSlides = @()
            $ExcludeSlides = @()
            $HeadroomFactor = 2.0
            $JpegQuality = 95
            $TransparencyThresholdPercent = 0.1
            $MinSavingsPercent = 0.0

            $result = Invoke-PptOptimization -CurrentInputPath $deck
            try {
                $result.Success | Should -BeTrue
                $rows = Import-Csv -LiteralPath $result.ReportPath
                # At least one image row reports a real source width (dims flowed through).
                ($rows | Where-Object { [int]($_.SourceWidthPx) -gt 0 }).Count | Should -BeGreaterThan 0
            } finally {
                foreach ($p in @($result.OutputPath, $result.ReportPath)) {
                    if ($p -and (Test-Path -LiteralPath $p)) { Remove-Item -LiteralPath $p -Force -ErrorAction SilentlyContinue }
                }
            }
        }
    }
}
