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
