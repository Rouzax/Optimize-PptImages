BeforeAll {
    Import-Module "$PSScriptRoot/../../Optimize-PptImages.psd1" -Force
}

# Dedicated tests for Update-OptimizeProbesParallel probe-format gating.
# The main probe-function tests (opaque PNG, JPEG quality, shared-path fan-out,
# WARNING-noise filtering) live in OptimizeJob.Tests.ps1.

Describe 'Update-OptimizeProbesParallel - BMP not probed for transparency' {
    It 'does not set EffectiveTransparencyPercent for a BMP (BMP excluded from transparency probe)' {
        InModuleScope Optimize-PptImages {
            $script:MagickExe = Find-ImageMagick -ProvidedPath $null
            $dir = Join-Path ([System.IO.Path]::GetTempPath()) "optprobe-bmp-$([guid]::NewGuid())"
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
            try {
                $bmpPath = Join-Path $dir 'test.bmp'
                & $script:MagickExe -size 40x30 'xc:red' $bmpPath

                $u = [ImageUsage]::new()
                $u.ImagePhysicalPath = $bmpPath
                $u.OriginalFileName  = 'test.bmp'
                # Set a sentinel value so we can tell if the probe overwrote it.
                $u.EffectiveTransparencyPercent = -99.0
                $u.SourceJpegQuality = -2

                $usages = [System.Collections.Generic.List[ImageUsage]]::new()
                $usages.Add($u)

                Update-OptimizeProbesParallel -usages $usages

                # BMP is excluded from the transparency probe; the sentinel must be unchanged.
                $usages[0].EffectiveTransparencyPercent | Should -Be ([double]-99.0)
                # BMP is not a JPEG either; quality sentinel must also be unchanged.
                $usages[0].SourceJpegQuality | Should -Be -2
            } finally {
                Remove-Item -LiteralPath $dir -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
    }

    It 'does set EffectiveTransparencyPercent for a GIF (non-BMP convertible is probed)' {
        InModuleScope Optimize-PptImages {
            $script:MagickExe = Find-ImageMagick -ProvidedPath $null
            $dir = Join-Path ([System.IO.Path]::GetTempPath()) "optprobe-gif-$([guid]::NewGuid())"
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
            try {
                $gifPath = Join-Path $dir 'test.gif'
                # Create a simple opaque GIF.
                & $script:MagickExe -size 40x30 'xc:red' $gifPath

                $u = [ImageUsage]::new()
                $u.ImagePhysicalPath = $gifPath
                $u.OriginalFileName  = 'test.gif'
                $u.EffectiveTransparencyPercent = -99.0
                $u.SourceJpegQuality = -2

                $usages = [System.Collections.Generic.List[ImageUsage]]::new()
                $usages.Add($u)

                Update-OptimizeProbesParallel -usages $usages

                # GIF is probed; value must have been overwritten from the sentinel.
                $usages[0].EffectiveTransparencyPercent | Should -Not -Be -99.0
                $usages[0].EffectiveTransparencyPercent | Should -BeGreaterOrEqual 0
            } finally {
                Remove-Item -LiteralPath $dir -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
    }
}
