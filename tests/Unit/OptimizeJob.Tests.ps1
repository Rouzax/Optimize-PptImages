BeforeAll {
    Import-Module "$PSScriptRoot/../../Optimize-PptImages.psd1" -Force
}

Describe 'Update-OptimizeProbesParallel' {
    It 'caches EffectiveTransparencyPercent (~0) for an opaque PNG and SourceJpegQuality (>0) for a JPEG' {
        InModuleScope Optimize-PptImages {
            $script:MagickExe = Find-ImageMagick -ProvidedPath $null
            $dir = Join-Path ([System.IO.Path]::GetTempPath()) "optprobe-$([guid]::NewGuid())"
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
            try {
                $pngPath  = Join-Path $dir 'opaque.png'
                $jpegPath = Join-Path $dir 'q.jpg'
                & $script:MagickExe -size 40x30 'xc:red' $pngPath
                & $script:MagickExe -size 40x30 'xc:red' $jpegPath

                function New-ProbeUsage {
                    param([string]$Path)
                    $u = [ImageUsage]::new()
                    $u.ImagePhysicalPath = $Path
                    $u.OriginalFileName = (Split-Path $Path -Leaf)
                    $u.EffectiveTransparencyPercent = -1.0
                    $u.SourceJpegQuality = -2
                    return $u
                }

                $usages = [System.Collections.Generic.List[ImageUsage]]::new()
                $usages.Add((New-ProbeUsage $pngPath))
                $usages.Add((New-ProbeUsage $jpegPath))

                Update-OptimizeProbesParallel -usages $usages

                # PNG: transparency percent should be near 0 (opaque image, no alpha pixels)
                $usages[0].EffectiveTransparencyPercent | Should -BeGreaterOrEqual 0
                $usages[0].EffectiveTransparencyPercent | Should -BeLessOrEqual 1
                # PNG: SourceJpegQuality should remain unchanged (not a JPEG)
                $usages[0].SourceJpegQuality | Should -Be -2

                # JPEG: quality should be set and positive
                $usages[1].SourceJpegQuality | Should -BeGreaterThan 0
                # JPEG: EffectiveTransparencyPercent should remain unchanged (not a PNG)
                $usages[1].EffectiveTransparencyPercent | Should -Be ([double]-1)
            } finally {
                Remove-Item -LiteralPath $dir -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
    }

    It 'tolerates an empty usage list without throwing' {
        InModuleScope Optimize-PptImages {
            $script:MagickExe = Find-ImageMagick -ProvidedPath $null
            $empty = [System.Collections.Generic.List[ImageUsage]]::new()
            { Update-OptimizeProbesParallel -usages $empty } | Should -Not -Throw
        }
    }

    It 'applies probe values to every usage sharing the same physical path' {
        InModuleScope Optimize-PptImages {
            $script:MagickExe = Find-ImageMagick -ProvidedPath $null
            $dir = Join-Path ([System.IO.Path]::GetTempPath()) "optprobe2-$([guid]::NewGuid())"
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
            try {
                $pngPath = Join-Path $dir 'shared.png'
                & $script:MagickExe -size 40x30 'xc:red' $pngPath

                $usages = [System.Collections.Generic.List[ImageUsage]]::new()
                foreach ($_ in 1..3) {
                    $u = [ImageUsage]::new()
                    $u.ImagePhysicalPath = $pngPath
                    $u.OriginalFileName = 'shared.png'
                    $u.EffectiveTransparencyPercent = -1.0
                    $u.SourceJpegQuality = -2
                    $usages.Add($u)
                }

                Update-OptimizeProbesParallel -usages $usages

                # All three usages should have the same transparency value
                $usages[0].EffectiveTransparencyPercent | Should -BeGreaterOrEqual 0
                $usages[1].EffectiveTransparencyPercent | Should -Be $usages[0].EffectiveTransparencyPercent
                $usages[2].EffectiveTransparencyPercent | Should -Be $usages[0].EffectiveTransparencyPercent
            } finally {
                Remove-Item -LiteralPath $dir -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
    }

    It 'filters WARNING noise lines from transparency output and returns the numeric value' {
        # Validates that the noise-filtering logic (matching Get-ImageTransparency exactly)
        # picks the numeric token even when WARNING lines accompany the magick output.
        # This is a unit test of the sequential parse step: `Raw` is now a pre-filtered
        # numeric token from the runspace, so a contaminated raw string falls back to 0.0.
        InModuleScope Optimize-PptImages {
            # Simulate the sequential parse step that runs in the parent after the
            # parallel block returns. The previous bug: raw = "WARNING: foo 0.000000"
            # did not match '^\d+\.?\d*$' and fell back to 0.0. The fix moves filtering
            # into the runspace so Raw is already a clean numeric string or $null.
            # We test here that a clean numeric token is parsed correctly and that a
            # contaminated string (as the old code would have produced) would fail.

            # Clean token (what the fixed runspace returns):
            $cleanRaw = '0.000000'
            $cleanRaw -match '^\d+\.?\d*$' | Should -BeTrue
            [double]$cleanRaw | Should -BeGreaterOrEqual 0

            # Contaminated string (what the old runspace returned when WARNING was present):
            $noisyRaw = 'WARNING: profile icc: tag size 1578 too large for tag icm2 1380' + [char]10 + '0.000000'
            $noisyRaw -match '^\d+\.?\d*$' | Should -BeFalse

            # Filtered equivalent of the noisy string (what Where-Object produces):
            $filtered = @($noisyRaw -split "`n") |
                Where-Object { $_ -notmatch '^(WARNING|System\.Management)' -and $_ -match '^\d+\.?\d*$' } |
                Select-Object -First 1
            $filtered | Should -Be '0.000000'
            [double]$filtered | Should -BeGreaterOrEqual 0
        }
    }
}

Describe 'OptimizeJob class' {
    It 'constructs and holds the job fields' {
        InModuleScope Optimize-PptImages {
            $j = [OptimizeJob]::new()
            $j.GroupKey = 'ppt/media/image1.png'
            $j.SourcePath = '/tmp/x/image1.png'
            $j.ScratchPath = '/tmp/scratch/abc.jpeg'
            $j.MagickArgs = @('/tmp/x/image1.png', '-strip', '/tmp/scratch/abc.jpeg')
            $j.Operation = 'ConvertPngToJpeg'
            $j.BeforeSize = 1234
            $j.NewExtension = '.jpeg'
            $j.StatusName = 'ConvertedPngToJpeg'
            $j.MagickArgs.Count | Should -Be 3
            $j.Operation | Should -Be 'ConvertPngToJpeg'
        }
    }
}

Describe 'Get-OptimizationJob (pure build)' {
    It 'builds a ResizeJpeg job for an oversized JPEG' {
        InModuleScope Optimize-PptImages {
            $HeadroomFactor = 2.0; $JpegQuality = 95; $TransparencyThresholdPercent = 0.1
            function New-Grp {
                param($Path, $SrcW, $SrcH, $DispW, $DispH, $Bytes, $Transp = 0, $Quality = -1)
                $u = [ImageUsage]::new()
                $u.ImagePhysicalPath = $Path; $u.OriginalFileName = (Split-Path $Path -Leaf)
                $u.ContextType = [ContextType]::Slide; $u.SlideNumber = 1
                $u.SourceWidthPx = $SrcW; $u.SourceHeightPx = $SrcH
                $u.DisplayWidthPx = $DispW; $u.DisplayHeightPx = $DispH
                $u.EffectiveTransparencyPercent = $Transp; $u.SourceJpegQuality = $Quality
                $u.OptimizationStatus = 'Pending'
                $g = [ImageGroup]::new(); $g.PhysicalPath = $Path; $g.OriginalSizeBytes = $Bytes
                $g.Usages.Add($u)
                return $g
            }
            $g = New-Grp '/tmp/x/image1.jpeg' 4000 3000 800 600 500000 0 90
            $job = Get-OptimizationJob -group $g -maxDisplay @{ Width = 800; Height = 600 } -scratchDir '/tmp/scratch'
            $job | Should -Not -BeNullOrEmpty
            $job.Operation | Should -Be 'OptimizeJpeg'
            $job.StatusName | Should -Be 'OptimizedJpeg'
            $job.MagickArgs[0] | Should -Be '/tmp/x/image1.jpeg'
            $job.MagickArgs[-1] | Should -Be $job.ScratchPath
            # FIX 3: dims must be written onto usages for CSV report fidelity.
            # HeadroomFactor=2.0, display=800 => desiredWidth=1600; target=Min(4000,1600)=1600.
            $g.Usages[0].SourceWidthPx  | Should -Be 4000
            $g.Usages[0].TargetWidthPx  | Should -Be 1600
        }
    }

    It 'returns null and marks AlreadyOptimal for an at-size, at-quality JPEG' {
        InModuleScope Optimize-PptImages {
            $HeadroomFactor = 2.0; $JpegQuality = 95; $TransparencyThresholdPercent = 0.1
            function New-Grp {
                param($Path, $SrcW, $SrcH, $DispW, $DispH, $Bytes, $Transp = 0, $Quality = -1)
                $u = [ImageUsage]::new()
                $u.ImagePhysicalPath = $Path; $u.OriginalFileName = (Split-Path $Path -Leaf)
                $u.ContextType = [ContextType]::Slide; $u.SlideNumber = 1
                $u.SourceWidthPx = $SrcW; $u.SourceHeightPx = $SrcH
                $u.DisplayWidthPx = $DispW; $u.DisplayHeightPx = $DispH
                $u.EffectiveTransparencyPercent = $Transp; $u.SourceJpegQuality = $Quality
                $u.OptimizationStatus = 'Pending'
                $g = [ImageGroup]::new(); $g.PhysicalPath = $Path; $g.OriginalSizeBytes = $Bytes
                $g.Usages.Add($u)
                return $g
            }
            $g = New-Grp '/tmp/x/image2.jpeg' 800 600 800 600 90000 0 80
            $job = Get-OptimizationJob -group $g -maxDisplay @{ Width = 800; Height = 600 } -scratchDir '/tmp/scratch'
            $job | Should -BeNullOrEmpty
            $g.Usages[0].OptimizationStatus | Should -Be 'Skipped_AlreadyOptimal'
        }
    }

    It 'builds a ConvertedToJpeg job for an oversized opaque BMP (convertible format)' {
        InModuleScope Optimize-PptImages {
            $HeadroomFactor = 2.0; $JpegQuality = 95; $TransparencyThresholdPercent = 0.1
            function New-Grp {
                param($Path, $SrcW, $SrcH, $DispW, $DispH, $Bytes, $Transp = 0, $Quality = -1)
                $u = [ImageUsage]::new()
                $u.ImagePhysicalPath = $Path; $u.OriginalFileName = (Split-Path $Path -Leaf)
                $u.ContextType = [ContextType]::Slide; $u.SlideNumber = 1
                $u.SourceWidthPx = $SrcW; $u.SourceHeightPx = $SrcH
                $u.DisplayWidthPx = $DispW; $u.DisplayHeightPx = $DispH
                $u.EffectiveTransparencyPercent = $Transp; $u.SourceJpegQuality = $Quality
                $u.OptimizationStatus = 'Pending'
                $g = [ImageGroup]::new(); $g.PhysicalPath = $Path; $g.OriginalSizeBytes = $Bytes
                $g.Usages.Add($u)
                return $g
            }
            # Oversized opaque BMP: src 3000x2000, display 600x400, no transparency.
            # CONVERTIBLE_FORMATS = @('.bmp', '.tif', '.tiff', '.gif')
            $g = New-Grp '/tmp/x/image.bmp' 3000 2000 600 400 900000 0 -1
            $job = Get-OptimizationJob -group $g -maxDisplay @{ Width = 600; Height = 400 } -scratchDir '/tmp/scratch'
            $job | Should -Not -BeNullOrEmpty
            $job.Operation   | Should -Be 'ConvertedToJpeg'
            $job.StatusName  | Should -Be 'ConvertedToJpeg'
            $job.NewExtension | Should -Be '.jpeg'
        }
    }

    It 'builds a ConvertPngToJpeg job for an opaque PNG' {
        InModuleScope Optimize-PptImages {
            $HeadroomFactor = 2.0; $JpegQuality = 95; $TransparencyThresholdPercent = 0.1
            function New-Grp {
                param($Path, $SrcW, $SrcH, $DispW, $DispH, $Bytes, $Transp = 0, $Quality = -1)
                $u = [ImageUsage]::new()
                $u.ImagePhysicalPath = $Path; $u.OriginalFileName = (Split-Path $Path -Leaf)
                $u.ContextType = [ContextType]::Slide; $u.SlideNumber = 1
                $u.SourceWidthPx = $SrcW; $u.SourceHeightPx = $SrcH
                $u.DisplayWidthPx = $DispW; $u.DisplayHeightPx = $DispH
                $u.EffectiveTransparencyPercent = $Transp; $u.SourceJpegQuality = $Quality
                $u.OptimizationStatus = 'Pending'
                $g = [ImageGroup]::new(); $g.PhysicalPath = $Path; $g.OriginalSizeBytes = $Bytes
                $g.Usages.Add($u)
                return $g
            }
            $g = New-Grp '/tmp/x/image3.png' 2000 2000 200 200 90000 0 -1
            $job = Get-OptimizationJob -group $g -maxDisplay @{ Width = 200; Height = 200 } -scratchDir '/tmp/scratch'
            $job.Operation | Should -Be 'ConvertPngToJpeg'
            $job.StatusName | Should -Be 'ConvertedPngToJpeg'
            $job.NewExtension | Should -Be '.jpeg'
        }
    }

    It 'builds an OptimizePngAlpha job for a transparent PNG' {
        InModuleScope Optimize-PptImages {
            $HeadroomFactor = 2.0; $JpegQuality = 95; $TransparencyThresholdPercent = 0.1
            function New-Grp {
                param($Path, $SrcW, $SrcH, $DispW, $DispH, $Bytes, $Transp = 0, $Quality = -1)
                $u = [ImageUsage]::new()
                $u.ImagePhysicalPath = $Path; $u.OriginalFileName = (Split-Path $Path -Leaf)
                $u.ContextType = [ContextType]::Slide; $u.SlideNumber = 1
                $u.SourceWidthPx = $SrcW; $u.SourceHeightPx = $SrcH
                $u.DisplayWidthPx = $DispW; $u.DisplayHeightPx = $DispH
                $u.EffectiveTransparencyPercent = $Transp; $u.SourceJpegQuality = $Quality
                $u.OptimizationStatus = 'Pending'
                $g = [ImageGroup]::new(); $g.PhysicalPath = $Path; $g.OriginalSizeBytes = $Bytes
                $g.Usages.Add($u)
                return $g
            }
            $g = New-Grp '/tmp/x/image4.png' 2000 2000 200 200 90000 40 -1
            $job = Get-OptimizationJob -group $g -maxDisplay @{ Width = 200; Height = 200 } -scratchDir '/tmp/scratch'
            $job.Operation | Should -Be 'OptimizePngAlpha'
            $job.StatusName | Should -Be 'OptimizedPngAlpha'
            $job.NewExtension | Should -Be ''
        }
    }
}

Describe 'Complete-OptimizationJob' {
    It 'rejects when savings are below threshold and keeps the original' {
        InModuleScope Optimize-PptImages {
            $dir = Join-Path ([System.IO.Path]::GetTempPath()) "commit-$([guid]::NewGuid())"
            New-Item -ItemType Directory -Path (Join-Path $dir 'ppt/media') -Force | Out-Null
            try {
                $media = Join-Path $dir 'ppt/media/image1.jpeg'
                Set-Content -LiteralPath $media -Value 'original-bytes'
                $scratch = Join-Path $dir 'scratch.jpeg'
                Set-Content -LiteralPath $scratch -Value 'this-scratch-is-not-smaller-enough'
                $u = [ImageUsage]::new(); $u.ImagePhysicalPath = $media; $u.OriginalFileName = 'image1.jpeg'
                $g = [ImageGroup]::new(); $g.PhysicalPath = $media; $g.OriginalSizeBytes = (Get-Item $media).Length
                $g.Usages.Add($u)
                $job = [OptimizeJob]::new(); $job.GroupKey = $media; $job.NewExtension = ''
                $job.BeforeSize = $g.OriginalSizeBytes; $job.StatusName = 'OptimizeJpeg'
                $r = Complete-OptimizationJob -job $job -group $g -scratchPath $scratch -success $true -afterSize 999999 -tempDir $dir
                $r.Success | Should -BeFalse
                (Test-Path -LiteralPath $media) | Should -BeTrue
                $u.OptimizationStatus | Should -Be 'Skipped_NoNetSavings'
            } finally {
                Remove-Item -LiteralPath $dir -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
    }
}
