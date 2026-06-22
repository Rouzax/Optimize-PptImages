BeforeAll {
    Import-Module "$PSScriptRoot/../../Optimize-PptImages.psd1" -Force
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
