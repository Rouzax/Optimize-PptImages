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
            $job.MagickArgs[0] | Should -Be '/tmp/x/image1.jpeg'
            $job.MagickArgs[-1] | Should -Be $job.ScratchPath
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
            $job.NewExtension | Should -Be ''
        }
    }
}
