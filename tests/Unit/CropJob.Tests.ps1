BeforeAll { Import-Module "$PSScriptRoot/../../Optimize-PptImages.psd1" -Force }

Describe 'CropJob class' {
    It 'holds the job fields' {
        InModuleScope Optimize-PptImages {
            $j = [CropJob]::new(); $j.Key = 'k'; $j.MagickArgs = @('a'); $j.BeforeSize = 5
            $j.Key | Should -Be 'k'; $j.MagickArgs.Count | Should -Be 1
        }
    }
}

Describe 'Get-CropJob decisions' {
    It 'builds a CropJob for a meaningful slide crop' {
        InModuleScope Optimize-PptImages {
            $script:CropSlides = $true
            $u = [ImageUsage]::new()
            $u.ImagePhysicalPath = '/tmp/x/img.png'; $u.ContextType = [ContextType]::Slide; $u.SlideNumber = 1
            $u.HasSrcRect = $true; $u.SrcRectLeft = 10000; $u.SrcRectTop = 10000; $u.SrcRectRight = 10000; $u.SrcRectBottom = 10000
            $u.SourceWidthPx = 1000; $u.SourceHeightPx = 800; $u.BeforeSizeBytes = 12345
            $job = Get-CropJob -usage $u -scratchDir '/tmp/scratch'
            $job | Should -Not -BeNullOrEmpty
            $job.MagickArgs[1] | Should -Be '-crop'
            $job.MagickArgs[-1] | Should -Be $job.ScratchPath
        }
    }
    It 'returns null and marks invalid geometry' {
        InModuleScope Optimize-PptImages {
            $script:CropSlides = $true
            $u = [ImageUsage]::new()
            $u.ImagePhysicalPath = '/tmp/x/img.png'; $u.ContextType = [ContextType]::Slide; $u.SlideNumber = 1
            $u.HasSrcRect = $true; $u.SrcRectLeft = 60000; $u.SrcRectRight = 60000  # >100% removed -> invalid
            $u.SourceWidthPx = 1000; $u.SourceHeightPx = 800
            $job = Get-CropJob -usage $u -scratchDir '/tmp/scratch'
            $job | Should -BeNullOrEmpty
            $u.OptimizationStatus | Should -Be 'Skipped_CropInvalidOrUnsafe'
        }
    }
    It 'skips when the crop flag is disabled' {
        InModuleScope Optimize-PptImages {
            $script:CropSlides = $false
            $u = [ImageUsage]::new()
            $u.ImagePhysicalPath = '/tmp/x/img.png'; $u.ContextType = [ContextType]::Slide; $u.SlideNumber = 1
            $u.HasSrcRect = $true; $u.SrcRectLeft = 10000
            $job = Get-CropJob -usage $u -scratchDir '/tmp/scratch'
            $job | Should -BeNullOrEmpty
            $u.OptimizationStatus | Should -Be 'Skipped_CropNotMaterialized'
            $u.ManualActionRequired | Should -BeTrue
        }
    }
    It 'marks both usages Skipped_MorphCropConflict when morph pair has meaningful crop' {
        InModuleScope Optimize-PptImages {
            $script:CropSlides = $true

            # Primary usage with a meaningful crop
            $u = [ImageUsage]::new()
            $u.ImagePhysicalPath = '/tmp/x/img.png'; $u.ContextType = [ContextType]::Slide; $u.SlideNumber = 2
            $u.HasSrcRect = $true; $u.SrcRectLeft = 10000; $u.SrcRectTop = 0; $u.SrcRectRight = 0; $u.SrcRectBottom = 0

            # Paired usage also with a meaningful crop
            $pair = [ImageUsage]::new()
            $pair.ImagePhysicalPath = '/tmp/x/img.png'; $pair.ContextType = [ContextType]::Slide; $pair.SlideNumber = 1
            $pair.HasSrcRect = $true; $pair.SrcRectLeft = 5000; $pair.SrcRectTop = 0; $pair.SrcRectRight = 0; $pair.SrcRectBottom = 0

            $u.MorphPair = $pair
            $u.SourceWidthPx = 1000; $u.SourceHeightPx = 800; $u.BeforeSizeBytes = 12345

            $job = Get-CropJob -usage $u -scratchDir '/tmp/scratch'
            $job | Should -BeNullOrEmpty
            $u.OptimizationStatus    | Should -Be 'Skipped_MorphCropConflict'
            $pair.OptimizationStatus | Should -Be 'Skipped_MorphCropConflict'
        }
    }
    It 'marks Skipped_AnimatedGif for a .gif path when Test-IsAnimatedGif returns true' {
        InModuleScope Optimize-PptImages {
            $script:CropSlides = $true
            Mock Test-IsAnimatedGif { $true }

            $u = [ImageUsage]::new()
            $u.ImagePhysicalPath = '/tmp/x/anim.gif'; $u.ContextType = [ContextType]::Slide; $u.SlideNumber = 1
            $u.HasSrcRect = $true; $u.SrcRectLeft = 10000; $u.SrcRectTop = 0; $u.SrcRectRight = 0; $u.SrcRectBottom = 0
            $u.SourceWidthPx = 400; $u.SourceHeightPx = 300; $u.BeforeSizeBytes = 5000

            $job = Get-CropJob -usage $u -scratchDir '/tmp/scratch'
            $job | Should -BeNullOrEmpty
            $u.OptimizationStatus | Should -Be 'Skipped_AnimatedGif'
        }
    }
}

Describe 'Complete-CropJob' {
    It 'on success: moves scratch, updates path, sets CropApplied, removes srcRect' {
        InModuleScope Optimize-PptImages {
            Initialize-Namespaces

            $dir = Join-Path ([System.IO.Path]::GetTempPath()) "cropcommit-$([guid]::NewGuid())"
            New-Item -ItemType Directory -Path (Join-Path $dir 'ppt/media') -Force | Out-Null
            New-Item -ItemType Directory -Path (Join-Path $dir 'ppt/slides/_rels') -Force | Out-Null
            try {
                # Create the original media file (simulates the uncropped image)
                $media = Join-Path $dir 'ppt/media/image1.png'
                Set-Content -LiteralPath $media -Value 'original-bytes'

                # Create a scratch file (simulates the already-produced cropped image)
                $scratch = Join-Path $dir 'scratch.png'
                Set-Content -LiteralPath $scratch -Value 'cropped-bytes'

                # Build a minimal rels file so Update-BlipRelationship can find it
                $relsPath = Join-Path $dir 'ppt/slides/_rels/slide1.xml.rels'
                $relsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
                    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>' +
                    '</Relationships>'
                Set-Content -LiteralPath $relsPath -Value $relsXml -Encoding utf8

                # Build a minimal blip XML doc with blipFill > blip + srcRect + stretch/fillRect
                $doc = [System.Xml.XmlDocument]::new()
                $aNs = 'http://schemas.openxmlformats.org/drawingml/2006/main'
                $rNs = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

                $blipFill = $doc.CreateElement('a', 'blipFill', $aNs)
                $null = $doc.AppendChild($blipFill)

                $blip = $doc.CreateElement('a', 'blip', $aNs)
                $blip.SetAttribute('embed', $rNs, 'rId1')
                $null = $blipFill.AppendChild($blip)

                $srcRect = $doc.CreateElement('a', 'srcRect', $aNs)
                $srcRect.SetAttribute('l', '10000')
                $srcRect.SetAttribute('t', '5000')
                $null = $blipFill.AppendChild($srcRect)

                $stretch = $doc.CreateElement('a', 'stretch', $aNs)
                $fillRect = $doc.CreateElement('a', 'fillRect', $aNs)
                $fillRect.SetAttribute('l', '-10000')
                $fillRect.SetAttribute('t', '-5000')
                $null = $stretch.AppendChild($fillRect)
                $null = $blipFill.AppendChild($stretch)

                # Build ImageUsage with the blip element and PartPath pointing at slide1.xml
                $u = [ImageUsage]::new()
                $u.ImagePhysicalPath = $media
                $u.OriginalFileName  = 'image1.png'
                $u.PartPath          = 'ppt/slides/slide1.xml'
                $u.BlipRId           = 'rId1'
                $u.BlipElement       = $blip
                $u.HasSrcRect        = $true
                $u.BeforeSizeBytes   = (Get-Item $media).Length
                $u.ShapeName         = 'TestShape'
                $u.Location          = 'Slide 1'

                # Build the CropJob
                $job = [CropJob]::new()
                $job.Key         = 'ppt/slides/slide1.xml|TestShape|rId1'
                $job.SourcePath  = $media
                $job.ScratchPath = $scratch
                $job.BeforeSize  = $u.BeforeSizeBytes
                $job.Usage       = $u

                $afterSize = (Get-Item $scratch).Length
                $r = Complete-CropJob -job $job -scratchPath $scratch -success $true -afterSize $afterSize -tempDir $dir

                $r.Success          | Should -BeTrue
                $u.CropApplied      | Should -BeTrue
                $u.HasSrcRect       | Should -BeFalse
                $u.AfterSizeBytes   | Should -Be $afterSize
                $u.OptimizedFile    | Should -BeLike '*_cropped*'
                # ImagePhysicalPath must point to the new cropped file
                $u.ImagePhysicalPath | Should -BeLike '*_cropped*'
                (Test-Path -LiteralPath $u.ImagePhysicalPath) | Should -BeTrue
                # The original media file must no longer exist (was not renamed; a new name is created)
                # srcRect must have been removed from the blipFill
                $blipFill.SelectSingleNode('a:srcRect', $script:NsMgr) | Should -BeNullOrEmpty
                # fillRect negative attrs should have been reset
                $fr = $blipFill.SelectSingleNode('a:stretch/a:fillRect', $script:NsMgr)
                $fr.GetAttribute('l') | Should -BeNullOrEmpty
                $fr.GetAttribute('t') | Should -BeNullOrEmpty
            } finally {
                Remove-Item -LiteralPath $dir -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
    }

    It 'on failure: sets Skipped_CropInvalidOrUnsafe and returns Success=false' {
        InModuleScope Optimize-PptImages {
            $u = [ImageUsage]::new()
            $u.ImagePhysicalPath = '/tmp/x/img.png'
            $u.ShapeName         = 'Fail'
            $u.Location          = 'Slide 1'

            $job = [CropJob]::new()
            $job.Key     = 'k'
            $job.Usage   = $u
            $job.BeforeSize = 0

            $r = Complete-CropJob -job $job -scratchPath '' -success $false -afterSize 0 -tempDir '/tmp'
            $r.Success                 | Should -BeFalse
            $u.OptimizationStatus      | Should -Be 'Skipped_CropInvalidOrUnsafe'
            $u.WhyNotOptimized         | Should -Be 'Crop operation failed'
        }
    }
}
