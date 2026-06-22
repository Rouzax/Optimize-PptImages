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
