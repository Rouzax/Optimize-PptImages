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
