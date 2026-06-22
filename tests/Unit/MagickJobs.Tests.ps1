BeforeAll { Import-Module "$PSScriptRoot/../../Optimize-PptImages.psd1" -Force }

Describe 'Invoke-MagickJobsParallel' {
    It 'runs each job to its scratch and returns keyed results' {
        InModuleScope Optimize-PptImages {
            $script:MagickExe = Find-ImageMagick -ProvidedPath $null
            $dir = Join-Path ([System.IO.Path]::GetTempPath()) "mjp-$([guid]::NewGuid())"
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
            try {
                $src = Join-Path $dir 'src.png'
                & $script:MagickExe -size 50x40 'xc:red' $src
                $jobs = @(
                    [PSCustomObject]@{ Key = 'a'; MagickArgs = @($src, '-crop', '20x10+0+0', '+repage', (Join-Path $dir 'a.png')); ScratchPath = (Join-Path $dir 'a.png') },
                    [PSCustomObject]@{ Key = 'b'; MagickArgs = @($src, '-crop', '10x10+0+0', '+repage', (Join-Path $dir 'b.png')); ScratchPath = (Join-Path $dir 'b.png') }
                )
                $res = Invoke-MagickJobsParallel -jobs $jobs -activity 'Test' -progressId 9
                $res['a'].Success | Should -BeTrue
                $res['a'].AfterSize | Should -BeGreaterThan 0
                $res['b'].Success | Should -BeTrue
            } finally { Remove-Item -LiteralPath $dir -Recurse -Force -ErrorAction SilentlyContinue }
        }
    }
}
