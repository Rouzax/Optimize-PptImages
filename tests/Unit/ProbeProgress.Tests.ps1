BeforeAll { Import-Module "$PSScriptRoot/../../Optimize-PptImages.psd1" -Force }

Describe 'Update-OptimizeProbesParallel color-probe progress' {
    It 'renders the Analyzing images bar during the color probe' {
        InModuleScope Optimize-PptImages {
            Mock Write-Progress {}
            $script:MagickExe = Find-ImageMagick -ProvidedPath $null
            $dir = Join-Path ([System.IO.Path]::GetTempPath()) "probe-$([guid]::NewGuid())"
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
            try {
                $p1 = Join-Path $dir 'a.png'; $p2 = Join-Path $dir 'b.jpg'
                & $script:MagickExe -size 40x30 'xc:red' $p1
                & $script:MagickExe -size 40x30 'xc:blue' $p2
                $usages = [System.Collections.Generic.List[ImageUsage]]::new()
                foreach ($p in @($p1, $p2)) {
                    $u = [ImageUsage]::new(); $u.ImagePhysicalPath = $p; $u.OriginalFileName = (Split-Path $p -Leaf)
                    $usages.Add($u)
                }
                Update-OptimizeProbesParallel -usages $usages
                Should -Invoke Write-Progress -ParameterFilter {
                    $Activity -eq 'Analyzing images' -and $Status -like 'Color profiles*'
                }
            } finally { Remove-Item -LiteralPath $dir -Recurse -Force -ErrorAction SilentlyContinue }
        }
    }
}
