BeforeAll {
    Import-Module "$PSScriptRoot/../../Optimize-PptImages.psd1" -Force
}

Describe 'Update-SourceDimensionsParallel' {
    It 'applies source dimensions and, with -IncludeOpacity, opacity' {
        InModuleScope Optimize-PptImages {
            $script:MagickExe = Find-ImageMagick -ProvidedPath $null
            $dir = Join-Path ([System.IO.Path]::GetTempPath()) "srcdim-$([guid]::NewGuid())"
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
            try {
                $opaque = Join-Path $dir 'opaque.png'
                $alpha  = Join-Path $dir 'alpha.png'
                & $script:MagickExe -size 120x80 'xc:red' $opaque
                & $script:MagickExe -size 60x40 'xc:none' $alpha   # fully transparent

                function New-U {
                    param($Path)
                    $u = [ImageUsage]::new()
                    $u.ImagePhysicalPath = $Path
                    $u.OriginalFileName = (Split-Path $Path -Leaf)
                    $u.EffectiveTransparencyPercent = 50   # sentinel to prove opacity is only touched with the switch
                    return $u
                }
                $usages = [System.Collections.Generic.List[ImageUsage]]::new()
                $usages.Add((New-U $opaque))
                $usages.Add((New-U $alpha))

                # Without -IncludeOpacity: dims set, opacity sentinel untouched.
                Update-SourceDimensionsParallel -usages $usages
                $usages[0].SourceWidthPx  | Should -Be 120
                $usages[0].SourceHeightPx | Should -Be 80
                $usages[1].SourceWidthPx  | Should -Be 60
                $usages[0].EffectiveTransparencyPercent | Should -Be 50
                $usages[1].EffectiveTransparencyPercent | Should -Be 50

                # With -IncludeOpacity: opacity now set (0 for opaque, 100 for transparent).
                Update-SourceDimensionsParallel -usages $usages -IncludeOpacity
                $usages[0].EffectiveTransparencyPercent | Should -Be 0
                $usages[1].EffectiveTransparencyPercent | Should -Be 100
            } finally {
                Remove-Item -LiteralPath $dir -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
    }

    It 'tolerates an empty usage list without throwing' {
        InModuleScope Optimize-PptImages {
            $script:MagickExe = Find-ImageMagick -ProvidedPath $null
            $empty = [System.Collections.Generic.List[ImageUsage]]::new()
            { Update-SourceDimensionsParallel -usages $empty } | Should -Not -Throw
        }
    }
}
