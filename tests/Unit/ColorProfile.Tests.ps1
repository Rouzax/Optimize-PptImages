BeforeAll { Import-Module "$PSScriptRoot/../../Optimize-PptImages.psd1" -Force }

Describe 'Test-NeedsSrgbConversion' {
    It 'returns false for sRGB colorspace with no profile' {
        InModuleScope Optimize-PptImages { Test-NeedsSrgbConversion -Colorspace 'sRGB' -IccDescription '' | Should -BeFalse }
    }
    It 'returns true for CMYK' {
        InModuleScope Optimize-PptImages { Test-NeedsSrgbConversion -Colorspace 'CMYK' -IccDescription '' | Should -BeTrue }
    }
    It 'returns true for an Adobe RGB profile' {
        InModuleScope Optimize-PptImages { Test-NeedsSrgbConversion -Colorspace 'sRGB' -IccDescription 'Compatible with Adobe RGB (1998)' | Should -BeTrue }
    }
    It 'returns false for an sRGB profile description' {
        InModuleScope Optimize-PptImages { Test-NeedsSrgbConversion -Colorspace 'sRGB' -IccDescription 'sRGB IEC61966-2.1' | Should -BeFalse }
    }
    It 'returns false for grayscale with no profile' {
        InModuleScope Optimize-PptImages { Test-NeedsSrgbConversion -Colorspace 'Gray' -IccDescription '' | Should -BeFalse }
    }
}

Describe 'bundled sRGB profile' {
    It 'is present and resolvable' {
        InModuleScope Optimize-PptImages { Test-Path -LiteralPath $script:SrgbProfilePath | Should -BeTrue }
    }
}

Describe 'Update-OptimizeProbesParallel color detection' {
    It 'flags an Adobe RGB image and leaves an sRGB image unflagged' {
        InModuleScope Optimize-PptImages {
            $script:MagickExe = Find-ImageMagick -ProvidedPath $null
            $dir = Join-Path ([System.IO.Path]::GetTempPath()) "color-$([guid]::NewGuid())"
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
            try {
                $srgb = Join-Path $dir 'srgb.jpg'
                $adobe = Join-Path $dir 'adobe.jpg'
                & $script:MagickExe -size 40x30 'xc:red' $srgb
                & $script:MagickExe -size 40x30 'xc:red' -profile $script:SrgbProfilePath $srgb    # ensure plain sRGB-ish
                & $script:MagickExe -size 40x30 'xc:red' -profile '/usr/share/color/icc/colord/AdobeRGB1998.icc' $adobe
                $usages = [System.Collections.Generic.List[ImageUsage]]::new()
                foreach ($p in @($srgb, $adobe)) {
                    $u = [ImageUsage]::new(); $u.ImagePhysicalPath = $p; $u.OriginalFileName = (Split-Path $p -Leaf)
                    $usages.Add($u)
                }
                Update-OptimizeProbesParallel -usages $usages
                ($usages | Where-Object { $_.ImagePhysicalPath -eq $adobe }).NeedsSrgbConversion | Should -BeTrue
                ($usages | Where-Object { $_.ImagePhysicalPath -eq $srgb }).NeedsSrgbConversion  | Should -BeFalse
            } finally { Remove-Item -LiteralPath $dir -Recurse -Force -ErrorAction SilentlyContinue }
        }
    }
}
