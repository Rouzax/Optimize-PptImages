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

Describe 'Get-OptimizationJob sRGB conversion' {
    It 'inserts -profile before -strip for a flagged JPEG' {
        InModuleScope Optimize-PptImages {
            $HeadroomFactor = 2.0; $JpegQuality = 95; $TransparencyThresholdPercent = 0.1
            function New-ColGrp {
                param($Path, $SrcW, $SrcH, $DispW, $DispH, $Bytes, $NeedsSrgb)
                $u = [ImageUsage]::new()
                $u.ImagePhysicalPath = $Path; $u.OriginalFileName = (Split-Path $Path -Leaf)
                $u.ContextType = [ContextType]::Slide; $u.SlideNumber = 1
                $u.SourceWidthPx = $SrcW; $u.SourceHeightPx = $SrcH
                $u.DisplayWidthPx = $DispW; $u.DisplayHeightPx = $DispH
                $u.EffectiveTransparencyPercent = 0; $u.SourceJpegQuality = 90
                $u.NeedsSrgbConversion = $NeedsSrgb; $u.OptimizationStatus = 'Pending'
                $g = [ImageGroup]::new(); $g.PhysicalPath = $Path; $g.OriginalSizeBytes = $Bytes; $g.Usages.Add($u)
                return $g
            }
            $g = New-ColGrp '/tmp/x/a.jpeg' 4000 3000 800 600 500000 $true
            $job = Get-OptimizationJob -group $g -maxDisplay @{ Width=800; Height=600 } -scratchDir '/tmp/s'
            $args = $job.MagickArgs
            $pi = [array]::IndexOf($args, '-profile'); $si = [array]::IndexOf($args, '-strip')
            $pi | Should -BeGreaterThan -1
            $args[$pi + 1] | Should -Be $script:SrgbProfilePath
            $pi | Should -BeLessThan $si    # -profile comes before -strip
        }
    }
    It 'does not add -profile for an unflagged (sRGB) JPEG' {
        InModuleScope Optimize-PptImages {
            $HeadroomFactor = 2.0; $JpegQuality = 95; $TransparencyThresholdPercent = 0.1
            function New-ColGrp {
                param($Path, $SrcW, $SrcH, $DispW, $DispH, $Bytes, $NeedsSrgb)
                $u = [ImageUsage]::new()
                $u.ImagePhysicalPath = $Path; $u.OriginalFileName = (Split-Path $Path -Leaf)
                $u.ContextType = [ContextType]::Slide; $u.SlideNumber = 1
                $u.SourceWidthPx = $SrcW; $u.SourceHeightPx = $SrcH
                $u.DisplayWidthPx = $DispW; $u.DisplayHeightPx = $DispH
                $u.EffectiveTransparencyPercent = 0; $u.SourceJpegQuality = 90
                $u.NeedsSrgbConversion = $NeedsSrgb; $u.OptimizationStatus = 'Pending'
                $g = [ImageGroup]::new(); $g.PhysicalPath = $Path; $g.OriginalSizeBytes = $Bytes; $g.Usages.Add($u)
                return $g
            }
            $g = New-ColGrp '/tmp/x/b.jpeg' 4000 3000 800 600 500000 $false
            $job = Get-OptimizationJob -group $g -maxDisplay @{ Width=800; Height=600 } -scratchDir '/tmp/s'
            ([array]::IndexOf($job.MagickArgs, '-profile')) | Should -Be -1
        }
    }
}

Describe 'Update-OptimizeProbesParallel color detection' {
    It 'flags an Adobe RGB image and leaves an sRGB image unflagged' {
        InModuleScope Optimize-PptImages {
            $adobeProfile = '/usr/share/color/icc/colord/AdobeRGB1998.icc'
            if (-not (Test-Path -LiteralPath $adobeProfile)) {
                Set-ItResult -Skipped -Because 'system AdobeRGB profile not available to build the non-sRGB test input'
                return
            }
            $script:MagickExe = Find-ImageMagick -ProvidedPath $null
            $dir = Join-Path ([System.IO.Path]::GetTempPath()) "color-$([guid]::NewGuid())"
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
            try {
                $srgb = Join-Path $dir 'srgb.jpg'
                $adobe = Join-Path $dir 'adobe.jpg'
                & $script:MagickExe -size 40x30 'xc:red' $srgb
                & $script:MagickExe -size 40x30 'xc:red' -profile $script:SrgbProfilePath $srgb    # ensure plain sRGB-ish
                & $script:MagickExe -size 40x30 'xc:red' -profile $adobeProfile $adobe
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
    It 'untagged grayscale image is not flagged for sRGB conversion' {
        InModuleScope Optimize-PptImages {
            $script:MagickExe = Find-ImageMagick -ProvidedPath $null
            $dir = Join-Path ([System.IO.Path]::GetTempPath()) "color-$([guid]::NewGuid())"
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
            try {
                $gray = Join-Path $dir 'u_gray.png'
                & $script:MagickExe -size 20x20 'xc:gray50' -colorspace Gray $gray
                $usages = [System.Collections.Generic.List[ImageUsage]]::new()
                $u = [ImageUsage]::new(); $u.ImagePhysicalPath = $gray; $u.OriginalFileName = (Split-Path $gray -Leaf)
                $usages.Add($u)
                Update-OptimizeProbesParallel -usages $usages
                $usages[0].NeedsSrgbConversion | Should -BeFalse
            } finally { Remove-Item -LiteralPath $dir -Recurse -Force -ErrorAction SilentlyContinue }
        }
    }
    It 'untagged sRGB image is not flagged for sRGB conversion' {
        InModuleScope Optimize-PptImages {
            $script:MagickExe = Find-ImageMagick -ProvidedPath $null
            $dir = Join-Path ([System.IO.Path]::GetTempPath()) "color-$([guid]::NewGuid())"
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
            try {
                $srgb = Join-Path $dir 'u_srgb.png'
                & $script:MagickExe -size 20x20 'xc:red' $srgb
                $usages = [System.Collections.Generic.List[ImageUsage]]::new()
                $u = [ImageUsage]::new(); $u.ImagePhysicalPath = $srgb; $u.OriginalFileName = (Split-Path $srgb -Leaf)
                $usages.Add($u)
                Update-OptimizeProbesParallel -usages $usages
                $usages[0].NeedsSrgbConversion | Should -BeFalse
            } finally { Remove-Item -LiteralPath $dir -Recurse -Force -ErrorAction SilentlyContinue }
        }
    }
    It 'untagged CMYK image is flagged for sRGB conversion' {
        InModuleScope Optimize-PptImages {
            $script:MagickExe = Find-ImageMagick -ProvidedPath $null
            $dir = Join-Path ([System.IO.Path]::GetTempPath()) "color-$([guid]::NewGuid())"
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
            try {
                $cmyk = Join-Path $dir 'u_cmyk.jpg'
                & $script:MagickExe -size 20x20 'xc:red' -colorspace CMYK $cmyk
                $usages = [System.Collections.Generic.List[ImageUsage]]::new()
                $u = [ImageUsage]::new(); $u.ImagePhysicalPath = $cmyk; $u.OriginalFileName = (Split-Path $cmyk -Leaf)
                $usages.Add($u)
                Update-OptimizeProbesParallel -usages $usages
                $usages[0].NeedsSrgbConversion | Should -BeTrue
            } finally { Remove-Item -LiteralPath $dir -Recurse -Force -ErrorAction SilentlyContinue }
        }
    }
}
