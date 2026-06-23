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
