BeforeAll {
    Import-Module "$PSScriptRoot/../../Optimize-PptImages.psd1" -Force
}

Describe 'ConvertFrom-MagickFacts' {
    It 'parses opaque facts' {
        InModuleScope Optimize-PptImages {
            $f = ConvertFrom-MagickFacts -Raw '120 80 True'
            $f.Width    | Should -Be 120
            $f.Height   | Should -Be 80
            $f.IsOpaque | Should -BeTrue
        }
    }

    It 'parses non-opaque facts with extra whitespace' {
        InModuleScope Optimize-PptImages {
            $f = ConvertFrom-MagickFacts -Raw "  640   480   False  "
            $f.Width    | Should -Be 640
            $f.Height   | Should -Be 480
            $f.IsOpaque | Should -BeFalse
        }
    }
}
