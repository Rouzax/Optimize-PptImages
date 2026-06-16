BeforeAll {
    Import-Module "$PSScriptRoot/../../Optimize-PptImages.psd1" -Force
}

Describe 'ConvertFrom-EMU' {
    It 'converts 914400 EMU to 96 px (1 inch at 96 DPI)' {
        InModuleScope Optimize-PptImages {
            ConvertFrom-EMU 914400 | Should -Be 96
        }
    }
}

Describe 'Test-ThinkCellName' {
    It 'returns true for a think-cell marker name' {
        InModuleScope Optimize-PptImages {
            Test-ThinkCellName 'think-cell data - do not delete' | Should -BeTrue
        }
    }

    It 'returns false for an ordinary shape name' {
        InModuleScope Optimize-PptImages {
            Test-ThinkCellName 'Picture 3' | Should -BeFalse
        }
    }

    It 'returns false for an empty string' {
        InModuleScope Optimize-PptImages {
            Test-ThinkCellName '' | Should -BeFalse
        }
    }

    It 'returns false for $null' {
        InModuleScope Optimize-PptImages {
            Test-ThinkCellName $null | Should -BeFalse
        }
    }
}

Describe 'Format-SlideList' {
    It 'returns (none) for an empty list' {
        InModuleScope Optimize-PptImages {
            $empty = [System.Collections.Generic.List[int]]::new()
            Format-SlideList $empty | Should -Be '(none)'
        }
    }

    It 'returns a sorted, de-duplicated, comma-separated list' {
        InModuleScope Optimize-PptImages {
            $list = [System.Collections.Generic.List[int]]::new()
            foreach ($n in 3, 1, 3, 2) { $list.Add($n) }
            Format-SlideList $list | Should -Be '1, 2, 3'
        }
    }
}
