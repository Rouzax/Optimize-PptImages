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

Describe 'Get-PptImageFindings' {
    It 'splits slides vs masters and flags resize and conversion candidates' {
        InModuleScope Optimize-PptImages {
            function New-Usage {
                param($Ctx, $Path, $Name, $SrcW, $SrcH, $DispW, $DispH, $Bytes, $Transparency = 0)
                $u = [ImageUsage]::new()
                $u.ContextType = $Ctx
                $u.ImagePhysicalPath = $Path
                $u.OriginalFileName = $Name
                $u.SourceWidthPx = $SrcW; $u.SourceHeightPx = $SrcH
                $u.DisplayWidthPx = $DispW; $u.DisplayHeightPx = $DispH
                $u.BeforeSizeBytes = $Bytes
                $u.EffectiveTransparencyPercent = $Transparency
                return $u
            }
            $list = [System.Collections.Generic.List[ImageUsage]]::new()
            $list.Add((New-Usage ([ContextType]::Slide)  'a/big.png'  'big.png'  8000 6000 1200 900 500000 0))
            $list.Add((New-Usage ([ContextType]::Slide)  'a/ok.jpg'   'ok.jpg'   1000  800 1000 800 120000 0))
            $list.Add((New-Usage ([ContextType]::Master) 'a/logo.png' 'logo.png' 2000 2000  200 200  90000 0))

            $f = Get-PptImageFindings -usages $list
            $f.Slides.UniqueCount               | Should -Be 2
            $f.Slides.ResizeCandidateCount      | Should -Be 1
            $f.Slides.ConversionCandidateCount  | Should -Be 1   # big.png opaque
            $f.Slides.TopOffenders[0].Name      | Should -Be 'big.png'
            $f.MastersLayouts.UniqueCount       | Should -Be 1
            $f.MastersLayouts.ResizeCandidateCount | Should -Be 1
            $f.Slides.FormatBreakdown['png']    | Should -Be 1
            $f.Slides.FormatBreakdown['jpg']    | Should -Be 1
        }
    }
}

Describe 'Wizard prompt helpers' {
    It 'maps a preset selection to its value' {
        InModuleScope Optimize-PptImages {
            $script:queue = [System.Collections.Generic.Queue[string]]::new()
            $script:queue.Enqueue('2')
            $script:WizardReadLine = { param($p) $script:queue.Dequeue() }
            $presets = @(
                @{ Label = 'Smallest (1.5)'; Value = 1.5 },
                @{ Label = 'Balanced (2.0)'; Value = 2.0 },
                @{ Label = 'Print (3.0)';    Value = 3.0 }
            )
            Read-WizardPreset -Prompt 'Headroom' -Hint 'hint' -Presets $presets -Default 2.0 -Min 0.5 -Max 4.0 |
                Should -Be 2.0
        }
    }

    It 're-asks on garbage then accepts a custom value' {
        InModuleScope Optimize-PptImages {
            $script:queue = [System.Collections.Generic.Queue[string]]::new()
            'x', '4', '2.5' | ForEach-Object { $script:queue.Enqueue($_) }  # garbage, choose Custom(4), then 2.5
            $script:WizardReadLine = { param($p) $script:queue.Dequeue() }
            $presets = @(
                @{ Label = 'Smallest (1.5)'; Value = 1.5 },
                @{ Label = 'Balanced (2.0)'; Value = 2.0 },
                @{ Label = 'Print (3.0)';    Value = 3.0 }
            )
            Read-WizardPreset -Prompt 'Headroom' -Hint 'hint' -Presets $presets -Default 2.0 -Min 0.5 -Max 4.0 |
                Should -Be 2.5
        }
    }

    It 'returns default on empty yes/no input' {
        InModuleScope Optimize-PptImages {
            $script:queue = [System.Collections.Generic.Queue[string]]::new()
            $script:queue.Enqueue('')
            $script:WizardReadLine = { param($p) $script:queue.Dequeue() }
            Read-WizardYesNo -Prompt 'Optimize slides?' -Default $true | Should -BeTrue
        }
    }
}
