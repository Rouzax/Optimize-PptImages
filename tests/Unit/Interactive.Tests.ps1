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

    It 'parses a custom decimal under a comma-decimal culture' {
        InModuleScope Optimize-PptImages {
            $orig = [System.Threading.Thread]::CurrentThread.CurrentCulture
            try {
                [System.Threading.Thread]::CurrentThread.CurrentCulture = [System.Globalization.CultureInfo]::GetCultureInfo('nl-NL')
                $script:queue = [System.Collections.Generic.Queue[string]]::new()
                '4', '2.5' | ForEach-Object { $script:queue.Enqueue($_) }   # choose Custom(4), then 2.5
                $script:WizardReadLine = { param($p) $script:queue.Dequeue() }
                $presets = @(
                    @{ Label = 'Smallest (1.5)'; Value = 1.5 },
                    @{ Label = 'Balanced (2.0)'; Value = 2.0 },
                    @{ Label = 'Print (3.0)';    Value = 3.0 }
                )
                Read-WizardPreset -Prompt 'Headroom' -Hint 'hint' -Presets $presets -Default 2.0 -Min 0.5 -Max 4.0 |
                    Should -Be 2.5
            } finally {
                [System.Threading.Thread]::CurrentThread.CurrentCulture = $orig
            }
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

Describe 'Format-PptCommandLine' {
    It 'composes switches and tuning values from answers' {
        InModuleScope Optimize-PptImages {
            $answers = [PSCustomObject]@{
                OptimizeSlides = $true; OptimizeMastersAndLayouts = $false
                CropSlides = $true; CropMastersAndLayouts = $false; CleanUnusedLayouts = $false
                HeadroomFactor = 2.0; JpegQuality = 90
                OutputPath = ''
                TransparencyThresholdPercent = 0.1; MinSavingsPercent = 0.0
                IncludeSlides = @(); ExcludeSlides = @()
            }
            $line = Format-PptCommandLine -InputPath 'C:/decks/q3.pptx' -Answers $answers
            $line | Should -BeLike '*-InputPath "C:/decks/q3.pptx"*'
            $line | Should -BeLike '*-OptimizeSlides*'
            $line | Should -BeLike '*-CropSlides*'
            $line | Should -BeLike '*-HeadroomFactor 2*'
            $line | Should -BeLike '*-JpegQuality 90*'
            $line | Should -Not -BeLike '*-OptimizeMastersAndLayouts*'
            $line | Should -Not -BeLike '*-MinSavingsPercent*'
            $line | Should -Not -BeLike '*-IncludeSlides*'
        }
    }

    It 'includes output path and slide filters when set' {
        InModuleScope Optimize-PptImages {
            $answers = [PSCustomObject]@{
                OptimizeSlides = $true; OptimizeMastersAndLayouts = $false
                CropSlides = $false; CropMastersAndLayouts = $false; CleanUnusedLayouts = $false
                HeadroomFactor = 1.5; JpegQuality = 95
                OutputPath = 'C:/out/small.pptx'
                TransparencyThresholdPercent = 0.1; MinSavingsPercent = 5.0
                IncludeSlides = @(1, 2, 5); ExcludeSlides = @()
            }
            $line = Format-PptCommandLine -InputPath 'in.pptx' -Answers $answers
            $line | Should -BeLike '*-OutputPath "C:/out/small.pptx"*'
            $line | Should -BeLike '*-MinSavingsPercent 5*'
            $line | Should -BeLike '*-IncludeSlides 1,2,5*'
        }
    }
}

Describe 'Format-PptFindingsSummary' {
    It 'renders two blocks and caps offenders at five' {
        InModuleScope Optimize-PptImages {
            $block = {
                param($unique, $offenderCount)
                $offenders = 1..$offenderCount | ForEach-Object {
                    [PSCustomObject]@{ Name = "img$_.png"; SourceW = 8000; SourceH = 6000; DisplayW = 1200; DisplayH = 900; Ratio = (50 - $_) }
                }
                [PSCustomObject]@{
                    UsageCount = $unique; UniqueCount = $unique; TotalBytes = 1048576
                    FormatBreakdown = [ordered]@{ png = $unique }
                    ResizeCandidateCount = $offenderCount; ConversionCandidateCount = 0
                    TopOffenders = @($offenders | Select-Object -First 5)
                }
            }
            $findings = [PSCustomObject]@{ Slides = (& $block 7 7); MastersLayouts = (& $block 2 0) }
            $lines = Format-PptFindingsSummary -Findings $findings
            ($lines -join "`n") | Should -BeLike '*Slides*'
            ($lines -join "`n") | Should -BeLike '*Masters*'
            ($lines -join "`n") | Should -BeLike '*+2 more*'   # 7 candidates, 5 shown
        }
    }
}

Describe '-Interactive batch guard' {
    It 'errors when interactive is combined with multiple pipeline inputs' {
        $err = $null
        try {
            'a.pptx', 'b.pptx' | Invoke-PptImageOptimization -Interactive -ErrorAction Stop 2>$null
        } catch {
            $err = $_
        }
        $err | Should -Not -BeNullOrEmpty
        "$err" | Should -BeLike '*single file*'
    }
}

Describe 'Show-PptInteractiveWizard' {
    BeforeEach {
        InModuleScope Optimize-PptImages {
            $script:findings = [PSCustomObject]@{
                Slides = [PSCustomObject]@{
                    UsageCount = 3; UniqueCount = 3; TotalBytes = 3145728
                    FormatBreakdown = [ordered]@{ png = 2; jpg = 1 }
                    ResizeCandidateCount = 1; ConversionCandidateCount = 2; TopOffenders = @()
                }
                MastersLayouts = [PSCustomObject]@{
                    UsageCount = 0; UniqueCount = 0; TotalBytes = 0
                    FormatBreakdown = [ordered]@{}; ResizeCandidateCount = 0
                    ConversionCandidateCount = 0; TopOffenders = @()
                }
            }
            $script:defaults = [PSCustomObject]@{
                OptimizeSlides = $true; OptimizeMastersAndLayouts = $false
                CropSlides = $false; CropMastersAndLayouts = $false; CleanUnusedLayouts = $false
                HeadroomFactor = 2.0; JpegQuality = 95; OutputPath = ''
                TransparencyThresholdPercent = 0.1; MinSavingsPercent = 0.0
                IncludeSlides = @(); ExcludeSlides = @()
            }
        }
    }

    It 'collects answers and returns Run' {
        InModuleScope Optimize-PptImages {
            $script:queue = [System.Collections.Generic.Queue[string]]::new()
            # ops: slides Y, masters N, crop slides N, crop masters N, clean N
            # headroom: 2, jpeg: 3, output: Enter, advanced gate: N, menu: R
            'y','n','n','n','n','2','3','','n','r' | ForEach-Object { $script:queue.Enqueue($_) }
            $script:WizardReadLine = { param($p) $script:queue.Dequeue() }
            $r = Show-PptInteractiveWizard -InputPath 'in.pptx' -Findings $script:findings -Defaults $script:defaults
            $r.Action | Should -Be 'Run'
            $r.Answers.OptimizeSlides | Should -BeTrue
            $r.Answers.HeadroomFactor | Should -Be 2.0
            $r.Answers.JpegQuality | Should -Be 95
        }
    }

    It 'returns Quit when the user quits at the menu' {
        InModuleScope Optimize-PptImages {
            $script:queue = [System.Collections.Generic.Queue[string]]::new()
            'y','n','n','n','n','2','3','','n','q' | ForEach-Object { $script:queue.Enqueue($_) }
            $script:WizardReadLine = { param($p) $script:queue.Dequeue() }
            $r = Show-PptInteractiveWizard -InputPath 'in.pptx' -Findings $script:findings -Defaults $script:defaults
            $r.Action | Should -Be 'Quit'
        }
    }
}
