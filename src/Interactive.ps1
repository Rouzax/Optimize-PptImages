#region Interactive Wizard

# Build a fixed-size findings summary from scanned usages. Pure: no disk or magick.
# SourceWidthPx/SourceHeightPx and EffectiveTransparencyPercent must already be set.
function Get-PptImageFindings {
    param([System.Collections.Generic.List[ImageUsage]]$usages)

    function New-FindingsBlock {
        param([object[]]$blockUsages)

        $groups = $blockUsages | Group-Object -Property ImagePhysicalPath
        $formats = [ordered]@{}
        $resize = 0
        $convert = 0
        $offenders = [System.Collections.Generic.List[object]]::new()
        $totalBytes = [long]0

        foreach ($g in $groups) {
            $first = $g.Group[0]
            $ext = [System.IO.Path]::GetExtension($first.OriginalFileName).TrimStart('.').ToLowerInvariant()
            if (-not $ext) { $ext = '(none)' }
            if ($formats.Contains($ext)) { $formats[$ext] = [int]$formats[$ext] + 1 } else { $formats[$ext] = 1 }

            $totalBytes += [long]$first.BeforeSizeBytes
            $maxDisplayArea = ($g.Group | ForEach-Object { [long]$_.DisplayWidthPx * [long]$_.DisplayHeightPx } | Measure-Object -Maximum).Maximum
            $sourceArea = [long]$first.SourceWidthPx * [long]$first.SourceHeightPx

            if ($first.SourceWidthPx -gt 0 -and $sourceArea -gt $maxDisplayArea) {
                $resize++
                $maxUsage = $g.Group | Sort-Object { [long]$_.DisplayWidthPx * [long]$_.DisplayHeightPx } -Descending | Select-Object -First 1
                $ratio = [math]::Round(($sourceArea / [math]::Max(1, $maxDisplayArea)), 1)
                $offenders.Add([PSCustomObject]@{
                    Name = $first.OriginalFileName
                    SourceW = $first.SourceWidthPx; SourceH = $first.SourceHeightPx
                    DisplayW = $maxUsage.DisplayWidthPx; DisplayH = $maxUsage.DisplayHeightPx
                    Ratio = $ratio
                })
            }

            $maxTransparency = ($g.Group | ForEach-Object { [double]$_.EffectiveTransparencyPercent } | Measure-Object -Maximum).Maximum
            if ($ext -eq 'png' -and $maxTransparency -eq 0) { $convert++ }
        }

        $top = $offenders | Sort-Object -Property Ratio -Descending | Select-Object -First 5

        return [PSCustomObject]@{
            UsageCount               = $blockUsages.Count
            UniqueCount              = @($groups).Count
            TotalBytes               = $totalBytes
            FormatBreakdown          = $formats
            ResizeCandidateCount     = $resize
            ConversionCandidateCount = $convert
            TopOffenders             = @($top)
        }
    }

    $slideUsages  = @($usages | Where-Object { $_.ContextType -eq [ContextType]::Slide })
    $masterUsages = @($usages | Where-Object { $_.ContextType -eq [ContextType]::Layout -or $_.ContextType -eq [ContextType]::Master })

    return [PSCustomObject]@{
        Slides         = New-FindingsBlock -blockUsages $slideUsages
        MastersLayouts = New-FindingsBlock -blockUsages $masterUsages
    }
}

    # Input seam so tests can feed scripted answers without a TTY.
    $script:WizardReadLine = { param($promptText) Read-Host -Prompt $promptText }

    function Read-WizardLine {
        param([string]$Prompt)
        return (& $script:WizardReadLine $Prompt)
    }

    function Read-WizardYesNo {
        param([string]$Prompt, [bool]$Default)
        $suffix = if ($Default) { '[Y/n]' } else { '[y/N]' }
        while ($true) {
            $raw = (Read-WizardLine "$Prompt $suffix").Trim().ToLowerInvariant()
            if ($raw -eq '') { return $Default }
            if ($raw -in @('y', 'yes')) { return $true }
            if ($raw -in @('n', 'no'))  { return $false }
            Write-Host "  Please answer y or n." -ForegroundColor DarkYellow
        }
    }

    function Read-WizardPreset {
        param(
            [string]$Prompt,
            [string]$Hint,
            [object[]]$Presets,
            [double]$Default,
            [double]$Min,
            [double]$Max
        )
        Write-Host ""
        Write-Host $Prompt -ForegroundColor Cyan
        if ($Hint) { Write-Host "  $Hint" -ForegroundColor Gray }
        $i = 1
        foreach ($p in $Presets) { Write-Host "  $i) $($p.Label)"; $i++ }
        $customIndex = $i
        Write-Host "  $customIndex) Custom"

        while ($true) {
            $raw = (Read-WizardLine "Choose 1-$customIndex [Enter for default $Default]").Trim()
            if ($raw -eq '') { return $Default }
            $sel = 0
            if (-not [int]::TryParse($raw, [ref]$sel)) {
                Write-Host "  Enter a number between 1 and $customIndex." -ForegroundColor DarkYellow
                continue
            }
            if ($sel -ge 1 -and $sel -lt $customIndex) {
                return [double]$Presets[$sel - 1].Value
            }
            if ($sel -eq $customIndex) {
                while ($true) {
                    $rawCustom = (Read-WizardLine "Custom value ($Min-$Max)").Trim()
                    $val = 0.0
                    if ([double]::TryParse($rawCustom, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$val) -and $val -ge $Min -and $val -le $Max) {
                        return $val
                    }
                    Write-Host "  Enter a number between $Min and $Max." -ForegroundColor DarkYellow
                }
            }
            Write-Host "  Enter a number between 1 and $customIndex." -ForegroundColor DarkYellow
        }
    }

    function Format-PptFindingsSummary {
        param([object]$Findings)

        $lines = [System.Collections.Generic.List[string]]::new()

        function Add-Block {
            param([System.Collections.Generic.List[string]]$out, [string]$title, [object]$b)
            $mb = [math]::Round($b.TotalBytes / 1MB, 1)
            $out.Add("$title")
            $out.Add("  $($b.UsageCount) usages across $($b.UniqueCount) unique media, $mb MB")
            if ($b.FormatBreakdown.Count -gt 0) {
                $fmt = ($b.FormatBreakdown.Keys | ForEach-Object { "$_ $($b.FormatBreakdown[$_])" }) -join ', '
                $out.Add("  Formats: $fmt")
            }
            $out.Add("  Resize candidates: $($b.ResizeCandidateCount)   PNG->JPEG candidates: $($b.ConversionCandidateCount)")
            if ($b.TopOffenders.Count -gt 0) {
                $out.Add("  Biggest resize opportunities (top $($b.TopOffenders.Count) of $($b.ResizeCandidateCount)):")
                foreach ($o in $b.TopOffenders) {
                    $out.Add(("    {0,-22} {1}x{2}  shown {3}x{4}  ~{5}x oversized" -f $o.Name, $o.SourceW, $o.SourceH, $o.DisplayW, $o.DisplayH, $o.Ratio))
                }
                $more = $b.ResizeCandidateCount - $b.TopOffenders.Count
                if ($more -gt 0) { $out.Add("    ... +$more more (see CSV report for the full list)") }
            }
            $out.Add('')
        }

        Add-Block -out $lines -title 'Slides' -b $Findings.Slides
        Add-Block -out $lines -title 'Masters & layouts' -b $Findings.MastersLayouts
        return $lines.ToArray()
    }

    function Show-PptInteractiveWizard {
        param([string]$InputPath, [object]$Findings, [object]$Defaults)

        foreach ($line in (Format-PptFindingsSummary -Findings $Findings)) { Write-Host $line }

        $answers = $Defaults.PSObject.Copy()

        $headroomPresets = @(
            @{ Label = 'Smallest file (1.5)'; Value = 1.5 },
            @{ Label = 'Balanced (2.0, recommended)'; Value = 2.0 },
            @{ Label = 'Print/HiDPI safe (3.0)'; Value = 3.0 }
        )
        $qualityPresets = @(
            @{ Label = 'Smaller (80)'; Value = 80 },
            @{ Label = 'Balanced (90)'; Value = 90 },
            @{ Label = 'High fidelity (95, recommended)'; Value = 95 }
        )

        function Invoke-OperationQuestions {
            param($a, $d)
            Write-Host ''
            Write-Host 'Operations' -ForegroundColor Cyan
            $a.OptimizeSlides            = Read-WizardYesNo -Prompt 'Optimize images on slides?' -Default $d.OptimizeSlides
            $a.OptimizeMastersAndLayouts = Read-WizardYesNo -Prompt 'Optimize images on masters/layouts?' -Default $d.OptimizeMastersAndLayouts
            $a.CropSlides                = Read-WizardYesNo -Prompt 'Apply crops on slides?' -Default $d.CropSlides
            $a.CropMastersAndLayouts     = Read-WizardYesNo -Prompt 'Apply crops on masters/layouts?' -Default $d.CropMastersAndLayouts
            $a.CleanUnusedLayouts        = Read-WizardYesNo -Prompt 'Remove unused layouts/masters?' -Default $d.CleanUnusedLayouts
        }

        function Invoke-TuningQuestions {
            param($a, $d)
            $a.HeadroomFactor = Read-WizardPreset -Prompt 'Headroom factor' `
                -Hint 'How much larger than display size an image may stay before resizing. Higher = sharper on zoom/HiDPI but larger files.' `
                -Presets $headroomPresets -Default $d.HeadroomFactor -Min 0.5 -Max 4.0
            $a.JpegQuality = [int](Read-WizardPreset -Prompt 'JPEG quality' `
                -Hint 'Quality for JPEG output (1-100). Higher = better image, larger files.' `
                -Presets $qualityPresets -Default $d.JpegQuality -Min 1 -Max 100)
        }

        function Invoke-OutputQuestion {
            param($a)
            Write-Host ''
            $raw = (Read-WizardLine "Output path [Enter for default derived name]").Trim()
            $a.OutputPath = $raw
        }

        function Invoke-AdvancedQuestions {
            param($a, $d)
            if (-not (Read-WizardYesNo -Prompt 'Tune advanced options (transparency threshold, min savings, slide filtering)?' -Default $false)) {
                return
            }
            $a.TransparencyThresholdPercent = Read-WizardPreset -Prompt 'Transparency threshold percent' `
                -Hint 'Below this percent of transparent pixels a PNG is treated as opaque (eligible for JPEG).' `
                -Presets @(@{ Label = 'Default (0.1)'; Value = 0.1 }) -Default $d.TransparencyThresholdPercent -Min 0.0 -Max 100.0
            $a.MinSavingsPercent = Read-WizardPreset -Prompt 'Minimum savings percent' `
                -Hint 'Skip an image unless optimizing saves at least this percent.' `
                -Presets @(@{ Label = 'Default (0)'; Value = 0.0 }) -Default $d.MinSavingsPercent -Min 0.0 -Max 50.0
            $rawIncl = (Read-WizardLine 'Include slides (comma-separated numbers, blank for all)').Trim()
            $a.IncludeSlides = if ($rawIncl) { @($rawIncl -split '\s*,\s*' | ForEach-Object { [int]$_ }) } else { @() }
            $rawExcl = (Read-WizardLine 'Exclude slides (comma-separated numbers, blank for none)').Trim()
            $a.ExcludeSlides = if ($rawExcl) { @($rawExcl -split '\s*,\s*' | ForEach-Object { [int]$_ }) } else { @() }
        }

        Invoke-OperationQuestions -a $answers -d $Defaults
        Invoke-TuningQuestions    -a $answers -d $Defaults
        Invoke-OutputQuestion     -a $answers
        Invoke-AdvancedQuestions  -a $answers -d $Defaults

        while ($true) {
            Write-Host ''
            Write-Host 'Equivalent command:' -ForegroundColor Cyan
            Write-Host "  $(Format-PptCommandLine -InputPath $InputPath -Answers $answers)" -ForegroundColor Green
            $choice = (Read-WizardLine '[R]un now / [C]opy & exit / [E]dit an answer / [Q]uit').Trim().ToLowerInvariant()
            switch ($choice) {
                'r' { return [PSCustomObject]@{ Action = 'Run';  Answers = $answers } }
                'c' { return [PSCustomObject]@{ Action = 'Copy'; Answers = $answers } }
                'q' { return [PSCustomObject]@{ Action = 'Quit'; Answers = $answers } }
                'e' {
                    Write-Host '  1) Operations  2) Headroom/Quality  3) Output path  4) Advanced'
                    $sec = (Read-WizardLine 'Edit which section? 1-4').Trim()
                    switch ($sec) {
                        '1' { Invoke-OperationQuestions -a $answers -d $answers }
                        '2' { Invoke-TuningQuestions    -a $answers -d $answers }
                        '3' { Invoke-OutputQuestion     -a $answers }
                        '4' { Invoke-AdvancedQuestions  -a $answers -d $answers }
                        default { Write-Host '  Enter 1-4.' -ForegroundColor DarkYellow }
                    }
                }
                default { Write-Host '  Enter R, C, E, or Q.' -ForegroundColor DarkYellow }
            }
        }
    }

    # Compose the equivalent CLI for the chosen answers. Operation switches only when
    # enabled, core tuning always, secondary/advanced options only when non-default.
    function Format-PptCommandLine {
        param([string]$InputPath, [object]$Answers)

        $parts = [System.Collections.Generic.List[string]]::new()
        $parts.Add('.\Optimize-PptImages.ps1')
        $parts.Add("-InputPath `"$InputPath`"")

        if ($Answers.OptimizeSlides)            { $parts.Add('-OptimizeSlides') }
        if ($Answers.OptimizeMastersAndLayouts) { $parts.Add('-OptimizeMastersAndLayouts') }
        if ($Answers.CropSlides)                { $parts.Add('-CropSlides') }
        if ($Answers.CropMastersAndLayouts)     { $parts.Add('-CropMastersAndLayouts') }
        if ($Answers.CleanUnusedLayouts)        { $parts.Add('-CleanUnusedLayouts') }

        $parts.Add("-HeadroomFactor $($Answers.HeadroomFactor)")
        $parts.Add("-JpegQuality $($Answers.JpegQuality)")

        if ($Answers.OutputPath) { $parts.Add("-OutputPath `"$($Answers.OutputPath)`"") }
        if ($Answers.TransparencyThresholdPercent -ne 0.1) { $parts.Add("-TransparencyThresholdPercent $($Answers.TransparencyThresholdPercent)") }
        if ($Answers.MinSavingsPercent -ne 0.0)            { $parts.Add("-MinSavingsPercent $($Answers.MinSavingsPercent)") }
        if ($Answers.IncludeSlides -and $Answers.IncludeSlides.Count -gt 0) { $parts.Add("-IncludeSlides $($Answers.IncludeSlides -join ',')") }
        if ($Answers.ExcludeSlides -and $Answers.ExcludeSlides.Count -gt 0) { $parts.Add("-ExcludeSlides $($Answers.ExcludeSlides -join ',')") }

        return ($parts -join ' ')
    }

#endregion
