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
