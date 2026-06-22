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

#endregion
