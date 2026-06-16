    #region Verbose Helper

    function Test-VerboseMode {
        return $VerbosePreference -eq 'Continue'
    }

    #endregion

    #region Utility Functions

    function Test-FileInUse {
        param([string]$Path)

        if (-not (Test-Path -LiteralPath $Path)) {
            return $false
        }

        try {
            $stream = [System.IO.File]::Open($Path, 'Open', 'Write', 'None')
            $stream.Close()
            $stream.Dispose()
            return $false
        } catch {
            return $true
        }
    }

    function ConvertFrom-EMU {
        param([long]$emu)
        return [Math]::Round($emu / $script:Config.EMU_PER_INCH * $script:Config.DPI)
    }

    function Get-CropPercentage {
        param(
            [double]$left,
            [double]$top,
            [double]$right,
            [double]$bottom
        )

        $widthRemaining = (100000 - $left - $right) / 1000.0
        $heightRemaining = (100000 - $top - $bottom) / 1000.0

        return @{
            WidthPercent = $widthRemaining
            HeightPercent = $heightRemaining
        }
    }

    function Test-IsNoOpCrop {
        param(
            [double]$left,
            [double]$top,
            [double]$right,
            [double]$bottom
        )

        $crop = Get-CropPercentage -left $left -top $top -right $right -bottom $bottom
        return ($crop.WidthPercent -ge $script:Config.NO_OP_CROP_THRESHOLD -and
                $crop.HeightPercent -ge $script:Config.NO_OP_CROP_THRESHOLD)
    }

    function Format-ByteSize {
        param([long]$bytes)

        if ($bytes -ge 1MB) {
            return "{0:N2} MB" -f ($bytes / 1MB)
        } elseif ($bytes -ge 1KB) {
            return "{0:N2} KB" -f ($bytes / 1KB)
        } else {
            return "$bytes B"
        }
    }

    function Format-Savings {
        param(
            [long]$before,
            [long]$after
        )

        if ($before -eq 0) { return "N/A" }

        $saved = $before - $after
        $percent = ($saved / $before) * 100

        return "{0} ({1:N1}%)" -f (Format-ByteSize $saved), $percent
    }

    function Get-UniqueRId {
        param(
            [System.Xml.XmlDocument]$relsDoc
        )

        # Force array collection to prevent pipeline leakage
        $existingIds = @($relsDoc.SelectNodes("//x:Relationship/@Id", $script:NsMgr) | ForEach-Object { $_.Value })
        $maxNum = 0

        foreach ($id in $existingIds) {
            if ($id -match '^rId(\d+)$') {
                $num = [int]$matches[1]
                if ($num -gt $maxNum) {
                    $maxNum = $num
                }
            }
        }

        return "rId$($maxNum + 1)"
    }

    function Test-SlideIncluded {
        param([int]$SlideNumber)

        # If IncludeSlides specified, slide must be in the list
        if ($IncludeSlides -and $IncludeSlides.Count -gt 0) {
            if ($SlideNumber -notin $IncludeSlides) {
                return $false
            }
        }

        # If ExcludeSlides specified, slide must not be in the list
        if ($ExcludeSlides -and $ExcludeSlides.Count -gt 0) {
            if ($SlideNumber -in $ExcludeSlides) {
                return $false
            }
        }

        return $true
    }

    #endregion

    # Heuristic: does a shape name look like think-cell content?
    # think-cell labels its shapes "think-cell data - do not delete" and "think-cell Slide".
    function Test-ThinkCellName {
        param([string]$Name)
        if (-not $Name) { return $false }
        return $Name -match 'think[\s-]?cell'
    }
