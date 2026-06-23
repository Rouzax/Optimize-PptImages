    #region ImageMagick Helpers

    function Get-ImageDimensions {
        param([string]$ImagePath)

        $identify = & $script:MagickExe identify -format "%w %h" $ImagePath 2>&1
        if ($LASTEXITCODE -ne 0) {
            throw "ImageMagick identify failed for '$ImagePath': $identify"
        }

        if ($identify -is [array]) {
            $identify = $identify[0]
        }

        $dims = $identify -split '\s+'
        return @{
            Width = [int]$dims[0]
            Height = [int]$dims[1]
        }
    }

    # Parse the output of `magick identify -format "%w %h %[opaque]"` into structured
    # facts. Pure (no magick invocation) so the parallel scan can call magick in
    # runspaces and parse the raw strings here, sequentially and unit-testably.
    function ConvertFrom-MagickFacts {
        param([string]$Raw)

        $parts = "$Raw".Trim() -split '\s+'
        return @{
            Width    = [int]$parts[0]
            Height   = [int]$parts[1]
            IsOpaque = ($parts[2] -eq 'True')
        }
    }

    # Pure decision: does this image need a color-managed conversion to sRGB before -strip?
    # True for CMYK, or for an embedded ICC profile whose description is not sRGB. Untagged
    # and sRGB/grayscale images return false (strip as today is safe for them).
    function Test-NeedsSrgbConversion {
        param([string]$Colorspace, [string]$IccDescription)
        if ($Colorspace -eq 'CMYK') { return $true }
        if ($IccDescription -and $IccDescription.Trim() -ne '' -and $IccDescription -notmatch '(?i)sRGB') { return $true }
        return $false
    }

    function Get-JpegQuality {
        param([string]$ImagePath)

        $quality = & $script:MagickExe identify -format "%Q" $ImagePath 2>&1
        if ($LASTEXITCODE -ne 0) {
            return -1  # Unknown
        }

        if ($quality -is [array]) {
            $quality = $quality[0]
        }

        if ($quality -match '^\d+$') {
            return [int]$quality
        }

        return -1
    }

    function Test-IsAnimatedGif {
        param([string]$ImagePath)

        $frameCount = & $script:MagickExe identify -format "%n\n" $ImagePath 2>&1 |
            Where-Object { $_ -match '^\d+$' } |
            Select-Object -First 1

        if ($frameCount -and [int]$frameCount -gt 1) {
            return $true
        }

        return $false
    }

    # Generic parallel magick runner. Each job exposes .Key, .MagickArgs (with its scratch
    # path as the output), and .ScratchPath. Runspaces invoke only magick and return raw
    # results; nothing module-specific runs inside the runspace. Returns a hashtable keyed
    # by Key. Progress is rendered from the parent via the streaming-downstream pattern.
    function Invoke-MagickJobsParallel {
        param(
            [object[]]$jobs,
            [string]$activity,
            [int]$progressId
        )

        $resultMap = @{}
        if (-not $jobs -or $jobs.Count -eq 0) { return $resultMap }

        $magickExe = $script:MagickExe
        $total = $jobs.Count
        $rawResults = @($jobs) | ForEach-Object -Parallel {
            $job = $_
            & $using:magickExe @($job.MagickArgs) 2>&1 | Out-Null
            $success   = ($LASTEXITCODE -eq 0 -and (Test-Path -LiteralPath $job.ScratchPath))
            $afterSize = if ($success) { (Get-Item -LiteralPath $job.ScratchPath).Length } else { 0 }
            [PSCustomObject]@{ Key = $job.Key; ScratchPath = $job.ScratchPath; Success = $success; AfterSize = $afterSize }
        } -ThrottleLimit ([Environment]::ProcessorCount) | ForEach-Object -Begin { $count = 0 } -Process {
            $count++
            Write-Progress -Activity $activity -Status "Running: $count of $total" `
                -PercentComplete (($count / [Math]::Max(1, $total)) * 100) -Id $progressId
            $_
        }
        Write-Progress -Activity $activity -Completed -Id $progressId

        foreach ($r in $rawResults) {
            $resultMap[$r.Key] = @{ Success = $r.Success; AfterSize = $r.AfterSize; ScratchPath = $r.ScratchPath }
        }
        return $resultMap
    }

    #endregion

    function Find-ImageMagick {
        param([string]$ProvidedPath)

        if ($ProvidedPath) {
            if (Test-Path -LiteralPath $ProvidedPath) {
                return $ProvidedPath
            }
            return $null
        }

        # Search PATH
        $magick = Get-Command 'magick' -ErrorAction SilentlyContinue
        if ($magick) {
            return $magick.Source
        }

        # Common locations
        $commonPaths = @(
            'C:\Program Files\ImageMagick-*\magick.exe',
            'C:\Program Files (x86)\ImageMagick-*\magick.exe',
            '/usr/bin/magick',
            '/usr/local/bin/magick',
            '/opt/homebrew/bin/magick'
        )

        foreach ($pattern in $commonPaths) {
            $found = Get-Item $pattern -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($found) {
                return $found.FullName
            }
        }

        return $null
    }

    #endregion
