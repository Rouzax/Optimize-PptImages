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
