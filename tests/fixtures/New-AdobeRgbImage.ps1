param([Parameter(Mandatory)][string]$OutputPath)
$magick = (Get-Command magick).Source
$adobe = '/usr/share/color/icc/colord/AdobeRGB1998.icc'
if (-not (Test-Path -LiteralPath $adobe)) { throw "AdobeRGB profile not found at $adobe (needed only to build the test input)" }
& $magick -size 64x48 'gradient:red-blue' -profile $adobe $OutputPath
if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $OutputPath)) { throw "Failed to create AdobeRGB test image" }
