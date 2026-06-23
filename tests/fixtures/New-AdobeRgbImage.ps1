param([Parameter(Mandatory)][string]$OutputPath)
$magick = (Get-Command magick).Source
# Bundled non-sRGB profile (free-redistributable "Compatible with Adobe RGB (1998)" from
# icc-profiles-free) so this test input builds on every platform, not just Linux dev boxes.
$adobe = Join-Path $PSScriptRoot 'AdobeRGB1998.icc'
if (-not (Test-Path -LiteralPath $adobe)) { throw "Bundled AdobeRGB profile not found at $adobe" }
& $magick -size 64x48 'gradient:red-blue' -profile $adobe $OutputPath
if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $OutputPath)) { throw "Failed to create AdobeRGB test image" }
