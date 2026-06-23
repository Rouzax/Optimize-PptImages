BeforeAll {
    Import-Module "$PSScriptRoot/../../Optimize-PptImages.psd1" -Force
    $script:gen = Join-Path $PSScriptRoot '../fixtures/New-AdobeRgbImage.ps1'
}

Describe 'Color-safe optimize of a non-sRGB image' {
    It 'converts an Adobe RGB image to sRGB (no profile, transformed pixels) not a naive strip' {
        $magick = (Get-Command magick).Source
        $dir = Join-Path ([System.IO.Path]::GetTempPath()) "colorsafe-$([guid]::NewGuid())"
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
        try {
            $src = Join-Path $dir 'adobe.jpg'
            & $script:gen -OutputPath $src

            # Reference: the CORRECT color-managed conversion (what optimize should match) and
            # the WRONG naive strip (what it must NOT match).
            $ref = Join-Path $dir 'ref_correct.jpg'
            $wrong = Join-Path $dir 'wrong_strip.jpg'
            $srgbIcc = InModuleScope Optimize-PptImages { $script:SrgbProfilePath }
            & $magick $src -profile $srgbIcc -strip $ref
            & $magick $src -strip $wrong

            # Run the optimizer's conversion path by building + running the job directly.
            $out = InModuleScope Optimize-PptImages -Parameters @{ src = $src; dir = $dir } {
                param($src, $dir)
                $script:MagickExe = Find-ImageMagick -ProvidedPath $null
                $u = [ImageUsage]::new(); $u.ImagePhysicalPath = $src; $u.OriginalFileName = 'adobe.jpg'
                $u.ContextType = [ContextType]::Slide; $u.SlideNumber = 1
                # Initialize probe fields so Get-OptimizationJob operates under StrictMode.
                $u.EffectiveTransparencyPercent = 0.0
                $u.SourceJpegQuality = 90
                $u.OptimizationStatus = 'Pending'
                $usages = [System.Collections.Generic.List[ImageUsage]]::new(); $usages.Add($u)
                Update-OptimizeProbesParallel -usages $usages
                # Source dims for the job.
                $facts = & $script:MagickExe identify -format '%w %h' $src
                $wh = "$facts".Trim() -split '\s+'; $u.SourceWidthPx = [int]$wh[0]; $u.SourceHeightPx = [int]$wh[1]
                $u.DisplayWidthPx = $u.SourceWidthPx; $u.DisplayHeightPx = $u.SourceHeightPx
                $g = [ImageGroup]::new(); $g.PhysicalPath = $src; $g.OriginalSizeBytes = (Get-Item $src).Length; $g.Usages.Add($u)
                $u.NeedsSrgbConversion | Should -BeTrue
                # Config vars required by Get-OptimizationJob under StrictMode.
                $HeadroomFactor = 2.0; $JpegQuality = 85; $TransparencyThresholdPercent = 0.1
                # Force a resize job by shrinking display so a job is produced.
                $job = Get-OptimizationJob -group $g -maxDisplay @{ Width = [int]($u.SourceWidthPx/2); Height = [int]($u.SourceHeightPx/2) } -scratchDir $dir
                & $script:MagickExe @($job.MagickArgs) | Out-Null
                $job.ScratchPath
            }

            # Output is sRGB and stripped (no embedded profile).
            (& $magick identify -format '%[profile:icc]' $out) | Should -BeNullOrEmpty
            (& $magick identify -format '%[colorspace]' $out) | Should -Be 'sRGB'

            # Prove the transform is color-managed, not a naive strip: $out must be closer
            # (lower MAE) to the color-managed reference than to the naive-strip reference.
            # Resize both references to $out's actual dimensions so the comparison is valid.
            $outDims      = (& $magick identify -format '%wx%h' $out).Trim()
            $refResized   = Join-Path $dir 'ref_resized.jpg'
            $wrongResized = Join-Path $dir 'wrong_resized.jpg'
            & $magick $ref   -resize "${outDims}!" $refResized   | Out-Null
            & $magick $wrong -resize "${outDims}!" $wrongResized | Out-Null
            $rawRef   = [string](& $magick compare -metric MAE $out $refResized   null: 2>&1)
            $rawWrong = [string](& $magick compare -metric MAE $out $wrongResized null: 2>&1)
            $maeRef   = [double](($rawRef   -split '\s+')[0])
            $maeWrong = [double](($rawWrong -split '\s+')[0])
            $maeRef | Should -BeLessThan $maeWrong
        } finally { Remove-Item -LiteralPath $dir -Recurse -Force -ErrorAction SilentlyContinue }
    }
}
