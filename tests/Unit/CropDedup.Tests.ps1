BeforeAll { Import-Module "$PSScriptRoot/../../Optimize-PptImages.psd1" -Force }

Describe 'Crop deduplication' {
    It 'two usages with the same source and geometry share one materialized file; a third (different geometry) gets its own' {
        InModuleScope Optimize-PptImages {
            $script:MagickExe = Find-ImageMagick -ProvidedPath $null
            $script:CropSlides = $true
            $script:CropMastersAndLayouts = $true
            $dir = Join-Path ([System.IO.Path]::GetTempPath()) "dedup-$([guid]::NewGuid())"
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
            try {
                # One source image shared by all usages.
                $src = Join-Path $dir 'shared.png'
                & $script:MagickExe -size 400x400 'gradient:red-blue' $src
                $geomA = '320x320+40+40'   # 10% inset on all sides of 400x400
                $jobs = foreach ($i in 0..2) {
                    $j = [CropJob]::new()
                    $j.Key = "part|Shape$i|rId$i"
                    $j.SourcePath = $src
                    $geom = if ($i -lt 2) { $geomA } else { '300x300+50+50' }  # 0,1 identical; 2 different
                    $scratch = Join-Path $dir "scratch$i.png"
                    $j.ScratchPath = $scratch
                    $j.MagickArgs = @($src, '-crop', $geom, '+repage', $scratch)
                    $j.CropKey = "$src|$geom"
                    $j
                }
                # Distinct CropKeys -> distinct leaders (the build-pass rule).
                $leaders = @{}
                foreach ($j in $jobs) { if (-not $leaders.ContainsKey($j.CropKey)) { $leaders[$j.CropKey] = $j } }
                $leaders.Count | Should -Be 2   # two unique crops from three usages

                # The two identical-geometry jobs share one CropKey.
                ($jobs[0].CropKey) | Should -Be ($jobs[1].CropKey)
                ($jobs[0].CropKey) | Should -Not -Be ($jobs[2].CropKey)
            } finally { Remove-Item -LiteralPath $dir -Recurse -Force -ErrorAction SilentlyContinue }
        }
    }
}
