BeforeAll {
    Import-Module "$PSScriptRoot/../../Optimize-PptImages.psd1" -Force
    $script:deck = Join-Path $PSScriptRoot '../fixtures/crop-sample.pptx'
    $script:baselineFile = Join-Path $PSScriptRoot '../fixtures/crop-identity-baseline.json'

    function Get-PackageManifest {
        param([string]$pptxPath)
        $work = Join-Path ([System.IO.Path]::GetTempPath()) "cropman-$([guid]::NewGuid())"
        New-Item -ItemType Directory -Path $work -Force | Out-Null
        try {
            Add-Type -Assembly 'System.IO.Compression.FileSystem'
            [System.IO.Compression.ZipFile]::ExtractToDirectory($pptxPath, $work)
            $entries = [ordered]@{}
            $targets = @()
            $md = Join-Path $work 'ppt/media'
            if (Test-Path -LiteralPath $md) { $targets += Get-ChildItem -LiteralPath $md -File }
            foreach ($sub in 'ppt/slides', 'ppt/slideLayouts', 'ppt/slideMasters') {
                $d = Join-Path $work $sub
                if (Test-Path -LiteralPath $d) { $targets += Get-ChildItem -LiteralPath $d -Recurse -File -Include '*.xml', '*.rels' }
            }
            foreach ($f in ($targets | Sort-Object { $_.FullName.Substring($work.Length).Replace('\', '/') })) {
                $rel = $f.FullName.Substring($work.Length).Replace('\', '/').TrimStart('/')
                $entries[$rel] = (Get-FileHash -LiteralPath $f.FullName -Algorithm SHA256).Hash
            }
            return $entries
        } finally { Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue }
    }
}

Describe 'Content identity of the crop phase' {
    It 'reproduces the crop baseline manifest for -All' {
        $out = Join-Path ([System.IO.Path]::GetTempPath()) "crop-id-out-$([guid]::NewGuid()).pptx"
        try {
            Invoke-PptImageOptimization -InputPath $script:deck -OutputPath $out -All -PassThru | Out-Null
            $manifest = Get-PackageManifest -pptxPath $out
            $expected = Get-Content -LiteralPath $script:baselineFile -Raw | ConvertFrom-Json
            foreach ($key in $manifest.Keys) { $manifest[$key] | Should -Be $expected.$key -Because "content of $key must match the sequential baseline" }
            @($manifest.Keys).Count | Should -Be @($expected.PSObject.Properties).Count
        } finally { if (Test-Path -LiteralPath $out) { Remove-Item -LiteralPath $out -Force -ErrorAction SilentlyContinue } }
    }
}
