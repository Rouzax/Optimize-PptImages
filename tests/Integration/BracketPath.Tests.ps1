BeforeAll { Import-Module "$PSScriptRoot/../../Optimize-PptImages.psd1" -Force }

Describe 'Bracket-safe paths' {
    It 'runs on an input and output path containing [ ] wildcard characters' {
        $src = Join-Path $PSScriptRoot '../fixtures/content-id-sample.pptx'
        $dir = Join-Path ([System.IO.Path]::GetTempPath()) "bracket-$([guid]::NewGuid())"
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
        try {
            $bin  = Join-Path $dir 'deck-[x].pptx'
            $bout = Join-Path $dir 'deck-[x].out.pptx'
            Copy-Item -LiteralPath $src -Destination $bin
            { Invoke-PptImageOptimization -InputPath $bin -OutputPath $bout -OptimizeSlides } | Should -Not -Throw
            Test-Path -LiteralPath $bout | Should -BeTrue
        } finally { Remove-Item -LiteralPath $dir -Recurse -Force -ErrorAction SilentlyContinue }
    }
}
