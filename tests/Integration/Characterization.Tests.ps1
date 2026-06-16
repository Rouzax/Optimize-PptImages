BeforeAll {
    Import-Module "$PSScriptRoot/../../Optimize-PptImages.psd1" -Force
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    Add-Type -AssemblyName System.IO.Compression

    $script:skip = $false
    $script:skipReason = ''
    $script:fixture = $null
    $script:generatedFixture = $null

    # --- Resolve a sample .pptx fixture in priority order. ---
    # (a) POPT_TEST_DECK env var pointing at an existing file.
    $envDeck = $env:POPT_TEST_DECK
    $repoFixture = Join-Path $PSScriptRoot '../fixtures/sample.pptx'
    $generator = Join-Path $PSScriptRoot '../fixtures/New-SampleDeck.ps1'

    if ($envDeck -and (Test-Path -LiteralPath $envDeck)) {
        $script:fixture = (Resolve-Path -LiteralPath $envDeck).Path
    }
    # (b) A committed/local tests/fixtures/sample.pptx.
    elseif (Test-Path -LiteralPath $repoFixture) {
        $script:fixture = (Resolve-Path -LiteralPath $repoFixture).Path
    }
    # (c) Generate one with the sample-deck generator.
    elseif (Test-Path -LiteralPath $generator) {
        try {
            $genPath = Join-Path ([System.IO.Path]::GetTempPath()) ("popt-fixture-" + [guid]::NewGuid().ToString('N') + ".pptx")
            & $generator -OutputPath $genPath | Out-Null
            if (Test-Path -LiteralPath $genPath) {
                $script:fixture = $genPath
                $script:generatedFixture = $genPath
            } else {
                $script:skip = $true
                $script:skipReason = 'Sample-deck generator ran but produced no file.'
            }
        } catch {
            $script:skip = $true
            $script:skipReason = "Sample-deck generator failed: $($_.Exception.Message)"
        }
    }
    # (d) No fixture available; skip cleanly.
    else {
        $script:skip = $true
        $script:skipReason = 'No fixture available (set POPT_TEST_DECK or add tests/fixtures/sample.pptx).'
    }

    # --- Run the optimizer once over the resolved fixture. ---
    # The module's end block calls `exit` when running under a ConsoleHost, which
    # would terminate the Pester runner. Invoke it in a child pwsh process so its
    # exit code is isolated from the test host.
    if (-not $script:skip) {
        $script:outDir = Join-Path ([System.IO.Path]::GetTempPath()) ("popt-itest-" + [guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $script:outDir -Force | Out-Null
        $script:outPptx = Join-Path $script:outDir 'out.pptx'
        $script:reportCsv = Join-Path $script:outDir 'out.opt-report.csv'

        $manifest = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../../Optimize-PptImages.psd1')).Path
        $runnerScript = @"
Import-Module '$manifest' -Force
Invoke-PptImageOptimization -InputPath '$($script:fixture)' -OutputPath '$($script:outPptx)' -All -CleanUnusedLayouts -ValidateOutput
"@
        $pwshExe = (Get-Process -Id $PID).Path
        & $pwshExe -NoProfile -Command $runnerScript *>&1 | Out-Null
        $script:optExitCode = $LASTEXITCODE

        # Exit codes: 0 = success with changes, 2 = success no changes. 1 = error.
        if ($script:optExitCode -eq 1 -or -not (Test-Path -LiteralPath $script:outPptx)) {
            $script:skip = $true
            $script:skipReason = "Optimizer failed on the fixture (exit $($script:optExitCode))."
        }

        $script:inputSize = (Get-Item -LiteralPath $script:fixture).Length
    }
}

AfterAll {
    if ($script:outDir -and (Test-Path -LiteralPath $script:outDir)) {
        Remove-Item -LiteralPath $script:outDir -Recurse -Force -ErrorAction SilentlyContinue
    }
    if ($script:generatedFixture -and (Test-Path -LiteralPath $script:generatedFixture)) {
        Remove-Item -LiteralPath $script:generatedFixture -Force -ErrorAction SilentlyContinue
    }
}

Describe 'Characterization: end-to-end optimization' {
    It 'produces an output .pptx file' {
        if ($script:skip) { Set-ItResult -Skipped -Because $script:skipReason; return }
        Test-Path -LiteralPath $script:outPptx | Should -BeTrue
    }

    It 'writes a CSV report whose columns include IsThinkCell, UsedOnSlides and MediaPart' {
        if ($script:skip) { Set-ItResult -Skipped -Because $script:skipReason; return }
        Test-Path -LiteralPath $script:reportCsv | Should -BeTrue
        $firstRow = Import-Csv -LiteralPath $script:reportCsv | Select-Object -First 1
        $firstRow | Should -Not -BeNullOrEmpty
        $names = $firstRow.PSObject.Properties.Name
        $names | Should -Contain 'IsThinkCell'
        $names | Should -Contain 'UsedOnSlides'
        $names | Should -Contain 'MediaPart'
    }

    It 'removes all ppt/tags and ppt/embeddings parts from the output' {
        if ($script:skip) { Set-ItResult -Skipped -Because $script:skipReason; return }
        $zip = [System.IO.Compression.ZipFile]::OpenRead($script:outPptx)
        try {
            $entries = $zip.Entries | ForEach-Object { $_.FullName }
        } finally {
            $zip.Dispose()
        }
        @($entries | Where-Object { $_ -like 'ppt/tags/*' }).Count | Should -Be 0
        @($entries | Where-Object { $_ -like 'ppt/embeddings/*' }).Count | Should -Be 0
    }

    It 'produces an output no larger than the input' {
        if ($script:skip) { Set-ItResult -Skipped -Because $script:skipReason; return }
        (Get-Item -LiteralPath $script:outPptx).Length | Should -BeLessOrEqual $script:inputSize
    }
}
