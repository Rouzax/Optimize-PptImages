BeforeAll {
    Import-Module "$PSScriptRoot/../../Optimize-PptImages.psd1" -Force
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    Add-Type -AssemblyName System.IO.Compression

    $script:skip = $false
    $script:skipReason = ''
    $script:inDir = $null

    # Batch processing needs MULTIPLE input decks, so always generate three with the
    # sample-deck generator. If the generator is unavailable or fails (e.g. no OpenXML
    # SDK on the runner), skip cleanly, matching the characterization test's behavior.
    $generator = Join-Path $PSScriptRoot '../fixtures/New-SampleDeck.ps1'
    if (-not (Test-Path -LiteralPath $generator)) {
        $script:skip = $true
        $script:skipReason = 'Sample-deck generator not found.'
    } else {
        $script:inDir = Join-Path ([System.IO.Path]::GetTempPath()) ("popt-batch-" + [guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $script:inDir -Force | Out-Null
        try {
            1..3 | ForEach-Object { & $generator -OutputPath (Join-Path $script:inDir "deck$_.pptx") | Out-Null }
            if (@(Get-ChildItem -LiteralPath $script:inDir -Filter 'deck*.pptx').Count -ne 3) {
                $script:skip = $true
                $script:skipReason = 'Generator did not produce three input decks.'
            }
        } catch {
            $script:skip = $true
            $script:skipReason = "Sample-deck generator failed: $($_.Exception.Message)"
        }
    }

    # Run the three decks through the module via the pipeline (batch mode). In batch mode
    # the end block does not call `exit`, but invoke in a child pwsh anyway to isolate the
    # host and to capture the console batch summary plus the -PassThru result count.
    if (-not $script:skip) {
        $manifest = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../../Optimize-PptImages.psd1')).Path
        # Enumerate inputs into an array BEFORE the run so the streamed pipeline never
        # picks up the *.optimized.pptx files written during processing.
        $runnerScript = @"
Import-Module '$manifest' -Force
`$inputs = @(Get-ChildItem -LiteralPath '$($script:inDir)' -Filter 'deck*.pptx')
`$results = `$inputs | Invoke-PptImageOptimization -OptimizeSlides -CleanUnusedLayouts -PassThru
Write-Output "PASSTHRU_COUNT=`$(@(`$results).Count)"
"@
        $pwshExe = (Get-Process -Id $PID).Path
        $script:batchOutput = (& $pwshExe -NoProfile -Command $runnerScript *>&1 | Out-String)
        $script:batchExit = $LASTEXITCODE
        $script:outputs = @(Get-ChildItem -LiteralPath $script:inDir -Filter '*.optimized.pptx')
    }
}

AfterAll {
    if ($script:inDir -and (Test-Path -LiteralPath $script:inDir)) {
        Remove-Item -LiteralPath $script:inDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Describe 'Batch processing: multiple decks via the pipeline' {
    It 'prints a batch summary reporting all three files processed' {
        if ($script:skip) { Set-ItResult -Skipped -Because $script:skipReason; return }
        $script:batchOutput | Should -Match 'BATCH SUMMARY'
        $script:batchOutput | Should -Match 'Files processed:\s*3'
        $script:batchOutput | Should -Match 'Successful:\s*3'
    }

    It 'returns one PassThru result object per input file' {
        if ($script:skip) { Set-ItResult -Skipped -Because $script:skipReason; return }
        $script:batchOutput | Should -Match 'PASSTHRU_COUNT=3'
    }

    It 'writes one optimized output per input deck' {
        if ($script:skip) { Set-ItResult -Skipped -Because $script:skipReason; return }
        $script:outputs.Count | Should -Be 3
    }

    It 'produces structurally valid .pptx outputs (openable ZIP with core OOXML parts)' {
        if ($script:skip) { Set-ItResult -Skipped -Because $script:skipReason; return }
        foreach ($out in $script:outputs) {
            $zip = [System.IO.Compression.ZipFile]::OpenRead($out.FullName)
            try {
                $entries = $zip.Entries | ForEach-Object { $_.FullName }
            } finally {
                $zip.Dispose()
            }
            $entries | Should -Contain '[Content_Types].xml'
            $entries | Should -Contain 'ppt/presentation.xml'
        }
    }
}
