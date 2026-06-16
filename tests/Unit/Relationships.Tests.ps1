BeforeAll {
    Import-Module "$PSScriptRoot/../../Optimize-PptImages.psd1" -Force
}

Describe 'Resolve-RelationshipTarget' {
    BeforeAll {
        $script:tempDir = Join-Path ([System.IO.Path]::GetTempPath()) ("popt-rels-" + [guid]::NewGuid().ToString('N'))
        $relsDir = Join-Path $script:tempDir 'ppt/slideLayouts/_rels'
        New-Item -ItemType Directory -Path $relsDir -Force | Out-Null
        $script:relsFile = Get-Item (New-Item -ItemType File -Path (Join-Path $relsDir 'slideLayout1.xml.rels') -Force).FullName
    }

    AfterAll {
        if ($script:tempDir -and (Test-Path -LiteralPath $script:tempDir)) {
            Remove-Item -LiteralPath $script:tempDir -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    It 'resolves a relative ../tags/tag1.xml target to ppt/tags/tag1.xml' {
        $td = $script:tempDir
        $rf = $script:relsFile
        InModuleScope Optimize-PptImages -Parameters @{ TempDir = $td; RelsFile = $rf } {
            param($TempDir, $RelsFile)
            Resolve-RelationshipTarget -tempDir $TempDir -relsFile $RelsFile -target '../tags/tag1.xml' |
                Should -Be 'ppt/tags/tag1.xml'
        }
    }

    It 'returns $null for a target that escapes the package root' {
        # Note: the function does not detect external (https) URLs itself; callers skip
        # those by checking TargetMode='External' before calling. The function only
        # returns $null for empty targets or targets that resolve outside the package
        # root. A path with enough leading "../" segments escapes the temp root.
        $td = $script:tempDir
        $rf = $script:relsFile
        InModuleScope Optimize-PptImages -Parameters @{ TempDir = $td; RelsFile = $rf } {
            param($TempDir, $RelsFile)
            Resolve-RelationshipTarget -tempDir $TempDir -relsFile $RelsFile -target '../../../../../../etc/passwd' |
                Should -BeNullOrEmpty
        }
    }
}
