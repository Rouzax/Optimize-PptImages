    #region Relationship Helpers

    function Update-BlipRelationship {
        param(
            [ImageUsage]$usage,
            [string]$tempDir,
            [string]$newMediaFileName
        )

        $partDir = Split-Path $usage.PartPath -Parent
        $partFile = Split-Path $usage.PartPath -Leaf
        $relsPath = Join-Path $tempDir (($partDir + '/_rels/' + $partFile + '.rels'))

        if (-not (Test-Path -LiteralPath $relsPath)) {
            throw "Rels file not found: $relsPath"
        }

        $relsDoc = Get-XmlDocument $relsPath

        # Create new relationship (old one is left intact; other blips on the same
        # part may still reference it. Orphaned rels are cleaned up by Invoke-MediaCleanup.)
        $newRId = Get-UniqueRId -relsDoc $relsDoc
        $rel = $relsDoc.CreateElement('Relationship', $relsDoc.DocumentElement.NamespaceURI)
        [void]$rel.SetAttribute('Id', $newRId)
        [void]$rel.SetAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image')
        [void]$rel.SetAttribute('Target', "../media/$newMediaFileName")
        $null = $relsDoc.DocumentElement.AppendChild($rel)
        Save-XmlDocument -doc $relsDoc -path $relsPath

        # Update blip element
        [void]$usage.BlipElement.SetAttribute('embed', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships', $newRId)

        # Update usage tracking
        $usage.BlipRId = $newRId

        return $newRId
    }

    function New-UniqueMediaPath {
        param(
            [string]$basePath,
            [string]$suffix,
            [string]$newExtension = $null
        )

        $dir = Split-Path $basePath -Parent
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($basePath)
        $ext = if ($newExtension) { $newExtension } else { [System.IO.Path]::GetExtension($basePath) }

        $outputName = "${baseName}${suffix}$ext"
        $outputPath = Join-Path $dir $outputName

        $counter = 1
        while (Test-Path $outputPath) {
            $outputName = "${baseName}${suffix}${counter}$ext"
            $outputPath = Join-Path $dir $outputName
            $counter++
        }

        return @{
            Path = $outputPath
            Name = $outputName
        }
    }

    #endregion

    # Resolve a relationship Target (relative to the part that owns the .rels file)
    # to a package-relative part path using forward slashes, e.g. 'ppt/tags/tag1.xml'.
    # Returns $null for external targets or anything outside the package.
    function Resolve-RelationshipTarget {
        param(
            [string]$tempDir,
            [System.IO.FileInfo]$relsFile,
            [string]$target
        )
        if (-not $target) { return $null }
        # Owning part directory is the parent of the _rels folder.
        $owningPartDir = Split-Path (Split-Path $relsFile.FullName -Parent) -Parent
        if ($target.StartsWith('/')) {
            # Absolute package path (rare): resolve against package root.
            $combined = Join-Path $tempDir ($target.TrimStart('/'))
        } else {
            $combined = Join-Path $owningPartDir $target
        }
        $full = [System.IO.Path]::GetFullPath($combined)
        $root = [System.IO.Path]::GetFullPath($tempDir)
        if (-not $full.StartsWith($root, [System.StringComparison]::OrdinalIgnoreCase)) {
            return $null
        }
        $rel = $full.Substring($root.Length).TrimStart('\', '/')
        return ($rel -replace '\\', '/')
    }
