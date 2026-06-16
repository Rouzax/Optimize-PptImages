    #region Cleanup and Packaging


    # Remove non-media parts that nothing references anymore. Generalizes the media
    # cleanup to satellite parts (ppt/tags, ppt/embeddings) that are left orphaned when
    # -CleanUnusedLayouts deletes the layouts/masters that referenced them (e.g. think-cell
    # marker tags and CSmartGrid OLE objects). Also prunes their [Content_Types].xml overrides.
    # Runs unconditionally; it only removes parts no remaining relationship points to.
    function Invoke-OrphanedPartCleanup {
        param([string]$tempDir)

        Write-Host "`n[CLEAN] Removing orphaned parts (tags/embeddings)..." -ForegroundColor Yellow

        $candidateDirs = @('ppt/tags', 'ppt/embeddings')
        $existingDirs = @($candidateDirs | Where-Object { Test-Path -LiteralPath (Join-Path $tempDir $_) })
        if ($existingDirs.Count -eq 0) {
            Write-Host "[CLEAN] No tags/embeddings parts present" -ForegroundColor Green
            return
        }

        # Build the set of every part path referenced by any relationship in the package.
        $referenced = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        $relsFiles = Get-ChildItem -Path $tempDir -Filter '*.rels' -Recurse
        foreach ($relsFile in $relsFiles) {
            $relsDoc = Get-XmlDocument $relsFile.FullName
            $rels = $relsDoc.SelectNodes('//x:Relationship', $script:NsMgr)
            foreach ($rel in $rels) {
                if ($rel.GetAttribute('TargetMode') -eq 'External') { continue }
                $resolved = Resolve-RelationshipTarget -tempDir $tempDir -relsFile $relsFile -target $rel.GetAttribute('Target')
                if ($resolved) { $null = $referenced.Add($resolved) }
            }
        }

        $contentTypesPath = Join-Path $tempDir '[Content_Types].xml'
        $contentTypesDoc = Get-XmlDocument $contentTypesPath
        $contentTypesModified = $false

        $removedByDir = @{}
        foreach ($dir in $existingDirs) {
            $removedByDir[$dir] = 0
            $partDirPath = Join-Path $tempDir $dir
            foreach ($file in (Get-ChildItem -Path $partDirPath -File)) {
                $partRel = "$dir/$($file.Name)"
                if ($referenced.Contains($partRel)) { continue }

                Remove-Item -LiteralPath $file.FullName -Force
                $removedByDir[$dir]++

                # Remove the part's own .rels if present.
                $partRelsPath = Join-Path $partDirPath "_rels/$($file.Name).rels"
                if (Test-Path -LiteralPath $partRelsPath) {
                    Remove-Item -LiteralPath $partRelsPath -Force
                }

                # Prune the matching [Content_Types].xml Override (leave shared Defaults).
                $override = $contentTypesDoc.SelectSingleNode(
                    "//ct:Override[@PartName='/$partRel']",
                    $script:NsMgr
                )
                if ($override) {
                    $null = $override.ParentNode.RemoveChild($override)
                    $contentTypesModified = $true
                }

                Write-Verbose "  Removed orphaned part: $partRel"
            }
        }

        if ($contentTypesModified) {
            Save-XmlDocument -doc $contentTypesDoc -path $contentTypesPath
        }

        $total = ($removedByDir.Values | Measure-Object -Sum).Sum
        if ($total -gt 0) {
            $detail = ($existingDirs | ForEach-Object { "$($_ -replace '^ppt/', ''): $($removedByDir[$_])" }) -join ', '
            Write-Host "[CLEAN] Orphaned parts removed: $total ($detail)" -ForegroundColor Green
        } else {
            Write-Host "[CLEAN] No orphaned parts found" -ForegroundColor Green
        }
    }

    function Invoke-MediaCleanup {
        param([string]$tempDir)

        Write-Host "`n[CLEAN] Cleaning up unreferenced media..." -ForegroundColor Yellow

        $relsFiles = Get-ChildItem -Path $tempDir -Filter "*.rels" -Recurse
        $mediaDir = Join-Path $tempDir 'ppt/media'
        Write-Verbose "Scanning $($relsFiles.Count) relationship files..."

        # Phase 1: Remove orphaned rels entries (rId not referenced by any
        # element in the parent part XML, or target file missing on disk).
        # This must run before media deletion so that stale rels entries
        # don't keep unreferenced files alive.
        $orphanedRelsCount = 0
        foreach ($relsFile in $relsFiles) {
            $relsDoc = Get-XmlDocument $relsFile.FullName

            $imageRels = $relsDoc.SelectNodes('//x:Relationship[@Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"]', $script:NsMgr)
            $hdPhotoRels = $relsDoc.SelectNodes('//x:Relationship[@Type="http://schemas.microsoft.com/office/2007/relationships/hdphoto"]', $script:NsMgr)

            $allRels = @($imageRels) + @($hdPhotoRels)
            if (-not $allRels -or $allRels.Count -eq 0) { continue }

            $relsDir = Split-Path $relsFile.FullName -Parent
            $partDir = Split-Path $relsDir -Parent
            $partFileName = $relsFile.Name -replace '\.rels$', ''
            $partPath = Join-Path $partDir $partFileName
            $partContent = $null
            if (Test-Path -LiteralPath $partPath) {
                $partContent = [System.IO.File]::ReadAllText($partPath)
            }

            $modified = $false
            foreach ($rel in $allRels) {
                if (-not $rel) { continue }
                $rId = $rel.GetAttribute('Id')
                $target = $rel.GetAttribute('Target')

                $isOrphaned = $false

                $mediaFile = $target -replace '^\.\./media/', ''
                $fullPath = Join-Path $mediaDir $mediaFile
                if (-not (Test-Path -LiteralPath $fullPath)) {
                    $isOrphaned = $true
                }

                if (-not $isOrphaned -and $partContent -and $partContent -notmatch [regex]::Escape("`"$rId`"")) {
                    $isOrphaned = $true
                }

                if ($isOrphaned) {
                    $null = $rel.ParentNode.RemoveChild($rel)
                    $orphanedRelsCount++
                    $modified = $true
                    Write-Verbose "  Removed orphaned rels entry: $rId -> $target (in $($relsFile.Name))"
                }
            }

            if ($modified) {
                Save-XmlDocument -doc $relsDoc -path $relsFile.FullName
            }
        }

        if ($orphanedRelsCount -gt 0) {
            Write-Host "[CLEAN] Orphaned relationship entries removed: $orphanedRelsCount" -ForegroundColor Green
        }

        # Phase 2: Delete media files not referenced by any remaining rels entry.
        $referenced = [System.Collections.Generic.HashSet[string]]::new()
        foreach ($relsFile in $relsFiles) {
            $relsDoc = Get-XmlDocument $relsFile.FullName

            $imageRels = $relsDoc.SelectNodes('//x:Relationship[@Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"]', $script:NsMgr)
            $hdPhotoRels = $relsDoc.SelectNodes('//x:Relationship[@Type="http://schemas.microsoft.com/office/2007/relationships/hdphoto"]', $script:NsMgr)

            foreach ($rel in $imageRels) {
                $target = $rel.GetAttribute('Target')
                $mediaFile = $target -replace '^\.\./media/', ''
                $null = $referenced.Add($mediaFile)
            }

            foreach ($rel in $hdPhotoRels) {
                $target = $rel.GetAttribute('Target')
                $mediaFile = $target -replace '^\.\./media/', ''
                $null = $referenced.Add($mediaFile)
            }
        }

        Write-Verbose "Found $($referenced.Count) referenced media file(s)"

        if (Test-Path -LiteralPath $mediaDir) {
            $allMedia = @(Get-ChildItem -Path $mediaDir)
            Write-Verbose "Checking $($allMedia.Count) media file(s) for references..."

            $deletedCount = 0

            foreach ($file in $allMedia) {
                if (-not $referenced.Contains($file.Name)) {
                    Remove-Item $file.FullName -Force
                    $deletedCount++
                    Write-Verbose "  Deleted unreferenced: $($file.Name)"
                }
            }

            Write-Host "[CLEAN] Media files removed: $deletedCount" -ForegroundColor Green
        }
    }

    function Update-ContentTypes {
        param([string]$tempDir)
        
        Write-Verbose "Updating [Content_Types].xml to ensure required content types..."
        
        $contentTypesPath = Join-Path $tempDir '[Content_Types].xml'
        
        $doc = Get-XmlDocument $contentTypesPath
        if (-not $doc) {
            return
        }
        
        # Ensure required defaults exist
        $requiredDefaults = @{
            'jpg' = 'image/jpeg'
            'jpeg' = 'image/jpeg'
            'png' = 'image/png'
            'svg' = 'image/svg+xml'
        }
        
        $addedCount = 0
        foreach ($ext in $requiredDefaults.Keys) {
            $existing = $doc.SelectSingleNode("//ct:Default[@Extension='$ext']", $script:NsMgr)
            if (-not $existing) {
                $default = $doc.CreateElement('Default', $doc.DocumentElement.NamespaceURI)
                [void]$default.SetAttribute('Extension', $ext)
                [void]$default.SetAttribute('ContentType', $requiredDefaults[$ext])
                $null = $doc.DocumentElement.AppendChild($default)
                Write-Verbose "  Added content type: .$ext -> $($requiredDefaults[$ext])"
                $addedCount++
            }
        }
        
        if ($addedCount -eq 0) {
            Write-Verbose "  All required content types already present"
        }
        
        Save-XmlDocument -doc $doc -path $contentTypesPath
    }

    function Invoke-Repack {
        param(
            [string]$tempDir
        )
        
        Write-Host "`n[PACK] Repacking presentation..." -ForegroundColor Yellow
        Write-Verbose "Creating optimized ZIP archive from: $tempDir"
        
        # Create temp file for output (outside the extraction temp dir)
        $tempOutputPath = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "ppt-opt-$([System.Guid]::NewGuid().ToString('N').Substring(0,8)).pptx")
        Write-Verbose "Temp output file: $tempOutputPath"
        
        # Force garbage collection to close any file handles
        Write-Verbose "Forcing garbage collection to release file handles..."
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        [System.GC]::Collect()
        
        try {
            # Use .NET ZipFile for better control
            Add-Type -Assembly 'System.IO.Compression.FileSystem'
            Write-Verbose "Creating ZIP archive with Optimal compression..."
            [System.IO.Compression.ZipFile]::CreateFromDirectory(
                $tempDir, 
                $tempOutputPath, 
                [System.IO.Compression.CompressionLevel]::Optimal, 
                $false  # Don't include base directory
            )
            
            # Verify the output file was created and is valid
            if (-not (Test-Path -LiteralPath $tempOutputPath)) {
                throw "Output file was not created: $tempOutputPath"
            }
            
            $outputSize = (Get-Item -LiteralPath $tempOutputPath).Length
            Write-Verbose "Repack complete: output file is $(Format-ByteSize $outputSize)"
            
            $fileInfo = Get-Item -LiteralPath $tempOutputPath
            if ($fileInfo.Length -eq 0) {
                throw "Output file is empty: $tempOutputPath"
            }
            
            Write-Host "[OK] Repacked: $([Math]::Round($fileInfo.Length / 1KB, 2)) KB" -ForegroundColor Green
            
            return $tempOutputPath
            
        } catch {
            throw "Failed to create ZIP archive: $_"
        }
    }


    #endregion
