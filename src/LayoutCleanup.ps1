    #region Layout Cleanup

    function Invoke-UnusedLayoutCleanup {
    [CmdletBinding(SupportsShouldProcess)]
        param([string]$tempDir)

        Write-Host "`n[CLEAN] Removing unused layouts and masters..." -ForegroundColor Yellow

        $slidesDir = Join-Path $tempDir 'ppt/slides'
        $layoutsDir = Join-Path $tempDir 'ppt/slideLayouts'
        $mastersDir = Join-Path $tempDir 'ppt/slideMasters'
        $contentTypesPath = Join-Path $tempDir '[Content_Types].xml'
        $presentationXmlPath = Join-Path $tempDir 'ppt/presentation.xml'
        $presentationRelsPath = Join-Path $tempDir 'ppt/_rels/presentation.xml.rels'

        # Step 1: Collect used layout filenames from slide rels
        $usedLayouts = [System.Collections.Generic.HashSet[string]]::new()
        $slideRelsDir = Join-Path $slidesDir '_rels'
        if (Test-Path -LiteralPath $slideRelsDir) {
            foreach ($relsFile in (Get-ChildItem -Path $slideRelsDir -Filter '*.rels')) {
                $relsDoc = Get-XmlDocument $relsFile.FullName
                $layoutRel = $relsDoc.SelectSingleNode(
                    "//x:Relationship[@Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout']",
                    $script:NsMgr
                )
                if ($layoutRel) {
                    $layoutFile = Split-Path $layoutRel.GetAttribute('Target') -Leaf
                    $null = $usedLayouts.Add($layoutFile)
                }
            }
        }

        # Step 2: Find unused layouts
        $allLayouts = @(Get-ChildItem -Path $layoutsDir -Filter 'slideLayout*.xml' -File)
        $unusedLayouts = @($allLayouts | Where-Object { -not $usedLayouts.Contains($_.Name) })

        if ($unusedLayouts.Count -eq 0) {
            Write-Host "[CLEAN] No unused layouts found" -ForegroundColor Green
            return
        }

        Write-Verbose "Found $($allLayouts.Count) total layouts, $($usedLayouts.Count) used, $($unusedLayouts.Count) unused"

        # Step 3: Delete unused layout files
        $contentTypesDoc = Get-XmlDocument $contentTypesPath
        $layoutRelsDir = Join-Path $layoutsDir '_rels'
        $removedLayoutCount = 0

        foreach ($layout in $unusedLayouts) {
            if ($PSCmdlet.ShouldProcess($layout.Name, "Remove unused layout")) {
                Remove-Item $layout.FullName -Force
                $relsFile = Join-Path $layoutRelsDir "$($layout.Name).rels"
                if (Test-Path -LiteralPath $relsFile) {
                    Remove-Item $relsFile -Force
                }

                $override = $contentTypesDoc.SelectSingleNode(
                    "//ct:Override[@PartName='/ppt/slideLayouts/$($layout.Name)']",
                    $script:NsMgr
                )
                if ($override) {
                    $null = $override.ParentNode.RemoveChild($override)
                }

                $removedLayoutCount++
                Write-Verbose "  Removed layout: $($layout.Name)"
            }
        }

        # Step 4: Clean up masters - remove references to deleted layouts,
        # then remove masters with zero remaining layouts
        $removedMasterCount = 0
        $masterFiles = @(Get-ChildItem -Path $mastersDir -Filter 'slideMaster*.xml' -File)
        $masterRelsDir = Join-Path $mastersDir '_rels'
        $mastersToRemove = [System.Collections.Generic.List[string]]::new()

        foreach ($masterFile in $masterFiles) {
            $masterDoc = Get-XmlDocument $masterFile.FullName
            $masterRelsPath = Join-Path $masterRelsDir "$($masterFile.Name).rels"
            $masterRelsDoc = if (Test-Path -LiteralPath $masterRelsPath) { Get-XmlDocument $masterRelsPath } else { $null }

            $layoutIdList = $masterDoc.SelectSingleNode('//p:sldLayoutIdLst', $script:NsMgr)
            if (-not $layoutIdList) { continue }

            $masterModified = $false
            $masterRelsModified = $false

            $layoutIds = @($layoutIdList.SelectNodes('p:sldLayoutId', $script:NsMgr))
            foreach ($layoutId in $layoutIds) {
                $rId = $layoutId.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
                if (-not $rId) { continue }

                if (-not $masterRelsDoc) { continue }
                $rel = $masterRelsDoc.SelectSingleNode("//x:Relationship[@Id='$rId']", $script:NsMgr)
                if (-not $rel) { continue }

                $targetFile = Split-Path $rel.GetAttribute('Target') -Leaf
                if (-not $usedLayouts.Contains($targetFile)) {
                    $null = $layoutId.ParentNode.RemoveChild($layoutId)
                    $null = $rel.ParentNode.RemoveChild($rel)
                    $masterModified = $true
                    $masterRelsModified = $true
                }
            }

            if ($masterModified) {
                Save-XmlDocument -doc $masterDoc -path $masterFile.FullName
            }
            if ($masterRelsModified -and $masterRelsDoc) {
                Save-XmlDocument -doc $masterRelsDoc -path $masterRelsPath
            }

            # Check if master has zero layouts remaining
            $remainingLayouts = $layoutIdList.SelectNodes('p:sldLayoutId', $script:NsMgr)
            if ($remainingLayouts.Count -eq 0) {
                $mastersToRemove.Add($masterFile.Name)
            }
        }

        # Step 5: Remove orphaned masters
        if ($mastersToRemove.Count -gt 0) {
            $presDoc = Get-XmlDocument $presentationXmlPath
            $presRelsDoc = Get-XmlDocument $presentationRelsPath
            $presModified = $false
            $presRelsModified = $false

            foreach ($masterName in $mastersToRemove) {
                if ($PSCmdlet.ShouldProcess($masterName, "Remove unused master")) {
                    # Find the master's rId in presentation.xml.rels
                    $masterRel = $presRelsDoc.SelectSingleNode(
                        "//x:Relationship[@Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster'][contains(@Target, '$masterName')]",
                        $script:NsMgr
                    )
                    if ($masterRel) {
                        $masterRId = $masterRel.GetAttribute('Id')

                        # Remove from presentation.xml sldMasterIdLst
                        $masterIdNode = $presDoc.SelectSingleNode(
                            "//p:sldMasterIdLst/p:sldMasterId[@r:id='$masterRId']",
                            $script:NsMgr
                        )
                        if ($masterIdNode) {
                            $null = $masterIdNode.ParentNode.RemoveChild($masterIdNode)
                            $presModified = $true
                        }

                        $null = $masterRel.ParentNode.RemoveChild($masterRel)
                        $presRelsModified = $true
                    }

                    # Delete files
                    $masterPath = Join-Path $mastersDir $masterName
                    if (Test-Path -LiteralPath $masterPath) {
                        Remove-Item $masterPath -Force
                    }
                    $masterRelsPath = Join-Path $masterRelsDir "$masterName.rels"
                    if (Test-Path -LiteralPath $masterRelsPath) {
                        Remove-Item $masterRelsPath -Force
                    }

                    # Remove from [Content_Types].xml
                    $override = $contentTypesDoc.SelectSingleNode(
                        "//ct:Override[@PartName='/ppt/slideMasters/$masterName']",
                        $script:NsMgr
                    )
                    if ($override) {
                        $null = $override.ParentNode.RemoveChild($override)
                    }

                    $removedMasterCount++
                    Write-Verbose "  Removed master: $masterName"
                }
            }

            if ($presModified) {
                Save-XmlDocument -doc $presDoc -path $presentationXmlPath
            }
            if ($presRelsModified) {
                Save-XmlDocument -doc $presRelsDoc -path $presentationRelsPath
            }
        }

        Save-XmlDocument -doc $contentTypesDoc -path $contentTypesPath

        Write-Host "[CLEAN] Removed $removedLayoutCount unused layouts, $removedMasterCount unused masters" -ForegroundColor Green
    }

    #endregion
