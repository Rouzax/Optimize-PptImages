    #region Deck Order and Structure

    function Get-SlideDimensions {
        <#
        .SYNOPSIS
            Gets the slide dimensions from presentation.xml.
        .DESCRIPTION
            Extracts the slide size (p:sldSz) from presentation.xml and converts to pixels.
            Used for background images that fill the entire slide.
        #>
        param([string]$tempDir)
        
        $presentationXml = Join-Path $tempDir 'ppt/presentation.xml'
        if (-not (Test-Path $presentationXml)) {
            # Return default widescreen dimensions
            return @{ WidthPx = 1280; HeightPx = 720 }
        }
        
        $presDoc = Get-XmlDocument $presentationXml
        if (-not $presDoc) {
            return @{ WidthPx = 1280; HeightPx = 720 }
        }
        
        $sldSz = $presDoc.SelectSingleNode('//p:sldSz', $script:NsMgr)
        if (-not $sldSz) {
            return @{ WidthPx = 1280; HeightPx = 720 }
        }
        
        $cx = $sldSz.GetAttribute('cx')
        $cy = $sldSz.GetAttribute('cy')
        
        $widthPx = if ($cx) { ConvertFrom-EMU -emu ([long]$cx) } else { 1280 }
        $heightPx = if ($cy) { ConvertFrom-EMU -emu ([long]$cy) } else { 720 }
        
        return @{ WidthPx = $widthPx; HeightPx = $heightPx }
    }

    function Get-CanonicalSlideOrder {
        param([string]$tempDir)
        
        $presentationXml = Join-Path $tempDir 'ppt/presentation.xml'
        if (-not (Test-Path $presentationXml)) {
            throw "presentation.xml not found"
        }
        
        $presDoc = Get-XmlDocument $presentationXml
        $presRelsPath = Join-Path $tempDir 'ppt/_rels/presentation.xml.rels'
        $presRels = Get-XmlDocument $presRelsPath
        
        # Parse slide list in document order
        $sldIdList = $presDoc.SelectNodes('//p:sldIdLst/p:sldId', $script:NsMgr)
        $slides = [System.Collections.Generic.List[SlideInfo]]::new()
        
        $slideNum = 1
        foreach ($sldId in $sldIdList) {
            $id = $sldId.GetAttribute('id')
            $relId = $sldId.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
            
            $rel = $presRels.SelectSingleNode("//x:Relationship[@Id='$relId']", $script:NsMgr)
            if ($rel) {
                $target = $rel.GetAttribute('Target')
                $partPath = "ppt/$target"
                
                $slideInfo = [SlideInfo]@{
                    Number = $slideNum
                    SlideId = $id
                    RelId = $relId
                    PartPath = $partPath
                    IsMorphTransition = $false
                }
                
                # Check for Morph transition (both old and new formats)
                $slidePath = Join-Path $tempDir $partPath
                if (Test-Path -LiteralPath $slidePath) {
                    $slideDoc = Get-XmlDocument $slidePath
                    # Check for 2015 format: p159:morph
                    $morphNew = $slideDoc.SelectSingleNode('//p:transition//p159:morph', $script:NsMgr)
                    # Check for 2010 format: p14:prst[@val="morph"]
                    $morphOld = $slideDoc.SelectSingleNode('//p:transition//p14:prst[@val="morph"]', $script:NsMgr)
                    
                    if ($morphNew -or $morphOld) {
                        $slideInfo.IsMorphTransition = $true
                    }
                }
                
                $slides.Add($slideInfo)
                $slideNum++
            }
        }
        
        if ($slides.Count -eq 0) {
            Write-Warning "No slides found in sldIdLst, falling back to file enumeration"
            # Fallback: enumerate slide files
            $slideFiles = Get-ChildItem -Path (Join-Path $tempDir 'ppt/slides') -Filter 'slide*.xml' | 
                Where-Object { $_.Name -match '^slide(\d+)\.xml$' } |
                Sort-Object { [int]$matches[1] }
            
            $slideNum = 1
            foreach ($file in $slideFiles) {
                $slideInfo = [SlideInfo]@{
                    Number = $slideNum
                    SlideId = $null
                    RelId = $null
                    PartPath = "ppt/slides/$($file.Name)"
                    IsMorphTransition = $false
                }
                $slides.Add($slideInfo)
                $slideNum++
            }
        }
        
        Write-Verbose "Processing $($slides.Count) slides (deck order from sldIdLst)"
        return ,$slides
    }

    #endregion

    #region SVG Detection

    function Test-IsSvgFallback {
        param([System.Xml.XmlElement]$blipElement)
        
        # Check for SVG extension in a:blip/a:extLst
        $svgExt = $blipElement.SelectSingleNode(
            "a:extLst/a:ext[@uri='{96DAC541-7B7A-43D3-8B79-37D633B846F1}']/asvg:svgBlip",
            $script:NsMgr
        )
        
        if ($svgExt) {
            $svgRId = $svgExt.GetAttribute('embed', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
            return @{
                IsSvgFallback = $true
                SvgRId = $svgRId
            }
        }
        
        return @{
            IsSvgFallback = $false
            SvgRId = $null
        }
    }

    #endregion

    #region Placeholder Resolution

    function Get-PlaceholderDimensionsFromLayout {
        <#
        .SYNOPSIS
            Resolves display dimensions for a placeholder-based picture from its layout.
        .DESCRIPTION
            When a picture on a slide uses a placeholder (has p:ph[@idx]), and doesn't have
            its own dimensions in p:spPr/a:xfrm, the dimensions are inherited from the
            corresponding placeholder in the slide layout.
        #>
        param(
            [string]$tempDir,
            [System.Xml.XmlDocument]$slideRels,
            [int]$placeholderIdx
        )
        
        # Find the slideLayout relationship
        $layoutRel = $slideRels.SelectSingleNode(
            "//x:Relationship[@Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout']",
            $script:NsMgr
        )
        
        if (-not $layoutRel) {
            return $null
        }
        
        # Construct path to layout file
        $layoutTarget = $layoutRel.GetAttribute('Target')
        # Target is typically "../slideLayouts/slideLayout1.xml" - resolve relative to slides folder
        $layoutPath = $layoutTarget -replace '^\.\./slideLayouts/', 'ppt/slideLayouts/'
        $layoutFullPath = Join-Path $tempDir ($layoutPath)
        
        if (-not (Test-Path -LiteralPath $layoutFullPath)) {
            Write-Verbose "    Layout file not found: $layoutFullPath"
            return $null
        }
        
        # Load layout XML
        $layoutDoc = Get-XmlDocument $layoutFullPath
        if (-not $layoutDoc) {
            return $null
        }
        
        # Find the placeholder shape with matching idx
        # Placeholders are defined as p:sp (shapes) with p:nvSpPr/p:nvPr/p:ph[@idx]
        $phShape = $layoutDoc.SelectSingleNode(
            "//p:sp[p:nvSpPr/p:nvPr/p:ph[@idx='$placeholderIdx']]",
            $script:NsMgr
        )
        
        if (-not $phShape) {
            Write-Verbose "    Placeholder idx=$placeholderIdx not found in layout"
            return $null
        }
        
        # Get dimensions from the placeholder's xfrm
        $xfrm = $phShape.SelectSingleNode('p:spPr/a:xfrm', $script:NsMgr)
        if (-not $xfrm) {
            return $null
        }
        
        $ext = $xfrm.SelectSingleNode('a:ext', $script:NsMgr)
        if (-not $ext) {
            return $null
        }
        
        $cx = $ext.GetAttribute('cx')
        $cy = $ext.GetAttribute('cy')
        
        if (-not $cx -or -not $cy) {
            return $null
        }
        
        $widthPx = ConvertFrom-EMU -emu ([long]$cx)
        $heightPx = ConvertFrom-EMU -emu ([long]$cy)
        
        Write-Verbose "    Resolved placeholder dimensions from layout: ${widthPx}x${heightPx}px (idx=$placeholderIdx)"
        
        return @{
            WidthPx = $widthPx
            HeightPx = $heightPx
            XfrmElement = $xfrm
        }
    }

    #endregion
