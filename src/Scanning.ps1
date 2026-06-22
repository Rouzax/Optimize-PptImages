    #region Image Scanning

    function Get-ImageUsages {
        param(
            [string]$tempDir,
            [System.Collections.Generic.List[SlideInfo]]$slides
        )
        
        $usages = [System.Collections.Generic.List[ImageUsage]]::new()
        $svgFallbackImages = [System.Collections.Generic.HashSet[string]]::new()
        
        # Progress tracking
        $totalItems = $slides.Count
        $currentItem = 0
        
        # Process slides
        foreach ($slide in $slides) {
            $currentItem++
            Write-Progress -Activity "Scanning for images" -Status "Slide $($slide.Number) of $($slides.Count)" `
                -PercentComplete (($currentItem / $totalItems) * 100) -Id 1
            
            # NOTE: We analyze ALL slides regardless of Include/ExcludeSlides filters
            # This ensures correct sizing (max display across all usages) and morph detection
            # Filtering is applied during processing phases (cropping, optimization)
            
            $slidePath = Join-Path $tempDir ($slide.PartPath)
            
            if (-not (Test-Path -LiteralPath $slidePath)) {
                Write-Warning "Slide file not found: $slidePath"
                continue
            }
            
            $slideDoc = Get-XmlDocument $slidePath
            if (-not $slideDoc) { 
                Write-Warning "Failed to load slide: $($slide.PartPath)"
                continue 
            }
            
            # Rels path: ppt/slides/slide1.xml -> ppt/slides/_rels/slide1.xml.rels
            $slidePartDir = Split-Path ($slide.PartPath) -Parent
            $slideFileName = Split-Path ($slide.PartPath) -Leaf
            $relsPath = Join-Path $tempDir (Join-Path $slidePartDir "_rels\$slideFileName.rels")
            $rels = if (Test-Path -LiteralPath $relsPath) { 
                Get-XmlDocument $relsPath 
            } else { 
                $null 
            }
            
            # Find all p:pic elements
            $pics = $slideDoc.GetElementsByTagName('pic', 'http://schemas.openxmlformats.org/presentationml/2006/main')
            
            foreach ($pic in $pics) {
                $usage = Get-PictureUsage -pic $pic -rels $rels -contextType ([ContextType]::Slide) `
                    -location "Slide $($slide.Number)" -slideNumber $slide.Number `
                    -tempDir $tempDir -isMorphSlide $slide.IsMorphTransition -partPath $slide.PartPath `
                    -slideDocument $slideDoc
                
                if ($usage -and $usage.IsSvgFallbackUsage) {
                    $null = $svgFallbackImages.Add($usage.ImagePhysicalPath)
                }
                
                if ($usage) {
                    $usages.Add($usage)
                }
            }
            
            # Process standalone blips (shape fills like Freeform, backgrounds, etc.)
            $blips = $slideDoc.GetElementsByTagName('blip', 'http://schemas.openxmlformats.org/drawingml/2006/main')
            
            foreach ($blip in $blips) {
                # Skip if this blip is already inside a <p:pic> (already processed above)
                $parentPic = $blip
                $isInsidePic = $false
                while ($parentPic.ParentNode) {
                    $parentPic = $parentPic.ParentNode
                    if ($parentPic.LocalName -eq 'pic' -and $parentPic.NamespaceURI -eq 'http://schemas.openxmlformats.org/presentationml/2006/main') {
                        $isInsidePic = $true
                        break
                    }
                }
                if ($isInsidePic) { continue }
                
                # Process standalone blip (shape fill, background, etc.)
                $usage = Get-BlipUsage -blip $blip -rels $rels -contextType ([ContextType]::Slide) `
                    -location "Slide $($slide.Number)" -tempDir $tempDir -partPath $slide.PartPath `
                    -slideDocument $slideDoc -slideNumber $slide.Number -isMorphSlide $slide.IsMorphTransition
                
                if ($usage -and $usage.IsSvgFallbackUsage) {
                    $null = $svgFallbackImages.Add($usage.ImagePhysicalPath)
                }
                
                if ($usage) {
                    $usages.Add($usage)
                }
            }
        }
        
        Write-Progress -Activity "Scanning for images" -Completed -Id 1
        
        # Mark all usages of SVG fallback images
        foreach ($usage in $usages) {
            if ($svgFallbackImages.Contains($usage.ImagePhysicalPath)) {
                $usage.IsSvgFallbackUsage = $true
            }
        }
        
        # Process masters and layouts
        $masterLayoutDirs = @(
            'ppt/slideMasters',
            'ppt/slideLayouts'
        )
        
        foreach ($dir in $masterLayoutDirs) {
            $fullDir = Join-Path $tempDir $dir
            if (-not (Test-Path $fullDir)) { continue }
            
            $files = @(Get-ChildItem -Path $fullDir -Filter '*.xml' -File)
            if ($files.Count -eq 0) { continue }
            
            foreach ($file in $files) {
                $partPath = ($file.FullName -replace [regex]::Escape($tempDir + [System.IO.Path]::DirectorySeparatorChar), '').Replace('\', '/')
                $contextType = if ($partPath -match 'slideMasters') { [ContextType]::Master } else { [ContextType]::Layout }
                
                $doc = Get-XmlDocument $file.FullName
                if (-not $doc) { continue }
                
                $partDir = Split-Path $partPath -Parent
                $partFile = Split-Path $partPath -Leaf
                $relsPath = Join-Path $tempDir (($partDir + '/_rels/' + $partFile + '.rels'))
                $rels = if (Test-Path -LiteralPath $relsPath) { Get-XmlDocument $relsPath } else { $null }
                
                $pics = $doc.GetElementsByTagName('pic', 'http://schemas.openxmlformats.org/presentationml/2006/main')
                $blips = $doc.GetElementsByTagName('blip', 'http://schemas.openxmlformats.org/drawingml/2006/main')
                
                # Process pictures (standard image containers)
                foreach ($pic in $pics) {
                    $usage = Get-PictureUsage -pic $pic -rels $rels -contextType $contextType `
                        -location "$contextType ($partFile)" -slideNumber 0 `
                        -tempDir $tempDir -isMorphSlide $false -partPath $partPath `
                        -slideDocument $doc
                    
                    if ($usage -and $usage.IsSvgFallbackUsage) {
                        $null = $svgFallbackImages.Add($usage.ImagePhysicalPath)
                    }
                    
                    if ($usage) {
                        $usages.Add($usage)
                    }
                }
                
                # Process standalone blips (backgrounds, shape fills, etc.)
                foreach ($blip in $blips) {
                    # Skip if this blip is already inside a <p:pic> (already processed above)
                    $parentPic = $blip
                    while ($parentPic.ParentNode) {
                        $parentPic = $parentPic.ParentNode
                        if ($parentPic.LocalName -eq 'pic' -and $parentPic.NamespaceURI -eq 'http://schemas.openxmlformats.org/presentationml/2006/main') {
                            $parentPic = $null
                            break
                        }
                    }
                    if (-not $parentPic) { continue }
                    
                    # Process standalone blip
                    $usage = Get-BlipUsage -blip $blip -rels $rels -contextType $contextType `
                        -location "$contextType ($partFile)" -tempDir $tempDir -partPath $partPath `
                        -slideDocument $doc
                    
                    if ($usage -and $usage.IsSvgFallbackUsage) {
                        $null = $svgFallbackImages.Add($usage.ImagePhysicalPath)
                    }
                    
                    if ($usage) {
                        $usages.Add($usage)
                    }
                }
            }
        }
        
        $slideUsages = @($usages | Where-Object { $_.ContextType -eq [ContextType]::Slide })
        $layoutUsages = @($usages | Where-Object { $_.ContextType -eq [ContextType]::Layout })
        $masterUsages = @($usages | Where-Object { $_.ContextType -eq [ContextType]::Master })
        Write-Verbose "Image usage breakdown: $($slideUsages.Count) in slides, $($layoutUsages.Count) in layouts, $($masterUsages.Count) in masters"
        
        return ,$usages
    }

    function Get-PictureUsage {
        param(
            [System.Xml.XmlElement]$pic,
            [System.Xml.XmlDocument]$rels,
            [ContextType]$contextType,
            [string]$location,
            [int]$slideNumber,
            [string]$tempDir,
            [bool]$isMorphSlide,
            [string]$partPath,
            [System.Xml.XmlDocument]$slideDocument
        )
        
        # Find blip element using GetElementsByTagName
        $blips = $pic.GetElementsByTagName('blip', 'http://schemas.openxmlformats.org/drawingml/2006/main')
        if ($blips.Count -eq 0) { 
            Write-Verbose "  No a:blip element found in picture"
            return $null 
        }
        
        $blip = $blips[0]
        
        # Check for r:embed attribute
        $rId = $blip.GetAttribute('embed', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
        if (-not $rId) {
            return $null
        }
        
        if (-not $rels) { 
            return $null 
        }
        
        $rel = $rels.SelectSingleNode("//x:Relationship[@Id='$rId']", $script:NsMgr)
        if (-not $rel) { 
            return $null 
        }
        
        $target = $rel.GetAttribute('Target')
        $mediaPath = $target -replace '^\.\./media/', 'ppt/media/'
        $fullPath = Join-Path $tempDir ($mediaPath)
        
        if (-not (Test-Path -LiteralPath $fullPath)) { 
            return $null 
        }
        
        # Get shape name
        $cNvPr = $pic.SelectSingleNode('.//p:cNvPr', $script:NsMgr)
        $shapeName = if ($cNvPr) { $cNvPr.GetAttribute('name') } else { 'Unknown' }
        
        # Check SVG fallback
        $svgCheck = Test-IsSvgFallback -blipElement $blip
        
        # Get display dimensions from xfrm
        $xfrm = $pic.SelectSingleNode('.//a:xfrm', $script:NsMgr)
        $ext = if ($xfrm) { $xfrm.SelectSingleNode('a:ext', $script:NsMgr) } else { $null }
        
        $displayWidth = 0
        $displayHeight = 0
        if ($ext) {
            $cx = $ext.GetAttribute('cx')
            $cy = $ext.GetAttribute('cy')
            if ($cx) { $displayWidth = ConvertFrom-EMU -emu ([long]$cx) }
            if ($cy) { $displayHeight = ConvertFrom-EMU -emu ([long]$cy) }
        }
        
        # If dimensions are 0, check if picture uses a placeholder and resolve from layout
        if ($displayWidth -eq 0 -and $displayHeight -eq 0 -and $contextType -eq [ContextType]::Slide) {
            # Check for placeholder reference in p:nvPicPr/p:nvPr/p:ph[@idx]
            $phElement = $pic.SelectSingleNode('p:nvPicPr/p:nvPr/p:ph[@idx]', $script:NsMgr)
            if ($phElement) {
                $phIdx = $phElement.GetAttribute('idx')
                if ($phIdx -and $rels) {
                    $layoutDims = Get-PlaceholderDimensionsFromLayout -tempDir $tempDir -slideRels $rels -placeholderIdx ([int]$phIdx)
                    if ($layoutDims) {
                        $displayWidth = $layoutDims.WidthPx
                        $displayHeight = $layoutDims.HeightPx
                        # Keep the original xfrm (from pic) for later updates - we just need the display dimensions for sizing
                    }
                }
            }
        }
        
        # Check for srcRect (sibling of blip within blipFill)
        $blipFill = $blip.ParentNode
        $srcRect = if ($blipFill) { 
            $blipFill.SelectSingleNode('a:srcRect', $script:NsMgr) 
        } else { 
            $null 
        }
        $hasSrcRect = $null -ne $srcRect
        $l = 0.0; $t = 0.0; $r = 0.0; $b = 0.0
        
        if ($srcRect) {
            $l = if ($srcRect.HasAttribute('l')) { [double]$srcRect.GetAttribute('l') } else { 0.0 }
            $t = if ($srcRect.HasAttribute('t')) { [double]$srcRect.GetAttribute('t') } else { 0.0 }
            $r = if ($srcRect.HasAttribute('r')) { [double]$srcRect.GetAttribute('r') } else { 0.0 }
            $b = if ($srcRect.HasAttribute('b')) { [double]$srcRect.GetAttribute('b') } else { 0.0 }
        }
        
        $usage = [ImageUsage]@{
            Location = $location
            ContextType = $contextType
            SlideNumber = $slideNumber
            ShapeName = $shapeName
            ImagePhysicalPath = $fullPath
            OriginalFileName = Split-Path $fullPath -Leaf
            SourceWidthPx = 0
            SourceHeightPx = 0
            DisplayWidthPx = $displayWidth
            DisplayHeightPx = $displayHeight
            HasSrcRect = $hasSrcRect
            SrcRectLeft = $l
            SrcRectTop = $t
            SrcRectRight = $r
            SrcRectBottom = $b
            IsMorphSlide = $isMorphSlide
            MorphPair = $null
            IsSvgFallbackUsage = $svgCheck.IsSvgFallback
            BlipRId = $rId
            PartPath = $partPath  # Keep original forward-slash format for XML operations
            BlipElement = $blip
            XfrmElement = $xfrm
            BeforeSizeBytes = (Get-Item $fullPath).Length
            AfterSizeBytes = 0
            CropApplied = $false
            CropNormalized = $false
            CropRemovedNoOp = $false
            OptimizationStatus = 'Pending'
            WhyNotOptimized = ''
            EffectiveTransparencyPercent = 0.0
            TargetWidthPx = 0
            TargetHeightPx = 0
            ManualActionRequired = $false
            ManualActionHint = ''
            OptimizedFile = ''
            SlideDocument = $slideDocument
            SourceJpegQuality = -1
        }
        
        # Build informative verbose description
        if (Test-VerboseMode) {
            $cropInfo = if ($hasSrcRect) {
                $crop = Get-CropPercentage -left $l -top $t -right $r -bottom $b
                "crop [$($crop.WidthPercent.ToString('F1'))%x$($crop.HeightPercent.ToString('F1'))%] (l:$l t:$t r:$r b:$b)"
            } else {
                "no crop"
            }
            
            $flags = @()
            if ($isMorphSlide) { $flags += 'Morph' }
            if ($svgCheck.IsSvgFallback) { $flags += 'SVG-fallback' }
            $flagStr = if ($flags.Count -gt 0) { " [$($flags -join ', ')]" } else { "" }
            
            Write-Verbose "  $location -> '$shapeName': $($usage.OriginalFileName) [${displayWidth}x${displayHeight}px, $cropInfo]$flagStr"
        }
        
        return $usage
    }

    function Get-BlipUsage {
        param(
            [System.Xml.XmlElement]$blip,
            [System.Xml.XmlDocument]$rels,
            [ContextType]$contextType,
            [string]$location,
            [string]$tempDir,
            [string]$partPath,
            [System.Xml.XmlDocument]$slideDocument,
            [int]$slideNumber = 0,
            [bool]$isMorphSlide = $false
        )
        
        # Check for r:embed attribute
        $rId = $blip.GetAttribute('embed', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
        if (-not $rId) {
            return $null
        }
        
        # Resolve relationship
        if (-not $rels) {
            Write-Verbose "  No relationships document available"
            return $null
        }
        
        $rel = $rels.SelectSingleNode("//x:Relationship[@Id='$rId']", $script:NsMgr)
        if (-not $rel) {
            return $null
        }
        
        $target = $rel.GetAttribute('Target')
        $mediaFile = $target -replace '^\.\./media/', ''
        $mediaPath = Join-Path $tempDir "ppt/media/$mediaFile"
        
        if (-not (Test-Path -LiteralPath $mediaPath)) {
            Write-Verbose "  Media file not found: $mediaFile"
            return $null
        }
        
        # Check SVG fallback
        $svgCheck = Test-IsSvgFallback -blipElement $blip
        
        # Walk up the DOM to find parent shape and determine context
        $shapeName = $null
        $displayWidth = 0
        $displayHeight = 0
        $isBackground = $false
        $xfrmElement = $null
        
        $parent = $blip.ParentNode
        while ($parent -and $parent.NodeType -eq [System.Xml.XmlNodeType]::Element) {
            $localName = $parent.LocalName
            
            # Check for background
            if ($localName -eq 'bgPr' -or $localName -eq 'bg') {
                $isBackground = $true
                if ($script:SlideDimensions) {
                    $displayWidth = $script:SlideDimensions.WidthPx
                    $displayHeight = $script:SlideDimensions.HeightPx
                }
                $shapeName = 'Slide Background'
                break
            }
            
            # Check for shape (sp) - this includes Freeform shapes
            if ($localName -eq 'sp' -and $parent.NamespaceURI -eq 'http://schemas.openxmlformats.org/presentationml/2006/main') {
                # Get shape name from p:nvSpPr/p:cNvPr
                $cNvPr = $parent.SelectSingleNode('p:nvSpPr/p:cNvPr', $script:NsMgr)
                if ($cNvPr) {
                    $shapeName = $cNvPr.GetAttribute('name')
                }
                
                # Get dimensions from p:spPr/a:xfrm/a:ext
                $spPr = $parent.SelectSingleNode('p:spPr', $script:NsMgr)
                if ($spPr) {
                    $xfrm = $spPr.SelectSingleNode('a:xfrm', $script:NsMgr)
                    if ($xfrm) {
                        $xfrmElement = $xfrm
                        $ext = $xfrm.SelectSingleNode('a:ext', $script:NsMgr)
                        if ($ext) {
                            $cx = $ext.GetAttribute('cx')
                            $cy = $ext.GetAttribute('cy')
                            if ($cx) { $displayWidth = ConvertFrom-EMU -emu ([long]$cx) }
                            if ($cy) { $displayHeight = ConvertFrom-EMU -emu ([long]$cy) }
                        }
                    }
                }
                break
            }
            
            # Check for other shape types (cxnSp/grpSp don't contain images we can optimize)
            if ($localName -eq 'cxnSp' -or $localName -eq 'grpSp') {
                break
            }
            
            $parent = $parent.ParentNode
        }
        
        # Fallback shape name if not found
        if (-not $shapeName) {
            $shapeName = Get-BlipParentShapeName -blip $blip
        }
        
        # Check for crop values - either srcRect or fillRect (with negative values)
        # srcRect: positive values = percentage to crop from each edge
        # fillRect: negative values = percentage image extends beyond shape (converts to srcRect equivalent)
        $blipFill = $blip.ParentNode
        $hasSrcRect = $false
        $l = 0.0; $t = 0.0; $r = 0.0; $b = 0.0
        
        if ($blipFill -and $blipFill.LocalName -eq 'blipFill') {
            # First check for srcRect (direct crop specification)
            $srcRect = $blipFill.SelectSingleNode('a:srcRect', $script:NsMgr)
            if ($srcRect) {
                $hasSrcRect = $true
                $lAttr = $srcRect.GetAttribute('l')
                $tAttr = $srcRect.GetAttribute('t')
                $rAttr = $srcRect.GetAttribute('r')
                $bAttr = $srcRect.GetAttribute('b')
                if ($lAttr) { $l = [double]$lAttr }
                if ($tAttr) { $t = [double]$tAttr }
                if ($rAttr) { $r = [double]$rAttr }
                if ($bAttr) { $b = [double]$bAttr }
            }
            
            # Check for fillRect with negative values (stretch crop)
            # fillRect is inside a:stretch element
            if (-not $hasSrcRect) {
                $stretch = $blipFill.SelectSingleNode('a:stretch', $script:NsMgr)
                if ($stretch) {
                    $fillRect = $stretch.SelectSingleNode('a:fillRect', $script:NsMgr)
                    if ($fillRect) {
                        $fl = 0.0; $ft = 0.0; $fr = 0.0; $fb = 0.0
                        $flAttr = $fillRect.GetAttribute('l')
                        $ftAttr = $fillRect.GetAttribute('t')
                        $frAttr = $fillRect.GetAttribute('r')
                        $fbAttr = $fillRect.GetAttribute('b')
                        if ($flAttr) { $fl = [double]$flAttr }
                        if ($ftAttr) { $ft = [double]$ftAttr }
                        if ($frAttr) { $fr = [double]$frAttr }
                        if ($fbAttr) { $fb = [double]$fbAttr }
                        
                        # Negative fillRect values indicate a crop (image extends beyond shape)
                        # Convert to srcRect-equivalent values
                        # Formula: srcRect% = |fillRect%| / (100000 + |l| + |r|) * 100000 for horizontal
                        #          srcRect% = |fillRect%| / (100000 + |t| + |b|) * 100000 for vertical
                        # Values are in 1/1000ths of percent (100000 = 100%)
                        
                        if ($fl -lt 0 -or $fr -lt 0 -or $ft -lt 0 -or $fb -lt 0) {
                            $hasSrcRect = $true
                            
                            # Convert negative fillRect to positive srcRect equivalent
                            $absL = [Math]::Abs($fl)
                            $absR = [Math]::Abs($fr)
                            $absT = [Math]::Abs($ft)
                            $absB = [Math]::Abs($fb)
                            
                            # Horizontal: total scale = 100% + left extension + right extension
                            $horzScale = 100000 + $absL + $absR
                            if ($horzScale -gt 100000) {
                                $l = ($absL / $horzScale) * 100000
                                $r = ($absR / $horzScale) * 100000
                            }
                            
                            # Vertical: total scale = 100% + top extension + bottom extension
                            $vertScale = 100000 + $absT + $absB
                            if ($vertScale -gt 100000) {
                                $t = ($absT / $vertScale) * 100000
                                $b = ($absB / $vertScale) * 100000
                            }
                            
                            Write-Verbose "    Converted fillRect (l=$fl, t=$ft, r=$fr, b=$fb) to srcRect equivalent (l=$([int]$l), t=$([int]$t), r=$([int]$r), b=$([int]$b))"
                        }
                    }
                }
            }
        }
        
        # Create usage object
        $usage = [ImageUsage]@{
            Location = $location
            ContextType = $contextType
            SlideNumber = $slideNumber
            ShapeName = $shapeName
            ImagePhysicalPath = $mediaPath
            OriginalFileName = $mediaFile
            SourceWidthPx = 0
            SourceHeightPx = 0
            DisplayWidthPx = $displayWidth
            DisplayHeightPx = $displayHeight
            HasSrcRect = $hasSrcRect
            SrcRectLeft = $l
            SrcRectTop = $t
            SrcRectRight = $r
            SrcRectBottom = $b
            IsMorphSlide = $isMorphSlide
            MorphPair = $null
            IsSvgFallbackUsage = $svgCheck.IsSvgFallback
            BlipRId = $rId
            PartPath = $partPath
            BlipElement = $blip
            XfrmElement = $xfrmElement
            BeforeSizeBytes = (Get-Item $mediaPath).Length
            AfterSizeBytes = 0
            CropApplied = $false
            CropNormalized = $false
            CropRemovedNoOp = $false
            OptimizationStatus = 'Pending'
            WhyNotOptimized = ''
            EffectiveTransparencyPercent = 0
            TargetWidthPx = 0
            TargetHeightPx = 0
            ManualActionRequired = $false
            ManualActionHint = ''
            OptimizedFile = ''
            SlideDocument = $slideDocument
            SourceJpegQuality = -1
        }
        
        if (Test-VerboseMode) {
            $svgFlag = if ($svgCheck.IsSvgFallback) { " [SVG-fallback]" } else { "" }
            $cropInfo = if ($hasSrcRect) { "crop" } else { "no crop" }
            $dimInfo = if ($displayWidth -gt 0 -and $displayHeight -gt 0) { 
                "[${displayWidth}x${displayHeight}px, $cropInfo]" 
            } elseif ($isBackground) { 
                "[${displayWidth}x${displayHeight}px, background]" 
            } else { 
                "[shape fill]" 
            }
            Write-Verbose "  $location -> '$shapeName': $mediaFile $dimInfo$svgFlag"
        }
        
        return $usage
    }

    function Get-BlipParentShapeName {
        param([System.Xml.XmlElement]$blip)
        
        # Walk up the DOM to identify context
        $current = $blip.ParentNode
        $depth = 0
        $maxDepth = 10
        
        while ($current -and $depth -lt $maxDepth) {
            $localName = $current.LocalName
            
            switch ($localName) {
                'bg' { return 'Slide Background' }
                'bgPr' { return 'Background Properties' }
                'spPr' {
                    # Check if parent is a specific shape type
                    $parentShape = $current.ParentNode
                    if ($parentShape) {
                        switch ($parentShape.LocalName) {
                            'sp' { return 'Shape Fill' }
                            'cxnSp' { return 'Connector Fill' }
                            'grpSp' { return 'Group Shape Fill' }
                            default { return "Shape Fill ($($parentShape.LocalName))" }
                        }
                    }
                    return 'Shape Properties'
                }
                'ln' { return 'Line/Border Fill' }
                'txPr' { return 'Text Properties Fill' }
                'tcPr' { return 'Table Cell Fill' }
                'tblPr' { return 'Table Fill' }
            }
            
            $current = $current.ParentNode
            $depth++
        }
        
        return 'Background/Fill'
    }

    # Shared scan core: given an already-extracted package directory, build the
    # canonical slide order, capture slide dimensions, and scan image usages.
    # Caller owns the temp directory lifecycle and must have run Initialize-Namespaces.
    function Get-PptImageScanData {
        param([string]$tempDir)

        $slides = Get-CanonicalSlideOrder -tempDir $tempDir
        if (-not $slides -or $slides.Count -eq 0) {
            Write-Warning "No slides found in presentation"
            $slides = [System.Collections.Generic.List[SlideInfo]]::new()
        }

        $script:SlideDimensions = Get-SlideDimensions -tempDir $tempDir
        Write-Verbose "Slide dimensions: $($script:SlideDimensions.WidthPx)x$($script:SlideDimensions.HeightPx)px"

        $usages = Get-ImageUsages -tempDir $tempDir -slides $slides
        if (-not $usages) {
            $usages = [System.Collections.Generic.List[ImageUsage]]::new()
        }

        return [PSCustomObject]@{
            Usages = $usages
            Slides = $slides
        }
    }

    #endregion
