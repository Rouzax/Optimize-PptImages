<#
.SYNOPSIS
    Optimizes PowerPoint presentations by cropping images and reducing file sizes.

.DESCRIPTION
    Creates an optimized copy of a PowerPoint file (.pptx/.potx) plus a CSV report
    that explains every decision, while preserving visual fidelity (rectangular crops,
    masters/layouts, Morph transitions).

.PARAMETER InputPath
    Path to the input PowerPoint file (.pptx or .potx).

.PARAMETER OutputPath
    Path for the optimized output file. Default: <name>.optimized<ext>

.PARAMETER MagickPath
    Path to ImageMagick 'magick' executable. If not specified, searches PATH.

.PARAMETER EnableCropping
    Allow materializing rectangular a:srcRect crops into physical images.

.PARAMETER CropMastersAndLayouts
    Requires -EnableCropping. Allows cropping images in masters/layouts.

.PARAMETER OptimizeMastersAndLayouts
    Allow optimizing images in masters/layouts.

.PARAMETER HeadroomFactor
    Target resolution multiplier (display × factor). Default: 2.0. Range: 1.0-4.0

.PARAMETER JpegQuality
    JPEG quality for optimization. Default: 95. Range: 1-100

.PARAMETER TransparencyThresholdPercent
    Threshold for treating images as transparent. Default: 0.5. Range: 0.0-100.0

.PARAMETER CsvReportPath
    Path for the CSV report. Default: alongside OutputPath as <OutputBaseName>.opt-report.csv

.EXAMPLE
    .\Optimize-PptImages.ps1 -InputPath "presentation.pptx" -EnableCropping

.EXAMPLE
    .\Optimize-PptImages.ps1 -InputPath "template.potx" -EnableCropping -CropMastersAndLayouts -OptimizeMastersAndLayouts -HeadroomFactor 1.5

#>

[CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='Medium')]
param(
    [Parameter(Mandatory=$true, Position=0)]
    [ValidateScript({
        if (-not (Test-Path $_)) {
            throw "Input file not found: $_"
        }
        if ($_ -notmatch '\.(pptx|potx)$') {
            throw "Input file must be .pptx or .potx"
        }
        $true
    })]
    [string]$InputPath,

    [Parameter(Mandatory=$false)]
    [string]$OutputPath,

    [Parameter(Mandatory=$false)]
    [string]$MagickPath,

    [Parameter(Mandatory=$false)]
    [switch]$EnableCropping,

    [Parameter(Mandatory=$false)]
    [switch]$CropMastersAndLayouts,

    [Parameter(Mandatory=$false)]
    [switch]$OptimizeMastersAndLayouts,

    [Parameter(Mandatory=$false)]
    [ValidateRange(1.0, 4.0)]
    [double]$HeadroomFactor = 2.0,

    [Parameter(Mandatory=$false)]
    [ValidateRange(1, 100)]
    [int]$JpegQuality = 95,

    [Parameter(Mandatory=$false)]
    [ValidateRange(0.0, 100.0)]
    [double]$TransparencyThresholdPercent = 0.1,

    [Parameter(Mandatory=$false)]
    [string]$CsvReportPath
)

#Requires -Version 5.1

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region Configuration Constants

$script:Config = @{
    EMU_PER_INCH = 914400
    DPI = 96
    NO_OP_CROP_THRESHOLD = 98.0  # percentage
    MIN_CROPPED_DIMENSION_PERCENT = 2.0  # percentage
    MIN_EMU_DIMENSION = 1
    CROP_THOUSANDTHS_MAX = 100000
    JPEG_SAMPLING_FACTOR_DEFAULT = '4:2:0'
    JPEG_SAMPLING_FACTOR_UI = '4:4:4'
    PNG_COMPRESSION_LEVEL = 9
    RESIZE_FILTER = 'Lanczos'
    IMAGEMAGICK_MEMORY_LIMIT = '2GB'
    IMAGEMAGICK_MAP_LIMIT = '4GB'
    PARALLEL_THROTTLE_LIMIT = 4
}

#endregion

#region Helper Classes and Structures

class ImageUsage {
    [string]$Location
    [string]$ContextType  # Slide|Layout|Master|Notes|Handout
    [int]$SlideNumber
    [string]$ShapeName
    [string]$ImagePhysicalPath
    [string]$OriginalFileName
    [int]$SourceWidthPx
    [int]$SourceHeightPx
    [int]$DisplayWidthPx
    [int]$DisplayHeightPx
    [bool]$HasSrcRect
    [double]$SrcRectLeft
    [double]$SrcRectTop
    [double]$SrcRectRight
    [double]$SrcRectBottom
    [bool]$IsMorphSlide
    [object]$MorphPair  # Link to paired usage
    [bool]$IsSvgFallbackUsage
    [string]$SvgPartPath
    [string]$BlipRId
    [string]$PartPath
    [object]$BlipElement
    [object]$XfrmElement
    [long]$BeforeSizeBytes
    [long]$AfterSizeBytes
    [bool]$CropApplied
    [bool]$CropNormalized
    [bool]$CropRemovedNoOp
    [string]$OptimizationStatus
    [string]$WhyNotOptimized
    [double]$EffectiveTransparencyPercent
    [int]$TargetWidthPx
    [int]$TargetHeightPx
    [bool]$ManualActionRequired
    [string]$ManualActionHint
    [string]$OptimizedFile
    [System.Xml.XmlDocument]$SlideDocument  # Reference to parent document for saving
}

class ImageGroup {
    [string]$PhysicalPath
    [System.Collections.Generic.List[ImageUsage]]$Usages
    [bool]$IsSvgFallbackImage
    [long]$OriginalSizeBytes
    [long]$OptimizedSizeBytes
    [string]$OptimizedPath

    ImageGroup() {
        $this.Usages = [System.Collections.Generic.List[ImageUsage]]::new()
    }
}

class SlideInfo {
    [int]$Number
    [string]$SlideId
    [string]$RelId
    [string]$PartPath
    [bool]$IsMorphTransition
}

#endregion

#region Validation and Initialization

function Initialize-Script {
    param(
        [string]$InputPath,
        [string]$MagickPath
    )

    # Validate CropMastersAndLayouts requires EnableCropping
    if ($CropMastersAndLayouts -and -not $EnableCropping) {
        throw "-CropMastersAndLayouts requires -EnableCropping"
    }

    # Resolve paths
    $script:InputFile = Resolve-Path $InputPath
    
    if (-not $OutputPath) {
        $dir = Split-Path $script:InputFile -Parent
        $base = [System.IO.Path]::GetFileNameWithoutExtension($script:InputFile)
        $ext = [System.IO.Path]::GetExtension($script:InputFile)
        $script:OutputFile = Join-Path $dir "$base.optimized$ext"
    } else {
        $script:OutputFile = $OutputPath
    }

    if (-not $CsvReportPath) {
        $dir = Split-Path $script:OutputFile -Parent
        $base = [System.IO.Path]::GetFileNameWithoutExtension($script:OutputFile)
        $script:CsvReport = Join-Path $dir "$base.opt-report.csv"
    } else {
        $script:CsvReport = $CsvReportPath
    }
    
    if ($PSCmdlet.MyInvocation.BoundParameters['Verbose']) {
        Write-Verbose "Configuration:"
        Write-Verbose "  Input:  $script:InputFile"
        Write-Verbose "  Output: $script:OutputFile"
        Write-Verbose "  Report: $script:CsvReport"
        Write-Verbose "Parameters:"
        Write-Verbose "  EnableCropping: $EnableCropping"
        if ($EnableCropping) {
            Write-Verbose "  CropMastersAndLayouts: $CropMastersAndLayouts"
        }
        Write-Verbose "  OptimizeMastersAndLayouts: $OptimizeMastersAndLayouts"
        Write-Verbose "  HeadroomFactor: $HeadroomFactor"
        Write-Verbose "  JpegQuality: $JpegQuality"
        Write-Verbose "  TransparencyThreshold: $TransparencyThresholdPercent%"
    }
    
    # Check if output file is accessible (fail early)
    if (Test-Path -LiteralPath $script:OutputFile) {
        try {
            $testStream = [System.IO.File]::Open($script:OutputFile, 'Open', 'Write', 'None')
            $testStream.Close()
            $testStream.Dispose()
        } catch {
            throw "Output file is in use by another process: $script:OutputFile. Please close it and try again."
        }
    }

    # Find ImageMagick
    $script:MagickExe = Find-ImageMagick -ProvidedPath $MagickPath
    if (-not $script:MagickExe) {
        throw "ImageMagick 'magick' command not found. Install ImageMagick or specify -MagickPath"
    }

    Write-Verbose "ImageMagick: $script:MagickExe"
}

function Find-ImageMagick {
    param([string]$ProvidedPath)

    if ($ProvidedPath) {
        if (Test-Path -LiteralPath $ProvidedPath) {
            return $ProvidedPath
        }
        return $null
    }

    # Search PATH
    $magick = Get-Command 'magick' -ErrorAction SilentlyContinue
    if ($magick) {
        return $magick.Source
    }

    # Common locations
    $commonPaths = @(
        'C:\Program Files\ImageMagick-*\magick.exe',
        'C:\Program Files (x86)\ImageMagick-*\magick.exe',
        '/usr/bin/magick',
        '/usr/local/bin/magick',
        '/opt/homebrew/bin/magick'
    )

    foreach ($pattern in $commonPaths) {
        $found = Get-Item $pattern -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($found) {
            return $found.FullName
        }
    }

    return $null
}

#endregion

#region Utility Functions

function ConvertFrom-EMU {
    param([long]$emu)
    return [Math]::Round($emu / $script:Config.EMU_PER_INCH * $script:Config.DPI)
}

function Get-CropPercentage {
    param(
        [double]$left,
        [double]$top,
        [double]$right,
        [double]$bottom
    )
    
    $widthRemaining = (100000 - $left - $right) / 1000.0
    $heightRemaining = (100000 - $top - $bottom) / 1000.0
    
    return @{
        WidthPercent = $widthRemaining
        HeightPercent = $heightRemaining
    }
}

function Test-IsNoOpCrop {
    param(
        [double]$left,
        [double]$top,
        [double]$right,
        [double]$bottom
    )
    
    $crop = Get-CropPercentage -left $left -top $top -right $right -bottom $bottom
    return ($crop.WidthPercent -ge $script:Config.NO_OP_CROP_THRESHOLD -and 
            $crop.HeightPercent -ge $script:Config.NO_OP_CROP_THRESHOLD)
}

function Format-ByteSize {
    param([long]$bytes)
    
    if ($bytes -ge 1MB) {
        return "{0:N2} MB" -f ($bytes / 1MB)
    } elseif ($bytes -ge 1KB) {
        return "{0:N2} KB" -f ($bytes / 1KB)
    } else {
        return "$bytes B"
    }
}

function Format-Savings {
    param(
        [long]$before,
        [long]$after
    )
    
    if ($before -eq 0) { return "N/A" }
    
    $saved = $before - $after
    $percent = ($saved / $before) * 100
    
    return "{0} ({1:N1}%)" -f (Format-ByteSize $saved), $percent
}

function Get-UniqueRId {
    param(
        [System.Xml.XmlDocument]$relsDoc,
        [string]$prefix = "rId"
    )
    
    # Force array collection to prevent pipeline leakage
    $existingIds = @($relsDoc.SelectNodes("//x:Relationship/@Id", $script:NsMgr) | ForEach-Object { $_.Value })
    $maxNum = 0
    
    foreach ($id in $existingIds) {
        if ($id -match '^rId(\d+)$') {
            $num = [int]$matches[1]
            if ($num -gt $maxNum) {
                $maxNum = $num
            }
        }
    }
    
    # Explicitly write output to avoid leakage
    Write-Output "rId$($maxNum + 1)"
}

function Get-UniqueMediaFileName {
    param(
        [string]$baseName,
        [string]$extension,
        [System.Collections.Generic.HashSet[string]]$existingNames
    )
    
    $counter = 1
    $candidate = "$baseName$extension"
    
    while ($existingNames.Contains($candidate)) {
        $candidate = "$baseName$counter$extension"
        $counter++
    }
    
    $null = $existingNames.Add($candidate)
    return $candidate
}

#endregion

#region XML and Namespace Helpers

function Initialize-Namespaces {
    $script:NsMgr = New-Object System.Xml.XmlNamespaceManager([System.Xml.XmlDocument]::new().NameTable)
    $script:NsMgr.AddNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main')
    $script:NsMgr.AddNamespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
    $script:NsMgr.AddNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main')
    $script:NsMgr.AddNamespace('x', 'http://schemas.openxmlformats.org/package/2006/relationships')
    $script:NsMgr.AddNamespace('asvg', 'http://schemas.microsoft.com/office/drawing/2016/SVG/main')
    $script:NsMgr.AddNamespace('ct', 'http://schemas.openxmlformats.org/package/2006/content-types')
    $script:NsMgr.AddNamespace('p14', 'http://schemas.microsoft.com/office/powerpoint/2010/main')
    $script:NsMgr.AddNamespace('p159', 'http://schemas.microsoft.com/office/powerpoint/2015/09/main')
}

function Get-XmlDocument {
    param([string]$path)
    
    if (-not (Test-Path -LiteralPath $path)) {
        Write-Warning "XML file not found: $path"
        return $null
    }
    
    try {
        $xml = New-Object System.Xml.XmlDocument
        $xml.XmlResolver = $null
        $xml.PreserveWhitespace = $false
        $xml.Load($path)
        return $xml
    } catch {
        Write-Warning "Failed to load XML from $path : $_"
        return $null
    }
}

function Save-XmlDocument {
    param(
        [System.Xml.XmlDocument]$doc,
        [string]$path
    )
    
    $settings = New-Object System.Xml.XmlWriterSettings
    $settings.Encoding = [System.Text.UTF8Encoding]::new($false)  # No BOM
    $settings.Indent = $false
    $settings.NewLineHandling = [System.Xml.NewLineHandling]::None
    
    $writer = [System.Xml.XmlWriter]::Create($path, $settings)
    try {
        $doc.Save($writer)
        $writer.Flush()
    } finally {
        if ($writer) {
            $writer.Dispose()  # Dispose instead of Close for proper cleanup
        }
    }
}

#endregion

#region Deck Order and Structure

function Get-CanonicalSlideOrder {
    param([string]$tempDir)
    
    $presentationXml = Join-Path $tempDir 'ppt\presentation.xml'
    if (-not (Test-Path $presentationXml)) {
        throw "presentation.xml not found"
    }
    
    $presDoc = Get-XmlDocument $presentationXml
    $presRelsPath = Join-Path $tempDir 'ppt\_rels\presentation.xml.rels'
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
        $slideFiles = Get-ChildItem -Path (Join-Path $tempDir 'ppt\slides') -Filter 'slide*.xml' | 
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

#region Image Scanning

function Get-ImageUsages {
    param(
        [string]$tempDir,
        [System.Collections.Generic.List[SlideInfo]]$slides
    )
    
    $usages = [System.Collections.Generic.List[ImageUsage]]::new()
    $svgFallbackImages = [System.Collections.Generic.HashSet[string]]::new()
    
    # Process slides
    foreach ($slide in $slides) {
        $slidePath = Join-Path $tempDir ($slide.PartPath -replace '/', '\')
        
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
        $slidePartDir = Split-Path ($slide.PartPath -replace '/', '\') -Parent
        $slideFileName = Split-Path ($slide.PartPath -replace '/', '\') -Leaf
        $relsPath = Join-Path $tempDir (Join-Path $slidePartDir "_rels\$slideFileName.rels")
        $rels = if (Test-Path -LiteralPath $relsPath) { 
            Get-XmlDocument $relsPath 
        } else { 
            $null 
        }
        
        # Find all p:pic elements
        $pics = $slideDoc.GetElementsByTagName('pic', 'http://schemas.openxmlformats.org/presentationml/2006/main')
        
        foreach ($pic in $pics) {
            $usage = Get-PictureUsage -pic $pic -rels $rels -contextType 'Slide' `
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
    }
    
    # Mark all usages of SVG fallback images
    foreach ($usage in $usages) {
        if ($svgFallbackImages.Contains($usage.ImagePhysicalPath)) {
            $usage.IsSvgFallbackUsage = $true
        }
    }
    
    # Process masters and layouts
    $masterLayoutDirs = @(
        'ppt\slideMasters',
        'ppt\slideLayouts'
    )
    
    foreach ($dir in $masterLayoutDirs) {
        $fullDir = Join-Path $tempDir $dir
        if (-not (Test-Path $fullDir)) { continue }
        
        $files = @(Get-ChildItem -Path $fullDir -Filter '*.xml' -File)
        if ($files.Count -eq 0) { continue }
        
        foreach ($file in $files) {
            $partPath = ($file.FullName -replace [regex]::Escape($tempDir + '\'), '').Replace('\', '/')
            $contextType = if ($partPath -match 'slideMasters') { 'Master' } else { 'Layout' }
            
            $doc = Get-XmlDocument $file.FullName
            if (-not $doc) { continue }
            
            $partDir = Split-Path $partPath -Parent
            $partFile = Split-Path $partPath -Leaf
            $relsPath = Join-Path $tempDir (($partDir + '/_rels/' + $partFile + '.rels') -replace '/', '\')
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
    
    $slideUsages = @($usages | Where-Object { $_.ContextType -eq 'Slide' })
    $layoutUsages = @($usages | Where-Object { $_.ContextType -eq 'Layout' })
    $masterUsages = @($usages | Where-Object { $_.ContextType -eq 'Master' })
    Write-Verbose "Image usage breakdown: $($slideUsages.Count) in slides, $($layoutUsages.Count) in layouts, $($masterUsages.Count) in masters"
    
    return ,$usages
}

function Get-PictureUsage {
    param(
        [System.Xml.XmlElement]$pic,
        [System.Xml.XmlDocument]$rels,
        [string]$contextType,
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
    $fullPath = Join-Path $tempDir ($mediaPath -replace '/', '\')
    
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
        SvgPartPath = $svgCheck.SvgRId
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
    }
    
    # Build informative verbose description
    $cropInfo = if ($hasSrcRect) {
        $crop = Get-CropPercentage -left $l -top $t -right $r -bottom $b
        "crop [$($crop.WidthPercent.ToString('F1'))%×$($crop.HeightPercent.ToString('F1'))%] (l:$l t:$t r:$r b:$b)"
    } else {
        "no crop"
    }
    
    $flags = @()
    if ($isMorphSlide) { $flags += 'Morph' }
    if ($svgCheck.IsSvgFallback) { $flags += 'SVG-fallback' }
    $flagStr = if ($flags.Count -gt 0) { " [$($flags -join ', ')]" } else { "" }
    
    Write-Verbose "  $location → '$shapeName': $($usage.OriginalFileName) [${displayWidth}×${displayHeight}px, $cropInfo]$flagStr"
    
    return $usage
}

function Get-BlipUsage {
    param(
        [System.Xml.XmlElement]$blip,
        [System.Xml.XmlDocument]$rels,
        [string]$contextType,
        [string]$location,
        [string]$tempDir,
        [string]$partPath,
        [System.Xml.XmlDocument]$slideDocument
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
    $mediaPath = Join-Path $tempDir "ppt\media\$mediaFile"
    
    if (-not (Test-Path -LiteralPath $mediaPath)) {
        Write-Verbose "  Media file not found: $mediaFile"
        return $null
    }
    
    # Check SVG fallback
    $svgCheck = Test-IsSvgFallback -blipElement $blip
    
    # Determine shape name (generic for standalone blips)
    $shapeName = "Background/Fill"
    if ($blip.ParentNode.ParentNode.LocalName -eq 'bg') {
        $shapeName = "Background"
    }
    
    # Create usage object
    $usage = [ImageUsage]@{
        Location = $location
        ContextType = $contextType
        SlideNumber = 0
        ShapeName = $shapeName
        ImagePhysicalPath = $mediaPath
        OriginalFileName = $mediaFile
        SourceWidthPx = 0
        SourceHeightPx = 0
        DisplayWidthPx = 0  # Unknown for backgrounds/fills
        DisplayHeightPx = 0
        HasSrcRect = $false
        SrcRectLeft = 0
        SrcRectTop = 0
        SrcRectRight = 0
        SrcRectBottom = 0
        IsMorphSlide = $false
        MorphPair = $null
        IsSvgFallbackUsage = $svgCheck.IsSvgFallback
        SvgPartPath = $svgCheck.SvgRId
        BlipRId = $rId
        PartPath = $partPath
        BlipElement = $blip
        XfrmElement = $null
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
    }
    
    $svgFlag = if ($svgCheck.IsSvgFallback) { " [SVG-fallback]" } else { "" }
    Write-Verbose "  $location → '$shapeName': $mediaFile [background/fill]$svgFlag"
    
    return $usage
}

#endregion

#region Morph Detection

function Find-MorphPairs {
    param([System.Collections.Generic.List[ImageUsage]]$usages)
    
    $pairs = @{}
    $morphSlideCount = 0
    $totalPairsFound = 0
    $unmatchedImages = 0
    
    $slideUsages = $usages | Where-Object { $_.ContextType -eq 'Slide' } | 
        Group-Object -Property SlideNumber
    
    foreach ($group in $slideUsages) {
        $slideNum = $group.Name
        if ($slideNum -le 1) { continue }
        
        $currentSlide = $group.Group
        $currentMorph = $currentSlide | Where-Object { $_.IsMorphSlide } | Select-Object -First 1
        if (-not $currentMorph) { continue }
        
        $morphSlideCount++
        $prevSlideNum = [int]$slideNum - 1
        $prevSlide = $usages | Where-Object { 
            $_.ContextType -eq 'Slide' -and $_.SlideNumber -eq $prevSlideNum 
        }
        
        if ($PSCmdlet.MyInvocation.BoundParameters['Verbose']) {
            Write-Verbose "Analyzing Morph transition: Slide $prevSlideNum → Slide $slideNum"
        }
        
        # Match by physical image path
        $slideMatchCount = 0
        foreach ($curr in $currentSlide) {
            $match = $prevSlide | Where-Object { $_.ImagePhysicalPath -eq $curr.ImagePhysicalPath } | 
                Select-Object -First 1
            
            if ($match) {
                $curr.MorphPair = $match
                $match.MorphPair = $curr
                
                $pairKey = "$($match.ImagePhysicalPath)|$prevSlideNum-$slideNum"
                $pairs[$pairKey] = @($match, $curr)
                $slideMatchCount++
                $totalPairsFound++
                
                if ($PSCmdlet.MyInvocation.BoundParameters['Verbose']) {
                    $fileName = Split-Path $match.ImagePhysicalPath -Leaf
                    Write-Verbose "  Linked pair: $fileName across slides $prevSlideNum ↔ $slideNum"
                }
            } else {
                $unmatchedImages++
            }
        }
        
        if ($PSCmdlet.MyInvocation.BoundParameters['Verbose'] -and $slideMatchCount -eq 0) {
            Write-Verbose "  Warning: No matching images found between slides $prevSlideNum and $slideNum"
        }
    }
    
    if ($morphSlideCount -gt 0) {
        Write-Host "🔗 Morph detection: $morphSlideCount transition(s), $totalPairsFound image pair(s) linked" -ForegroundColor Cyan
        
        if ($PSCmdlet.MyInvocation.BoundParameters['Verbose']) {
            if ($unmatchedImages -gt 0) {
                Write-Verbose "Note: $unmatchedImages image(s) on morph slides had no match on previous slide"
            }
            Write-Verbose "Morph pairs will be protected from conflicting crop operations"
        }
    }
    
    return $pairs
}

#endregion

#region Crop Normalization

function Invoke-CropNormalization {
    param([ImageUsage]$usage)
    
    if (-not $usage.HasSrcRect) { return }
    
    $l = $usage.SrcRectLeft
    $t = $usage.SrcRectTop
    $r = $usage.SrcRectRight
    $b = $usage.SrcRectBottom
    
    $needsNormalization = $false
    $originalValues = @{}
    
    # Detect out-of-bounds
    if ($l -lt 0 -or $l -gt $script:Config.CROP_THOUSANDTHS_MAX -or
        $t -lt 0 -or $t -gt $script:Config.CROP_THOUSANDTHS_MAX -or
        $r -lt 0 -or $r -gt $script:Config.CROP_THOUSANDTHS_MAX -or
        $b -lt 0 -or $b -gt $script:Config.CROP_THOUSANDTHS_MAX) {
        $needsNormalization = $true
        $originalValues = @{ l=$l; t=$t; r=$r; b=$b }
    }
    
    if (-not $needsNormalization) { return }
    
    # Clamp to valid range
    $l = [Math]::Max(0, [Math]::Min($l, $script:Config.CROP_THOUSANDTHS_MAX))
    $t = [Math]::Max(0, [Math]::Min($t, $script:Config.CROP_THOUSANDTHS_MAX))
    $r = [Math]::Max(0, [Math]::Min($r, $script:Config.CROP_THOUSANDTHS_MAX))
    $b = [Math]::Max(0, [Math]::Min($b, $script:Config.CROP_THOUSANDTHS_MAX))
    
    # Compensate xfrm for normalized crops
    if ($usage.XfrmElement) {
        $ext = $usage.XfrmElement.SelectSingleNode('a:ext', $script:NsMgr)
        $off = $usage.XfrmElement.SelectSingleNode('a:off', $script:NsMgr)
        
        if ($ext -and $off) {
            $origCx = [long]$ext.GetAttribute('cx')
            $origCy = [long]$ext.GetAttribute('cy')
            $origX = [long]$off.GetAttribute('x')
            $origY = [long]$off.GetAttribute('y')
            
            # Calculate visible proportions (accounting for negative = extension)
            $origWidthProp = (100000.0 - $originalValues.l - $originalValues.r) / 100000.0
            $origHeightProp = (100000.0 - $originalValues.t - $originalValues.b) / 100000.0
            $newWidthProp = (100000.0 - $l - $r) / 100000.0
            $newHeightProp = (100000.0 - $t - $b) / 100000.0
            
            # Scale dimensions
            $newCx = [long]($origCx * ($newWidthProp / $origWidthProp))
            $newCy = [long]($origCy * ($newHeightProp / $origHeightProp))
            
            # Adjust position if negative crops were normalized
            $newX = $origX
            $newY = $origY
            
            if ($originalValues.l -lt 0 -and $l -eq 0) {
                $newX = $origX + [long]($origCx * ([Math]::Abs($originalValues.l) / (100000.0 - $originalValues.l - $originalValues.r)))
            }
            if ($originalValues.t -lt 0 -and $t -eq 0) {
                $newY = $origY + [long]($origCy * ([Math]::Abs($originalValues.t) / (100000.0 - $originalValues.t - $originalValues.b)))
            }
            
            # Clamp to minimums
            $newCx = [Math]::Max($script:Config.MIN_EMU_DIMENSION, $newCx)
            $newCy = [Math]::Max($script:Config.MIN_EMU_DIMENSION, $newCy)
            $newX = [Math]::Max(0, $newX)
            $newY = [Math]::Max(0, $newY)
            
            $ext.SetAttribute('cx', $newCx.ToString())
            $ext.SetAttribute('cy', $newCy.ToString())
            $off.SetAttribute('x', $newX.ToString())
            $off.SetAttribute('y', $newY.ToString())
        }
    }
    
    # Update srcRect in XML (sibling of blip within blipFill)
    $blipFill = $usage.BlipElement.ParentNode
    if ($blipFill) {
        $srcRect = $blipFill.SelectSingleNode('a:srcRect', $script:NsMgr)
        if ($srcRect) {
            if ($l -eq 0) { $srcRect.RemoveAttribute('l') } else { $srcRect.SetAttribute('l', ([int]$l).ToString()) }
            if ($t -eq 0) { $srcRect.RemoveAttribute('t') } else { $srcRect.SetAttribute('t', ([int]$t).ToString()) }
            if ($r -eq 0) { $srcRect.RemoveAttribute('r') } else { $srcRect.SetAttribute('r', ([int]$r).ToString()) }
            if ($b -eq 0) { $srcRect.RemoveAttribute('b') } else { $srcRect.SetAttribute('b', ([int]$b).ToString()) }
        }
    }
    
    # Update usage
    $usage.SrcRectLeft = $l
    $usage.SrcRectTop = $t
    $usage.SrcRectRight = $r
    $usage.SrcRectBottom = $b
    $usage.CropNormalized = $true
    
    # Recompute display size
    if ($usage.XfrmElement) {
        $ext = $usage.XfrmElement.SelectSingleNode('a:ext', $script:NsMgr)
        if ($ext) {
            $usage.DisplayWidthPx = ConvertFrom-EMU -emu ([long]$ext.GetAttribute('cx'))
            $usage.DisplayHeightPx = ConvertFrom-EMU -emu ([long]$ext.GetAttribute('cy'))
        }
    }
    
    # Build change description
    $changes = @()
    $details = @()
    foreach ($key in @('l', 't', 'r', 'b')) {
        $oldVal = $originalValues[$key]
        $newVal = Get-Variable -Name $key -ValueOnly
        if ($oldVal -ne $newVal) {
            $changes += "${key}:${oldVal}→${newVal}"
            
            # Build detailed explanation for verbose
            $side = switch ($key) {
                'l' { 'left' }
                't' { 'top' }
                'r' { 'right' }
                'b' { 'bottom' }
            }
            
            $reason = if ($oldVal -lt 0) {
                "extended beyond image boundary (negative value)"
            } elseif ($oldVal -gt $script:Config.CROP_THOUSANDTHS_MAX) {
                "exceeded maximum allowed value ($($script:Config.CROP_THOUSANDTHS_MAX))"
            } else {
                "was out of valid range [0-$($script:Config.CROP_THOUSANDTHS_MAX)]"
            }
            
            $details += "  • $side edge: $oldVal → $newVal (was $reason)"
        }
    }
    
    if ($changes.Count -gt 0) {
        Write-Host "📝 Corrected illegal crop values for '$($usage.ShapeName)' on $($usage.Location) ($($changes -join ', '))" -ForegroundColor Cyan
        
        if ($PSCmdlet.MyInvocation.BoundParameters['Verbose']) {
            Write-Verbose "Illegal crop correction details for '$($usage.ShapeName)':"
            foreach ($detail in $details) {
                Write-Verbose $detail
            }
            
            # Show xfrm adjustments if made
            if ($usage.XfrmElement) {
                $ext = $usage.XfrmElement.SelectSingleNode('a:ext', $script:NsMgr)
                $off = $usage.XfrmElement.SelectSingleNode('a:off', $script:NsMgr)
                if ($ext -and $off) {
                    Write-Verbose "  Adjusted shape transform (xfrm) to maintain visual appearance:"
                    Write-Verbose "    • Position: x=$($off.GetAttribute('x')) EMU, y=$($off.GetAttribute('y')) EMU"
                    Write-Verbose "    • Size: cx=$($ext.GetAttribute('cx')) EMU ($($usage.DisplayWidthPx)px), cy=$($ext.GetAttribute('cy')) EMU ($($usage.DisplayHeightPx)px)"
                }
            }
            
            Write-Verbose "  Crop values are in thousandths (0-100000 range, where 100000 = 100%)"
            Write-Verbose "  Negative values indicate the image was extended beyond its boundaries"
        }
    }
}

#endregion

#region Cropping

function Invoke-ImageCropping {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [System.Collections.Generic.List[ImageUsage]]$usages,
        [string]$tempDir,
        [hashtable]$morphPairs
    )
    
    Write-Host "`n✂️ Processing image crops..." -ForegroundColor Yellow
    
    $croppedCount = 0
    $skippedCount = 0
    
    foreach ($usage in $usages) {
        if (-not $usage.HasSrcRect) { continue }
        
        # SVG fallback check (highest precedence)
        if ($usage.IsSvgFallbackUsage) {
            Write-Host "⏭️ Skipped crop for '$($usage.ShapeName)' on $($usage.Location): SVG fallback (PowerPoint regenerates)" -ForegroundColor Gray
            $usage.OptimizationStatus = 'Skipped_SvgFallback'
            $usage.WhyNotOptimized = 'PNG is auto-generated SVG fallback; PowerPoint regenerates'
            $usage.ManualActionRequired = $false
            $skippedCount++
            continue
        }
        
        # Check if cropping is enabled
        if (-not $EnableCropping) {
            Write-Host "⏭️ Skipped crop for '$($usage.ShapeName)' on $($usage.Location): cropping disabled; enable -EnableCropping" -ForegroundColor Gray
            $usage.OptimizationStatus = 'Skipped_CropNotMaterialized'
            $usage.WhyNotOptimized = 'Cropping disabled; enable -EnableCropping'
            $usage.ManualActionRequired = $true
            $usage.ManualActionHint = 'Run with -EnableCropping'
            $skippedCount++
            continue
        }
        
        # Check scope (masters/layouts)
        if (($usage.ContextType -eq 'Master' -or $usage.ContextType -eq 'Layout') -and -not $CropMastersAndLayouts) {
            Write-Host "⏭️ Skipped crop for '$($usage.ShapeName)' on $($usage.Location): enable -CropMastersAndLayouts" -ForegroundColor Gray
            $usage.OptimizationStatus = 'Skipped_CropNotMaterialized'
            $usage.WhyNotOptimized = 'Master/Layout cropping disabled; enable -CropMastersAndLayouts'
            $usage.ManualActionRequired = $true
            $usage.ManualActionHint = 'Run with -EnableCropping -CropMastersAndLayouts'
            $skippedCount++
            continue
        }
        
        # Check for no-op crop
        if (Test-IsNoOpCrop -left $usage.SrcRectLeft -top $usage.SrcRectTop -right $usage.SrcRectRight -bottom $usage.SrcRectBottom) {
            $blipFill = $usage.BlipElement.ParentNode
            if ($blipFill) {
                $srcRect = $blipFill.SelectSingleNode('a:srcRect', $script:NsMgr)
                if ($srcRect) {
                    $null = $srcRect.ParentNode.RemoveChild($srcRect)
                }
            }
            $usage.HasSrcRect = $false
            $usage.CropRemovedNoOp = $true
            Write-Host "📝 Removed no-op crop for '$($usage.ShapeName)' on $($usage.Location) (full image)" -ForegroundColor Cyan
            continue
        }
        
        # Check Morph pairs
        if ($usage.MorphPair) {
            $pair = $usage.MorphPair
            $usageHasMeaningful = -not (Test-IsNoOpCrop -left $usage.SrcRectLeft -top $usage.SrcRectTop -right $usage.SrcRectRight -bottom $usage.SrcRectBottom)
            $pairHasMeaningful = $pair.HasSrcRect -and -not (Test-IsNoOpCrop -left $pair.SrcRectLeft -top $pair.SrcRectTop -right $pair.SrcRectRight -bottom $pair.SrcRectBottom)
            
            if ($usageHasMeaningful -or $pairHasMeaningful) {
                Write-Host "🚫 Morph crop conflict for '$($usage.ShapeName)' across Slide $($pair.SlideNumber) ↔ $($usage.SlideNumber) (skipping crop & optimization)" -ForegroundColor Red
                $usage.OptimizationStatus = 'Skipped_MorphCropConflict'
                $pair.OptimizationStatus = 'Skipped_MorphCropConflict'
                $usage.WhyNotOptimized = 'Morph transition with crop conflict'
                $pair.WhyNotOptimized = 'Morph transition with crop conflict'
                $usage.ManualActionRequired = $true
                $pair.ManualActionRequired = $true
                $usage.ManualActionHint = 'Manually align crops or remove Morph'
                $pair.ManualActionHint = 'Manually align crops or remove Morph'
                $skippedCount += 2
                continue
            }
        }
        
        # Validate crop geometry
        if (-not (Test-CropValidity -usage $usage)) {
            Write-Host "🚫 Skipped crop for '$($usage.ShapeName)' on $($usage.Location): invalid crop geometry" -ForegroundColor Red
            $usage.OptimizationStatus = 'Skipped_CropInvalidOrUnsafe'
            $usage.WhyNotOptimized = 'Invalid crop geometry'
            $usage.ManualActionRequired = $true
            $usage.ManualActionHint = 'Review and fix crop values'
            $skippedCount++
            continue
        }
        
        # Perform crop
        if ($PSCmdlet.ShouldProcess("$($usage.ShapeName) on $($usage.Location)", "Crop image")) {
            $result = Invoke-CropOperation -usage $usage -tempDir $tempDir
            if ($result.Success) {
                $croppedCount++
            } else {
                $skippedCount++
            }
        }
    }
    
    Write-Host "📊 Cropping phase complete: $croppedCount cropped, $skippedCount skipped" -ForegroundColor Green
}

function Test-CropValidity {
    param([ImageUsage]$usage)
    
    $l = $usage.SrcRectLeft
    $t = $usage.SrcRectTop
    $r = $usage.SrcRectRight
    $b = $usage.SrcRectBottom
    
    # Check bounds
    if ($l -lt 0 -or $l -gt $script:Config.CROP_THOUSANDTHS_MAX) {
        if ($PSCmdlet.MyInvocation.BoundParameters['Verbose']) {
            Write-Verbose "  Crop validation failed for '$($usage.ShapeName)': left=$l out of range [0-$($script:Config.CROP_THOUSANDTHS_MAX)]"
        }
        return $false
    }
    if ($t -lt 0 -or $t -gt $script:Config.CROP_THOUSANDTHS_MAX) {
        if ($PSCmdlet.MyInvocation.BoundParameters['Verbose']) {
            Write-Verbose "  Crop validation failed for '$($usage.ShapeName)': top=$t out of range [0-$($script:Config.CROP_THOUSANDTHS_MAX)]"
        }
        return $false
    }
    if ($r -lt 0 -or $r -gt $script:Config.CROP_THOUSANDTHS_MAX) {
        if ($PSCmdlet.MyInvocation.BoundParameters['Verbose']) {
            Write-Verbose "  Crop validation failed for '$($usage.ShapeName)': right=$r out of range [0-$($script:Config.CROP_THOUSANDTHS_MAX)]"
        }
        return $false
    }
    if ($b -lt 0 -or $b -gt $script:Config.CROP_THOUSANDTHS_MAX) {
        if ($PSCmdlet.MyInvocation.BoundParameters['Verbose']) {
            Write-Verbose "  Crop validation failed for '$($usage.ShapeName)': bottom=$b out of range [0-$($script:Config.CROP_THOUSANDTHS_MAX)]"
        }
        return $false
    }
    
    # Check sum
    if (($l + $r) -ge $script:Config.CROP_THOUSANDTHS_MAX) {
        if ($PSCmdlet.MyInvocation.BoundParameters['Verbose']) {
            Write-Verbose "  Crop validation failed for '$($usage.ShapeName)': left+right=$($l+$r) >= $($script:Config.CROP_THOUSANDTHS_MAX) (crops overlap horizontally)"
        }
        return $false
    }
    if (($t + $b) -ge $script:Config.CROP_THOUSANDTHS_MAX) {
        if ($PSCmdlet.MyInvocation.BoundParameters['Verbose']) {
            Write-Verbose "  Crop validation failed for '$($usage.ShapeName)': top+bottom=$($t+$b) >= $($script:Config.CROP_THOUSANDTHS_MAX) (crops overlap vertically)"
        }
        return $false
    }
    
    # Check minimum dimensions
    $crop = Get-CropPercentage -left $l -top $t -right $r -bottom $b
    if ($crop.WidthPercent -lt $script:Config.MIN_CROPPED_DIMENSION_PERCENT) {
        if ($PSCmdlet.MyInvocation.BoundParameters['Verbose']) {
            Write-Verbose "  Crop validation failed for '$($usage.ShapeName)': resulting width $($crop.WidthPercent.ToString('F1'))% < minimum $($script:Config.MIN_CROPPED_DIMENSION_PERCENT)%"
        }
        return $false
    }
    if ($crop.HeightPercent -lt $script:Config.MIN_CROPPED_DIMENSION_PERCENT) {
        if ($PSCmdlet.MyInvocation.BoundParameters['Verbose']) {
            Write-Verbose "  Crop validation failed for '$($usage.ShapeName)': resulting height $($crop.HeightPercent.ToString('F1'))% < minimum $($script:Config.MIN_CROPPED_DIMENSION_PERCENT)%"
        }
        return $false
    }
    
    if ($PSCmdlet.MyInvocation.BoundParameters['Verbose']) {
        Write-Verbose "  Crop validation passed for '$($usage.ShapeName)': l=$l, t=$t, r=$r, b=$b → $($crop.WidthPercent.ToString('F1'))%×$($crop.HeightPercent.ToString('F1'))% of original"
    }
    
    return $true
}

function Invoke-CropOperation {
    param(
        [ImageUsage]$usage,
        [string]$tempDir
    )
    
    try {
        # Get source image dimensions
        $identify = & $script:MagickExe identify -format "%w %h" $usage.ImagePhysicalPath 2>&1
        if ($LASTEXITCODE -ne 0) {
            throw "ImageMagick identify failed: $identify"
        }
        
        # Handle array results
        if ($identify -is [array]) {
            $identify = $identify[0]
        }
        
        $dims = $identify -split '\s+'
        $srcWidth = [int]$dims[0]
        $srcHeight = [int]$dims[1]
        
        # Store source dimensions
        $usage.SourceWidthPx = $srcWidth
        $usage.SourceHeightPx = $srcHeight
        
        # Calculate crop rectangle
        $leftPx = [Math]::Round($srcWidth * $usage.SrcRectLeft / 100000.0)
        $topPx = [Math]::Round($srcHeight * $usage.SrcRectTop / 100000.0)
        $rightPx = [Math]::Round($srcWidth * $usage.SrcRectRight / 100000.0)
        $bottomPx = [Math]::Round($srcHeight * $usage.SrcRectBottom / 100000.0)
        
        $cropWidth = $srcWidth - $leftPx - $rightPx
        $cropHeight = $srcHeight - $topPx - $bottomPx
        
        # Clamp
        $cropWidth = [Math]::Max(1, [Math]::Min($cropWidth, $srcWidth))
        $cropHeight = [Math]::Max(1, [Math]::Min($cropHeight, $srcHeight))
        $leftPx = [Math]::Max(0, [Math]::Min($leftPx, $srcWidth - 1))
        $topPx = [Math]::Max(0, [Math]::Min($topPx, $srcHeight - 1))
        
        # Create output filename
        $ext = [System.IO.Path]::GetExtension($usage.ImagePhysicalPath)
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($usage.ImagePhysicalPath)
        $outputName = "${baseName}_cropped$ext"
        $outputPath = Join-Path (Split-Path $usage.ImagePhysicalPath -Parent) $outputName
        
        # Make unique
        $counter = 1
        while (Test-Path $outputPath) {
            $outputName = "${baseName}_cropped${counter}$ext"
            $outputPath = Join-Path (Split-Path $usage.ImagePhysicalPath -Parent) $outputName
            $counter++
        }
        
        # Execute crop
        $cropGeometry = "${cropWidth}x${cropHeight}+${leftPx}+${topPx}"
        
        if ($PSCmdlet.MyInvocation.BoundParameters['Verbose']) {
            Write-Verbose "  Materializing crop for '$($usage.ShapeName)':"
            Write-Verbose "    Source: ${srcWidth}×${srcHeight}px"
            Write-Verbose "    Crop values: l=$($usage.SrcRectLeft), t=$($usage.SrcRectTop), r=$($usage.SrcRectRight), b=$($usage.SrcRectBottom)"
            Write-Verbose "    Resulting geometry: $cropGeometry (${cropWidth}×${cropHeight}px)"
        }
        
        & $script:MagickExe $usage.ImagePhysicalPath -crop $cropGeometry +repage $outputPath 2>&1 | Out-Null
        
        if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $outputPath)) {
            throw "Crop operation failed"
        }
        
        $beforeSize = $usage.BeforeSizeBytes
        $afterSize = (Get-Item $outputPath).Length
        $savedPercent = (($beforeSize - $afterSize) / $beforeSize) * 100
        
        if ($PSCmdlet.MyInvocation.BoundParameters['Verbose']) {
            Write-Verbose "    Output: $outputName [$(Format-ByteSize $beforeSize) → $(Format-ByteSize $afterSize), saved $($savedPercent.ToString('F1'))%]"
        }
        
        # Update relationships
        $partDir = Split-Path $usage.PartPath -Parent
        $partFile = Split-Path $usage.PartPath -Leaf
        $relsPath = Join-Path $tempDir (($partDir + '/_rels/' + $partFile + '.rels') -replace '/', '\')
        
        if (-not (Test-Path -LiteralPath $relsPath)) {
            throw "Rels file not found: $relsPath"
        }
        
        $relsDoc = Get-XmlDocument $relsPath
        
        # Remove old relationship
        $oldRId = $usage.BlipRId
        if ($oldRId) {
            $oldRel = $relsDoc.SelectSingleNode("//x:Relationship[@Id='$oldRId']", $script:NsMgr)
            if ($oldRel) {
                $null = $oldRel.ParentNode.RemoveChild($oldRel)
            }
        }
        
        # Create new relationship
        $newRId = Get-UniqueRId -relsDoc $relsDoc
        $rel = $relsDoc.CreateElement('Relationship', $relsDoc.DocumentElement.NamespaceURI)
        [void]$rel.SetAttribute('Id', $newRId)
        [void]$rel.SetAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image')
        [void]$rel.SetAttribute('Target', "../media/$outputName")
        $null = $relsDoc.DocumentElement.AppendChild($rel)
        Save-XmlDocument -doc $relsDoc -path $relsPath
        
        # Update blip
        [void]$usage.BlipElement.SetAttribute('embed', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships', $newRId)
        
        # CRITICAL: Update BlipRId so subsequent operations (like PNG→JPEG) can find the current relationship
        $usage.BlipRId = $newRId
        
        # Remove srcRect (sibling of blip within blipFill)
        $blipFill = $usage.BlipElement.ParentNode
        if ($blipFill) {
            $srcRect = $blipFill.SelectSingleNode('a:srcRect', $script:NsMgr)
            if ($srcRect) {
                $null = $srcRect.ParentNode.RemoveChild($srcRect)
            }
        }
        
        $usage.HasSrcRect = $false
        $usage.CropApplied = $true
        $usage.AfterSizeBytes = $afterSize
        $usage.OptimizedFile = $outputName
        $usage.ImagePhysicalPath = $outputPath  # CRITICAL: Update path to cropped file for optimization phase
        
        if ($PSCmdlet.MyInvocation.BoundParameters['Verbose']) {
            Write-Verbose "    Updated ImagePhysicalPath: $outputPath (optimization will use cropped file)"
        }
        
        $savedPercent = (($beforeSize - $afterSize) / $beforeSize) * 100
        Write-Host "✂️ Cropped '$($usage.ShapeName)' on $($usage.Location) (saved $($savedPercent.ToString('F1'))%)" -ForegroundColor Green
        
        return @{ Success = $true }
        
    } catch {
        Write-Warning "Crop failed for '$($usage.ShapeName)' on $($usage.Location): $_"
        $usage.OptimizationStatus = 'Skipped_CropInvalidOrUnsafe'
        $usage.WhyNotOptimized = "Crop operation failed: $_"
        return @{ Success = $false }
    }
}

#endregion

#region Optimization

function Invoke-ImageOptimization {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [System.Collections.Generic.List[ImageUsage]]$usages,
        [string]$tempDir,
        [hashtable]$morphPairs
    )
    
    Write-Host "`n⚙️ Optimizing images..." -ForegroundColor Yellow
    
    # Group by physical image
    $imageGroups = @{}
    foreach ($usage in $usages) {
        if (-not $imageGroups.ContainsKey($usage.ImagePhysicalPath)) {
            $group = [ImageGroup]::new()
            $group.PhysicalPath = $usage.ImagePhysicalPath
            $group.IsSvgFallbackImage = $usage.IsSvgFallbackUsage
            # Read file size from disk - will be cropped file if crop was applied
            if (Test-Path -LiteralPath $usage.ImagePhysicalPath) {
                $group.OriginalSizeBytes = (Get-Item -LiteralPath $usage.ImagePhysicalPath).Length
            } else {
                $group.OriginalSizeBytes = $usage.BeforeSizeBytes
            }
            $imageGroups[$usage.ImagePhysicalPath] = $group
        }
        $imageGroups[$usage.ImagePhysicalPath].Usages.Add($usage)
    }
    
    $optimizedCount = 0
    $skippedCount = 0
    
    # Process groups in slide order (by minimum slide number of usages)
    $sortedGroups = $imageGroups.Values | Sort-Object { 
        $slideUsages = $_.Usages | Where-Object { $_.ContextType -eq 'Slide' }
        if ($slideUsages) {
            ($slideUsages | Measure-Object -Property SlideNumber -Minimum).Minimum
        } else {
            9999  # No slide usages (master/layout only), sort to end
        }
    }
    
    foreach ($group in $sortedGroups) {
        # Check if optimization forbidden
        $skipReason = Get-OptimizationSkipReason -group $group
        if ($skipReason) {
            foreach ($usage in $group.Usages) {
                if ($usage.OptimizationStatus -eq 'Pending') {
                    $usage.OptimizationStatus = $skipReason.Status
                    $usage.WhyNotOptimized = $skipReason.Reason
                    $usage.ManualActionRequired = $skipReason.ManualAction
                    $usage.ManualActionHint = $skipReason.Hint
                }
            }
            # Log first usage as representative
            $firstUsage = $group.Usages[0]
            Write-Host "⏭️ Skipped optimization for '$($firstUsage.ShapeName)' on $($firstUsage.Location): $($skipReason.Reason)" -ForegroundColor Gray
            $skippedCount += $group.Usages.Count
            continue
        }
        
        # Calculate target dimensions
        $maxDisplay = @{
            Width = ($group.Usages | Measure-Object -Property DisplayWidthPx -Maximum).Maximum
            Height = ($group.Usages | Measure-Object -Property DisplayHeightPx -Maximum).Maximum
        }
        
        $mediaFile = Split-Path $group.PhysicalPath -Leaf
        $firstUsage = $group.Usages[0]
        
        if ($PSCmdlet.ShouldProcess($mediaFile, "Optimize image")) {
            Write-Verbose "Optimizing $mediaFile [display: $($maxDisplay.Width)×$($maxDisplay.Height)px] from $($firstUsage.Location)"
            
            $result = Invoke-OptimizationOperation -group $group -maxDisplay $maxDisplay -tempDir $tempDir
            if ($result.Success) {
                $optimizedCount++
            } else {
                $skippedCount += $group.Usages.Count
            }
        }
    }
    
    Write-Host "📊 Optimization phase complete: $optimizedCount optimized, $skippedCount skipped" -ForegroundColor Green
}

function Get-OptimizationSkipReason {
    param([ImageGroup]$group)
    
    # Check file format - skip vector formats and other non-optimizable types
    $ext = [System.IO.Path]::GetExtension($group.PhysicalPath).ToLower()
    $supportedFormats = @('.png', '.jpg', '.jpeg')
    
    if ($ext -notin $supportedFormats) {
        $fileName = Split-Path $group.PhysicalPath -Leaf
        $formatType = switch ($ext) {
            '.emf' { 'vector graphic (EMF)' }
            '.wmf' { 'vector graphic (WMF)' }
            '.svg' { 'vector graphic (SVG)' }
            '.gif' { 'animated/indexed image (GIF)' }
            '.bmp' { 'uncompressed bitmap (BMP)' }
            '.tif' { 'TIFF image' }
            '.tiff' { 'TIFF image' }
            default { "format ($ext)" }
        }
        
        return @{
            Status = 'Skipped_UnsupportedFormat'
            Reason = "Unsupported $formatType - only PNG/JPEG optimization supported"
            ManualAction = $false
            Hint = ''
        }
    }
    
    # Capture source dimensions for reporting, even if image will be skipped
    try {
        $identify = & $script:MagickExe identify -format "%w %h" $group.PhysicalPath 2>&1
        if ($LASTEXITCODE -eq 0) {
            if ($identify -is [array]) {
                $identify = $identify[0]
            }
            $dims = $identify -split '\s+'
            $srcWidth = [int]$dims[0]
            $srcHeight = [int]$dims[1]
            
            # Store in all usages for CSV reporting
            foreach ($usage in $group.Usages) {
                if ($usage.SourceWidthPx -eq 0) {
                    $usage.SourceWidthPx = $srcWidth
                    $usage.SourceHeightPx = $srcHeight
                }
            }
        }
    } catch {
        # Silently continue if dimension capture fails
    }
    
    # Check SVG fallback
    if ($group.IsSvgFallbackImage) {
        return @{
            Status = 'Skipped_SvgFallback'
            Reason = 'PNG is auto-generated SVG fallback; PowerPoint regenerates'
            ManualAction = $false
            Hint = ''
        }
    }
    
    # Check for any already-set skip status (crop-related or Morph conflicts)
    $alreadySkipped = $group.Usages | Where-Object { 
        $_.OptimizationStatus -like 'Skipped_*' 
    } | Select-Object -First 1
    
    if ($alreadySkipped) {
        # Return the existing skip reason to prevent optimization and log it properly
        return @{
            Status = $alreadySkipped.OptimizationStatus
            Reason = $alreadySkipped.WhyNotOptimized
            ManualAction = $alreadySkipped.ManualActionRequired
            Hint = $alreadySkipped.ManualActionHint
        }
    }
    
    # Check masters/layouts without opt-in
    $hasTemplate = $group.Usages | Where-Object { $_.ContextType -in @('Master', 'Layout') }
    
    $contextTypes = ($group.Usages | Select-Object -ExpandProperty ContextType -Unique) -join ', '
    $fileName = Split-Path $group.PhysicalPath -Leaf
    Write-Verbose "  Checking skip reasons for $($fileName): contexts=[$contextTypes], hasTemplate=$($null -ne $hasTemplate), OptimizeMastersAndLayouts=$OptimizeMastersAndLayouts"
    
    if ($hasTemplate -and -not $OptimizeMastersAndLayouts) {
        return @{
            Status = 'Skipped_SharedWithTemplatesFlagMissing'
            Reason = 'Used in master/layout; enable -OptimizeMastersAndLayouts'
            ManualAction = $true
            Hint = 'Run with -OptimizeMastersAndLayouts'
        }
    }
    
    # Check unresolved crops
    $hasUnresolvedCrop = @($group.Usages | Where-Object { $_.HasSrcRect -and -not $_.CropApplied })
    if ($hasUnresolvedCrop.Count -gt 0) {
        return @{
            Status = 'Skipped_CropNotMaterialized'
            Reason = 'Crop present; enable -EnableCropping'
            ManualAction = $true
            Hint = 'Run with -EnableCropping'
        }
    }
    
    return $null
}

function Invoke-OptimizationOperation {
    param(
        [ImageGroup]$group,
        [hashtable]$maxDisplay,
        [string]$tempDir
    )
    
    try {
        # Get source dimensions
        $identify = & $script:MagickExe identify -format "%w %h" $group.PhysicalPath 2>&1
        if ($LASTEXITCODE -ne 0) {
            throw "ImageMagick identify failed: $identify"
        }
        
        # Handle array results
        if ($identify -is [array]) {
            $identify = $identify[0]
        }
        
        $info = $identify -split '\s+'
        $srcWidth = [int]$info[0]
        $srcHeight = [int]$info[1]
        
        # Calculate target size
        $desiredWidth = [Math]::Ceiling($maxDisplay.Width * $HeadroomFactor)
        $desiredHeight = [Math]::Ceiling($maxDisplay.Height * $HeadroomFactor)
        
        # For background/fill images with no display size, use source dimensions
        if ($desiredWidth -eq 0 -or $desiredHeight -eq 0) {
            $desiredWidth = $srcWidth
            $desiredHeight = $srcHeight
        }
        
        $targetWidth = [Math]::Min($srcWidth, $desiredWidth)
        $targetHeight = [Math]::Min($srcHeight, $desiredHeight)
        
        # Update usages
        foreach ($usage in $group.Usages) {
            $usage.SourceWidthPx = $srcWidth
            $usage.SourceHeightPx = $srcHeight
            $usage.TargetWidthPx = $targetWidth
            $usage.TargetHeightPx = $targetHeight
        }
        
        Write-Verbose "  Dimensions: source ${srcWidth}×${srcHeight}px → target ${targetWidth}×${targetHeight}px [HeadroomFactor: $HeadroomFactor]"
        
        # Check if already at target
        if ($srcWidth -eq $targetWidth -and $srcHeight -eq $targetHeight) {
            # Check if PNG can be converted to JPEG
            if ($group.PhysicalPath -match '\.png$') {
                $transp = Get-ImageTransparency -imagePath $group.PhysicalPath
                foreach ($usage in $group.Usages) {
                    $usage.EffectiveTransparencyPercent = $transp
                }
                
                if ($transp -le $TransparencyThresholdPercent) {
                    # Opaque - convert PNG to JPEG
                    Write-Host "📝 Effective transparency $($transp.ToString('F1'))% (≤$($TransparencyThresholdPercent.ToString('F1'))% threshold) — treating as opaque" -ForegroundColor Cyan
                    $result = ConvertTo-Jpeg -group $group -tempDir $tempDir -resize $false
                    return $result
                }
            }
            
            $firstUsage = $group.Usages[0]
            foreach ($usage in $group.Usages) {
                $usage.OptimizationStatus = 'NoChangeNeeded'
                $usage.WhyNotOptimized = 'Already at target size'
                $usage.ManualActionRequired = $false
            }
            Write-Host "📝 No change needed for '$($firstUsage.ShapeName)' on $($firstUsage.Location) (already at target)" -ForegroundColor Cyan
            return @{ Success = $false }
        }
        
        # Determine optimization strategy
        $ext = [System.IO.Path]::GetExtension($group.PhysicalPath).ToLower()
        
        if ($ext -eq '.jpg' -or $ext -eq '.jpeg') {
            $result = Optimize-Jpeg -group $group -targetWidth $targetWidth -targetHeight $targetHeight -tempDir $tempDir
        } elseif ($ext -eq '.png') {
            # Always check real transparency for PNGs (even if alpha channel exists)
            $transp = Get-ImageTransparency -imagePath $group.PhysicalPath
            foreach ($usage in $group.Usages) {
                $usage.EffectiveTransparencyPercent = $transp
            }
            
            if ($transp -le $TransparencyThresholdPercent) {
                # Opaque or negligible transparency - convert to JPEG
                Write-Host "📝 Effective transparency $($transp.ToString('F1'))% (≤$($TransparencyThresholdPercent.ToString('F1'))% threshold) — treating as opaque" -ForegroundColor Cyan
                $result = ConvertTo-Jpeg -group $group -tempDir $tempDir -resize $true
            } else {
                # Has meaningful transparency - keep as PNG
                $result = Optimize-PngWithAlpha -group $group -targetWidth $targetWidth -targetHeight $targetHeight -tempDir $tempDir
            }
        } else {
            foreach ($usage in $group.Usages) {
                $usage.OptimizationStatus = 'Skipped_Other'
                $usage.WhyNotOptimized = 'Unsupported format'
            }
            return @{ Success = $false }
        }
        
        return $result
        
    } catch {
        Write-Warning "Optimization failed for $($group.PhysicalPath): $_"
        foreach ($usage in $group.Usages) {
            $usage.OptimizationStatus = 'Skipped_Other'
            $usage.WhyNotOptimized = "Optimization failed: $_"
        }
        return @{ Success = $false }
    }
}

function Get-ImageTransparency {
    param([string]$imagePath)
    
    $fileName = Split-Path $imagePath -Leaf
    
    # Use magick (not convert) to check if image has actual transparency
    # Returns percentage of pixels that are NOT fully opaque
    $result = & $script:MagickExe $imagePath -alpha extract -format "%[fx:100*(1-mean)]" info: 2>&1 | 
        Where-Object { $_ -notmatch '^(WARNING|System\.Management)' -and $_ -match '^\d+\.?\d*$' } |
        Select-Object -First 1
    
    if ($result -and $result -match '^\d+\.?\d*$') {
        $transp = [double]$result
        
        if ($PSCmdlet.MyInvocation.BoundParameters['Verbose']) {
            $decision = if ($transp -le $TransparencyThresholdPercent) { "opaque (can convert to JPEG)" } else { "transparent (keep as PNG)" }
            Write-Verbose "  Transparency analysis for $($fileName): $($transp.ToString('F2'))% transparent pixels → $decision"
        }
        
        return $transp
    }
    
    if ($PSCmdlet.MyInvocation.BoundParameters['Verbose']) {
        Write-Verbose "  Transparency analysis for $($fileName): unable to determine, assuming opaque (0%)"
    }
    
    return 0.0
}

function Optimize-Jpeg {
    param(
        [ImageGroup]$group,
        [int]$targetWidth,
        [int]$targetHeight,
        [string]$tempDir
    )
    
    Write-Verbose "  Optimizing JPEG: target ${targetWidth}×${targetHeight}px, quality $JpegQuality"
    
    $outputPath = "$($group.PhysicalPath).tmp"
    
    $magickArgs = @(
        $group.PhysicalPath,
        '-filter', $script:Config.RESIZE_FILTER,
        '-resize', "${targetWidth}x${targetHeight}>",
        '-strip',
        '-quality', $JpegQuality,
        '-define', 'jpeg:optimize-coding=true',
        '-define', 'jpeg:dct-method=float',
        '-sampling-factor', $script:Config.JPEG_SAMPLING_FACTOR_DEFAULT,
        $outputPath
    )
    
    & $script:MagickExe $magickArgs 2>&1 | Out-Null
    
    if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $outputPath)) {
        throw "JPEG optimization failed"
    }
    
    $beforeSize = $group.OriginalSizeBytes
    $afterSize = (Get-Item $outputPath).Length
    
    if ($afterSize -ge $beforeSize) {
        Remove-Item -LiteralPath $outputPath -Force
        $firstUsage = $group.Usages[0]
        foreach ($usage in $group.Usages) {
            $usage.OptimizationStatus = 'Skipped_NoNetSavings'
            $usage.WhyNotOptimized = 'No net savings achieved'
            $usage.ManualActionRequired = $false
        }
        
        if ($PSCmdlet.MyInvocation.BoundParameters['Verbose']) {
            $increase = $afterSize - $beforeSize
            Write-Verbose "  JPEG optimization rejected (no worse than before): $(Format-ByteSize $beforeSize) → $(Format-ByteSize $afterSize) (+$(Format-ByteSize $increase))"
            Write-Verbose "    Likely already well-optimized or at target quality — keeping original"
        }
        
        Write-Host "📝 No net savings for '$($firstUsage.ShapeName)' on $($firstUsage.Location) — kept original" -ForegroundColor Cyan
        return @{ Success = $false }
    }
    
    # Replace original
    Move-Item $outputPath $group.PhysicalPath -Force
    $group.OptimizedSizeBytes = $afterSize
    
    $savedPercent = (($beforeSize - $afterSize) / $beforeSize) * 100
    $firstUsage = $group.Usages[0]
    
    foreach ($usage in $group.Usages) {
        $usage.AfterSizeBytes = $afterSize
        $usage.OptimizationStatus = 'OptimizedJpeg'
    }
    
    Write-Host "⚙️ Optimized JPEG for '$($firstUsage.ShapeName)' on $($firstUsage.Location) (saved $($savedPercent.ToString('F1'))%)" -ForegroundColor Green
    
    return @{ Success = $true }
}

function Optimize-PngWithAlpha {
    param(
        [ImageGroup]$group,
        [int]$targetWidth,
        [int]$targetHeight,
        [string]$tempDir
    )
    
    Write-Verbose "  Optimizing PNG with transparency: target ${targetWidth}×${targetHeight}px, compression level $($script:Config.PNG_COMPRESSION_LEVEL)"
    
    $outputPath = "$($group.PhysicalPath).tmp"
    
    $magickArgs = @(
        $group.PhysicalPath,
        '-filter', $script:Config.RESIZE_FILTER,
        '-resize', "${targetWidth}x${targetHeight}>",
        '-strip',
        '-define', "png:compression-level=$($script:Config.PNG_COMPRESSION_LEVEL)",
        '-define', 'png:color-type=6',
        $outputPath
    )
    
    & $script:MagickExe $magickArgs 2>&1 | Out-Null
    
    if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $outputPath)) {
        throw "PNG optimization failed"
    }
    
    $beforeSize = $group.OriginalSizeBytes
    $afterSize = (Get-Item $outputPath).Length
    
    if ($afterSize -ge $beforeSize) {
        Remove-Item -LiteralPath $outputPath -Force
        $firstUsage = $group.Usages[0]
        foreach ($usage in $group.Usages) {
            $usage.OptimizationStatus = 'Skipped_NoNetSavings'
            $usage.WhyNotOptimized = 'No net savings achieved'
            $usage.ManualActionRequired = $false
        }
        
        if ($PSCmdlet.MyInvocation.BoundParameters['Verbose']) {
            $increase = $afterSize - $beforeSize
            Write-Verbose "  PNG optimization rejected (no worse than before): $(Format-ByteSize $beforeSize) → $(Format-ByteSize $afterSize) (+$(Format-ByteSize $increase))"
            Write-Verbose "    Likely already well-compressed or target size creates larger file — keeping original"
        }
        
        Write-Host "📝 No net savings for '$($firstUsage.ShapeName)' on $($firstUsage.Location) — kept original" -ForegroundColor Cyan
        return @{ Success = $false }
    }
    
    # Replace original
    Move-Item $outputPath $group.PhysicalPath -Force
    $group.OptimizedSizeBytes = $afterSize
    
    $savedPercent = (($beforeSize - $afterSize) / $beforeSize) * 100
    $firstUsage = $group.Usages[0]
    
    foreach ($usage in $group.Usages) {
        $usage.AfterSizeBytes = $afterSize
        $usage.OptimizationStatus = 'OptimizedPngAlpha'
    }
    
    Write-Host "⚙️ Optimized PNG with transparency for '$($firstUsage.ShapeName)' on $($firstUsage.Location) (saved $($savedPercent.ToString('F1'))%)" -ForegroundColor Green
    
    return @{ Success = $true }
}

function ConvertTo-Jpeg {
    param(
        [ImageGroup]$group,
        [string]$tempDir,
        [bool]$resize
    )
    
    $srcFile = Split-Path $group.PhysicalPath -Leaf
    Write-Verbose "  Converting PNG→JPEG: $srcFile [resize: $resize, quality: $JpegQuality]"
    
    # Generate unique JPEG filename
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($group.PhysicalPath)
    $outputName = "$baseName.jpeg"
    $outputPath = Join-Path (Split-Path $group.PhysicalPath -Parent) $outputName
    
    $counter = 1
    while (Test-Path $outputPath) {
        $outputName = "$baseName$counter.jpeg"
        $outputPath = Join-Path (Split-Path $group.PhysicalPath -Parent) $outputName
        $counter++
    }
    
    $magickArgs = @($group.PhysicalPath)
    
    if ($resize) {
        $usage = $group.Usages[0]
        $magickArgs += @(
            '-filter', $script:Config.RESIZE_FILTER,
            '-resize', "$($usage.TargetWidthPx)x$($usage.TargetHeightPx)>"
        )
    }
    
    $magickArgs += @(
        '-strip',
        '-quality', $JpegQuality,
        '-define', 'jpeg:optimize-coding=true',
        '-define', 'jpeg:dct-method=float',
        '-sampling-factor', $script:Config.JPEG_SAMPLING_FACTOR_DEFAULT,
        $outputPath
    )
    
    & $script:MagickExe $magickArgs 2>&1 | Out-Null
    
    if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $outputPath)) {
        throw "PNG to JPEG conversion failed"
    }
    
    $beforeSize = $group.OriginalSizeBytes
    $afterSize = (Get-Item $outputPath).Length
    
    if ($afterSize -ge $beforeSize) {
        Remove-Item -LiteralPath $outputPath -Force
        $firstUsage = $group.Usages[0]
        foreach ($usage in $group.Usages) {
            $usage.OptimizationStatus = 'Skipped_NoNetSavings'
            $usage.WhyNotOptimized = 'Conversion yielded no savings'
            $usage.ManualActionRequired = $false
        }
        
        if ($PSCmdlet.MyInvocation.BoundParameters['Verbose']) {
            $increase = $afterSize - $beforeSize
            Write-Verbose "  PNG→JPEG conversion rejected (no worse than before): $(Format-ByteSize $beforeSize) → $(Format-ByteSize $afterSize) (+$(Format-ByteSize $increase))"
            Write-Verbose "    Original PNG is more efficient than JPEG at this quality — keeping original"
        }
        
        Write-Host "📝 PNG→JPEG conversion yielded no savings for '$($firstUsage.ShapeName)' on $($firstUsage.Location) — kept original" -ForegroundColor Cyan
        return @{ Success = $false }
    }
    
    # Update all relationships pointing to this image
    Write-Verbose "  Updating $($group.Usages.Count) relationship(s) to point to new JPEG: $outputName"
    
    foreach ($usage in $group.Usages) {
        $partDir = Split-Path $usage.PartPath -Parent
        $partFile = Split-Path $usage.PartPath -Leaf
        $relsPath = Join-Path $tempDir (($partDir + '/_rels/' + $partFile + '.rels') -replace '/', '\')
        
        if (-not (Test-Path -LiteralPath $relsPath)) {
            Write-Warning "Rels file not found for $($usage.PartPath): $relsPath"
            continue
        }
        
        $relsDoc = Get-XmlDocument $relsPath
        if (-not $relsDoc) {
            Write-Warning "Failed to load rels file: $relsPath"
            continue
        }
        
        # Remove old relationship
        $oldRId = $usage.BlipRId
        if ($oldRId) {
            $oldRel = $relsDoc.SelectSingleNode("//x:Relationship[@Id='$oldRId']", $script:NsMgr)
            if ($oldRel) {
                $null = $oldRel.ParentNode.RemoveChild($oldRel)
                Write-Verbose "    Removed old relationship $oldRId from $($usage.PartPath)"
            }
        }
        
        # Create new relationship
        $newRId = Get-UniqueRId -relsDoc $relsDoc
        $rel = $relsDoc.CreateElement('Relationship', $relsDoc.DocumentElement.NamespaceURI)
        [void]$rel.SetAttribute('Id', $newRId)
        [void]$rel.SetAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image')
        [void]$rel.SetAttribute('Target', "../media/$outputName")
        $null = $relsDoc.DocumentElement.AppendChild($rel)
        
        Write-Verbose "    Added new relationship $newRId → ../media/$outputName"
        
        Save-XmlDocument -doc $relsDoc -path $relsPath
        
        # Update blip
        [void]$usage.BlipElement.SetAttribute('embed', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships', $newRId)
        
        # Update BlipRId for consistency
        $usage.BlipRId = $newRId
        
        $usage.AfterSizeBytes = $afterSize
        $usage.OptimizationStatus = 'ConvertedPngToJpeg'
        $usage.OptimizedFile = $outputName
    }
    
    # Remove original PNG
    Remove-Item -LiteralPath $group.PhysicalPath -Force
    $group.OptimizedSizeBytes = $afterSize
    $group.OptimizedPath = $outputPath
    
    $savedPercent = (($beforeSize - $afterSize) / $beforeSize) * 100
    $firstUsage = $group.Usages[0]
    Write-Host "🔁 Converted PNG to JPEG for '$($firstUsage.ShapeName)' on $($firstUsage.Location) (saved $($savedPercent.ToString('F1'))%)" -ForegroundColor Green
    
    return @{ Success = $true }
}

#endregion

#region Cleanup and Packaging

function Invoke-MediaCleanup {
    param([string]$tempDir)
    
    Write-Host "`n🧹 Cleaning up unreferenced media..." -ForegroundColor Yellow
    
    # Collect all referenced media paths
    $referenced = [System.Collections.Generic.HashSet[string]]::new()
    
    $relsFiles = Get-ChildItem -Path $tempDir -Filter "*.rels" -Recurse
    Write-Verbose "Scanning $($relsFiles.Count) relationship files for image references..."
    
    foreach ($relsFile in $relsFiles) {
        $relsDoc = Get-XmlDocument $relsFile.FullName
        
        # Check both regular images and HD Photo layers
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
    
    # Find and remove unreferenced files
    $mediaDir = Join-Path $tempDir 'ppt\media'
    if (Test-Path -LiteralPath $mediaDir) {
        $allMedia = Get-ChildItem -Path $mediaDir
        Write-Verbose "Checking $($allMedia.Count) media file(s) for references..."
        
        $deletedCount = 0
        
        foreach ($file in $allMedia) {
            if (-not $referenced.Contains($file.Name)) {
                Remove-Item $file.FullName -Force
                $deletedCount++
                Write-Verbose "  Deleted unreferenced: $($file.Name)"
            }
        }
        
        Write-Host "🧹 Media files removed: $deletedCount" -ForegroundColor Green
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
            Write-Verbose "  Added content type: .$ext → $($requiredDefaults[$ext])"
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
        [string]$tempDir,
        [string]$outputPath
    )
    
    Write-Host "`nRepacking presentation..." -ForegroundColor Yellow
    Write-Verbose "Creating optimized ZIP archive from: $tempDir"
    Write-Verbose "Output file: $outputPath"
    
    if (Test-Path -LiteralPath $outputPath) {
        Write-Verbose "Removing existing output file..."
        Remove-Item -LiteralPath $outputPath -Force
    }
    
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
            $outputPath, 
            [System.IO.Compression.CompressionLevel]::Optimal, 
            $false  # Don't include base directory
        )
        
        # Verify the output file was created and is valid
        if (-not (Test-Path -LiteralPath $outputPath)) {
            throw "Output file was not created: $outputPath"
        }
        
        $outputSize = (Get-Item -LiteralPath $outputPath).Length
        Write-Verbose "Repack complete: output file is $(Format-ByteSize $outputSize)"
        
        $fileInfo = Get-Item -LiteralPath $outputPath
        if ($fileInfo.Length -eq 0) {
            throw "Output file is empty: $outputPath"
        }
        
        Write-Host "✅ Created: $outputPath ($([Math]::Round($fileInfo.Length / 1KB, 2)) KB)" -ForegroundColor Green
        
    } catch {
        throw "Failed to create ZIP archive: $_"
    }
}

#endregion

#region CSV Reporting

function Export-CsvReport {
    param(
        [System.Collections.Generic.List[ImageUsage]]$usages,
        [string]$csvPath
    )
    
    Write-Host "`nGenerating CSV report..." -ForegroundColor Yellow
    
    if (-not $usages -or $usages.Count -eq 0) {
        Write-Host "No usages to report" -ForegroundColor Gray
        @() | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
        return
    }
    
    # Sort by savings (descending), then context, then slide number
    $sorted = $usages | Sort-Object @(
        @{ Expression = { $_.BeforeSizeBytes - $_.AfterSizeBytes }; Descending = $true }
        @{ Expression = { $_.ContextType }; Ascending = $true }
        @{ Expression = { $_.SlideNumber }; Ascending = $true }
    )
    
    $rows = foreach ($usage in $sorted) {
        [PSCustomObject]@{
            Location = $usage.Location
            ContextType = $usage.ContextType
            SlideNumber = $usage.SlideNumber
            ShapeName = $usage.ShapeName
            SourceWidthPx = $usage.SourceWidthPx
            SourceHeightPx = $usage.SourceHeightPx
            DisplayWidthPx = $usage.DisplayWidthPx
            DisplayHeightPx = $usage.DisplayHeightPx
            TargetWidthPx = $usage.TargetWidthPx
            TargetHeightPx = $usage.TargetHeightPx
            OriginalFile = $usage.OriginalFileName
            OptimizedFile = if ($usage.OptimizedFile) { $usage.OptimizedFile } else { $usage.OriginalFileName }
            BeforeSizeBytes = $usage.BeforeSizeBytes
            AfterSizeBytes = if ($usage.AfterSizeBytes -gt 0) { $usage.AfterSizeBytes } else { $usage.BeforeSizeBytes }
            Savings = Format-Savings -before $usage.BeforeSizeBytes -after $(if ($usage.AfterSizeBytes -gt 0) { $usage.AfterSizeBytes } else { $usage.BeforeSizeBytes })
            HasSrcRect = $usage.HasSrcRect
            CropApplied = if ($usage.CropApplied) { 'Yes' } else { 'No' }
            CropNormalized = if ($usage.CropNormalized) { 'Yes' } else { 'No' }
            CropRemovedNoOp = if ($usage.CropRemovedNoOp) { 'Yes' } else { 'No' }
            SvgFallback = if ($usage.IsSvgFallbackUsage) { 'Yes' } else { 'No' }
            OptimizationStatus = $usage.OptimizationStatus
            WhyNotOptimized = $usage.WhyNotOptimized
            HeadroomFactor = $HeadroomFactor
            JpegQuality = $JpegQuality
            TransparencyThresholdPercent = $TransparencyThresholdPercent
            EffectiveTransparencyPercent = $usage.EffectiveTransparencyPercent.ToString('F2')
            ManualActionRequired = if ($usage.ManualActionRequired) { 'Yes' } else { 'No' }
            ManualActionHint = $usage.ManualActionHint
        }
    }
    
    $rows | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
    Write-Host "📊 CSV report: $csvPath" -ForegroundColor Green
}

#endregion

#region Main Processing

try {
    Write-Host "🔎 Analyzing presentation structure..." -ForegroundColor Cyan
    
    # Initialize
    Initialize-Script -InputPath $InputPath -MagickPath $MagickPath
    Initialize-Namespaces
    
    # Create temp directory
    $tempDir = Join-Path $env:TEMP "ppt-optimize-$(New-Guid)"
    New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
    
    # Normalize to full path (avoid 8.3 filename issues)
    $tempDir = (Get-Item -LiteralPath $tempDir).FullName
    
    try {
        # Extract
        Expand-Archive -Path $script:InputFile -DestinationPath $tempDir -Force
        
        # Get canonical slide order
        $slides = Get-CanonicalSlideOrder -tempDir $tempDir
        
        if (-not $slides -or $slides.Count -eq 0) {
            Write-Warning "No slides found in presentation"
            $slides = [System.Collections.Generic.List[SlideInfo]]::new()
        }
        
        # Scan for image usages
        $usages = Get-ImageUsages -tempDir $tempDir -slides $slides
        
        if (-not $usages) {
            $usages = [System.Collections.Generic.List[ImageUsage]]::new()
        }
        
        Write-Host "📊 Found $($usages.Count) image usages" -ForegroundColor Cyan
        $uniqueImages = if ($usages.Count -gt 0) { 
            ($usages | Select-Object -Property ImagePhysicalPath -Unique).Count 
        } else { 
            0 
        }
        Write-Host "📊 Across $uniqueImages unique images" -ForegroundColor Cyan
        
        if ($usages.Count -eq 0) {
            Write-Host "`n📝 No images found to optimize" -ForegroundColor Yellow
            
            # Still create output file and empty CSV
            Copy-Item $script:InputFile $script:OutputFile
            
            $emptyReport = @()
            $emptyReport | Export-Csv -Path $script:CsvReport -NoTypeInformation -Encoding UTF8
            
            Write-Host "`n🏁 Complete (no changes needed)" -ForegroundColor Green
            Write-Host "Output: $script:OutputFile" -ForegroundColor Green
            Write-Host "Report: $script:CsvReport" -ForegroundColor Green
            return
        }
        
        # Find Morph pairs
        $morphPairs = Find-MorphPairs -usages $usages
        
        # Phase 1: Normalize crops (only if cropping enabled)
        if ($EnableCropping) {
            $docsToSave = @{}
            
            foreach ($usage in $usages) {
                Invoke-CropNormalization -usage $usage
                if ($usage.CropNormalized -and $usage.SlideDocument) {
                    $slidePath = Join-Path $tempDir ($usage.PartPath -replace '/', '\')
                    $docsToSave[$slidePath] = $usage.SlideDocument
                }
            }
            
            # Save normalized XMLs
            foreach ($path in $docsToSave.Keys) {
                Save-XmlDocument -doc $docsToSave[$path] -path $path
            }
        }
        
        # Phase 2: Cropping
        if ($EnableCropping) {
            Invoke-ImageCropping -usages $usages -tempDir $tempDir -morphPairs $morphPairs
            
            # Save XML changes
            $docsToSave = @{}
            foreach ($usage in $usages | Where-Object { $_.CropApplied -or $_.CropRemovedNoOp }) {
                if ($usage.SlideDocument) {
                    $slidePath = Join-Path $tempDir ($usage.PartPath -replace '/', '\')
                    $docsToSave[$slidePath] = $usage.SlideDocument
                }
            }
            foreach ($path in $docsToSave.Keys) {
                Save-XmlDocument -doc $docsToSave[$path] -path $path
            }
        }
        
        # Phase 3: Optimization
        Invoke-ImageOptimization -usages $usages -tempDir $tempDir -morphPairs $morphPairs
        
        # Save XML changes (format conversions update blip refs)
        $docsToSave = @{}
        foreach ($usage in $usages | Where-Object { $_.OptimizationStatus -eq 'ConvertedPngToJpeg' }) {
            if ($usage.SlideDocument) {
                $slidePath = Join-Path $tempDir ($usage.PartPath -replace '/', '\')
                $docsToSave[$slidePath] = $usage.SlideDocument
                $relativePath = $usage.PartPath
                Write-Verbose "Marking document for save (PNG→JPEG conversion): $relativePath"
            } else {
                Write-Warning "Converted usage on $($usage.Location) missing SlideDocument reference"
            }
        }
        
        if ($docsToSave.Count -gt 0) {
            Write-Host "💾 Saving $($docsToSave.Count) updated document(s)..." -ForegroundColor Yellow
            foreach ($path in $docsToSave.Keys) {
                try {
                    $relativePath = $path -replace [regex]::Escape($tempDir + '\'), ''
                    Write-Verbose "Saving updated document: $relativePath"
                    Save-XmlDocument -doc $docsToSave[$path] -path $path
                } catch {
                    Write-Error "Failed to save $path : $_"
                    throw
                }
            }
        }
        
        # Clear XML references to reduce memory usage (no longer needed after saves)
        foreach ($usage in $usages) {
            $usage.SlideDocument = $null
            $usage.BlipElement = $null
            $usage.XfrmElement = $null
        }
        
        # Phase 4: Cleanup
        Invoke-MediaCleanup -tempDir $tempDir
        Update-ContentTypes -tempDir $tempDir
        
        # Validate critical files before repacking
        $criticalFiles = @(
            '[Content_Types].xml',
            '_rels\.rels',
            'ppt\presentation.xml',
            'ppt\_rels\presentation.xml.rels'
        )
        
        Write-Verbose "Validating $($criticalFiles.Count) critical files before repack..."
        foreach ($file in $criticalFiles) {
            $filePath = Join-Path $tempDir $file
            if (-not (Test-Path -LiteralPath $filePath)) {
                throw "Critical file missing before repack: $file"
            }
            Write-Verbose "  ✓ $file"
        }
        Write-Verbose "Pre-repack validation passed - all critical files present"
        
        # Phase 5: Repack
        Invoke-Repack -tempDir $tempDir -outputPath $script:OutputFile
        
        # Generate report
        Export-CsvReport -usages $usages -csvPath $script:CsvReport
        
        # Final statistics
        $originalSize = (Get-Item $script:InputFile).Length
        $optimizedSize = (Get-Item $script:OutputFile).Length
        $savedBytes = $originalSize - $optimizedSize
        $savedPercent = ($savedBytes / $originalSize) * 100
        
        $croppedCount = @($usages | Where-Object { $_.CropApplied }).Count
        $optimizedCount = @($usages | Where-Object { $_.OptimizationStatus -match '^(Optimized|Converted)' }).Count
        
        Write-Host "`n🏁 Optimization Complete" -ForegroundColor Green
        Write-Host "📊 Images cropped: $croppedCount" -ForegroundColor Green
        Write-Host "📊 Images optimized: $optimizedCount" -ForegroundColor Green
        Write-Host "`n📉 File size:" -ForegroundColor Green
        Write-Host "  Original: $((Format-ByteSize $originalSize))" -ForegroundColor Green
        Write-Host "  Optimized: $((Format-ByteSize $optimizedSize))" -ForegroundColor Green
        Write-Host "  Saved: $(Format-ByteSize $savedBytes) ($($savedPercent.ToString('F1'))%)" -ForegroundColor Green
        
    } finally {
        # Cleanup temp directory
        if (Test-Path -LiteralPath $tempDir) {
            Remove-Item -LiteralPath $tempDir -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
    
} catch {
    Write-Error "Fatal error: $_"
    Write-Error $_.ScriptStackTrace
    exit 1
}

#endregion