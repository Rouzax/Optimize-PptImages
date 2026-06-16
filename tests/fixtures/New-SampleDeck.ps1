<#
.SYNOPSIS
    Generates a minimal but valid .pptx for characterization tests.

.DESCRIPTION
    Produces a deck with synthetic content only (no real text, solid colors and
    a procedural image). The deck contains:
      - One slide that displays a large raster PNG (2000x2000) inside a small
        (~2 inch) box, so the optimizer has oversized images to resize.
      - Two slide layouts where only one is used by the slide, so
        -CleanUnusedLayouts has an unused layout to remove.

    The PNG is created with ImageMagick (`magick`). The package is assembled with
    the DocumentFormat.OpenXml SDK, loaded with the two-DLL pattern (Framework
    first, then the main assembly, picking the netstandard builds).

.PARAMETER OutputPath
    Destination .pptx path. Created or overwritten.

.PARAMETER MagickPath
    Optional path to the ImageMagick executable. Defaults to 'magick' on PATH.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$OutputPath,

    [Parameter(Mandatory = $false)]
    [string]$MagickPath = 'magick'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# --- Load the Open XML SDK (Framework DLL first, then the main assembly) ---
$fwPkg = Get-Package -Name DocumentFormat.OpenXml.Framework -ErrorAction Stop
$mainPkg = Get-Package -Name DocumentFormat.OpenXml -ErrorAction Stop
$fwDll = Get-ChildItem (Split-Path $fwPkg.Source -Parent) -Recurse -Filter '*.dll' |
    Where-Object { $_.FullName -match 'netstandard' } | Select-Object -First 1
$mainDll = Get-ChildItem (Split-Path $mainPkg.Source -Parent) -Recurse -Filter '*.dll' |
    Where-Object { $_.FullName -match 'netstandard' } | Select-Object -First 1
if (-not $fwDll -or -not $mainDll) {
    throw "Could not locate the DocumentFormat.OpenXml netstandard DLLs."
}
Add-Type -Path $fwDll.FullName
Add-Type -Path $mainDll.FullName

# --- Create two oversized PNGs with ImageMagick ---
# Two distinct images so the package always ends up with more than one media file.
$tempPng1 = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), ("popt-sample-a-" + [guid]::NewGuid().ToString('N') + ".png"))
$tempPng2 = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), ("popt-sample-b-" + [guid]::NewGuid().ToString('N') + ".png"))
& $MagickPath -size 2000x2000 'gradient:skyblue-navy' $tempPng1
if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $tempPng1)) {
    throw "ImageMagick failed to create the first sample PNG (exit $LASTEXITCODE)."
}
& $MagickPath -size 2000x2000 'gradient:orange-firebrick' $tempPng2
if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $tempPng2)) {
    throw "ImageMagick failed to create the second sample PNG (exit $LASTEXITCODE)."
}

try {
    $outDir = Split-Path -Parent $OutputPath
    if ($outDir -and -not (Test-Path -LiteralPath $outDir)) {
        New-Item -ItemType Directory -Path $outDir -Force | Out-Null
    }
    if (Test-Path -LiteralPath $OutputPath) { Remove-Item -LiteralPath $OutputPath -Force }

    $doc = [DocumentFormat.OpenXml.Packaging.PresentationDocument]::Create(
        $OutputPath,
        [DocumentFormat.OpenXml.PresentationDocumentType]::Presentation)

    # ---- Presentation part ----
    $presPart = $doc.AddPresentationPart()
    $presPart.Presentation = [DocumentFormat.OpenXml.Presentation.Presentation]::new()

    # ---- Slide master + theme ----
    $masterPart = $presPart.AddNewPart[DocumentFormat.OpenXml.Packaging.SlideMasterPart]('rIdSm1')
    $themePart = $masterPart.AddNewPart[DocumentFormat.OpenXml.Packaging.ThemePart]('rIdTheme1')

    # Minimal but valid theme. Built from raw XML to avoid the many strongly-typed
    # Append calls; the SDK parses this into the Theme part root element.
    $themeXml = @'
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="1F497D"/></a:dk2>
      <a:lt2><a:srgbClr val="EEECE1"/></a:lt2>
      <a:accent1><a:srgbClr val="4F81BD"/></a:accent1>
      <a:accent2><a:srgbClr val="C0504D"/></a:accent2>
      <a:accent3><a:srgbClr val="9BBB59"/></a:accent3>
      <a:accent4><a:srgbClr val="8064A2"/></a:accent4>
      <a:accent5><a:srgbClr val="4BACC6"/></a:accent5>
      <a:accent6><a:srgbClr val="F79646"/></a:accent6>
      <a:hlink><a:srgbClr val="0000FF"/></a:hlink>
      <a:folHlink><a:srgbClr val="800080"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
      <a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
        <a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
        <a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
      </a:lnStyleLst>
      <a:effectStyleLst>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
      </a:effectStyleLst>
      <a:bgFillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
</a:theme>
'@
    $theme = [DocumentFormat.OpenXml.Drawing.Theme]::new($themeXml)
    $themePart.Theme = $theme

    # ---- Build a reusable common-slide-data factory ----
    function New-CommonSlideData {
        $csd = [DocumentFormat.OpenXml.Presentation.CommonSlideData]::new()
        $spTree = [DocumentFormat.OpenXml.Presentation.ShapeTree]::new()
        $nvGrp = [DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeProperties]::new()
        $cnvP = [DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties]::new()
        $cnvP.Id = [System.UInt32]1; $cnvP.Name = ''
        [void]$nvGrp.Append($cnvP)
        [void]$nvGrp.Append([DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeDrawingProperties]::new())
        [void]$nvGrp.Append([DocumentFormat.OpenXml.Presentation.ApplicationNonVisualDrawingProperties]::new())
        [void]$spTree.Append($nvGrp)
        $grpSpPr = [DocumentFormat.OpenXml.Presentation.GroupShapeProperties]::new()
        [void]$spTree.Append($grpSpPr)
        [void]$csd.Append($spTree)
        return @{ Csd = $csd; SpTree = $spTree }
    }

    # ---- Slide master content ----
    $smCsd = New-CommonSlideData
    $slideMaster = [DocumentFormat.OpenXml.Presentation.SlideMaster]::new()
    $slideMaster.Append($smCsd.Csd)
    $clrMap = [DocumentFormat.OpenXml.Presentation.ColorMap]::new()
    $clrMap.Background1 = [DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues]::Light1
    $clrMap.Text1 = [DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues]::Dark1
    $clrMap.Background2 = [DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues]::Light2
    $clrMap.Text2 = [DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues]::Dark2
    $clrMap.Accent1 = [DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues]::Accent1
    $clrMap.Accent2 = [DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues]::Accent2
    $clrMap.Accent3 = [DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues]::Accent3
    $clrMap.Accent4 = [DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues]::Accent4
    $clrMap.Accent5 = [DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues]::Accent5
    $clrMap.Accent6 = [DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues]::Accent6
    $clrMap.Hyperlink = [DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues]::Hyperlink
    $clrMap.FollowedHyperlink = [DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues]::FollowedHyperlink
    $slideMaster.Append($clrMap)
    $sldLayoutIdList = [DocumentFormat.OpenXml.Presentation.SlideLayoutIdList]::new()

    # ---- Two layouts: layout1 (used) and layout2 (unused) ----
    function New-LayoutPart {
        param($MasterPart, $RelId)
        $layoutPart = $MasterPart.AddNewPart[DocumentFormat.OpenXml.Packaging.SlideLayoutPart]($RelId)
        $lCsd = New-CommonSlideData
        $layout = [DocumentFormat.OpenXml.Presentation.SlideLayout]::new()
        $layout.Type = [DocumentFormat.OpenXml.Presentation.SlideLayoutValues]::Blank
        [void]$layout.Append($lCsd.Csd)
        $lClrMap = [DocumentFormat.OpenXml.Presentation.ColorMapOverride]::new()
        [void]$lClrMap.Append([DocumentFormat.OpenXml.Drawing.MasterColorMapping]::new())
        [void]$layout.Append($lClrMap)
        $layoutPart.SlideLayout = $layout
        return $layoutPart
    }

    # layout1 is referenced by the slide; layout2 is left unused so -CleanUnusedLayouts has work.
    $layout1Part = New-LayoutPart -MasterPart $masterPart -RelId 'rIdLayout1'
    $layout2Part = New-LayoutPart -MasterPart $masterPart -RelId 'rIdLayout2'

    $layout1Id = [DocumentFormat.OpenXml.Presentation.SlideLayoutId]::new()
    $layout1Id.Id = [System.UInt32]2147483649
    $layout1Id.RelationshipId = $masterPart.GetIdOfPart($layout1Part)
    $layout2Id = [DocumentFormat.OpenXml.Presentation.SlideLayoutId]::new()
    $layout2Id.Id = [System.UInt32]2147483650
    $layout2Id.RelationshipId = $masterPart.GetIdOfPart($layout2Part)
    $sldLayoutIdList.Append($layout1Id)
    $sldLayoutIdList.Append($layout2Id)
    $slideMaster.Append($sldLayoutIdList)
    $masterPart.SlideMaster = $slideMaster

    # ---- Slide that uses layout1 and displays the oversized images ----
    $slidePart = $presPart.AddNewPart[DocumentFormat.OpenXml.Packaging.SlidePart]('rIdSlide1')
    # Link the slide to layout1 (the used layout). layout2 stays unreferenced by any slide.
    $slidePart.AddPart($layout1Part) | Out-Null

    function Add-ImagePartFromFile {
        param($SlidePart, $RelId, $PngPath)
        $part = $SlidePart.AddNewPart[DocumentFormat.OpenXml.Packaging.ImagePart]('image/png', $RelId)
        $stream = [System.IO.File]::OpenRead($PngPath)
        try { $part.FeedData($stream) } finally { $stream.Dispose() }
        return $SlidePart.GetIdOfPart($part)
    }

    $imgRelId1 = Add-ImagePartFromFile -SlidePart $slidePart -RelId 'rIdImg1' -PngPath $tempPng1
    $imgRelId2 = Add-ImagePartFromFile -SlidePart $slidePart -RelId 'rIdImg2' -PngPath $tempPng2

    $slideCsd = New-CommonSlideData
    $slide = [DocumentFormat.OpenXml.Presentation.Slide]::new()
    [void]$slide.Append($slideCsd.Csd)
    $slideClrMap = [DocumentFormat.OpenXml.Presentation.ColorMapOverride]::new()
    [void]$slideClrMap.Append([DocumentFormat.OpenXml.Drawing.MasterColorMapping]::new())
    [void]$slide.Append($slideClrMap)

    function New-PictureShape {
        param($ShapeId, $ShapeName, $EmbedRelId, $OffsetXEmu)
        # 2000x2000 PNG shown in a ~2 inch box (1828800 EMU).
        $pic = [DocumentFormat.OpenXml.Presentation.Picture]::new()
        $nvPicPr = [DocumentFormat.OpenXml.Presentation.NonVisualPictureProperties]::new()
        $picCnvP = [DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties]::new()
        $picCnvP.Id = [System.UInt32]$ShapeId; $picCnvP.Name = $ShapeName
        [void]$nvPicPr.Append($picCnvP)
        [void]$nvPicPr.Append([DocumentFormat.OpenXml.Presentation.NonVisualPictureDrawingProperties]::new())
        [void]$nvPicPr.Append([DocumentFormat.OpenXml.Presentation.ApplicationNonVisualDrawingProperties]::new())
        [void]$pic.Append($nvPicPr)

        $blipFill = [DocumentFormat.OpenXml.Presentation.BlipFill]::new()
        $blip = [DocumentFormat.OpenXml.Drawing.Blip]::new()
        $blip.Embed = $EmbedRelId
        [void]$blipFill.Append($blip)
        $stretch = [DocumentFormat.OpenXml.Drawing.Stretch]::new()
        [void]$stretch.Append([DocumentFormat.OpenXml.Drawing.FillRectangle]::new())
        [void]$blipFill.Append($stretch)
        [void]$pic.Append($blipFill)

        $spPr = [DocumentFormat.OpenXml.Presentation.ShapeProperties]::new()
        $xfrm = [DocumentFormat.OpenXml.Drawing.Transform2D]::new()
        $off = [DocumentFormat.OpenXml.Drawing.Offset]::new(); $off.X = [System.Int64]$OffsetXEmu; $off.Y = [System.Int64]914400
        $ext = [DocumentFormat.OpenXml.Drawing.Extents]::new(); $ext.Cx = [System.Int64]1828800; $ext.Cy = [System.Int64]1828800
        [void]$xfrm.Append($off); [void]$xfrm.Append($ext)
        [void]$spPr.Append($xfrm)
        $presetGeom = [DocumentFormat.OpenXml.Drawing.PresetGeometry]::new()
        $presetGeom.Preset = [DocumentFormat.OpenXml.Drawing.ShapeTypeValues]::Rectangle
        [void]$presetGeom.Append([DocumentFormat.OpenXml.Drawing.AdjustValueList]::new())
        [void]$spPr.Append($presetGeom)
        [void]$pic.Append($spPr)
        # Unary comma prevents PowerShell from enumerating the Picture (an
        # OpenXmlElement is IEnumerable over its children) when returning it.
        return , $pic
    }

    [void]$slideCsd.SpTree.Append((New-PictureShape -ShapeId 2 -ShapeName 'Picture 1' -EmbedRelId $imgRelId1 -OffsetXEmu 914400))
    [void]$slideCsd.SpTree.Append((New-PictureShape -ShapeId 3 -ShapeName 'Picture 2' -EmbedRelId $imgRelId2 -OffsetXEmu 3657600))
    $slidePart.Slide = $slide

    # ---- Slide id list and slide size ----
    $sldIdList = [DocumentFormat.OpenXml.Presentation.SlideIdList]::new()
    $sldId = [DocumentFormat.OpenXml.Presentation.SlideId]::new()
    $sldId.Id = [System.UInt32]256
    $sldId.RelationshipId = $presPart.GetIdOfPart($slidePart)
    $sldIdList.Append($sldId)

    $sldMasterIdList = [DocumentFormat.OpenXml.Presentation.SlideMasterIdList]::new()
    $sldMasterId = [DocumentFormat.OpenXml.Presentation.SlideMasterId]::new()
    $sldMasterId.Id = [System.UInt32]2147483648
    $sldMasterId.RelationshipId = $presPart.GetIdOfPart($masterPart)
    $sldMasterIdList.Append($sldMasterId)

    $presPart.Presentation.Append($sldMasterIdList)
    $presPart.Presentation.Append($sldIdList)
    $sldSz = [DocumentFormat.OpenXml.Presentation.SlideSize]::new()
    $sldSz.Cx = [System.Int32]9144000; $sldSz.Cy = [System.Int32]6858000
    $presPart.Presentation.Append($sldSz)
    $notesSz = [DocumentFormat.OpenXml.Presentation.NotesSize]::new()
    $notesSz.Cx = [System.Int64]6858000; $notesSz.Cy = [System.Int64]9144000
    $presPart.Presentation.Append($notesSz)

    $presPart.Presentation.Save()
    $masterPart.SlideMaster.Save()
    $layout1Part.SlideLayout.Save()
    $layout2Part.SlideLayout.Save()
    $slidePart.Slide.Save()
    $themePart.Theme.Save()

    $doc.Dispose()
}
finally {
    if (Test-Path -LiteralPath $tempPng1) { Remove-Item -LiteralPath $tempPng1 -Force -ErrorAction SilentlyContinue }
    if (Test-Path -LiteralPath $tempPng2) { Remove-Item -LiteralPath $tempPng2 -Force -ErrorAction SilentlyContinue }
}

if (-not (Test-Path -LiteralPath $OutputPath)) {
    throw "Sample deck was not created at: $OutputPath"
}

# --- Normalize relationship targets to the PowerPoint convention ---
# The SDK writes absolute package targets (Target="/ppt/slides/slide1.xml"). Real
# PowerPoint decks (and this module's scanner, which does "ppt/$target") expect
# part-relative targets (Target="slides/slide1.xml", "../media/image.png"). Unpack,
# rewrite every .rels Target from absolute to relative, then repack.
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.IO.Compression

$workDir = Join-Path ([System.IO.Path]::GetTempPath()) ("popt-sample-norm-" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $workDir -Force | Out-Null
try {
    [System.IO.Compression.ZipFile]::ExtractToDirectory($OutputPath, $workDir)

    function ConvertTo-RelativeTarget {
        param([string]$OwningPartPath, [string]$AbsoluteTarget)
        # Both are package-absolute paths using forward slashes; $AbsoluteTarget starts with '/'.
        $targetSegs = @($AbsoluteTarget.TrimStart('/').Split('/'))
        $ownerDirSegs = (Split-Path $OwningPartPath -Parent) -replace '\\', '/'
        $ownerSegs = @(if ($ownerDirSegs) { $ownerDirSegs.Split('/') } else { @() })
        # Strip common prefix.
        $i = 0
        while ($i -lt $ownerSegs.Count -and $i -lt ($targetSegs.Count - 1) -and $ownerSegs[$i] -eq $targetSegs[$i]) { $i++ }
        $up = $ownerSegs.Count - $i
        $rel = @()
        for ($u = 0; $u -lt $up; $u++) { $rel += '..' }
        for ($t = $i; $t -lt $targetSegs.Count; $t++) { $rel += $targetSegs[$t] }
        return ($rel -join '/')
    }

    $relsFiles = Get-ChildItem -Path $workDir -Filter '*.rels' -Recurse -File
    foreach ($relsFile in $relsFiles) {
        # Package path of the part that owns this .rels file.
        $relsRel = $relsFile.FullName.Substring($workDir.Length).TrimStart('\', '/') -replace '\\', '/'
        # e.g. ppt/slides/_rels/slide1.xml.rels -> owning part ppt/slides/slide1.xml
        $owningPart = ($relsRel -replace '/_rels/', '/') -replace '\.rels$', ''
        [xml]$xml = Get-Content -LiteralPath $relsFile.FullName -Raw
        $changed = $false
        foreach ($rel in $xml.Relationships.Relationship) {
            $isExternal = ($rel.PSObject.Properties.Name -contains 'TargetMode') -and ($rel.TargetMode -eq 'External')
            if ($isExternal) { continue }
            if ($rel.Target -and $rel.Target.StartsWith('/')) {
                $rel.Target = ConvertTo-RelativeTarget -OwningPartPath $owningPart -AbsoluteTarget $rel.Target
                $changed = $true
            }
        }
        if ($changed) { $xml.Save($relsFile.FullName) }
    }

    Remove-Item -LiteralPath $OutputPath -Force
    [System.IO.Compression.ZipFile]::CreateFromDirectory($workDir, $OutputPath)
}
finally {
    if (Test-Path -LiteralPath $workDir) { Remove-Item -LiteralPath $workDir -Recurse -Force -ErrorAction SilentlyContinue }
}

return $OutputPath
