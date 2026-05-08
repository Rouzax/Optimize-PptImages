<#
.SYNOPSIS
    Optimizes PowerPoint presentations by cropping images and reducing file sizes.

.DESCRIPTION
    Creates an optimized copy of a PowerPoint file (.pptx/.potx) plus a CSV report
    that explains every decision, while preserving visual fidelity (rectangular crops,
    masters/layouts, Morph transitions).

    Supports batch processing via pipeline input.

.PARAMETER InputPath
    Path to the input PowerPoint file (.pptx or .potx). Accepts pipeline input.

.PARAMETER OutputPath
    Path for the optimized output file. Default: <n>.optimized<ext>

.PARAMETER MagickPath
    Path to ImageMagick 'magick' executable. If not specified, searches PATH.

.PARAMETER OptimizeSlides
    Allow resizing and format conversion of images in slides.

.PARAMETER OptimizeMastersAndLayouts
    Allow resizing and format conversion of images in masters/layouts.

.PARAMETER CropSlides
    Allow materializing rectangular a:srcRect crops into physical images in slides.

.PARAMETER CropMastersAndLayouts
    Allow materializing rectangular a:srcRect crops into physical images in masters/layouts.

.PARAMETER CleanUnusedLayouts
    Remove slide layouts not used by any slide and slide masters that become
    unreferenced. Not included in -All. Reduces the layout gallery in PowerPoint.

.PARAMETER All
    Enable all optimization and cropping operations.

.PARAMETER HeadroomFactor
    Target resolution multiplier (display x factor). Default: 2.0. Range: 0.5-4.0

.PARAMETER JpegQuality
    JPEG quality for optimization. Default: 95. Range: 1-100

.PARAMETER TransparencyThresholdPercent
    Threshold for treating images as transparent. Default: 0.5. Range: 0.0-100.0

.PARAMETER MinSavingsPercent
    Minimum savings percentage required to apply optimization. Default: 0.0. Range: 0.0-50.0

.PARAMETER IncludeSlides
    Array of slide numbers to process. If specified, only these slides are modified.
    Note: All slides are still analyzed for correct image sizing and morph detection.
    Shared images are optimized using max dimensions from all usages (including excluded slides).

.PARAMETER ExcludeSlides
    Array of slide numbers to exclude from processing.
    Note: All slides are still analyzed for correct image sizing and morph detection.

.PARAMETER CsvReportPath
    Path for the CSV report. Default: alongside OutputPath as <OutputBaseName>.opt-report.csv

.PARAMETER PassThru
    Return the result object to the pipeline. By default, results are not emitted
    to avoid noise during interactive use. Use -PassThru for programmatic access.

.OUTPUTS
    PSCustomObject with optimization results including file sizes, savings, and counts.

.EXAMPLE
    .\Optimize-PptImages.ps1 -InputPath "presentation.pptx" -OptimizeSlides

.EXAMPLE
    .\Optimize-PptImages.ps1 -InputPath "presentation.pptx" -OptimizeSlides -CropSlides

.EXAMPLE
    .\Optimize-PptImages.ps1 -InputPath "template.potx" -All -HeadroomFactor 1.5

.EXAMPLE
    Get-ChildItem *.pptx | .\Optimize-PptImages.ps1 -OptimizeSlides -CropSlides
    # Batch process all PPTX files in current directory

.EXAMPLE
    .\Optimize-PptImages.ps1 -InputPath "presentation.pptx" -IncludeSlides 1,5,10 -OptimizeSlides -CropSlides
    # Only process slides 1, 5, and 10

.LINK
    https://github.com/yourusername/Optimize-PptImages

#>

[CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='Medium')]
param(
    [Parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
    [Alias('FullName', 'Path')]
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
    [switch]$OptimizeSlides,

    [Parameter(Mandatory=$false)]
    [switch]$OptimizeMastersAndLayouts,

    [Parameter(Mandatory=$false)]
    [switch]$CropSlides,

    [Parameter(Mandatory=$false)]
    [switch]$CropMastersAndLayouts,

    [Parameter(Mandatory=$false)]
    [switch]$CleanUnusedLayouts,

    [Parameter(Mandatory=$false)]
    [switch]$All,

    [Parameter(Mandatory=$false)]
    [ValidateRange(0.5, 4.0)]
    [double]$HeadroomFactor = 2.0,

    [Parameter(Mandatory=$false)]
    [ValidateRange(1, 100)]
    [int]$JpegQuality = 95,

    [Parameter(Mandatory=$false)]
    [ValidateRange(0.0, 100.0)]
    [double]$TransparencyThresholdPercent = 0.1,

    [Parameter(Mandatory=$false)]
    [ValidateRange(0.0, 50.0)]
    [double]$MinSavingsPercent = 0.0,

    [Parameter(Mandatory=$false)]
    [int[]]$IncludeSlides,

    [Parameter(Mandatory=$false)]
    [int[]]$ExcludeSlides,

    [Parameter(Mandatory=$false)]
    [string]$CsvReportPath,

    [Parameter(Mandatory=$false)]
    [switch]$PassThru
)

#Requires -Version 5.1

begin {
    Set-StrictMode -Version Latest
    $ErrorActionPreference = 'Stop'
    
    # Track batch processing results
    $script:BatchResults = [System.Collections.Generic.List[PSCustomObject]]::new()
    $script:BatchMode = $false
    $script:FileCount = 0

    # Expand -All flag into individual flags
    if ($All) {
        $script:OptimizeSlides = $true
        $script:OptimizeMastersAndLayouts = $true
        $script:CropSlides = $true
        $script:CropMastersAndLayouts = $true
    } else {
        $script:OptimizeSlides = $OptimizeSlides.IsPresent
        $script:OptimizeMastersAndLayouts = $OptimizeMastersAndLayouts.IsPresent
        $script:CropSlides = $CropSlides.IsPresent
        $script:CropMastersAndLayouts = $CropMastersAndLayouts.IsPresent
    }
    $script:CleanUnusedLayouts = $CleanUnusedLayouts.IsPresent

    #region Enums

    enum ContextType {
        Slide
        Layout
        Master
        Notes
        Handout
    }

    #endregion

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
        SUPPORTED_RASTER_FORMATS = @('.png', '.jpg', '.jpeg')
        CONVERTIBLE_FORMATS = @('.bmp', '.tif', '.tiff', '.gif')
    }

    #endregion

    #region Helper Classes and Structures

    class ImageUsage {
        [string]$Location
        [ContextType]$ContextType
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
        [int]$SourceJpegQuality  # For JPEG quality assessment
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

    #region Verbose Helper

    function Test-VerboseMode {
        return $VerbosePreference -eq 'Continue'
    }

    #endregion

    #region ImageMagick Helpers

    function Get-ImageDimensions {
        param([string]$ImagePath)
        
        $identify = & $script:MagickExe identify -format "%w %h" $ImagePath 2>&1
        if ($LASTEXITCODE -ne 0) {
            throw "ImageMagick identify failed for '$ImagePath': $identify"
        }
        
        if ($identify -is [array]) {
            $identify = $identify[0]
        }
        
        $dims = $identify -split '\s+'
        return @{
            Width = [int]$dims[0]
            Height = [int]$dims[1]
        }
    }

    function Get-JpegQuality {
        param([string]$ImagePath)
        
        $quality = & $script:MagickExe identify -format "%Q" $ImagePath 2>&1
        if ($LASTEXITCODE -ne 0) {
            return -1  # Unknown
        }
        
        if ($quality -is [array]) {
            $quality = $quality[0]
        }
        
        if ($quality -match '^\d+$') {
            return [int]$quality
        }
        
        return -1
    }

    function Test-IsAnimatedGif {
        param([string]$ImagePath)
        
        $frameCount = & $script:MagickExe identify -format "%n\n" $ImagePath 2>&1 | 
            Where-Object { $_ -match '^\d+$' } | 
            Select-Object -First 1
        
        if ($frameCount -and [int]$frameCount -gt 1) {
            return $true
        }
        
        return $false
    }

    #endregion

    #region Relationship Helpers

    function Update-BlipRelationship {
        param(
            [ImageUsage]$usage,
            [string]$tempDir,
            [string]$newMediaFileName
        )

        $partDir = Split-Path $usage.PartPath -Parent
        $partFile = Split-Path $usage.PartPath -Leaf
        $relsPath = Join-Path $tempDir (($partDir + '/_rels/' + $partFile + '.rels') -replace '/', '\')

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

    #region Validation and Initialization

    function Initialize-Script {
        param(
            [string]$InputPath,
            [string]$MagickPath
        )

        # Returns: $null on success, or error message string on validation failure

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
        
        if (Test-VerboseMode) {
            Write-Verbose "Configuration:"
            Write-Verbose "  Input:  $script:InputFile"
            Write-Verbose "  Output: $script:OutputFile"
            Write-Verbose "  Report: $script:CsvReport"
            Write-Verbose "Operations enabled:"
            Write-Verbose "  OptimizeSlides: $script:OptimizeSlides"
            Write-Verbose "  OptimizeMastersAndLayouts: $script:OptimizeMastersAndLayouts"
            Write-Verbose "  CropSlides: $script:CropSlides"
            Write-Verbose "  CropMastersAndLayouts: $script:CropMastersAndLayouts"
            Write-Verbose "Parameters:"
            Write-Verbose "  HeadroomFactor: $HeadroomFactor"
            Write-Verbose "  JpegQuality: $JpegQuality"
            Write-Verbose "  TransparencyThreshold: $TransparencyThresholdPercent%"
            Write-Verbose "  MinSavingsPercent: $MinSavingsPercent%"
            if ($IncludeSlides) {
                Write-Verbose "  IncludeSlides: $($IncludeSlides -join ', ')"
            }
            if ($ExcludeSlides) {
                Write-Verbose "  ExcludeSlides: $($ExcludeSlides -join ', ')"
            }
        }
        
        # Warn about flags that are ignored when slide filters are active
        $slideFiltersActive = ($IncludeSlides -and $IncludeSlides.Count -gt 0) -or ($ExcludeSlides -and $ExcludeSlides.Count -gt 0)
        if ($slideFiltersActive) {
            $ignoredFlags = @()
            if ($script:CropMastersAndLayouts) { $ignoredFlags += '-CropMastersAndLayouts' }
            if ($script:OptimizeMastersAndLayouts) { $ignoredFlags += '-OptimizeMastersAndLayouts' }
            
            if ($ignoredFlags.Count -gt 0) {
                Write-Host "[WARN]  Note: $($ignoredFlags -join ', ') ignored when slide filters are active" -ForegroundColor DarkYellow
                Write-Host "   (Masters/layouts affect all slides, not just filtered selection)" -ForegroundColor DarkYellow
            }
        }
        
        # Check if output files are accessible (fail early)
        if (Test-FileInUse -Path $script:OutputFile) {
            return "Output file is in use: $($script:OutputFile)`n  -> Please close the file and try again."
        }
        if (Test-FileInUse -Path $script:CsvReport) {
            return "CSV report file is in use: $($script:CsvReport)`n  -> Please close the file (Excel?) and try again."
        }

        # Find ImageMagick
        $script:MagickExe = Find-ImageMagick -ProvidedPath $MagickPath
        if (-not $script:MagickExe) {
            return "ImageMagick not found.`n  -> Install ImageMagick and ensure 'magick' is in PATH, or use -MagickPath parameter."
        }

        Write-Verbose "ImageMagick: $script:MagickExe"
        
        # Return $null to indicate success
        return $null
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

    function Test-FileInUse {
        param([string]$Path)
        
        if (-not (Test-Path -LiteralPath $Path)) {
            return $false
        }
        
        try {
            $stream = [System.IO.File]::Open($Path, 'Open', 'Write', 'None')
            $stream.Close()
            $stream.Dispose()
            return $false
        } catch {
            return $true
        }
    }

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
            [System.Xml.XmlDocument]$relsDoc
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
        
        return "rId$($maxNum + 1)"
    }

    function Test-SlideIncluded {
        param([int]$SlideNumber)
        
        # If IncludeSlides specified, slide must be in the list
        if ($IncludeSlides -and $IncludeSlides.Count -gt 0) {
            if ($SlideNumber -notin $IncludeSlides) {
                return $false
            }
        }
        
        # If ExcludeSlides specified, slide must not be in the list
        if ($ExcludeSlides -and $ExcludeSlides.Count -gt 0) {
            if ($SlideNumber -in $ExcludeSlides) {
                return $false
            }
        }
        
        return $true
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

    function Get-SlideDimensions {
        <#
        .SYNOPSIS
            Gets the slide dimensions from presentation.xml.
        .DESCRIPTION
            Extracts the slide size (p:sldSz) from presentation.xml and converts to pixels.
            Used for background images that fill the entire slide.
        #>
        param([string]$tempDir)
        
        $presentationXml = Join-Path $tempDir 'ppt\presentation.xml'
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
        $layoutFullPath = Join-Path $tempDir ($layoutPath -replace '/', '\')
        
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
                $contextType = if ($partPath -match 'slideMasters') { [ContextType]::Master } else { [ContextType]::Layout }
                
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
        $mediaPath = Join-Path $tempDir "ppt\media\$mediaFile"
        
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

    #endregion

    #region Morph Detection

    function Find-MorphPairs {
        param([System.Collections.Generic.List[ImageUsage]]$usages)
        
        $pairs = @{}
        $morphSlideCount = 0
        $totalPairsFound = 0
        $unmatchedImages = 0
        
        $slideUsages = $usages | Where-Object { $_.ContextType -eq [ContextType]::Slide } | 
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
                $_.ContextType -eq [ContextType]::Slide -and $_.SlideNumber -eq $prevSlideNum 
            }
            
            if (Test-VerboseMode) {
                Write-Verbose "Analyzing Morph transition: Slide $prevSlideNum -> Slide $slideNum"
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
                    
                    if (Test-VerboseMode) {
                        $fileName = Split-Path $match.ImagePhysicalPath -Leaf
                        Write-Verbose "  Linked pair: $fileName across slides $prevSlideNum <-> $slideNum"
                    }
                } else {
                    $unmatchedImages++
                }
            }
            
            if (Test-VerboseMode -and $slideMatchCount -eq 0) {
                Write-Verbose "  Warning: No matching images found between slides $prevSlideNum and $slideNum"
            }
        }
        
        if ($morphSlideCount -gt 0) {
            Write-Host "[LINK] Morph detection: $morphSlideCount transition(s), $totalPairsFound image pair(s) linked" -ForegroundColor Cyan
            
            if (Test-VerboseMode) {
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
                $changes += "${key}:${oldVal}->${newVal}"
                
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
                
                $details += "  * $side edge: $oldVal -> $newVal (was $reason)"
            }
        }
        
        if ($changes.Count -gt 0) {
            Write-Host "[NOTE] Corrected illegal crop values for '$($usage.ShapeName)' on $($usage.Location) ($($changes -join ', '))" -ForegroundColor Cyan
            
            if (Test-VerboseMode) {
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
                        Write-Verbose "    * Position: x=$($off.GetAttribute('x')) EMU, y=$($off.GetAttribute('y')) EMU"
                        Write-Verbose "    * Size: cx=$($ext.GetAttribute('cx')) EMU ($($usage.DisplayWidthPx)px), cy=$($ext.GetAttribute('cy')) EMU ($($usage.DisplayHeightPx)px)"
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
        
        Write-Host "`n[CROP] Processing image crops..." -ForegroundColor Yellow
        
        $croppedCount = 0
        $skippedCount = 0
        
        # Determine if slide filters are active (used for scoping)
        $slideFiltersActive = ($IncludeSlides -and $IncludeSlides.Count -gt 0) -or ($ExcludeSlides -and $ExcludeSlides.Count -gt 0)
        
        # Get items with crops that are in scope for progress counter
        # (excludes items silently skipped due to slide filters, includes SVG fallbacks since they get skip messages)
        $itemsWithCrops = @($usages | Where-Object { 
            $_.HasSrcRect -and
            # Check slide filter
            ($_.ContextType -ne [ContextType]::Slide -or (Test-SlideIncluded -SlideNumber $_.SlideNumber)) -and
            # When slide filters active, exclude masters/layouts
            (-not $slideFiltersActive -or $_.ContextType -eq [ContextType]::Slide)
        })
        $totalItems = $itemsWithCrops.Count
        $currentItem = 0
        
        foreach ($usage in $usages) {
            if (-not $usage.HasSrcRect) { continue }
            
            # Silent skip checks FIRST (before progress update)
            
            # Check slide filter (silently skip usages not in included slides)
            if ($usage.ContextType -eq [ContextType]::Slide -and -not (Test-SlideIncluded -SlideNumber $usage.SlideNumber)) {
                continue
            }
            
            # When slide filters are active, skip masters/layouts entirely
            # (they affect all slides, not just the filtered ones)
            if ($slideFiltersActive -and ($usage.ContextType -eq [ContextType]::Master -or $usage.ContextType -eq [ContextType]::Layout)) {
                continue
            }
            
            # Now update progress (item is in scope)
            $currentItem++
            Write-Progress -Activity "Processing crops" -Status "$currentItem of $totalItems" `
                -PercentComplete (($currentItem / [Math]::Max(1, $totalItems)) * 100) -Id 2
            
            # SVG fallback check
            if ($usage.IsSvgFallbackUsage) {
                Write-Host "   [SKIP] Skipped crop for '$($usage.ShapeName)' on $($usage.Location): SVG fallback (PowerPoint regenerates)" -ForegroundColor Gray
                $usage.OptimizationStatus = 'Skipped_SvgFallback'
                $usage.WhyNotOptimized = 'PNG is auto-generated SVG fallback; PowerPoint regenerates'
                $usage.ManualActionRequired = $false
                $skippedCount++
                continue
            }
            
            # Animated GIF check - cropping would break animation
            $ext = [System.IO.Path]::GetExtension($usage.ImagePhysicalPath).ToLower()
            if ($ext -eq '.gif' -and (Test-IsAnimatedGif -ImagePath $usage.ImagePhysicalPath)) {
                Write-Host "   [SKIP] Skipped crop for '$($usage.ShapeName)' on $($usage.Location): animated GIF (would break animation)" -ForegroundColor Gray
                $usage.OptimizationStatus = 'Skipped_AnimatedGif'
                $usage.WhyNotOptimized = 'Animated GIF - would break animation'
                $usage.ManualActionRequired = $false
                $skippedCount++
                continue
            }
            
            # Check if cropping is enabled for this context
            if ($usage.ContextType -eq [ContextType]::Slide) {
                if (-not $script:CropSlides) {
                    Write-Host "   [SKIP] Skipped crop for '$($usage.ShapeName)' on $($usage.Location): enable -CropSlides" -ForegroundColor Gray
                    $usage.OptimizationStatus = 'Skipped_CropNotMaterialized'
                    $usage.WhyNotOptimized = 'Slide cropping disabled; enable -CropSlides'
                    $usage.ManualActionRequired = $true
                    $usage.ManualActionHint = 'Run with -CropSlides'
                    $skippedCount++
                    continue
                }
            } elseif ($usage.ContextType -eq [ContextType]::Master -or $usage.ContextType -eq [ContextType]::Layout) {
                if (-not $script:CropMastersAndLayouts) {
                    Write-Host "   [SKIP] Skipped crop for '$($usage.ShapeName)' on $($usage.Location): enable -CropMastersAndLayouts" -ForegroundColor Gray
                    $usage.OptimizationStatus = 'Skipped_CropNotMaterialized'
                    $usage.WhyNotOptimized = 'Master/Layout cropping disabled; enable -CropMastersAndLayouts'
                    $usage.ManualActionRequired = $true
                    $usage.ManualActionHint = 'Run with -CropMastersAndLayouts'
                    $skippedCount++
                    continue
                }
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
                Write-Host "   [NOTE] Removed no-op crop for '$($usage.ShapeName)' on $($usage.Location) (full image)" -ForegroundColor Cyan
                continue
            }
            
            # Check Morph pairs
            if ($usage.MorphPair) {
                $pair = $usage.MorphPair
                $usageHasMeaningful = -not (Test-IsNoOpCrop -left $usage.SrcRectLeft -top $usage.SrcRectTop -right $usage.SrcRectRight -bottom $usage.SrcRectBottom)
                $pairHasMeaningful = $pair.HasSrcRect -and -not (Test-IsNoOpCrop -left $pair.SrcRectLeft -top $pair.SrcRectTop -right $pair.SrcRectRight -bottom $pair.SrcRectBottom)
                
                if ($usageHasMeaningful -or $pairHasMeaningful) {
                    Write-Host "   [BLOCK] Morph crop conflict for '$($usage.ShapeName)' across Slide $($pair.SlideNumber) <-> $($usage.SlideNumber) (skipping crop & optimization)" -ForegroundColor Red
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
                Write-Host "[BLOCK] Skipped crop for '$($usage.ShapeName)' on $($usage.Location): invalid crop geometry" -ForegroundColor Red
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
        
        Write-Progress -Activity "Processing crops" -Completed -Id 2
        Write-Host "[STATS] Cropping phase complete: $croppedCount cropped, $skippedCount skipped" -ForegroundColor Green
    }

    function Test-CropValidity {
        param([ImageUsage]$usage)
        
        $l = $usage.SrcRectLeft
        $t = $usage.SrcRectTop
        $r = $usage.SrcRectRight
        $b = $usage.SrcRectBottom
        
        # Check bounds
        if ($l -lt 0 -or $l -gt $script:Config.CROP_THOUSANDTHS_MAX) {
            if (Test-VerboseMode) {
                Write-Verbose "  Crop validation failed for '$($usage.ShapeName)': left=$l out of range [0-$($script:Config.CROP_THOUSANDTHS_MAX)]"
            }
            return $false
        }
        if ($t -lt 0 -or $t -gt $script:Config.CROP_THOUSANDTHS_MAX) {
            if (Test-VerboseMode) {
                Write-Verbose "  Crop validation failed for '$($usage.ShapeName)': top=$t out of range [0-$($script:Config.CROP_THOUSANDTHS_MAX)]"
            }
            return $false
        }
        if ($r -lt 0 -or $r -gt $script:Config.CROP_THOUSANDTHS_MAX) {
            if (Test-VerboseMode) {
                Write-Verbose "  Crop validation failed for '$($usage.ShapeName)': right=$r out of range [0-$($script:Config.CROP_THOUSANDTHS_MAX)]"
            }
            return $false
        }
        if ($b -lt 0 -or $b -gt $script:Config.CROP_THOUSANDTHS_MAX) {
            if (Test-VerboseMode) {
                Write-Verbose "  Crop validation failed for '$($usage.ShapeName)': bottom=$b out of range [0-$($script:Config.CROP_THOUSANDTHS_MAX)]"
            }
            return $false
        }
        
        # Check sum
        if (($l + $r) -ge $script:Config.CROP_THOUSANDTHS_MAX) {
            if (Test-VerboseMode) {
                Write-Verbose "  Crop validation failed for '$($usage.ShapeName)': left+right=$($l+$r) >= $($script:Config.CROP_THOUSANDTHS_MAX) (crops overlap horizontally)"
            }
            return $false
        }
        if (($t + $b) -ge $script:Config.CROP_THOUSANDTHS_MAX) {
            if (Test-VerboseMode) {
                Write-Verbose "  Crop validation failed for '$($usage.ShapeName)': top+bottom=$($t+$b) >= $($script:Config.CROP_THOUSANDTHS_MAX) (crops overlap vertically)"
            }
            return $false
        }
        
        # Check minimum dimensions
        $crop = Get-CropPercentage -left $l -top $t -right $r -bottom $b
        if ($crop.WidthPercent -lt $script:Config.MIN_CROPPED_DIMENSION_PERCENT) {
            if (Test-VerboseMode) {
                Write-Verbose "  Crop validation failed for '$($usage.ShapeName)': resulting width $($crop.WidthPercent.ToString('F1'))% < minimum $($script:Config.MIN_CROPPED_DIMENSION_PERCENT)%"
            }
            return $false
        }
        if ($crop.HeightPercent -lt $script:Config.MIN_CROPPED_DIMENSION_PERCENT) {
            if (Test-VerboseMode) {
                Write-Verbose "  Crop validation failed for '$($usage.ShapeName)': resulting height $($crop.HeightPercent.ToString('F1'))% < minimum $($script:Config.MIN_CROPPED_DIMENSION_PERCENT)%"
            }
            return $false
        }
        
        if (Test-VerboseMode) {
            Write-Verbose "  Crop validation passed for '$($usage.ShapeName)': l=$l, t=$t, r=$r, b=$b -> $($crop.WidthPercent.ToString('F1'))%x$($crop.HeightPercent.ToString('F1'))% of original"
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
            $dims = Get-ImageDimensions -ImagePath $usage.ImagePhysicalPath
            $srcWidth = $dims.Width
            $srcHeight = $dims.Height
            
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
            $output = New-UniqueMediaPath -basePath $usage.ImagePhysicalPath -suffix '_cropped'
            $outputPath = $output.Path
            $outputName = $output.Name
            
            # Execute crop
            $cropGeometry = "${cropWidth}x${cropHeight}+${leftPx}+${topPx}"
            
            if (Test-VerboseMode) {
                Write-Verbose "  Materializing crop for '$($usage.ShapeName)':"
                Write-Verbose "    Source: ${srcWidth}x${srcHeight}px"
                Write-Verbose "    Crop values: l=$($usage.SrcRectLeft), t=$($usage.SrcRectTop), r=$($usage.SrcRectRight), b=$($usage.SrcRectBottom)"
                Write-Verbose "    Resulting geometry: $cropGeometry (${cropWidth}x${cropHeight}px)"
            }
            
            & $script:MagickExe $usage.ImagePhysicalPath -crop $cropGeometry +repage $outputPath 2>&1 | Out-Null
            
            if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $outputPath)) {
                throw "Crop operation failed"
            }
            
            $beforeSize = $usage.BeforeSizeBytes
            $afterSize = (Get-Item $outputPath).Length
            $savedPercent = (($beforeSize - $afterSize) / $beforeSize) * 100
            
            if (Test-VerboseMode) {
                Write-Verbose "    Output: $outputName [$(Format-ByteSize $beforeSize) -> $(Format-ByteSize $afterSize), saved $($savedPercent.ToString('F1'))%]"
            }
            
            # Update relationship
            $null = Update-BlipRelationship -usage $usage -tempDir $tempDir -newMediaFileName $outputName
            
            # Remove srcRect (sibling of blip within blipFill)
            $blipFill = $usage.BlipElement.ParentNode
            if ($blipFill) {
                $srcRect = $blipFill.SelectSingleNode('a:srcRect', $script:NsMgr)
                if ($srcRect) {
                    $null = $srcRect.ParentNode.RemoveChild($srcRect)
                }
                
                # Also reset fillRect if it has negative values (stretch crop)
                # The crop has been materialized, so fillRect should be reset to no offset
                $stretch = $blipFill.SelectSingleNode('a:stretch', $script:NsMgr)
                if ($stretch) {
                    $fillRect = $stretch.SelectSingleNode('a:fillRect', $script:NsMgr)
                    if ($fillRect) {
                        # Check if any attributes were negative (indicating crop)
                        $hasNegative = $false
                        foreach ($attr in @('l', 't', 'r', 'b')) {
                            $val = $fillRect.GetAttribute($attr)
                            if ($val -and [double]$val -lt 0) {
                                $hasNegative = $true
                                break
                            }
                        }
                        
                        if ($hasNegative) {
                            # Remove all attributes to reset to default (no offset)
                            $fillRect.RemoveAttribute('l')
                            $fillRect.RemoveAttribute('t')
                            $fillRect.RemoveAttribute('r')
                            $fillRect.RemoveAttribute('b')
                            Write-Verbose "    Reset fillRect attributes (crop materialized into image)"
                        }
                    }
                }
            }
            
            $usage.HasSrcRect = $false
            $usage.CropApplied = $true
            $usage.AfterSizeBytes = $afterSize
            $usage.OptimizedFile = $outputName
            $usage.ImagePhysicalPath = $outputPath  # CRITICAL: Update path to cropped file for optimization phase
            
            if (Test-VerboseMode) {
                Write-Verbose "    Updated ImagePhysicalPath: $outputPath (optimization will use cropped file)"
            }
            
            $savedPercent = (($beforeSize - $afterSize) / $beforeSize) * 100
            Write-Host "   [CROP] Cropped '$($usage.ShapeName)' on $($usage.Location) (saved $($savedPercent.ToString('F1'))%)" -ForegroundColor Green
            
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
        
        Write-Host "`n[PROC] Optimizing images..." -ForegroundColor Yellow
        
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
            $slideUsages = $_.Usages | Where-Object { $_.ContextType -eq [ContextType]::Slide }
            if ($slideUsages) {
                ($slideUsages | Measure-Object -Property SlideNumber -Minimum).Minimum
            } else {
                9999  # No slide usages (master/layout only), sort to end
            }
        }
        
        # Determine if slide filters are active (used for scoping)
        $slideFiltersActive = ($IncludeSlides -and $IncludeSlides.Count -gt 0) -or ($ExcludeSlides -and $ExcludeSlides.Count -gt 0)
        
        # Pre-filter groups to only those in scope for accurate progress count
        $inScopeGroups = @($sortedGroups | Where-Object {
            $grp = $_
            $hasIncluded = $grp.Usages | Where-Object {
                if ($slideFiltersActive) {
                    $_.ContextType -eq [ContextType]::Slide -and (Test-SlideIncluded -SlideNumber $_.SlideNumber)
                } else {
                    $_.ContextType -ne [ContextType]::Slide -or (Test-SlideIncluded -SlideNumber $_.SlideNumber)
                }
            }
            $hasIncluded
        })
        
        $totalGroups = $inScopeGroups.Count
        $currentGroup = 0
        
        foreach ($group in $sortedGroups) {
            $mediaFile = Split-Path $group.PhysicalPath -Leaf
            
            # Check slide filter FIRST - skip groups where no usage is on an included slide
            # When slide filters are active, also skip masters/layouts (they affect all slides)
            $hasIncludedUsage = $group.Usages | Where-Object {
                if ($slideFiltersActive) {
                    # Only include slide usages that pass the filter
                    $_.ContextType -eq [ContextType]::Slide -and (Test-SlideIncluded -SlideNumber $_.SlideNumber)
                } else {
                    # No filter - include all (slides pass through, masters/layouts always included)
                    $_.ContextType -ne [ContextType]::Slide -or (Test-SlideIncluded -SlideNumber $_.SlideNumber)
                }
            }
            if (-not $hasIncludedUsage) {
                continue
            }
            
            # Now update progress (group is in scope)
            $currentGroup++
            Write-Progress -Activity "Optimizing images" -Status "$mediaFile ($currentGroup of $totalGroups)" `
                -PercentComplete (($currentGroup / [Math]::Max(1, $totalGroups)) * 100) -Id 3
            
            # Check if optimization forbidden (only after confirming group is in scope)
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
                # Log first included usage as representative
                $firstUsage = $hasIncludedUsage | Select-Object -First 1
                Write-Host "   [SKIP] Skipped optimization for '$($firstUsage.ShapeName)' on $($firstUsage.Location): $($skipReason.Reason)" -ForegroundColor Gray
                $skippedCount += $group.Usages.Count
                continue
            }
            
            # Calculate target dimensions
            $maxDisplay = @{
                Width = ($group.Usages | Measure-Object -Property DisplayWidthPx -Maximum).Maximum
                Height = ($group.Usages | Measure-Object -Property DisplayHeightPx -Maximum).Maximum
            }
            
            $firstUsage = $hasIncludedUsage | Select-Object -First 1
            
            if ($PSCmdlet.ShouldProcess($mediaFile, "Optimize image")) {
                Write-Verbose "Optimizing $mediaFile [display: $($maxDisplay.Width)x$($maxDisplay.Height)px] from $($firstUsage.Location)"
                
                $result = Invoke-OptimizationOperation -group $group -maxDisplay $maxDisplay -tempDir $tempDir
                if ($result.Success) {
                    $optimizedCount++
                } else {
                    $skippedCount += $group.Usages.Count
                }
            }
        }
        
        Write-Progress -Activity "Optimizing images" -Completed -Id 3
        Write-Host "[STATS] Optimization phase complete: $optimizedCount optimized, $skippedCount skipped" -ForegroundColor Green
    }

    function Get-OptimizationSkipReason {
        param([ImageGroup]$group)
        
        # Check file format - skip vector formats and other non-optimizable types
        $ext = [System.IO.Path]::GetExtension($group.PhysicalPath).ToLower()
        $supportedFormats = $script:Config.SUPPORTED_RASTER_FORMATS
        $convertibleFormats = $script:Config.CONVERTIBLE_FORMATS
        
        # Check for animated GIF (must skip - cropping/resizing breaks animation)
        if ($ext -eq '.gif' -and (Test-IsAnimatedGif -ImagePath $group.PhysicalPath)) {
            return @{
                Status = 'Skipped_AnimatedGif'
                Reason = 'Animated GIF - would break animation'
                ManualAction = $false
                Hint = ''
            }
        }
        
        if ($ext -notin $supportedFormats -and $ext -notin $convertibleFormats) {
            $fileName = Split-Path $group.PhysicalPath -Leaf
            $formatType = switch ($ext) {
                '.emf' { 'vector graphic (EMF)' }
                '.wmf' { 'vector graphic (WMF)' }
                '.svg' { 'vector graphic (SVG)' }
                default { "format ($ext)" }
            }
            
            return @{
                Status = 'Skipped_UnsupportedFormat'
                Reason = "Unsupported $formatType"
                ManualAction = $false
                Hint = ''
            }
        }
        
        # Capture source dimensions for reporting, even if image will be skipped
        try {
            $dims = Get-ImageDimensions -ImagePath $group.PhysicalPath
            
            # Store in all usages for CSV reporting
            foreach ($usage in $group.Usages) {
                if ($usage.SourceWidthPx -eq 0) {
                    $usage.SourceWidthPx = $dims.Width
                    $usage.SourceHeightPx = $dims.Height
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
        
        # Check if optimization is enabled for this context
        $hasSlide = $group.Usages | Where-Object { $_.ContextType -eq [ContextType]::Slide }
        $hasTemplate = $group.Usages | Where-Object { $_.ContextType -in @([ContextType]::Master, [ContextType]::Layout) }
        
        $contextTypes = ($group.Usages | Select-Object -ExpandProperty ContextType -Unique) -join ', '
        $fileName = Split-Path $group.PhysicalPath -Leaf
        Write-Verbose "  Checking skip reasons for $($fileName): contexts=[$contextTypes], hasSlide=$($null -ne $hasSlide), hasTemplate=$($null -ne $hasTemplate), OptimizeSlides=$script:OptimizeSlides, OptimizeMastersAndLayouts=$script:OptimizeMastersAndLayouts"
        
        if ($hasSlide -and -not $script:OptimizeSlides) {
            return @{
                Status = 'Skipped_SlidesFlagMissing'
                Reason = 'Used in slide; enable -OptimizeSlides'
                ManualAction = $true
                Hint = 'Run with -OptimizeSlides'
            }
        }
        
        if ($hasTemplate -and -not $script:OptimizeMastersAndLayouts) {
            return @{
                Status = 'Skipped_TemplatesFlagMissing'
                Reason = 'Used in master/layout; enable -OptimizeMastersAndLayouts'
                ManualAction = $true
                Hint = 'Run with -OptimizeMastersAndLayouts'
            }
        }
        
        # Check unresolved crops - determine appropriate hint based on context
        $hasUnresolvedCrop = @($group.Usages | Where-Object { $_.HasSrcRect -and -not $_.CropApplied })
        if ($hasUnresolvedCrop.Count -gt 0) {
            $firstUnresolved = $hasUnresolvedCrop[0]
            $hint = if ($firstUnresolved.ContextType -in @([ContextType]::Master, [ContextType]::Layout)) {
                'Run with -CropMastersAndLayouts'
            } else {
                'Run with -CropSlides'
            }
            return @{
                Status = 'Skipped_CropNotMaterialized'
                Reason = 'Crop present; enable cropping'
                ManualAction = $true
                Hint = $hint
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
            $dims = Get-ImageDimensions -ImagePath $group.PhysicalPath
            $srcWidth = $dims.Width
            $srcHeight = $dims.Height
            
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
            
            Write-Verbose "  Dimensions: source ${srcWidth}x${srcHeight}px -> target ${targetWidth}x${targetHeight}px [HeadroomFactor: $HeadroomFactor]"
            
            # Check JPEG quality for quality preservation
            $ext = [System.IO.Path]::GetExtension($group.PhysicalPath).ToLower()
            if ($ext -in @('.jpg', '.jpeg')) {
                $sourceQuality = Get-JpegQuality -ImagePath $group.PhysicalPath
                foreach ($usage in $group.Usages) {
                    $usage.SourceJpegQuality = $sourceQuality
                }
                
                # If source quality is at or below target and dimensions match, skip
                if ($sourceQuality -gt 0 -and $sourceQuality -le $JpegQuality -and 
                    $srcWidth -eq $targetWidth -and $srcHeight -eq $targetHeight) {
                    $firstUsage = $group.Usages[0]
                    foreach ($usage in $group.Usages) {
                        $usage.OptimizationStatus = 'Skipped_AlreadyOptimal'
                        $usage.WhyNotOptimized = "Already at optimal quality (Q$sourceQuality <= Q$JpegQuality) and target size"
                        $usage.ManualActionRequired = $false
                    }
                    Write-Host "   [SKIP] Skipped '$($firstUsage.ShapeName)' on $($firstUsage.Location): already at target size and quality (Q$sourceQuality)" -ForegroundColor Gray
                    return @{ Success = $false }
                }
            }
            
            # Handle convertible formats (BMP, TIFF, GIF)
            if ($ext -in $script:Config.CONVERTIBLE_FORMATS) {
                $result = ConvertTo-OptimizedFormat -group $group -targetWidth $targetWidth -targetHeight $targetHeight -tempDir $tempDir
                return $result
            }
            
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
                        $firstUsage = $group.Usages[0]
                        Write-Host "   [INFO] Effective transparency $($transp.ToString('F1'))% for '$($firstUsage.ShapeName)' on $($firstUsage.Location) - treating as opaque" -ForegroundColor Cyan
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
                Write-Host "   [SKIP] No change needed for '$($firstUsage.ShapeName)' on $($firstUsage.Location) (already at target)" -ForegroundColor Gray
                return @{ Success = $false }
            }
            
            # Determine optimization strategy
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
                    $firstUsage = $group.Usages[0]
                    Write-Host "   [INFO] Effective transparency $($transp.ToString('F1'))% for '$($firstUsage.ShapeName)' on $($firstUsage.Location) - treating as opaque" -ForegroundColor Cyan
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
            
            if (Test-VerboseMode) {
                $decision = if ($transp -le $TransparencyThresholdPercent) { "opaque (can convert to JPEG)" } else { "transparent (keep as PNG)" }
                Write-Verbose "  Transparency analysis for $($fileName): $($transp.ToString('F2'))% transparent pixels -> $decision"
            }
            
            return $transp
        }
        
        if (Test-VerboseMode) {
            Write-Verbose "  Transparency analysis for $($fileName): unable to determine, assuming opaque (0%)"
        }
        
        return 0.0
    }

    function Test-MinSavingsThreshold {
        param(
            [long]$beforeSize,
            [long]$afterSize
        )
        
        # Always reject if file grew larger (negative savings)
        if ($afterSize -ge $beforeSize) {
            return $false
        }
        
        if ($MinSavingsPercent -le 0) {
            return $true  # No minimum threshold (but still reject growth)
        }
        
        $savingsPercent = (($beforeSize - $afterSize) / $beforeSize) * 100
        return $savingsPercent -ge $MinSavingsPercent
    }

    function Optimize-Jpeg {
        param(
            [ImageGroup]$group,
            [int]$targetWidth,
            [int]$targetHeight,
            [string]$tempDir
        )
        
        Write-Verbose "  Optimizing JPEG: target ${targetWidth}x${targetHeight}px, quality $JpegQuality"
        
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
        
        # Check minimum savings threshold
        if ($afterSize -ge $beforeSize -or -not (Test-MinSavingsThreshold -beforeSize $beforeSize -afterSize $afterSize)) {
            Remove-Item -LiteralPath $outputPath -Force
            $firstUsage = $group.Usages[0]
            
            $reason = if ($afterSize -ge $beforeSize) {
                'No net savings achieved'
            } else {
                $actualSavings = (($beforeSize - $afterSize) / $beforeSize) * 100
                "Savings $($actualSavings.ToString('F1'))% below $($MinSavingsPercent.ToString('F1'))% threshold"
            }
            
            foreach ($usage in $group.Usages) {
                $usage.OptimizationStatus = 'Skipped_NoNetSavings'
                $usage.WhyNotOptimized = $reason
                $usage.ManualActionRequired = $false
            }
            
            if (Test-VerboseMode) {
                Write-Verbose "  JPEG optimization rejected: $(Format-ByteSize $beforeSize) -> $(Format-ByteSize $afterSize)"
            }
            
            Write-Host "   [INFO] $reason for '$($firstUsage.ShapeName)' on $($firstUsage.Location) - kept original" -ForegroundColor Cyan
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
        
        Write-Host "   [OK] Optimized JPEG for '$($firstUsage.ShapeName)' on $($firstUsage.Location) (saved $($savedPercent.ToString('F1'))%)" -ForegroundColor Green
        
        return @{ Success = $true }
    }

    function Optimize-PngWithAlpha {
        param(
            [ImageGroup]$group,
            [int]$targetWidth,
            [int]$targetHeight,
            [string]$tempDir
        )
        
        Write-Verbose "  Optimizing PNG with transparency: target ${targetWidth}x${targetHeight}px, compression level $($script:Config.PNG_COMPRESSION_LEVEL)"
        
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
        
        # Check minimum savings threshold
        if ($afterSize -ge $beforeSize -or -not (Test-MinSavingsThreshold -beforeSize $beforeSize -afterSize $afterSize)) {
            Remove-Item -LiteralPath $outputPath -Force
            $firstUsage = $group.Usages[0]
            
            $reason = if ($afterSize -ge $beforeSize) {
                'No net savings achieved'
            } else {
                $actualSavings = (($beforeSize - $afterSize) / $beforeSize) * 100
                "Savings $($actualSavings.ToString('F1'))% below $($MinSavingsPercent.ToString('F1'))% threshold"
            }
            
            foreach ($usage in $group.Usages) {
                $usage.OptimizationStatus = 'Skipped_NoNetSavings'
                $usage.WhyNotOptimized = $reason
                $usage.ManualActionRequired = $false
            }
            
            if (Test-VerboseMode) {
                Write-Verbose "  PNG optimization rejected: $(Format-ByteSize $beforeSize) -> $(Format-ByteSize $afterSize)"
            }
            
            Write-Host "   [INFO] $reason for '$($firstUsage.ShapeName)' on $($firstUsage.Location) - kept original" -ForegroundColor Cyan
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
        
        Write-Host "   [OK] Optimized PNG with transparency for '$($firstUsage.ShapeName)' on $($firstUsage.Location) (saved $($savedPercent.ToString('F1'))%)" -ForegroundColor Green
        
        return @{ Success = $true }
    }

    function ConvertTo-Jpeg {
        param(
            [ImageGroup]$group,
            [string]$tempDir,
            [bool]$resize
        )
        
        $srcFile = Split-Path $group.PhysicalPath -Leaf
        Write-Verbose "  Converting PNG->JPEG: $srcFile [resize: $resize, quality: $JpegQuality]"
        
        # Generate unique JPEG filename
        $output = New-UniqueMediaPath -basePath $group.PhysicalPath -suffix '' -newExtension '.jpeg'
        $outputPath = $output.Path
        $outputName = $output.Name
        
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
        
        # Check minimum savings threshold
        if ($afterSize -ge $beforeSize -or -not (Test-MinSavingsThreshold -beforeSize $beforeSize -afterSize $afterSize)) {
            Remove-Item -LiteralPath $outputPath -Force
            $firstUsage = $group.Usages[0]
            
            $reason = if ($afterSize -ge $beforeSize) {
                'Conversion yielded no savings'
            } else {
                $actualSavings = (($beforeSize - $afterSize) / $beforeSize) * 100
                "Savings $($actualSavings.ToString('F1'))% below $($MinSavingsPercent.ToString('F1'))% threshold"
            }
            
            foreach ($usage in $group.Usages) {
                $usage.OptimizationStatus = 'Skipped_NoNetSavings'
                $usage.WhyNotOptimized = $reason
                $usage.ManualActionRequired = $false
            }
            
            if (Test-VerboseMode) {
                Write-Verbose "  PNG->JPEG conversion rejected: $(Format-ByteSize $beforeSize) -> $(Format-ByteSize $afterSize)"
            }
            
            Write-Host "   [INFO] $reason for '$($firstUsage.ShapeName)' on $($firstUsage.Location) - kept original" -ForegroundColor Cyan
            return @{ Success = $false }
        }
        
        # Update all relationships pointing to this image
        Write-Verbose "  Updating $($group.Usages.Count) relationship(s) to point to new JPEG: $outputName"
        
        foreach ($usage in $group.Usages) {
            $null = Update-BlipRelationship -usage $usage -tempDir $tempDir -newMediaFileName $outputName
            
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
        Write-Host "   [CONV] Converted PNG to JPEG for '$($firstUsage.ShapeName)' on $($firstUsage.Location) (saved $($savedPercent.ToString('F1'))%)" -ForegroundColor Green
        
        return @{ Success = $true }
    }

    function ConvertTo-OptimizedFormat {
        param(
            [ImageGroup]$group,
            [int]$targetWidth,
            [int]$targetHeight,
            [string]$tempDir
        )
        
        $srcFile = Split-Path $group.PhysicalPath -Leaf
        $ext = [System.IO.Path]::GetExtension($group.PhysicalPath).ToLower()
        
        # Check if image has transparency (for TIFF and GIF)
        $hasTransparency = $false
        if ($ext -in @('.tif', '.tiff', '.gif')) {
            $transp = Get-ImageTransparency -imagePath $group.PhysicalPath
            $hasTransparency = $transp -gt $TransparencyThresholdPercent
            foreach ($usage in $group.Usages) {
                $usage.EffectiveTransparencyPercent = $transp
            }
            
            if ($hasTransparency) {
                $firstUsage = $group.Usages[0]
                Write-Host "   [INFO] Effective transparency $($transp.ToString('F1').Replace('.', ','))% for '$($firstUsage.ShapeName)' on $($firstUsage.Location) - keeping as PNG" -ForegroundColor Cyan
            } else {
                $firstUsage = $group.Usages[0]
                Write-Host "   [INFO] Effective transparency $($transp.ToString('F1').Replace('.', ','))% for '$($firstUsage.ShapeName)' on $($firstUsage.Location) - treating as opaque" -ForegroundColor Cyan
            }
        }
        
        # Determine output format
        $outputExt = if ($hasTransparency) { '.png' } else { '.jpeg' }
        
        Write-Verbose "  Converting $ext->$($outputExt): $srcFile [target: ${targetWidth}x${targetHeight}px]"
        
        # Generate unique filename
        $output = New-UniqueMediaPath -basePath $group.PhysicalPath -suffix '' -newExtension $outputExt
        $outputPath = $output.Path
        $outputName = $output.Name
        
        $magickArgs = @(
            $group.PhysicalPath,
            '-filter', $script:Config.RESIZE_FILTER,
            '-resize', "${targetWidth}x${targetHeight}>",
            '-strip'
        )
        
        if ($outputExt -eq '.jpeg') {
            $magickArgs += @(
                '-quality', $JpegQuality,
                '-define', 'jpeg:optimize-coding=true',
                '-define', 'jpeg:dct-method=float',
                '-sampling-factor', $script:Config.JPEG_SAMPLING_FACTOR_DEFAULT
            )
        } else {
            $magickArgs += @(
                '-define', "png:compression-level=$($script:Config.PNG_COMPRESSION_LEVEL)",
                '-define', 'png:color-type=6'
            )
        }
        
        $magickArgs += $outputPath
        
        & $script:MagickExe $magickArgs 2>&1 | Out-Null
        
        if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $outputPath)) {
            throw "$ext conversion failed"
        }
        
        $beforeSize = $group.OriginalSizeBytes
        $afterSize = (Get-Item $outputPath).Length
        
        # Check minimum savings threshold
        if (-not (Test-MinSavingsThreshold -beforeSize $beforeSize -afterSize $afterSize)) {
            Remove-Item -LiteralPath $outputPath -Force
            $firstUsage = $group.Usages[0]
            
            $reason = if ($afterSize -ge $beforeSize) {
                'Conversion yielded no savings'
            } else {
                $actualSavings = (($beforeSize - $afterSize) / $beforeSize) * 100
                "Savings $($actualSavings.ToString('F1'))% below $($MinSavingsPercent.ToString('F1'))% threshold"
            }
            
            foreach ($usage in $group.Usages) {
                $usage.OptimizationStatus = 'Skipped_NoNetSavings'
                $usage.WhyNotOptimized = $reason
                $usage.ManualActionRequired = $false
            }
            
            Write-Host "   [INFO] $reason for '$($firstUsage.ShapeName)' on $($firstUsage.Location) - kept original" -ForegroundColor Cyan
            return @{ Success = $false }
        }
        
        # Update all relationships
        Write-Verbose "  Updating $($group.Usages.Count) relationship(s) to point to new file: $outputName"
        
        $statusName = if ($outputExt -eq '.jpeg') { 'ConvertedToJpeg' } else { 'ConvertedToPng' }
        
        foreach ($usage in $group.Usages) {
            $null = Update-BlipRelationship -usage $usage -tempDir $tempDir -newMediaFileName $outputName
            
            $usage.AfterSizeBytes = $afterSize
            $usage.OptimizationStatus = $statusName
            $usage.OptimizedFile = $outputName
        }
        
        # Remove original
        Remove-Item -LiteralPath $group.PhysicalPath -Force
        $group.OptimizedSizeBytes = $afterSize
        $group.OptimizedPath = $outputPath
        
        $savedPercent = (($beforeSize - $afterSize) / $beforeSize) * 100
        $firstUsage = $group.Usages[0]
        Write-Host "   [CONV] Converted $ext to $outputExt for '$($firstUsage.ShapeName)' on $($firstUsage.Location) (saved $($savedPercent.ToString('F1'))%)" -ForegroundColor Green
        
        return @{ Success = $true }
    }

    #endregion

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

    #region Cleanup and Packaging

    function Invoke-MediaCleanup {
        param([string]$tempDir)

        Write-Host "`n[CLEAN] Cleaning up unreferenced media..." -ForegroundColor Yellow

        $relsFiles = Get-ChildItem -Path $tempDir -Filter "*.rels" -Recurse
        $mediaDir = Join-Path $tempDir 'ppt\media'
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

    #region CSV Reporting

    function Export-CsvReport {
        param(
            [System.Collections.Generic.List[ImageUsage]]$usages
        )
        
        Write-Host "`n[CSV] Generating report..." -ForegroundColor Yellow
        
        # Create temp file for CSV (outside the extraction temp dir)
        $tempCsvPath = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "ppt-opt-$([System.Guid]::NewGuid().ToString('N').Substring(0,8)).csv")
        Write-Verbose "Temp CSV file: $tempCsvPath"
        
        if (-not $usages -or $usages.Count -eq 0) {
            Write-Host "No usages to report" -ForegroundColor Gray
            @() | Export-Csv -Path $tempCsvPath -NoTypeInformation -Encoding UTF8
            return $tempCsvPath
        }
        
        # Filter usages based on slide selection (same logic as processing phases)
        $slideFiltersActive = ($IncludeSlides -and $IncludeSlides.Count -gt 0) -or ($ExcludeSlides -and $ExcludeSlides.Count -gt 0)
        if ($slideFiltersActive) {
            $filteredUsages = @($usages | Where-Object {
                # Only include slide usages that pass the filter (exclude masters/layouts when filtering)
                $_.ContextType -eq [ContextType]::Slide -and (Test-SlideIncluded -SlideNumber $_.SlideNumber)
            })
            Write-Verbose "CSV filtered to $($filteredUsages.Count) usages (slide filter active)"
        } else {
            $filteredUsages = $usages
        }
        
        if ($filteredUsages.Count -eq 0) {
            Write-Host "No usages in scope to report" -ForegroundColor Gray
            @() | Export-Csv -Path $tempCsvPath -NoTypeInformation -Encoding UTF8
            return $tempCsvPath
        }
        
        # Sort by context type (Slide, Layout, Master), then slide number, then layout/master number
        $sorted = $filteredUsages | Sort-Object @(
            @{ Expression = { $_.ContextType }; Ascending = $true }
            @{ Expression = { $_.SlideNumber }; Ascending = $true }
            @{ Expression = { 
                if ($_.PartPath -match '(\d+)\.xml$') { 
                    [int]$matches[1] 
                } else { 
                    0 
                }
            }; Ascending = $true }
        )
        
        $rows = foreach ($usage in $sorted) {
            [PSCustomObject]@{
                Location = $usage.Location
                ContextType = $usage.ContextType.ToString()
                SlideNumber = $usage.SlideNumber
                SlideModified = if ($usage.CropApplied -or $usage.OptimizationStatus -match '^(Optimized|Converted)') { 'Yes' } else { 'No' }
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
                SourceJpegQuality = if ($usage.SourceJpegQuality -gt 0) { $usage.SourceJpegQuality } else { ' ' }
                TransparencyThresholdPercent = $TransparencyThresholdPercent
                EffectiveTransparencyPercent = $usage.EffectiveTransparencyPercent.ToString('F2')
                MinSavingsPercent = $MinSavingsPercent
                ManualActionRequired = if ($usage.ManualActionRequired) { 'Yes' } else { 'No' }
                ManualActionHint = $usage.ManualActionHint
            }
        }
        
        $rows | Export-Csv -Path $tempCsvPath -NoTypeInformation -Encoding UTF8
        Write-Host "[OK] CSV report generated" -ForegroundColor Green
        
        return $tempCsvPath
    }

    #endregion

    #region Output File Handling

    function Copy-OutputFilesToFinal {
        param(
            [string]$tempPptxPath,
            [string]$tempCsvPath,
            [string]$finalPptxPath,
            [string]$finalCsvPath
        )
        
        $result = @{
            PptxSuccess = $false
            CsvSuccess = $false
            PptxTempPath = $tempPptxPath
            CsvTempPath = $tempCsvPath
        }
        
        # Try to copy PPTX
        try {
            if (Test-FileInUse -Path $finalPptxPath) {
                Write-Warning "Cannot write to output file (in use): $finalPptxPath"
            } else {
                if (Test-Path -LiteralPath $finalPptxPath) {
                    Remove-Item -LiteralPath $finalPptxPath -Force
                }
                Copy-Item -LiteralPath $tempPptxPath -Destination $finalPptxPath -Force
                $result.PptxSuccess = $true
                Write-Host "[OUT] Output: $finalPptxPath" -ForegroundColor Green
            }
        } catch {
            Write-Warning "Failed to copy PPTX to final destination: $_"
        }
        
        # Try to copy CSV
        try {
            if (Test-FileInUse -Path $finalCsvPath) {
                Write-Warning "Cannot write to CSV file (in use): $finalCsvPath"
            } else {
                if (Test-Path -LiteralPath $finalCsvPath) {
                    Remove-Item -LiteralPath $finalCsvPath -Force
                }
                Copy-Item -LiteralPath $tempCsvPath -Destination $finalCsvPath -Force
                $result.CsvSuccess = $true
                Write-Host "[OUT] Report: $finalCsvPath" -ForegroundColor Green
            }
        } catch {
            Write-Warning "Failed to copy CSV to final destination: $_"
        }
        
        return $result
    }

    #endregion

    #region Main Processing Function

    function Invoke-PptOptimization {
        param([string]$CurrentInputPath)
        
        $exitCode = 0
        
        try {
            Write-Host "[SCAN] Analyzing presentation structure..." -ForegroundColor Cyan
            Write-Host "   File: $CurrentInputPath" -ForegroundColor Gray
            
            # Initialize and validate
            $validationError = Initialize-Script -InputPath $CurrentInputPath -MagickPath $MagickPath
            if ($validationError) {
                Write-Host "`n[ERROR] $validationError" -ForegroundColor Red
                return [PSCustomObject]@{
                    InputPath = $CurrentInputPath
                    OutputPath = $null
                    ReportPath = $null
                    OriginalSizeBytes = 0
                    OptimizedSizeBytes = 0
                    SavingsBytes = 0
                    SavingsPercent = 0.0
                    ImagesCropped = 0
                    ImagesOptimized = 0
                    ImagesSkipped = 0
                    TotalImageUsages = 0
                    UniqueImages = 0
                    ExitCode = 1
                    Success = $false
                    ErrorType = 'Validation'
                    Error = $validationError
                }
            }
            
            Initialize-Namespaces
            
            # Create temp directory
            $tempDir = Join-Path ([System.IO.Path]::GetTempPath()) "ppt-optimize-$(New-Guid)"
            New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
            
            # Normalize to full path (avoid 8.3 filename issues)
            $tempDir = (Get-Item -LiteralPath $tempDir).FullName
            
            try {
                # Initialize temp output file tracking (for cleanup in finally)
                $tempPptxPath = $null
                $tempCsvPath = $null
                $copyFailed = $true  # Assume failure until copy succeeds
                
                # Extract (use .NET ZipFile for PS 5.1 compatibility - Expand-Archive only accepts .zip extension)
                Add-Type -Assembly 'System.IO.Compression.FileSystem' -ErrorAction SilentlyContinue
                [System.IO.Compression.ZipFile]::ExtractToDirectory($script:InputFile, $tempDir)
                
                # Get canonical slide order
                $slides = Get-CanonicalSlideOrder -tempDir $tempDir
                
                if (-not $slides -or $slides.Count -eq 0) {
                    Write-Warning "No slides found in presentation"
                    $slides = [System.Collections.Generic.List[SlideInfo]]::new()
                }
                
                # Get slide dimensions for background image sizing
                $script:SlideDimensions = Get-SlideDimensions -tempDir $tempDir
                Write-Verbose "Slide dimensions: $($script:SlideDimensions.WidthPx)x$($script:SlideDimensions.HeightPx)px"
                
                # Scan for image usages
                $usages = Get-ImageUsages -tempDir $tempDir -slides $slides
                
                if (-not $usages) {
                    $usages = [System.Collections.Generic.List[ImageUsage]]::new()
                }
                
                Write-Host "[STATS] Found $($usages.Count) image usages" -ForegroundColor Cyan
                $uniqueImages = if ($usages.Count -gt 0) { 
                    ($usages | Select-Object -Property ImagePhysicalPath -Unique).Count 
                } else { 
                    0 
                }
                Write-Host "[STATS] Across $uniqueImages unique images" -ForegroundColor Cyan
                
                if ($usages.Count -eq 0) {
                    Write-Host "`n[INFO] No images found to optimize" -ForegroundColor Yellow
                    
                    # Still create output file and empty CSV
                    Copy-Item $script:InputFile $script:OutputFile
                    
                    $emptyReport = @()
                    $emptyReport | Export-Csv -Path $script:CsvReport -NoTypeInformation -Encoding UTF8
                    
                    Write-Host "`n[DONE] Complete (no changes needed)" -ForegroundColor Green
                    Write-Host "Output: $script:OutputFile" -ForegroundColor Green
                    Write-Host "Report: $script:CsvReport" -ForegroundColor Green
                    
                    $exitCode = 2  # No changes needed
                    
                    # Return summary
                    return [PSCustomObject]@{
                        InputPath = $script:InputFile.ToString()
                        OutputPath = $script:OutputFile
                        ReportPath = $script:CsvReport
                        OriginalSizeBytes = (Get-Item $script:InputFile).Length
                        OptimizedSizeBytes = (Get-Item $script:InputFile).Length
                        SavingsBytes = 0
                        SavingsPercent = 0.0
                        ImagesCropped = 0
                        ImagesOptimized = 0
                        ImagesSkipped = 0
                        TotalImageUsages = 0
                        UniqueImages = 0
                        ExitCode = $exitCode
                        Success = $true
                        ErrorType = 'None'
                    }
                }
                
                # Find Morph pairs
                $morphPairs = Find-MorphPairs -usages $usages
                
                # Phase 1: Normalize crops (only if any cropping enabled)
                if ($script:CropSlides -or $script:CropMastersAndLayouts) {
                    $docsToSave = @{}
                    $slideFiltersActive = ($IncludeSlides -and $IncludeSlides.Count -gt 0) -or ($ExcludeSlides -and $ExcludeSlides.Count -gt 0)
                    
                    foreach ($usage in $usages) {
                        # Skip usages not in scope (same logic as cropping phase)
                        if ($usage.ContextType -eq [ContextType]::Slide -and -not (Test-SlideIncluded -SlideNumber $usage.SlideNumber)) {
                            continue
                        }
                        if ($slideFiltersActive -and ($usage.ContextType -eq [ContextType]::Master -or $usage.ContextType -eq [ContextType]::Layout)) {
                            continue
                        }
                        
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
                if ($script:CropSlides -or $script:CropMastersAndLayouts) {
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
                foreach ($usage in $usages | Where-Object { $_.OptimizationStatus -match '^Converted' }) {
                    if ($usage.SlideDocument) {
                        $slidePath = Join-Path $tempDir ($usage.PartPath -replace '/', '\')
                        $docsToSave[$slidePath] = $usage.SlideDocument
                        $relativePath = $usage.PartPath
                        Write-Verbose "Marking document for save (format conversion): $relativePath"
                    } else {
                        Write-Warning "Converted usage on $($usage.Location) missing SlideDocument reference"
                    }
                }
                
                if ($docsToSave.Count -gt 0) {
                    Write-Host "[SAVE] Saving $($docsToSave.Count) updated document(s)..." -ForegroundColor Yellow
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
                
                # Phase 3.5: Layout/Master cleanup (if requested)
                if ($script:CleanUnusedLayouts) {
                    Invoke-UnusedLayoutCleanup -tempDir $tempDir
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
                    Write-Verbose "  [v] $file"
                }
                Write-Verbose "Pre-repack validation passed - all critical files present"
                
                # Phase 5: Repack (to temp file)
                $tempPptxPath = Invoke-Repack -tempDir $tempDir
                
                # Generate report (to temp file)
                $tempCsvPath = Export-CsvReport -usages $usages
                
                # Final statistics (calculated from temp file)
                $originalSize = (Get-Item $script:InputFile).Length
                $optimizedSize = (Get-Item $tempPptxPath).Length
                $savedBytes = $originalSize - $optimizedSize
                $savedPercent = ($savedBytes / $originalSize) * 100
                
                $croppedCount = @($usages | Where-Object { $_.CropApplied } | Select-Object -Property OriginalFileName -Unique).Count
                $optimizedCount = @($usages | Where-Object { $_.OptimizationStatus -match '^(Optimized|Converted)' } | Select-Object -Property OriginalFileName -Unique).Count
                $skippedCount = @($usages | Where-Object { $_.OptimizationStatus -match '^Skipped' -or $_.OptimizationStatus -eq 'NoChangeNeeded' }).Count
                
                Write-Host "`n[DONE] Optimization Complete" -ForegroundColor Green
                Write-Host "[STATS] Images cropped: $croppedCount" -ForegroundColor Green
                Write-Host "[STATS] Images optimized: $optimizedCount" -ForegroundColor Green
                Write-Host "`n[SIZE] File size:" -ForegroundColor Green
                Write-Host "  Original: $((Format-ByteSize $originalSize))" -ForegroundColor Green
                Write-Host "  Optimized: $((Format-ByteSize $optimizedSize))" -ForegroundColor Green
                Write-Host "  Saved: $(Format-ByteSize $savedBytes) ($($savedPercent.ToString('F1'))%)" -ForegroundColor Green
                
                # Phase 6: Copy to final destinations
                Write-Host "`n[SAVE] Saving output files..." -ForegroundColor Yellow
                $copyResult = Copy-OutputFilesToFinal `
                    -tempPptxPath $tempPptxPath `
                    -tempCsvPath $tempCsvPath `
                    -finalPptxPath $script:OutputFile `
                    -finalCsvPath $script:CsvReport
                
                # Handle any copy failures
                $copyFailed = -not $copyResult.PptxSuccess -or -not $copyResult.CsvSuccess
                if ($copyFailed) {
                    Write-Host "`n[WARN]  Some output files could not be saved (files may be in use):" -ForegroundColor Yellow
                    if (-not $copyResult.PptxSuccess) {
                        Write-Host "  Optimized PPTX saved at: $tempPptxPath" -ForegroundColor Cyan
                        Write-Host "  -> Close the target file and copy manually, or re-run the script." -ForegroundColor Gray
                    }
                    if (-not $copyResult.CsvSuccess) {
                        Write-Host "  CSV report saved at: $tempCsvPath" -ForegroundColor Cyan
                        Write-Host "  -> Close the target file and copy manually, or re-run the script." -ForegroundColor Gray
                    }
                    $exitCode = 1
                } else {
                    $exitCode = 0
                }
                
                # Return summary object
                return [PSCustomObject]@{
                    InputPath = $script:InputFile.ToString()
                    OutputPath = if ($copyResult.PptxSuccess) { $script:OutputFile } else { $tempPptxPath }
                    ReportPath = if ($copyResult.CsvSuccess) { $script:CsvReport } else { $tempCsvPath }
                    OriginalSizeBytes = $originalSize
                    OptimizedSizeBytes = $optimizedSize
                    SavingsBytes = $savedBytes
                    SavingsPercent = [Math]::Round($savedPercent, 2)
                    ImagesCropped = $croppedCount
                    ImagesOptimized = $optimizedCount
                    ImagesSkipped = $skippedCount
                    TotalImageUsages = $usages.Count
                    UniqueImages = $uniqueImages
                    ExitCode = $exitCode
                    Success = (-not $copyFailed)
                    ErrorType = 'None'
                }
                
            } finally {
                # Cleanup extraction temp directory (always)
                if (Test-Path -LiteralPath $tempDir) {
                    Remove-Item -LiteralPath $tempDir -Recurse -Force -ErrorAction SilentlyContinue
                }
                
                # Cleanup temp output files only if copy succeeded
                if (-not $copyFailed) {
                    if ($tempPptxPath -and (Test-Path -LiteralPath $tempPptxPath)) {
                        Remove-Item -LiteralPath $tempPptxPath -Force -ErrorAction SilentlyContinue
                    }
                    if ($tempCsvPath -and (Test-Path -LiteralPath $tempCsvPath)) {
                        Remove-Item -LiteralPath $tempCsvPath -Force -ErrorAction SilentlyContinue
                    }
                } elseif ($tempPptxPath -or $tempCsvPath) {
                    # If we have temp files but copy failed (or error occurred), tell user where they are
                    $hasTempPptx = $tempPptxPath -and (Test-Path -LiteralPath $tempPptxPath)
                    $hasTempCsv = $tempCsvPath -and (Test-Path -LiteralPath $tempCsvPath)
                    if ($hasTempPptx -or $hasTempCsv) {
                        Write-Host "`n[SAVE] Temporary files preserved:" -ForegroundColor Yellow
                        if ($hasTempPptx) {
                            Write-Host "  PPTX: $tempPptxPath" -ForegroundColor Cyan
                        }
                        if ($hasTempCsv) {
                            Write-Host "  CSV:  $tempCsvPath" -ForegroundColor Cyan
                        }
                    }
                }
            }
            
        } catch {
            Write-Host "`n[ERROR] Unexpected error: $_" -ForegroundColor Red
            if (Test-VerboseMode) {
                Write-Host "`nStack trace:" -ForegroundColor DarkGray
                Write-Host $_.ScriptStackTrace -ForegroundColor DarkGray
            } else {
                Write-Host "   Run with -Verbose for detailed error information." -ForegroundColor Gray
            }
            $exitCode = 1
            
            return [PSCustomObject]@{
                InputPath = $CurrentInputPath
                OutputPath = $null
                ReportPath = $null
                OriginalSizeBytes = 0
                OptimizedSizeBytes = 0
                SavingsBytes = 0
                SavingsPercent = 0.0
                ImagesCropped = 0
                ImagesOptimized = 0
                ImagesSkipped = 0
                TotalImageUsages = 0
                UniqueImages = 0
                ExitCode = $exitCode
                Success = $false
                ErrorType = 'Processing'
                Error = $_.ToString()
            }
        }
    }

    #endregion
}

process {
    $script:FileCount++
    
    # Detect batch mode
    if ($script:FileCount -gt 1 -or $MyInvocation.ExpectingInput) {
        $script:BatchMode = $true
    }
    
    if ($script:BatchMode -and $script:FileCount -gt 1) {
        Write-Host "`n" + ("=" * 60) -ForegroundColor DarkGray
    }
    
    Write-Host "`n[FILE] Processing file $($script:FileCount): $(Split-Path $InputPath -Leaf)" -ForegroundColor Cyan
    
    $result = Invoke-PptOptimization -CurrentInputPath $InputPath
    $script:BatchResults.Add($result)
}

end {
    # Output results
    if ($script:BatchMode -and $script:BatchResults.Count -gt 1) {
        Write-Host "`n" + ("=" * 60) -ForegroundColor DarkGray
        Write-Host "[STATS] BATCH SUMMARY" -ForegroundColor Cyan
        Write-Host ("=" * 60) -ForegroundColor DarkGray
        
        $totalOriginal = ($script:BatchResults | Measure-Object -Property OriginalSizeBytes -Sum).Sum
        $totalOptimized = ($script:BatchResults | Measure-Object -Property OptimizedSizeBytes -Sum).Sum
        $totalSaved = $totalOriginal - $totalOptimized
        $totalSavedPercent = if ($totalOriginal -gt 0) { ($totalSaved / $totalOriginal) * 100 } else { 0 }
        
        $successCount = @($script:BatchResults | Where-Object { $_.Success }).Count
        $failCount = @($script:BatchResults | Where-Object { -not $_.Success }).Count
        
        Write-Host "Files processed: $($script:BatchResults.Count)" -ForegroundColor Green
        Write-Host "  Successful: $successCount" -ForegroundColor Green
        if ($failCount -gt 0) {
            Write-Host "  Failed: $failCount" -ForegroundColor Red
        }
        Write-Host ""
        Write-Host "Total savings:" -ForegroundColor Green
        Write-Host "  Original: $(Format-ByteSize $totalOriginal)" -ForegroundColor Green
        Write-Host "  Optimized: $(Format-ByteSize $totalOptimized)" -ForegroundColor Green
        Write-Host "  Saved: $(Format-ByteSize $totalSaved) ($($totalSavedPercent.ToString('F1'))%)" -ForegroundColor Green
        
        # In batch mode, output results if -PassThru specified
        if ($PassThru) {
            $script:BatchResults
        }
    } else {
        # Single file mode - only output result if -PassThru specified
        if ($PassThru) {
            $result = $script:BatchResults[0]
            if ($result.ErrorType -ne 'Validation') {
                $result
            }
        }
    }
    
    # Set exit code based on results
    $finalExitCode = 0
    foreach ($result in $script:BatchResults) {
        if ($result.ExitCode -gt $finalExitCode) {
            $finalExitCode = $result.ExitCode
        }
    }
    
    # Exit with appropriate code
    # 0 = Success with changes
    # 1 = Error occurred
    # 2 = Success but no changes needed
    if ($Host.Name -ne 'ConsoleHost' -or $script:BatchMode) {
        # Don't exit in ISE or batch mode - let caller handle
    } else {
        exit $finalExitCode
    }
}