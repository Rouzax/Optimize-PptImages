# Optimize-PptImages

A cross-platform PowerShell script that intelligently optimizes PowerPoint presentations by resizing oversized images, applying physical crops, converting inefficient formats, and reducing file sizes--often by 50-95% without visible quality loss.

## Table of Contents

- [Optimize-PptImages](#optimize-pptimages)
  - [Table of Contents](#table-of-contents)
  - [Quick Start](#quick-start)
  - [Features](#features)
    - [Core Optimization](#core-optimization)
    - [Intelligent Processing](#intelligent-processing)
    - [Batch \& Automation](#batch--automation)
    - [Output \& Reporting](#output--reporting)
  - [Requirements](#requirements)
    - [Core Requirements](#core-requirements)
    - [Platform Support](#platform-support)
  - [Installation](#installation)
    - [1. Install ImageMagick](#1-install-imagemagick)
    - [2. Clone the Repository](#2-clone-the-repository)
    - [3. Verify Installation](#3-verify-installation)
    - [Repository Layout](#repository-layout)
  - [Usage](#usage)
    - [Basic Examples](#basic-examples)
    - [Batch Processing](#batch-processing)
    - [Parameters Reference](#parameters-reference)
  - [How It Works](#how-it-works)
    - [Display Size vs Physical Size](#display-size-vs-physical-size)
    - [Cropping in PowerPoint](#cropping-in-powerpoint)
    - [Format Conversion](#format-conversion)
    - [Shared Images](#shared-images)
    - [Masters and Layouts](#masters-and-layouts)
    - [Morph Transition Protection](#morph-transition-protection)
  - [Understanding the Output](#understanding-the-output)
    - [Console Output](#console-output)
    - [Status Indicators](#status-indicators)
    - [CSV Report](#csv-report)
    - [Exit Codes](#exit-codes)
  - [Example Results](#example-results)
    - [Analysis Only (Default)](#analysis-only-default)
    - [Optimize Slides (No Cropping)](#optimize-slides-no-cropping)
    - [Optimize and Crop Slides](#optimize-and-crop-slides)
    - [Full Optimization (Including Masters/Layouts)](#full-optimization-including-masterslayouts)
  - [Best Practices](#best-practices)
    - [Before Using This Script](#before-using-this-script)
    - [When This Script Helps Most](#when-this-script-helps-most)
    - [When NOT to Use](#when-not-to-use)
  - [Limitations](#limitations)
    - [What This Script Doesn't Handle](#what-this-script-doesnt-handle)
    - [Known Behaviors](#known-behaviors)
    - [Backup Recommendation](#backup-recommendation)
  - [Troubleshooting](#troubleshooting)
    - ["ImageMagick not found"](#imagemagick-not-found)
    - ["File in use" / "Cannot open file"](#file-in-use--cannot-open-file)
    - ["Output file is in use"](#output-file-is-in-use)
    - ["Optimized file is corrupt"](#optimized-file-is-corrupt)
    - [Unexpected File Size Increase](#unexpected-file-size-increase)
  - [Performance Notes](#performance-notes)
    - [Processing Time](#processing-time)
    - [Resource Usage](#resource-usage)
    - [Reporting Issues](#reporting-issues)
  - [Acknowledgments](#acknowledgments)

---

## Quick Start

```powershell
# Analysis only (default - no changes, generates report)
.\Optimize-PptImages.ps1 -InputPath "presentation.pptx"

# Optimize slides (most common)
.\Optimize-PptImages.ps1 -InputPath "presentation.pptx" -OptimizeSlides

# Optimize with cropping
.\Optimize-PptImages.ps1 -InputPath "presentation.pptx" -OptimizeSlides -CropSlides

# Full optimization (all slides, masters, layouts, with cropping)
.\Optimize-PptImages.ps1 -InputPath "presentation.pptx" -All

# Batch process all PPTX files
Get-ChildItem *.pptx | .\Optimize-PptImages.ps1 -OptimizeSlides -CropSlides
```

---

## Features

### Core Optimization

- **Image resizing** -- Reduces oversized images to match display size with configurable headroom (default 2x)
- **Physical cropping** -- Materializes PowerPoint's soft crops into actual image files, removing hidden data
- **Format conversion** -- Converts PNG->JPEG when transparency isn't needed; converts BMP, TIFF, GIF->JPEG/PNG
- **Unreferenced media cleanup** -- Removes orphaned files from `ppt/media/`

### Intelligent Processing

- **Transparency analysis** -- Detects actual transparent pixels (not just alpha channel presence)
- **JPEG quality preservation** -- Avoids recompression when source quality is already optimal
- **Morph transition protection** -- Preserves linked images across Morph slide pairs
- **Shared image awareness** -- Uses maximum display size across all usages
- **Minimum savings threshold** -- Skip optimizations below configurable savings percentage

### Batch & Automation

- **Pipeline input** -- Process multiple files via `Get-ChildItem | .\Optimize-PptImages.ps1`
- **Batch summary** -- Aggregated statistics when processing multiple files
- **PassThru output** -- Return result objects for programmatic use
- **Selective processing** -- Include or exclude specific slides
- **WhatIf/Confirm support** -- Preview changes before applying
- **Unused layout cleanup** -- Removes slide layouts not referenced by any slide, plus masters that become completely unreferenced (opt-in via `-CleanUnusedLayouts`). Orphaned add-in parts (think-cell marker tags and OLE objects under `ppt/tags`/`ppt/embeddings`) that were referenced only by the removed layouts are automatically removed on every run once nothing references them.

### Output & Reporting

- **Detailed console output** -- Tag-based status indicators with progress
- **CSV report** -- Per-image statistics for auditing and analysis; includes think-cell discovery so you can see which images are think-cell content and which slides depend on their layouts
- **think-cell summary** -- When think-cell content is detected, a `[THINK-CELL]` console block shows total think-cell size, which layouts carry it, which slides use those layouts, and orphaned add-in parts
- **Output validation** -- Optional post-repack validation via `-ValidateOutput`; reports Open XML SDK errors as warnings without aborting the run
- **Verbose mode** -- Detailed logging with `-Verbose`

---

## Requirements

### Core Requirements

- **PowerShell:** 7.0 or later (PowerShell Core). Windows PowerShell 5.1 is no longer supported.
- **ImageMagick:** Version 7.x or later

### Optional Requirements

- **DocumentFormat.OpenXml** -- required only when using `-ValidateOutput`. Install via NuGet. Not needed for any other feature.

### Platform Support

| Platform | PowerShell Version | Status |
|----------|-------------------|--------|
| Windows | PowerShell 7.0+ | Supported |
| macOS | PowerShell 7.0+ | Supported |
| Linux | PowerShell 7.0+ | Supported |

---

## Installation

### 1. Install ImageMagick

**Windows:**
```powershell
# Via Chocolatey
choco install imagemagick

# Or download from https://imagemagick.org/script/download.php#windows
```

**macOS:**
```bash
brew install imagemagick
```

**Linux (Ubuntu/Debian):**
```bash
sudo apt update && sudo apt install imagemagick
```

### 2. Clone the Repository

The tool is a PowerShell module spread across multiple files, so cloning the repository is required.

```powershell
git clone https://github.com/Rouzax/Optimize-PptImages.git
cd Optimize-PptImages
```

**Running the tool:**

Option A -- run the entry-point script directly (same invocation as always):

```powershell
.\Optimize-PptImages.ps1 -InputPath "presentation.pptx" -All
```

Option B -- import the module and call the exported function directly (useful for scripting and automation):

```powershell
Import-Module ./Optimize-PptImages.psd1
Invoke-PptImageOptimization -InputPath "presentation.pptx" -All
```

Both options accept the same parameters and support pipeline input.

### 3. Verify Installation

```powershell
magick --version
pwsh --version
```

### Repository Layout

| Path | Purpose |
|------|---------|
| `Optimize-PptImages.ps1` | Thin entry-point script; imports the module and forwards all parameters |
| `Optimize-PptImages.psm1` | Module root; dot-sources the `src/` files |
| `Optimize-PptImages.psd1` | Module manifest |
| `src/` | Focused source files (Config, Classes, Common, ImageMagick, Xml, Relationships, Initialization, Structure, Scanning, Morph, Cropping, Optimization, LayoutCleanup, Cleanup, Validation, Reporting, Pipeline) |
| `src/Public/Invoke-PptImageOptimization.ps1` | The exported public function |
| `tests/` | Pester 5 test suite |

---

## Usage

### Basic Examples

```powershell
# Analysis only (default - generates report, no image changes)
.\Optimize-PptImages.ps1 -InputPath "MyPresentation.pptx"

# Optimize slides (resize and convert formats)
.\Optimize-PptImages.ps1 -InputPath "MyPresentation.pptx" -OptimizeSlides

# Optimize slides with cropping
.\Optimize-PptImages.ps1 -InputPath "MyPresentation.pptx" -OptimizeSlides -CropSlides

# Full optimization (everything enabled)
.\Optimize-PptImages.ps1 -InputPath "MyPresentation.pptx" -All

# Custom output location
.\Optimize-PptImages.ps1 -InputPath "input.pptx" -OutputPath "output.pptx" -OptimizeSlides

# With verbose logging
.\Optimize-PptImages.ps1 -InputPath "MyPresentation.pptx" -OptimizeSlides -CropSlides -Verbose

# Process only specific slides
.\Optimize-PptImages.ps1 -InputPath "presentation.pptx" -IncludeSlides 1,5,10 -OptimizeSlides -CropSlides

# Exclude certain slides
.\Optimize-PptImages.ps1 -InputPath "presentation.pptx" -ExcludeSlides 3,7 -OptimizeSlides

# Custom quality settings
.\Optimize-PptImages.ps1 -InputPath "presentation.pptx" -OptimizeSlides -HeadroomFactor 1.5 -JpegQuality 90

# Require minimum 10% savings to apply changes
.\Optimize-PptImages.ps1 -InputPath "presentation.pptx" -OptimizeSlides -MinSavingsPercent 10

# Remove unused layouts and masters
.\Optimize-PptImages.ps1 -InputPath "presentation.pptx" -OptimizeSlides -CropSlides -CleanUnusedLayouts
```

### Batch Processing

```powershell
# Process all PPTX files in current directory
Get-ChildItem *.pptx | .\Optimize-PptImages.ps1 -OptimizeSlides -CropSlides

# Process files matching pattern
Get-ChildItem -Path "C:\Presentations" -Filter "Report*.pptx" | .\Optimize-PptImages.ps1 -OptimizeSlides

# Get results for programmatic use
$results = Get-ChildItem *.pptx | .\Optimize-PptImages.ps1 -OptimizeSlides -PassThru
$results | Where-Object { $_.SavingsPercent -gt 50 } | Select-Object InputPath, SavingsPercent

# Preview what would be processed (WhatIf)
Get-ChildItem *.pptx | .\Optimize-PptImages.ps1 -OptimizeSlides -WhatIf
```

**Batch Output Example:**
```
============================================================
[STATS] BATCH SUMMARY
============================================================
Files processed: 5
  Successful: 5

Total savings:
  Original: 245.32 MB
  Optimized: 48.76 MB
  Saved: 196.56 MB (80.1%)
```

### Parameters Reference

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `InputPath` | String | *Required* | Path to input PPTX/POTX file. Accepts pipeline input. |
| `OutputPath` | String | `<n>.optimized.<ext>` | Path for optimized output file |
| `MagickPath` | String | Auto-detect | Path to ImageMagick `magick` executable |
| `OptimizeSlides` | Switch | `$false` | Resize and convert images in slides |
| `OptimizeMastersAndLayouts` | Switch | `$false` | Resize and convert images in master slides and layouts |
| `CropSlides` | Switch | `$false` | Apply physical crops to images in slides |
| `CropMastersAndLayouts` | Switch | `$false` | Apply physical crops to images in master slides and layouts |
| `CleanUnusedLayouts` | Switch | `$false` | Remove slide layouts not used by any slide and masters that become unreferenced. Not included in `-All`. |
| `ValidateOutput` | Switch | `$false` | Validate the repacked file with the Open XML SDK and report any errors. Non-fatal. Requires the optional DocumentFormat.OpenXml package. Not included in `-All`. |
| `All` | Switch | `$false` | Enable all optimization and cropping operations |
| `HeadroomFactor` | Decimal | `2.0` | Target resolution multiplier (0.5-4.0). 2.0 = 2x display size |
| `JpegQuality` | Int | `95` | JPEG quality for compression (1-100) |
| `TransparencyThresholdPercent` | Decimal | `0.1` | Minimum transparency % to keep PNG format (0.0-100.0) |
| `MinSavingsPercent` | Decimal | `0.0` | Minimum savings % required to apply optimization (0.0-50.0) |
| `IncludeSlides` | Int[] | *None* | Only process these slide numbers |
| `ExcludeSlides` | Int[] | *None* | Skip these slide numbers |
| `CsvReportPath` | String | `<o>.opt-report.csv` | Path for CSV report file |
| `PassThru` | Switch | `$false` | Return result object(s) to pipeline |

**Note:** When using `-IncludeSlides` or `-ExcludeSlides`, all slides are still analyzed for correct shared image sizing and Morph detection--only the modification is filtered.

---

## How It Works

### Display Size vs Physical Size

PowerPoint stores the **full original image** regardless of how you resize it on the slide:

- **Physical Size:** Actual pixel dimensions (e.g., 3840x2160px photo at 2 MB)
- **Display Size:** How large it appears on slide (e.g., scaled to 400x225px)

If you display a 3840x2160px image at 400x225px, you're storing ~95% unnecessary data.

**How PowerPoint Measures Size:**

PowerPoint uses EMUs (English Metric Units): 914,400 EMUs = 1 inch. At 96 DPI, 9,525 EMUs = 1 pixel.

### Cropping in PowerPoint

When you crop in PowerPoint's UI:
1. The **entire original image** remains embedded
2. Crop percentages are stored as XML (`srcRect` attributes)
3. Only the visible portion renders, but all data is stored

With `-CropSlides`, this script physically removes the cropped portions.

### Format Conversion

The script intelligently converts formats when beneficial:

| Source Format | Conversion | Condition |
|---------------|------------|-----------|
| PNG | -> JPEG | Transparency < threshold (default 0.1%) |
| BMP | -> JPEG/PNG | Based on transparency analysis |
| TIFF | -> JPEG/PNG | Based on transparency analysis |
| GIF (static) | -> JPEG/PNG | Based on transparency analysis |
| GIF (animated) | *Skipped* | Would break animation |

**Transparency Analysis:** The script counts actual transparent pixels, not just alpha channel presence. A PNG with an alpha channel but no actual transparency gets converted to JPEG.

### Shared Images

PowerPoint can share a single media file across multiple slides. When you copy-paste an image between slides, they reference the same file.

**Optimization Implications:**
- The script tracks the **largest display size** across all usages
- All instances benefit from a single optimized file
- Cropped shared images are handled carefully

**Best Practice:** Copy-paste images between slides (don't re-insert from disk) to leverage sharing.

### Masters and Layouts

PowerPoint's three-tier hierarchy:
1. **Slide Masters** -- Top-level templates
2. **Slide Layouts** -- Layout templates based on masters  
3. **Slides** -- Actual content based on layouts

By default, the script **skips master/layout images** to avoid widespread changes. Enable with:
- `-OptimizeMastersAndLayouts` -- Resize and convert formats
- `-CropMastersAndLayouts` -- Apply physical crops

Corporate templates often accumulate dozens or hundreds of unused layouts from merged or inherited slide masters. The `-CleanUnusedLayouts` switch removes layouts not referenced by any slide, along with any masters that have no remaining layouts. This is opt-in and not included in `-All`, since it reduces the layout gallery available when creating new slides.

### Morph Transition Protection

The script detects Morph transitions and:
- Identifies linked image pairs across consecutive slides
- Prevents conflicting crops that would break animations
- Ensures synchronized optimization of paired images

---

## Understanding the Output

### Console Output

```
[FILE] Processing file 1: Test Deck.pptx
[SCAN] Analyzing presentation structure...
   File: D:\Temp\PowerPoint\Test Deck.pptx
[STATS] Found 28 image usages
[STATS] Across 19 unique images
[LINK] Morph detection: 2 transition(s), 2 image pair(s) linked
[NOTE] Corrected illegal crop values for 'Picture 3' on Slide 20 (r:-12038->0, b:-182512->0)

[CROP] Processing image crops...
   [CROP] Cropped 'Picture 6' on Slide 4 (saved 71.8%)
   [BLOCK] Morph crop conflict for 'Picture 3' across Slide 18 <-> 17 (skipping crop & optimization)
   [SKIP] Skipped crop for 'Picture 3' on Slide 26: animated GIF (would break animation)
[STATS] Cropping phase complete: 6 cropped, 3 skipped

[PROC] Optimizing images...
   [INFO] Effective transparency 0.0% for 'Picture 2' on Slide 1 -- treating as opaque
   [CONV] Converted PNG to JPEG for 'Picture 2' on Slide 1 (saved 97.7%)
   [OK] Optimized PNG with transparency for 'Picture 3' on Slide 2 (saved 91.7%)
   [SKIP] Skipped 'Picture 3' on Slide 20: already at target size and quality (Q95)
[STATS] Optimization phase complete: 12 optimized, 10 skipped

[CLEAN] Cleaning up unreferenced media...
[CLEAN] Media files removed: 4

[PACK] Repacking presentation...
[OK] Repacked: 6557.56 KB

[DONE] Optimization Complete
[STATS] Images cropped: 6
[STATS] Images optimized: 11

[SIZE] File size:
  Original: 62.17 MB
  Optimized: 6.40 MB
  Saved: 55.76 MB (89.7%)

[SAVE] Saving output files...
[OUT] Output: D:\Temp\PowerPoint\Test Deck.optimized.pptx
[OUT] Report: D:\Temp\PowerPoint\Test Deck.optimized.opt-report.csv
```

### Status Indicators

| Tag | Meaning |
|-----|---------|
| `[FILE]` | Processing file (batch mode) |
| `[SCAN]` | Analysis phase |
| `[STATS]` | Statistics/counts |
| `[LINK]` | Morph transition detection |
| `[CROP]` | Cropping operation |
| `[PROC]` | Optimization phase |
| `[CONV]` | Format conversion (PNG->JPEG, BMP->JPEG, etc.) |
| `[NOTE]` | Corrections (illegal crop values, removed no-op crop) |
| `[SKIP]` | Skipped operation |
| `[INFO]` | Info/notes (transparency detection, kept original) |
| `[BLOCK]` | Blocked operation (Morph conflict, safety) |
| `[CLEAN]` | Cleanup operation |
| `[PACK]` | Repacking presentation |
| `[CSV]` | Generating CSV report |
| `[OK]` | Success |
| `[DONE]` | Completion |
| `[SIZE]` | File size information |
| `[SAVE]` | File save operation |
| `[OUT]` | Output file paths |
| `[WARN]` | Warning |
| `[ERROR]` | Error |

### CSV Report

Each run generates a detailed CSV report with these columns:

| Column | Description |
|--------|-------------|
| `Location` | Where the image appears (e.g., "Slide 5", "Layout (slideLayout1.xml)") |
| `ContextType` | Type: Slide, Layout, Master, Notes, or Handout |
| `SlideNumber` | Slide number (0 for masters/layouts) |
| `SlideModified` | Whether the slide XML was modified |
| `ShapeName` | PowerPoint shape name (e.g., "Picture 3") |
| `IsThinkCell` | `Yes` if the shape is think-cell content (by shape name), else `No` |
| `UsedOnSlides` | For Layout/Master rows, the slide numbers that use this layout/master (`(none)` if unused); blank for Slide rows |
| `MediaPart` | Path of the media file inside the package (e.g., `ppt/media/image16.emf`) |
| `SourceWidthPx` / `SourceHeightPx` | Original image dimensions |
| `DisplayWidthPx` / `DisplayHeightPx` | Display dimensions on slide |
| `TargetWidthPx` / `TargetHeightPx` | Target dimensions after headroom factor |
| `OriginalFile` | Original media filename |
| `OptimizedFile` | Filename after optimization (may change if format converted) |
| `BeforeSizeBytes` / `AfterSizeBytes` | File sizes |
| `Savings` | Human-readable savings (e.g., "443.4 KB (93.0%)") |
| `HasSrcRect` | Whether PowerPoint crop was present |
| `CropApplied` | Whether physical crop was applied |
| `CropNormalized` | Whether illegal crop values were corrected |
| `CropRemovedNoOp` | Whether a no-op crop (>=98% visible) was removed |
| `SvgFallback` | Whether this is an SVG fallback image |
| `OptimizationStatus` | Result status (see below) |
| `WhyNotOptimized` | Reason if skipped |
| `HeadroomFactor` | Headroom factor used |
| `JpegQuality` | JPEG quality setting used |
| `SourceJpegQuality` | Original JPEG quality (for JPEG sources) |
| `TransparencyThresholdPercent` | Transparency threshold used |
| `EffectiveTransparencyPercent` | Actual transparency detected |
| `MinSavingsPercent` | Minimum savings threshold used |
| `ManualActionRequired` | Whether manual intervention is recommended |
| `ManualActionHint` | Suggested action |

**Sort Order:** Rows are sorted by context type (Slides first, then Layouts, then Masters), then by slide number, then by layout/master number. This follows presentation order for easy reference.

**Optimization Status Values:**

| Status | Description |
|--------|-------------|
| `ConvertedPngToJpeg` | PNG converted to JPEG (no significant transparency) |
| `ConvertedToJpeg` | BMP/TIFF/GIF converted to JPEG |
| `ConvertedToPng` | TIFF/GIF converted to PNG (has transparency) |
| `OptimizedPngAlpha` | PNG optimized while preserving transparency |
| `OptimizedJpeg` | JPEG resized/recompressed |
| `NoChangeNeeded` | Already at target size |
| `Skipped_AlreadyOptimal` | Already at target size and quality |
| `Skipped_SlidesFlagMissing` | In slide; enable `-OptimizeSlides` |
| `Skipped_TemplatesFlagMissing` | In master/layout; enable `-OptimizeMastersAndLayouts` |
| `Skipped_CropNotMaterialized` | Has crop; enable `-CropSlides` or `-CropMastersAndLayouts` |
| `Skipped_MorphCropConflict` | Morph transition with conflicting crops |
| `Skipped_SvgFallback` | SVG fallback PNG; PowerPoint regenerates |
| `Skipped_UnsupportedFormat` | Unsupported format (EMF, WMF, SVG) |
| `Skipped_AnimatedGif` | Animated GIF; conversion would break animation |
| `Skipped_NoNetSavings` | No savings achieved or below threshold |

### Exit Codes

| Code | Meaning |
|------|---------|
| `0` | Success with changes applied |
| `1` | Error occurred |
| `2` | Success but no changes needed |

---

## Example Results

### Analysis Only (Default)

```powershell
.\Optimize-PptImages.ps1 -InputPath "Test Deck.pptx"
```

| Metric | Value |
|--------|-------|
| Original | 63.91 MB |
| Optimized | 63.84 MB |
| **Saved** | **75 KB (0.1%)** |
| Images cropped | 0 |
| Images optimized | 0 |

*Savings from repack only. All images reported as skipped with hints for which flags to enable.*

### Optimize Slides (No Cropping)

```powershell
.\Optimize-PptImages.ps1 -InputPath "Test Deck.pptx" -OptimizeSlides
```

| Metric | Value |
|--------|-------|
| Original | 63.91 MB |
| Optimized | 56.59 MB |
| **Saved** | **7.32 MB (11.5%)** |
| Images cropped | 0 |
| Images optimized | 5 |

*Images with crops are skipped until `-CropSlides` is enabled.*

### Optimize and Crop Slides

```powershell
.\Optimize-PptImages.ps1 -InputPath "Test Deck.pptx" -OptimizeSlides -CropSlides
```

| Metric | Value |
|--------|-------|
| Original | 63.91 MB |
| Optimized | 34.90 MB |
| **Saved** | **29.01 MB (45.4%)** |
| Images cropped | 5 |
| Images optimized | 9 |

### Full Optimization (Including Masters/Layouts)

```powershell
.\Optimize-PptImages.ps1 -InputPath "Test Deck.pptx" -All
```

| Metric | Value |
|--------|-------|
| Original | 63.91 MB |
| Optimized | 6.56 MB |
| **Saved** | **57.35 MB (89.7%)** |
| Images cropped | 6 |
| Images optimized | 12 |

---

## Best Practices

### Before Using This Script

1. **Resize images before inserting** -- Don't insert 4000x3000px photos for 400x300px display
2. **Crop in image editor** -- Don't rely on PowerPoint's soft crop
3. **Choose right format** -- Photos->JPEG, transparency->PNG, vectors->SVG
4. **Copy-paste images** -- Don't re-insert same image from disk (enables sharing)

### When This Script Helps Most

- Inheriting poorly-optimized presentations from others
- Batch processing legacy presentations
- Automated workflows where manual optimization isn't feasible
- Recovering from "quick deadline" work
- Applying accumulated soft crops

### When NOT to Use

- Print-quality presentations requiring high DPI (>220 PPI)
- Intentional oversizing for zoom effects
- Already well-optimized files
- When you need original high-res images preserved

---

## Limitations

### What This Script Doesn't Handle

| Item | Reason |
|------|--------|
| Embedded videos | Only processes images |
| Embedded fonts | Not optimizable |
| SVG files | Vector format (no raster optimization) |
| EMF/WMF graphics | Vector formats (common in think-cell, Excel objects) |
| Audio files | Not processed |
| SmartArt | Rendered as vectors |

### Known Behaviors

- **SVG Fallbacks:** PowerPoint auto-generates PNG fallbacks for SVGs. The script skips these as PowerPoint regenerates them.
- **Animated GIFs:** Skipped to preserve animation.
- **Shared Image Sizing:** Uses largest display size across all usages.
- **Illegal Crop Correction:** Automatically corrects invalid (negative) crop values.
- **JPEG Quality Detection:** Avoids recompression if source quality <= target.

### Backup Recommendation

**Always keep a backup.** Optimization is lossy (especially JPEG compression). Original high-resolution images cannot be recovered.

---

## Troubleshooting

### "ImageMagick not found"

1. Install ImageMagick (see [Installation](#installation))
2. Ensure `magick` is in your PATH
3. Restart terminal/PowerShell
4. Or use `-MagickPath "C:\Path\To\magick.exe"`

### "File in use" / "Cannot open file"

- Close PowerPoint before running
- Ensure file isn't open elsewhere
- Check file permissions

### "Output file is in use"

The script checks file accessibility before processing. If locked:
- Close the output file
- Check console for temporary file locations

### "Optimized file is corrupt"

Possible causes:
1. Unusual XML structure in original
2. Disk space issues
3. PowerPoint version incompatibility

Solutions:
- Run with `-Verbose` for details
- Check available disk space
- Report issue with sample file

### Unexpected File Size Increase

If optimized file is larger:
- Original was already well-optimized
- Images were at/below target sizes
- Run with `-Verbose` to see what was skipped

---

## Performance Notes

### Processing Time

| Deck Size | Typical Time |
|-----------|--------------|
| Small (10 slides, 5 images) | 10-20 seconds |
| Medium (50 slides, 30 images) | 1-2 minutes |
| Large (100 slides, 100 images) | 3-5 minutes |

### Resource Usage

- **Peak memory:** ~500MB-1GB depending on image sizes
- **Temp disk space:** Up to 2x original file size
- **Cleanup:** Automatic on completion or error

### Reporting Issues

Include:
1. PowerShell version: `$PSVersionTable`
2. ImageMagick version: `magick --version`
3. Operating system
4. Sample PPTX (if possible)
5. Full console output with `-Verbose`

---

## Acknowledgments

- **ImageMagick** for high-quality image processing
- PowerPoint Open XML format documentation

---

**Remember:** This script rescues bloated presentations and automates best practices--but the best optimization is careful design from the start. Size images appropriately, choose the right formats, crop before inserting, and reuse via copy-paste.