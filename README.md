# Optimize-PptImages

A cross-platform PowerShell script that intelligently optimizes PowerPoint presentations by resizing oversized images, applying physical crops, and reducing file sizes—often by 50-95% without visible quality loss.

## Table of Contents

- [Overview](#overview)
- [How PowerPoint Handles Images](#how-powerpoint-handles-images)
  - [Display Size vs Physical Size](#display-size-vs-physical-size)
  - [Cropping in PowerPoint](#cropping-in-powerpoint)
  - [Shared Images](#shared-images)
  - [Masters and Layouts](#masters-and-layouts)
- [Image Format Best Practices](#image-format-best-practices)
- [The Importance of Careful Design](#the-importance-of-careful-design)
- [What This Script Does](#what-this-script-does)
- [Requirements](#requirements)
- [Installation](#installation)
- [Usage](#usage)
  - [Basic Usage](#basic-usage)
  - [Advanced Options](#advanced-options)
  - [Parameters](#parameters)
- [Features](#features)
- [Understanding the Output](#understanding-the-output)
- [CSV Report](#csv-report)
- [Example Results](#example-results)
- [Limitations and Considerations](#limitations-and-considerations)
- [Troubleshooting](#troubleshooting)

---

## Overview

PowerPoint presentations often become bloated when users insert high-resolution images (e.g., 4000×3000px photos from cameras) and then resize them to small display sizes (e.g., 400×300px on the slide). PowerPoint stores the **full original image** in the file, even though only a fraction is visible. Similarly, when you crop an image in PowerPoint, the cropped portions remain embedded in the file.

This script addresses these issues by:

1. **Resizing images** to match their actual display size (with configurable headroom for quality)
2. **Physically cropping images** that have PowerPoint crop attributes
3. **Converting unnecessary PNGs** to JPEGs when transparency isn't truly needed
4. **Removing unreferenced media** files
5. **Protecting special features** like Morph transitions

**Typical savings:** 50-95% file size reduction, depending on how inefficiently the original presentation was constructed.

---

## How PowerPoint Handles Images

### Display Size vs Physical Size

PowerPoint's image handling can be counterintuitive:

- **Physical Size:** The actual pixel dimensions and file size of the embedded image (e.g., a 3840×2160px photo at 2 MB)
- **Display Size:** How large the image appears on the slide (e.g., scaled down to 400×225px)

When you insert a high-resolution image and resize it on your slide, PowerPoint keeps the full-resolution original embedded in the PPTX file. If you display a 3840×2160px image at 400×225px (5.7% of original size), you're storing ~95% unnecessary data.

**How PowerPoint Measures Size:**

PowerPoint internally uses **EMUs (English Metric Units)** where 914,400 EMUs = 1 inch. Since PowerPoint assumes 96 DPI for screen display:
- 1 inch = 96 pixels
- Therefore: 9,525 EMUs = 1 pixel

Display dimensions are stored in the slide XML as EMU values, which this script converts to pixels for processing.

### Cropping in PowerPoint

When you crop an image in PowerPoint's UI:

1. PowerPoint stores the **entire original image** in the file
2. Crop information is saved as XML attributes (`srcRect` with left/top/right/bottom percentages)
3. These percentages represent how much to trim from each edge
4. PowerPoint renders only the visible portion, but the file contains all the cropped-away data

**Example:** A 1000×1000px image cropped to show only the center 500×500px still embeds the full 1000×1000px image.

This script can **physically apply these crops**, removing the hidden portions and updating the XML to reflect a non-cropped image, resulting in substantial file size savings.

### Shared Images

PowerPoint's architecture allows **multiple slides to reference the same physical image file**:

- Images are stored in `ppt/media/` folder
- Each slide's relationship file (`ppt/slides/_rels/slideX.xml.rels`) references media by ID
- Multiple slides can reference the same `image5.jpg`

#### How Image Sharing Works

Understanding the difference between proper image reuse and wasteful duplication:

**❌ INEFFICIENT: Inserting the same image 3 times**

```
Slide 1: [Logo]  --------->  ppt/media/image1.png (400 KB)

Slide 2: [Logo]  --------->  ppt/media/image2.png (400 KB)

Slide 3: [Logo]  --------->  ppt/media/image3.png (400 KB)

Total storage: 1,200 KB (3 separate files)
```

**✅ EFFICIENT: Copy-pasting the image (shared reference)**

```
                              +---> ppt/media/image1.png
Slide 1: [Logo]  ------------+      (400 KB - shared file)
                              |
Slide 2: [Logo]  ------------+
                              |
Slide 3: [Logo]  ------------+

Total storage: 400 KB (1 shared file, 3 references)
Savings: 800 KB (67% reduction)
```

**The key difference:**
- **Inefficient approach** creates 3 separate 400 KB files = 1,200 KB total
- **Efficient approach** creates 1 shared 400 KB file with 3 references = 400 KB total
- **Result:** Same visual appearance, 67% less storage

#### How to Reuse Images as a User

PowerPoint **automatically shares images when you copy-paste** them within the same presentation:

**Method 1: Copy-Paste (Recommended)**
1. Insert an image on Slide 1 **at the largest size you'll need**
2. Select the image and copy it (Ctrl+C / Cmd+C)
3. Navigate to Slide 2 and paste (Ctrl+V / Cmd+V)
4. Resize smaller if needed on Slide 2
5. ✅ PowerPoint reuses the same physical image file

**Method 2: Duplicate Slides**
1. Create a slide with your image
2. Right-click the slide thumbnail → Duplicate Slide
3. ✅ The duplicated slide references the same image file

**Important: Size Strategy**
When you know you'll use an image at multiple sizes:
- ✅ **Insert at the LARGEST display size first** (e.g., 800×600px on Slide 3)
- ✅ **Then copy-paste and resize smaller** for other slides (e.g., 400×300px on Slide 5)
- ❌ **Don't insert at small size then try to enlarge** (quality degradation)

This ensures the shared image file is high enough quality for all uses. When this script optimizes, it will detect the largest display size (800×600px) and resize the source image appropriately, benefiting all instances.

**What NOT to do:**
- ❌ Don't re-insert the same image file from disk on each slide
- ❌ Don't use Insert → Pictures multiple times for the same image
- ❌ Don't insert at small size if you'll later need it larger
- ✅ Instead, copy-paste the image placeholder between slides

**Real-World Example:**
- 20 slides each showing your company logo
- **Inserted separately:** 20 copies of `logo.png` (400 KB × 20 = **8,000 KB**)
- **Copy-pasted:** 1 copy of `logo.png` (400 KB × 1 = **400 KB**)
- **Savings: 7,600 KB (95%)**

#### Optimization Implications

When this script processes shared images:
- Optimizing affects **all slides** using that image
- The script tracks the **largest display size** across all usages
- Cropped instances are handled carefully to avoid breaking other slides
- You can check the CSV report to see which images are shared and how they're used

**Verification:**
You can verify image reuse in your presentation:
1. Save your PPTX file
2. Rename extension from `.pptx` to `.zip`
3. Extract and navigate to `ppt/media/` folder
4. Count unique image files vs. total image instances
5. Fewer files = better sharing = smaller file size

### Masters and Layouts

PowerPoint uses a three-tier hierarchy:

1. **Slide Masters** – Top-level templates that define themes
2. **Slide Layouts** – Individual layout templates based on masters
3. **Slides** – Actual presentation slides based on layouts

Images can appear in any tier, and changes to masters/layouts affect multiple slides. By default, this script **skips master/layout images** to avoid unintended widespread visual changes. You can enable optimization for these with `-OptimizeMastersAndLayouts` and `-CropMastersAndLayouts`.

---

## Image Format Best Practices

Understanding when to use each format is crucial for **starting with an optimized presentation**:

### JPEG (.jpg)

**Best for:** Photographs, complex images with gradients, realistic graphics

**Characteristics:**
- Lossy compression (controlled by quality setting)
- No transparency support
- Excellent compression ratios for photographic content
- Quality 85-95 is typically imperceptible to humans

**When to use:**
- Photos from cameras or stock photography
- Screenshots without transparency
- Complex graphics with millions of colors
- Any image where you don't need transparency

### PNG (.png)

**Best for:** Graphics requiring transparency, simple diagrams, text-based images

**Characteristics:**
- Lossless compression
- Supports full alpha transparency
- Larger file sizes for complex/photographic content
- Excellent for simple graphics with few colors

**When to use:**
- Logos with transparency
- Icons with transparent backgrounds
- Diagrams with transparency
- Screenshots of UI elements with shadows/transparency
- Simple graphics with text

**When NOT to use:**
- Photographs (use JPEG instead)
- Complex images without transparency (use JPEG instead)

### SVG (.svg)

**Best for:** Vector graphics, logos, icons

**Characteristics:**
- Vector format (scales infinitely without quality loss)
- Modern PowerPoint (Office 2016+) displays SVG natively
- PowerPoint auto-generates a PNG fallback for backward compatibility
- Both the SVG and PNG fallback are stored in the presentation

**How PowerPoint handles SVG:**
- Current versions (Office 2016+) render the SVG directly (true vector display)
- Older versions fall back to the auto-generated PNG raster
- Both files are embedded: `image5.svg` and `image5.png`

**Script behavior:** This script detects auto-generated PNG fallbacks and skips optimizing them, as PowerPoint will regenerate them if needed. The SVG files themselves are not modified (vector formats don't benefit from raster optimization).

### Conversion Strategy

This script intelligently converts PNG to JPEG when:
- Transparency is below threshold (default 0.1%)
- The conversion results in smaller file size
- The image quality is maintained at specified JPEG quality (default 95%)

---

## The Importance of Careful Design

### Optimization ≠ Substitute for Good Design

**This script can save 50-95% in poorly constructed decks, but there's no substitute for careful design upfront.**

#### Start Right: Pre-Optimization Best Practices

1. **Resize images before inserting**
   - If your slide needs a 400×300px image, don't insert a 4000×3000px photo
   - Use image editing software to resize to appropriate dimensions first
   - Rule of thumb: 2× display size at 96 DPI (retina-ready)

2. **Crop images before inserting**
   - Don't rely on PowerPoint's crop tool
   - Physically crop in an image editor (Photoshop, GIMP, Paint.NET)
   - Only insert the portion you'll actually use

3. **Choose the right format**
   - Photos → JPEG (quality 85-95%)
   - Logos/icons with transparency → PNG
   - Simple diagrams → PNG or SVG
   - Never use PNG for photographs unless transparency is required

4. **Reuse images intelligently**
   - Copy-paste images between slides instead of re-inserting
   - Use slide masters for recurring elements (logos, backgrounds)
   - Leverage PowerPoint's image sharing to minimize file size

5. **Use PowerPoint's compression**
   - File → Info → Compress Pictures
   - Choose appropriate resolution (220 PPI for print, 150 PPI for screen, 96 PPI for email)
   - This is quick but less intelligent than this script

6. **Avoid "spray and pray" image insertion**
   - Don't insert 50 high-res photos and hope for the best
   - Be intentional about image selection and sizing

#### When This Script Helps Most

This script is most valuable when:
- **Inheriting presentations** from others who didn't follow best practices
- **Batch processing** multiple legacy presentations
- **Automated workflows** where manual optimization isn't feasible
- **Recovering from "quick deadline" mode** where optimization was deferred
- **Applying crops physically** that accumulate over many edits

#### What This Script Can't Fix

- **Poor visual design** – Oversized images that look good need to stay oversized
- **Intentional quality** – If you deliberately want super high-res for printing, don't optimize
- **Already-optimized content** – If images are already right-sized, there's nothing to save
- **Video file bloat** – This script only handles images, not embedded videos

#### The Golden Rule

> **Design for the medium first, optimize second.**

A well-constructed presentation with right-sized JPEGs and minimal transparency PNGs will be lean from the start. This script exists to **rescue poorly constructed files** and **automate best practices**, not to compensate for lazy workflow.

---

## What This Script Does

### Core Operations

1. **Image Resizing**
   - Analyzes display size of each image on each slide
   - Calculates target dimensions with headroom factor (default 2×)
   - Resizes only when original is significantly larger than needed
   - Uses ImageMagick for high-quality resampling

2. **Physical Cropping** (with `-EnableCropping`)
   - Detects PowerPoint crop attributes (`srcRect` in XML)
   - Physically crops the image file
   - Removes crop attributes from XML
   - Validates crop safety for Morph transitions

3. **Format Conversion**
   - Analyzes PNG files for transparency
   - Converts to JPEG when transparency < threshold (default 0.1%)
   - Ensures conversion doesn't increase file size ("no worse than before")
   - Maintains quality at specified level (default 95%)

4. **Safety Features**
   - **Morph Protection:** Detects Morph transitions and prevents breaking linked images
   - **"No Worse Than Before":** Rejects optimizations that increase file size
   - **Shared Image Awareness:** Handles images used across multiple slides
   - **Validation:** Checks critical XML files before repacking

5. **Cleanup**
   - Removes unreferenced media files
   - Updates content type definitions
   - Validates relationships

6. **Reporting**
   - Console output with emoji indicators and progress
   - Detailed verbose logging with `-Verbose`
   - CSV report with per-image statistics

---

## Requirements

### Core Requirements

- **PowerShell:** 5.1 or later (Windows) / PowerShell Core 7+ (macOS/Linux)
- **ImageMagick:** Version 7.x or later
  - Windows: Install from [imagemagick.org](https://imagemagick.org/script/download.php#windows)
  - macOS: `brew install imagemagick`
  - Linux: `sudo apt install imagemagick` or equivalent

### Platform Support

- ✅ Windows (PowerShell 5.1+, PowerShell Core 7+)
- ✅ macOS (PowerShell Core 7+)
- ✅ Linux (PowerShell Core 7+)

The script automatically detects the OS and adjusts paths accordingly.

---

## Installation

1. **Install ImageMagick**

   **Windows:**
   ```powershell
   # Download from https://imagemagick.org/script/download.php#windows
   # Or use Chocolatey:
   choco install imagemagick
   ```

   **macOS:**
   ```bash
   brew install imagemagick
   ```

   **Linux (Ubuntu/Debian):**
   ```bash
   sudo apt update
   sudo apt install imagemagick
   ```

2. **Download the Script**

   ```powershell
   # Clone the repository
   git clone https://github.com/Rouzax/Optimize-PptImages.git
   cd Optimize-PptImages

   # Or download directly
   Invoke-WebRequest -Uri "https://raw.githubusercontent.com/Rouzax/Optimize-PptImages/main/Optimize-PptImages.ps1" -OutFile "Optimize-PptImages.ps1"
   ```

3. **Verify ImageMagick Installation**

   ```powershell
   magick --version
   ```

---

## Usage

### Basic Usage

```powershell
# Optimize a presentation (resizing only)
.\Optimize-PptImages.ps1 -InputPath "MyPresentation.pptx"

# With verbose output
.\Optimize-PptImages.ps1 -InputPath "MyPresentation.pptx" -Verbose

# Specify output location
.\Optimize-PptImages.ps1 -InputPath "MyPresentation.pptx" -OutputPath "MyPresentation_Small.pptx"
```

### Advanced Options

```powershell
# Enable physical cropping for maximum savings
.\Optimize-PptImages.ps1 -InputPath "MyPresentation.pptx" -EnableCropping

# Optimize everything including masters and layouts
.\Optimize-PptImages.ps1 -InputPath "MyPresentation.pptx" `
    -EnableCropping `
    -CropMastersAndLayouts `
    -OptimizeMastersAndLayouts

# Custom quality and headroom settings
.\Optimize-PptImages.ps1 -InputPath "MyPresentation.pptx" `
    -HeadroomFactor 1.5 `
    -JpegQuality 90 `
    -TransparencyThresholdPercent 1.0

# Process only specific slides
.\Optimize-PptImages.ps1 -InputPath "MyPresentation.pptx" `
    -IncludeSlides 1,5,10 `
    -EnableCropping

# Exclude certain slides from processing
.\Optimize-PptImages.ps1 -InputPath "MyPresentation.pptx" `
    -ExcludeSlides 3,7 `
    -EnableCropping

# Require minimum 10% savings before applying optimization
.\Optimize-PptImages.ps1 -InputPath "MyPresentation.pptx" `
    -MinSavingsPercent 10

# Batch processing with pipeline input
Get-ChildItem *.pptx | .\Optimize-PptImages.ps1 -EnableCropping

# Get result object for programmatic use
$result = .\Optimize-PptImages.ps1 -InputPath "MyPresentation.pptx" -PassThru
```

### Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `InputPath` | String | *Required* | Path to the input PPTX or POTX file. Accepts pipeline input. |
| `OutputPath` | String | `<n>.optimized.<ext>` | Path for the optimized output file |
| `MagickPath` | String | Auto-detected | Path to ImageMagick `magick` executable |
| `EnableCropping` | Switch | `$false` | Apply physical crops to images with PowerPoint crop attributes |
| `CropMastersAndLayouts` | Switch | `$false` | Apply crops to images in master slides and layouts (requires `-EnableCropping`) |
| `OptimizeMastersAndLayouts` | Switch | `$false` | Resize/optimize images in master slides and layouts |
| `HeadroomFactor` | Decimal | `2.0` | Multiplier for target dimensions (2.0 = 2× display size). Range: 1.0-4.0 |
| `JpegQuality` | Int | `95` | JPEG quality for compression (1-100, higher = better quality) |
| `TransparencyThresholdPercent` | Decimal | `0.1` | Minimum transparency % to keep PNG format (convert to JPEG if below). Range: 0.0-100.0 |
| `MinSavingsPercent` | Decimal | `0.0` | Minimum savings percentage required to apply optimization. Range: 0.0-50.0 |
| `IncludeSlides` | Int[] | *None* | Array of slide numbers to process. Only these slides will be modified. All slides are still analyzed for correct image sizing and Morph detection. |
| `ExcludeSlides` | Int[] | *None* | Array of slide numbers to exclude from processing. All slides are still analyzed for correct image sizing and Morph detection. |
| `CsvReportPath` | String | `<o>.opt-report.csv` | Path for the CSV report file |
| `PassThru` | Switch | `$false` | Return the result object to the pipeline for programmatic access. |

---

## Features

### Intelligent Optimization

- **Headroom Factor:** Maintains quality by targeting 2× display size (configurable)
- **Selective Processing:** Only optimizes images that will benefit (skips already-optimal images)
- **No Worse Than Before:** Rejects optimizations that would increase file size

### Morph Transition Protection

Detects and protects PowerPoint Morph transitions:
- Identifies image pairs across consecutive Morph slides
- Prevents applying conflicting crops that would break animations
- Ensures both slides in a Morph pair remain synchronized

### Transparency Analysis

Intelligent PNG handling:
- Counts actual transparent pixels (not just alpha channel presence)
- Converts "fake transparency" PNGs to JPEG
- Respects images with genuine transparency needs
- Configurable threshold (default: 0.1% transparent pixels)

### Cross-Platform

- Automatic path handling for Windows/macOS/Linux
- Platform-aware ImageMagick detection
- UTF-8 encoding for international characters
- Case-sensitive filesystem support

### Comprehensive Reporting

**Console Output:**
```
📂 Processing file 1: Test Deck.pptx
🔎 Analyzing presentation structure...
   File: D:\Temp\PowerPoint\Test Deck.pptx
📊 Found 28 image usages
📊 Across 19 unique images
🔗 Morph detection: 2 transition(s), 2 image pair(s) linked
📝 Corrected illegal crop values for 'Picture 3' on Slide 20 (r:-12038→0, b:-182512→0)

✂️ Processing image crops...
   ✂️ Cropped 'Picture 6' on Slide 4 (saved 71,8%)
   ✂️ Cropped 'Picture 3' on Slide 5 (saved 84,1%)
   ✂️ Cropped 'Picture 3' on Slide 7 (saved 85,5%)
   🚫 Morph crop conflict for 'Picture 3' across Slide 18 ↔ 17 (skipping crop & optimization)
   ✂️ Cropped 'Picture 3' on Slide 20 (saved 52,5%)
   ✂️ Cropped 'Picture 3' on Slide 24 (saved 30,7%)
   ⏭️ Skipped crop for 'Picture 3' on Slide 26: animated GIF (would break animation)
   ✂️ Cropped 'Picture 8' on Layout (slideLayout1.xml) (saved 71,6%)
📊 Cropping phase complete: 6 cropped, 3 skipped

⚙️ Optimizing images...
   ℹ️ Effective transparency 0,0% for 'Picture 2' on Slide 1 — treating as opaque
   🔁 Converted PNG to JPEG for 'Picture 2' on Slide 1 (saved 97,7%)
   ✅ Optimized PNG with transparency for 'Picture 3' on Slide 2 (saved 91,7%)
   ℹ️ Effective transparency 0,0% for 'Picture 6' on Slide 4 — treating as opaque
   🔁 Converted PNG to JPEG for 'Picture 6' on Slide 4 (saved 93,0%)
   ℹ️ Effective transparency 0,0% for 'Picture 3' on Slide 5 — treating as opaque
   🔁 Converted PNG to JPEG for 'Picture 3' on Slide 5 (saved 89,6%)
   ✅ Optimized JPEG for 'Picture 5' on Slide 6 (saved 93,8%)
   ✅ Optimized JPEG for 'Picture 3' on Slide 7 (saved 89,1%)
   ✅ Optimized PNG with transparency for 'Picture 3' on Slide 8 (saved 71,2%)
   ℹ️ Effective transparency 0,0% for 'Picture 3' on Slide 13 — treating as opaque
   🔁 Converted PNG to JPEG for 'Picture 3' on Slide 13 (saved 98,7%)
   ✅ Optimized JPEG for 'Picture 4' on Slide 14 (saved 84,5%)
   ✅ Optimized JPEG for 'Picture 3' on Slide 15 (saved 88,7%)
   ⏭️ Skipped optimization for 'Picture 3' on Slide 17: Morph transition with crop conflict
   ℹ️ No net savings achieved for 'Picture 5' on Slide 19 — kept original
   ⏭️ Skipped 'Picture 3' on Slide 20: already at target size and quality (Q95)
   ⏭️ Skipped 'Picture 3' on Slide 21: already at target size and quality (Q95)
   ⏭️ Skipped optimization for 'Graphic 3' on Slide 22: PNG is auto-generated SVG fallback; PowerPoint regenerates
   ℹ️ Effective transparency 0,0% for 'Picture 3' on Slide 23 — treating as opaque
   ℹ️ Conversion yielded no savings for 'Picture 3' on Slide 23 — kept original
   ℹ️ Effective transparency 0,0% for 'Picture 3' on Slide 24 — treating as opaque
   🔁 Converted .gif to .jpeg for 'Picture 3' on Slide 24 (saved 2,5%)
   ⏭️ Skipped optimization for 'Picture 3' on Slide 25: Animated GIF — would break animation
   ⏭️ Skipped optimization for 'Picture 3' on Slide 26: Animated GIF — would break animation
   ℹ️ Effective transparency 0,0% for 'Picture 3' on Slide 27 — treating as opaque
   ℹ️ Conversion yielded no savings for 'Picture 3' on Slide 27 — kept original
   ℹ️ Effective transparency 0,0% for 'Picture 8' on Layout (slideLayout1.xml) — treating as opaque
   🔁 Converted PNG to JPEG for 'Picture 8' on Layout (slideLayout1.xml) (saved 98,7%)
📊 Optimization phase complete: 12 optimized, 10 skipped
💾 Saving 6 updated document(s)...

🧹 Cleaning up unreferenced media...
🧹 Media files removed: 4

📦 Repacking presentation...
✅ Repacked: 6557.56 KB

📄 Generating CSV report...
✅ CSV report generated

🏁 Optimization Complete
📊 Images cropped: 6
📊 Images optimized: 11

📉 File size:
  Original: 62,17 MB
  Optimized: 6,40 MB
  Saved: 55,76 MB (89,7%)

💾 Saving output files...
📁 Output: D:\Temp\PowerPoint\Test Deck.optimized.pptx
📊 Report: D:\Temp\PowerPoint\Test Deck.optimized.opt-report.csv
```

**CSV Report:** Detailed per-image statistics (see [CSV Report](#csv-report) section)

---

## Understanding the Output

### Exit Codes

- `0` – Success
- `1` – Error (check console output)

### Status Indicators

| Emoji | Meaning |
|-------|---------|
| 📂 | Processing file (batch mode) |
| 🔎 | Analysis phase |
| 📊 | Statistics/summary |
| 🔗 | Morph transition detection |
| ✂️ | Cropping operation |
| ⚙️ | Optimization operation |
| 🔁 | Format conversion |
| 📝 | Corrections (illegal crop values, removed no-op crop) |
| ⏭️ | Skipped operation |
| ℹ️ | Info/notes (transparency info, kept original) |
| 🚫 | Blocked operation (safety, e.g., Morph conflict) |
| 🧹 | Cleanup operation |
| 📦 | Repacking presentation |
| 📄 | Generating CSV report |
| ✅ | Success |
| 🏁 | Completion |
| 📉 | Size comparison |
| 💾 | File save operation |
| 📁 | Output file path |

### Optimization Results

**Optimized:**
- Image was resized, cropped, or converted
- File size reduced
- Changes written to output

**Skipped:**
- Already at optimal size and quality
- Used in master/layout (protection unless `-OptimizeMastersAndLayouts` is set)
- Has crop but cropping disabled
- Auto-generated PNG fallback for SVG (PowerPoint regenerates)
- Morph conflict detected
- Unsupported format (EMF, WMF vector graphics)
- No net savings achieved

**Rejected:**
- Optimization would increase file size
- Quality loss unacceptable
- Safety check failed

---

## CSV Report

Each optimization run generates a detailed CSV report with the following columns:

| Column | Description |
|--------|-------------|
| `Location` | Where the image appears (e.g., "Slide 5", "Layout (slideLayout1.xml)") |
| `ContextType` | Type of context: Slide, Layout, Master, Notes, or Handout |
| `SlideNumber` | Slide number (0 for masters/layouts) |
| `SlideModified` | Whether the slide XML was modified (Yes/No) |
| `ShapeName` | PowerPoint shape name (e.g., "Picture 3") |
| `SourceWidthPx` | Original image width in pixels |
| `SourceHeightPx` | Original image height in pixels |
| `DisplayWidthPx` | Display width on slide in pixels |
| `DisplayHeightPx` | Display height on slide in pixels |
| `TargetWidthPx` | Target width after applying headroom factor |
| `TargetHeightPx` | Target height after applying headroom factor |
| `OriginalFile` | Original media filename (e.g., "image5.png") |
| `OptimizedFile` | Filename after optimization (may differ if format changed) |
| `BeforeSizeBytes` | File size before optimization in bytes |
| `AfterSizeBytes` | File size after optimization in bytes |
| `Savings` | Human-readable savings (e.g., "443.4 KB (93.0%)") |
| `HasSrcRect` | Whether PowerPoint crop was present (True/False) |
| `CropApplied` | Whether physical crop was applied (Yes/No) |
| `CropNormalized` | Whether crop values were normalized (Yes/No) |
| `CropRemovedNoOp` | Whether a no-op crop was removed (Yes/No) |
| `SvgFallback` | Whether this is an SVG fallback image (Yes/No) |
| `OptimizationStatus` | Result status (see [Optimization Status Values](#optimization-status-values) below) |
| `WhyNotOptimized` | Reason if skipped or rejected |
| `HeadroomFactor` | Headroom factor used for this run |
| `JpegQuality` | JPEG quality setting used for this run |
| `SourceJpegQuality` | Original JPEG quality detected (for JPEG source images) |
| `TransparencyThresholdPercent` | Transparency threshold used for this run |
| `EffectiveTransparencyPercent` | Actual transparency percentage detected in image |
| `MinSavingsPercent` | Minimum savings threshold used for this run |
| `ManualActionRequired` | Whether manual intervention is recommended (Yes/No) |
| `ManualActionHint` | Description of recommended manual action |

### Optimization Status Values

| Status | Description |
|--------|-------------|
| `ConvertedPngToJpeg` | PNG converted to JPEG (no significant transparency) |
| `OptimizedPngAlpha` | PNG optimized while preserving transparency |
| `OptimizedJpeg` | JPEG resized/recompressed |
| `Skipped_SharedWithTemplatesFlagMissing` | Used in master/layout; enable `-OptimizeMastersAndLayouts` |
| `Skipped_MorphCropConflict` | Morph transition with crop conflict |
| `Skipped_SvgFallback` | PNG is auto-generated SVG fallback; PowerPoint regenerates |
| `Skipped_UnsupportedFormat` | Unsupported format (e.g., EMF vector graphic) |
| `Skipped_AlreadyOptimal` | Already at optimal quality and target size |
| `Skipped_NoNetSavings` | No net savings achieved |

**Example:**

```csv
Location,ContextType,SlideNumber,SlideModified,ShapeName,SourceWidthPx,SourceHeightPx,DisplayWidthPx,DisplayHeightPx,TargetWidthPx,TargetHeightPx,OriginalFile,OptimizedFile,BeforeSizeBytes,AfterSizeBytes,Savings,HasSrcRect,CropApplied,CropNormalized,CropRemovedNoOp,SvgFallback,OptimizationStatus,WhyNotOptimized,HeadroomFactor,JpegQuality,SourceJpegQuality,TransparencyThresholdPercent,EffectiveTransparencyPercent,MinSavingsPercent,ManualActionRequired,ManualActionHint
Slide 4,Slide,4,Yes,Picture 6,1749,1180,350,236,700,472,image4.png,image4_cropped.jpeg,13728601,271442,"12.83 MB (98.0%)",False,Yes,No,No,No,ConvertedPngToJpeg,,2,95, ,0.1,0.00,0,No,
Slide 8,Slide,8,Yes,Picture 3,2500,1535,331,203,1368,840,image8.png,image8.png,5862925,1689117,"3.98 MB (71.2%)",False,No,No,No,No,OptimizedPngAlpha,,2,95, ,0.1,13.59,0,No,
Slide 7,Slide,7,Yes,Picture 3,1583,600,274,104,548,208,image7.jpg,image7_cropped.jpg,681435,10740,"654.98 KB (98.4%)",False,Yes,No,No,No,OptimizedJpeg,,2,95,95,0.1,0.00,0,No,
```

---

## Example Results

### Conservative Optimization

```powershell
.\Optimize-PptImages.ps1 -InputPath "Test Deck.pptx"
```

**Result:**
- Original: 59.56 MB
- Optimized: 51.99 MB
- **Saved: 7.57 MB (12.7%)**
- Images cropped: 0
- Images optimized: 10

**What happened:** Only resized oversized images, skipped cropped images and master/layout images.

---

### Aggressive Optimization with Cropping

```powershell
.\Optimize-PptImages.ps1 -InputPath "Test Deck.pptx" -EnableCropping -CropMastersAndLayouts
```

**Result:**
- Original: 59.56 MB
- Optimized: 21.15 MB
- **Saved: 38.42 MB (64.5%)**
- Images cropped: 5
- Images optimized: 14

**What happened:** Resized oversized images AND applied physical crops to slides and layouts. Significant savings from removing cropped-away image data. Morph transitions with crop conflicts were automatically protected.

---

### Maximum Optimization

```powershell
.\Optimize-PptImages.ps1 -InputPath "Test Deck.pptx" `
    -EnableCropping `
    -CropMastersAndLayouts `
    -OptimizeMastersAndLayouts
```

**Result:**
- Original: 59.56 MB
- Optimized: 3.26 MB
- **Saved: 56.30 MB (94.5%)**
- Images cropped: 5
- Images optimized: 17

**What happened:** Optimized everything including master slides and layouts. Maximum possible savings for this deck.

---

### Selective Slide Processing

```powershell
.\Optimize-PptImages.ps1 -InputPath "presentation.pptx" -IncludeSlides 1,5,10 -EnableCropping
```

**What happens:** Only slides 1, 5, and 10 are modified. All slides are still analyzed for correct image sizing (shared images use max dimensions from all usages) and Morph detection.

---

## Limitations and Considerations

### What This Script Doesn't Handle

1. **Embedded Videos** – Only processes image files (.png, .jpg, .jpeg, .svg)
2. **Embedded Fonts** – Doesn't optimize font embedding
3. **SVG Vector Optimization** – SVG files remain unchanged (vector formats don't benefit from raster optimization); only auto-generated PNG fallbacks are skipped
4. **EMF/WMF Vector Graphics** – Embedded vector graphics (common in think-cell charts, Excel objects) are skipped
5. **Audio Files** – Not processed
6. **SmartArt Graphics** – Rendered as vectors, not optimized as images

### When NOT to Use This Script

- **Print-quality presentations** requiring high DPI (>220 PPI)
- **Presentations with intentional oversizing** for zoom effects
- **Files already optimized** by careful manual work
- **When you need the original high-res images** preserved for future editing

### Known Behaviors

- **SVG Fallbacks:** PowerPoint auto-generates PNG fallback versions of SVG images for backward compatibility with older Office versions. This script skips optimizing these auto-generated PNGs because PowerPoint will regenerate them as needed. The SVG files themselves are not modified.
  
- **Morph Conflicts:** Images involved in Morph transitions with different crops on consecutive slides will skip both crop and optimization to preserve animation integrity.

- **Master/Layout Protection:** By default, the script doesn't touch master or layout images to avoid widespread unintended changes. Use flags to override.

- **Shared Image Sizing:** When an image is shared across multiple slides at different display sizes, the script uses the **largest display size** to ensure no visual degradation on any slide.

- **Illegal Crop Correction:** PowerPoint can sometimes save illegal crop values (negative numbers). The script automatically corrects these to 0 and reports the correction.

- **JPEG Quality Detection:** For JPEG source images, the script detects the original quality and avoids recompression if the source is already at or below the target quality (prevents generation loss).

- **EMF/WMF Vector Graphics:** These embedded vector formats (often from think-cell or other add-ins) are skipped as they cannot be optimized as raster images.

### Backup Recommendation

**Always keep a backup of your original file.** While this script includes extensive safety checks, optimization is a lossy process (especially JPEG compression). The original high-resolution images cannot be recovered after optimization.

---

## Troubleshooting

### "ImageMagick not found"

**Solution:**
1. Install ImageMagick (see [Installation](#installation))
2. Ensure `magick` command is in your PATH
3. Restart your terminal/PowerShell session

### "Cannot open file" or "File in use"

**Solution:**
- Close PowerPoint before running the script
- Ensure the file isn't open in another program
- Check file permissions

### "Output file is in use"

The script checks if output files are accessible before processing. If the optimized PPTX or CSV report file is locked by another application (e.g., PowerPoint or Excel), the script will either fail early or preserve temporary files that you can manually copy.

**Solution:**
- Close the output file if it's open
- Check the console output for temporary file locations

### "Optimized file is corrupt"

**Possible causes:**
1. PowerPoint version incompatibility (script supports Office 2007+)
2. Unusual XML structure in original file
3. Disk space issues during processing

**Solution:**
- Try with `-Verbose` to see where it fails
- Check available disk space
- Report issue with sample file if reproducible

### Unexpected File Size Increase

If the optimized file is larger:
- The original was already well-optimized
- Images were at or below target sizes
- Run with `-Verbose` to see what was skipped

---

## Performance Notes

### Processing Time

Typical processing times (varies by hardware):
- Small deck (10 slides, 5 images): ~10-20 seconds
- Medium deck (50 slides, 30 images): ~1-2 minutes
- Large deck (100 slides, 100 images): ~3-5 minutes

### Memory Usage

- Peak memory: ~500MB-1GB depending on image sizes
- Temp disk space: Up to 2× original file size during processing
- Automatically cleans up temp files on completion or error

---

### Reporting Issues

When reporting issues, please include:
1. PowerShell version: `$PSVersionTable`
2. ImageMagick version: `magick --version`
3. Operating system
4. Sample PPTX file (if possible)
5. Full console output with `-Verbose`


---

## Acknowledgments

- **ImageMagick** for high-quality image processing
- PowerPoint Open XML format documentation

---

**Remember:** This script is a powerful tool for rescuing bloated presentations, but the best optimization is **careful design from the start**. Size your images appropriately, choose the right formats, crop before inserting, and reuse images via copy-paste. This script exists to fix mistakes and automate best practices—not to replace thoughtful design.