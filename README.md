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
- [Contributing](#contributing)

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

```
┌─────────────────────────────────────────────────────────────────────┐
│  ❌ INEFFICIENT: Inserting the same image 3 times                   │
├─────────────────────────────────────────────────────────────────────┤
│                                                                       │
│  Slide 1: [Logo]  ──────────►  ppt/media/image1.png (400 KB)       │
│                                                                       │
│  Slide 2: [Logo]  ──────────►  ppt/media/image2.png (400 KB)       │
│                                                                       │
│  Slide 3: [Logo]  ──────────►  ppt/media/image3.png (400 KB)       │
│                                                                       │
│  Total: 1,200 KB                                                     │
└─────────────────────────────────────────────────────────────────────┘

┌─────────────────────────────────────────────────────────────────────┐
│  ✅ EFFICIENT: Copy-pasting the image (shared reference)            │
├─────────────────────────────────────────────────────────────────────┤
│                                    ┌──► ppt/media/image1.png        │
│  Slide 1: [Logo]  ────────────────┤     (400 KB)                    │
│                                    │                                 │
│  Slide 2: [Logo]  ────────────────┤                                 │
│                                    │                                 │
│  Slide 3: [Logo]  ────────────────┘                                 │
│                                                                       │
│  Total: 400 KB                                                       │
│  Savings: 800 KB (67%)                                               │
└─────────────────────────────────────────────────────────────────────┘
```

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

**Script behavior:** This script detects auto-generated PNG fallbacks and skips optimizing them, as PowerPoint will regenerate them if needed. The SVG file itself is not modified (vector formats don't benefit from raster optimization).

### Conversion Strategy

This script intelligently converts PNG to JPEG when:
- Transparency is below threshold (default 0.5%)
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
   - Converts to JPEG when transparency < threshold (default 0.5%)
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
    -TransparencyThreshold 1.0
```

### Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `InputPath` | String | *Required* | Path to the input PPTX file |
| `OutputPath` | String | `<input>_optimized.pptx` | Path for the optimized output file |
| `EnableCropping` | Switch | `$false` | Apply physical crops to images with PowerPoint crop attributes |
| `CropMastersAndLayouts` | Switch | `$false` | Apply crops to images in master slides and layouts |
| `OptimizeMastersAndLayouts` | Switch | `$false` | Resize/optimize images in master slides and layouts |
| `HeadroomFactor` | Decimal | `2.0` | Multiplier for target dimensions (2.0 = 2× display size) |
| `JpegQuality` | Int | `95` | JPEG quality for compression (1-100, higher = better quality) |
| `TransparencyThreshold` | Decimal | `0.5` | Minimum transparency % to keep PNG format (convert to JPEG if below) |

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
- Configurable threshold (default: 0.5% transparent pixels)

### Cross-Platform

- Automatic path handling for Windows/macOS/Linux
- Platform-aware ImageMagick detection
- UTF-8 encoding for international characters
- Case-sensitive filesystem support

### Comprehensive Reporting

**Console Output:**
```
🔎 Analyzing presentation structure...
📊 Found 23 image usages
📊 Across 15 unique images
🔗 Morph detection: 2 transition(s), 2 image pair(s) linked

✂️ Processing image crops...
✂️ Cropped 'Picture 6' on Slide 4 (saved 71,8%)
✂️ Cropped 'Picture 3' on Slide 5 (saved 84,1%)

⚙️ Optimizing images...
⚙️ Optimized PNG with transparency for 'Picture 3' on Slide 2 (saved 92,8%)
🔁 Converted PNG to JPEG for 'Picture 6' on Slide 4 (saved 93,0%)

🏁 Optimization Complete
📊 Images cropped: 4
📊 Images optimized: 13

📉 File size:
  Original: 59,56 MB
  Optimized: 30,02 MB
  Saved: 29,55 MB (49,6%)
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
| 🔎 | Analysis phase |
| 📊 | Statistics/summary |
| 🔗 | Morph transition detection |
| ✂️ | Cropping operation |
| ⚙️ | Optimization operation |
| 🔁 | Format conversion |
| 📝 | Information/note |
| ⏭️ | Skipped operation |
| 🚫 | Blocked operation (safety) |
| 🧹 | Cleanup operation |
| ✅ | Success |
| 🏁 | Completion |
| 📉 | Size comparison |

### Optimization Results

**Optimized:**
- Image was resized, cropped, or converted
- File size reduced
- Changes written to output

**Skipped:**
- Already at optimal size
- Used in master/layout (protection)
- Has crop but cropping disabled
- Auto-generated PNG fallback for SVG (PowerPoint regenerates)
- Morph conflict detected

**Rejected:**
- Optimization would increase file size
- Quality loss unacceptable
- Safety check failed

---

## CSV Report

Each optimization run generates a detailed CSV report with the following columns:

| Column | Description |
|--------|-------------|
| `Location` | Where the image appears (e.g., "Slide 5", "Layout 1") |
| `ObjectName` | PowerPoint shape name (e.g., "Picture 3") |
| `ImageFile` | Media filename (e.g., "image5.jpg") |
| `DisplaySize` | How large the image appears on slide (e.g., "415×233px") |
| `OriginalFormat` | Original format (PNG, JPEG) |
| `OriginalSize` | Original file size (human-readable) |
| `HasCrop` | Whether PowerPoint crop was present (Yes/No) |
| `CropRegion` | Crop percentages if present (e.g., "45.6%×54.7%") |
| `HasTransparency` | Whether image has alpha channel (Yes/No) |
| `TransparencyPercent` | Percentage of transparent pixels |
| `Action` | What was done (e.g., "Cropped+Optimized", "Skipped") |
| `FinalFormat` | Format after optimization |
| `FinalSize` | File size after optimization |
| `SavingsBytes` | Bytes saved |
| `SavingsPercent` | Percentage saved (e.g., "71.8%") |
| `Reason` | Why action was taken or skipped |

**Example:**

```csv
Location,ObjectName,ImageFile,DisplaySize,OriginalFormat,OriginalSize,HasCrop,CropRegion,HasTransparency,TransparencyPercent,Action,FinalFormat,FinalSize,SavingsBytes,SavingsPercent,Reason
Slide 4,Picture 6,image4.png,350×236px,PNG,456.2 KB,Yes,45.6%×54.7%,Yes,0.0%,Cropped+Converted,JPEG,12.8 KB,443.4 KB,93.0%,Effective transparency below threshold
Slide 5,Picture 3,image5.png,270×294px,PNG,789.3 KB,Yes,27.8%×49.2%,Yes,0.0%,Cropped+Converted,JPEG,18.5 KB,770.8 KB,90.6%,Effective transparency below threshold
Slide 8,Picture 3,image8.png,684×420px,PNG,1.2 MB,No,,Yes,13.6%,Optimized,PNG,298.7 KB,901.3 KB,75.1%,PNG with transparency
Slide 19,Picture 5,image12.jpg,256×144px,JPEG,10.2 KB,No,,No,0.0%,Kept Original,JPEG,10.2 KB,0 KB,0.0%,No net savings
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

### Aggressive Optimization (Recommended)

```powershell
.\Optimize-PptImages.ps1 -InputPath "Test Deck.pptx" -EnableCropping
```

**Result:**
- Original: 59.56 MB
- Optimized: 30.02 MB
- **Saved: 29.55 MB (49.6%)**
- Images cropped: 4
- Images optimized: 13

**What happened:** Resized oversized images AND applied physical crops, significant savings from removing cropped-away image data.

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

## Limitations and Considerations

### What This Script Doesn't Handle

1. **Embedded Videos** – Only processes image files (.png, .jpg, .jpeg, .svg)
2. **Embedded Fonts** – Doesn't optimize font embedding
3. **SVG Vector Optimization** – SVG files remain unchanged (vector formats don't benefit from raster optimization); only auto-generated PNG fallbacks are skipped
4. **Audio Files** – Not processed
5. **SmartArt Graphics** – Rendered as vectors, not optimized as images

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