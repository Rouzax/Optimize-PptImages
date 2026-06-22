    #region Enums

    enum ContextType {
        Slide
        Layout
        Master
        Notes
        Handout
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

    class OptimizeJob {
        [string]$GroupKey
        [string]$SourcePath
        [string]$ScratchPath
        [string[]]$MagickArgs
        [string]$Operation
        [long]$BeforeSize
        [string]$NewExtension
        [string]$StatusName
    }

    #endregion
