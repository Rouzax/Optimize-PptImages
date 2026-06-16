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
