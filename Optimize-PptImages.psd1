@{
    RootModule        = 'Optimize-PptImages.psm1'
    ModuleVersion     = '1.0.0'
    GUID              = 'e3451dc6-acf1-45ee-b831-77c09ec71210'
    Author            = 'Rouzax'
    Description       = 'Optimize PowerPoint images: resize, materialize crops, convert formats, clean unused parts.'
    PowerShellVersion = '7.0'
    FunctionsToExport = @('Invoke-PptImageOptimization')
    FileList          = @('resources/sRGB.icc')
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()
}
