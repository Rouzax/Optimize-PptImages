#Requires -Version 7.0
Set-StrictMode -Version Latest

. $PSScriptRoot/src/Config.ps1
. $PSScriptRoot/src/Classes.ps1
. $PSScriptRoot/src/Common.ps1
. $PSScriptRoot/src/Xml.ps1
. $PSScriptRoot/src/Relationships.ps1
. $PSScriptRoot/src/ImageMagick.ps1
. $PSScriptRoot/src/Initialization.ps1
. $PSScriptRoot/src/Structure.ps1
. $PSScriptRoot/src/Scanning.ps1
. $PSScriptRoot/src/Morph.ps1
. $PSScriptRoot/src/Cropping.ps1
. $PSScriptRoot/src/Optimization.ps1
. $PSScriptRoot/src/LayoutCleanup.ps1
. $PSScriptRoot/src/Cleanup.ps1
. $PSScriptRoot/src/Validation.ps1
. $PSScriptRoot/src/Reporting.ps1
. $PSScriptRoot/src/Pipeline.ps1
. $PSScriptRoot/src/Public/Invoke-PptImageOptimization.ps1

Export-ModuleMember -Function Invoke-PptImageOptimization
