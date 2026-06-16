#Requires -Version 7.0
Set-StrictMode -Version Latest

. $PSScriptRoot/src/_monolith.ps1
. $PSScriptRoot/src/Public/Invoke-PptImageOptimization.ps1

Export-ModuleMember -Function Invoke-PptImageOptimization
