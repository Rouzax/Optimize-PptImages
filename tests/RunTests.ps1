param([string]$Path = "$PSScriptRoot")
$config = New-PesterConfiguration
$config.Run.Path = $Path
$config.Output.Verbosity = 'Detailed'
Invoke-Pester -Configuration $config
