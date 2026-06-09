param(
  [Parameter(Mandatory = $false)]
  [switch]$IncludePass
)

$ErrorActionPreference = "Stop"
$script = Join-Path $PSScriptRoot "run_reference_equity.ps1"
& $script -IncludePass:$IncludePass
