[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)]
  [string]$TemplatePath,

  [string]$CacheDir,

  [switch]$Verified,

  [string]$OutFile
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot "universal-report-core.ps1")

$resolvedTemplatePath = (Resolve-Path -LiteralPath $TemplatePath).Path
$arguments = @("analyze-template", $resolvedTemplatePath)
if (-not [string]::IsNullOrWhiteSpace($CacheDir)) {
  $arguments += @("--cache-dir", [System.IO.Path]::GetFullPath($CacheDir))
}
if ($Verified) {
  $arguments += "--verified"
}
if (-not [string]::IsNullOrWhiteSpace($OutFile)) {
  $arguments += @("--output", [System.IO.Path]::GetFullPath($OutFile))
}

Invoke-UniversalReportCore -Arguments $arguments
