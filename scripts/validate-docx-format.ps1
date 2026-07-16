[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)]
  [string]$TemplatePath,

  [Parameter(Mandatory = $true)]
  [string]$DocxPath,

  [string]$OutFile
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot "universal-report-core.ps1")

$arguments = @(
  "validate-format",
  (Resolve-Path -LiteralPath $TemplatePath).Path,
  (Resolve-Path -LiteralPath $DocxPath).Path
)
if (-not [string]::IsNullOrWhiteSpace($OutFile)) {
  $arguments += @("--output", [System.IO.Path]::GetFullPath($OutFile))
}

Invoke-UniversalReportCore -Arguments $arguments
