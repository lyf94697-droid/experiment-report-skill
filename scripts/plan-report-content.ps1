[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)]
  [string]$CourseName,

  [Parameter(Mandatory = $true)]
  [string]$ExperimentName,

  [ValidateSet("standard", "long")]
  [string]$DetailLevel = "standard",

  [string]$VariantSeed,

  [string]$OutFile
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot "universal-report-core.ps1")

$arguments = @(
  "plan-content",
  "--course", $CourseName,
  "--experiment", $ExperimentName,
  "--detail", $DetailLevel
)
if (-not [string]::IsNullOrWhiteSpace($VariantSeed)) {
  $arguments += @("--variant-seed", $VariantSeed)
}
if (-not [string]::IsNullOrWhiteSpace($OutFile)) {
  $arguments += @("--output", [System.IO.Path]::GetFullPath($OutFile))
}

Invoke-UniversalReportCore -Arguments $arguments
