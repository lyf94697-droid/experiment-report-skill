[CmdletBinding()]
param(
  [string]$OutputDir,
  [string]$WorkDir,
  [string[]]$ExampleId,
  [switch]$Strict,
  [switch]$AllowOfficeCom
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$catalogPath = Join-Path $repoRoot "examples\template-examples\catalog.json"
$catalog = (Get-Content -LiteralPath $catalogPath -Raw -Encoding UTF8) | ConvertFrom-Json

$resolvedOutputDir = if ([string]::IsNullOrWhiteSpace($OutputDir)) {
  Join-Path $repoRoot "examples\template-examples\docx"
} else {
  [System.IO.Path]::GetFullPath($OutputDir)
}
$resolvedWorkDir = if ([string]::IsNullOrWhiteSpace($WorkDir)) {
  Join-Path $repoRoot ("tests-output\template-examples-" + (Get-Date -Format "yyyyMMdd-HHmmss"))
} else {
  [System.IO.Path]::GetFullPath($WorkDir)
}

New-Item -ItemType Directory -Path $resolvedOutputDir -Force | Out-Null
New-Item -ItemType Directory -Path $resolvedWorkDir -Force | Out-Null

$python = if (-not [string]::IsNullOrWhiteSpace($env:EXPERIMENT_REPORT_PYTHON)) {
  $env:EXPERIMENT_REPORT_PYTHON
} else {
  [string](Get-Command python -ErrorAction Stop).Source
}
$pythonDir = Split-Path -Parent $python
if (-not [string]::IsNullOrWhiteSpace($pythonDir)) {
  $env:PATH = $pythonDir + [System.IO.Path]::PathSeparator + $env:PATH
}

$selectedExamples = @()
if ($null -eq $ExampleId -or $ExampleId.Count -eq 0) {
  $selectedExamples = @($catalog.examples)
} else {
  $selectedExamples = @($catalog.examples | Where-Object { $ExampleId -contains [string]$_.id })
}
if ($selectedExamples.Count -eq 0) {
  throw "No matching template examples were found."
}

$results = New-Object System.Collections.Generic.List[object]
foreach ($example in $selectedExamples) {
  $caseOutput = Join-Path $resolvedWorkDir ([string]$example.id)
  $buildParams = @{
    BuiltInTemplateId = [string]$example.id
    ReportPath = Join-Path $repoRoot ([string]$example.report)
    MetadataPath = Join-Path $repoRoot ([string]$example.metadata)
    ReportProfileName = [string]$example.profile
    OutputDir = $caseOutput
    PipelineMode = $(if ($Strict) { "full" } else { "fast" })
    QualityMode = $(if ($Strict) { "strict" } else { "fast" })
    TemplateStyleMode = "preserve"
    FailOnFormatValidation = $true
  }
  if ($null -ne $example.PSObject.Properties["imageSpecs"]) {
    $buildParams.ImageSpecsPath = Join-Path $repoRoot ([string]$example.imageSpecs)
  }
  if ($AllowOfficeCom) {
    $buildParams.AllowOfficeCom = $true
  }

  & (Join-Path $repoRoot "scripts\build-report.ps1") @buildParams | Out-Null
  $summaryPath = Join-Path $caseOutput "summary.json"
  $summary = (Get-Content -LiteralPath $summaryPath -Raw -Encoding UTF8) | ConvertFrom-Json
  if ([string]$summary.generationStatus -ne "completed") {
    throw "Example build did not complete: $($example.id)"
  }
  if (-not [bool]$summary.formatValidationPassed) {
    throw "Example format validation failed: $($example.id)"
  }

  $sourceDocx = [string]$summary.finalDocxPath
  $polishedDocx = Join-Path $caseOutput (([string]$example.id) + ".example.docx")
  & $python (Join-Path $repoRoot "scripts\polish-template-example.py") `
    --catalog $catalogPath `
    --example-id ([string]$example.id) `
    --input $sourceDocx `
    --output $polishedDocx
  if ($LASTEXITCODE -ne 0) {
    throw "Example polish step failed: $($example.id)"
  }

  $templatePath = Join-Path $repoRoot ("examples\report-templates\" + [string]$example.id + ".docx")
  $finalValidationPath = Join-Path $caseOutput "example-format-validation.json"
  & (Join-Path $repoRoot "scripts\validate-docx-format.ps1") `
    -TemplatePath $templatePath `
    -DocxPath $polishedDocx `
    -OutFile $finalValidationPath | Out-Null
  $finalValidation = (Get-Content -LiteralPath $finalValidationPath -Raw -Encoding UTF8) | ConvertFrom-Json
  if (-not [bool]$finalValidation.passed) {
    throw "Polished example format validation failed: $($example.id)"
  }

  $finalVisualPassed = $null
  $finalPageCount = $null
  if ($Strict) {
    $finalVisualDir = Join-Path $caseOutput "final-visual"
    $finalVisualPath = Join-Path $caseOutput "final-visual-validation.json"
    $visualParams = @{
      DocxPath = $polishedDocx
      OutputDir = $finalVisualDir
      OutFile = $finalVisualPath
    }
    if ($AllowOfficeCom) {
      $visualParams.AllowOfficeCom = $true
    }
    & (Join-Path $repoRoot "scripts\run-visual-validation.ps1") @visualParams | Out-Null
    $finalVisual = (Get-Content -LiteralPath $finalVisualPath -Raw -Encoding UTF8) | ConvertFrom-Json
    $finalVisualPassed = [bool]$finalVisual.passed
    $finalPageCount = [int]$finalVisual.pageCount
    if (-not $finalVisualPassed) {
      throw "Polished example visual validation failed: $($example.id)"
    }
  }

  $destination = Join-Path $resolvedOutputDir ([string]$example.output)
  Copy-Item -LiteralPath $polishedDocx -Destination $destination -Force
  [void]$results.Add([pscustomobject]@{
      id = [string]$example.id
      topic = [string]$example.topic
      output = $destination
      generationStatus = [string]$summary.generationStatus
      formatValidationPassed = [bool]$finalValidation.passed
      visualValidationPassed = $finalVisualPassed
      pageCount = $finalPageCount
      sha256 = (Get-FileHash -LiteralPath $destination -Algorithm SHA256).Hash.ToLowerInvariant()
    })
}

$resultPath = Join-Path $resolvedWorkDir "template-examples-summary.json"
[System.IO.File]::WriteAllText(
  $resultPath,
  ($results | ConvertTo-Json -Depth 5),
  (New-Object System.Text.UTF8Encoding($true))
)
Write-Output "Template examples built: $resultPath"
