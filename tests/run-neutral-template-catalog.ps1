[CmdletBinding()]
param(
  [string]$OutputDir
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Assert-True {
  param(
    [Parameter(Mandatory = $true)]
    [bool]$Condition,

    [Parameter(Mandatory = $true)]
    [string]$Message
  )

  if (-not $Condition) {
    throw $Message
  }
}

function Get-DocxEntryXml {
  param(
    [Parameter(Mandatory = $true)]
    [string]$DocxPath,

    [Parameter(Mandatory = $true)]
    [string]$EntryName
  )

  $archive = [System.IO.Compression.ZipFile]::OpenRead($DocxPath)
  try {
    $entry = $archive.GetEntry($EntryName)
    if ($null -eq $entry) {
      return $null
    }
    $stream = $entry.Open()
    $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8, $true)
    try {
      [xml]$xmlDocument = $reader.ReadToEnd()
      Write-Output -NoEnumerate $xmlDocument
      return
    } finally {
      $reader.Dispose()
      $stream.Dispose()
    }
  } finally {
    $archive.Dispose()
  }
}

function Get-DocxEntryNames {
  param(
    [Parameter(Mandatory = $true)]
    [string]$DocxPath,

    [Parameter(Mandatory = $true)]
    [string]$Pattern
  )

  $archive = [System.IO.Compression.ZipFile]::OpenRead($DocxPath)
  try {
    return @($archive.Entries | Where-Object { $_.FullName -like $Pattern } | ForEach-Object { $_.FullName })
  } finally {
    $archive.Dispose()
  }
}

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$python = if (-not [string]::IsNullOrWhiteSpace($env:EXPERIMENT_REPORT_PYTHON)) {
  $env:EXPERIMENT_REPORT_PYTHON
} else {
  [string](Get-Command python -ErrorAction Stop).Source
}
$resolvedOutputDir = if ([string]::IsNullOrWhiteSpace($OutputDir)) {
  Join-Path $repoRoot ("tests-output\neutral-template-catalog-" + (Get-Date -Format "yyyyMMdd-HHmmss"))
} else {
  [System.IO.Path]::GetFullPath($OutputDir)
}
New-Item -ItemType Directory -Path $resolvedOutputDir -Force | Out-Null

$auditPath = Join-Path $resolvedOutputDir "template-catalog-audit.json"
& $python -m universal_report audit-template-catalog --repo-root $repoRoot --output $auditPath
if ($LASTEXITCODE -ne 0) {
  throw "Built-in template catalog audit failed. See $auditPath"
}
$audit = (Get-Content -LiteralPath $auditPath -Raw -Encoding UTF8) | ConvertFrom-Json
Assert-True -Condition ([bool]$audit.passed) -Message "All ten built-in templates must pass identity and provenance checks."
Assert-True -Condition (@($audit.templates).Count -eq 10) -Message "The built-in template catalog must contain exactly ten templates."

$exampleCatalogPath = Join-Path $repoRoot "examples\template-examples\catalog.json"
$exampleCatalog = (Get-Content -LiteralPath $exampleCatalogPath -Raw -Encoding UTF8) | ConvertFrom-Json
$cases = @(
  foreach ($example in @($exampleCatalog.examples)) {
    $case = [ordered]@{
      id = [string]$example.id
      profile = [string]$example.profile
      report = [string]$example.report
      metadata = [string]$example.metadata
    }
    if ($null -ne $example.PSObject.Properties["imageSpecs"]) {
      $case.imageSpecs = [string]$example.imageSpecs
    }
    [pscustomobject]$case
  }
)

$results = New-Object System.Collections.Generic.List[object]
foreach ($case in $cases) {
  $caseOutput = Join-Path $resolvedOutputDir ([string]$case.id)
  $buildParams = @{
    BuiltInTemplateId = [string]$case.id
    ReportPath = Join-Path $repoRoot ([string]$case.report)
    MetadataPath = Join-Path $repoRoot ([string]$case.metadata)
    ReportProfileName = [string]$case.profile
    OutputDir = $caseOutput
    PipelineMode = "fast"
    QualityMode = "fast"
    TemplateStyleMode = "preserve"
  }
  if ($null -ne $case.PSObject.Properties["imageSpecs"]) {
    $buildParams.ImageSpecsPath = Join-Path $repoRoot ([string]$case.imageSpecs)
  }
  & (Join-Path $repoRoot "scripts\build-report.ps1") @buildParams | Out-Null

  $summaryPath = Join-Path $caseOutput "summary.json"
  Assert-True -Condition (Test-Path -LiteralPath $summaryPath) -Message "Missing summary for built-in template: $($case.id)"
  $summary = (Get-Content -LiteralPath $summaryPath -Raw -Encoding UTF8) | ConvertFrom-Json
  Assert-True -Condition ([string]$summary.builtInTemplateId -eq [string]$case.id) -Message "Resolved the wrong built-in template for: $($case.id)"
  Assert-True -Condition ([string]$summary.generationStatus -eq "completed") -Message "Built-in template did not complete: $($case.id)"
  Assert-True -Condition ([bool]$summary.templateStylePreserved) -Message "Built-in template style was not preserved: $($case.id)"
  Assert-True -Condition ([bool]$summary.formatValidationPassed) -Message "Format validation failed for: $($case.id)"
  Assert-True -Condition ([string]::IsNullOrWhiteSpace([string]$summary.templateFrameDocxPath)) -Message "Legacy template-frame conversion should not run for: $($case.id)"
  Assert-True -Condition (Test-Path -LiteralPath ([string]$summary.finalDocxPath) -PathType Leaf) -Message "Final DOCX was not created for: $($case.id)"
  $fieldMap = (Get-Content -LiteralPath (Join-Path $caseOutput "generated-field-map.json") -Raw -Encoding UTF8) | ConvertFrom-Json
  Assert-True -Condition ([int]$fieldMap.summary.mappedSectionCount -eq [int]$fieldMap.summary.reportSectionCount) -Message "Built-in template did not map every report section: $($case.id)"
  if ([string]$case.id -eq "neutral-course-design") {
    Assert-True -Condition ([bool]$summary.layoutCheckPassed) -Message "Course-design built-in template failed layout validation."
    $layoutCheck = (Get-Content -LiteralPath ([string]$summary.layoutCheckPath) -Raw -Encoding UTF8) | ConvertFrom-Json
    Assert-True -Condition ([int]$layoutCheck.actual.invalidTableChildParagraphCount -eq 0) -Message "Course-design image insertion produced invalid paragraphs directly under a table."
    Assert-True -Condition ([int]$layoutCheck.actual.imageDrawingCount -eq 6) -Message "Course-design image catalog should render all six requested images."
    Assert-True -Condition ([int]$layoutCheck.actual.captionCount -eq 6) -Message "Course-design image catalog should render all six captions."
    $finalDocxPath = [string]$summary.finalDocxPath
    $documentXml = Get-DocxEntryXml -DocxPath $finalDocxPath -EntryName "word/document.xml"
    $documentText = [string]$documentXml.DocumentElement.InnerText
    $tablesMatch = [regex]::Match($documentText, "\d+\.\d+\s*\u529F\u80FD\u6A21\u5757\u8BBE\u8BA1")
    $resultsMatch = [regex]::Match($documentText, "\u4E03\u3001\u5B9E\u73B0\u7ED3\u679C")
    $tablesIndex = if ($tablesMatch.Success) { $tablesMatch.Index } else { -1 }
    $resultsIndex = if ($resultsMatch.Success) { $resultsMatch.Index } else { -1 }
    Assert-True -Condition ($tablesIndex -ge 0 -and $resultsIndex -ge 0 -and $tablesIndex -lt $resultsIndex) -Message "Course-design tables must stay inside the design section before implementation results."

    foreach ($entryName in @(Get-DocxEntryNames -DocxPath $finalDocxPath -Pattern "word/header*.xml")) {
      $headerText = [string](Get-DocxEntryXml -DocxPath $finalDocxPath -EntryName $entryName).DocumentElement.InnerText
      $headerCount = ([regex]::Matches($headerText, "\u8BFE\u7A0B\u8BBE\u8BA1\u62A5\u544A")).Count
      Assert-True -Condition ($headerCount -le 1) -Message "Course-design header text was duplicated in $entryName."
    }
    foreach ($entryName in @(Get-DocxEntryNames -DocxPath $finalDocxPath -Pattern "word/footer*.xml")) {
      $footerText = [string](Get-DocxEntryXml -DocxPath $finalDocxPath -EntryName $entryName).DocumentElement.InnerText
      $pagePrefixCount = ([regex]::Matches($footerText, "\u7B2C\s*")).Count
      Assert-True -Condition ($pagePrefixCount -le 1) -Message "Course-design page number was duplicated in $entryName."
    }
  }

  [void]$results.Add([pscustomobject]@{
      id = [string]$case.id
      generationStatus = [string]$summary.generationStatus
      mappedSectionCount = [int]$fieldMap.summary.mappedSectionCount
      reportSectionCount = [int]$fieldMap.summary.reportSectionCount
      formatValidationPassed = [bool]$summary.formatValidationPassed
      layoutCheckPassed = $summary.layoutCheckPassed
      finalDocxPath = [string]$summary.finalDocxPath
    })
}

$resultPath = Join-Path $resolvedOutputDir "ten-template-summary.json"
[System.IO.File]::WriteAllText(
  $resultPath,
  ($results | ConvertTo-Json -Depth 5),
  (New-Object System.Text.UTF8Encoding($true))
)
Write-Output "Neutral template catalog tests passed: $resultPath"
