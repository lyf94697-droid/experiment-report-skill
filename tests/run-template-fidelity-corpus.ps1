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

function Get-PreservedEntryDigests {
  param(
    [Parameter(Mandatory = $true)]
    [string]$DocxPath
  )

  $digests = [ordered]@{}
  $archive = [System.IO.Compression.ZipFile]::OpenRead($DocxPath)
  try {
    $entries = @(
      $archive.Entries |
        Where-Object {
          $_.FullName -in @(
            "word/styles.xml",
            "word/settings.xml",
            "word/numbering.xml"
          ) -or
          $_.FullName -like "word/theme/*" -or
          $_.FullName -like "word/header*.xml" -or
          $_.FullName -like "word/footer*.xml"
        } |
        Sort-Object FullName
    )
    foreach ($entry in $entries) {
      $stream = $entry.Open()
      $sha256 = [System.Security.Cryptography.SHA256]::Create()
      try {
        $digest = $sha256.ComputeHash($stream)
        $digests[$entry.FullName] = [System.BitConverter]::ToString($digest).Replace("-", "").ToLowerInvariant()
      } finally {
        $sha256.Dispose()
        $stream.Dispose()
      }
    }
  } finally {
    $archive.Dispose()
  }
  return $digests
}

function Assert-PreservedEntries {
  param(
    [Parameter(Mandatory = $true)]
    [string]$TemplatePath,

    [Parameter(Mandatory = $true)]
    [string]$DocumentPath,

    [Parameter(Mandatory = $true)]
    [string]$CaseId
  )

  $templateEntries = Get-PreservedEntryDigests -DocxPath $TemplatePath
  $documentEntries = Get-PreservedEntryDigests -DocxPath $DocumentPath
  $templateNames = @($templateEntries.Keys)
  $documentNames = @($documentEntries.Keys)
  Assert-True -Condition (($templateNames -join "`n") -eq ($documentNames -join "`n")) -Message "$CaseId changed the preserved DOCX package-part set."
  foreach ($name in $templateNames) {
    Assert-True -Condition ([string]$templateEntries[$name] -eq [string]$documentEntries[$name]) -Message "$CaseId changed preserved package part: $name"
  }
}

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$resolvedOutputDir = if ([string]::IsNullOrWhiteSpace($OutputDir)) {
  Join-Path $repoRoot ("tests-output\template-fidelity-corpus-" + (Get-Date -Format "yyyyMMdd-HHmmss"))
} else {
  [System.IO.Path]::GetFullPath($OutputDir)
}
New-Item -ItemType Directory -Path $resolvedOutputDir -Force | Out-Null

$cases = @(
  [pscustomobject]@{
    id = "compact-header"
    template = "examples\report-templates\neutral-compact-header-lab.docx"
    profile = "experiment-report"
    report = "examples\sample-report.txt"
    metadata = "examples\docx-report-metadata.json"
  },
  [pscustomobject]@{
    id = "review-panel"
    template = "examples\report-templates\neutral-review-panel-lab.docx"
    profile = "experiment-report"
    report = "examples\sample-report.txt"
    metadata = "examples\docx-report-metadata.json"
  },
  [pscustomobject]@{
    id = "code-notebook"
    template = "examples\report-templates\neutral-code-notebook-lab.docx"
    profile = "experiment-report"
    report = "examples\sample-report.txt"
    metadata = "examples\docx-report-metadata.json"
  },
  [pscustomobject]@{
    id = "data-analysis"
    template = "examples\report-templates\neutral-data-analysis-lab.docx"
    profile = "experiment-report"
    report = "examples\sample-report.txt"
    metadata = "examples\docx-report-metadata.json"
  },
  [pscustomobject]@{
    id = "project-dossier"
    template = "examples\report-templates\neutral-project-dossier.docx"
    profile = "course-design-report"
    report = "examples\cases\course-design-student-management\report.txt"
    metadata = "examples\cases\course-design-student-management\metadata.json"
  }
)

$results = New-Object System.Collections.Generic.List[object]
$uploadedTemplatesDir = Join-Path $resolvedOutputDir "uploaded-templates"
New-Item -ItemType Directory -Path $uploadedTemplatesDir -Force | Out-Null
foreach ($case in $cases) {
  $sourceTemplatePath = Join-Path $repoRoot ([string]$case.template)
  $templatePath = Join-Path $uploadedTemplatesDir ("uploaded-" + [string]$case.id + ".docx")
  Copy-Item -LiteralPath $sourceTemplatePath -Destination $templatePath -Force
  $caseOutput = Join-Path $resolvedOutputDir ([string]$case.id)
  & (Join-Path $repoRoot "scripts\build-report.ps1") `
    -TemplatePath $templatePath `
    -ReportPath (Join-Path $repoRoot ([string]$case.report)) `
    -MetadataPath (Join-Path $repoRoot ([string]$case.metadata)) `
    -ReportProfileName ([string]$case.profile) `
    -OutputDir $caseOutput `
    -PipelineMode fast `
    -QualityMode fast `
    -TemplateStyleMode preserve | Out-Null

  $summaryPath = Join-Path $caseOutput "summary.json"
  Assert-True -Condition (Test-Path -LiteralPath $summaryPath -PathType Leaf) -Message "$($case.id) did not create summary.json."
  $summary = (Get-Content -LiteralPath $summaryPath -Raw -Encoding UTF8) | ConvertFrom-Json
  Assert-True -Condition ([string]$summary.generationStatus -eq "completed") -Message "$($case.id) did not complete."
  Assert-True -Condition ([string]$summary.templateSelectionSource -eq "user") -Message "$($case.id) was not treated as an uploaded user template."
  Assert-True -Condition ([string]::IsNullOrWhiteSpace([string]$summary.builtInTemplateId)) -Message "$($case.id) unexpectedly resolved to a built-in template ID."
  Assert-True -Condition ([bool]$summary.templateStylePreserved) -Message "$($case.id) did not use preserve mode."
  Assert-True -Condition ([bool]$summary.formatValidationPassed) -Message "$($case.id) failed template contract validation."
  Assert-True -Condition ([string]::IsNullOrWhiteSpace([string]$summary.templateFrameDocxPath)) -Message "$($case.id) unexpectedly ran legacy template-frame conversion."
  Assert-True -Condition (Test-Path -LiteralPath ([string]$summary.finalDocxPath) -PathType Leaf) -Message "$($case.id) did not create a final DOCX."

  $fieldMap = (Get-Content -LiteralPath (Join-Path $caseOutput "generated-field-map.json") -Raw -Encoding UTF8) | ConvertFrom-Json
  Assert-True -Condition ([int]$fieldMap.summary.mappedSectionCount -eq [int]$fieldMap.summary.reportSectionCount) -Message "$($case.id) did not map every report section."

  $formatValidation = (Get-Content -LiteralPath ([string]$summary.formatValidationPath) -Raw -Encoding UTF8) | ConvertFrom-Json
  $requiredChecks = @(
    "page-size",
    "page-margins",
    "page-border",
    "report-title-font",
    "report-title-size",
    "report-title-alignment",
    "body-font",
    "body-size",
    "body-line-spacing",
    "body-first-line-indent",
    "metadata-table-grid",
    "metadata-vertical-text-risk"
  )
  foreach ($checkCode in $requiredChecks) {
    $check = @($formatValidation.checks | Where-Object { [string]$_.code -eq $checkCode })
    Assert-True -Condition ($check.Count -eq 1 -and [bool]$check[0].passed) -Message "$($case.id) failed required fidelity check: $checkCode"
  }

  Assert-PreservedEntries -TemplatePath $templatePath -DocumentPath ([string]$summary.finalDocxPath) -CaseId ([string]$case.id)
  [void]$results.Add([pscustomobject]@{
      id = [string]$case.id
      templateSelectionSource = [string]$summary.templateSelectionSource
      templateStylePreserved = [bool]$summary.templateStylePreserved
      formatValidationPassed = [bool]$summary.formatValidationPassed
      mappedSectionCount = [int]$fieldMap.summary.mappedSectionCount
      reportSectionCount = [int]$fieldMap.summary.reportSectionCount
      finalDocxPath = [string]$summary.finalDocxPath
    })
}

$resultPath = Join-Path $resolvedOutputDir "template-fidelity-summary.json"
$resultPayload = [ordered]@{
  schemaVersion = "1.0"
  caseCount = $results.Count
  cases = @($results.ToArray())
}
[System.IO.File]::WriteAllText(
  $resultPath,
  ($resultPayload | ConvertTo-Json -Depth 8),
  (New-Object System.Text.UTF8Encoding($true))
)
Write-Output "Template fidelity corpus passed: $resultPath"
