[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

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

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("universal-report-core-" + [guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null

try {
  $fixturesDir = Join-Path $tempRoot "fixtures"
  & (Join-Path $repoRoot "scripts\new-report-test-fixtures.ps1") -OutputDir $fixturesDir | Out-Null
  $fixtureManifestPath = Join-Path $fixturesDir "fixture-manifest.json"
  Assert-True -Condition (Test-Path -LiteralPath $fixtureManifestPath) -Message "Fixture generator did not create fixture-manifest.json."
  $fixtureManifest = (Get-Content -LiteralPath $fixtureManifestPath -Raw -Encoding UTF8) | ConvertFrom-Json
  Assert-True -Condition (@($fixtureManifest.fixtures).Count -eq 10) -Message "Fixture generator should create exactly ten required scenarios."

  $fourColumnPath = [string](@($fixtureManifest.fixtures | Where-Object { $_.name -eq "four-column-metadata" })[0].path)
  $contractPath = Join-Path $tempRoot "template-style-contract.json"
  & (Join-Path $repoRoot "scripts\analyze-docx-template.ps1") -TemplatePath $fourColumnPath -OutFile $contractPath | Out-Null
  $contract = (Get-Content -LiteralPath $contractPath -Raw -Encoding UTF8) | ConvertFrom-Json
  Assert-True -Condition ([string]$contract.styles.roles.reportTitle.font.eastAsia -eq "黑体") -Message "Template analyzer did not preserve the inherited eastAsia title font."
  Assert-True -Condition ([int]$contract.structure.metadataTable.columnCount -eq 4) -Message "Template analyzer did not detect the four-column metadata table."

  $formatValidationPath = Join-Path $tempRoot "format-validation.json"
  & (Join-Path $repoRoot "scripts\validate-docx-format.ps1") -TemplatePath $fourColumnPath -DocxPath $fourColumnPath -OutFile $formatValidationPath | Out-Null
  $formatValidation = (Get-Content -LiteralPath $formatValidationPath -Raw -Encoding UTF8) | ConvertFrom-Json
  Assert-True -Condition ([bool]$formatValidation.passed) -Message "Format validation should pass when comparing a template with itself."

  $preserveMapPath = Join-Path $tempRoot "preserve-style-field-map.json"
  [System.IO.File]::WriteAllText(
    $preserveMapPath,
    (@{
        fieldMap = @{
          "实验目的" = "新的实验目的正文用于验证：标题段落保持标题样式，正文复用模板已有正文段落样式。"
        }
      } | ConvertTo-Json -Depth 5),
    (New-Object System.Text.UTF8Encoding($true))
  )
  $preservedDocxPath = Join-Path $tempRoot "preserved-style.docx"
  & (Join-Path $repoRoot "scripts\apply-docx-field-map.ps1") `
    -TemplatePath $fourColumnPath `
    -MappingPath $preserveMapPath `
    -OutPath $preservedDocxPath `
    -Overwrite | Out-Null
  $preservedFormatPath = Join-Path $tempRoot "preserved-format-validation.json"
  & (Join-Path $repoRoot "scripts\validate-docx-format.ps1") `
    -TemplatePath $fourColumnPath `
    -DocxPath $preservedDocxPath `
    -OutFile $preservedFormatPath | Out-Null
  $preservedFormat = (Get-Content -LiteralPath $preservedFormatPath -Raw -Encoding UTF8) | ConvertFrom-Json
  Assert-True -Condition ([bool]$preservedFormat.passed) -Message "Section body insertion should preserve the template heading and body styles."

  $legacyDocPath = [string](@($fixtureManifest.fixtures | Where-Object { $_.name -eq "legacy-doc-conversion-failure" })[0].path)
  $conversionFailed = $false
  try {
    & (Join-Path $repoRoot "scripts\convert-report-template.ps1") -SourcePath $legacyDocPath -OutputDir (Join-Path $tempRoot "converted") | Out-Null
  } catch {
    $conversionFailed = $true
    Assert-True -Condition ($_.Exception.Message -match "LibreOffice") -Message "Legacy DOC failure should recommend LibreOffice."
    Assert-True -Condition ($_.Exception.Message -match "AllowOfficeCom") -Message "Legacy DOC failure should explain explicit COM authorization."
  }
  Assert-True -Condition $conversionFailed -Message "Legacy DOC conversion should fail without LibreOffice or explicit COM authorization."

  Write-Output "Universal report core smoke tests passed."
} finally {
  if (Test-Path -LiteralPath $tempRoot) {
    Remove-Item -LiteralPath $tempRoot -Recurse -Force
  }
}
