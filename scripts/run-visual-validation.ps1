[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)]
  [string]$DocxPath,

  [Parameter(Mandatory = $true)]
  [string]$OutputDir,

  [string]$PdfPath,

  [string]$OutFile,

  [switch]$RequireClosedFrame,

  [switch]$AllowOfficeCom
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot "universal-report-core.ps1")

$resolvedDocxPath = (Resolve-Path -LiteralPath $DocxPath).Path
$resolvedOutputDir = [System.IO.Path]::GetFullPath($OutputDir)
New-Item -ItemType Directory -Path $resolvedOutputDir -Force | Out-Null
$resolvedPdfPath = if ([string]::IsNullOrWhiteSpace($PdfPath)) {
  Join-Path $resolvedOutputDir (([System.IO.Path]::GetFileNameWithoutExtension($resolvedDocxPath)) + ".pdf")
} else {
  [System.IO.Path]::GetFullPath($PdfPath)
}
$resolvedOutFile = if ([string]::IsNullOrWhiteSpace($OutFile)) {
  Join-Path $resolvedOutputDir "visual-validation.json"
} else {
  [System.IO.Path]::GetFullPath($OutFile)
}
$previewDir = Join-Path $resolvedOutputDir "preview"

try {
  $exportParams = @{
    DocxPath = $resolvedDocxPath
    PdfPath = $resolvedPdfPath
  }
  if ($AllowOfficeCom) {
    $exportParams.AllowOfficeCom = $true
  }
  & (Join-Path $PSScriptRoot "export-docx-pdf.ps1") @exportParams | Out-Null
} catch {
  $failure = [pscustomobject]@{
    schemaVersion = "1.0"
    docxPath = $resolvedDocxPath
    pdfPath = $resolvedPdfPath
    previewDir = $previewDir
    status = "needs-fix"
    passed = $false
    pageCount = 0
    pages = @()
    errors = @(
      [pscustomobject]@{
        code = "pdf-export-failed"
        message = $_.Exception.Message
        suggestion = "安装 LibreOffice 后重试；仅在明确允许时使用 -AllowOfficeCom。DOCX 已保留。"
      }
    )
  }
  [System.IO.File]::WriteAllText($resolvedOutFile, ($failure | ConvertTo-Json -Depth 8), (New-Object System.Text.UTF8Encoding($true)))
  Write-Output $failure
  exit 0
}

$arguments = @(
  "visual-validate",
  $resolvedPdfPath,
  "--preview-dir", $previewDir,
  "--output", $resolvedOutFile
)
if ($RequireClosedFrame) {
  $arguments += "--require-closed-frame"
}
Invoke-UniversalReportCore -Arguments $arguments

$result = (Get-Content -LiteralPath $resolvedOutFile -Raw -Encoding UTF8) | ConvertFrom-Json
Write-Output $result
