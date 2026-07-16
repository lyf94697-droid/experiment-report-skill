[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)]
  [string]$DocxPath,

  [Parameter(Mandatory = $true)]
  [string]$PdfPath,

  [switch]$AllowOfficeCom,

  [int]$LibreOfficeTimeoutSeconds = 45
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Resolve-SofficePath {
  foreach ($commandName in @("soffice", "libreoffice")) {
    $command = Get-Command $commandName -ErrorAction SilentlyContinue
    if ($null -ne $command -and -not [string]::IsNullOrWhiteSpace([string]$command.Source)) {
      return [string]$command.Source
    }
  }

  $candidates = @(
    "$env:ProgramFiles\LibreOffice\program\soffice.exe",
    "${env:ProgramFiles(x86)}\LibreOffice\program\soffice.exe",
    "$env:ProgramFiles\OpenOffice 4\program\soffice.exe",
    "${env:ProgramFiles(x86)}\OpenOffice 4\program\soffice.exe"
  )

  foreach ($candidate in $candidates) {
    if (-not [string]::IsNullOrWhiteSpace($candidate) -and (Test-Path -LiteralPath $candidate -PathType Leaf)) {
      return $candidate
    }
  }

  return $null
}

function Release-ComObjectIfNeeded {
  param(
    [AllowNull()]
    [object]$ComObject
  )

  if ($null -eq $ComObject) {
    return
  }

  try {
    if ([System.Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
      [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObject)
    }
  } catch {
    # Best-effort cleanup only.
  }
}

$docx = (Resolve-Path -LiteralPath $DocxPath).Path
$pdf = [System.IO.Path]::GetFullPath($PdfPath)
$pdfParent = Split-Path -Parent $pdf
if (-not [string]::IsNullOrWhiteSpace($pdfParent)) {
  New-Item -ItemType Directory -Path $pdfParent -Force | Out-Null
}
if (Test-Path -LiteralPath $pdf) {
  Remove-Item -LiteralPath $pdf -Force
}

$errors = New-Object System.Collections.Generic.List[string]

$soffice = Resolve-SofficePath
if (-not [string]::IsNullOrWhiteSpace($soffice)) {
  $process = $null
  try {
    $process = Start-Process -FilePath $soffice -ArgumentList @(
      "--headless",
      "--convert-to", "pdf",
      "--outdir", $pdfParent,
      $docx
    ) -WindowStyle Hidden -PassThru

    if (-not $process.WaitForExit([Math]::Max(5, $LibreOfficeTimeoutSeconds) * 1000)) {
      try { $process.Kill() } catch {}
      throw "LibreOffice timed out after $LibreOfficeTimeoutSeconds seconds."
    }

    $converted = Join-Path $pdfParent (([System.IO.Path]::GetFileNameWithoutExtension($docx)) + ".pdf")
    if (Test-Path -LiteralPath $converted -PathType Leaf) {
      if (-not [string]::Equals($converted, $pdf, [System.StringComparison]::OrdinalIgnoreCase)) {
        Move-Item -LiteralPath $converted -Destination $pdf -Force
      }
      Write-Output "PDF exported with LibreOffice: $pdf"
      exit 0
    }

    throw "LibreOffice finished but did not create the expected PDF."
  } catch {
    [void]$errors.Add("LibreOffice: $($_.Exception.Message)")
  }
} else {
  [void]$errors.Add("LibreOffice: not found")
}

if ($AllowOfficeCom) {
  foreach ($progId in @("KWPS.Application", "Word.Application")) {
    $app = $null
    $doc = $null
    try {
      $app = New-Object -ComObject $progId
      $app.Visible = $false
      try { $app.DisplayAlerts = 0 } catch {}
      $doc = $app.Documents.Open($docx, $false, $true)
      $doc.ExportAsFixedFormat($pdf, 17)
      $doc.Close($false)
      $doc = $null
      $app.Quit()
      $app = $null
      if (Test-Path -LiteralPath $pdf -PathType Leaf) {
        Write-Output "PDF exported with ${progId}: $pdf"
        exit 0
      }
      throw "$progId finished but did not create the PDF."
    } catch {
      [void]$errors.Add("${progId}: $($_.Exception.Message)")
    } finally {
      if ($null -ne $doc) {
        try { $doc.Close($false) | Out-Null } catch {}
        Release-ComObjectIfNeeded -ComObject $doc
      }
      if ($null -ne $app) {
        try { $app.Quit() | Out-Null } catch {}
        Release-ComObjectIfNeeded -ComObject $app
      }
      [GC]::Collect()
      [GC]::WaitForPendingFinalizers()
    }
  }
} else {
  [void]$errors.Add("Office COM fallback is disabled. Set EXPERIMENT_REPORT_ALLOW_OFFICE_COM=1 to allow WPS/Word fallback.")
}

throw ("PDF export failed. DOCX has already been generated. Install LibreOffice for stable PDF export, or enable Office COM fallback after closing WPS/Word. Details: " + ($errors -join " | "))
