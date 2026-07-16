[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)]
  [string]$SourcePath,

  [Parameter(Mandatory = $true)]
  [string]$OutputDir,

  [switch]$AllowOfficeCom,

  [int]$TimeoutSeconds = 60
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

  foreach ($candidate in @(
      "$env:ProgramFiles\LibreOffice\program\soffice.exe",
      "${env:ProgramFiles(x86)}\LibreOffice\program\soffice.exe"
    )) {
    if (-not [string]::IsNullOrWhiteSpace($candidate) -and (Test-Path -LiteralPath $candidate -PathType Leaf)) {
      return $candidate
    }
  }

  return $null
}

function Get-OfficeProcessSnapshot {
  $snapshot = @{}
  foreach ($name in @("WINWORD", "wps", "wpp", "et")) {
    foreach ($process in @(Get-Process -Name $name -ErrorAction SilentlyContinue)) {
      $snapshot[$process.Id] = $true
    }
  }
  return $snapshot
}

function Stop-NewOfficeProcesses {
  param(
    [Parameter(Mandatory = $true)]
    [hashtable]$Before
  )

  foreach ($name in @("WINWORD", "wps", "wpp", "et")) {
    foreach ($process in @(Get-Process -Name $name -ErrorAction SilentlyContinue)) {
      if (-not $Before.ContainsKey($process.Id)) {
        Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
      }
    }
  }
}

$source = (Resolve-Path -LiteralPath $SourcePath).Path
$extension = [System.IO.Path]::GetExtension($source).ToLowerInvariant()
if ($extension -eq ".docx") {
  Write-Output ([pscustomobject]@{
      templatePath = $source
      sourceTemplatePath = $source
      status = "none"
      converter = "none"
      convertedTemplatePath = $null
    })
  exit 0
}
if ($extension -ne ".doc") {
  throw "仅支持 DOCX，或将旧版 DOC 转换为 DOCX：$source"
}

$resolvedOutputDir = [System.IO.Path]::GetFullPath($OutputDir)
New-Item -ItemType Directory -Path $resolvedOutputDir -Force | Out-Null
$convertedPath = Join-Path $resolvedOutputDir (([System.IO.Path]::GetFileNameWithoutExtension($source)) + ".docx")
if (Test-Path -LiteralPath $convertedPath) {
  Remove-Item -LiteralPath $convertedPath -Force
}

$errors = New-Object System.Collections.Generic.List[string]
$soffice = Resolve-SofficePath
if (-not [string]::IsNullOrWhiteSpace($soffice)) {
  $profileDir = Join-Path $resolvedOutputDir (".lo-profile-" + [guid]::NewGuid().ToString("N"))
  New-Item -ItemType Directory -Path $profileDir -Force | Out-Null
  try {
    $profileUri = ([System.Uri]$profileDir).AbsoluteUri
    $process = Start-Process -FilePath $soffice -ArgumentList @(
      "--headless",
      "-env:UserInstallation=$profileUri",
      "--convert-to", "docx",
      "--outdir", $resolvedOutputDir,
      $source
    ) -WindowStyle Hidden -PassThru
    if (-not $process.WaitForExit([Math]::Max(5, $TimeoutSeconds) * 1000)) {
      try { $process.Kill() } catch {}
      throw "LibreOffice 转换超过 $TimeoutSeconds 秒。"
    }
    if (Test-Path -LiteralPath $convertedPath -PathType Leaf) {
      Write-Output ([pscustomobject]@{
          templatePath = (Resolve-Path -LiteralPath $convertedPath).Path
          sourceTemplatePath = $source
          status = "converted"
          converter = "libreoffice"
          convertedTemplatePath = (Resolve-Path -LiteralPath $convertedPath).Path
        })
      exit 0
    }
    throw "LibreOffice 已退出，但没有生成 DOCX。"
  } catch {
    [void]$errors.Add("LibreOffice: $($_.Exception.Message)")
  } finally {
    if (Test-Path -LiteralPath $profileDir) {
      Remove-Item -LiteralPath $profileDir -Recurse -Force -ErrorAction SilentlyContinue
    }
  }
} else {
  [void]$errors.Add("LibreOffice: 未安装或不在 PATH 中")
}

if (-not $AllowOfficeCom) {
  throw ("旧版 DOC 转换失败。请安装 LibreOffice 后重试；如果明确允许使用 WPS/Word 自动化，请添加 -AllowOfficeCom。源文件未被修改。详情：" + ($errors -join " | "))
}

$beforeProcesses = Get-OfficeProcessSnapshot
$job = Start-Job -ScriptBlock {
  param($InputPath, $OutputPath)
  $ErrorActionPreference = "Stop"

  function Release-ComObject {
    param([AllowNull()][object]$Value)
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
      [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value)
    }
  }

  $attemptErrors = New-Object System.Collections.Generic.List[string]
  foreach ($progId in @("KWPS.Application", "Word.Application")) {
    $app = $null
    $document = $null
    try {
      $app = New-Object -ComObject $progId
      $app.Visible = $false
      try { $app.DisplayAlerts = 0 } catch {}
      $document = $app.Documents.Open($InputPath)
      if ($progId -eq "Word.Application") {
        $document.SaveAs2($OutputPath, 16)
      } else {
        $document.SaveAs($OutputPath, 16)
      }
      $document.Close($false)
      $document = $null
      $app.Quit()
      $app = $null
      if (Test-Path -LiteralPath $OutputPath -PathType Leaf) {
        return $progId
      }
      throw "$progId 未生成输出文件。"
    } catch {
      [void]$attemptErrors.Add("${progId}: $($_.Exception.Message)")
    } finally {
      if ($null -ne $document) {
        try { $document.Close($false) } catch {}
        Release-ComObject $document
      }
      if ($null -ne $app) {
        try { $app.Quit() } catch {}
        Release-ComObject $app
      }
      [GC]::Collect()
      [GC]::WaitForPendingFinalizers()
    }
  }
  throw ($attemptErrors -join " | ")
} -ArgumentList $source, $convertedPath

try {
  if (-not (Wait-Job -Job $job -Timeout ([Math]::Max(5, $TimeoutSeconds)))) {
    Stop-Job -Job $job -ErrorAction SilentlyContinue
    Stop-NewOfficeProcesses -Before $beforeProcesses
    throw "WPS/Word COM 转换超过 $TimeoutSeconds 秒，已停止并清理本次启动的 Office 进程。"
  }
  $converter = Receive-Job -Job $job -ErrorAction Stop
  if (-not (Test-Path -LiteralPath $convertedPath -PathType Leaf)) {
    throw "WPS/Word COM 已结束，但没有生成 DOCX。"
  }
  Write-Output ([pscustomobject]@{
      templatePath = (Resolve-Path -LiteralPath $convertedPath).Path
      sourceTemplatePath = $source
      status = "converted"
      converter = [string]$converter
      convertedTemplatePath = (Resolve-Path -LiteralPath $convertedPath).Path
    })
} finally {
  Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
}
