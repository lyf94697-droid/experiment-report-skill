Set-StrictMode -Version Latest

function Resolve-UniversalReportPython {
  if (-not [string]::IsNullOrWhiteSpace($env:EXPERIMENT_REPORT_PYTHON)) {
    if (Test-Path -LiteralPath $env:EXPERIMENT_REPORT_PYTHON -PathType Leaf) {
      return [pscustomobject]@{
        command = (Resolve-Path -LiteralPath $env:EXPERIMENT_REPORT_PYTHON).Path
        prefix = @()
      }
    }
    throw "EXPERIMENT_REPORT_PYTHON points to a missing file: $env:EXPERIMENT_REPORT_PYTHON"
  }

  foreach ($candidate in @(
      @{ command = "python"; prefix = @() },
      @{ command = "py"; prefix = @("-3") }
    )) {
    $resolved = Get-Command $candidate.command -ErrorAction SilentlyContinue
    if ($null -ne $resolved) {
      return [pscustomobject]@{
        command = [string]$resolved.Source
        prefix = @($candidate.prefix)
      }
    }
  }

  throw "未找到 Python 3。请安装 Python 3，或设置 EXPERIMENT_REPORT_PYTHON 指向 python.exe。"
}

function Invoke-UniversalReportCore {
  param(
    [Parameter(Mandatory = $true)]
    [string[]]$Arguments,

    [AllowNull()]
    [string]$RepoRoot
  )

  if ([string]::IsNullOrWhiteSpace($RepoRoot)) {
    $RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
  }

  $python = Resolve-UniversalReportPython
  $oldPythonPath = $env:PYTHONPATH
  $oldPythonIoEncoding = $env:PYTHONIOENCODING
  try {
    $env:PYTHONPATH = if ([string]::IsNullOrWhiteSpace($oldPythonPath)) {
      $RepoRoot
    } else {
      $RepoRoot + [System.IO.Path]::PathSeparator + $oldPythonPath
    }
    $env:PYTHONIOENCODING = "utf-8"
    & $python.command @($python.prefix + @("-m", "universal_report") + $Arguments)
    if ($LASTEXITCODE -ne 0) {
      throw "Universal report core failed with exit code $LASTEXITCODE."
    }
  } finally {
    $env:PYTHONPATH = $oldPythonPath
    $env:PYTHONIOENCODING = $oldPythonIoEncoding
  }
}
