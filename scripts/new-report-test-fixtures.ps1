[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)]
  [string]$OutputDir
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot "universal-report-core.ps1")

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$resolvedOutputDir = [System.IO.Path]::GetFullPath($OutputDir)
New-Item -ItemType Directory -Path $resolvedOutputDir -Force | Out-Null
$python = Resolve-UniversalReportPython
$oldPythonPath = $env:PYTHONPATH
$oldPythonIoEncoding = $env:PYTHONIOENCODING
try {
  $env:PYTHONPATH = if ([string]::IsNullOrWhiteSpace($oldPythonPath)) {
    $repoRoot
  } else {
    $repoRoot + [System.IO.Path]::PathSeparator + $oldPythonPath
  }
  $env:PYTHONIOENCODING = "utf-8"
  $code = @'
import json
import sys
from pathlib import Path
from tests.fixture_catalog import build_fixture_catalog

output_dir = Path(sys.argv[1]).resolve()
catalog = build_fixture_catalog(output_dir)
payload = {
    "schemaVersion": "1.0",
    "fixtures": [
        {
            "name": name,
            "path": str(item["path"].resolve()),
            "description": item["description"],
        }
        for name, item in catalog.items()
    ],
}
(output_dir / "fixture-manifest.json").write_text(
    json.dumps(payload, ensure_ascii=False, indent=2),
    encoding="utf-8-sig",
)
'@
  $code | & $python.command @($python.prefix + @("-", $resolvedOutputDir))
  if ($LASTEXITCODE -ne 0) {
    throw "Failed to generate report test fixtures."
  }
} finally {
  $env:PYTHONPATH = $oldPythonPath
  $env:PYTHONIOENCODING = $oldPythonIoEncoding
}

Write-Output ("Fixture manifest: {0}" -f (Join-Path $resolvedOutputDir "fixture-manifest.json"))
