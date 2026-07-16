param(
  [string]$BrowserProfile = $env:OPENCLAW_BROWSER_PROFILE,
  [string]$OpenClawCmd = $env:OPENCLAW_CMD
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($BrowserProfile)) {
  $BrowserProfile = "openclaw"
}

function Invoke-OpenClawBrowserStatusWithRetry {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Cli,

    [Parameter(Mandatory = $true)]
    [string]$Profile,

    [int]$Attempts = 3,

    [int]$DelaySeconds = 2
  )

  $lastOutput = ""
  $lastExitCode = 0

  for ($attempt = 1; $attempt -le $Attempts; $attempt++) {
    $output = (& $Cli browser status --browser-profile $Profile 2>&1 | Out-String).Trim()
    $exitCode = $LASTEXITCODE

    if ($exitCode -eq 0) {
      return [pscustomobject]@{
        Succeeded = $true
        Output = $output
        ExitCode = $exitCode
        Attempts = $attempt
      }
    }

    $lastOutput = $output
    $lastExitCode = $exitCode
    if ($attempt -lt $Attempts) {
      Start-Sleep -Seconds $DelaySeconds
    }
  }

  return [pscustomobject]@{
    Succeeded = $false
    Output = $lastOutput
    ExitCode = $lastExitCode
    Attempts = $Attempts
  }
}

function Resolve-OpenClawCommand {
  param(
    [string]$Candidate
  )

  if (-not [string]::IsNullOrWhiteSpace($Candidate)) {
    if (Test-Path $Candidate) {
      return (Resolve-Path $Candidate).Path
    }
    throw "OPENCLAW_CMD does not exist: $Candidate"
  }

  foreach ($name in @("openclaw.cmd", "openclaw")) {
    $cmd = Get-Command $name -ErrorAction SilentlyContinue
    if ($null -ne $cmd -and $cmd.Source) {
      return $cmd.Source
    }
  }

  throw "OpenClaw CLI not found. Set OPENCLAW_CMD or add openclaw.cmd to PATH."
}

$cli = Resolve-OpenClawCommand -Candidate $OpenClawCmd

Write-Output "OpenClaw CLI: $cli"
Write-Output "Browser profile: $BrowserProfile"
Write-Output ""
Write-Output "browser status:"
$statusResult = Invoke-OpenClawBrowserStatusWithRetry -Cli $cli -Profile $BrowserProfile
if (-not [string]::IsNullOrWhiteSpace($statusResult.Output)) {
  Write-Output $statusResult.Output
}
if (-not $statusResult.Succeeded) {
  throw "browser status failed after $($statusResult.Attempts) attempts."
}
