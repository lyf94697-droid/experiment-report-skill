param(
  [Parameter(Mandatory = $true)]
  [string]$Url,

  [string]$BrowserProfile = $env:OPENCLAW_BROWSER_PROFILE,

  [string]$OpenClawCmd = $env:OPENCLAW_CMD,

  [int]$MaxChars = 30000,

  [int]$TimeoutMs = 30000
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

function Parse-JsonFromOutput {
  param(
    [string]$Text
  )

  $trimmed = $Text.Trim()
  if ([string]::IsNullOrWhiteSpace($trimmed)) {
    throw "Expected JSON output, got empty text."
  }

  try {
    return $trimmed | ConvertFrom-Json
  } catch {
    $starts = @($trimmed.IndexOf("{"), $trimmed.IndexOf("["))
    $start = ($starts | Where-Object { $_ -ge 0 } | Measure-Object -Minimum).Minimum
    if ($null -eq $start) {
      throw "Failed to locate JSON in output: $trimmed"
    }
    return $trimmed.Substring([int]$start) | ConvertFrom-Json
  }
}

$cli = Resolve-OpenClawCommand -Candidate $OpenClawCmd
try {
  $uri = [Uri]$Url
} catch {
  throw "URL is not a valid URI: $Url"
}

if (-not $uri.IsAbsoluteUri -or @("http", "https") -notcontains $uri.Scheme.ToLowerInvariant()) {
  throw "URL must use http or https: $Url"
}

$statusResult = Invoke-OpenClawBrowserStatusWithRetry -Cli $cli -Profile $BrowserProfile
if (-not $statusResult.Succeeded -or $statusResult.Output -notmatch "running:\s*true") {
  & $cli browser start --browser-profile $BrowserProfile | Out-Null
  $statusResult = Invoke-OpenClawBrowserStatusWithRetry -Cli $cli -Profile $BrowserProfile
}
if (-not $statusResult.Succeeded -or $statusResult.Output -notmatch "running:\s*true") {
  throw "OpenClaw browser did not become ready after start. Last status: $($statusResult.Output)"
}

$openRaw = (& $cli browser --browser-profile $BrowserProfile open $Url --json 2>&1 | Out-String).Trim()
$opened = Parse-JsonFromOutput -Text $openRaw
$targetId = $opened.targetId
if ([string]::IsNullOrWhiteSpace($targetId)) {
  throw "Failed to open URL: $Url"
}

& $cli browser --browser-profile $BrowserProfile focus $targetId | Out-Null
try {
  & $cli browser --browser-profile $BrowserProfile wait --target-id $targetId --load networkidle --timeout-ms $TimeoutMs | Out-Null
} catch {
  Start-Sleep -Seconds 3
}

# Keep using the target id returned by browser open/focus.
# Avoid reparsing the global tabs list here because page titles from unrelated
# tabs can contain unescaped quotes and break ConvertFrom-Json in legacy shells.

$titleFn = "() => document.title || ''"
$contentFn = "() => { const selectors = ['article','main article','[role=""main""] article','main','[role=""main""]','.article-content','.article-content-box','.blog-content-box','.markdown-body','.post-content','.entry-content','#content','.content','.article','.post']; let text = ''; for (const selector of selectors) { const node = document.querySelector(selector); if (node && node.innerText && node.innerText.trim().length > 200) { text = node.innerText; break; } } if (!text) { text = document.body ? document.body.innerText : ''; } return text.replace(/\u00a0/g, ' ').replace(/\r/g, '').replace(/\n{3,}/g, '\n\n').trim().slice(0, $MaxChars); }"

$titleRaw = (& $cli browser --browser-profile $BrowserProfile evaluate --target-id $targetId --fn $titleFn 2>&1 | Out-String).Trim()
$contentRaw = (& $cli browser --browser-profile $BrowserProfile evaluate --target-id $targetId --fn $contentFn 2>&1 | Out-String).Trim()

$title = $titleRaw.Trim('"')
$content = $contentRaw.Trim('"')
$content = $content -replace "\\n", "`n"
$content = $content -replace '\\"', '"'
$content = $content -replace "\\t", "`t"

Write-Output "TITLE: $title"
Write-Output "URL: $Url"
Write-Output "TARGET: $targetId"
Write-Output ""
Write-Output $content
