[CmdletBinding()]
param(
  [string]$TemplatePath,

  [string]$BuiltInTemplateId,

  [Parameter(Mandatory = $true)]
  [string]$ReportPath,

  [string]$MetadataPath,

  [string]$MetadataJson,

  [string]$ImageSpecsPath,

  [string]$ImageSpecsJson,

  [string[]]$ImagePaths,

  [int]$RequestedImageCount = -1,

  [string]$ReportProfileName = "experiment-report",

  [string]$ReportProfilePath,

  [string]$RequirementsPath,

  [string]$RequirementsJson,

  [string]$OutputDir,

  [string]$FieldMapOutPath,

  [string]$FilledDocxOutPath,

  [string]$ImagePlanOutPath,

  [string]$ImageMapOutPath,

  [string]$FilledDocxWithImagesOutPath,

  [string]$StyledDocxOutPath,

  [string]$TemplateFrameDocxOutPath,

  [switch]$StyleFinalDocx,

  [switch]$CreateTemplateFrameDocx,

  [ValidateSet("fast", "full")]
  [string]$PipelineMode = "fast",

  [ValidateSet("fast", "strict")]
  [string]$QualityMode = "fast",

  [ValidateSet("standard", "long")]
  [string]$DetailLevel = "standard",

  [ValidateSet("preserve", "normalize")]
  [string]$TemplateStyleMode = "preserve",

  [string]$TemplateCacheDir,

  [switch]$AllowOfficeCom,

  [switch]$FailOnFormatValidation,

  [ValidateSet("auto", "default", "compact", "school", "excellent")]
  [string]$StyleProfile = "auto",

  [string]$StyleProfilePath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot "report-defaults.ps1")
. (Join-Path $PSScriptRoot "report-profiles.ps1")
. (Join-Path $PSScriptRoot "universal-report-core.ps1")

function Ensure-ParentDirectory {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Path
  )

  $parent = Split-Path -Parent $Path
  if (-not [string]::IsNullOrWhiteSpace($parent)) {
    New-Item -ItemType Directory -Path $parent -Force | Out-Null
  }
}

function Add-PipelineStage {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Name,

    [ValidateSet("completed", "failed", "skipped")]
    [string]$Status = "completed",

    [AllowNull()]
    [object]$Output,

    [AllowNull()]
    [object]$Error
  )

  if ($null -eq $script:pipelineStages) {
    $script:pipelineStages = New-Object System.Collections.Generic.List[object]
  }

  $timestamp = (Get-Date).ToString("s")
  [void]$script:pipelineStages.Add([pscustomobject]@{
      name = $Name
      status = $Status
      startedAt = $timestamp
      finishedAt = $timestamp
      output = $Output
      error = $Error
    })
}

function Get-FirstObjectPropertyValue {
  param(
    [AllowNull()]
    [object]$Object,

    [Parameter(Mandatory = $true)]
    [string[]]$Names,

    [AllowNull()]
    [string]$Fallback
  )

  if ($null -ne $Object) {
    foreach ($name in $Names) {
      if ($Object.PSObject.Properties.Name -contains $name) {
        $value = [string]$Object.$name
        if (-not [string]::IsNullOrWhiteSpace($value)) {
          return $value
        }
      }
    }
  }

  return $Fallback
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

function Clear-ComReferences {
  [System.GC]::Collect()
  [System.GC]::WaitForPendingFinalizers()
}

function Get-OfficeProcessDiagnostic {
  $processNames = @("WINWORD", "wps", "wpp", "et")
  $runningProcesses = New-Object System.Collections.Generic.List[string]

  foreach ($processName in $processNames) {
    try {
      foreach ($process in @(Get-Process -Name $processName -ErrorAction SilentlyContinue)) {
        [void]$runningProcesses.Add(("{0}(pid={1})" -f $process.ProcessName, $process.Id))
      }
    } catch {
      # Process diagnostics should never hide the conversion error.
    }
  }

  if ($runningProcesses.Count -eq 0) {
    return "Running Office/WPS processes: none detected."
  }

  return ("Running Office/WPS processes: {0}" -f ((@($runningProcesses) | Sort-Object -Unique) -join ", "))
}

function Get-TemplateConversionFailureMessage {
  param(
    [Parameter(Mandatory = $true)]
    [string]$SourcePath,

    [Parameter(Mandatory = $true)]
    [string]$ConvertedPath,

    [Parameter(Mandatory = $true)]
    [System.Collections.Generic.List[string]]$WpsErrors,

    [Parameter(Mandatory = $true)]
    [System.Collections.Generic.List[string]]$WordErrors
  )

  $wpsDetail = if ($WpsErrors.Count -gt 0) { (@($WpsErrors) -join " | ") } else { "not attempted or no error captured" }
  $wordDetail = if ($WordErrors.Count -gt 0) { (@($WordErrors) -join " | ") } else { "not attempted or no error captured" }
  $processDetail = Get-OfficeProcessDiagnostic

  return ("Failed to convert .doc template to .docx. Source: {0}. Output: {1}. WPS attempts: {2}. Word attempts: {3}. {4}. Common causes: Word/WPS is busy or showing a modal dialog, the template is open or locked, Protected View or first-run dialogs are pending, or Office/WPS was started with a different privilege level. Close Word/WPS dialogs and the template file, then retry; if it still fails, manually save the template as .docx and pass that .docx to -TemplatePath." -f $SourcePath, $ConvertedPath, $wpsDetail, $wordDetail, $processDetail)
}

function Convert-TemplateToDocxIfNeeded {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Path,

    [Parameter(Mandatory = $true)]
    [string]$OutputDir
  )

  $sourcePath = (Resolve-Path -LiteralPath $Path).Path
  $extension = [System.IO.Path]::GetExtension($sourcePath).ToLowerInvariant()
  if ($extension -eq ".docx") {
    return [pscustomobject]@{
      templatePath = $sourcePath
      sourceTemplatePath = $sourcePath
      status = "none"
      converter = "none"
      convertedTemplatePath = $null
    }
  }

  if ($extension -ne ".doc") {
    throw "Only .docx templates are supported directly; .doc templates can be converted with WPS/Word COM if Office automation is available: $sourcePath"
  }

  $convertedDir = Join-Path ([System.IO.Path]::GetFullPath($OutputDir)) "converted-templates"
  New-Item -ItemType Directory -Path $convertedDir -Force | Out-Null
  $convertedPath = Join-Path $convertedDir (([System.IO.Path]::GetFileNameWithoutExtension($sourcePath)) + ".docx")
  if (Test-Path -LiteralPath $convertedPath) {
    Remove-Item -LiteralPath $convertedPath -Force
  }

  $wpsErrors = New-Object System.Collections.Generic.List[string]
  for ($attempt = 1; $attempt -le 3; $attempt++) {
    $wpsApp = $null
    $wpsDoc = $null
    try {
      $wpsApp = New-Object -ComObject KWPS.Application
      $wpsApp.Visible = $false
      try {
        $wpsApp.DisplayAlerts = 0
      } catch {
        # Some WPS automation hosts do not expose DisplayAlerts.
      }
      Start-Sleep -Milliseconds (300 * $attempt)
      $wpsDoc = $wpsApp.Documents.Open($sourcePath)
      Start-Sleep -Milliseconds (500 * $attempt)
      $wpsDoc.SaveAs($convertedPath, 16)
      Start-Sleep -Milliseconds 300
      if (Test-Path -LiteralPath $convertedPath -PathType Leaf) {
        $resolvedConvertedPath = (Resolve-Path -LiteralPath $convertedPath).Path
        return [pscustomobject]@{
          templatePath = $resolvedConvertedPath
          sourceTemplatePath = $sourcePath
          status = "converted"
          converter = "wps"
          convertedTemplatePath = $resolvedConvertedPath
        }
      }
      throw "WPS SaveAs completed but did not create the output file: $convertedPath"
    } catch {
      [void]$wpsErrors.Add(("attempt {0}: {1}" -f $attempt, $_.Exception.Message))
      if (Test-Path -LiteralPath $convertedPath) {
        Remove-Item -LiteralPath $convertedPath -Force -ErrorAction SilentlyContinue
      }
      if ($attempt -lt 3) {
        Start-Sleep -Milliseconds (900 * $attempt)
      }
    } finally {
      if ($null -ne $wpsDoc) {
        try {
          $wpsDoc.Close($false)
        } catch {
          # Best-effort cleanup only.
        } finally {
          Release-ComObjectIfNeeded -ComObject $wpsDoc
        }
      }
      if ($null -ne $wpsApp) {
        try {
          $wpsApp.Quit()
        } catch {
          # Best-effort cleanup only.
        } finally {
          Release-ComObjectIfNeeded -ComObject $wpsApp
        }
      }
      Clear-ComReferences
    }
  }

  $wordErrors = New-Object System.Collections.Generic.List[string]
  for ($attempt = 1; $attempt -le 3; $attempt++) {
    $app = $null
    $doc = $null
    try {
      $app = New-Object -ComObject Word.Application
      $app.Visible = $false
      $app.DisplayAlerts = 0
      try {
        $app.AutomationSecurity = 3
      } catch {
        # Older Word automation hosts may not expose AutomationSecurity.
      }
      Start-Sleep -Milliseconds (300 * $attempt)
      $doc = $app.Documents.Open($sourcePath, $false, $true, $false)
      Start-Sleep -Milliseconds (500 * $attempt)
      $doc.SaveAs2($convertedPath, 16)
      Start-Sleep -Milliseconds 300
      if (-not (Test-Path -LiteralPath $convertedPath -PathType Leaf)) {
        throw "Word SaveAs2 completed but did not create the output file: $convertedPath"
      }
      break
    } catch {
      [void]$wordErrors.Add(("attempt {0}: {1}" -f $attempt, $_.Exception.Message))
      if (Test-Path -LiteralPath $convertedPath) {
        Remove-Item -LiteralPath $convertedPath -Force -ErrorAction SilentlyContinue
      }
      if ($attempt -eq 3) {
        throw (Get-TemplateConversionFailureMessage -SourcePath $sourcePath -ConvertedPath $convertedPath -WpsErrors $wpsErrors -WordErrors $wordErrors)
      }
      Start-Sleep -Milliseconds (900 * $attempt)
    } finally {
      if ($null -ne $doc) {
        try {
          $doc.Close($false)
        } catch {
          # Best-effort cleanup only.
        } finally {
          Release-ComObjectIfNeeded -ComObject $doc
        }
      }
      if ($null -ne $app) {
        try {
          $app.Quit()
        } catch {
          # Best-effort cleanup only.
        } finally {
          Release-ComObjectIfNeeded -ComObject $app
        }
      }
      Clear-ComReferences
    }
  }

  if (-not (Test-Path -LiteralPath $convertedPath -PathType Leaf)) {
    throw (Get-TemplateConversionFailureMessage -SourcePath $sourcePath -ConvertedPath $convertedPath -WpsErrors $wpsErrors -WordErrors $wordErrors)
  }

  $resolvedConvertedPath = (Resolve-Path -LiteralPath $convertedPath).Path
  return [pscustomobject]@{
    templatePath = $resolvedConvertedPath
    sourceTemplatePath = $sourcePath
    status = "converted"
    converter = "word"
    convertedTemplatePath = $resolvedConvertedPath
  }
}

function Get-OptionalTextContent {
  param(
    [AllowNull()]
    [string]$Path,

    [AllowNull()]
    [string]$InlineText
  )

  if (-not [string]::IsNullOrWhiteSpace($Path)) {
    return Get-Content -LiteralPath $Path -Raw -Encoding UTF8
  }

  if (-not [string]::IsNullOrWhiteSpace($InlineText)) {
    return $InlineText
  }

  return ""
}

function Get-JsonObjectOrNull {
  param(
    [AllowNull()]
    [string]$JsonText,

    [string]$SourceLabel = "JSON input"
  )

  if ([string]::IsNullOrWhiteSpace($JsonText)) {
    return $null
  }

  try {
    return $JsonText | ConvertFrom-Json
  } catch {
    throw "$SourceLabel is not valid JSON. $($_.Exception.Message)"
  }
}

function Get-RepoScriptPath {
  param(
    [Parameter(Mandatory = $true)]
    [string]$RepoRoot,

    [Parameter(Mandatory = $true)]
    [string]$ScriptName
  )

  return [System.IO.Path]::Combine($RepoRoot, "scripts", $ScriptName)
}

function Get-ImageInputItems {
  param(
    [Parameter(Mandatory = $true)]
    [string]$InputMode,

    [AllowNull()]
    [string]$SpecsPath,

    [AllowNull()]
    [string]$SpecsJson,

    [AllowNull()]
    [string[]]$Paths
  )

  if ([string]::Equals($InputMode, "image-paths", [System.StringComparison]::OrdinalIgnoreCase)) {
    return @(@($Paths | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | ForEach-Object {
          [pscustomobject]@{
            path = [string]$_
          }
        }))
  }

  $jsonText = Get-OptionalTextContent -Path $SpecsPath -InlineText $SpecsJson
  $rootObject = Get-JsonObjectOrNull -JsonText $jsonText -SourceLabel "Image specs"
  if ($null -eq $rootObject) {
    return @()
  }

  if ($rootObject -is [System.Collections.IEnumerable] -and $rootObject -isnot [string]) {
    return @($rootObject)
  }

  if ($rootObject.PSObject.Properties.Name -contains "images") {
    return @($rootObject.images)
  }

  return @($rootObject)
}

function Get-ImageItemValue {
  param(
    [Parameter(Mandatory = $true)]
    [object]$Item,

    [Parameter(Mandatory = $true)]
    [string[]]$Keys
  )

  foreach ($key in $Keys) {
    if ($Item -is [System.Collections.IDictionary]) {
      if ($Item.Contains($key) -and -not [string]::IsNullOrWhiteSpace([string]$Item[$key])) {
        return ([string]$Item[$key]).Trim()
      }
      continue
    }

    $property = $Item.PSObject.Properties[$key]
    if ($null -ne $property -and -not [string]::IsNullOrWhiteSpace([string]$property.Value)) {
      return ([string]$property.Value).Trim()
    }
  }

  return $null
}

function Test-IsFlowchartSignal {
  param(
    [AllowNull()]
    [string]$Text
  )

  return (
    -not [string]::IsNullOrWhiteSpace($Text) -and
    $Text -match '(?i)(流程图|总体架构|系统架构|体系结构图|flowchart|flow-chart|system architecture|architecture diagram)'
  )
}

function Test-ImageItemsContainFlowchart {
  param(
    [AllowEmptyCollection()]
    [object[]]$Items
  )

  foreach ($item in @($Items)) {
    foreach ($signal in @(
        (Get-ImageItemValue -Item $item -Keys @("caption", "title", "figureCaption")),
        (Get-ImageItemValue -Item $item -Keys @("section", "sectionName", "heading")),
        (Get-ImageItemValue -Item $item -Keys @("path", "imagePath", "file"))
      )) {
      if (Test-IsFlowchartSignal -Text $signal) {
        return $true
      }
    }
  }

  return $false
}

function Test-ImageItemsContainExplicitCaption {
  param(
    [AllowEmptyCollection()]
    [object[]]$Items
  )

  foreach ($item in @($Items)) {
    if (-not [string]::IsNullOrWhiteSpace((Get-ImageItemValue -Item $item -Keys @("caption", "title", "figureCaption")))) {
      return $true
    }
  }

  return $false
}

function Get-CourseDesignFlowchartTitle {
  param(
    [AllowNull()]
    [string]$MetadataPath,

    [AllowNull()]
    [string]$MetadataJson,

    [Parameter(Mandatory = $true)]
    [string]$ReportPath
  )

  $metadataText = Get-OptionalTextContent -Path $MetadataPath -InlineText $MetadataJson
  $metadataRoot = Get-JsonObjectOrNull -JsonText $metadataText -SourceLabel "Metadata"
  if ($null -ne $metadataRoot) {
    foreach ($key in @("课题名称", "项目名称", "题目", "标题", "实验名称")) {
      $value = Get-ImageItemValue -Item $metadataRoot -Keys @($key)
      if (-not [string]::IsNullOrWhiteSpace($value)) {
        return ("{0}流程图" -f $value)
      }
    }
  }

  $reportText = Get-Content -LiteralPath $ReportPath -Raw -Encoding UTF8
  foreach ($pattern in @(
      '课题名称[:：]\s*(?<value>.+)',
      '项目名称[:：]\s*(?<value>.+)',
      '题目[:：]\s*(?<value>.+)'
    )) {
    if ($reportText -match $pattern) {
      return ("{0}流程图" -f $matches["value"].Trim())
    }
  }

  return "课程设计实现流程图"
}

function Get-CourseDesignFlowchartSteps {
  param(
    [Parameter(Mandatory = $true)]
    [string]$ReportPath,

    [AllowNull()]
    [string]$RequirementsPath,

    [AllowNull()]
    [string]$RequirementsJson
  )

  $reportText = Get-Content -LiteralPath $ReportPath -Raw -Encoding UTF8
  $requirementsText = Get-OptionalTextContent -Path $RequirementsPath -InlineText $RequirementsJson
  $combinedText = ($reportText + [Environment]::NewLine + $requirementsText)
  $rootTitle = "课程设计系统"
  foreach ($pattern in @(
      '课题名称[:：]\s*(?<value>.+)',
      '项目名称[:：]\s*(?<value>.+)',
      '题目[:：]\s*(?<value>.+)'
    )) {
    if ($reportText -match $pattern) {
      $rootTitle = $matches["value"].Trim()
      break
    }
  }

  $frontendModules = New-Object System.Collections.Generic.List[string]
  if ($combinedText -match '分类|目录|导航|首页|列表') {
    $frontendModules.Add("分类浏览") | Out-Null
  }
  if ($combinedText -match '搜索|查询|检索') {
    $frontendModules.Add("信息检索") | Out-Null
  }
  if ($combinedText -match '详情|展示|说明|结果') {
    $frontendModules.Add("详情展示") | Out-Null
  }
  foreach ($fallbackLabel in @("页面展示", "交互入口", "结果详情")) {
    if ($frontendModules.Count -ge 3) {
      break
    }
    if ($frontendModules -notcontains $fallbackLabel) {
      $frontendModules.Add($fallbackLabel) | Out-Null
    }
  }

  $backendModules = New-Object System.Collections.Generic.List[string]
  if ($combinedText -match '逻辑|算法|调度|推荐|接口|服务端|业务') {
    $backendModules.Add("业务处理") | Out-Null
  }
  if ($combinedText -match '数据库|数据表|SQL|MySQL|SQLite|ER图|存储') {
    $backendModules.Add("数据管理") | Out-Null
  }
  if ($combinedText -match '收藏|订单|成绩|权限|日志|状态|审核') {
    $backendModules.Add("状态维护") | Out-Null
  }
  $backendModules.Add("测试验证") | Out-Null
  foreach ($fallbackLabel in @("异常处理", "日志管理", "配置维护")) {
    if ($backendModules.Count -ge 4) {
      break
    }
    if ($backendModules -notcontains $fallbackLabel) {
      $backendModules.Add($fallbackLabel) | Out-Null
    }
  }

  return @(
    "@TREE $rootTitle",
    ("@GROUP 前台模块|{0}" -f ((@($frontendModules | Select-Object -Unique -First 3)) -join "|")),
    ("@GROUP 后台模块|{0}" -f ((@($backendModules | Select-Object -Unique -First 4)) -join "|"))
  )
}

function Try-NewCourseDesignAutoFlowchart {
  param(
    [Parameter(Mandatory = $true)]
    [string]$ReportPath,

    [AllowNull()]
    [string]$MetadataPath,

    [AllowNull()]
    [string]$MetadataJson,

    [AllowNull()]
    [string]$RequirementsPath,

    [AllowNull()]
    [string]$RequirementsJson,

    [Parameter(Mandatory = $true)]
    [string]$OutputDir,

    [Parameter(Mandatory = $true)]
    [string]$RepoRoot,

    [Parameter(Mandatory = $true)]
    [string]$ImageInputMode,

    [AllowNull()]
    [string]$ImageSpecsPath,

    [AllowNull()]
    [string]$ImageSpecsJson,

    [AllowNull()]
    [string[]]$ImagePaths
  )

  $reportText = Get-Content -LiteralPath $ReportPath -Raw -Encoding UTF8
  $requirementsText = Get-OptionalTextContent -Path $RequirementsPath -InlineText $RequirementsJson
  if (($reportText + [Environment]::NewLine + $requirementsText) -match '(?i)(不要流程图|不需要流程图|no flowchart)') {
    return $null
  }

  $existingItems = @(Get-ImageInputItems -InputMode $ImageInputMode -SpecsPath $ImageSpecsPath -SpecsJson $ImageSpecsJson -Paths $ImagePaths)
  if (Test-ImageItemsContainFlowchart -Items $existingItems) {
    return $null
  }

  $rendererPath = [System.IO.Path]::Combine($RepoRoot, "scripts", "render-vertical-lab-flowchart.py")
  if (-not (Test-Path -LiteralPath $rendererPath)) {
    return $null
  }

  $artifactsDir = Join-Path $OutputDir "artifacts"
  New-Item -ItemType Directory -Path $artifactsDir -Force | Out-Null

  $stepsPath = Join-Path $artifactsDir "course-design-auto-flowchart.steps.txt"
  $flowchartPath = Join-Path $artifactsDir "course-design-auto-flowchart.png"
  $mergedSpecsPath = Join-Path $artifactsDir "course-design-auto-image-specs.json"
  $flowchartSteps = @(Get-CourseDesignFlowchartSteps -ReportPath $ReportPath -RequirementsPath $RequirementsPath -RequirementsJson $RequirementsJson)
  $flowchartTitle = Get-CourseDesignFlowchartTitle -MetadataPath $MetadataPath -MetadataJson $MetadataJson -ReportPath $ReportPath

  [System.IO.File]::WriteAllLines($stepsPath, $flowchartSteps, (New-Object System.Text.UTF8Encoding($true)))

  $rendered = $false
  foreach ($pythonOption in @(
      @{ command = "python"; prefix = @() },
      @{ command = "py"; prefix = @("-3") }
    )) {
    if ($null -eq (Get-Command $pythonOption.command -ErrorAction SilentlyContinue)) {
      continue
    }

    try {
      & $pythonOption.command @($pythonOption.prefix + @($rendererPath, "--out", $flowchartPath, "--title", $flowchartTitle, "--steps-file", $stepsPath))
      if (Test-Path -LiteralPath $flowchartPath) {
        $rendered = $true
        break
      }
    } catch {
      if (Test-Path -LiteralPath $flowchartPath) {
        Remove-Item -LiteralPath $flowchartPath -Force -ErrorAction SilentlyContinue
      }
    }
  }

  if (-not $rendered) {
    Write-Warning "Skipped course-design auto flowchart because the renderer could not run successfully."
    return $null
  }

  $flowchartCaption = if (Test-ImageItemsContainExplicitCaption -Items $existingItems) {
    "系统总体设计图"
  } else {
    "图1 系统总体设计图"
  }

  $mergedImages = @(
    [ordered]@{
      path = $flowchartPath
      section = "方案设计与实现"
      caption = $flowchartCaption
      widthCm = 15.8
    }
  ) + @($existingItems)

  $mergedSpecsRoot = [ordered]@{
    images = $mergedImages
  }
  [System.IO.File]::WriteAllText($mergedSpecsPath, ($mergedSpecsRoot | ConvertTo-Json -Depth 8), (New-Object System.Text.UTF8Encoding($true)))

  return [pscustomobject]@{
    flowchartPath = $flowchartPath
    imageSpecsPath = $mergedSpecsPath
  }
}

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$script:pipelineStages = New-Object System.Collections.Generic.List[object]
$templatePathDefaulted = (-not $PSBoundParameters.ContainsKey("TemplatePath") -or [string]::IsNullOrWhiteSpace($TemplatePath))
$reportProfile = Get-ReportProfile -ProfileName $ReportProfileName -ProfilePath $ReportProfilePath -RepoRoot $repoRoot
$resolvedReportProfilePath = [string]$reportProfile.resolvedProfilePath
$resolvedTemplatePath = Resolve-ExperimentReportTemplatePath `
  -TemplatePath $TemplatePath `
  -ReportProfileName ([string]$reportProfile.name) `
  -ReportProfilePath $resolvedReportProfilePath `
  -RepoRoot $repoRoot `
  -BuiltInTemplateId $BuiltInTemplateId
$sourceTemplatePath = $resolvedTemplatePath
$resolvedBuiltInTemplateId = Get-BuiltInTemplateIdForPath -RepoRoot $repoRoot -TemplatePath $resolvedTemplatePath
$templateSelectionSource = if (-not [string]::IsNullOrWhiteSpace($TemplatePath)) { "user" } else { "builtin" }
$templateConversion = $null
$resolvedReportPath = (Resolve-Path -LiteralPath $ReportPath).Path

$resolvedMetadataPath = $null
if (-not [string]::IsNullOrWhiteSpace($MetadataPath)) {
  $resolvedMetadataPath = (Resolve-Path -LiteralPath $MetadataPath).Path
}

$resolvedRequirementsPath = $null
if (-not [string]::IsNullOrWhiteSpace($RequirementsPath)) {
  $resolvedRequirementsPath = (Resolve-Path -LiteralPath $RequirementsPath).Path
}

$effectiveStyleProfile = if ($PSBoundParameters.ContainsKey("StyleProfile")) {
  $StyleProfile
} else {
  Get-ReportProfileDefaultStyleProfile -Profile $reportProfile
}
$templateFrameDefaulted = ($templatePathDefaulted -and (-not [bool]$CreateTemplateFrameDocx) -and [string]::IsNullOrWhiteSpace($TemplateFrameDocxOutPath) -and (Test-ExperimentReportTemplateFrameDefault -ReportProfileName ([string]$reportProfile.name) -ReportProfilePath $resolvedReportProfilePath))
$shouldCreateTemplateFrameDocx = ([bool]$CreateTemplateFrameDocx) -or (-not [string]::IsNullOrWhiteSpace($TemplateFrameDocxOutPath)) -or $templateFrameDefaulted
$metadataInputMode = if (-not [string]::IsNullOrWhiteSpace($resolvedMetadataPath)) {
  "path"
} elseif (-not [string]::IsNullOrWhiteSpace($MetadataJson)) {
  "inline"
} else {
  "none"
}
$requirementsInputMode = if (-not [string]::IsNullOrWhiteSpace($resolvedRequirementsPath)) {
  "path"
} elseif (-not [string]::IsNullOrWhiteSpace($RequirementsJson)) {
  "inline"
} else {
  "none"
}

$imageInputModes = 0
if (-not [string]::IsNullOrWhiteSpace($ImageSpecsPath)) { $imageInputModes++ }
if (-not [string]::IsNullOrWhiteSpace($ImageSpecsJson)) { $imageInputModes++ }
if ($null -ne $ImagePaths -and @($ImagePaths).Count -gt 0) { $imageInputModes++ }
$imageInputsProvided = ($imageInputModes -gt 0)
if ($imageInputModes -gt 1) {
  throw "Provide zero or one of -ImageSpecsPath, -ImageSpecsJson, or -ImagePaths."
}
$imageInputMode = if (-not [string]::IsNullOrWhiteSpace($ImageSpecsPath)) {
  "specs-path"
} elseif (-not [string]::IsNullOrWhiteSpace($ImageSpecsJson)) {
  "specs-json"
} elseif ($null -ne $ImagePaths -and @($ImagePaths).Count -gt 0) {
  "image-paths"
} else {
  "none"
}

$styleOutputRequested = $StyleFinalDocx -or (-not [string]::IsNullOrWhiteSpace($StyledDocxOutPath))
$allowLegacyCourseDesignEnhancements = (
  $templatePathDefaulted -or
  [string]::Equals($TemplateStyleMode, "normalize", [System.StringComparison]::OrdinalIgnoreCase) -or
  $styleOutputRequested
)
$runFullPipeline = [string]::Equals($PipelineMode, "full", [System.StringComparison]::OrdinalIgnoreCase)
$shouldRunValidation = $runFullPipeline -or ($requirementsInputMode -ne "none")
$shouldGenerateDebugOutlines = $runFullPipeline

if ([string]::IsNullOrWhiteSpace($OutputDir)) {
  $OutputDir = [System.IO.Path]::Combine($repoRoot, "tests-output", ("build-" + (Get-Date -Format "yyyyMMdd-HHmmss")))
}

$resolvedOutputDir = [System.IO.Path]::GetFullPath($OutputDir)
New-Item -ItemType Directory -Path $resolvedOutputDir -Force | Out-Null

$pipelineTracePath = Join-Path $resolvedOutputDir "pipeline-trace.json"
$materialsPath = Join-Path $resolvedOutputDir "materials-analysis.json"
$contentPlanPath = Join-Path $resolvedOutputDir "content-plan.json"
$templateContractPath = Join-Path $resolvedOutputDir "template-style-contract.json"
$formatValidationPath = Join-Path $resolvedOutputDir "format-validation.json"
$imageManifestPath = Join-Path $resolvedOutputDir "image-manifest.json"
$visualValidationPath = Join-Path $resolvedOutputDir "visual-validation.json"
$strictPdfPath = Join-Path $resolvedOutputDir "strict-preview.pdf"
$strictPreviewDir = Join-Path $resolvedOutputDir "strict-preview"

$materialSummary = [pscustomobject]@{
  schemaVersion = "1.0"
  report = [pscustomobject]@{
    path = $resolvedReportPath
    sizeBytes = (Get-Item -LiteralPath $resolvedReportPath).Length
  }
  metadata = [pscustomobject]@{
    mode = $metadataInputMode
    path = $resolvedMetadataPath
  }
  requirements = [pscustomobject]@{
    mode = $requirementsInputMode
    path = $resolvedRequirementsPath
  }
  template = [pscustomobject]@{
    sourcePath = $sourceTemplatePath
    defaulted = $templatePathDefaulted
    selectionSource = $templateSelectionSource
    builtInTemplateId = $resolvedBuiltInTemplateId
  }
  images = [pscustomobject]@{
    mode = $imageInputMode
    requestedCount = $(if ($null -ne $ImagePaths) { @($ImagePaths).Count } else { $null })
  }
}
[System.IO.File]::WriteAllText($materialsPath, ($materialSummary | ConvertTo-Json -Depth 6), (New-Object System.Text.UTF8Encoding($true)))
Add-PipelineStage -Name "materials-parsing" -Output ([pscustomobject]@{ path = $materialsPath })

$metadataObject = if (-not [string]::IsNullOrWhiteSpace($resolvedMetadataPath)) {
  (Get-Content -LiteralPath $resolvedMetadataPath -Raw -Encoding UTF8) | ConvertFrom-Json
} elseif (-not [string]::IsNullOrWhiteSpace($MetadataJson)) {
  $MetadataJson | ConvertFrom-Json
} else {
  $null
}
$planCourseName = Get-FirstObjectPropertyValue -Object $metadataObject -Names @("CourseName", "courseName", "Course") -Fallback ([string]$reportProfile.displayName)
$planExperimentName = Get-FirstObjectPropertyValue -Object $metadataObject -Names @("ExperimentName", "experimentName", "Title") -Fallback ([System.IO.Path]::GetFileNameWithoutExtension($resolvedReportPath))
$planVariantSeed = Get-FirstObjectPropertyValue -Object $metadataObject -Names @("StudentId", "studentId", "Name", "studentName") -Fallback $planExperimentName
& (Get-RepoScriptPath -RepoRoot $repoRoot -ScriptName "plan-report-content.ps1") `
  -CourseName $planCourseName `
  -ExperimentName $planExperimentName `
  -DetailLevel $DetailLevel `
  -VariantSeed $planVariantSeed `
  -OutFile $contentPlanPath | Out-Null
Add-PipelineStage -Name "content-planning" -Output ([pscustomobject]@{ path = $contentPlanPath })
Add-PipelineStage -Name "body-generation" -Output ([pscustomobject]@{ source = "provided-report"; path = $resolvedReportPath })

$effectiveImageInputMode = $imageInputMode
$effectiveImageInputsProvided = $imageInputsProvided
$effectiveImageSpecsPath = $ImageSpecsPath
$effectiveImageSpecsJson = $ImageSpecsJson
$effectiveImagePaths = $ImagePaths
$imageSelectionManifest = $null
$autoCourseDesignFlowchart = $null
if (
  [string]::Equals([string]$reportProfile.name, "course-design-report", [System.StringComparison]::OrdinalIgnoreCase) -and
  $allowLegacyCourseDesignEnhancements
) {
  $autoCourseDesignFlowchart = Try-NewCourseDesignAutoFlowchart `
    -ReportPath $resolvedReportPath `
    -MetadataPath $resolvedMetadataPath `
    -MetadataJson $MetadataJson `
    -RequirementsPath $resolvedRequirementsPath `
    -RequirementsJson $RequirementsJson `
    -OutputDir $resolvedOutputDir `
    -RepoRoot $repoRoot `
    -ImageInputMode $imageInputMode `
    -ImageSpecsPath $ImageSpecsPath `
    -ImageSpecsJson $ImageSpecsJson `
    -ImagePaths $ImagePaths
  if ($null -ne $autoCourseDesignFlowchart) {
    $effectiveImageInputMode = "specs-path"
    $effectiveImageInputsProvided = $true
    $effectiveImageSpecsPath = [string]$autoCourseDesignFlowchart.imageSpecsPath
    $effectiveImageSpecsJson = $null
    $effectiveImagePaths = @()
  }
}

if (
  [string]::Equals($effectiveImageInputMode, "image-paths", [System.StringComparison]::OrdinalIgnoreCase) -and
  $null -ne $effectiveImagePaths -and
  @($effectiveImagePaths).Count -gt 0
) {
  $imageSelectionArguments = @("image-manifest") + @($effectiveImagePaths)
  if ($RequestedImageCount -ge 0) {
    $imageSelectionArguments += @("--count", [string]$RequestedImageCount)
  }
  $imageSelectionArguments += @("--output", $imageManifestPath)
  Invoke-UniversalReportCore -Arguments $imageSelectionArguments -RepoRoot $repoRoot
  $imageSelectionManifest = (Get-Content -LiteralPath $imageManifestPath -Raw -Encoding UTF8) | ConvertFrom-Json
  $effectiveImagePaths = @($imageSelectionManifest.images | ForEach-Object { [string]$_.path })
  $effectiveImageInputsProvided = ($effectiveImagePaths.Count -gt 0)
  if (-not $effectiveImageInputsProvided -and $RequestedImageCount -ne 0) {
    throw "No usable images remained after path validation and duplicate filtering. See $imageManifestPath"
  }
}
if ($RequestedImageCount -gt 0 -and -not $effectiveImageInputsProvided) {
  throw "Requested $RequestedImageCount images, but no image input was provided."
}

$shouldGenerateImagePlan = $effectiveImageInputsProvided -or (-not [string]::IsNullOrWhiteSpace($ImagePlanOutPath))
$shouldRunLayoutCheck = $runFullPipeline -or $effectiveImageInputsProvided

$templateConversionParams = @{
  SourcePath = $resolvedTemplatePath
  OutputDir = (Join-Path $resolvedOutputDir "converted-templates")
}
if ($AllowOfficeCom) {
  $templateConversionParams.AllowOfficeCom = $true
}
$templateConversion = & (Get-RepoScriptPath -RepoRoot $repoRoot -ScriptName "convert-report-template.ps1") @templateConversionParams
$resolvedTemplatePath = [string]$templateConversion.templatePath

$resolvedTemplateCacheDir = if (-not [string]::IsNullOrWhiteSpace($TemplateCacheDir)) {
  [System.IO.Path]::GetFullPath($TemplateCacheDir)
} elseif (-not [string]::IsNullOrWhiteSpace($env:EXPERIMENT_REPORT_CACHE_ROOT)) {
  [System.IO.Path]::GetFullPath($env:EXPERIMENT_REPORT_CACHE_ROOT)
} else {
  Join-Path $HOME ".cache\experiment-report\templates"
}
New-Item -ItemType Directory -Path $resolvedTemplateCacheDir -Force | Out-Null
& (Get-RepoScriptPath -RepoRoot $repoRoot -ScriptName "analyze-docx-template.ps1") `
  -TemplatePath $resolvedTemplatePath `
  -CacheDir $resolvedTemplateCacheDir `
  -OutFile $templateContractPath | Out-Null
$templateContract = (Get-Content -LiteralPath $templateContractPath -Raw -Encoding UTF8) | ConvertFrom-Json
$qualityRecommendation = $templateContract.qualityRecommendation
Add-PipelineStage -Name "template-analysis" -Output ([pscustomobject]@{
    path = $templateContractPath
    cacheHit = $(if ($templateContract.PSObject.Properties.Name -contains "cache") { [bool]$templateContract.cache.hit } else { $false })
    riskLevel = [string]$templateContract.risk.level
    recommendedMode = [string]$qualityRecommendation.recommendedMode
  })

$resolvedFieldMapOutPath = if ([string]::IsNullOrWhiteSpace($FieldMapOutPath)) {
  Join-Path $resolvedOutputDir "generated-field-map.json"
} else {
  [System.IO.Path]::GetFullPath($FieldMapOutPath)
}
Ensure-ParentDirectory -Path $resolvedFieldMapOutPath

$resolvedFilledDocxOutPath = if ([string]::IsNullOrWhiteSpace($FilledDocxOutPath)) {
  Join-Path $resolvedOutputDir (([System.IO.Path]::GetFileNameWithoutExtension($resolvedTemplatePath)) + ".filled.docx")
} else {
  [System.IO.Path]::GetFullPath($FilledDocxOutPath)
}
Ensure-ParentDirectory -Path $resolvedFilledDocxOutPath

$resolvedImagePlanOutPath = $null
$resolvedImageMapOutPath = $null
$resolvedFilledDocxWithImagesOutPath = $null
$resolvedCourseDesignTablesDocxOutPath = $null
$resolvedStyledDocxOutPath = $null
$validationPath = $null
$filledOutlinePath = $null
$filledWithImagesOutlinePath = $null
$styledOutlinePath = $null
$resolvedTemplateFrameDocxOutPath = $null
$styleResult = $null
$summaryPath = Join-Path $resolvedOutputDir "summary.json"
$layoutCheckPath = Join-Path $resolvedOutputDir "layout-check.json"
$layoutCheckResult = $null
$courseDesignTablesResult = $null
$expectedLayoutImageCount = -1
$expectedLayoutCaptionCount = -1
$imagePlanLowConfidenceCount = $null
$imagePlanNeedsReview = $null

$validationResult = $null
if ($shouldRunValidation) {
  $validationPath = Join-Path $resolvedOutputDir "validation.json"
  $validationParams = @{
    Path = $resolvedReportPath
    Format = "json"
  }
  if (-not [string]::IsNullOrWhiteSpace($ReportProfileName)) {
    $validationParams.ReportProfileName = $ReportProfileName
  }
  if (-not [string]::IsNullOrWhiteSpace($resolvedReportProfilePath)) {
    $validationParams.ReportProfilePath = $resolvedReportProfilePath
  }
  if (-not [string]::IsNullOrWhiteSpace($resolvedRequirementsPath)) {
    $validationParams.RequirementsPath = $resolvedRequirementsPath
  } elseif (-not [string]::IsNullOrWhiteSpace($RequirementsJson)) {
    $validationParams.RequirementsJson = $RequirementsJson
  }

  $validationJson = & (Get-RepoScriptPath -RepoRoot $repoRoot -ScriptName "validate-report-draft.ps1") @validationParams | Out-String
  [System.IO.File]::WriteAllText($validationPath, $validationJson, (New-Object System.Text.UTF8Encoding($true)))
  $validationResult = $validationJson | ConvertFrom-Json
}

$fieldMapParams = @{
  TemplatePath = $resolvedTemplatePath
  ReportPath = $resolvedReportPath
  ReportProfileName = $ReportProfileName
  ReportProfilePath = $resolvedReportProfilePath
  Format = "json"
  OutFile = $resolvedFieldMapOutPath
}
if (-not [string]::IsNullOrWhiteSpace($resolvedMetadataPath)) {
  $fieldMapParams.MetadataPath = $resolvedMetadataPath
} elseif (-not [string]::IsNullOrWhiteSpace($MetadataJson)) {
  $fieldMapParams.MetadataJson = $MetadataJson
}

& (Get-RepoScriptPath -RepoRoot $repoRoot -ScriptName "generate-docx-field-map.ps1") @fieldMapParams | Out-Null
& (Get-RepoScriptPath -RepoRoot $repoRoot -ScriptName "apply-docx-field-map.ps1") -TemplatePath $resolvedTemplatePath -MappingPath $resolvedFieldMapOutPath -OutPath $resolvedFilledDocxOutPath -Overwrite | Out-Null

if ($shouldGenerateDebugOutlines) {
  $filledOutlinePath = Join-Path $resolvedOutputDir "filled-template-outline.md"
  $filledOutline = & (Get-RepoScriptPath -RepoRoot $repoRoot -ScriptName "extract-docx-template.ps1") -Path $resolvedFilledDocxOutPath -Format markdown | Out-String
  [System.IO.File]::WriteAllText($filledOutlinePath, $filledOutline, (New-Object System.Text.UTF8Encoding($true)))
}

if ($effectiveImageInputsProvided) {
  if ($shouldGenerateImagePlan) {
    $resolvedImagePlanOutPath = if ([string]::IsNullOrWhiteSpace($ImagePlanOutPath)) {
      Join-Path $resolvedOutputDir "image-placement-plan.md"
    } else {
      [System.IO.Path]::GetFullPath($ImagePlanOutPath)
    }
    Ensure-ParentDirectory -Path $resolvedImagePlanOutPath
  }

  $resolvedImageMapOutPath = if ([string]::IsNullOrWhiteSpace($ImageMapOutPath)) {
    Join-Path $resolvedOutputDir "generated-image-map.json"
  } else {
    [System.IO.Path]::GetFullPath($ImageMapOutPath)
  }
  Ensure-ParentDirectory -Path $resolvedImageMapOutPath

  $resolvedFilledDocxWithImagesOutPath = if ([string]::IsNullOrWhiteSpace($FilledDocxWithImagesOutPath)) {
    Join-Path $resolvedOutputDir (([System.IO.Path]::GetFileNameWithoutExtension($resolvedFilledDocxOutPath)) + ".images.docx")
  } else {
    [System.IO.Path]::GetFullPath($FilledDocxWithImagesOutPath)
  }
  Ensure-ParentDirectory -Path $resolvedFilledDocxWithImagesOutPath

  $imageInputParams = @{
    DocxPath = $resolvedFilledDocxOutPath
    ReportProfileName = $ReportProfileName
    ReportProfilePath = $resolvedReportProfilePath
  }
  if ([string]::Equals($effectiveImageInputMode, "specs-path", [System.StringComparison]::OrdinalIgnoreCase)) {
    $imageInputParams.ImageSpecsPath = (Resolve-Path -LiteralPath $effectiveImageSpecsPath).Path
  } elseif ([string]::Equals($effectiveImageInputMode, "specs-json", [System.StringComparison]::OrdinalIgnoreCase)) {
    $imageInputParams.ImageSpecsJson = $effectiveImageSpecsJson
  } else {
    $imageInputParams.ImagePaths = $effectiveImagePaths
  }

  if ($shouldGenerateImagePlan) {
    $imagePlanJsonParams = $imageInputParams.Clone()
    $imagePlanJsonParams.Format = "json"
    $imagePlanJsonParams.PlanOnly = $true
    $imagePlanResult = ((& (Get-RepoScriptPath -RepoRoot $repoRoot -ScriptName "generate-docx-image-map.ps1") @imagePlanJsonParams) | Out-String) | ConvertFrom-Json
    $imagePlanEntries = if ($null -ne $imagePlanResult -and $imagePlanResult.PSObject.Properties.Name -contains "plan") {
      @($imagePlanResult.plan)
    } else {
      @()
    }
    $imagePlanLowConfidenceCount = @($imagePlanEntries | Where-Object { [string]$_.confidence -eq "low" }).Count
    $imagePlanNeedsReview = ($imagePlanLowConfidenceCount -gt 0)

    $imagePlanMarkdownParams = $imageInputParams.Clone()
    $imagePlanMarkdownParams.Format = "markdown"
    $imagePlanMarkdownParams.PlanOnly = $true
    $imagePlanMarkdownParams.OutFile = $resolvedImagePlanOutPath
    & (Get-RepoScriptPath -RepoRoot $repoRoot -ScriptName "generate-docx-image-map.ps1") @imagePlanMarkdownParams | Out-Null
  }

  $imageMapParams = $imageInputParams.Clone()
  $imageMapParams.Format = "json"
  $imageMapParams.OutFile = $resolvedImageMapOutPath
  & (Get-RepoScriptPath -RepoRoot $repoRoot -ScriptName "generate-docx-image-map.ps1") @imageMapParams | Out-Null
  $generatedImageMap = (Get-Content -LiteralPath $resolvedImageMapOutPath -Raw -Encoding UTF8) | ConvertFrom-Json
  if ($RequestedImageCount -ge 0 -and @($generatedImageMap.images).Count -ne $RequestedImageCount) {
    throw "Requested $RequestedImageCount images, but the resolved image map contains $(@($generatedImageMap.images).Count)."
  }
  if ($shouldRunLayoutCheck -and $null -ne $generatedImageMap -and $generatedImageMap.PSObject.Properties.Name -contains "images") {
    $expectedLayoutImageCount = @($generatedImageMap.images).Count
    $expectedLayoutCaptionCount = @($generatedImageMap.images | Where-Object {
        $_.PSObject.Properties.Name -contains "caption" -and -not [string]::IsNullOrWhiteSpace([string]$_.caption)
      }).Count
  }

  $manifestImages = New-Object System.Collections.Generic.List[object]
  $selectionByPath = @{}
  if ($null -ne $imageSelectionManifest) {
    foreach ($selectedItem in @($imageSelectionManifest.images)) {
      $selectionByPath[[System.IO.Path]::GetFullPath([string]$selectedItem.path).ToLowerInvariant()] = $selectedItem
    }
  }
  $manifestIndex = 0
  foreach ($imageItem in @($generatedImageMap.images)) {
    $manifestIndex++
    $planItem = if ($manifestIndex -le @($imagePlanEntries).Count) { @($imagePlanEntries)[$manifestIndex - 1] } else { $null }
    $selectionItem = $null
    $selectionKey = [System.IO.Path]::GetFullPath([string]$imageItem.path).ToLowerInvariant()
    if ($selectionByPath.ContainsKey($selectionKey)) {
      $selectionItem = $selectionByPath[$selectionKey]
    }
    [void]$manifestImages.Add([pscustomobject]@{
        path = [string]$imageItem.path
        section = $(if ($imageItem.PSObject.Properties.Name -contains "section") { [string]$imageItem.section } elseif ($null -ne $planItem -and $planItem.PSObject.Properties.Name -contains "section") { [string]$planItem.section } else { $null })
        anchor = $(if ($imageItem.PSObject.Properties.Name -contains "anchor") { [string]$imageItem.anchor } else { $null })
        caption = $(if ($imageItem.PSObject.Properties.Name -contains "caption") { [string]$imageItem.caption } else { $null })
        order = $manifestIndex
        layout = $(if ($imageItem.PSObject.Properties.Name -contains "layout") { $imageItem.layout } else { [pscustomobject]@{ mode = "single"; columns = 1 } })
        selectionReason = $(if ($null -ne $selectionItem -and $selectionItem.PSObject.Properties.Name -contains "selectionReason") { [string]$selectionItem.selectionReason } elseif ($null -ne $planItem -and $planItem.PSObject.Properties.Name -contains "reason") { [string]$planItem.reason } elseif ($null -ne $planItem -and $planItem.PSObject.Properties.Name -contains "confidence") { "模板锚点匹配，置信度：$([string]$planItem.confidence)" } else { "按用户顺序和章节锚点选取" })
      })
  }
  $imageManifest = [pscustomobject]@{
    schemaVersion = "1.0"
    requestedCount = $(if ($RequestedImageCount -ge 0) { $RequestedImageCount } elseif ($null -ne $ImagePaths) { @($ImagePaths).Count } else { $expectedLayoutImageCount })
    selectedCount = $manifestImages.Count
    duplicatesFiltered = $(if ($null -ne $imageSelectionManifest) { [int]$imageSelectionManifest.duplicatesFiltered } else { 0 })
    rejected = $(if ($null -ne $imageSelectionManifest) { @($imageSelectionManifest.rejected) } else { @() })
    images = @($manifestImages.ToArray())
  }
  [System.IO.File]::WriteAllText($imageManifestPath, ($imageManifest | ConvertTo-Json -Depth 10), (New-Object System.Text.UTF8Encoding($true)))
  Add-PipelineStage -Name "image-selection-matching" -Output ([pscustomobject]@{
      path = $imageManifestPath
      selectedCount = $manifestImages.Count
      needsReview = $imagePlanNeedsReview
    })

  & (Get-RepoScriptPath -RepoRoot $repoRoot -ScriptName "insert-docx-images.ps1") `
    -DocxPath $resolvedFilledDocxOutPath `
    -MappingPath $resolvedImageMapOutPath `
    -ReportProfileName $ReportProfileName `
    -ReportProfilePath $resolvedReportProfilePath `
    -OutPath $resolvedFilledDocxWithImagesOutPath `
    -Overwrite | Out-Null

  if ($shouldGenerateDebugOutlines) {
    $filledWithImagesOutlinePath = Join-Path $resolvedOutputDir "filled-template-with-images-outline.md"
    $filledWithImagesOutline = & (Get-RepoScriptPath -RepoRoot $repoRoot -ScriptName "extract-docx-template.ps1") -Path $resolvedFilledDocxWithImagesOutPath -Format markdown | Out-String
    [System.IO.File]::WriteAllText($filledWithImagesOutlinePath, $filledWithImagesOutline, (New-Object System.Text.UTF8Encoding($true)))
  }
}

if (-not $effectiveImageInputsProvided) {
  $emptyImageManifest = [pscustomobject]@{
    schemaVersion = "1.0"
    requestedCount = 0
    selectedCount = 0
    duplicatesFiltered = 0
    images = @()
  }
  [System.IO.File]::WriteAllText($imageManifestPath, ($emptyImageManifest | ConvertTo-Json -Depth 5), (New-Object System.Text.UTF8Encoding($true)))
  Add-PipelineStage -Name "image-selection-matching" -Status "skipped" -Output ([pscustomobject]@{
      path = $imageManifestPath
      reason = "no-images-provided"
    })
}

if (
  [string]::Equals([string]$reportProfile.name, "course-design-report", [System.StringComparison]::OrdinalIgnoreCase) -and
  $allowLegacyCourseDesignEnhancements
) {
  $courseDesignTablesInputPath = if ($null -ne $resolvedFilledDocxWithImagesOutPath) { $resolvedFilledDocxWithImagesOutPath } else { $resolvedFilledDocxOutPath }
  $resolvedCourseDesignTablesDocxOutPath = Join-Path $resolvedOutputDir (([System.IO.Path]::GetFileNameWithoutExtension($courseDesignTablesInputPath)) + ".course-tables.docx")
  Ensure-ParentDirectory -Path $resolvedCourseDesignTablesDocxOutPath
  $courseDesignTablesResult = & (Get-RepoScriptPath -RepoRoot $repoRoot -ScriptName "insert-course-design-tables.ps1") `
    -DocxPath $courseDesignTablesInputPath `
    -OutPath $resolvedCourseDesignTablesDocxOutPath `
    -Overwrite
  Add-PipelineStage -Name "course-design-enhancements" -Output ([pscustomobject]@{
      autoFlowchartPath = $(if ($null -ne $autoCourseDesignFlowchart) { [string]$autoCourseDesignFlowchart.flowchartPath } else { $null })
      tablesDocxPath = $resolvedCourseDesignTablesDocxOutPath
    })
} elseif ([string]::Equals([string]$reportProfile.name, "course-design-report", [System.StringComparison]::OrdinalIgnoreCase)) {
  Add-PipelineStage -Name "course-design-enhancements" -Status "skipped" -Output ([pscustomobject]@{
      reason = "custom-template-preserve-mode"
    })
}

if ($styleOutputRequested) {
  $styleInputPath = if ($null -ne $resolvedCourseDesignTablesDocxOutPath) { $resolvedCourseDesignTablesDocxOutPath } elseif ($null -ne $resolvedFilledDocxWithImagesOutPath) { $resolvedFilledDocxWithImagesOutPath } else { $resolvedFilledDocxOutPath }
  $resolvedStyledDocxOutPath = if ([string]::IsNullOrWhiteSpace($StyledDocxOutPath)) {
    Join-Path $resolvedOutputDir (([System.IO.Path]::GetFileNameWithoutExtension($styleInputPath)) + ".styled.docx")
  } else {
    [System.IO.Path]::GetFullPath($StyledDocxOutPath)
  }

  $styleParams = @{
    DocxPath = $styleInputPath
    OutPath = $resolvedStyledDocxOutPath
    Overwrite = $true
    Profile = $effectiveStyleProfile
    ReportProfileName = [string]$reportProfile.name
    ReportProfilePath = [string]$reportProfile.resolvedProfilePath
  }
  if (-not [string]::IsNullOrWhiteSpace($StyleProfilePath)) {
    $styleParams.ProfilePath = (Resolve-Path -LiteralPath $StyleProfilePath).Path
  }

  $styleResult = & (Get-RepoScriptPath -RepoRoot $repoRoot -ScriptName "format-docx-report-style.ps1") @styleParams

  if ($shouldGenerateDebugOutlines) {
    $styledOutlinePath = Join-Path $resolvedOutputDir "styled-template-outline.md"
    $styledOutline = & (Get-RepoScriptPath -RepoRoot $repoRoot -ScriptName "extract-docx-template.ps1") -Path $resolvedStyledDocxOutPath -Format markdown | Out-String
    [System.IO.File]::WriteAllText($styledOutlinePath, $styledOutline, (New-Object System.Text.UTF8Encoding($true)))
  }
}

$preFrameFinalDocxPath = if ($null -ne $resolvedStyledDocxOutPath) {
  $resolvedStyledDocxOutPath
} elseif ($null -ne $resolvedCourseDesignTablesDocxOutPath) {
  $resolvedCourseDesignTablesDocxOutPath
} elseif ($null -ne $resolvedFilledDocxWithImagesOutPath) {
  $resolvedFilledDocxWithImagesOutPath
} else {
  $resolvedFilledDocxOutPath
}
$finalDocxPath = $preFrameFinalDocxPath

if ($shouldCreateTemplateFrameDocx) {
  $resolvedTemplateFrameDocxOutPath = if ([string]::IsNullOrWhiteSpace($TemplateFrameDocxOutPath)) {
    Join-Path $resolvedOutputDir (([System.IO.Path]::GetFileNameWithoutExtension($preFrameFinalDocxPath)) + ".template-frame.docx")
  } else {
    [System.IO.Path]::GetFullPath($TemplateFrameDocxOutPath)
  }
  Ensure-ParentDirectory -Path $resolvedTemplateFrameDocxOutPath
  & (Get-RepoScriptPath -RepoRoot $repoRoot -ScriptName "convert-docx-template-frame.ps1") `
    -DocxPath $preFrameFinalDocxPath `
    -OutPath $resolvedTemplateFrameDocxOutPath `
    -Overwrite | Out-Null
  $finalDocxPath = $resolvedTemplateFrameDocxOutPath
}

Add-PipelineStage -Name "docx-rendering" -Output ([pscustomobject]@{
    finalDocxPath = $finalDocxPath
    templateStyleMode = $TemplateStyleMode
    explicitStyleNormalization = [bool]$styleOutputRequested
    explicitTemplateFrame = [bool]$shouldCreateTemplateFrameDocx
  })

if ($shouldRunLayoutCheck) {
  $layoutCheckParams = @{
    DocxPath = $finalDocxPath
    Format = "json"
    OutFile = $layoutCheckPath
  }
  if (-not [string]::IsNullOrWhiteSpace($ReportProfileName)) {
    $layoutCheckParams.ReportProfileName = $ReportProfileName
  }
  if (-not [string]::IsNullOrWhiteSpace($resolvedReportProfilePath)) {
    $layoutCheckParams.ReportProfilePath = $resolvedReportProfilePath
  }
  if ($expectedLayoutImageCount -ge 0) {
    $layoutCheckParams.ExpectedImageCount = $expectedLayoutImageCount
  }
  if ($expectedLayoutCaptionCount -ge 0) {
    $layoutCheckParams.ExpectedCaptionCount = $expectedLayoutCaptionCount
  }
  if ($shouldCreateTemplateFrameDocx) {
    $layoutCheckParams.RequireTemplateFrame = $true
  }
  if ([string]::Equals($QualityMode, "strict", [System.StringComparison]::OrdinalIgnoreCase)) {
    $layoutCheckParams.RequireReadableMetadata = $true
  }
  & (Get-RepoScriptPath -RepoRoot $repoRoot -ScriptName "check-docx-layout.ps1") @layoutCheckParams | Out-Null
  $layoutCheckResult = (Get-Content -LiteralPath $layoutCheckPath -Raw -Encoding UTF8) | ConvertFrom-Json
  Add-PipelineStage `
    -Name "layout-validation" `
    -Status $(if ([bool]$layoutCheckResult.passed) { "completed" } else { "failed" }) `
    -Output ([pscustomobject]@{
        path = $layoutCheckPath
        passed = [bool]$layoutCheckResult.passed
        errorCount = [int]$layoutCheckResult.summary.errorCount
        warningCount = [int]$layoutCheckResult.summary.warningCount
      }) `
    -Error $(if (-not [bool]$layoutCheckResult.passed) {
        [pscustomobject]@{
          code = "layout-validation-failed"
          message = "文档结构与布局检查未通过。"
          suggestion = "查看 layout-check.json，修复图片、图注、分页或模板框架问题。"
        }
      } else { $null })
} else {
  Add-PipelineStage -Name "layout-validation" -Status "skipped" -Output ([pscustomobject]@{
      reason = "fast-mode-without-images"
    })
}

& (Get-RepoScriptPath -RepoRoot $repoRoot -ScriptName "validate-docx-format.ps1") `
  -TemplatePath $resolvedTemplatePath `
  -DocxPath $finalDocxPath `
  -OutFile $formatValidationPath | Out-Null
$formatValidationResult = (Get-Content -LiteralPath $formatValidationPath -Raw -Encoding UTF8) | ConvertFrom-Json
$formatValidationRequired = (
  [string]::Equals($TemplateStyleMode, "preserve", [System.StringComparison]::OrdinalIgnoreCase) -and
  (-not $templatePathDefaulted) -and
  (-not $styleOutputRequested) -and
  (-not $shouldCreateTemplateFrameDocx)
)
Add-PipelineStage `
  -Name "format-validation" `
  -Status $(if ([bool]$formatValidationResult.passed -or (-not $formatValidationRequired)) { "completed" } else { "failed" }) `
  -Output ([pscustomobject]@{
      path = $formatValidationPath
      passed = [bool]$formatValidationResult.passed
      required = $formatValidationRequired
      failedCodes = @($formatValidationResult.summary.failedCodes)
    }) `
  -Error $(if ((-not [bool]$formatValidationResult.passed) -and $formatValidationRequired) {
      [pscustomobject]@{
        code = "template-format-drift"
        message = "最终文档与用户模板的关键格式不一致。"
        suggestion = "查看 format-validation.json，修复对应角色或表格后重试。"
      }
    } else { $null })

$visualValidationResult = $null
if ([string]::Equals($QualityMode, "strict", [System.StringComparison]::OrdinalIgnoreCase)) {
  $sourceHasPageBorder = $false
  foreach ($section in @($templateContract.page.sections)) {
    if ($null -ne $section.pageBorder -and @($section.pageBorder.sides.PSObject.Properties).Count -ge 4) {
      $sourceHasPageBorder = $true
      break
    }
  }
  $visualValidationParams = @{
    DocxPath = $finalDocxPath
    OutputDir = $strictPreviewDir
    PdfPath = $strictPdfPath
    OutFile = $visualValidationPath
  }
  if ($shouldCreateTemplateFrameDocx -or $sourceHasPageBorder) {
    $visualValidationParams.RequireClosedFrame = $true
  }
  if ($AllowOfficeCom) {
    $visualValidationParams.AllowOfficeCom = $true
  }
  $visualValidationResult = & (Get-RepoScriptPath -RepoRoot $repoRoot -ScriptName "run-visual-validation.ps1") @visualValidationParams
  if (Test-Path -LiteralPath $visualValidationPath) {
    $visualValidationResult = (Get-Content -LiteralPath $visualValidationPath -Raw -Encoding UTF8) | ConvertFrom-Json
  }
  Add-PipelineStage `
    -Name "visual-validation" `
    -Status $(if ($null -ne $visualValidationResult -and [bool]$visualValidationResult.passed) { "completed" } else { "failed" }) `
    -Output ([pscustomobject]@{
        path = $visualValidationPath
        pdfPath = $(if (Test-Path -LiteralPath $strictPdfPath) { $strictPdfPath } else { $null })
        previewDir = $strictPreviewDir
        passed = $(if ($null -ne $visualValidationResult) { [bool]$visualValidationResult.passed } else { $false })
      }) `
    -Error $(if ($null -eq $visualValidationResult -or (-not [bool]$visualValidationResult.passed)) {
        [pscustomobject]@{
          code = "visual-validation-failed"
          message = "严格模式视觉验收未通过。"
          suggestion = "查看 visual-validation.json；确认 LibreOffice 与预览依赖可用后修复版式。"
        }
      } else { $null })
} else {
  Add-PipelineStage -Name "visual-validation" -Status "skipped" -Output ([pscustomobject]@{
      reason = "quality-mode-fast"
      recommendation = [string]$qualityRecommendation.recommendedMode
    })
}

$validationWarningSummary = @()
$validationErrorCodes = @()
$validationWarningCodes = @()
if ($null -ne $validationResult -and $validationResult.PSObject.Properties.Name -contains "findings") {
  $validationWarningSummary = @(
    $validationResult.findings |
      Where-Object { [string]$_.severity -eq "warning" } |
      ForEach-Object {
        [pscustomobject]@{
          severity = [string]$_.severity
          code = [string]$_.code
          category = $(if ($_.PSObject.Properties.Name -contains "category") { [string]$_.category } else { $null })
          message = [string]$_.message
          remediation = $(if ($_.PSObject.Properties.Name -contains "remediation") { [string]$_.remediation } else { $null })
        }
      }
  )
}
if ($null -ne $validationResult -and $validationResult.summary.PSObject.Properties.Name -contains "errorCodes") {
  $validationErrorCodes = @($validationResult.summary.errorCodes)
}
if ($null -ne $validationResult -and $validationResult.summary.PSObject.Properties.Name -contains "warningCodes") {
  $validationWarningCodes = @($validationResult.summary.warningCodes)
}

$pipelineFailedStages = @($script:pipelineStages | Where-Object { [string]$_.status -eq "failed" })
$pipelineStatus = if ($pipelineFailedStages.Count -gt 0) { "needs-fix" } else { "completed" }
$pipelineErrors = @(
  $pipelineFailedStages |
    Where-Object { $null -ne $_.error } |
    ForEach-Object {
      [pscustomobject]@{
        stage = [string]$_.name
        code = [string]$_.error.code
        message = [string]$_.error.message
        suggestion = [string]$_.error.suggestion
      }
    }
)
$pipelineTrace = [pscustomobject]@{
  schemaVersion = "1.0"
  status = $pipelineStatus
  currentStage = $(if ($script:pipelineStages.Count -gt 0) { [string]$script:pipelineStages[$script:pipelineStages.Count - 1].name } else { $null })
  stages = @($script:pipelineStages.ToArray())
  errors = $pipelineErrors
}
[System.IO.File]::WriteAllText($pipelineTracePath, ($pipelineTrace | ConvertTo-Json -Depth 12), (New-Object System.Text.UTF8Encoding($true)))

$summary = [pscustomobject]@{
  outputDir = $resolvedOutputDir
  pipelineMode = $PipelineMode
  qualityMode = $QualityMode
  detailLevel = $DetailLevel
  templateStyleMode = $TemplateStyleMode
  generationStatus = $pipelineStatus
  pipelineTracePath = $pipelineTracePath
  materialsAnalysisPath = $materialsPath
  contentPlanPath = $contentPlanPath
  reportProfileName = [string]$reportProfile.name
  reportProfilePath = $resolvedReportProfilePath
  templatePath = $resolvedTemplatePath
  sourceTemplatePath = $sourceTemplatePath
  templatePathDefaulted = $templatePathDefaulted
  templateSelectionSource = $templateSelectionSource
  builtInTemplateId = $resolvedBuiltInTemplateId
  templateContractPath = $templateContractPath
  templateCacheDir = $resolvedTemplateCacheDir
  templateCacheHit = $(if ($templateContract.PSObject.Properties.Name -contains "cache") { [bool]$templateContract.cache.hit } else { $false })
  templateRiskLevel = [string]$templateContract.risk.level
  templateRiskScore = [int]$templateContract.risk.score
  recommendedQualityMode = [string]$qualityRecommendation.recommendedMode
  qualityRecommendationReasons = @($qualityRecommendation.reasons)
  templateFrameDefaulted = $templateFrameDefaulted
  templateStylePreserved = (
    [string]::Equals($TemplateStyleMode, "preserve", [System.StringComparison]::OrdinalIgnoreCase) -and
    (-not $styleOutputRequested) -and
    (-not $shouldCreateTemplateFrameDocx)
  )
  fixedExperimentReportStyle = (Test-IsExperimentReportProfile -ReportProfileName ([string]$reportProfile.name) -ReportProfilePath $resolvedReportProfilePath)
  templateConversionStatus = $(if ($null -ne $templateConversion) { [string]$templateConversion.status } else { "none" })
  templateConversionConverter = $(if ($null -ne $templateConversion) { [string]$templateConversion.converter } else { "none" })
  convertedTemplatePath = $(if ($null -ne $templateConversion) { [string]$templateConversion.convertedTemplatePath } else { $null })
  reportPath = $resolvedReportPath
  reportInputMode = "path"
  metadataPath = $resolvedMetadataPath
  metadataInputMode = $metadataInputMode
  requirementsInputMode = $requirementsInputMode
  imageInputMode = $imageInputMode
  fieldMapPath = $resolvedFieldMapOutPath
  filledDocxPath = $resolvedFilledDocxOutPath
  filledOutlinePath = $filledOutlinePath
  imagePlanPath = $resolvedImagePlanOutPath
  imagePlanLowConfidenceCount = $imagePlanLowConfidenceCount
  imagePlanNeedsReview = $imagePlanNeedsReview
  imageMapPath = $resolvedImageMapOutPath
  imageManifestPath = $imageManifestPath
  filledDocxWithImagesPath = $resolvedFilledDocxWithImagesOutPath
  filledWithImagesOutlinePath = $filledWithImagesOutlinePath
  courseDesignTablesDocxPath = $resolvedCourseDesignTablesDocxOutPath
  courseDesignTablesInserted = $(if ($null -ne $courseDesignTablesResult -and $courseDesignTablesResult.PSObject.Properties.Name -contains "inserted") { [bool]$courseDesignTablesResult.inserted } else { $null })
  courseDesignTablesCount = $(if ($null -ne $courseDesignTablesResult -and $courseDesignTablesResult.PSObject.Properties.Name -contains "tableCount") { [int]$courseDesignTablesResult.tableCount } else { $null })
  styledDocxPath = $resolvedStyledDocxOutPath
  preFrameFinalDocxPath = $preFrameFinalDocxPath
  styledOutlinePath = $styledOutlinePath
  templateFrameDocxPath = $resolvedTemplateFrameDocxOutPath
  layoutCheckPath = $(if ($shouldRunLayoutCheck) { $layoutCheckPath } else { $null })
  layoutCheckPassed = $(if ($null -ne $layoutCheckResult) { [bool]$layoutCheckResult.passed } else { $null })
  layoutCheckMessage = $(if ($null -ne $layoutCheckResult -and $layoutCheckResult.PSObject.Properties.Name -contains "message") { [string]$layoutCheckResult.message } else { $null })
  layoutCheckErrorCount = $(if ($null -ne $layoutCheckResult) { [int]$layoutCheckResult.summary.errorCount } else { $null })
  layoutCheckWarningCount = $(if ($null -ne $layoutCheckResult) { [int]$layoutCheckResult.summary.warningCount } else { $null })
  formatValidationPath = $formatValidationPath
  formatValidationPassed = [bool]$formatValidationResult.passed
  formatValidationRequired = $formatValidationRequired
  formatValidationFailedCodes = @($formatValidationResult.summary.failedCodes)
  visualValidationPath = $(if ([string]::Equals($QualityMode, "strict", [System.StringComparison]::OrdinalIgnoreCase)) { $visualValidationPath } else { $null })
  visualValidationPassed = $(if ($null -ne $visualValidationResult) { [bool]$visualValidationResult.passed } else { $null })
  strictPdfPath = $(if (Test-Path -LiteralPath $strictPdfPath) { $strictPdfPath } else { $null })
  strictPreviewDir = $(if (Test-Path -LiteralPath $strictPreviewDir) { $strictPreviewDir } else { $null })
  expectedLayoutImageCount = $(if ($expectedLayoutImageCount -ge 0) { $expectedLayoutImageCount } else { $null })
  expectedLayoutCaptionCount = $(if ($expectedLayoutCaptionCount -ge 0) { $expectedLayoutCaptionCount } else { $null })
  requestedStyleProfile = $(if ($styleOutputRequested) { $effectiveStyleProfile } else { $null })
  styleProfilePath = $(if ($null -ne $styleResult) { [string]$styleResult.profilePath } elseif (-not [string]::IsNullOrWhiteSpace($StyleProfilePath)) { (Resolve-Path -LiteralPath $StyleProfilePath).Path } else { $null })
  styleProfile = $(if ($null -ne $styleResult) { [string]$styleResult.styleProfile } else { $null })
  resolvedStyleProfile = $(if ($null -ne $styleResult) { [string]$styleResult.resolvedProfile } else { $null })
  styleProfileReason = $(if ($null -ne $styleResult) { [string]$styleResult.profileReason } else { $null })
  appliedStyleSettings = $(if ($null -ne $styleResult) { $styleResult.appliedSettings } else { $null })
  finalDocxPath = $finalDocxPath
  validationPath = $(if ($shouldRunValidation) { $validationPath } else { $null })
  validationPassed = $(if ($null -ne $validationResult) { [bool]$validationResult.passed } else { $null })
  validationErrorCount = $(if ($null -ne $validationResult) { [int]$validationResult.summary.errorCount } else { $null })
  validationWarningCount = $(if ($null -ne $validationResult) { [int]$validationResult.summary.warningCount } else { $null })
  validationPaginationRiskCount = $(if ($null -ne $validationResult -and $validationResult.summary.PSObject.Properties.Name -contains "paginationRiskCount") { [int]$validationResult.summary.paginationRiskCount } else { $null })
  validationPaginationRiskThresholds = $(if ($null -ne $validationResult -and $validationResult.summary.PSObject.Properties.Name -contains "paginationRiskThresholds") { $validationResult.summary.paginationRiskThresholds } else { $null })
  validationPaginationRiskRemediations = $(if ($null -ne $validationResult -and $validationResult.summary.PSObject.Properties.Name -contains "paginationRiskRemediations") { $validationResult.summary.paginationRiskRemediations } else { $null })
  validationStructuralIssueCount = $(if ($null -ne $validationResult -and $validationResult.summary.PSObject.Properties.Name -contains "structuralIssueCount") { [int]$validationResult.summary.structuralIssueCount } else { $null })
  validationFindingCountsByCode = $(if ($null -ne $validationResult -and $validationResult.summary.PSObject.Properties.Name -contains "findingCountsByCode") { $validationResult.summary.findingCountsByCode } else { $null })
  validationFindingCountsByCategory = $(if ($null -ne $validationResult -and $validationResult.summary.PSObject.Properties.Name -contains "findingCountsByCategory") { $validationResult.summary.findingCountsByCategory } else { $null })
  validationErrorCodes = $validationErrorCodes
  validationWarningCodes = $validationWarningCodes
  validationWarningSummary = $validationWarningSummary
}
[System.IO.File]::WriteAllText($summaryPath, ($summary | ConvertTo-Json -Depth 8), (New-Object System.Text.UTF8Encoding($true)))

Write-Output ("Field-map path: {0}" -f $resolvedFieldMapOutPath)
Write-Output ("Filled docx path: {0}" -f $resolvedFilledDocxOutPath)
if ($null -ne $resolvedFilledDocxWithImagesOutPath) {
  if (-not [string]::IsNullOrWhiteSpace($resolvedImagePlanOutPath)) {
    Write-Output ("Image-plan path: {0}" -f $resolvedImagePlanOutPath)
  }
  Write-Output ("Image-map path: {0}" -f $resolvedImageMapOutPath)
  Write-Output ("Filled docx with images path: {0}" -f $resolvedFilledDocxWithImagesOutPath)
}
if ($null -ne $resolvedCourseDesignTablesDocxOutPath) {
  Write-Output ("Course-design tables docx path: {0}" -f $resolvedCourseDesignTablesDocxOutPath)
}
if ($null -ne $resolvedStyledDocxOutPath) {
  Write-Output ("Styled docx path: {0}" -f $resolvedStyledDocxOutPath)
}
if ($null -ne $resolvedTemplateFrameDocxOutPath) {
  Write-Output ("Template-frame docx path: {0}" -f $resolvedTemplateFrameDocxOutPath)
}
Write-Output ("Final docx path: {0}" -f $finalDocxPath)
if ($shouldRunLayoutCheck) {
  Write-Output ("Layout check path: {0}" -f $layoutCheckPath)
}
Write-Output ("Template contract path: {0}" -f $templateContractPath)
Write-Output ("Format validation path: {0}" -f $formatValidationPath)
Write-Output ("Image manifest path: {0}" -f $imageManifestPath)
Write-Output ("Pipeline trace path: {0}" -f $pipelineTracePath)
if ([string]::Equals($QualityMode, "strict", [System.StringComparison]::OrdinalIgnoreCase)) {
  Write-Output ("Visual validation path: {0}" -f $visualValidationPath)
}
Write-Output ("Summary path: {0}" -f $summaryPath)

if ($null -ne $validationResult -and -not $validationResult.passed) {
  throw "Report validation failed. See $validationPath"
}
if (
  $null -ne $layoutCheckResult -and
  (-not [bool]$layoutCheckResult.passed) -and
  [string]::Equals($QualityMode, "strict", [System.StringComparison]::OrdinalIgnoreCase)
) {
  throw "Strict layout validation failed. See $layoutCheckPath"
}
if (
  (-not [bool]$formatValidationResult.passed) -and
  (
    $FailOnFormatValidation -or
    (
      $formatValidationRequired -and
      [string]::Equals($QualityMode, "strict", [System.StringComparison]::OrdinalIgnoreCase)
    )
  )
) {
  throw "Template format validation failed. See $formatValidationPath"
}
if (
  [string]::Equals($QualityMode, "strict", [System.StringComparison]::OrdinalIgnoreCase) -and
  ($null -eq $visualValidationResult -or (-not [bool]$visualValidationResult.passed))
) {
  throw "Strict visual validation failed. See $visualValidationPath"
}
