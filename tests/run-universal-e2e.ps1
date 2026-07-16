[CmdletBinding()]
param(
  [string]$OutputDir
)

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

function Write-Utf8BomText {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Path,

    [Parameter(Mandatory = $true)]
    [string]$Text
  )

  $parent = Split-Path -Parent $Path
  if (-not [string]::IsNullOrWhiteSpace($parent)) {
    New-Item -ItemType Directory -Path $parent -Force | Out-Null
  }
  [System.IO.File]::WriteAllText($Path, $Text, (New-Object System.Text.UTF8Encoding($true)))
}

function New-CaseInputs {
  param(
    [Parameter(Mandatory = $true)]
    [string]$CaseDir,

    [Parameter(Mandatory = $true)]
    [string]$CourseName,

    [Parameter(Mandatory = $true)]
    [string]$ExperimentName,

    [Parameter(Mandatory = $true)]
    [string]$DomainText
  )

  New-Item -ItemType Directory -Path $CaseDir -Force | Out-Null
  $reportPath = Join-Path $CaseDir "report.txt"
  $metadataPath = Join-Path $CaseDir "metadata.json"
  $reportText = @"
实验目的
理解 $CourseName 中与“$ExperimentName”相关的核心概念，能够完成环境配置、操作验证与结果分析。

实验环境
Windows 11、课程实验软件、浏览器与命令行工具。所有步骤均在本地隔离目录中完成。

实验原理或任务要求
$DomainText
本实验要求记录关键参数、执行顺序、观察结果和异常处理方法，并用可复核的证据说明任务是否完成。

实验步骤
1. 阅读任务材料，提取输入条件、约束和验收标准。
2. 准备实验环境，记录软件版本与关键配置。
3. 按顺序执行核心操作，对每一步保存必要的输出结果。
4. 对照预期结果检查关键状态；若出现异常，定位配置、权限、依赖或输入数据问题。
5. 整理结果并复查报告中的参数、图片与结论是否一致。

实验结果
核心任务已完成，关键状态与预期一致。结果记录能够对应到具体步骤，未使用无法验证的占位内容。

问题分析
实验过程中重点检查了输入合法性、环境差异和执行顺序。对于可复现问题，采用逐项排除法确认根因。

实验总结
本次实验形成了从材料解析、过程执行、结果验证到问题复盘的完整闭环，并保留了可继续改进的记录。
"@
  Write-Utf8BomText -Path $reportPath -Text $reportText

  $metadata = [ordered]@{
    姓名 = "示例学生"
    学号 = "20260001"
    班级 = "计科2401"
    课程名称 = $CourseName
    实验名称 = $ExperimentName
    指导教师 = "示例教师"
    实验性质 = "综合性实验"
    日期 = "2026-07-16"
    实验时间 = "2026-07-16"
    实验地点 = "实验室 A"
    Name = "示例学生"
    StudentId = "20260001"
    ClassName = "计科2401"
    CourseName = $CourseName
    ExperimentName = $ExperimentName
    TeacherName = "示例教师"
    ExperimentProperty = "综合性实验"
    ExperimentDate = "2026-07-16"
    ExperimentLocation = "实验室 A"
  }
  Write-Utf8BomText -Path $metadataPath -Text ($metadata | ConvertTo-Json -Depth 5)
  return [pscustomobject]@{
    reportPath = $reportPath
    metadataPath = $metadataPath
  }
}

function Invoke-E2ECase {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Name,

    [Parameter(Mandatory = $true)]
    [string]$CourseName,

    [Parameter(Mandatory = $true)]
    [string]$ExperimentName,

    [Parameter(Mandatory = $true)]
    [string]$DomainText,

    [AllowNull()]
    [string]$TemplatePath,

    [string[]]$ImagePaths = @(),

    [int]$ExpectedImageCount = -1,

    [switch]$ExpectTemplatePreserved
  )

  $caseDir = Join-Path $resolvedOutputDir $Name
  $inputs = New-CaseInputs -CaseDir $caseDir -CourseName $CourseName -ExperimentName $ExperimentName -DomainText $DomainText
  $buildDir = Join-Path $caseDir "build"
  $params = @{
    ReportPath = $inputs.reportPath
    MetadataPath = $inputs.metadataPath
    OutputDir = $buildDir
    ReportProfileName = "experiment-report"
    PipelineMode = "fast"
    QualityMode = "fast"
    DetailLevel = "standard"
    TemplateStyleMode = "preserve"
  }
  if (-not [string]::IsNullOrWhiteSpace($TemplatePath)) {
    $params.TemplatePath = $TemplatePath
  }
  if ($ImagePaths.Count -gt 0) {
    $params.ImagePaths = $ImagePaths
  }
  if ($ExpectedImageCount -ge 0) {
    $params.RequestedImageCount = $ExpectedImageCount
  }

  & (Join-Path $repoRoot "scripts\build-report.ps1") @params | Out-Null
  $summaryPath = Join-Path $buildDir "summary.json"
  Assert-True -Condition (Test-Path -LiteralPath $summaryPath) -Message "$Name did not create summary.json."
  $summary = (Get-Content -LiteralPath $summaryPath -Raw -Encoding UTF8) | ConvertFrom-Json
  Assert-True -Condition (Test-Path -LiteralPath ([string]$summary.finalDocxPath) -PathType Leaf) -Message "$Name did not create the final DOCX."
  foreach ($artifact in @("pipelineTracePath", "materialsAnalysisPath", "contentPlanPath", "templateContractPath", "formatValidationPath", "imageManifestPath")) {
    Assert-True -Condition (Test-Path -LiteralPath ([string]$summary.$artifact) -PathType Leaf) -Message "$Name is missing $artifact."
  }
  if ($ExpectTemplatePreserved) {
    Assert-True -Condition ([bool]$summary.templateStylePreserved) -Message "$Name did not preserve the custom template mode."
  }
  $manifest = (Get-Content -LiteralPath ([string]$summary.imageManifestPath) -Raw -Encoding UTF8) | ConvertFrom-Json
  $expectedCount = if ($ExpectedImageCount -ge 0) { $ExpectedImageCount } else { $ImagePaths.Count }
  Assert-True -Condition ([int]$manifest.selectedCount -eq $expectedCount) -Message "$Name selected an unexpected image count."

  return [pscustomobject]@{
    name = $Name
    courseName = $CourseName
    templatePath = $summary.templatePath
    templateStylePreserved = [bool]$summary.templateStylePreserved
    generationStatus = [string]$summary.generationStatus
    imageCount = [int]$manifest.selectedCount
    finalDocxPath = [string]$summary.finalDocxPath
    summaryPath = $summaryPath
  }
}

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
if ([string]::IsNullOrWhiteSpace($OutputDir)) {
  $OutputDir = Join-Path $repoRoot ("tests-output\universal-e2e-" + (Get-Date -Format "yyyyMMdd-HHmmss"))
}
$resolvedOutputDir = [System.IO.Path]::GetFullPath($OutputDir)
New-Item -ItemType Directory -Path $resolvedOutputDir -Force | Out-Null

$fixturesDir = Join-Path $resolvedOutputDir "fixtures"
& (Join-Path $repoRoot "scripts\new-report-test-fixtures.ps1") -OutputDir $fixturesDir | Out-Null
$fixtureManifest = (Get-Content -LiteralPath (Join-Path $fixturesDir "fixture-manifest.json") -Raw -Encoding UTF8) | ConvertFrom-Json
function Get-FixturePath([string]$Name) {
  return [string](@($fixtureManifest.fixtures | Where-Object { [string]$_.name -eq $Name })[0].path)
}

$imageOne = Join-Path $repoRoot "demo\assets\result-ping.png"
$imageTwo = Join-Path $repoRoot "demo\assets\result-arp.png"
$results = New-Object System.Collections.Generic.List[object]
$oldTemplatePath = $env:EXPERIMENT_REPORT_TEMPLATE_PATH
try {
  $env:EXPERIMENT_REPORT_TEMPLATE_PATH = Join-Path $repoRoot "examples\report-templates\experiment-report-template.docx"

  [void]$results.Add((Invoke-E2ECase `
        -Name "network" `
        -CourseName "计算机网络" `
        -ExperimentName "局域网连通性与常用命令验证" `
        -DomainText "通过 ipconfig、ping 与 arp 等命令观察主机地址、连通性和邻居缓存，解释输出字段与网络状态的关系。" `
        -TemplatePath (Get-FixturePath "repository-default-template") `
        -ImagePaths @($imageOne, $imageOne) `
        -ExpectedImageCount 1 `
        -ExpectTemplatePreserved))

  [void]$results.Add((Invoke-E2ECase `
        -Name "operating-system" `
        -CourseName "操作系统" `
        -ExperimentName "进程调度算法分析" `
        -DomainText "比较先来先服务、短作业优先与时间片轮转的调度过程，记录等待时间、周转时间和响应特征。" `
        -TemplatePath (Get-FixturePath "five-column-metadata") `
        -ExpectTemplatePreserved))

  [void]$results.Add((Invoke-E2ECase `
        -Name "java-web" `
        -CourseName "Java Web 应用开发" `
        -ExperimentName "用户登录与会话管理" `
        -DomainText "实现请求参数校验、登录状态保存、受保护页面访问与退出流程，分析会话生命周期和常见安全边界。" `
        -TemplatePath (Get-FixturePath "cover-body-sections") `
        -ExpectTemplatePreserved))

  [void]$results.Add((Invoke-E2ECase `
        -Name "no-template" `
        -CourseName "数据库原理" `
        -ExperimentName "事务隔离级别验证" `
        -DomainText "通过两个并发会话观察脏读、不可重复读或幻读现象，说明隔离级别对一致性与并发性的影响。"))

  [void]$results.Add((Invoke-E2ECase `
        -Name "non-default-template" `
        -CourseName "面向复杂工程系统的软件体系结构与综合实践" `
        -ExperimentName "组件边界与接口协作验证" `
        -DomainText "识别模块职责、接口契约和依赖方向，验证关键调用链并分析长课程名称下的信息表适配风险。" `
        -TemplatePath (Get-FixturePath "long-course-name-cell-pressure") `
        -ExpectTemplatePreserved))

  $strictDir = Join-Path $resolvedOutputDir "strict-mode"
  $strictInputs = New-CaseInputs `
    -CaseDir $strictDir `
    -CourseName "软件测试" `
    -ExperimentName "严格模式视觉验收" `
    -DomainText "验证严格模式在 PDF 导出或逐页视觉检查失败时返回 needs-fix，而不是成功状态。"
  $strictBuildDir = Join-Path $strictDir "build"
  $strictError = $null
  try {
    & (Join-Path $repoRoot "scripts\build-report.ps1") `
      -ReportPath $strictInputs.reportPath `
      -MetadataPath $strictInputs.metadataPath `
      -OutputDir $strictBuildDir `
      -QualityMode strict `
      -TemplateStyleMode preserve | Out-Null
  } catch {
    $strictError = $_.Exception.Message
  }
  $strictSummary = (Get-Content -LiteralPath (Join-Path $strictBuildDir "summary.json") -Raw -Encoding UTF8) | ConvertFrom-Json
  Assert-True -Condition (
    ([bool]$strictSummary.visualValidationPassed) -or
    (
      [string]$strictSummary.generationStatus -eq "needs-fix" -and
      -not [string]::IsNullOrWhiteSpace($strictError)
    )
  ) -Message "Strict mode reported success after visual validation failure."

  $e2eSummary = [pscustomobject]@{
    schemaVersion = "1.0"
    outputDir = $resolvedOutputDir
    cases = @($results.ToArray())
    strictMode = [pscustomobject]@{
      generationStatus = [string]$strictSummary.generationStatus
      visualValidationPassed = $strictSummary.visualValidationPassed
      visualValidationPath = [string]$strictSummary.visualValidationPath
      error = $strictError
    }
  }
  $summaryPath = Join-Path $resolvedOutputDir "e2e-summary.json"
  Write-Utf8BomText -Path $summaryPath -Text ($e2eSummary | ConvertTo-Json -Depth 8)
  Write-Output ("Universal E2E summary: {0}" -f $summaryPath)
} finally {
  $env:EXPERIMENT_REPORT_TEMPLATE_PATH = $oldTemplatePath
}
