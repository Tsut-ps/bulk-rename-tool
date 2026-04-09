[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [string]$CsvPath,

    [Parameter(Mandatory = $true)]
    [string]$TargetDirectory,

    [ValidateSet("Preview", "Apply")]
    [string]$Mode = "Preview",

    [string]$SourceColumn = "File",

    [string]$TargetColumn,

    [string]$NameTemplate,

    [string]$PreviewPath,

    [string]$ReportCsvPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$scriptRoot = Split-Path -Parent $PSCommandPath
Import-Module (Join-Path -Path $scriptRoot -ChildPath "RenameCore.psm1") -Force

function Ensure-OutputDirectory {
    param([Parameter(Mandatory = $true)][string]$DirectoryPath)

    if (-not (Test-Path -LiteralPath $DirectoryPath)) {
        New-Item -ItemType Directory -Path $DirectoryPath -Force | Out-Null
    }
}

function Get-ModeLabel {
    param([Parameter(Mandatory = $true)][string]$ModeName)

    if ($ModeName -eq "Preview") {
        return "プレビュー"
    }

    return "適用"
}

function HtmlEncode {
    param(
        [AllowNull()]
        [string]$Value
    )

    if ($null -eq $Value) {
        return ""
    }

    return [System.Security.SecurityElement]::Escape($Value)
}

function New-PreviewHtml {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Rows,

        [Parameter(Mandatory = $true)]
        [string]$CsvPath,

        [Parameter(Mandatory = $true)]
        [string]$TargetDirectory,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath,

        [Parameter(Mandatory = $true)]
        [string]$Mode
    )

    $rowList = @($Rows)
    $summary = Get-RenameSummary -Rows $rowList

    $summaryMap = [ordered]@{
        "生成日時" = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        "モード" = $Mode
        "CSVパス" = $CsvPath
        "対象フォルダ" = $TargetDirectory
        "総件数" = $summary.Total
        "変更対象" = $summary.Changed
        "変更なし" = $summary.Unchanged
        "元ファイルなし" = $summary.Missing
        "要確認" = $summary.Blocked
        "実行可能" = $summary.Ready
    }

    $summaryRows = foreach ($entry in $summaryMap.GetEnumerator()) {
        "<tr><th>$($entry.Key)</th><td>$(HtmlEncode([string]$entry.Value))</td></tr>"
    }

    $bodyRows = foreach ($row in $rowList) {
        $issues = if ($row.Issues.Count -gt 0) {
            ($row.Issues | ForEach-Object { HtmlEncode($_) }) -join "<br>"
        } else {
            ""
        }

        @"
<tr class="$($row.StatusCss)">
  <td>$($row.No)</td>
  <td>$(HtmlEncode($row.SourceName))</td>
  <td>$(HtmlEncode($row.TargetName))</td>
  <td>$(HtmlEncode($row.Action))</td>
  <td>$(HtmlEncode($row.Status))</td>
  <td>$issues</td>
</tr>
"@
    }

    $html = @"
<!doctype html>
<html lang="ja">
<head>
<meta charset="utf-8">
<title>一括リネーム プレビュー</title>
<style>
body { font-family: "Yu Gothic UI", "Segoe UI", sans-serif; margin: 24px; color: #1f2937; background: #f7f7f5; }
h1 { margin-bottom: 8px; }
p { margin-top: 0; }
table { border-collapse: collapse; width: 100%; background: white; }
th, td { border: 1px solid #d1d5db; padding: 8px 10px; text-align: left; vertical-align: top; }
th { background: #e5e7eb; }
.summary { width: auto; margin-bottom: 20px; }
.ready { background: #edfdf3; }
.blocked { background: #fff1f2; }
.missing { background: #fff7ed; }
.unchanged { background: #f9fafb; }
code { background: #eef2ff; padding: 1px 4px; border-radius: 4px; }
</style>
</head>
<body>
<h1>一括リネーム プレビュー</h1>
<p>差分を確認して、問題がなければ <code>-Mode Apply</code> で実行してください。</p>
<table class="summary">
<tbody>
$($summaryRows -join "`n")
</tbody>
</table>
<table>
<thead>
<tr>
  <th>番号</th>
  <th>現在の名前</th>
  <th>変更後の名前</th>
  <th>操作</th>
  <th>状態</th>
  <th>問題点</th>
</tr>
</thead>
<tbody>
$($bodyRows -join "`n")
</tbody>
</table>
</body>
</html>
"@

    Set-Content -LiteralPath $OutputPath -Value $html -Encoding UTF8
}

$csvFullPath = Resolve-RenamePath -Path $CsvPath
$targetFullPath = Resolve-RenamePath -Path $TargetDirectory -Directory
$outputsDir = Join-Path -Path $scriptRoot -ChildPath "outputs"
$modeLabel = Get-ModeLabel -ModeName $Mode

if (-not $PreviewPath) {
    Ensure-OutputDirectory -DirectoryPath $outputsDir
    $PreviewPath = Join-Path -Path $outputsDir -ChildPath "bulk-rename-preview.html"
}

if (-not $ReportCsvPath) {
    Ensure-OutputDirectory -DirectoryPath $outputsDir
    $ReportCsvPath = Join-Path -Path $outputsDir -ChildPath "bulk-rename-preview.csv"
}

$previewRows = @(New-RenamePlan `
    -CsvPath $csvFullPath `
    -TargetDirectory $targetFullPath `
    -SourceColumn $SourceColumn `
    -TargetColumn $TargetColumn `
    -NameTemplate $NameTemplate)

$reportRows = $previewRows | Select-Object `
    @{Name = "番号順"; Expression = { $_.Index } }, `
    @{Name = "番号"; Expression = { $_.No } }, `
    @{Name = "変更元ファイル名"; Expression = { $_.SourceName } }, `
    @{Name = "変更後ファイル名"; Expression = { $_.TargetName } }, `
    @{Name = "操作"; Expression = { $_.Action } }, `
    @{Name = "状態"; Expression = { $_.Status } }, `
    @{Name = "元ファイル存在"; Expression = { $_.SourceExists } }, `
    @{Name = "問題点"; Expression = { $_.Issues -join " | " } }, `
    @{Name = "変更元パス"; Expression = { $_.SourcePath } }, `
    @{Name = "変更後パス"; Expression = { $_.TargetPath } }

$reportRows | Export-Csv -LiteralPath $ReportCsvPath -NoTypeInformation -Encoding UTF8
New-PreviewHtml -Rows $previewRows -CsvPath $csvFullPath -TargetDirectory $targetFullPath -OutputPath $PreviewPath -Mode $modeLabel

$blockingRows = @($previewRows | Where-Object { $_.Status -in @("要確認", "元ファイルなし") })
$renameRows = @($previewRows | Where-Object { $_.Action -eq "変更" -and $_.Status -eq "実行可能" })

if ($Mode -eq "Apply") {
    if ($blockingRows.Count -gt 0) {
        throw "問題のある行が残っているため適用できません。プレビュー結果を確認してください。"
    }

    foreach ($row in $renameRows) {
        if ($PSCmdlet.ShouldProcess($row.SourcePath, "$($row.TargetName) に名前を変更")) {
            Rename-Item -LiteralPath $row.SourcePath -NewName $row.TargetName
        }
    }
}

[pscustomobject]@{
    モード = $modeLabel
    CSVパス = $csvFullPath
    対象フォルダ = $targetFullPath
    HTML出力 = $PreviewPath
    CSV出力 = $ReportCsvPath
    総件数 = $previewRows.Count
    実行可能件数 = $renameRows.Count
    要確認件数 = $blockingRows.Count
    変更なし件数 = @($previewRows | Where-Object { $_.Status -eq "変更なし" }).Count
}
