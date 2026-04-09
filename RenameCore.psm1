Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Resolve-RenamePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [switch]$Directory
    )

    $item = Get-Item -LiteralPath $Path
    if ($Directory -and -not $item.PSIsContainer) {
        throw "フォルダを指定してください: $Path"
    }
    if (-not $Directory -and $item.PSIsContainer) {
        throw "ファイルを指定してください: $Path"
    }
    return $item.FullName
}

function Get-SafeLeafName {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Name
    )

    $invalidChars = [IO.Path]::GetInvalidFileNameChars()
    $builder = New-Object System.Text.StringBuilder

    foreach ($char in $Name.ToCharArray()) {
        if ($invalidChars -contains $char) {
            [void]$builder.Append("_")
        } else {
            [void]$builder.Append($char)
        }
    }

    $safe = $builder.ToString().Trim()
    $safe = $safe.TrimEnd(".")
    return $safe
}

function Expand-RenameTemplate {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Template,

        [Parameter(Mandatory = $true)]
        [psobject]$Row
    )

    $rowProps = @{}
    foreach ($prop in $Row.PSObject.Properties) {
        $rowProps[$prop.Name] = [string]$prop.Value
    }

    if ($rowProps.ContainsKey("File")) {
        $rowProps["FileBase"] = [IO.Path]::GetFileNameWithoutExtension($rowProps["File"])
        $rowProps["Ext"] = [IO.Path]::GetExtension($rowProps["File"])
    }

    $matches = [regex]::Matches($Template, "\{(?<name>[A-Za-z0-9_]+)(:(?<format>[^}]+))?\}")
    $processedTokens = @{}
    $expanded = $Template

    foreach ($match in $matches) {
        $fullToken = $match.Value
        if ($processedTokens.ContainsKey($fullToken)) {
            continue
        }
        $processedTokens[$fullToken] = $true

        $name = $match.Groups["name"].Value
        $format = $match.Groups["format"].Value

        if (-not $rowProps.ContainsKey($name)) {
            throw "テンプレート内の `{$name}` に対応するCSV列が見つかりません。"
        }

        $rawValue = $rowProps[$name]
        if ($format) {
            $number = 0
            if (-not [int]::TryParse($rawValue, [ref]$number)) {
                throw ("テンプレート内の `{0}:{1}` は数値列が必要ですが、実際の値は `{2}` でした。" -f $name, $format, $rawValue)
            }
            $replacement = $number.ToString($format)
        } else {
            $replacement = $rawValue
        }

        $expanded = $expanded.Replace($fullToken, $replacement)
    }

    return $expanded
}

function Add-RenameIssue {
    param(
        [System.Collections.Generic.List[string]]$Issues,
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    $Issues.Add($Message) | Out-Null
}

function New-RenamePlan {
    param(
        [Parameter(Mandatory = $true)]
        [string]$CsvPath,

        [Parameter(Mandatory = $true)]
        [string]$TargetDirectory,

        [string]$SourceColumn = "File",

        [string]$TargetColumn,

        [string]$NameTemplate
    )

    if ([string]::IsNullOrWhiteSpace($TargetColumn) -and [string]::IsNullOrWhiteSpace($NameTemplate)) {
        throw "変更先列名または名前テンプレートを指定してください。"
    }

    $csvRows = @(Import-Csv -LiteralPath $CsvPath)
    if ($csvRows.Count -eq 0) {
        throw "CSVにデータ行がありません。"
    }

    if (-not ($csvRows[0].PSObject.Properties.Name -contains $SourceColumn)) {
        throw ("CSV列 '{0}' が見つかりません。" -f $SourceColumn)
    }

    if ($TargetColumn -and -not ($csvRows[0].PSObject.Properties.Name -contains $TargetColumn)) {
        throw ("CSV列 '{0}' が見つかりません。" -f $TargetColumn)
    }

    $previewRows = New-Object System.Collections.Generic.List[object]
    $seenSources = @{}
    $seenTargets = @{}

    for ($index = 0; $index -lt $csvRows.Count; $index++) {
        $row = $csvRows[$index]
        $sourceName = ([string]$row.$SourceColumn).Trim()
        $issues = New-Object 'System.Collections.Generic.List[string]'

        if ([string]::IsNullOrWhiteSpace($sourceName)) {
            Add-RenameIssue -Issues $issues -Message "変更元のファイル名が空です。"
        }

        if ($TargetColumn) {
            $targetLeaf = [string]$row.$TargetColumn
        } else {
            $targetLeaf = Expand-RenameTemplate -Template $NameTemplate -Row $row
        }

        $targetLeaf = Get-SafeLeafName -Name $targetLeaf
        $sourceLeaf = Get-SafeLeafName -Name $sourceName

        if ([string]::IsNullOrWhiteSpace($targetLeaf)) {
            Add-RenameIssue -Issues $issues -Message "変更後のファイル名が空になりました。"
        }

        if ($targetLeaf.Length -gt 255) {
            Add-RenameIssue -Issues $issues -Message "変更後のファイル名が255文字を超えています。"
        }

        if ($sourceLeaf -ne $sourceName) {
            Add-RenameIssue -Issues $issues -Message "変更元のファイル名に使えない文字が含まれています。"
        }

        $sourcePath = Join-Path -Path $TargetDirectory -ChildPath $sourceName
        $targetPath = Join-Path -Path $TargetDirectory -ChildPath $targetLeaf
        $sourceExists = Test-Path -LiteralPath $sourcePath

        if (-not $sourceExists) {
            Add-RenameIssue -Issues $issues -Message "対象フォルダ内に変更元ファイルが見つかりません。"
        }

        $sourceKey = $sourceName.ToLowerInvariant()
        if ($seenSources.ContainsKey($sourceKey)) {
            Add-RenameIssue -Issues $issues -Message "同じ変更元ファイル名がCSV内で重複しています。"
        } else {
            $seenSources[$sourceKey] = $true
        }

        $targetKey = $targetLeaf.ToLowerInvariant()
        if ($seenTargets.ContainsKey($targetKey)) {
            Add-RenameIssue -Issues $issues -Message "同じ変更後ファイル名に割り当てられている行が複数あります。"
        } else {
            $seenTargets[$targetKey] = $true
        }

        $action = if ($sourceName -ieq $targetLeaf) { "変更なし" } else { "変更" }

        if ($action -eq "変更" -and (Test-Path -LiteralPath $targetPath) -and ($sourceName -ine $targetLeaf)) {
            Add-RenameIssue -Issues $issues -Message "変更後ファイル名と同じ名前のファイルが既に存在します。"
        }

        $status = "実行可能"
        $statusCss = "ready"
        if ($issues.Count -gt 0) {
            $status = "要確認"
            $statusCss = "blocked"
            if (-not $sourceExists) {
                $status = "元ファイルなし"
                $statusCss = "missing"
            }
        } elseif ($action -eq "変更なし") {
            $status = "変更なし"
            $statusCss = "unchanged"
        }

        $previewRows.Add([pscustomobject]@{
            Index = $index + 1
            No = if ($row.PSObject.Properties.Name -contains "No") { [string]$row.No } else { [string]($index + 1) }
            SourceName = $sourceName
            TargetName = $targetLeaf
            Action = $action
            Status = $status
            StatusCss = $statusCss
            SourceExists = $sourceExists
            SourcePath = $sourcePath
            TargetPath = $targetPath
            Issues = @($issues)
        }) | Out-Null
    }

    return @($previewRows.ToArray())
}

function Get-RenameSummary {
    param([Parameter(Mandatory = $true)][object[]]$Rows)

    $rowList = @($Rows)
    return [pscustomobject]@{
        Total = $rowList.Count
        Ready = @($rowList | Where-Object { $_.Status -eq "実行可能" }).Count
        Blocked = @($rowList | Where-Object { $_.Status -in @("要確認", "元ファイルなし") }).Count
        Unchanged = @($rowList | Where-Object { $_.Status -eq "変更なし" }).Count
        Missing = @($rowList | Where-Object { $_.Status -eq "元ファイルなし" }).Count
        Changed = @($rowList | Where-Object { $_.Action -eq "変更" }).Count
    }
}

function Get-RenameSummaryText {
    param([Parameter(Mandatory = $true)][object[]]$Rows)

    $summary = Get-RenameSummary -Rows $Rows
    return "総件数: $($summary.Total)    実行可能: $($summary.Ready)    要確認: $($summary.Blocked)    変更なし: $($summary.Unchanged)"
}

Export-ModuleMember -Function Resolve-RenamePath, Get-SafeLeafName, Expand-RenameTemplate, New-RenamePlan, Get-RenameSummary, Get-RenameSummaryText
