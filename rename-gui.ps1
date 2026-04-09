Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$scriptRoot = Split-Path -Parent $PSCommandPath
Import-Module (Join-Path -Path $scriptRoot -ChildPath "RenameCore.psm1") -Force

$theme = @{
    WindowBg = [System.Drawing.Color]::FromArgb(11, 18, 32)
    PanelBg = [System.Drawing.Color]::FromArgb(15, 23, 42)
    InputBg = [System.Drawing.Color]::FromArgb(30, 41, 59)
    HeaderBg = [System.Drawing.Color]::FromArgb(30, 41, 59)
    GridBg = [System.Drawing.Color]::FromArgb(15, 23, 42)
    GridAltBg = [System.Drawing.Color]::FromArgb(17, 24, 39)
    GridLine = [System.Drawing.Color]::FromArgb(51, 65, 85)
    Title = [System.Drawing.Color]::FromArgb(241, 245, 249)
    Text = [System.Drawing.Color]::FromArgb(226, 232, 240)
    Muted = [System.Drawing.Color]::FromArgb(148, 163, 184)
    Accent = [System.Drawing.Color]::FromArgb(16, 185, 129)
    AccentDark = [System.Drawing.Color]::FromArgb(6, 78, 59)
    WarningDark = [System.Drawing.Color]::FromArgb(120, 53, 15)
    WarningText = [System.Drawing.Color]::FromArgb(255, 237, 213)
    DangerDark = [System.Drawing.Color]::FromArgb(127, 29, 29)
    DangerText = [System.Drawing.Color]::FromArgb(254, 226, 226)
    Info = [System.Drawing.Color]::FromArgb(96, 165, 250)
    SuccessText = [System.Drawing.Color]::FromArgb(167, 243, 208)
    WarningTextStrong = [System.Drawing.Color]::FromArgb(253, 230, 138)
    ErrorText = [System.Drawing.Color]::FromArgb(252, 165, 165)
}

$metrics = @{
    FormWidth = 1180
    FormHeight = 820
    Margin = 24
    PanelWidth = 1120
    ControlHeight = 34
    ActionButtonWidth = 138
    PrimaryButtonWidth = 120
    SecondaryActionWidth = 120
    InputWidth = 922
}

function Set-StatusMessage {
    param(
        [System.Windows.Forms.Label]$Label,
        [string]$Message,
        [System.Drawing.Color]$Color
    )

    $Label.Text = $Message
    $Label.ForeColor = $Color
}

function Clear-PreviewState {
    $state.LastRows = @()
    $grid.Items.Clear()
    $btnApply.Enabled = $false
    $lblSummary.Text = "入力待ちです。"
    Set-StatusMessage -Label $lblStatus -Message "CSVと対象フォルダを指定して、プレビューを押してください。" -Color $theme.Muted
}

function Show-WarningMessage {
    param([string]$Message)

    [System.Windows.Forms.MessageBox]::Show(
        $Message,
        "一括リネーム",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    ) | Out-Null
}

function Show-FolderPickerDialog {
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = "フォルダ|*.folder"
    $dialog.Title = "対象フォルダを選択"
    $dialog.CheckFileExists = $false
    $dialog.CheckPathExists = $true
    $dialog.ValidateNames = $false
    $dialog.FileName = "このフォルダを選択"

    if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        return $null
    }

    $selectedPath = Split-Path -Path $dialog.FileName -Parent
    if ([string]::IsNullOrWhiteSpace($selectedPath)) {
        return $null
    }

    return $selectedPath
}

function Set-PrimaryButtonStyle {
    param([System.Windows.Forms.Button]$Button)

    $Button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $Button.FlatAppearance.BorderSize = 0
    $Button.BackColor = $theme.Accent
    $Button.ForeColor = [System.Drawing.Color]::White
    $Button.Font = New-Object System.Drawing.Font("Yu Gothic UI", 9.5, [System.Drawing.FontStyle]::Bold)
    $Button.Cursor = [System.Windows.Forms.Cursors]::Hand
}

function Set-SecondaryButtonStyle {
    param([System.Windows.Forms.Button]$Button)

    $Button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $Button.FlatAppearance.BorderSize = 0
    $Button.BackColor = [System.Drawing.Color]::FromArgb(39, 49, 66)
    $Button.ForeColor = $theme.Text
    $Button.Font = New-Object System.Drawing.Font("Yu Gothic UI", 9)
    $Button.Cursor = [System.Windows.Forms.Cursors]::Hand
}

function Get-PreviewRowColors {
    param(
        [psobject]$Row,
        [bool]$IsSelected
    )

    if ($IsSelected) {
        return [pscustomobject]@{
            BackColor = [System.Drawing.Color]::FromArgb(30, 64, 175)
            ForeColor = [System.Drawing.Color]::White
        }
    }

    switch ([string]$Row.Status) {
        "実行可能" {
            return [pscustomobject]@{
                BackColor = $theme.AccentDark
                ForeColor = [System.Drawing.Color]::FromArgb(236, 253, 245)
            }
        }
        "要確認" {
            return [pscustomobject]@{
                BackColor = $theme.DangerDark
                ForeColor = $theme.DangerText
            }
        }
        "元ファイルなし" {
            return [pscustomobject]@{
                BackColor = $theme.WarningDark
                ForeColor = $theme.WarningText
            }
        }
        default {
            return [pscustomobject]@{
                BackColor = $theme.InputBg
                ForeColor = [System.Drawing.Color]::FromArgb(203, 213, 225)
            }
        }
    }
}

function Add-TextBoxShell {
    param(
        [System.Windows.Forms.Control]$Parent,
        [System.Windows.Forms.TextBox]$TextBox,
        [int]$X,
        [int]$Y,
        [int]$Width,
        [int]$Height
    )

    $shell = New-Object System.Windows.Forms.Panel
    $shell.Location = New-Object System.Drawing.Point($X, $Y)
    $shell.Size = New-Object System.Drawing.Size($Width, $Height)
    $shell.BackColor = $theme.InputBg
    $shell.BorderStyle = [System.Windows.Forms.BorderStyle]::None

    $TextBox.Multiline = $false
    $TextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $TextBox.BackColor = $theme.InputBg
    $TextBox.ForeColor = $theme.Title
    $TextBox.Location = New-Object System.Drawing.Point(10, [Math]::Max(8, [int](($Height - $TextBox.PreferredHeight) / 2)))
    $TextBox.Size = New-Object System.Drawing.Size(($Width - 20), ($TextBox.PreferredHeight + 2))

    $shell.Controls.Add($TextBox)
    $Parent.Controls.Add($shell)
    return $shell
}

function New-FileDropHandler {
    param(
        [System.Windows.Forms.TextBox]$TextBox,
        [ValidateSet("File", "Folder")]
        [string]$Kind
    )

    $TextBox.AllowDrop = $true

    $TextBox.add_DragEnter({
        if ($_.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
            $_.Effect = [System.Windows.Forms.DragDropEffects]::Copy
        } else {
            $_.Effect = [System.Windows.Forms.DragDropEffects]::None
        }
    })

    $TextBox.add_DragDrop({
        $items = $_.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
        if (-not $items -or $items.Length -eq 0) {
            return
        }

        $path = [string]$items[0]
        if ($Kind -eq "File" -and (Test-Path -LiteralPath $path -PathType Leaf)) {
            $TextBox.Text = $path
        }
        if ($Kind -eq "Folder" -and (Test-Path -LiteralPath $path -PathType Container)) {
            $TextBox.Text = $path
        }
    })
}

function Get-PlanArguments {
    return @{
        CsvPath = $txtCsv.Text
        TargetDirectory = $txtFolder.Text
        TargetColumn = $txtTargetColumn.Text
        NameTemplate = $txtTemplate.Text
    }
}

function Set-BusyState {
    param([bool]$IsBusy)

    $form.UseWaitCursor = $IsBusy
    $btnApply.Enabled = $false
    $btnCsv.Enabled = -not $IsBusy
    $btnFolder.Enabled = -not $IsBusy
}

function Update-PreviewStatus {
    param([object[]]$Rows)

    $lblSummary.Text = Get-RenameSummaryText -Rows $Rows
    $summary = Get-RenameSummary -Rows $Rows

    if ($summary.Ready -gt 0) {
        Set-StatusMessage -Label $lblStatus -Message "プレビュー完了。適用したい行を一覧から選択してください。" -Color $theme.Muted
    } elseif ($summary.Blocked -gt 0) {
        Set-StatusMessage -Label $lblStatus -Message "プレビュー完了。問題のある行があります。内容を確認してください。" -Color $theme.ErrorText
    } else {
        Set-StatusMessage -Label $lblStatus -Message "変更対象はありません。" -Color $theme.Muted
    }
}

function Bind-PreviewRows {
    param([object[]]$Rows)

    $grid.BeginUpdate()
    $grid.Items.Clear()

    foreach ($row in @($Rows)) {
        $item = New-Object System.Windows.Forms.ListViewItem([string]$row.No)
        [void]$item.SubItems.Add([string]$row.SourceName)
        [void]$item.SubItems.Add([string]$row.TargetName)
        [void]$item.SubItems.Add([string]$row.Action)
        [void]$item.SubItems.Add([string]$row.Status)
        [void]$item.SubItems.Add([string]($row.Issues -join " | "))
        $item.Tag = $row
        $colors = Get-PreviewRowColors -Row $row -IsSelected:$false
        $item.BackColor = $colors.BackColor
        $item.ForeColor = $colors.ForeColor

        [void]$grid.Items.Add($item)
    }

    $grid.EndUpdate()

    $state.LastRows = @($Rows)
    Update-PreviewStatus -Rows $Rows
    Update-ApplyButtonState
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "一括リネーム"
$form.Size = New-Object System.Drawing.Size($metrics.FormWidth, $metrics.FormHeight)
$form.StartPosition = "CenterScreen"
$form.MinimumSize = New-Object System.Drawing.Size(1060, 760)
$form.BackColor = $theme.WindowBg

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = "CSVで一括リネーム"
$lblTitle.Location = New-Object System.Drawing.Point($metrics.Margin, 20)
$lblTitle.AutoSize = $true
$lblTitle.Font = New-Object System.Drawing.Font("Yu Gothic UI", 18, [System.Drawing.FontStyle]::Bold)
$lblTitle.ForeColor = $theme.Title
$form.Controls.Add($lblTitle)

$lblLead = New-Object System.Windows.Forms.Label
$lblLead.Text = "CSVと対象フォルダを指定して、一覧で確認してから名前変更できます。プレビュー用ファイルは作りません。"
$lblLead.Location = New-Object System.Drawing.Point(($metrics.Margin + 2), 56)
$lblLead.Size = New-Object System.Drawing.Size(980, 24)
$lblLead.Font = New-Object System.Drawing.Font("Yu Gothic UI", 9.5)
$lblLead.ForeColor = $theme.Muted
$form.Controls.Add($lblLead)

$panelInput = New-Object System.Windows.Forms.Panel
$panelInput.Location = New-Object System.Drawing.Point($metrics.Margin, 96)
$panelInput.Size = New-Object System.Drawing.Size($metrics.PanelWidth, 270)
$panelInput.Anchor = "Top,Left,Right"
$panelInput.BackColor = $theme.PanelBg
$panelInput.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$form.Controls.Add($panelInput)

$txtCsv = New-Object System.Windows.Forms.TextBox
$txtCsv.Font = New-Object System.Drawing.Font("Yu Gothic UI", 10)
$null = Add-TextBoxShell -Parent $panelInput -TextBox $txtCsv -X 20 -Y 46 -Width $metrics.InputWidth -Height $metrics.ControlHeight
New-FileDropHandler -TextBox $txtCsv -Kind File

$txtFolder = New-Object System.Windows.Forms.TextBox
$txtFolder.Font = New-Object System.Drawing.Font("Yu Gothic UI", 10)
$null = Add-TextBoxShell -Parent $panelInput -TextBox $txtFolder -X 20 -Y 120 -Width $metrics.InputWidth -Height $metrics.ControlHeight
New-FileDropHandler -TextBox $txtFolder -Kind Folder

$txtTemplate = New-Object System.Windows.Forms.TextBox
$txtTemplate.Text = "{No:0000}_{Content}{Ext}"
$txtTemplate.Font = New-Object System.Drawing.Font("Consolas", 10)
$null = Add-TextBoxShell -Parent $panelInput -TextBox $txtTemplate -X 20 -Y 194 -Width 470 -Height $metrics.ControlHeight

$txtTargetColumn = New-Object System.Windows.Forms.TextBox
$txtTargetColumn.Font = New-Object System.Drawing.Font("Yu Gothic UI", 10)
$null = Add-TextBoxShell -Parent $panelInput -TextBox $txtTargetColumn -X 560 -Y 194 -Width 220 -Height $metrics.ControlHeight

foreach ($spec in @(
    @{ Text = "CSVファイル"; X = 20; Y = 20 },
    @{ Text = "対象フォルダ"; X = 20; Y = 94 },
    @{ Text = "名前テンプレート"; X = 20; Y = 168 },
    @{ Text = "変更先列名"; X = 560; Y = 168 }
)) {
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $spec.Text
    $label.Location = New-Object System.Drawing.Point($spec.X, $spec.Y)
    $label.AutoSize = $true
    $label.Font = New-Object System.Drawing.Font("Yu Gothic UI", 9.5, [System.Drawing.FontStyle]::Bold)
    $label.ForeColor = $theme.Text
    $panelInput.Controls.Add($label)
}

$btnCsv = New-Object System.Windows.Forms.Button
$btnCsv.Text = "CSVを選ぶ"
$btnCsv.Location = New-Object System.Drawing.Point(954, 46)
$btnCsv.Size = New-Object System.Drawing.Size($metrics.ActionButtonWidth, $metrics.ControlHeight)
Set-SecondaryButtonStyle -Button $btnCsv
$panelInput.Controls.Add($btnCsv)

$btnFolder = New-Object System.Windows.Forms.Button
$btnFolder.Text = "フォルダを選ぶ"
$btnFolder.Location = New-Object System.Drawing.Point(954, 120)
$btnFolder.Size = New-Object System.Drawing.Size($metrics.ActionButtonWidth, $metrics.ControlHeight)
Set-SecondaryButtonStyle -Button $btnFolder
$panelInput.Controls.Add($btnFolder)

$lblOr = New-Object System.Windows.Forms.Label
$lblOr.Text = "または"
$lblOr.Location = New-Object System.Drawing.Point(505, 201)
$lblOr.AutoSize = $true
$lblOr.Font = New-Object System.Drawing.Font("Yu Gothic UI", 8.5)
$lblOr.ForeColor = [System.Drawing.Color]::FromArgb(100, 116, 139)
$panelInput.Controls.Add($lblOr)

$lblHint = New-Object System.Windows.Forms.Label
$lblHint.Text = "テンプレートか変更先列名のどちらかを使います。CSV とフォルダはドラッグ&ドロップにも対応しています。"
$lblHint.Location = New-Object System.Drawing.Point(20, 234)
$lblHint.Size = New-Object System.Drawing.Size(760, 22)
$lblHint.Font = New-Object System.Drawing.Font("Yu Gothic UI", 8.5)
$lblHint.ForeColor = $theme.Muted
$panelInput.Controls.Add($lblHint)

$btnPreview = New-Object System.Windows.Forms.Button
$btnPreview.Text = "プレビュー"
$btnPreview.Location = New-Object System.Drawing.Point(16, 14)
$btnPreview.Size = New-Object System.Drawing.Size($metrics.SecondaryActionWidth, $metrics.ControlHeight)
Set-SecondaryButtonStyle -Button $btnPreview

$btnApply = New-Object System.Windows.Forms.Button
$btnApply.Text = "適用"
$btnApply.Location = New-Object System.Drawing.Point(148, 14)
$btnApply.Size = New-Object System.Drawing.Size($metrics.PrimaryButtonWidth, $metrics.ControlHeight)
$btnApply.Enabled = $false
Set-PrimaryButtonStyle -Button $btnApply

$lblSummary = New-Object System.Windows.Forms.Label
$lblSummary.Text = "入力待ちです。"
$lblSummary.Location = New-Object System.Drawing.Point(292, 12)
$lblSummary.AutoSize = $true
$lblSummary.Font = New-Object System.Drawing.Font("Yu Gothic UI", 11, [System.Drawing.FontStyle]::Bold)
$lblSummary.ForeColor = $theme.Title

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = "CSVと対象フォルダを指定して、プレビューを押してください。"
$lblStatus.Location = New-Object System.Drawing.Point(292, 42)
$lblStatus.Size = New-Object System.Drawing.Size(812, 28)
$lblStatus.Font = New-Object System.Drawing.Font("Yu Gothic UI", 9.5)
$lblStatus.ForeColor = $theme.Muted

$panelGrid = New-Object System.Windows.Forms.Panel
$panelGrid.Location = New-Object System.Drawing.Point($metrics.Margin, 382)
$panelGrid.Size = New-Object System.Drawing.Size($metrics.PanelWidth, 394)
$panelGrid.Anchor = "Top,Bottom,Left,Right"
$panelGrid.BackColor = $theme.PanelBg
$panelGrid.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$form.Controls.Add($panelGrid)
$panelGrid.Controls.Add($btnPreview)
$panelGrid.Controls.Add($btnApply)
$panelGrid.Controls.Add($lblSummary)
$panelGrid.Controls.Add($lblStatus)

$grid = New-Object System.Windows.Forms.ListView
$grid.Location = New-Object System.Drawing.Point(16, 84)
$grid.Size = New-Object System.Drawing.Size(1088, 294)
$grid.Anchor = "Top,Bottom,Left,Right"
$grid.Visible = $true
$grid.View = [System.Windows.Forms.View]::Details
$grid.FullRowSelect = $true
$grid.MultiSelect = $true
$grid.BackColor = $theme.GridBg
$grid.ForeColor = $theme.Text
$grid.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$grid.HideSelection = $false
$grid.GridLines = $false
$grid.HeaderStyle = [System.Windows.Forms.ColumnHeaderStyle]::Nonclickable
$grid.Font = New-Object System.Drawing.Font("Yu Gothic UI", 9.5)
$grid.OwnerDraw = $true
[void]$grid.Columns.Add("番号", 70)
[void]$grid.Columns.Add("変更元ファイル名", 250)
[void]$grid.Columns.Add("変更後ファイル名", 250)
[void]$grid.Columns.Add("操作", 90)
[void]$grid.Columns.Add("状態", 110)
[void]$grid.Columns.Add("問題点", 300)
$panelGrid.Controls.Add($grid)
$grid.BringToFront()

$grid.Add_DrawColumnHeader({
    param($sender, $e)

    $e.Graphics.FillRectangle((New-Object System.Drawing.SolidBrush($theme.HeaderBg)), $e.Bounds)
    $textBounds = New-Object System.Drawing.Rectangle(($e.Bounds.X + 8), ($e.Bounds.Y + 1), ($e.Bounds.Width - 16), $e.Bounds.Height)
    [System.Windows.Forms.TextRenderer]::DrawText(
        $e.Graphics,
        $e.Header.Text,
        $grid.Font,
        $textBounds,
        $theme.Title,
        [System.Windows.Forms.TextFormatFlags]::Left -bor [System.Windows.Forms.TextFormatFlags]::VerticalCenter -bor [System.Windows.Forms.TextFormatFlags]::EndEllipsis
    )
    $e.Graphics.DrawLine((New-Object System.Drawing.Pen($theme.GridLine)), ($e.Bounds.Right - 1), $e.Bounds.Top, ($e.Bounds.Right - 1), $e.Bounds.Bottom)
})

$grid.Add_DrawItem({
    param($sender, $e)

    if ($e.Item.ListView.View -ne [System.Windows.Forms.View]::Details) {
        $e.DrawDefault = $true
    }
})

$grid.Add_DrawSubItem({
    param($sender, $e)

    $row = $e.Item.Tag
    $colors = Get-PreviewRowColors -Row $row -IsSelected:$e.Item.Selected
    $e.Graphics.FillRectangle((New-Object System.Drawing.SolidBrush($colors.BackColor)), $e.Bounds)

    $textBounds = New-Object System.Drawing.Rectangle(($e.Bounds.X + 6), ($e.Bounds.Y + 1), ($e.Bounds.Width - 12), ($e.Bounds.Height - 2))
    [System.Windows.Forms.TextRenderer]::DrawText(
        $e.Graphics,
        $e.SubItem.Text,
        $grid.Font,
        $textBounds,
        $colors.ForeColor,
        [System.Windows.Forms.TextFormatFlags]::Left -bor [System.Windows.Forms.TextFormatFlags]::VerticalCenter -bor [System.Windows.Forms.TextFormatFlags]::EndEllipsis
    )

    $e.Graphics.DrawRectangle((New-Object System.Drawing.Pen($theme.GridLine)), $e.Bounds.X, $e.Bounds.Y, ($e.Bounds.Width - 1), ($e.Bounds.Height - 1))
})

$openCsvDialog = New-Object System.Windows.Forms.OpenFileDialog
$openCsvDialog.Filter = "CSVファイル (*.csv)|*.csv|すべてのファイル (*.*)|*.*"
$openCsvDialog.Title = "CSVを選択"

$state = [pscustomobject]@{
    LastRows = @()
}

function Get-SelectedPreviewRows {
    $selected = @()
    foreach ($item in $grid.SelectedItems) {
        if ($null -ne $item.Tag) {
            $selected += $item.Tag
        }
    }
    return @($selected)
}

function Update-ApplyButtonState {
    $selectedRows = @(Get-SelectedPreviewRows)
    $runnableRows = @($selectedRows | Where-Object { $_.Status -eq "実行可能" -and $_.Action -eq "変更" })
    $nonRunnableRows = @($selectedRows | Where-Object { $_.Status -ne "実行可能" -or $_.Action -ne "変更" })

    $btnApply.Enabled = ($runnableRows.Count -gt 0)

    if (@($state.LastRows).Count -eq 0) {
        $lblSummary.Text = "入力待ちです。"
        Set-StatusMessage -Label $lblStatus -Message "CSVと対象フォルダを指定して、プレビューを押してください。" -Color $theme.Muted
        return
    }

    $baseSummary = Get-RenameSummaryText -Rows $state.LastRows
    if ($selectedRows.Count -eq 0) {
        $lblSummary.Text = "$baseSummary    選択: 0"
        Set-StatusMessage -Label $lblStatus -Message "適用したい行を一覧から選択してください。" -Color $theme.Muted
        return
    }

    $lblSummary.Text = "$baseSummary    選択: $($selectedRows.Count)"
    if ($runnableRows.Count -gt 0 -and $nonRunnableRows.Count -eq 0) {
        Set-StatusMessage -Label $lblStatus -Message "選択中の $($runnableRows.Count) 行を適用できます。" -Color $theme.SuccessText
    } elseif ($runnableRows.Count -gt 0) {
        Set-StatusMessage -Label $lblStatus -Message "選択中に適用できない行があります。実行可能な行だけが対象になります。" -Color $theme.WarningTextStrong
    } else {
        Set-StatusMessage -Label $lblStatus -Message "選択中の行は適用できません。実行可能な行を選択してください。" -Color $theme.ErrorText
    }
}

function Validate-Inputs {
    param([switch]$Silent)

    if ([string]::IsNullOrWhiteSpace($txtCsv.Text) -or -not (Test-Path -LiteralPath $txtCsv.Text -PathType Leaf)) {
        if (-not $Silent) {
            Show-WarningMessage -Message "CSVファイルを指定してください。"
        }
        return $false
    }

    if ([string]::IsNullOrWhiteSpace($txtFolder.Text) -or -not (Test-Path -LiteralPath $txtFolder.Text -PathType Container)) {
        if (-not $Silent) {
            Show-WarningMessage -Message "対象フォルダを指定してください。"
        }
        return $false
    }

    if ([string]::IsNullOrWhiteSpace($txtTargetColumn.Text) -and [string]::IsNullOrWhiteSpace($txtTemplate.Text)) {
        if (-not $Silent) {
            Show-WarningMessage -Message "テンプレートまたは変更先列名を指定してください。"
        }
        return $false
    }

    return $true
}

function Refresh-Preview {
    param([switch]$Silent)

    if (-not (Validate-Inputs -Silent:$Silent)) {
        Clear-PreviewState
        return
    }

    Set-BusyState -IsBusy $true
    [System.Windows.Forms.Application]::DoEvents()

    try {
        Set-StatusMessage -Label $lblStatus -Message "処理中..." -Color $theme.Info

        $planArguments = Get-PlanArguments
        $rows = @(New-RenamePlan @planArguments)

        Bind-PreviewRows -Rows $rows
    } catch {
        Clear-PreviewState
        Set-StatusMessage -Label $lblStatus -Message ([string]$_.Exception.Message) -Color $theme.ErrorText
        if (-not $Silent) {
            [System.Windows.Forms.MessageBox]::Show(
                [string]$_.Exception.Message,
                "一括リネーム",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            ) | Out-Null
        }
    } finally {
        Set-BusyState -IsBusy $false
    }
}

function Queue-Preview {
    $state.LastRows = @()
    $grid.Items.Clear()
    $btnApply.Enabled = $false
    $lblSummary.Text = "未プレビューです。"
    Set-StatusMessage -Label $lblStatus -Message "入力内容が変わりました。プレビューを押して更新してください。" -Color $theme.Muted
}

function Apply-RenamePlan {
    $selectedRows = @(Get-SelectedPreviewRows | Where-Object { $_.Status -eq "実行可能" -and $_.Action -eq "変更" })
    foreach ($row in $selectedRows) {
        Rename-Item -LiteralPath $row.SourcePath -NewName $row.TargetName
    }
}

$btnCsv.Add_Click({
    if ($openCsvDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtCsv.Text = $openCsvDialog.FileName
    }
})

$btnFolder.Add_Click({
    $selectedFolder = Show-FolderPickerDialog
    if (-not [string]::IsNullOrWhiteSpace($selectedFolder)) {
        $txtFolder.Text = $selectedFolder
    }
})

$txtCsv.Add_TextChanged({ Queue-Preview })
$txtFolder.Add_TextChanged({ Queue-Preview })
$txtTemplate.Add_TextChanged({ Queue-Preview })
$txtTargetColumn.Add_TextChanged({ Queue-Preview })
$grid.Add_SelectedIndexChanged({ Update-ApplyButtonState })

$btnPreview.Add_Click({
    Refresh-Preview
})

$btnApply.Add_Click({
    $selectedRows = @(Get-SelectedPreviewRows | Where-Object { $_.Status -eq "実行可能" -and $_.Action -eq "変更" })
    if ($selectedRows.Count -eq 0) {
        Show-WarningMessage -Message "適用したい実行可能な行を選択してください。"
        return
    }

    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "選択中の $($selectedRows.Count) 行の名前を変更します。続けますか？",
        "一括リネーム",
        [System.Windows.Forms.MessageBoxButtons]::OKCancel,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )

    if ($confirm -ne [System.Windows.Forms.DialogResult]::OK) {
        return
    }

    try {
        $btnApply.Enabled = $false
        Apply-RenamePlan
        Refresh-Preview -Silent
        Set-StatusMessage -Label $lblStatus -Message "選択中の行に対してリネームを実行しました。プレビュー用のファイルは作っていません。" -Color $theme.SuccessText
    } catch {
        Set-StatusMessage -Label $lblStatus -Message ([string]$_.Exception.Message) -Color $theme.ErrorText
        [System.Windows.Forms.MessageBox]::Show(
            [string]$_.Exception.Message,
            "一括リネーム",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }
})

[void]$form.ShowDialog()
