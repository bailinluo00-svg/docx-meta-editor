# DocxMetaGui.ps1 -- simple WinForms front-end for DocxMeta.ps1
# Launch: run.bat   (or: powershell -ExecutionPolicy Bypass -File DocxMetaGui.ps1)
#
# For headless verification only:  -Probe
#   builds the whole form and wires every event handler, but does NOT open a
#   window, letting a test harness drive the real button handlers in-process.
#
# Two PowerShell traps this file is written around, both covered by the tests:
#   1. A .ps1 containing non-ASCII text is read as GBK by Windows PowerShell 5.1
#      unless it carries a UTF-8 BOM (re-added below if it ever goes missing).
#   2. Call a PowerShell function as  Fn 'a' 'b'  -- never  Fn('a', 'b').
#      The parenthesised form is parsed as a .NET method call and fails with
#      "Cannot convert value to type System.String" on typed parameters.

[CmdletBinding()]
param([switch] $Probe)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms | Out-Null
Add-Type -AssemblyName System.Drawing | Out-Null

# Derive our own directory from myCommand: that works when dot-sourced too,
# unlike $PSScriptRoot which belongs to the *calling* script.
$selfFile = $MyInvocation.MyCommand.Path
$selfDir  = Split-Path -Parent $selfFile
$rawBytes = [System.IO.File]::ReadAllBytes($selfFile)
if (-not ($rawBytes.Length -ge 3 -and $rawBytes[0] -eq 0xEF -and $rawBytes[1] -eq 0xBB -and $rawBytes[2] -eq 0xBF)) {
    $text = [System.Text.Encoding]::UTF8.GetString($rawBytes)
    [System.IO.File]::WriteAllText($selfFile, $text, (New-Object System.Text.UTF8Encoding($true)))
}

. (Join-Path $selfDir 'DocxMeta.ps1')

[System.Windows.Forms.Application]::EnableVisualStyles()

# ---------------------------------------------------------------- helpers
function ConvertFrom-Minutes {
    param([Parameter(Mandatory)][int] $Minutes)
    if ($Minutes -lt 0) { return '-' }
    return ("{0} 小时 {1} 分" -f [math]::Floor($Minutes / 60), ($Minutes % 60))
}

function Set-Status {
    param([string] $Message, [string] $Color = 'Black')
    $lblStatus.Text = $Message
    $lblStatus.ForeColor = [System.Drawing.Color]::FromName($Color)
    $lblStatus.Refresh()
}

# All user-facing dialogs funnel through these wrappers: it keeps the event
# handlers readable, gives the verification harness one place to intercept them,
# and leaves a single place to localise later.
function Show-UiMessage {
    param(
        [Parameter(Mandatory)][string] $MessageText,
        [string] $Caption = '提示',
        [string] $Icon = 'Information'
    )
    # The hooks let the verification harness intercept dialogs. Event handlers
    # resolve functions from the scope they were DEFINED in, so a test cannot
    # shadow these wrappers after the fact -- it has to go through the hooks.
    if ($script:UiMessageHook) { return & $script:UiMessageHook $MessageText $Caption $Icon }
    return [System.Windows.Forms.MessageBox]::Show($MessageText, $Caption, 'OK', $Icon)
}

function Show-UiConfirm {
    param(
        [Parameter(Mandatory)][string] $MessageText,
        [string] $Caption = '确认'
    )
    if ($script:UiConfirmHook) { return & $script:UiConfirmHook $MessageText $Caption }
    return [System.Windows.Forms.MessageBox]::Show($MessageText, $Caption, 'YesNo', 'Question')
}

# ---------------------------------------------------------------- form
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Word 文档信息修改器  v1.0'
$form.ClientSize = New-Object System.Drawing.Size(716, 636)
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false
$form.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 9.5)

$fileFilter = 'Word 文档 (*.docx)|*.docx|所有文件 (*.*)|*.*'

function Pick-DocxFile {
    <#
      Returns a chosen .docx path, or $null if the user cancelled.
      $global:DocxMetaPickFile is a test-only hook: the verification harness sets
      it so the real button handlers can be driven without a modal dialog. It is
      $null in normal use, so end users always get the real picker.
    #>
    param([Parameter(Mandatory)][string] $Title)
    if ($null -ne $global:DocxMetaPickFile) { return $global:DocxMetaPickFile }
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = $fileFilter
    $dlg.Title = $Title
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { return $dlg.FileName }
    return $null
}

# ============================ 1. 目标文档 ============================
$grpFile = New-Object System.Windows.Forms.GroupBox
$grpFile.Text = '1. 选择要修改的 Word 文档'
$grpFile.Location = New-Object System.Drawing.Point(12, 8)
$grpFile.Size = New-Object System.Drawing.Size(692, 96)

$btnPick = New-Object System.Windows.Forms.Button
$btnPick.Text = '选择文档...'
$btnPick.Location = New-Object System.Drawing.Point(14, 26)
$btnPick.Size = New-Object System.Drawing.Size(104, 28)

$txtFile = New-Object System.Windows.Forms.TextBox
$txtFile.Location = New-Object System.Drawing.Point(126, 27)
$txtFile.Size = New-Object System.Drawing.Size(548, 25)
$txtFile.ReadOnly = $true
$txtFile.BackColor = [System.Drawing.Color]::White

$lblCurrent = New-Object System.Windows.Forms.Label
$lblCurrent.Location = New-Object System.Drawing.Point(16, 62)
$lblCurrent.Size = New-Object System.Drawing.Size(660, 22)
$lblCurrent.Text = '当前文档：尚未选择'
$lblCurrent.ForeColor = [System.Drawing.Color]::DimGray

$grpFile.Controls.AddRange(@($btnPick, $txtFile, $lblCurrent))

# ============================ 2. 编辑总时间 ============================
$grpTime = New-Object System.Windows.Forms.GroupBox
$grpTime.Text = '2. 编辑总时间（不勾选则保持原值不变）'
$grpTime.Location = New-Object System.Drawing.Point(12, 112)
$grpTime.Size = New-Object System.Drawing.Size(692, 116)

$chkTime = New-Object System.Windows.Forms.CheckBox
$chkTime.Text = '修改编辑总时间'
$chkTime.Location = New-Object System.Drawing.Point(14, 22)
$chkTime.Size = New-Object System.Drawing.Size(150, 24)

$tblTime = New-Object System.Windows.Forms.TableLayoutPanel
$tblTime.Location = New-Object System.Drawing.Point(14, 48)
$tblTime.Size = New-Object System.Drawing.Size(560, 58)
$tblTime.ColumnCount = 4
$tblTime.RowCount = 1
$null = $tblTime.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Absolute', 70)))
$null = $tblTime.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Absolute', 70)))
$null = $tblTime.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Absolute', 60)))
$null = $tblTime.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Absolute', 340)))

$numHours = New-Object System.Windows.Forms.NumericUpDown
$numHours.Minimum = 0
$numHours.Maximum = 9999
$numHours.Value = 6
$numHours.Width = 60

$numMins = New-Object System.Windows.Forms.NumericUpDown
$numMins.Minimum = 0
$numMins.Maximum = 59
$numMins.Value = 0
$numMins.Width = 60

$lblH = New-Object System.Windows.Forms.Label
$lblH.Text = '小时'
$lblH.TextAlign = 'MiddleLeft'
$lblH.Width = 56

$lblHint = New-Object System.Windows.Forms.Label
$lblHint.Text = '（Word 只精确到分钟；6 小时 = 360 分钟）'
$lblHint.TextAlign = 'MiddleLeft'
$lblHint.ForeColor = [System.Drawing.Color]::DimGray
$lblHint.Width = 340

$null = $tblTime.Controls.Add($numHours, 0, 0)
$null = $tblTime.Controls.Add($lblH, 1, 0)
$null = $tblTime.Controls.Add($numMins, 2, 0)
$null = $tblTime.Controls.Add($lblHint, 3, 0)

$grpTime.Controls.AddRange(@($chkTime, $tblTime))

# ============================ 3. 作者名称 ============================
$grpAuthor = New-Object System.Windows.Forms.GroupBox
$grpAuthor.Text = '3. 作者名称（不勾选则保持原值不变）'
$grpAuthor.Location = New-Object System.Drawing.Point(12, 234)
$grpAuthor.Size = New-Object System.Drawing.Size(692, 158)

$chkAuthor = New-Object System.Windows.Forms.CheckBox
$chkAuthor.Text = '修改作者名称'
$chkAuthor.Location = New-Object System.Drawing.Point(14, 22)
$chkAuthor.Size = New-Object System.Drawing.Size(150, 24)

$lblAuthorCap = New-Object System.Windows.Forms.Label
$lblAuthorCap.Text = '作者：'
$lblAuthorCap.Location = New-Object System.Drawing.Point(14, 52)
$lblAuthorCap.Size = New-Object System.Drawing.Size(52, 24)
$lblAuthorCap.TextAlign = 'MiddleLeft'

$txtAuthor = New-Object System.Windows.Forms.TextBox
$txtAuthor.Location = New-Object System.Drawing.Point(66, 53)
$txtAuthor.Size = New-Object System.Drawing.Size(300, 25)

$chkModBy = New-Object System.Windows.Forms.CheckBox
$chkModBy.Text = '同时修改「上次保存者」'
$chkModBy.Location = New-Object System.Drawing.Point(380, 54)
$chkModBy.Size = New-Object System.Drawing.Size(200, 24)
$chkModBy.Checked = $true

$lblDonor = New-Object System.Windows.Forms.Label
$lblDonor.Text = '从另一个 Word 文档读取作者：'
$lblDonor.Location = New-Object System.Drawing.Point(14, 96)
$lblDonor.Size = New-Object System.Drawing.Size(210, 24)
$lblDonor.TextAlign = 'MiddleLeft'

$txtDonor = New-Object System.Windows.Forms.TextBox
$txtDonor.Location = New-Object System.Drawing.Point(14, 124)
$txtDonor.Size = New-Object System.Drawing.Size(470, 25)
$txtDonor.ReadOnly = $true
$txtDonor.BackColor = [System.Drawing.Color]::White

$btnDonor = New-Object System.Windows.Forms.Button
$btnDonor.Text = '读取作者并填入...'
$btnDonor.Location = New-Object System.Drawing.Point(492, 122)
$btnDonor.Size = New-Object System.Drawing.Size(182, 28)

$grpAuthor.Controls.AddRange(@($chkAuthor, $lblAuthorCap, $txtAuthor, $chkModBy, $lblDonor, $txtDonor, $btnDonor))

# ============================ 4. 保存方式 ============================
$grpSave = New-Object System.Windows.Forms.GroupBox
$grpSave.Text = '4. 保存方式'
$grpSave.Location = New-Object System.Drawing.Point(12, 398)
$grpSave.Size = New-Object System.Drawing.Size(692, 76)

$rbInPlace = New-Object System.Windows.Forms.RadioButton
$rbInPlace.Text = '直接修改原文件（自动生成 .backup.docx 备份）'
$rbInPlace.Location = New-Object System.Drawing.Point(14, 22)
$rbInPlace.Size = New-Object System.Drawing.Size(340, 24)
$rbInPlace.Checked = $true

$rbSaveAs = New-Object System.Windows.Forms.RadioButton
$rbSaveAs.Text = '另存为新文件'
$rbSaveAs.Location = New-Object System.Drawing.Point(14, 46)
$rbSaveAs.Size = New-Object System.Drawing.Size(120, 24)

$lblSaveAs = New-Object System.Windows.Forms.Label
$lblSaveAs.Text = '新文件名（留空则自动加 _modified）'
$lblSaveAs.Location = New-Object System.Drawing.Point(138, 47)
$lblSaveAs.Size = New-Object System.Drawing.Size(250, 22)
$lblSaveAs.ForeColor = [System.Drawing.Color]::DimGray

$txtSaveAs = New-Object System.Windows.Forms.TextBox
$txtSaveAs.Location = New-Object System.Drawing.Point(392, 45)
$txtSaveAs.Size = New-Object System.Drawing.Size(284, 25)
$txtSaveAs.Enabled = $false

$grpSave.Controls.AddRange(@($rbInPlace, $rbSaveAs, $lblSaveAs, $txtSaveAs))

# ============================ 5. 执行 ============================
$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = '开始修改'
$btnRun.Location = New-Object System.Drawing.Point(12, 482)
$btnRun.Size = New-Object System.Drawing.Size(692, 38)
$btnRun.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 11, [System.Drawing.FontStyle]::Bold)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Location = New-Object System.Drawing.Point(14, 526)
$lblStatus.Size = New-Object System.Drawing.Size(690, 22)
$lblStatus.Text = '就绪。'

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(14, 552)
$txtLog.Size = New-Object System.Drawing.Size(690, 76)
$txtLog.Multiline = $true
$txtLog.ReadOnly = $true
$txtLog.ScrollBars = 'Vertical'
$txtLog.BackColor = [System.Drawing.Color]::FromArgb(250, 250, 250)
$txtLog.Font = New-Object System.Drawing.Font('Consolas', 9)

$form.Controls.AddRange(@($grpFile, $grpTime, $grpAuthor, $grpSave, $btnRun, $lblStatus, $txtLog))

function Write-Log {
    param([string] $Message)
    $txtLog.AppendText(("[{0}] {1}{2}" -f (Get-Date -Format 'HH:mm:ss'), $Message, [Environment]::NewLine))
    $txtLog.SelectionStart = $txtLog.TextLength
    $txtLog.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

# ---------------------------------------------------------------- state
$script:currentFile = $null
$script:donorAuthor = $null
$script:donorModBy  = $null

function Update-CurrentInfo {
    if (-not $script:currentFile) {
        $lblCurrent.Text = '当前文档：尚未选择'
        $lblCurrent.ForeColor = [System.Drawing.Color]::DimGray
        return
    }
    try {
        $meta = Get-DocxMeta -Path $script:currentFile
        $authorTxt = if ([string]::IsNullOrEmpty($meta.Author)) { '(空)' } else { $meta.Author }
        $timeTxt = if ($meta.HasTotalTime) { ConvertFrom-Minutes -Minutes $meta.TotalMinutes } else { '(无此字段)' }
        $lblCurrent.Text = ("当前文档：作者 = {0}    |    编辑总时间 = {1}" -f $authorTxt, $timeTxt)
        $lblCurrent.ForeColor = [System.Drawing.Color]::FromArgb(0, 100, 0)
    } catch {
        $lblCurrent.Text = "当前文档：读取失败 - $($_.Exception.Message)"
        $lblCurrent.ForeColor = [System.Drawing.Color]::Firebrick
    }
}

function Reset-ForNewFile {
    $script:currentFile = $null
    $script:donorAuthor = $null
    $script:donorModBy  = $null
    $txtFile.Text = ''
    $txtDonor.Text = ''
    $txtAuthor.Text = ''
    $chkAuthor.Checked = $false
    $chkTime.Checked = $false
    $txtLog.Clear()
    $lblStatus.Text = '就绪。'
    $lblStatus.ForeColor = [System.Drawing.Color]::Black
    Update-CurrentInfo
}

# ---------------------------------------------------------------- events
$btnPick.Add_Click({
    $picked = Pick-DocxFile -Title '选择要修改的 Word 文档'
    if (-not $picked) { return }

    Reset-ForNewFile
    $script:currentFile = $picked
    $txtFile.Text = $picked
    Update-CurrentInfo
    Set-Status -Message '已选择文档，请设置要修改的项目。'
    Write-Log -Message "已载入：$picked"

    try {
        $m = Get-DocxMeta -Path $picked
        $timeTxt = if ($m.HasTotalTime) { ConvertFrom-Minutes -Minutes $m.TotalMinutes } else { '无' }
        Write-Log -Message ("  原作者 = '{0}'，上次保存者 = '{1}'，编辑总时间 = {2}" -f $m.Author, $m.LastModifiedBy, $timeTxt)

        # prefill the controls with the current values for convenience
        if (-not [string]::IsNullOrEmpty($m.Author)) { $txtAuthor.Text = $m.Author }
        if ($m.HasTotalTime) {
            $numHours.Value = [math]::Min([int]$numHours.Maximum, [int][math]::Floor($m.TotalMinutes / 60))
            $numMins.Value = [int]($m.TotalMinutes % 60)
        }
    } catch {
        Write-Log -Message "  读取信息失败：$($_.Exception.Message)"
    }
})

$btnDonor.Add_Click({
    if (-not $script:currentFile) {
        Show-UiMessage '请先选择要修改的目标文档。' '提示'
        return
    }
    $donorPath = Pick-DocxFile -Title '选择提供作者信息的 Word 文档'
    if (-not $donorPath) { return }

    try {
        $donor = Get-DocxMeta -Path $donorPath
    } catch {
        Show-UiMessage "读取失败：$($_.Exception.Message)" '错误' 'Error'
        return
    }

    if ([string]::IsNullOrEmpty($donor.Author)) {
        Show-UiMessage "该文档的作者字段为空，无法复制。`n`n可以改为直接在上面的「作者」框里手动输入。" '没有作者信息' 'Warning'
        Write-Log -Message "来源文档作者为空：$donorPath"
        return
    }

    $script:donorAuthor = $donor.Author
    $script:donorModBy  = $donor.LastModifiedBy
    $txtDonor.Text = $donorPath
    $txtAuthor.Text = $donor.Author
    $chkAuthor.Checked = $true
    $chkModBy.Checked = $true
    Write-Log -Message "已从来源文档读取作者：'$($donor.Author)'"
    Set-Status -Message "作者已填入：$($donor.Author)，点「开始修改」应用。" -Color 'DarkGreen'
})

$rbInPlace.Add_CheckedChanged({
    $txtSaveAs.Enabled = $rbSaveAs.Checked
    $lblSaveAs.Enabled = $rbSaveAs.Checked
})

$btnRun.Add_Click({
    if (-not $script:currentFile) {
        Show-UiMessage '请先选择要修改的 Word 文档。' '提示'
        return
    }

    $doTime   = $chkTime.Checked
    $doAuthor = $chkAuthor.Checked

    if (-not ($doTime -or $doAuthor)) {
        Show-UiMessage '请至少勾选一项要修改的内容（编辑总时间 / 作者名称）。' '提示'
        return
    }

    $author = $txtAuthor.Text
    if ($doAuthor -and [string]::IsNullOrWhiteSpace($author)) {
        Show-UiMessage '作者名称不能是空白。如果想把作者清空，请告诉我，我再加这个功能。' '提示' 'Warning'
        return
    }

    $minutes = ([int]$numHours.Value * 60) + [int]$numMins.Value

    $params = @{ Path = $script:currentFile }
    if ($doTime)   { $params['TotalMinutes'] = $minutes }
    if ($doAuthor) {
        $params['Author'] = $author
        if ($chkModBy.Checked) { $params['ModifiedBy'] = $author }
    }

    if ($rbInPlace.Checked) {
        $params['InPlace'] = $true
    } else {
        $stem = [System.IO.Path]::GetFileNameWithoutExtension($script:currentFile)
        $dir  = [System.IO.Path]::GetDirectoryName($script:currentFile)
        $name = $txtSaveAs.Text.Trim()
        if (-not $name) { $name = $stem + '_modified' }
        if ($name -notmatch '\.docx$') { $name = $name + '.docx' }
        $params['OutputPath'] = Join-Path $dir $name
    }

    # confirm first
    $summary = @()
    if ($doTime)   { $summary += ("编辑总时间 -> {0}（{1} 分钟）" -f (ConvertFrom-Minutes -Minutes $minutes), $minutes) }
    if ($doAuthor) { $summary += ("作者名称 -> {0}" -f $author) }
    $where = if ($rbInPlace.Checked) { '直接修改原文件（自动备份）' } else { "另存为：$($params['OutputPath'])" }
    $msg = "文档：$($script:currentFile)`n`n" + ($summary -join "`n") + "`n`n保存方式：$where`n`n确认执行吗？"
    if ((Show-UiConfirm -MessageText $msg) -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $btnRun.Enabled = $false
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    Set-Status -Message '正在处理...' -Color 'DarkOrange'
    [System.Windows.Forms.Application]::DoEvents()

    try {
        $result = Set-DocxMeta @params
        Write-Log -Message "修改成功：$($result.Changed)"
        Write-Log -Message "输出文件：$($result.Path)"
        if ($result.InPlace) { Write-Log -Message '备份文件：同目录下 <文件名>.backup.docx' }

        $verify = Get-DocxMeta -Path $result.Path
        $vTime = if ($verify.HasTotalTime) { ConvertFrom-Minutes -Minutes $verify.TotalMinutes } else { '无' }
        Write-Log -Message ("回读验证 -> 作者 = '{0}'，上次保存者 = '{1}'，编辑总时间 = {2}" -f $verify.Author, $verify.LastModifiedBy, $vTime)

        $ok = $true
        if ($doTime -and $verify.TotalMinutes -ne $minutes) { $ok = $false }
        if ($doAuthor -and $verify.Author -ne $author) { $ok = $false }

        if ($ok) {
            Set-Status -Message '修改成功，并已回读验证通过。' -Color 'DarkGreen'
            Write-Log -Message '验证通过。'
            # Measured behaviour (Word 16.0, Microsoft 365): TotalTime ACCUMULATES.
            # Word keeps the stored value as a baseline and adds each new session
            # onto it, so this write is a baseline, not a final value.
            $advice = '编辑总时间已设定。' + "`n`n" +
                'Word 的编辑总时间是【累计】的：它会以这个值为基线，以后每次编辑并保存都会在这个基础上继续往上加，不会把它抹掉重算。' + "`n`n" +
                '所以：如果你需要一个精确的最终数字，请在改完全部内容后最后再设一次，之后别再让 Word 保存；如果不需要精确值，正常编辑即可。'
            if ($doAuthor) {
                $advice += "`n`n" + '注意：「作者 / 上次保存者」不一样 —— Word 保存时一般会用当前 Word 用户名把作者冲掉（这一条我没实测，是推断）。要保留作者，也请放在最后一步改。'
            }
            Show-UiMessage ("处理完成，已验证。" + "`n`n" + $advice) '完成'
        } else {
            Set-Status -Message '已写入，但回读验证不一致，请检查。' -Color 'Firebrick'
            Write-Log -Message '警告：回读验证不一致。'
        }

        Update-CurrentInfo
    } catch {
        Set-Status -Message "失败：$($_.Exception.Message)" -Color 'Firebrick'
        Write-Log -Message "失败：$($_.Exception.Message)"
        Show-UiMessage "处理失败：`n`n$($_.Exception.Message)`n`n（提示：如果文件正被 Word 打开，请先关闭它。）" '错误' 'Error'
    } finally {
        $btnRun.Enabled = $true
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

# ---------------------------------------------------------------- show
Reset-ForNewFile
Write-Log -Message '工具已启动。请先「选择文档...」。'

if ($Probe) {
    # Headless: hand the form and every control to the caller and return, so the
    # harness can drive the real button handlers without a window.
    $script:gui = [pscustomobject]@{
        Form       = $form
        BtnPick    = $btnPick
        BtnDonor   = $btnDonor
        BtnRun     = $btnRun
        TxtFile    = $txtFile
        TxtAuthor  = $txtAuthor
        TxtDonor   = $txtDonor
        TxtSaveAs  = $txtSaveAs
        ChkTime    = $chkTime
        ChkAuthor  = $chkAuthor
        ChkModBy   = $chkModBy
        NumHours   = $numHours
        NumMins    = $numMins
        RbInPlace  = $rbInPlace
        RbSaveAs   = $rbSaveAs
        LblStatus  = $lblStatus
        LblCurrent = $lblCurrent
        Log        = $txtLog
        StartupLog = $txtLog.Text
    }
    $script:guiReady = $true
    return
}

$null = $form.ShowDialog()
