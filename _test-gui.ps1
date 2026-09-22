# _test-gui.ps1 -- drives the REAL GUI event handlers in-process (-Probe mode).
#
# Nothing may block on a dialog, so the GUI exposes two explicit hooks
# ($script:UiMessageHook / $script:UiConfirmHook) which this harness fills in.
# Re-defining the wrapper functions would NOT work: event handlers resolve
# functions from the scope where they were defined.
#
# ASCII-only on purpose.

$ErrorActionPreference = 'Stop'
$dir = $PSScriptRoot

$script:failed = 0
function Assert($cond, $msg) {
    if ($cond) { Write-Output "PASS: $msg" } else { Write-Output "FAIL: $msg"; $script:failed++ }
}

Add-Type -AssemblyName System.Windows.Forms

$script:msgLog = New-Object System.Collections.Generic.List[string]
$script:answer  = [System.Windows.Forms.DialogResult]::OK
function Msg-Text { ($script:msgLog -join ' || ') }
function Set-MsgAnswer($r) { $script:answer = $r }

function Invoke-Click {
    <#
      PerformClick() refuses to fire when the control tree was never shown
      (CanSelect is false), so reach the real Click event the way a mouse does:
      through Control.OnClick, which raises every subscribed handler.
    #>
    param([Parameter(Mandatory)][System.Windows.Forms.Control] $Control)
    $flags = [System.Reflection.BindingFlags]::Instance -bor [System.Reflection.BindingFlags]::NonPublic -bor [System.Reflection.BindingFlags]::InvokeMethod
    $null = $Control.GetType().InvokeMember('OnClick', $flags, $null, $Control, @([System.EventArgs]::Empty))
}

# ---------- load the GUI in probe mode (no window) ----------
$global:DocxMetaPickFile = $null
. (Join-Path $dir 'DocxMetaGui.ps1') -Probe

# ---------- fill the dialog hooks the GUI offers ----------
$script:UiMessageHook = {
    param($Text, $Caption, $Icon)
    $script:msgLog.Add("MSG[$Caption] $Text")
    return [System.Windows.Forms.DialogResult]::OK
}
$script:UiConfirmHook = {
    param($Text, $Caption)
    $script:msgLog.Add("CONFIRM[$Caption] $Text")
    return $script:answer
}

Write-Output '--- 0) form construction and initial state ---'
Assert ($script:guiReady -eq $true) 'GUI built in probe mode without opening a window'
Assert ($null -ne $script:gui) 'form controls exposed to the harness'
Assert ($script:gui.StartupLog -match '工具已启动') 'startup log written with correct Chinese'
Assert ($script:gui.Form.Text -match 'Word 文档信息修改器') 'form title is correct Chinese'
Assert ($script:gui.Form.ClientSize.Width -eq 716) 'expected client width'
Assert ($script:gui.ChkTime.Checked -eq $false) 'time checkbox starts unchecked'
Assert ($script:gui.ChkAuthor.Checked -eq $false) 'author checkbox starts unchecked'
Assert ($script:gui.RbInPlace.Checked -eq $true) 'in-place is the default save mode'
Assert ($script:gui.TxtSaveAs.Enabled -eq $false) 'save-as box disabled while in-place selected'
Assert ($script:gui.LblCurrent.Text -match '尚未选择') 'current-document line starts empty'

$script:gui.NumHours.Value = 5
$script:gui.NumMins.Value = 30
Assert (($script:gui.NumHours.Value * 60 + $script:gui.NumMins.Value) -eq 330) 'hours+minutes combine to 330 min'
$script:gui.RbSaveAs.Checked = $true
Assert ($script:gui.TxtSaveAs.Enabled -eq $true) 'save-as box enables when save-as mode is picked'
$script:gui.RbInPlace.Checked = $true

# ---------- fixtures ----------
$tmp = Join-Path $dir '_gui-test'
if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Recurse -Force }
$null = New-Item -ItemType Directory -Path $tmp -Force

. (Join-Path $dir 'Make-TestDocx.ps1')
$sample = Join-Path $dir '_fixtures\sample.docx'
$null = New-TestDocx -Path $sample -Text 'fixture 测试文档'
$baseMeta = Get-DocxMeta -Path $sample
$target = Join-Path $tmp 'target.docx'
$donor  = Join-Path $tmp 'donor.docx'
Copy-Item -LiteralPath $sample -Destination $target
Copy-Item -LiteralPath $sample -Destination $donor

Set-DocxMeta -Path $donor -InPlace -NoBackup -Author 'ZHANG SAN' -ModifiedBy 'ZHANG SAN' | Out-Null
Assert ((Get-DocxMeta -Path $donor).Author -eq 'ZHANG SAN') 'donor fixture author written'

Write-Output ''
Write-Output '--- 1) pick button ---'
$global:DocxMetaPickFile = $target
Invoke-Click -Control $script:gui.BtnPick
Assert ($script:gui.TxtFile.Text -eq $target) 'pick handler stored the chosen path'
Assert ($script:gui.LblCurrent.Text -match '当前文档') 'current-document line refreshed'
Assert ($script:gui.Log.Text -match '已载入') 'log recorded the load'
Write-Output "      $($script:gui.LblCurrent.Text)"

Write-Output ''
Write-Output '--- 2) donor button: copy author from another document ---'
$global:DocxMetaPickFile = $donor
Invoke-Click -Control $script:gui.BtnDonor
Assert ($script:gui.TxtAuthor.Text -eq 'ZHANG SAN') "author box filled from donor (got '$($script:gui.TxtAuthor.Text)')"
Assert ($script:gui.TxtDonor.Text -eq $donor) 'donor path shown in the UI'
Assert ($script:gui.ChkAuthor.Checked -eq $true) 'author checkbox auto-ticked by donor import'
Assert ($script:gui.ChkModBy.Checked -eq $true) 'last-saved-by checkbox ticked by donor import'

Write-Output ''
Write-Output '--- 3) run: donor author + 2h30m, save as result.docx ---'
$script:gui.ChkTime.Checked = $true
$script:gui.NumHours.Value = 2
$script:gui.NumMins.Value = 30
$script:gui.RbSaveAs.Checked = $true
$script:gui.TxtSaveAs.Text = 'result'
$script:msgLog.Clear()
Set-MsgAnswer ([System.Windows.Forms.DialogResult]::Yes)
Invoke-Click -Control $script:gui.BtnRun

$outFile = Join-Path $tmp 'result.docx'
Assert (Test-Path -LiteralPath $outFile) 'save-as produced result.docx'
$vm = Get-DocxMeta -Path $outFile
Assert ($vm.Author -eq 'ZHANG SAN') "result author == donor author (got '$($vm.Author)')"
Assert ($vm.LastModifiedBy -eq 'ZHANG SAN') 'result lastModifiedBy == donor lastModifiedBy'
Assert ($vm.TotalMinutes -eq 150) "2h30m applied as 150 minutes (got $($vm.TotalMinutes))"
Assert ($vm.Path -eq $outFile) 'wrote to the requested path'
Assert ((Get-DocxMeta -Path $target).TotalMinutes -eq $baseMeta.TotalMinutes) 'target untouched in save-as mode'
Assert (Msg-Text -match 'CONFIRM') 'the confirm hook was the one invoked'
Assert (Msg-Text -match '确认执行') 'confirmation dialog shown before writing'
Assert (Msg-Text -match '处理完成') 'completion dialog shown after writing'
Assert (Msg-Text -match '累计') "completion dialog explains Word ACCUMULATES (got: $(Msg-Text))"
Assert ($script:gui.LblStatus.Text -match '成功') "status reports success (got '$($script:gui.LblStatus.Text)')"
Assert ($script:gui.Log.Text -match '回读验证') 'log contains the read-back verification line'
Write-Output "      $($script:gui.LblStatus.Text)"

Write-Output ''
Write-Output '--- 4) time only, in place, with backup ---'
$target2 = Join-Path $tmp 'target2.docx'
Copy-Item -LiteralPath $sample -Destination $target2
$global:DocxMetaPickFile = $target2
Invoke-Click -Control $script:gui.BtnPick
Assert ($script:gui.TxtFile.Text -eq $target2) 'picked the second target'

$script:gui.RbInPlace.Checked = $true
$script:gui.ChkAuthor.Checked = $false
$script:gui.TxtAuthor.Text = ''
$script:gui.ChkTime.Checked = $true
$script:gui.NumHours.Value = 6
$script:gui.NumMins.Value = 0
$script:gui.ChkModBy.Checked = $false
$authorBefore = (Get-DocxMeta -Path $target2).Author
$script:msgLog.Clear()
Invoke-Click -Control $script:gui.BtnRun

$vm2 = Get-DocxMeta -Path $target2
Assert ($vm2.TotalMinutes -eq 360) "in-place 6h applied (got $($vm2.TotalMinutes))"
Assert ($vm2.Author -eq $authorBefore) 'author left alone when its checkbox is off'
$bak2 = Join-Path $tmp 'target2.backup.docx'
Assert (Test-Path -LiteralPath $bak2) 'backup created next to the in-place target'
Assert ((Get-DocxMeta -Path $bak2).TotalMinutes -eq $baseMeta.TotalMinutes) 'backup holds the original time'
Assert ((Get-DocxMeta -Path $bak2).Author -eq $authorBefore) 'backup holds the original author'
Assert ($script:gui.LblStatus.Text -match '成功') "in-place run reported success (got '$($script:gui.LblStatus.Text)')"

Write-Output ''
Write-Output '--- 5) guard rails ---'
$script:gui.ChkTime.Checked = $false
$script:gui.ChkAuthor.Checked = $false
$script:msgLog.Clear()
Invoke-Click -Control $script:gui.BtnRun
Assert (Msg-Text -match '至少勾选一项') "refuses when nothing is ticked (got: $(Msg-Text))"
Assert ((Get-DocxMeta -Path $target2).TotalMinutes -eq 360) 'refusal wrote nothing'

$script:gui.ChkAuthor.Checked = $true
$script:gui.TxtAuthor.Text = '   '
$script:msgLog.Clear()
Invoke-Click -Control $script:gui.BtnRun
Assert (Msg-Text -match '不能是空白') "refuses a blank author (got: $(Msg-Text))"
Assert ((Get-DocxMeta -Path $target2).Author -eq $authorBefore) 'blank-author refusal wrote nothing'

$script:gui.TxtAuthor.Text = 'SHOULD NOT APPLY'
$script:gui.ChkTime.Checked = $false
$script:msgLog.Clear()
Set-MsgAnswer ([System.Windows.Forms.DialogResult]::No)
Invoke-Click -Control $script:gui.BtnRun
Assert (Msg-Text -match 'CONFIRM\[确认\]') 'the confirmation dialog was reached'
Assert ((Get-DocxMeta -Path $target2).Author -eq $authorBefore) 'cancelling the confirmation leaves the file alone'
Set-MsgAnswer ([System.Windows.Forms.DialogResult]::OK)

$blankDonor = Join-Path $tmp 'blank.docx'
Copy-Item -LiteralPath $sample -Destination $blankDonor
Set-DocxMeta -Path $blankDonor -InPlace -NoBackup -Author '' | Out-Null
Assert ([string]::IsNullOrEmpty((Get-DocxMeta -Path $blankDonor).Author)) 'blank donor really has an empty author'
$script:gui.TxtAuthor.Text = ''
$global:DocxMetaPickFile = $blankDonor
$script:msgLog.Clear()
Invoke-Click -Control $script:gui.BtnDonor
Assert (Msg-Text -match '作者字段为空') "empty donor author produces a warning (got: $(Msg-Text))"
Assert ($script:gui.TxtAuthor.Text -eq '') 'empty donor does not fill the author box'

Write-Output ''
Write-Output '--- 6) locked file surfaces a clear failure ---'
$script:gui.TxtAuthor.Text = 'LOCK TEST'
$script:gui.ChkAuthor.Checked = $true
Set-MsgAnswer ([System.Windows.Forms.DialogResult]::Yes)   # section 5 left it on No
$lock = [System.IO.File]::Open($target2, 'Open', 'Read', 'None')
$script:msgLog.Clear()
Invoke-Click -Control $script:gui.BtnRun
$lock.Close()
Assert (Msg-Text -match 'MSG\[错误\] 处理失败') "locked file shows the failure dialog (got: $(Msg-Text))"
Assert ($script:gui.LblStatus.Text -match '失败') "status line reports the failure (got '$($script:gui.LblStatus.Text)')"
Assert ($script:gui.BtnRun.Enabled -eq $true) 'run button re-enabled after a failure'

Write-Output ''
Write-Output '--- 7) the scratch folder must be the only place touched ---'
# scratch is wiped at the start of every run, so THIS run must have created exactly
# the two backups it expects, and nothing else anywhere.
# scratch is wiped at the start of every run. Section 3 saved as a NEW file (no backup),
# section 4 edited in place (with backup), so exactly one backup must exist.
$strays = @(Get-ChildItem -LiteralPath $tmp -Filter '*.backup.docx' -File | Sort-Object Name)
Assert ($strays.Count -eq 1) "exactly 1 backup in scratch (got $($strays.Count))"
Assert ($strays[0].Name -eq 'target2.backup.docx') "the backup is the in-place target's (got $($strays[0].Name))"
Write-Output ''
Write-Output "=== FAILURES: $script:failed ==="

# non-zero exit code so _test-all.bat and CI can actually detect failures
if ($script:failed -gt 0) { exit 1 } else { exit 0 }
