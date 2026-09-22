# _gui-smoke.ps1 -- launch the GUI for real, screenshot the window, then close it.
# Usage: _gui-smoke.ps1 [-GuiPath <path to DocxMetaGui.ps1>]   (default: next to this file)
# ASCII-only.

param([string] $GuiPath)

$ErrorActionPreference = 'Stop'
if (-not $GuiPath) { $GuiPath = Join-Path $PSScriptRoot 'DocxMetaGui.ps1' }
$outPng = Join-Path (Split-Path -Parent (Resolve-Path $GuiPath).Path) '_gui.png'

Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class W {
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
    [DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr h, uint msg, IntPtr w, IntPtr l);
    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left, Top, Right, Bottom; }
}
"@

$proc = Start-Process powershell.exe -ArgumentList @('-NoProfile','-ExecutionPolicy','Bypass','-STA','-File',"`"$GuiPath`"") -PassThru
Write-Output "STARTED pid=$($proc.Id) gui=$GuiPath"

$hwnd = [IntPtr]::Zero
$deadline = (Get-Date).AddSeconds(25)
while ((Get-Date) -lt $deadline) {
    Start-Sleep -Milliseconds 400
    $p = Get-Process -Id $proc.Id -ErrorAction SilentlyContinue
    if ($p -and $p.MainWindowHandle -ne 0) { $hwnd = $p.MainWindowHandle; break }
}

if ($hwnd -eq [IntPtr]::Zero) {
    Write-Output 'WINDOW_NOT_FOUND'
    if (-not $proc.HasExited) { $proc.Kill() }
    exit 1
}

Write-Output "WINDOW_FOUND hwnd=$hwnd"
Write-Output ("TITLE=" + (Get-Process -Id $proc.Id).MainWindowTitle)
$null = [W]::SetForegroundWindow($hwnd)
Start-Sleep -Milliseconds 900

$r = New-Object W+RECT
$null = [W]::GetWindowRect($hwnd, [ref]$r)
$w = $r.Right - $r.Left; $h = $r.Bottom - $r.Top
Write-Output "RECT=${w}x${h} at ($($r.Left),$($r.Top))"

$bmp = New-Object System.Drawing.Bitmap($w, $h)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.CopyFromScreen($r.Left, $r.Top, 0, 0, (New-Object System.Drawing.Size($w, $h)))
$bmp.Save($outPng, [System.Drawing.Imaging.ImageFormat]::Png)
$g.Dispose(); $bmp.Dispose()
Write-Output "SCREENSHOT=$outPng"

$null = [W]::SendMessage($hwnd, 0x0010, [IntPtr]::Zero, [IntPtr]::Zero)
if ($proc.WaitForExit(8000)) { Write-Output "CLOSED_CLEANLY exit=$($proc.ExitCode)" }
else { Write-Output 'DID_NOT_CLOSE_CLEANLY'; $proc.Kill() }
