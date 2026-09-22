# _test-core.ps1 -- logic-only test harness (no GUI). Kept ASCII-only.
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'DocxMeta.ps1')

$tmp = Join-Path $PSScriptRoot '_tmp'
if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Recurse -Force }
$null = New-Item -ItemType Directory -Path $tmp -Force

# self-contained fixture: build a valid docx rather than depending on any local file
. (Join-Path $PSScriptRoot 'Make-TestDocx.ps1')
$fixtureDir = Join-Path $PSScriptRoot '_fixtures'
$sample = Join-Path $fixtureDir 'sample.docx'
$null = New-TestDocx -Path $sample -Text 'fixture 测试文档'
Write-Output "SAMPLE (generated): $sample"

$script:failed = 0
function Assert($cond, $msg) {
    if ($cond) { Write-Output "PASS: $msg" } else { Write-Output "FAIL: $msg"; $script:failed++ }
}

# Strict part check: parse raw BYTES with the encoding taken from the declaration.
Add-Type -AssemblyName System.IO.Compression.FileSystem
function Test-PartsStrict($file) {
    $zip = [System.IO.Compression.ZipFile]::OpenRead($file)
    $bad = @()
    foreach ($e in $zip.Entries) {
        $s = $e.Open(); $ms = New-Object System.IO.MemoryStream
        $s.CopyTo($ms); $s.Close()
        $bytes = $ms.ToArray()
        try {
            $ms2 = New-Object System.IO.MemoryStream(,$bytes)
            $rd = [System.Xml.XmlReader]::Create($ms2)
            while ($rd.Read()) {}
            $rd.Close()
        } catch { $bad += "$($e.FullName): $($_.Exception.Message)" }
    }
    $n = $zip.Entries.Count
    $zip.Dispose()
    [pscustomobject]@{ Entries = $n; Bad = $bad }
}

function Get-AppXmlText($file) {
    $zip = [System.IO.Compression.ZipFile]::OpenRead($file)
    $bytes = Get-ZipBytes $zip 'docProps/app.xml'
    $zip.Dispose()
    ConvertFrom-Utf8Bytes $bytes
}

$f1 = Join-Path $tmp 'a.docx'    # time-only, in place
$f2 = Join-Path $tmp 'b.docx'    # author donor
Copy-Item -LiteralPath $sample -Destination $f1
Copy-Item -LiteralPath $sample -Destination $f2

Write-Output '--- baseline ---'
$base = Get-DocxMeta $f1
Write-Output ("author='{0}' lastMod='{1}' total={2} app={3}" -f $base.Author, $base.LastModifiedBy, $base.TotalMinutes, $base.App)

Write-Output ''
Write-Output '--- 0) source part encoding is sane (regression guard) ---'
$appText = Get-AppXmlText $f1
Assert ($appText -match 'encoding="utf-8"') "app.xml declares utf-8 (got: $($appText.Substring(0,50)))"

Write-Output ''
Write-Output '--- 1) set author on donor file B ---'
Set-DocxMeta -Path $f2 -Author 'ZHANG SAN' -ModifiedBy 'ZHANG SAN' | Out-Null
$m2 = Get-DocxMeta $f2
Assert ($m2.Author -eq 'ZHANG SAN') "author written (got '$($m2.Author)')"
Assert ($m2.LastModifiedBy -eq 'ZHANG SAN') "lastModifiedBy written"
Assert ($m2.TotalMinutes -eq $base.TotalMinutes) "TotalTime untouched by author-only edit"
Assert ((Get-DocxMeta $sample).Author -eq $base.Author) "source file on disk untouched"

Write-Output ''
Write-Output '--- 2) time only, in place, backup created ---'
Set-DocxMeta -Path $f1 -InPlace -TotalMinutes 200 | Out-Null
$m1 = Get-DocxMeta $f1
Assert ($m1.TotalMinutes -eq 200) "TotalTime=200 (got $($m1.TotalMinutes))"
$bakPath = Join-Path $tmp 'a.backup.docx'
Assert (Test-Path -LiteralPath $bakPath) "backup file created"
Assert ((Get-DocxMeta $bakPath).TotalMinutes -eq $base.TotalMinutes) "backup keeps original time"

Write-Output ''
Write-Output '--- 3) author + time together, explicit output path ---'
$f3 = Join-Path $tmp 'c.docx'
Set-DocxMeta -Path $f1 -OutputPath $f3 -Author 'LI SI' -TotalMinutes 15 | Out-Null
$m3 = Get-DocxMeta $f3
Assert ($m3.Author -eq 'LI SI') "combined: author"
Assert ($m3.TotalMinutes -eq 15) "combined: time"
Assert ((Get-DocxMeta $f1).TotalMinutes -eq 200) "source of a copy is not mutated"

Write-Output ''
Write-Output '--- 4) copy author from B into a fresh copy of the original ---'
$f4 = Join-Path $tmp 'd.docx'
$srcMeta = Get-DocxMeta $f2
Set-DocxMeta -Path $sample -OutputPath $f4 -Author $srcMeta.Author -ModifiedBy $srcMeta.LastModifiedBy | Out-Null
$m4 = Get-DocxMeta $f4
Assert ($m4.Author -eq $m2.Author) "author copied from B ('$($m4.Author)')"
Assert ($m4.LastModifiedBy -eq $m2.LastModifiedBy) "lastModifiedBy copied from B"

Write-Output ''
Write-Output '--- 5) unicode author, 6h exact, idempotency ---'
$f5 = Join-Path $tmp 'e.docx'
$cn = [string]([char]0x5F20 + [char]0x4E09)
Set-DocxMeta -Path $sample -OutputPath $f5 -Author $cn -TotalMinutes 360 | Out-Null
$m5 = Get-DocxMeta $f5
Assert ($m5.Author -eq $cn) "unicode author round-trips (got '$($m5.Author)')"
Assert ($m5.TotalMinutes -eq 360) "6h = 360 minutes"
$len1 = (Get-Item -LiteralPath $f5).Length
Set-DocxMeta -Path $f5 -InPlace -NoBackup -TotalMinutes 360 -Author $cn | Out-Null
$len2 = (Get-Item -LiteralPath $f5).Length
Assert ($len1 -eq $len2) "idempotent re-run ($len1 vs $len2 bytes)"

Write-Output ''
Write-Output '--- 6) strict part validation (bytes parsed per declared encoding) ---'
foreach ($f in @($f1, $f2, $f3, $f4, $f5)) {
    $r = Test-PartsStrict $f
    Assert ($r.Bad.Count -eq 0) "all $($r.Entries) parts parse: $(Split-Path $f -Leaf)"
    foreach ($b in $r.Bad) { Write-Output "      bad -> $b" }
}
$rApp = Get-AppXmlText $f5
Assert ($rApp -match 'encoding="utf-8"') "rewritten app.xml still declares utf-8"
Assert ($rApp -notmatch 'utf-16') "no stale utf-16 declaration anywhere"

Write-Output ''
Write-Output '--- 7) error handling ---'
$threw = $false; try { Set-DocxMeta -Path $f1 -InPlace | Out-Null } catch { $threw = $true }
Assert $threw "throws when no operation requested"
$threw = $false; try { Set-DocxMeta -Path $f1 -InPlace -TotalMinutes -5 | Out-Null } catch { $threw = $true }
Assert $threw "throws on negative minutes"
# A blank author is not an error: it clears the field (used by the GUI's donor flow)
Set-DocxMeta -Path $f2 -InPlace -NoBackup -Author '' | Out-Null
Assert ([string]::IsNullOrEmpty((Get-DocxMeta $f2).Author)) "blank author clears the field"
Set-DocxMeta -Path $f2 -InPlace -NoBackup -Author 'REFILL' | Out-Null
Assert ((Get-DocxMeta $f2).Author -eq 'REFILL') "author can be written again after clearing"
# PowerShell coerces $null to "" for a [string] parameter, so this is the same clear-the-field path
Set-DocxMeta -Path $f2 -InPlace -NoBackup -Author $null | Out-Null
Assert ([string]::IsNullOrEmpty((Get-DocxMeta $f2).Author)) "null author is coerced to empty and clears the field"
$threw = $false
try { Set-DocxMeta -Path $f3 -OutputPath $f4 | Out-Null } catch { $threw = $true }
Assert $threw "throws when -OutputPath given without any change"

Write-Output ''
Write-Output '--- 8) locked file gives a clear error ---'
$lock = [System.IO.File]::Open($f1, 'Open', 'Read', 'None')
$threw = $false
try { Set-DocxMeta -Path $f1 -InPlace -TotalMinutes 10 | Out-Null }
catch { $threw = $true; Write-Output "      (msg: $($_.Exception.Message))" }
$lock.Close()
Assert $threw "locked file reported, not silently ignored"

Write-Output ''
Write-Output "=== FAILURES: $script:failed ==="

# non-zero exit code so _test-all.bat and CI can actually detect failures
if ($script:failed -gt 0) { exit 1 } else { exit 0 }
