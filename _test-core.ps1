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
Write-Output '--- 9) created / content-modified dates ---'
# A local time must be stored as its UTC equivalent and read back as the same local time.
$wantC  = Get-Date '2015-06-07 08:09:00'
$wantCM = Get-Date '2016-07-08 09:10:11'
Set-DocxMeta -Path $f1 -InPlace -NoBackup -Created $wantC -ContentModified $wantCM | Out-Null
$dc = Get-DocxMeta $f1
Assert ($dc.Created.ToString('yyyy-MM-dd HH:mm:ss') -eq $wantC.ToString('yyyy-MM-dd HH:mm:ss')) "Created round-trips as local time (got $($dc.Created))"
Assert ($dc.ContentModified.ToString('yyyy-MM-dd HH:mm:ss') -eq $wantCM.ToString('yyyy-MM-dd HH:mm:ss')) "ContentModified round-trips as local time"

# stored value must be UTC, with the xsi:type attribute Word expects
$zipC = [System.IO.Compression.ZipFile]::OpenRead($f1)
$ec = $zipC.GetEntry('docProps/core.xml'); $sc = $ec.Open(); $mc = New-Object System.IO.MemoryStream; $sc.CopyTo($mc); $sc.Close(); $zipC.Dispose()
$xmlC = [System.Text.Encoding]::UTF8.GetString($mc.ToArray())
$stamp = [regex]::Match($xmlC, '<dcterms:created[^>]*>([^<]*)</dcterms:created>').Groups[1].Value
Assert ($stamp -eq $wantC.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')) "created stored as UTC (got $stamp)"
Assert ($xmlC -match '<dcterms:created[^>]*xsi:type="dcterms:W3CDTF"' ) "created carries xsi:type"
Assert ($xmlC -match 'xmlns:xsi=') "xmlns:xsi still declared"
Assert ([regex]::Matches($xmlC, '<dcterms:created').Count -eq 1) "exactly one dcterms:created (no duplicate)"

# a part that lacks the element must get it created in the RIGHT namespace
$f6 = Join-Path $tmp 'f.docx'
$null = New-TestDocx -Path $f6 -Text 'no dates'
$zipD = [System.IO.Compression.ZipFile]::Open($f6, 'Update')
$ed = $zipD.GetEntry('docProps/core.xml'); $sd = $ed.Open(); $md = New-Object System.IO.MemoryStream; $sd.CopyTo($md); $sd.Close()
$td = [System.Text.Encoding]::UTF8.GetString($md.ToArray()) -replace '<dcterms:created[^>]*>[^<]*</dcterms:created>', ''
$ed.Delete()
$nd = $zipD.CreateEntry('docProps/core.xml'); $od = $nd.Open(); $bd = (New-Object System.Text.UTF8Encoding($false)).GetBytes($td); $od.Write($bd,0,$bd.Length); $od.Close()
$zipD.Dispose()
Assert ($null -eq (Get-DocxMeta $f6).Created) "fixture without the element reads back as none"
Set-DocxMeta -Path $f6 -InPlace -NoBackup -Created (Get-Date '2019-09-09 09:09:09') | Out-Null
$zipE = [System.IO.Compression.ZipFile]::OpenRead($f6)
$ee = $zipE.GetEntry('docProps/core.xml'); $se = $ee.Open(); $me = New-Object System.IO.MemoryStream; $se.CopyTo($me); $se.Close(); $zipE.Dispose()
$xmlE = [System.Text.Encoding]::UTF8.GetString($me.ToArray())
Assert ($xmlE -match '<dcterms:created[^>]*xsi:type="dcterms:W3CDTF"' ) "missing element created with the right namespace and attribute"
Assert (-not ($xmlE -match '<cp:created')) "no stray cp:created element"
Assert ((Get-DocxMeta $f6).Created.ToString('yyyy-MM-dd HH:mm:ss') -eq '2019-09-09 09:09:09') "newly created element round-trips"

# app.xml must keep working (its vocabulary is in a different namespace)
Set-DocxMeta -Path $f6 -InPlace -NoBackup -TotalMinutes 42 | Out-Null
Assert ((Get-DocxMeta $f6).TotalMinutes -eq 42) "TotalTime still writes and reads (app.xml namespace)"

Write-Output ''
Write-Output '--- 10) filesystem timestamps ---'
$created = Get-Date '2001-01-01 01:01:01'
$written = Get-Date '2002-02-02 02:02:02'
$tsr = Set-FileTimestamps -Path $f6 -CreationTime $created -LastWriteTime $written
Assert ($tsr.CreationTime.ToString('yyyy-MM-dd HH:mm:ss') -eq '2001-01-01 01:01:01') "filesystem CreationTime set"
Assert ($tsr.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss') -eq '2002-02-02 02:02:02') "filesystem LastWriteTime set"
$threw = $false; try { Set-FileTimestamps -Path $f6 | Out-Null } catch { $threw = $true }
Assert $threw "Set-FileTimestamps rejects an empty request"

Write-Output ''
Write-Output '--- 11) parts stay well-formed after the date edits ---'
foreach ($f in @($f6)) {
    $r6 = Test-PartsStrict $f
    Assert ($r6.Bad.Count -eq 0) "all $($r6.Entries) parts parse: $(Split-Path $f -Leaf)"
    foreach ($b in $r6.Bad) { Write-Output "      bad -> $b" }
}
Write-Output ''
Write-Output "=== FAILURES: $script:failed ==="

# non-zero exit code so _test-all.bat and CI can actually detect failures
if ($script:failed -gt 0) { exit 1 } else { exit 0 }
