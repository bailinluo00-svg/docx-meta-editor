# DocxMeta.ps1 -- core logic: read/write docx metadata (author, last modifier, total editing time)
#
# Everything here works on raw BYTES, never on .NET strings fed back into a
# StringWriter. That matters: XmlDocument.Save(StringWriter) always emits
# encoding="utf-16" in the declaration while the zip entry holds UTF-8 bytes,
# which produces a package strict parsers reject.
#
# This file is ASCII-only on purpose; user-facing Chinese strings live in callers.

Add-Type -AssemblyName System.IO.Compression | Out-Null
Add-Type -AssemblyName System.IO.Compression.FileSystem | Out-Null

function Get-ZipBytes {
    param(
        [Parameter(Mandatory)][System.IO.Compression.ZipArchive] $Zip,
        [Parameter(Mandatory)][string] $EntryName
    )
    $e = $Zip.GetEntry($EntryName)
    if ($null -eq $e) { return $null }
    $s = $e.Open()
    try {
        $ms = New-Object System.IO.MemoryStream
        $s.CopyTo($ms)
        return $ms.ToArray()
    } finally { $s.Close() }
}

function ConvertFrom-Utf8Bytes {
    param([Parameter(Mandatory)][byte[]] $Bytes)
    $enc = New-Object System.Text.UTF8Encoding($false)
    $text = $enc.GetString($Bytes)
    # tolerate a UTF-8 BOM and a stale UTF-16 declaration from older Word builds
    $text = $text.TrimStart([char]0xFEFF)
    $text = [regex]::Replace($text, '^(\s*<\?xml[^>]*?encoding\s*=\s*")[^"]*(")', '${1}utf-8${2}')
    return $text
}

function Set-ZipBytes {
    param(
        [Parameter(Mandatory)][System.IO.Compression.ZipArchive] $Zip,
        [Parameter(Mandatory)][string] $EntryName,
        [Parameter(Mandatory)][byte[]] $Bytes
    )
    $old = $Zip.GetEntry($EntryName)
    if ($old) { $null = $old.Delete() }
    $new = $Zip.CreateEntry($EntryName, [System.IO.Compression.CompressionLevel]::Optimal)
    $s = $new.Open()
    try {
        $s.Write($Bytes, 0, $Bytes.Length)
        $s.Flush()
    } finally { $s.Close() }
}

function Get-DocxMeta {
    <#
      Returns: Path, Author, LastModifiedBy, TotalMinutes, HasTotalTime, App
      Author / LastModifiedBy / TotalMinutes are $null when the part or element is absent.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string] $Path)

    $Path = (Resolve-Path -LiteralPath $Path).Path
    $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $coreBytes = Get-ZipBytes $zip 'docProps/core.xml'
        $appBytes  = Get-ZipBytes $zip 'docProps/app.xml'
    } finally { $zip.Dispose() }

    function Get-ElementText {
        param([byte[]] $Bytes, [string] $LocalName)
        if ($null -eq $Bytes) { return $null }
        $doc = New-Object System.Xml.XmlDocument
        $doc.LoadXml((ConvertFrom-Utf8Bytes $Bytes))
        $ns = $doc.DocumentElement.NamespaceURI
        $n = $doc.SelectSingleNode("//*[local-name()='$LocalName' and namespace-uri()='$ns']")
        if ($n) { return $n.InnerText }
        return $null
    }

    $author   = Get-ElementText $coreBytes 'creator'
    $modBy    = Get-ElementText $coreBytes 'lastModifiedBy'
    $totalRaw = Get-ElementText $appBytes  'TotalTime'
    $app      = Get-ElementText $appBytes  'Application'

    $total = $null
    if ($null -ne $totalRaw -and $totalRaw -match '^\s*\d+\s*$') { $total = [int]$totalRaw }

    [pscustomobject]@{
        Path           = $Path
        Author         = $author
        LastModifiedBy = $modBy
        TotalMinutes   = $total
        HasTotalTime   = ($null -ne $total)
        App            = $app
    }
}

function Set-DocxMeta {
    <#
      Rewrites only the zip entries that actually change.
        -TotalMinutes : set/insert <TotalTime> in docProps/app.xml
        -Author       : set/insert <dc:creator> in docProps/core.xml
        -ModifiedBy   : set/insert <cp:lastModifiedBy> in docProps/core.xml
        -OutputPath   : write a NEW file, original untouched
        -InPlace      : modify the file itself (makes <stem>.backup.docx unless -NoBackup)
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string] $Path,
        [int]    $TotalMinutes = 0,
        [string] $Author,
        [string] $ModifiedBy,
        [string] $OutputPath,
        [switch] $InPlace,
        [switch] $NoBackup
    )

    $setTime   = $PSBoundParameters.ContainsKey('TotalMinutes')
    $setAuthor = $PSBoundParameters.ContainsKey('Author')
    $setModBy  = $PSBoundParameters.ContainsKey('ModifiedBy')

    if (-not ($setTime -or $setAuthor -or $setModBy)) {
        throw 'nothing to do: pass -TotalMinutes and/or -Author and/or -ModifiedBy'
    }
    if ($setTime -and $TotalMinutes -lt 0) { throw '-TotalMinutes must be >= 0' }
    # An empty string is deliberate: it clears the field. Only $null is rejected.
    if ($setAuthor -and $null -eq $Author) { throw '-Author must not be $null (pass "" to clear it)' }

    $src = (Resolve-Path -LiteralPath $Path).Path
    # default is to edit the file itself (with a backup); -OutputPath switches to copy mode
    $inPlaceMode = (-not $OutputPath)

    if ($inPlaceMode) {
        $target = $src
        if (-not $NoBackup) {
            $stem = [System.IO.Path]::GetFileNameWithoutExtension($src)
            $bak  = Join-Path ([System.IO.Path]::GetDirectoryName($src)) ($stem + '.backup.docx')
            if (-not (Test-Path -LiteralPath $bak)) {
                Copy-Item -LiteralPath $src -Destination $bak
            }
        }
    } else {
        $outDir = [System.IO.Path]::GetDirectoryName($OutputPath)
        if (-not $outDir) { $OutputPath = Join-Path (Get-Location).Path $OutputPath; $outDir = [System.IO.Path]::GetDirectoryName($OutputPath) }
        if (-not (Test-Path -LiteralPath $outDir)) { $null = New-Item -ItemType Directory -Path $outDir -Force }
        $target = [System.IO.Path]::GetFullPath($OutputPath)
        if ($target -ne $src) { Copy-Item -LiteralPath $src -Destination $target -Force }
    }

    $changed = New-Object System.Collections.Generic.List[string]
    $utf8 = New-Object System.Text.UTF8Encoding($false)

    $zip = [System.IO.Compression.ZipFile]::Open($target, [System.IO.Compression.ZipArchiveMode]::Update)
    try {
        # ---------- docProps/app.xml : TotalTime ----------
        if ($setTime) {
            $appBytes = Get-ZipBytes $zip 'docProps/app.xml'
            if ($null -eq $appBytes) {
                $blank = '<?xml version="1.0" encoding="utf-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"></Properties>'
                $appBytes = $utf8.GetBytes($blank)
            }
            $appBytes = Set-XmlElement -Bytes $appBytes -LocalName 'TotalTime' -Value ([string]$TotalMinutes)
            Set-ZipBytes $zip 'docProps/app.xml' $appBytes
            $changed.Add("TotalTime=$TotalMinutes")
        }

        # ---------- docProps/core.xml : author / lastModifiedBy ----------
        if ($setAuthor -or $setModBy) {
            $coreBytes = Get-ZipBytes $zip 'docProps/core.xml'
            if ($null -eq $coreBytes) {
                $blank = '<?xml version="1.0" encoding="utf-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"></cp:coreProperties>'
                $coreBytes = $utf8.GetBytes($blank)
            }
            if ($setAuthor) {
                $coreBytes = Set-XmlElement -Bytes $coreBytes -LocalName 'creator' -Value $Author
                $changed.Add("Author=$Author")
            }
            if ($setModBy) {
                $coreBytes = Set-XmlElement -Bytes $coreBytes -LocalName 'lastModifiedBy' -Value $ModifiedBy
                $changed.Add("LastModifiedBy=$ModifiedBy")
            }
            Set-ZipBytes $zip 'docProps/core.xml' $coreBytes
        }
    } finally { $zip.Dispose() }

    [pscustomobject]@{
        Path     = $target
        Original = $src
        Changed  = ($changed -join ', ')
        InPlace  = $inPlaceMode
    }
}

function Set-XmlElement {
    <#
      Sets the text of the first element matching $LocalName (any prefix), creating it
      before the root close tag using the prefix the root already declares.
      Returns UTF-8 BYTES so the XML declaration always matches the stored encoding.
    #>
    param(
        [Parameter(Mandatory)][byte[]] $Bytes,
        [Parameter(Mandatory)][string] $LocalName,
        [Parameter(Mandatory)][AllowEmptyString()][string] $Value
    )

    $doc = New-Object System.Xml.XmlDocument
    $doc.LoadXml((ConvertFrom-Utf8Bytes $Bytes))

    $rootNs = $doc.DocumentElement.NamespaceURI
    $node = $doc.SelectSingleNode("//*[local-name()='$LocalName' and namespace-uri()='$rootNs']")

    if ($node) {
        $node.InnerText = $Value
    } else {
        $prefix = $null
        $attrs = $doc.DocumentElement.Attributes
        for ($i = 0; $i -lt $attrs.Count; $i++) {
            $a = $attrs.Item($i)
            if ($a.Name -eq 'xmlns') {
                if ($a.Value -eq $rootNs) { $prefix = '' }
            } elseif ($a.Name -like 'xmlns:*') {
                if ($a.Value -eq $rootNs) { $prefix = $a.Name.Substring(6) }
            }
        }
        if ($null -eq $prefix) { throw "root element does not declare namespace $rootNs" }
        $qualified = if ($prefix -eq '') { $LocalName } else { "$prefix`:$LocalName" }
        $el = $doc.CreateElement($qualified, $rootNs)
        $el.InnerText = $Value
        $null = $doc.DocumentElement.AppendChild($el)
    }

    $settings = New-Object System.Xml.XmlWriterSettings
    $settings.Encoding = New-Object System.Text.UTF8Encoding($false)
    $settings.OmitXmlDeclaration = $false
    $settings.Indent = $false
    $ms = New-Object System.IO.MemoryStream
    $xw = [System.Xml.XmlWriter]::Create($ms, $settings)
    try {
        $doc.Save($xw)
        $xw.Flush()
        return $ms.ToArray()
    } finally { $xw.Close() }
}

function ConvertTo-TimeParts {
    param([Parameter(Mandatory)][int] $Minutes)
    [pscustomobject]@{
        Hours   = [math]::Floor($Minutes / 60)
        Minutes = ($Minutes % 60)
        Text    = ("{0} h {1} min" -f [math]::Floor($Minutes / 60), ($Minutes % 60))
    }
}
