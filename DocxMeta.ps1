# DocxMeta.ps1 -- core logic: read/write docx metadata
#   author / last modified by / total editing time / created & modified dates
#
# Everything here works on raw BYTES, never on .NET strings fed back into a
# StringWriter. That matters: XmlDocument.Save(StringWriter) always emits
# encoding="utf-16" in the declaration while the zip entry holds UTF-8 bytes,
# which produces a package strict parsers reject.
#
# This file is ASCII-only on purpose; user-facing Chinese strings live in callers.

Add-Type -AssemblyName System.IO.Compression | Out-Null
Add-Type -AssemblyName System.IO.Compression.FileSystem | Out-Null

# ---------------------------------------------------------------- zip helpers
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

function ConvertFrom-Utf8Bytes {
    param([Parameter(Mandatory)][byte[]] $Bytes)
    $enc = New-Object System.Text.UTF8Encoding($false)
    $text = $enc.GetString($Bytes)
    # tolerate a UTF-8 BOM and a stale UTF-16 declaration from older Word builds
    $text = $text.TrimStart([char]0xFEFF)
    $text = [regex]::Replace($text, '^(\s*<\?xml[^>]*?encoding\s*=\s*")[^"]*(")', '${1}utf-8${2}')
    return $text
}

function Write-XmlBytes {
    <# Serialises an XmlDocument to UTF-8 bytes so the declaration always matches. #>
    param([Parameter(Mandatory)][System.Xml.XmlDocument] $Doc)
    $settings = New-Object System.Xml.XmlWriterSettings
    $settings.Encoding = New-Object System.Text.UTF8Encoding($false)
    $settings.OmitXmlDeclaration = $false
    $settings.Indent = $false
    $ms = New-Object System.IO.MemoryStream
    $xw = [System.Xml.XmlWriter]::Create($ms, $settings)
    try {
        $Doc.Save($xw)
        $xw.Flush()
        return $ms.ToArray()
    } finally { $xw.Close() }
}

# ---------------------------------------------------------------- namespaces
function Get-PreferredElementNamespace {
    <#
      PREFERENCE list of namespace URIs for a local name, best first. Only a hint:
      docx parts use different vocabularies and app.xml shares no namespace with
      core.xml. Callers always prefer the namespace an element already lives in.
    #>
    param([Parameter(Mandatory)][string] $LocalName)

    $dcterms = 'http://purl.org/dc/terms/'
    $dc      = 'http://purl.org/dc/elements/1.1/'
    $cp      = 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties'
    $app     = 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties'

    switch ($LocalName) {
        'creator'        { return @($dc) }
        'lastModifiedBy' { return @($cp) }
        'created'        { return @($dcterms, $cp) }
        'modified'       { return @($dcterms, $cp) }
        'TotalTime'      { return @($app) }
        'Application'    { return @($app) }
        default          { return @($cp) }
    }
}

function Get-ElementNamespace {
    <#
      The namespace URI the element with this local name already lives in, or $null.
      Lets a part whose vocabulary is not in the preference table still work.
    #>
    param(
        [Parameter(Mandatory)][System.Xml.XmlDocument] $Doc,
        [Parameter(Mandatory)][string] $LocalName
    )
    $n = $Doc.SelectSingleNode('//*[local-name()="' + $LocalName + '"]')
    if ($n) { return $n.NamespaceURI }
    return $null
}

function Select-XmlElement {
    <#
      First element with this local name, ANY namespace. Deliberately
      namespace-agnostic: the same local name appears in several vocabularies and the
      parts we touch use different ones. Get-TargetNamespace decides where a newly
      created element belongs.
    #>
    param(
        [Parameter(Mandatory)][System.Xml.XmlDocument] $Doc,
        [Parameter(Mandatory)][string] $LocalName
    )
    return $Doc.SelectSingleNode('//*[local-name()="' + $LocalName + '"]')
}

function Get-TargetNamespace {
    <#
      Namespace an element should live in: whatever it already uses, else the first
      preferred candidate the root already declares, else the preferred one newly
      declared on the root (the caller needs a usable prefix).

      Note: GetPrefixOfNamespace returns an EMPTY STRING - not $null - for an
      undeclared namespace, so test the declarations directly instead.
    #>
    param(
        [Parameter(Mandatory)][System.Xml.XmlDocument] $Doc,
        [Parameter(Mandatory)][string] $LocalName
    )

    $existing = Get-ElementNamespace -Doc $Doc -LocalName $LocalName
    if ($existing) { return $existing }

    $declared = @{}
    foreach ($a in $Doc.DocumentElement.Attributes) {
        if ($a.Name -eq 'xmlns') { $declared[$a.Value] = '' }
        elseif ($a.Name -like 'xmlns:*') { $declared[$a.Value] = $a.Name.Substring(6) }
    }

    foreach ($uri in (Get-PreferredElementNamespace -LocalName $LocalName)) {
        if ($declared.ContainsKey($uri)) { return $uri }
    }

    $wantUri = (Get-PreferredElementNamespace -LocalName $LocalName)[0]
    $wantPrefix = switch ($wantUri) {
        'http://purl.org/dc/terms/'                                                 { 'dcterms' }
        'http://purl.org/dc/elements/1.1/'                                          { 'dc' }
        'http://schemas.openxmlformats.org/package/2006/metadata/core-properties'    { 'cp' }
        'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties'  { '' }
        default                                                                     { 'ns' }
    }
    if ($wantPrefix -eq '') {
        $null = $Doc.DocumentElement.SetAttribute('xmlns', $wantUri)
    } else {
        $null = $Doc.DocumentElement.SetAttribute('xmlns:' + $wantPrefix, $wantUri)
    }
    return $wantUri
}

function Set-XmlElement {
    <#
      Sets the text of the first element matching $LocalName (any namespace), creating
      it if absent, and optionally sets one attribute on it (xsi:type for the dcterms
      date elements). Returns UTF-8 BYTES.
    #>
    param(
        [Parameter(Mandatory)][byte[]] $Bytes,
        [Parameter(Mandatory)][string] $LocalName,
        [Parameter(Mandatory)][AllowEmptyString()][string] $Value,
        [string] $AttributeName,
        [string] $AttributeValue
    )

    $doc = New-Object System.Xml.XmlDocument
    $doc.LoadXml((ConvertFrom-Utf8Bytes $Bytes))

    $root = $doc.DocumentElement
    $node = Select-XmlElement -Doc $doc -LocalName $LocalName

    if (-not $node) {
        $targetNs = Get-TargetNamespace -Doc $doc -LocalName $LocalName

        $prefix = $null
        foreach ($a in $root.Attributes) {
            if ($a.Name -eq 'xmlns') {
                if ($a.Value -eq $targetNs) { $prefix = '' }
            } elseif ($a.Name -like 'xmlns:*') {
                if ($a.Value -eq $targetNs) { $prefix = $a.Name.Substring(6) }
            }
        }
        if ($null -eq $prefix) {
            throw ('root element does not declare namespace ' + $targetNs)
        }
        $qualified = if ($prefix -eq '') { $LocalName } else { $prefix + ':' + $LocalName }
        $node = $doc.CreateElement($qualified, $targetNs)
        $null = $root.AppendChild($node)
    }

    $node.InnerText = $Value

    if ($AttributeName) {
        # Go through SetAttributeNode with a real namespace URI. Plain
        # SetAttribute("xsi:type", ...) is parsed as prefix "xsi" + local "type" and
        # collides with the xmlns:xsi declaration ("duplicate attribute name").
        $attrPrefix = if ($AttributeName.Contains(':')) { $AttributeName.Split(':')[0] } else { '' }
        if ($attrPrefix) {
            $uri = $node.GetNamespaceOfPrefix($attrPrefix)
            if ([string]::IsNullOrEmpty($uri)) {
                $nsUri = $null
                foreach ($a in $root.Attributes) {
                    if ($a.Name -eq 'xmlns:' + $attrPrefix) { $nsUri = $a.Value }
                }
                if ([string]::IsNullOrEmpty($nsUri)) {
                    $nsUri = 'http://www.w3.org/2001/XMLSchema-instance'
                    $null = $root.SetAttribute('xmlns:' + $attrPrefix, $nsUri)
                }
            }
            $attrUri = $node.GetNamespaceOfPrefix($attrPrefix)
            if ([string]::IsNullOrEmpty($attrUri)) { $attrUri = 'http://www.w3.org/2001/XMLSchema-instance' }
            $attrNode = $doc.CreateAttribute($AttributeName, $attrUri)
            $attrNode.Value = $AttributeValue
            $null = $node.SetAttributeNode($attrNode)
        } else {
            $node.SetAttribute($AttributeName, $AttributeValue)
        }
    }

    return Write-XmlBytes -Doc $doc
}

# ---------------------------------------------------------------- dates
function ConvertFrom-W3CDTF {
    <#
      Parses a W3CDTF stamp (dcterms:created / dcterms:modified) into a LOCAL [datetime].
      Word writes these in UTC with a trailing Z, e.g. 2026-09-22T02:30:00Z.
      Returns $null when the value is missing or unparseable.
    #>
    param([AllowNull()][string] $Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
    $dt = [datetime]::MinValue
    $ok = [datetime]::TryParse($Value.Trim(),
        [System.Globalization.CultureInfo]::InvariantCulture,
        [System.Globalization.DateTimeStyles]::AdjustToUniversal -bor [System.Globalization.DateTimeStyles]::AssumeUniversal,
        [ref]$dt)
    if (-not $ok) { return $null }
    return $dt.ToLocalTime()
}

function ConvertTo-W3CDTF {
    <#
      Formats a [datetime] as the UTC W3CDTF stamp Word stores. A local input is
      converted to UTC first, because that is what Word and Windows both save and then
      render in the viewer's own timezone.
    #>
    param([Parameter(Mandatory)][datetime] $Value)
    return $Value.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ',
        [System.Globalization.CultureInfo]::InvariantCulture)
}

# ---------------------------------------------------------------- read
function Get-DocxMeta {
    <#
      Returns: Path, Author, LastModifiedBy, Created, ContentModified, TotalMinutes,
               HasTotalTime, App.
      Dates come back as LOCAL [datetime]. Absent fields are $null.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string] $Path)

    $Path = (Resolve-Path -LiteralPath $Path).Path
    $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $coreBytes = Get-ZipBytes $zip 'docProps/core.xml'
        $appBytes  = Get-ZipBytes $zip 'docProps/app.xml'
    } finally { $zip.Dispose() }

    function Get-Text {
        param([byte[]] $Buf, [string] $LocalName, [string[]] $Namespace)
        if ($null -eq $Buf) { return $null }
        $doc = New-Object System.Xml.XmlDocument
        $doc.LoadXml((ConvertFrom-Utf8Bytes $Buf))

        if ($Namespace) {
            foreach ($uri in $Namespace) {
                $x = '//*[local-name()="' + $LocalName + '" and namespace-uri()="' + $uri + '"]'
                $n = $doc.SelectSingleNode($x)
                if ($n) { return $n.InnerText }
            }
            return $null
        }

        $all = @($doc.SelectNodes('//*[local-name()="' + $LocalName + '"]'))
        if ($all.Count -eq 0) { return $null }
        # a duplicate in another namespace is leftover corruption from an older bug
        if ($all.Count -gt 1) {
            for ($i = 1; $i -lt $all.Count; $i++) {
                $null = $all[$i].ParentNode.RemoveChild($all[$i])
            }
        }
        return $all[0].InnerText
    }

    $dcterms = 'http://purl.org/dc/terms/'
    $dc      = 'http://purl.org/dc/elements/1.1/'
    $cp      = 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties'

    $author     = Get-Text $coreBytes 'creator'        @($dc)
    $modBy      = Get-Text $coreBytes 'lastModifiedBy' @($cp)
    $createdRaw = Get-Text $coreBytes 'created'        @($dcterms)
    $cmodRaw    = Get-Text $coreBytes 'modified'       @($dcterms)
    $totalRaw   = Get-Text $appBytes  'TotalTime'
    $app        = Get-Text $appBytes  'Application'

    $total = $null
    if ($null -ne $totalRaw -and $totalRaw -match '^\s*\d+\s*$') { $total = [int]$totalRaw }

    [pscustomobject]@{
        Path            = $Path
        Author          = $author
        LastModifiedBy  = $modBy
        Created         = ConvertFrom-W3CDTF $createdRaw
        ContentModified = ConvertFrom-W3CDTF $cmodRaw
        TotalMinutes    = $total
        HasTotalTime    = ($null -ne $total)
        App             = $app
    }
}

# ---------------------------------------------------------------- write
function Set-DocxMeta {
    <#
      Rewrites only the zip entries that actually change.
        -TotalMinutes    : set/insert <TotalTime> in docProps/app.xml
        -Author          : set/insert <dc:creator> in docProps/core.xml
        -ModifiedBy      : set/insert <cp:lastModifiedBy> in docProps/core.xml
        -Created         : set/insert <dcterms:created>  (document "content created")
        -ContentModified : set/insert <dcterms:modified> (document "content modified")
        -OutputPath      : write a NEW file, original untouched
        -InPlace         : modify the file itself (makes <stem>.backup.docx unless -NoBackup)

      The two date parameters write XML metadata only; they do NOT change the file's
      own filesystem timestamps. Use Set-FileTimestamps for those.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string] $Path,
        [int]      $TotalMinutes = 0,
        [string]   $Author,
        [string]   $ModifiedBy,
        [datetime] $Created,
        [datetime] $ContentModified,
        [string]   $OutputPath,
        [switch]   $InPlace,
        [switch]   $NoBackup
    )

    $setTime    = $PSBoundParameters.ContainsKey('TotalMinutes')
    $setAuthor  = $PSBoundParameters.ContainsKey('Author')
    $setModBy   = $PSBoundParameters.ContainsKey('ModifiedBy')
    $setCreated = $PSBoundParameters.ContainsKey('Created')
    $setCMod    = $PSBoundParameters.ContainsKey('ContentModified')

    if (-not ($setTime -or $setAuthor -or $setModBy -or $setCreated -or $setCMod)) {
        throw 'nothing to do: pass at least one of -TotalMinutes / -Author / -ModifiedBy / -Created / -ContentModified'
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
        if (-not $outDir) {
            $OutputPath = Join-Path (Get-Location).Path $OutputPath
            $outDir = [System.IO.Path]::GetDirectoryName($OutputPath)
        }
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
            $null = $changed.Add('TotalTime=' + $TotalMinutes)
        }

        # ---------- docProps/core.xml : author / lastModifiedBy / dates ----------
        if ($setAuthor -or $setModBy -or $setCreated -or $setCMod) {
            $coreBytes = Get-ZipBytes $zip 'docProps/core.xml'
            if ($null -eq $coreBytes) {
                $blank = '<?xml version="1.0" encoding="utf-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"></cp:coreProperties>'
                $coreBytes = $utf8.GetBytes($blank)
            }
            if ($setAuthor) {
                $coreBytes = Set-XmlElement -Bytes $coreBytes -LocalName 'creator' -Value $Author
                $null = $changed.Add('Author=' + $Author)
            }
            if ($setModBy) {
                $coreBytes = Set-XmlElement -Bytes $coreBytes -LocalName 'lastModifiedBy' -Value $ModifiedBy
                $null = $changed.Add('LastModifiedBy=' + $ModifiedBy)
            }
            if ($setCreated) {
                $stamp = ConvertTo-W3CDTF -Value $Created
                $coreBytes = Set-XmlElement -Bytes $coreBytes -LocalName 'created' -Value $stamp -AttributeName 'xsi:type' -AttributeValue 'dcterms:W3CDTF'
                $null = $changed.Add('Created=' + $stamp)
            }
            if ($setCMod) {
                $stamp = ConvertTo-W3CDTF -Value $ContentModified
                $coreBytes = Set-XmlElement -Bytes $coreBytes -LocalName 'modified' -Value $stamp -AttributeName 'xsi:type' -AttributeValue 'dcterms:W3CDTF'
                $null = $changed.Add('ContentModified=' + $stamp)
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

function Set-FileTimestamps {
    <#
      Sets FILESYSTEM timestamps. A different thing from the docx XML metadata:

        filesystem CreationTime  -> Explorer column "Date created"  / 属性 -> 创建时间
        filesystem LastWriteTime -> Explorer column "Date modified"
        dcterms:created          -> Explorer column "Content created"
        dcterms:modified         -> Explorer column "Date last saved"

      Only the parameters you pass are changed.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string] $Path,
        [datetime] $CreationTime,
        [datetime] $LastWriteTime,
        [datetime] $LastAccessTime
    )

    if (-not ($PSBoundParameters.ContainsKey('CreationTime') -or
              $PSBoundParameters.ContainsKey('LastWriteTime') -or
              $PSBoundParameters.ContainsKey('LastAccessTime'))) {
        throw 'nothing to do: pass -CreationTime and/or -LastWriteTime and/or -LastAccessTime'
    }

    $full = (Resolve-Path -LiteralPath $Path).Path
    $item = Get-Item -LiteralPath $full -Force
    $applied = New-Object System.Collections.Generic.List[string]

    if ($PSBoundParameters.ContainsKey('CreationTime')) {
        $item.CreationTime = $CreationTime
        $null = $applied.Add('CreationTime=' + $CreationTime)
    }
    if ($PSBoundParameters.ContainsKey('LastWriteTime')) {
        $item.LastWriteTime = $LastWriteTime
        $null = $applied.Add('LastWriteTime=' + $LastWriteTime)
    }
    if ($PSBoundParameters.ContainsKey('LastAccessTime')) {
        $item.LastAccessTime = $LastAccessTime
        $null = $applied.Add('LastAccessTime=' + $LastAccessTime)
    }

    $after = Get-Item -LiteralPath $full -Force
    [pscustomobject]@{
        Path           = $full
        CreationTime   = $after.CreationTime
        LastWriteTime  = $after.LastWriteTime
        LastAccessTime = $after.LastAccessTime
        Changed        = ($applied -join ', ')
    }
}

function ConvertTo-TimeParts {
    param([Parameter(Mandatory)][int] $Minutes)
    [pscustomobject]@{
        Hours   = [math]::Floor($Minutes / 60)
        Minutes = ($Minutes % 60)
        Text    = ('{0} h {1} min' -f [math]::Floor($Minutes / 60), ($Minutes % 60))
    }
}
