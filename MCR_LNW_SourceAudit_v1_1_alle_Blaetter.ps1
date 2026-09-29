& {
    $ErrorActionPreference = 'Stop'
    Set-StrictMode -Version 2.0
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.IO.Compression

    $out = Join-Path $env:LOCALAPPDATA ('MCR_LNW_AlleBlaetter_' + (Get-Date -Format 'yyyyMMdd_HHmmss_fff'))
    [void][IO.Directory]::CreateDirectory($out)
    $utf8 = [Text.UTF8Encoding]::new($true)
    $summary = New-Object 'System.Collections.Generic.List[string]'
    $summary.Add('MCR_LNW_SOURCE_AUDIT=1.1')
    $summary.Add('Statischer Export. Kein Excel, keine Makros, kein Zugriff auf verknuepfte Dateien.')
    $summary.Add('Ein Dateidialog. Alle vorhandenen Arbeitsblaetter, auch ausgeblendete; kein fester Blattname.')
    $summary.Add('Zeile 1 wird als HEADER_KANDIDAT gelesen, NICHT als gesicherte Kopfzeile.')
    $summary.Add('Keine weiteren Zellwerte/Caches. Tabellennamen, Spaltendefinitionen und Formeln werden gelesen.')
    $summary.Add('Auch Zeile 1 und Formeln koennen personenbezogene oder vertrauliche Angaben enthalten.')
    $summary.Add('Je Spalte erste zwei und letzte Formel; keine Vollstaendigkeitszertifizierung.')
    $summary.Add('Header, Formeln, Dateinamen und Fehler vor Weitergabe pruefen.')

    function Read-Hash([IO.Stream]$stream) {
        $sha = [Security.Cryptography.SHA256]::Create()
        try {
            $stream.Position = 0
            $hash = [BitConverter]::ToString($sha.ComputeHash($stream)).Replace('-', '')
            if ($hash -notmatch '\A[0-9A-F]{64}\z') { throw 'SHA256 ungueltig.' }
            return $hash
        } finally { $sha.Dispose(); $stream.Position = 0 }
    }

    function Read-Xml($archive, [string]$part) {
        $entry = $archive.GetEntry($part)
        if ($null -eq $entry) { throw "XML fehlt: $part" }
        if ($entry.Length -gt 67108864) { throw "XML ueber 64 MiB: $part" }
        $settings = New-Object System.Xml.XmlReaderSettings
        $settings.DtdProcessing = [Xml.DtdProcessing]::Prohibit
        $settings.XmlResolver = $null
        $settings.MaxCharactersInDocument = 67108864
        $stream = $entry.Open()
        $reader = $null
        try {
            $reader = [Xml.XmlReader]::Create($stream, $settings)
            $doc = New-Object System.Xml.XmlDocument
            $doc.XmlResolver = $null
            $doc.PreserveWhitespace = $true
            $doc.Load($reader)
            return ,$doc
        } finally {
            if ($null -ne $reader) { $reader.Dispose() }
            $stream.Dispose()
        }
    }

    function Rel-Id($node) {
        $id = $node.SelectSingleNode("@*[local-name()='id' and namespace-uri()!='']")
        if ($null -eq $id) { throw 'Relationship-ID fehlt.' }
        return $id.Value
    }

    function Relation($archive, [string]$part, [string]$id) {
        $p = $part.LastIndexOf('/')
        $rp = $part.Substring(0, $p + 1) + '_rels/' + $part.Substring($p + 1) + '.rels'
        $doc = Read-Xml $archive $rp
        $hits = @($doc.DocumentElement.SelectNodes("*[local-name()='Relationship']") |
            Where-Object { $_.GetAttribute('Id') -eq $id })
        if ($hits.Count -ne 1) { throw "Relationship nicht eindeutig: $part | $id" }
        return ,($hits[0])
    }

    function Package-Part([string]$from, $rel) {
        if ($rel.GetAttribute('TargetMode') -eq 'External') { throw 'Externes Oeffnen abgelehnt.' }
        $target = $rel.GetAttribute('Target')
        if ([string]::IsNullOrWhiteSpace($target)) { throw 'Leeres Paket-Ziel.' }
        # Nur ZIP-Pfadaufloesung, kein Netzwerkzugriff.
        $u = [Uri]::new([Uri]('https://package.invalid/' + $from), $target)
        if ($u.Authority -ne 'package.invalid' -or $u.Scheme -ne 'https' -or
            $u.Query -ne '' -or $u.Fragment -ne '') { throw 'Ungueltiger Paketpfad.' }
        return [Uri]::UnescapeDataString($u.AbsolutePath.TrimStart('/'))
    }

    function Header($cell, $strings) {
        if ($null -eq $cell) { return '[LEER]' }
        if ($null -ne $cell.SelectSingleNode("*[local-name()='f']")) { return '[FORMEL; SIEHE UNTEN]' }
        $v = $cell.SelectSingleNode("*[local-name()='v']")
        if ($cell.GetAttribute('t') -eq 's') {
            if ($null -eq $v) { throw 'Shared-String-Index fehlt.' }
            $i = [int]$v.InnerText
            if ($i -lt 0 -or $i -ge $strings.Count) { throw 'Shared-String-Index ungueltig.' }
            $node = $strings[$i]
        } elseif ($cell.GetAttribute('t') -eq 'inlineStr') {
            $node = $cell.SelectSingleNode("*[local-name()='is']")
        } else {
            if ($null -eq $v) { return '[LEER]' }
            return $v.InnerText
        }
        if ($null -eq $node) { return '[LEER]' }
        return (@($node.SelectNodes(".//*[local-name()='t' and not(ancestor::*[local-name()='rPh'])]") |
            ForEach-Object { $_.InnerText }) -join '')
    }

    $jobs = @(
        @{ Label='VORLAGE'; File='Contribution Calculation IPP 2026 - Template - neu.xlsx' }
    )

    foreach ($job in $jobs) {
        $dialog = New-Object System.Windows.Forms.OpenFileDialog
        $dialog.Title = 'Vorhandene Vorlage auswaehlen - alle Arbeitsblaetter werden untersucht'
        $dialog.Filter = 'Excel-Arbeitsmappen|*.xlsm;*.xlsx'
        $dialog.FileName = $job.File
        try {
            if ($dialog.ShowDialog() -ne [Windows.Forms.DialogResult]::OK) {
                $summary.Add("$($job.Label): STATUS=SKIPPED_NOT_SELECTED")
                continue
            }
            $path = $dialog.FileName
        } finally { $dialog.Dispose() }

        $source = $zip = $null
        $before = $after = ''
        $integrity = 'NOT_VERIFIED'
        $status = 'ERROR'
        $phase = 'OPEN_SOURCE'
        $complete = $false
        $data = New-Object 'System.Collections.Generic.List[string]'
        $errorText = ''
        $inventory = New-Object 'System.Collections.Generic.List[string]'
        try {
            Write-Host "Lese ausgewaehlte Datei statisch ..." -ForegroundColor Cyan
            $source = [IO.FileStream]::new($path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
            $phase = 'READ_SOURCE'
            $before = Read-Hash $source
            $zip = [IO.Compression.ZipArchive]::new($source, [IO.Compression.ZipArchiveMode]::Read, $true)
            if (@($zip.Entries | Group-Object FullName | Where-Object { $_.Count -gt 1 }).Count -gt 0) {
                throw 'Doppelte ZIP-Bestandteile.'
            }
            $wb = Read-Xml $zip 'xl/workbook.xml'
            $sheets = @($wb.SelectNodes("/*[local-name()='workbook']/*[local-name()='sheets']/*[local-name()='sheet']"))
            if ($sheets.Count -eq 0) { throw 'Die Arbeitsmappe hat keine Blattliste.' }
            $inventory.Add("BLATTEINTRAEGE_GESAMT=$($sheets.Count)")
            $strings = @()
            if ($null -ne $zip.GetEntry('xl/sharedStrings.xml')) {
                $ss = Read-Xml $zip 'xl/sharedStrings.xml'
                $strings = @($ss.SelectNodes("/*[local-name()='sst']/*[local-name()='si']"))
            }
            foreach ($sheet in $sheets) {
                $sheetName = $sheet.GetAttribute('name')
                $phase = 'READ_SHEET: ' + $sheetName
                $visibility = $sheet.GetAttribute('state')
                if ($visibility -eq '') { $visibility = 'visible' }
                $rel = Relation $zip 'xl/workbook.xml' (Rel-Id $sheet)
                $kind = $rel.GetAttribute('Type')
                $inventory.Add("BLATT=$sheetName | SICHTBARKEIT=$visibility | TYP=$kind")
                $data.Add("`r`n=== BLATT: $sheetName | SICHTBARKEIT=$visibility ===")
                if (-not $kind.EndsWith('/worksheet')) {
                    $data.Add('STATUS=NUR_INVENTAR; kein regulaeres XML-Arbeitsblatt, nicht als Zellblatt untersucht.')
                    continue
                }
                Write-Host "Lese Blatt: $sheetName" -ForegroundColor Cyan
                $part = Package-Part 'xl/workbook.xml' $rel
                $xml = Read-Xml $zip $part
                $dim = $xml.DocumentElement.SelectSingleNode("*[local-name()='dimension']")
                if ($null -ne $dim) {
                    $data.Add('BEREICH_XML=' + $dim.GetAttribute('ref'))
                }

                # Tabellenmetadaten helfen auch bei Kopfzeilen unterhalb Zeile 1.
                foreach ($tp in $xml.DocumentElement.SelectNodes("*[local-name()='tableParts']/*[local-name()='tablePart']")) {
                    $tr = Relation $zip $part (Rel-Id $tp)
                    $tablePart = Package-Part $part $tr
                    $td = Read-Xml $zip $tablePart
                    $table = $td.DocumentElement
                    if ($table.LocalName -ne 'table') { throw 'Unerwarteter Tabellenbestandteil.' }
                    $tableName = $table.GetAttribute('name')
                    $tableRef = $table.GetAttribute('ref')
                    $inventory.Add("TABELLE=$tableName | BLATT=$sheetName | BEREICH=$tableRef")
                    $data.Add("TABELLE=$tableName | BEREICH=$tableRef | HEADER_ROWS_XML=$($table.GetAttribute('headerRowCount'))")
                    $position = 0
                    foreach ($tc in $table.SelectNodes("*[local-name()='tableColumns']/*[local-name()='tableColumn']")) {
                        $position++
                        $data.Add("TABELLENSPALTE=$position | NAME=$($tc.GetAttribute('name'))")
                        foreach ($tf in $tc.SelectNodes("*[local-name()='calculatedColumnFormula' or local-name()='totalsRowFormula']")) {
                            $data.Add("TABELLENFORMEL=$($tf.OuterXml)")
                        }
                    }
                }
                $cells = @($xml.SelectNodes("//*[local-name()='sheetData']/*[local-name()='row']/*[local-name()='c'][*[local-name()='f' or local-name()='v' or local-name()='is']]"))
                $shared = @{}
                foreach ($c in $cells) {
                    $f = $c.SelectSingleNode("*[local-name()='f']")
                    if ($null -ne $f -and $f.GetAttribute('t') -eq 'shared' -and $f.InnerText.Length -gt 0) {
                        $shared[$f.GetAttribute('si')] = $c.GetAttribute('r') + ' | ' + $f.InnerText
                    }
                }
                $groups = $cells | Group-Object { $_.GetAttribute('r') -replace '\d+', '' } |
                    Sort-Object @{Expression={ $_.Name.Length }}, Name
                foreach ($g in $groups) {
                    $h = $g.Group | Where-Object { $_.GetAttribute('r') -eq "$($g.Name)1" } | Select-Object -First 1
                    $data.Add("`r`nSPALTE=$($g.Name) | HEADER_KANDIDAT_ZEILE_1=$(Header $h $strings)")
                    $fc = @($g.Group | Where-Object { $null -ne $_.SelectSingleNode("*[local-name()='f']") } |
                        Sort-Object { [int]($_.GetAttribute('r') -replace '\D+', '') })
                    $data.Add("FORMELZELLEN_GESAMT=$($fc.Count)")
                    $seen = @{}
                    foreach ($c in (@($fc | Select-Object -First 2) + @($fc | Select-Object -Last 1))) {
                        $a = $c.GetAttribute('r')
                        if ($seen.ContainsKey($a)) { continue }
                        $seen[$a] = $true
                        $f = $c.SelectSingleNode("*[local-name()='f']")
                        $data.Add("ZELLE=$a | FORMEL_XML=$($f.OuterXml)")
                        if ($f.GetAttribute('t') -eq 'shared') {
                            $si = $f.GetAttribute('si')
                            if (-not $shared.ContainsKey($si)) { throw "Shared-Formelbasis fehlt: $a" }
                            $data.Add("SHARED_ID=$si | BASIS=$($shared[$si]) | NICHT AUF ZIELZELLE UMGERECHNET")
                        }
                    }
                }
            }
            # Indizes gelten nur in DIESER ausgewaehlten Quelldatei.
            $refs = @($wb.SelectNodes("/*[local-name()='workbook']/*[local-name()='externalReferences']/*[local-name()='externalReference']"))
            $data.Add("`r`nEXTERNE_REFERENZEN_DIESER_DATEI=$($refs.Count)")
            for ($i = 0; $i -lt $refs.Count; $i++) {
                $er = Relation $zip 'xl/workbook.xml' (Rel-Id ($refs[$i]))
                $ep = Package-Part 'xl/workbook.xml' $er
                $ed = Read-Xml $zip $ep
                $eb = $ed.DocumentElement.SelectSingleNode("*[local-name()='externalBook']")
                if ($null -eq $eb) { $data.Add("EXTERN=[$($i+1)] | KEIN_EXTERNALBOOK"); continue }
                $tr = Relation $zip $ep (Rel-Id $eb)
                $target = $tr.GetAttribute('Target')
                if ($target -match '^https?://') { $target = ([Uri]$target).AbsolutePath }
                $name = ([Uri]::UnescapeDataString($target).Replace('\', '/') -split '/')[-1]
                $data.Add("EXTERN=[$($i+1)] | DATEINAME=$name")
            }
            $calc = $wb.DocumentElement.SelectSingleNode("*[local-name()='calcPr']")
            if ($null -ne $calc) { $data.Add("CALCPR_XML=$($calc.OuterXml)") }
            $data.Add('STATUS_DER_ANALYSE=FORMELSTICHPROBE; keine Berechnung und keine Freigabe als LNW-Ersatz.')
            $complete = $true
        } catch {
            $ex = $_.Exception.GetBaseException()
            $errorText = $ex.Message
            if ($phase -eq 'OPEN_SOURCE' -and ($ex.HResult -band 65535) -in @(32,33)) { $status = 'BLOCKED_FILE_LOCK' }
        } finally {
            if ($null -ne $zip) {
                try { $zip.Dispose() } catch { $complete = $false; $errorText += ' | ZIP_CLOSE_ERROR' }
            }
            if ($null -ne $source) {
                try {
                    if ($before -cmatch '\A[0-9A-F]{64}\z') {
                        $after = Read-Hash $source
                        if ($after -ceq $before) { $integrity = 'PASS' } else { $integrity = 'FAIL' }
                    }
                } catch { $errorText += ' | HASH_ERROR: ' + $_.Exception.Message }
                finally { $source.Dispose() }
            }
        }
        if ($complete -and $integrity -eq 'PASS') { $status = 'EXPORTED_SAMPLE' }
        elseif ($complete) { $status = 'INTEGRITY_NOT_CONFIRMED' }
        $summary.Add("`r`n=== $($job.Label) ===")
        $summary.Add('DATEI=' + [IO.Path]::GetFileName($path))
        $summary.Add("STATUS=$status | QUELLE_UNVERAENDERT=$integrity")
        $summary.Add("SHA256_VORHER=$(if ($before) { $before } else { 'NOT_AVAILABLE' })")
        $summary.Add("SHA256_NACHHER=$(if ($after) { $after } else { 'NOT_AVAILABLE' })")
        if ($errorText) { $summary.Add("FEHLER=$errorText") }
        if ($status -ne 'EXPORTED_SAMPLE') { continue }
        foreach ($item in $inventory) { $summary.Add($item) }
        $chunks = New-Object 'System.Collections.Generic.List[string]'
        $buffer = New-Object Text.StringBuilder
        foreach ($line in $data) {
            if ($buffer.Length -gt 0 -and $buffer.Length + $line.Length -gt 12000) {
                $chunks.Add($buffer.ToString()); [void]$buffer.Clear()
            }
            [void]$buffer.AppendLine($line)
        }
        if ($buffer.Length -gt 0) { $chunks.Add($buffer.ToString()) }
        for ($i = 0; $i -lt $chunks.Count; $i++) {
            $name = '{0}_TEIL_{1:D3}.txt' -f $job.Label, ($i+1)
            $head = "MCR_SOURCE_AUDIT=1.1 | $($job.Label) | TEIL=$($i+1)/$($chunks.Count)`r`nSHA256=$before`r`n"
            [IO.File]::WriteAllText((Join-Path $out $name), $head + $chunks[$i], $utf8)
            $summary.Add("BERICHT=$name")
        }
    }
    $report = Join-Path $out '00_START_HIER.txt'
    [IO.File]::WriteAllLines($report, $summary.ToArray(), $utf8)
    Write-Host "Berichte: $out" -ForegroundColor Cyan
    Start-Process notepad.exe -ArgumentList ('"' + $report + '"')
    Start-Process explorer.exe -ArgumentList ('"' + $out + '"')
}
