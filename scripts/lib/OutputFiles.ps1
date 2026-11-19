function Open-StaDocument($Session, [string]$Collection, [string]$Path, [switch]$Writable) {
    $documents = Invoke-StaCall 'Open' "Application.$Collection" { ,$Session.App.$Collection }
    try {
        if ($Writable) {
            Invoke-StaCall 'Open' "$Collection.Open" { ,($documents.Open($Path, $false)) }
        } else {
            Invoke-StaCall 'Open' "$Collection.OpenReadOnly" { ,($documents.OpenReadOnly($Path, $false)) }
        }
    } finally {
        Release-StaCom $documents
    }
}

function Copy-StaInput([string]$Source, [string]$Stage) {
    Invoke-StaCall 'Stage' 'File.Copy' {
        [IO.File]::Copy($Source, $Stage)
        [IO.File]::SetAttributes($Stage, [IO.FileAttributes]::Normal)
    }
}

function Assert-StaWritable($Sheet) {
    Invoke-StaCall 'Open' 'Spreadsheet.ReadOnly' {
        if ($Sheet.ReadOnly) { throw "The working spreadsheet opened read-only: $($Sheet.FullName)" }
    }
}

function Save-StaDocument($Document, [string]$Path) {
    Invoke-StaCall 'Save' 'Document.SaveAs' {
        # The stage file is unique; the parent applies the requested overwrite policy on publish.
        if (-not $Document.SaveAs($Path, $true)) { throw "SaveAs returned false for $Path." }
    }
}

function Save-StaSpreadsheet($Sheet) {
    Invoke-StaCall 'Save' 'Spreadsheet.Save' {
        if ($Sheet.ReadOnly) { throw "The working spreadsheet is read-only: $($Sheet.FullName)" }
        $Sheet.Save()
        if (-not $Sheet.Saved) { throw 'STATISTICA still reports unsaved changes after Save().' }
    }
}

function Confirm-StaFile([string]$Path) {
    Invoke-StaCall 'Verify' 'Output file' {
        $file = [IO.FileInfo]::new($Path)
        if (-not $file.Exists -or $file.Length -eq 0) { throw "Expected a nonempty output file: $Path" }
        switch ($file.Extension) {
            '.png' {
                $stream = $file.OpenRead()
                try {
                    $head = [byte[]]::new(24)
                    $read = $stream.Read($head, 0, $head.Length)
                    if ($read -ne 24 -or [BitConverter]::ToString($head, 0, 8) -ne '89-50-4E-47-0D-0A-1A-0A' -or
                        [Text.Encoding]::ASCII.GetString($head, 12, 4) -ne 'IHDR' -or
                        [Net.IPAddress]::NetworkToHostOrder([BitConverter]::ToInt32($head, 16)) -le 0 -or
                        [Net.IPAddress]::NetworkToHostOrder([BitConverter]::ToInt32($head, 20)) -le 0) {
                        throw 'Invalid PNG header or image dimensions.'
                    }
                } finally { $stream.Dispose() }
            }
            '.svg' {
                $settings = [Xml.XmlReaderSettings]::new()
                $settings.DtdProcessing = 'Ignore'
                $settings.XmlResolver = $null
                $reader = [Xml.XmlReader]::Create($Path, $settings)
                try {
                    [void]$reader.MoveToContent()
                    if ($reader.LocalName -ne 'svg' -or $reader.NamespaceURI -ne 'http://www.w3.org/2000/svg') {
                        throw 'The output is missing its SVG root element.'
                    }
                    while ($reader.Read()) {}
                } finally { $reader.Dispose() }
            }
        }
    }
}
