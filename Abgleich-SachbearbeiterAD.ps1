#Requires -Version 5.1
<#
.SYNOPSIS
    OFFLINE-Abgleich: ordnet den Sachbearbeitern aus einer Excel-Liste die
    passenden AD-Benutzernamen (sAMAccountName) aus einem AD-Export zu und
    erzeugt EINE Ergebnisliste zum manuellen Ergaenzen/Weiterverarbeiten.

.DESCRIPTION
    KEINE Verbindung zum Active Directory noetig (kein RSAT). Es werden nur
    zwei Dateien verglichen:

      1) -ExcelPath   : Sachbearbeiter-Liste (Quelle der Wahrheit), Spalten
                        "Vorname" und "Nachname".
      2) -AdListPath  : Export ALLER (relevanten) AD-Benutzer mit Vorname,
                        Nachname und sAMAccountName (CSV oder XLSX).

    Das Skript matcht ueber givenName + sn (mit Varianten fuer Umlaute,
    "genannt"-Namen und mehrteilige Vornamen) und schreibt eine Ergebnisliste:

      Vorname | Nachname | Status | sAMAccountName | Kandidaten | Hinweis

    Status:
      GEFUNDEN      -> genau 1 Treffer, sAMAccountName ist eingetragen
      MEHRDEUTIG    -> mehrere Treffer, bitte Kandidaten pruefen (Spalte leer)
      NICHT_GEFUNDEN-> kein Treffer, bitte manuell ergaenzen (Spalte leer)

    Die Zeilen mit leerem sAMAccountName sind genau die, die du noch von Hand
    fuellen musst.

.PARAMETER ExcelPath
    Pfad zur Sachbearbeiter-Excel (.xlsx).

.PARAMETER AdListPath
    Pfad zum AD-Export (.csv oder .xlsx) mit Vorname/Nachname/sAMAccountName.

.PARAMETER WorksheetName
    Optionales Arbeitsblatt in der Sachbearbeiter-Excel (bei mehreren Blaettern).

.PARAMETER OutputPath
    Zielpfad der Ergebnisliste (.csv). Standard: neben der Excel mit Zeitstempel.
    Zusaetzlich wird - falls moeglich - eine gleichnamige .xlsx geschrieben.

.PARAMETER AdGivenNameColumn / AdSurnameColumn / AdSamColumn / AdDisplayNameColumn
    Optional: Spaltennamen im AD-Export ueberschreiben, falls die automatische
    Erkennung nicht passt.

.EXAMPLE
    .\Abgleich-SachbearbeiterAD.ps1 `
        -ExcelPath  .\Source\Liste_SachbearbeiterLeistung.xlsx `
        -AdListPath .\Source\AD_Export.csv

.EXAMPLE
    # Mit Beispieldateien zum Ausprobieren:
    .\Abgleich-SachbearbeiterAD.ps1 `
        -ExcelPath  .\Source\Beispiel_SachbearbeiterListe.xlsx -WorksheetName "Leistung" `
        -AdListPath .\Source\Beispiel_AD_Export.csv

.NOTES
    * Excel lesen: Modul "ImportExcel" (empfohlen) ODER Microsoft Excel (COM).
      CSV wird ohne Zusatzmodul gelesen.
    * Das Skript aendert NICHTS - es erzeugt nur die Ergebnisliste.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ExcelPath,

    [Parameter(Mandatory = $true)]
    [string]$AdListPath,

    [Parameter(Mandatory = $false)]
    [string]$WorksheetName,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath,

    [Parameter(Mandatory = $false)]
    [string]$AdGivenNameColumn,

    [Parameter(Mandatory = $false)]
    [string]$AdSurnameColumn,

    [Parameter(Mandatory = $false)]
    [string]$AdSamColumn,

    [Parameter(Mandatory = $false)]
    [string]$AdDisplayNameColumn
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

# ==========================================================================
#  HILFSFUNKTIONEN
# ==========================================================================

# Namens-Normalisierung: Umlaute/Diakritika/Sonderzeichen vergleichbar machen
#   "Grünberg" -> "gruenberg", "Marie-Christin" -> "marie christin"
function ConvertTo-NormalizedName {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return '' }
    $n = $Name.Trim().ToLowerInvariant()
    $n = $n -replace 'ä', 'ae' -replace 'ö', 'oe' -replace 'ü', 'ue' -replace 'ß', 'ss'
    $sb = [System.Text.StringBuilder]::new()
    foreach ($ch in $n.Normalize([Text.NormalizationForm]::FormD).ToCharArray()) {
        if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch) -ne
            [Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$sb.Append($ch)
        }
    }
    $n = $sb.ToString()
    $n = $n -replace '[^a-z0-9]', ' '
    $n = ($n -replace '\s+', ' ').Trim()
    return $n
}

# Nachnamen-Varianten (behandelt "genannt")
#   "Böhme genannt Klein" -> @("boehme genannt klein","boehme","klein")
function Get-SurnameVariants {
    param([string]$Surname)
    $variants = New-Object System.Collections.Generic.List[string]
    $full = ConvertTo-NormalizedName $Surname
    if ($full) { $variants.Add($full) }
    if ($Surname -match '(?i)\bgenannt\b') {
        foreach ($p in ($Surname -split '(?i)\bgenannt\b')) {
            $v = ConvertTo-NormalizedName $p
            if ($v -and -not $variants.Contains($v)) { $variants.Add($v) }
        }
    }
    return $variants
}

# Vornamen-Varianten (voll + erstes Token)
#   "Leslie Jenny" -> @("leslie jenny","leslie")
function Get-GivenNameVariants {
    param([string]$GivenName)
    $variants = New-Object System.Collections.Generic.List[string]
    $full = ConvertTo-NormalizedName $GivenName
    if ($full) { $variants.Add($full) }
    $first = ($full -split ' ')[0]
    if ($first -and -not $variants.Contains($first)) { $variants.Add($first) }
    return $variants
}

# Wert einer Spalte tolerant beziehen (Alias-Liste, Gross/Klein + Whitespace egal)
function Get-ColumnValue {
    param($Row, [string[]]$Names)
    foreach ($p in $Row.PSObject.Properties) {
        $pn = ($p.Name).Trim()
        foreach ($n in $Names) {
            if ($pn -ieq $n) { return "$($p.Value)".Trim() }
        }
    }
    return ''
}

# Ersten passenden vorhandenen Spaltennamen finden (fuer Auto-Erkennung)
function Find-ColumnName {
    param($Row, [string[]]$Candidates)
    foreach ($p in $Row.PSObject.Properties) {
        $pn = ($p.Name).Trim()
        foreach ($c in $Candidates) {
            if ($pn -ieq $c) { return $p.Name }
        }
    }
    return $null
}

# --------------------------------------------------------------------------
# Excel-Arbeitsblatt lesen (ImportExcel bevorzugt, sonst Excel-COM)
# --------------------------------------------------------------------------
function Import-XlsxRows {
    param([string]$Path, [string]$WorksheetName)

    if (Get-Module -ListAvailable -Name ImportExcel) {
        Import-Module ImportExcel -ErrorAction Stop
        if ($WorksheetName) { return Import-Excel -Path $Path -WorksheetName $WorksheetName }
        return Import-Excel -Path $Path
    }

    # COM-Fallback (Excel muss installiert sein)
    $excel = $null; $wb = $null; $ws = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false; $excel.DisplayAlerts = $false
        $wb = $excel.Workbooks.Open((Resolve-Path $Path).Path, $null, $true)
        $ws = if ($WorksheetName) { $wb.Worksheets.Item($WorksheetName) } else { $wb.Worksheets.Item(1) }
        $used = $ws.UsedRange
        $data = $used.Value2
        $rowCount = $used.Rows.Count; $colCount = $used.Columns.Count
        $headers = @{}
        for ($c = 1; $c -le $colCount; $c++) {
            $h = "$($data[1, $c])".Trim()
            if ($h) { $headers[$h] = $c }
        }
        $result = New-Object System.Collections.Generic.List[object]
        for ($rr = 2; $rr -le $rowCount; $rr++) {
            $obj = [ordered]@{}
            foreach ($h in $headers.Keys) { $obj[$h] = "$($data[$rr, $headers[$h]])" }
            $result.Add([pscustomobject]$obj)
        }
        return $result
    }
    finally {
        if ($wb)    { $wb.Close($false) | Out-Null }
        if ($excel) { $excel.Quit() | Out-Null }
        foreach ($o in @($ws, $wb, $excel)) {
            if ($o) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) }
        }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}

# --------------------------------------------------------------------------
# AD-Datei lesen (CSV mit Auto-Trennzeichen, oder XLSX)
# --------------------------------------------------------------------------
function Import-AdRows {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { throw "AD-Datei nicht gefunden: $Path" }

    $ext = [IO.Path]::GetExtension($Path).ToLowerInvariant()
    if ($ext -eq '.csv') {
        # Trennzeichen automatisch bestimmen (deutsche Exporte oft ';')
        $firstLine = (Get-Content -LiteralPath $Path -TotalCount 1 -Encoding UTF8)
        $delim = if ($firstLine -match ';') { ';' } else { ',' }
        return Import-Csv -LiteralPath $Path -Delimiter $delim -Encoding UTF8
    }
    elseif ($ext -in @('.xlsx', '.xlsm')) {
        return Import-XlsxRows -Path $Path -WorksheetName $null
    }
    else {
        throw "Nicht unterstuetztes Format der AD-Datei: '$ext' (erwartet .csv oder .xlsx)"
    }
}

# ==========================================================================
#  HAUPTPROGRAMM
# ==========================================================================
try {
    Write-Host "=== OFFLINE-Abgleich Sachbearbeiter <-> AD-Export ===" -ForegroundColor Cyan

    if (-not (Test-Path -LiteralPath $ExcelPath))  { throw "Excel-Datei nicht gefunden: $ExcelPath" }

    # --- 1) Sachbearbeiter aus Excel lesen (nur Zeilen MIT Namen) ---
    $srcRows = Import-XlsxRows -Path $ExcelPath -WorksheetName $WorksheetName
    $persons = New-Object System.Collections.Generic.List[object]
    $seen = @{}
    foreach ($r in $srcRows) {
        $vor  = Get-ColumnValue $r @('Vorname', 'GivenName', 'Rufname')
        $nach = Get-ColumnValue $r @('Nachname', 'Surname', 'Name', 'sn')
        if (-not $vor -and -not $nach) { continue }   # leere/Fortsetzungszeile
        $key = (ConvertTo-NormalizedName $vor) + '|' + (ConvertTo-NormalizedName $nach)
        if ($seen.ContainsKey($key)) { continue }
        $seen[$key] = $true
        $persons.Add([pscustomobject]@{ Vorname = $vor; Nachname = $nach })
    }
    Write-Host ("Sachbearbeiter (eindeutig) : {0}" -f $persons.Count) -ForegroundColor Gray

    # --- 2) AD-Export lesen + Spalten erkennen ---
    $adRows = @(Import-AdRows -Path $AdListPath)
    if ($adRows.Count -eq 0) { throw "AD-Datei enthaelt keine Datenzeilen." }

    $sample = $adRows[0]
    $colGiven = if ($AdGivenNameColumn) { $AdGivenNameColumn } else { Find-ColumnName $sample @('GivenName', 'Vorname', 'givenName', 'Rufname') }
    $colSn    = if ($AdSurnameColumn)   { $AdSurnameColumn }   else { Find-ColumnName $sample @('Surname', 'Nachname', 'sn', 'SN') }
    $colSam   = if ($AdSamColumn)       { $AdSamColumn }       else { Find-ColumnName $sample @('SamAccountName', 'sAMAccountName', 'Benutzername', 'Login', 'AccountName', 'UserName', 'Anmeldename') }
    $colDisp  = if ($AdDisplayNameColumn) { $AdDisplayNameColumn } else { Find-ColumnName $sample @('DisplayName', 'Anzeigename', 'CN', 'Name') }

    if (-not $colGiven -or -not $colSn) {
        throw ("Im AD-Export fehlen Vorname-/Nachname-Spalten. Gefundene Spalten: {0}. " +
               "Bitte -AdGivenNameColumn / -AdSurnameColumn angeben." -f
               (($sample.PSObject.Properties | ForEach-Object { $_.Name }) -join ', '))
    }
    if (-not $colSam) {
        throw ("Im AD-Export fehlt die sAMAccountName-Spalte. Gefundene Spalten: {0}. " +
               "Bitte -AdSamColumn angeben." -f
               (($sample.PSObject.Properties | ForEach-Object { $_.Name }) -join ', '))
    }
    Write-Host ("AD-Benutzer im Export      : {0}" -f $adRows.Count) -ForegroundColor Gray
    Write-Host ("Erkannte AD-Spalten        : Vorname='{0}', Nachname='{1}', sAMAccountName='{2}', Anzeige='{3}'" -f
                $colGiven, $colSn, $colSam, ($(if ($colDisp) { $colDisp } else { '-' }))) -ForegroundColor Gray

    # --- 3) AD-Index aufbauen (normalisiert) ---
    $byGivenSn = @{}   # "given|sn" -> Liste von AD-Zeilen
    $byDisplay = @{}   # normalisierter Anzeigename -> Liste
    foreach ($a in $adRows) {
        $g  = ConvertTo-NormalizedName ("$($a.$colGiven)")
        $s  = ConvertTo-NormalizedName ("$($a.$colSn)")
        $sam = "$($a.$colSam)".Trim()
        if (-not $sam) { continue }
        $entry = [pscustomobject]@{ Sam = $sam; Given = "$($a.$colGiven)"; Sn = "$($a.$colSn)" }

        if ($g -and $s) {
            $k = "$g|$s"
            if (-not $byGivenSn.ContainsKey($k)) { $byGivenSn[$k] = New-Object System.Collections.Generic.List[object] }
            if (-not ($byGivenSn[$k] | Where-Object { $_.Sam -ieq $sam })) { $byGivenSn[$k].Add($entry) }
        }
        if ($colDisp) {
            $d = ConvertTo-NormalizedName ("$($a.$colDisp)")
            if ($d) {
                if (-not $byDisplay.ContainsKey($d)) { $byDisplay[$d] = New-Object System.Collections.Generic.List[object] }
                if (-not ($byDisplay[$d] | Where-Object { $_.Sam -ieq $sam })) { $byDisplay[$d].Add($entry) }
            }
        }
    }

    # --- 4) Abgleich je Person ---
    $results = New-Object System.Collections.Generic.List[object]
    $countFound = 0; $countAmbig = 0; $countMiss = 0

    foreach ($p in $persons) {
        $cands = New-Object System.Collections.Generic.List[object]

        # Primaer: givenName + sn (alle Varianten kombinieren)
        foreach ($g in (Get-GivenNameVariants $p.Vorname)) {
            foreach ($s in (Get-SurnameVariants $p.Nachname)) {
                $k = "$g|$s"
                if ($byGivenSn.ContainsKey($k)) {
                    foreach ($e in $byGivenSn[$k]) {
                        if (-not ($cands | Where-Object { $_.Sam -ieq $e.Sam })) { $cands.Add($e) }
                    }
                }
            }
        }
        # Fallback: Anzeigename == "Vorname Nachname"
        if ($cands.Count -eq 0 -and $byDisplay.Count -gt 0) {
            foreach ($dk in @((ConvertTo-NormalizedName "$($p.Vorname) $($p.Nachname)"),
                              (ConvertTo-NormalizedName "$($p.Nachname) $($p.Vorname)"))) {
                if ($byDisplay.ContainsKey($dk)) {
                    foreach ($e in $byDisplay[$dk]) {
                        if (-not ($cands | Where-Object { $_.Sam -ieq $e.Sam })) { $cands.Add($e) }
                    }
                }
            }
        }

        if ($cands.Count -eq 1) {
            $status = 'GEFUNDEN'; $sam = $cands[0].Sam; $kand = ''; $hint = ''
            $countFound++
        }
        elseif ($cands.Count -gt 1) {
            $status = 'MEHRDEUTIG'; $sam = ''; $countAmbig++
            $kand = ($cands | ForEach-Object { $_.Sam }) -join ' | '
            $hint = 'Bitte richtigen Benutzer waehlen und sAMAccountName eintragen.'
        }
        else {
            $status = 'NICHT_GEFUNDEN'; $sam = ''; $kand = ''; $countMiss++
            $hint = 'Im AD-Export nicht gefunden - bitte manuell ergaenzen.'
        }

        $results.Add([pscustomobject]@{
            Vorname        = $p.Vorname
            Nachname       = $p.Nachname
            Status         = $status
            sAMAccountName = $sam
            Kandidaten     = $kand
            Hinweis        = $hint
        })
    }

    # Sortierung: offene Faelle (zu ergaenzen) zuerst, dann gefundene
    $order = @{ 'NICHT_GEFUNDEN' = 0; 'MEHRDEUTIG' = 1; 'GEFUNDEN' = 2 }
    $sorted = $results | Sort-Object @{ Expression = { $order[$_.Status] } }, Nachname, Vorname

    # --- 5) Ergebnis schreiben (CSV, plus XLSX falls moeglich) ---
    if (-not $OutputPath) {
        $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
        $dir   = Split-Path -Parent (Resolve-Path $ExcelPath)
        $OutputPath = Join-Path $dir "Abgleich_Ergebnis_$stamp.csv"
    }
    $sorted | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding UTF8 -Delimiter ';'
    Write-Host ("`nErgebnisliste (CSV)        : {0}" -f $OutputPath) -ForegroundColor Green

    if (Get-Module -ListAvailable -Name ImportExcel) {
        try {
            $xlsxOut = [IO.Path]::ChangeExtension($OutputPath, '.xlsx')
            $sorted | Export-Excel -Path $xlsxOut -AutoSize -FreezeTopRow -BoldTopRow -WorksheetName 'Abgleich'
            Write-Host ("Ergebnisliste (XLSX)       : {0}" -f $xlsxOut) -ForegroundColor Green
        } catch {
            Write-Host ("XLSX konnte nicht geschrieben werden: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
        }
    }

    # --- 6) Zusammenfassung ---
    Write-Host "`n==================== ZUSAMMENFASSUNG ====================" -ForegroundColor Cyan
    Write-Host ("Sachbearbeiter gesamt : {0}" -f $persons.Count)
    Write-Host ("Gefunden              : {0}" -f $countFound) -ForegroundColor Green
    Write-Host ("Mehrdeutig            : {0}" -f $countAmbig) -ForegroundColor Yellow
    Write-Host ("Nicht gefunden        : {0}" -f $countMiss)  -ForegroundColor Yellow
    Write-Host "========================================================="
    if (($countAmbig + $countMiss) -gt 0) {
        Write-Host ("HINWEIS: {0} Zeile(n) haben einen leeren sAMAccountName und muessen manuell ergaenzt werden." -f ($countAmbig + $countMiss)) -ForegroundColor Yellow
    }
}
catch {
    Write-Host ("ABBRUCH: {0}" -f $_.Exception.Message) -ForegroundColor Red
    exit 1
}
