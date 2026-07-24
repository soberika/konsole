#Requires -Version 5.1
<#
.SYNOPSIS
    ANALYSIERT den Soll/Ist-Abgleich einer AD-Sicherheitsgruppe
    (z. B. "Liste_Jobcenter_Leistung") gegen eine Excel-Liste und erzeugt
    daraus einen visuellen Report + ein fertiges, pruefbares Ausfuehrungsskript.

.DESCRIPTION
    Zwei-Phasen-Prinzip (bewusst getrennt, damit NIE versehentlich etwas passiert):

      PHASE 1 - ANALYSE (Standard, NUR LESEND):
        * liest die Excel-Liste (einzige Quelle der Wahrheit),
        * matcht jede Person gegen das AD (primaer givenName + sn, mit Fallbacks),
        * vergleicht mit den *direkten* Benutzer-Mitgliedern der Gruppe,
        * schreibt einen HTML-Report (Vorher/Nachher) und eine CSV,
        * erzeugt ein GENERIERTES Ausfuehrungsskript "Apply-*.ps1" mit den
          konkreten Add-/Remove-Befehlen (aufgeloeste DNs, je 1 Zeile pro Person).
        => Es wird NICHTS im AD veraendert.

      PHASE 2 - FREIGABE & AUSFUEHRUNG (durch den Menschen):
        * Report ansehen, Apply-Skript pruefen ("absegnen"),
        * Apply-Skript starten -> DANN erst werden die Namen ergaenzt/entfernt.
          (Das Apply-Skript fragt vor dem Schreiben noch einmal nach und
           unterstuetzt selbst -WhatIf.)

    Verschachtelte Gruppen (nested groups) werden NIE angefasst - nur direkte
    Benutzer-Mitglieder. Mehrdeutige/nicht gefundene Personen werden NICHT
    automatisiert, sondern nur zur manuellen Pruefung ausgewiesen.

.PARAMETER ExcelPath
    Pfad zur .xlsx-Datei mit den Sachbearbeitern.

.PARAMETER GroupName
    Name der Ziel-AD-Gruppe. Standard: "Liste_Jobcenter_Leistung".

.PARAMETER WorksheetName
    Name des zu verwendenden Arbeitsblatts (Tabellenblatt) in der xlsx.
    Wird der Parameter NICHT angegeben und die Datei hat mehrere Blaetter,
    fragt das Skript interaktiv per Menue nach. Fuer unbeaufsichtigte Laeufe
    (Aufgabenplaner) diesen Parameter immer setzen.

.PARAMETER Server
    Optionaler Domain Controller / Domaenenname (z. B. "kreis-meissen.de").

.PARAMETER OutputDir
    Zielordner fuer Report/CSV/Apply-Skript/Log. Standard: Ordner "Sync-Output"
    neben diesem Skript.

.PARAMETER LogPath
    Optionaler expliziter Pfad zur Logdatei.

.PARAMETER HtmlReportPath
    Optionaler expliziter Pfad fuer den HTML-Report.

.PARAMETER ApplyScriptPath
    Optionaler expliziter Pfad fuer das generierte Ausfuehrungsskript.

.PARAMETER CsvPath
    Optionaler expliziter Pfad fuer den CSV-Export.

.EXAMPLE
    # Analyse (Standard) - liest nur, erzeugt Report + Apply-Skript:
    .\Sync-JobcenterGroup.ps1 -ExcelPath .\Source\Liste_SachbearbeiterLeistung.xlsx

.EXAMPLE
    # Danach das generierte Skript pruefen und ausfuehren:
    .\Sync-Output\Apply-Liste_Jobcenter_Leistung_20260724.ps1 -WhatIf   # nochmal Trockenlauf
    .\Sync-Output\Apply-Liste_Jobcenter_Leistung_20260724.ps1           # echte Aenderung (fragt nach)

.EXAMPLE
    # Bestimmtes Arbeitsblatt direkt waehlen (kein interaktives Menue):
    .\Sync-JobcenterGroup.ps1 -ExcelPath .\Source\Beispiel_SachbearbeiterListe.xlsx `
        -WorksheetName "Leistung"

.EXAMPLE
    # Andere Jobcenter-Gruppe (leicht erweiterbar):
    .\Sync-JobcenterGroup.ps1 -ExcelPath .\Source\Liste_Vermittlung.xlsx `
        -GroupName "Liste_Jobcenter_Vermittlung"

.NOTES
    ANLEITUNG / VORAUSSETZUNGEN
    ---------------------------
    * PowerShell 5.1+ (Windows).
    * Modul "ActiveDirectory" (Teil der RSAT) muss installiert sein.
    * Zum Lesen der Excel-Datei wird EINES von beiden benoetigt:
        a) Modul "ImportExcel" (empfohlen, KEIN Excel noetig):
             Install-Module ImportExcel -Scope CurrentUser
        b) ODER lokal installiertes Microsoft Excel (COM-Fallback).
    * Ablauf: dieses Skript ausfuehren -> Report + Apply-Skript pruefen ->
      Apply-Skript starten. Nur das Apply-Skript veraendert das AD.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ExcelPath,

    [Parameter(Mandatory = $false)]
    [string]$GroupName = 'Liste_Jobcenter_Leistung',

    [Parameter(Mandatory = $false)]
    [string]$WorksheetName,

    [Parameter(Mandatory = $false)]
    [string]$Server,

    [Parameter(Mandatory = $false)]
    [string]$OutputDir,

    [Parameter(Mandatory = $false)]
    [string]$LogPath,

    [Parameter(Mandatory = $false)]
    [string]$HtmlReportPath,

    [Parameter(Mandatory = $false)]
    [string]$ApplyScriptPath,

    [Parameter(Mandatory = $false)]
    [string]$CsvPath
)

# --------------------------------------------------------------------------
# 0) Grundgeruest: Fehler sollen hart stoppen, damit wir sie sauber fangen
# --------------------------------------------------------------------------
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

# Ausgabeordner + Standardpfade (Report/CSV/Apply-Skript/Log) festlegen
$stamp   = Get-Date -Format 'yyyyMMdd_HHmmss'
$baseDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
if (-not $OutputDir) { $OutputDir = Join-Path $baseDir 'Sync-Output' }
if (-not (Test-Path -LiteralPath $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
}
# Fuer Dateinamen unzulaessige Zeichen der Gruppe entschaerfen
$safeGroup = ($GroupName -replace '[^\w\-]', '_')

if (-not $LogPath)         { $LogPath         = Join-Path $OutputDir "Sync_${safeGroup}_$stamp.log" }
if (-not $HtmlReportPath)  { $HtmlReportPath  = Join-Path $OutputDir "Report_${safeGroup}_$stamp.html" }
if (-not $CsvPath)         { $CsvPath         = Join-Path $OutputDir "Analyse_${safeGroup}_$stamp.csv" }
if (-not $ApplyScriptPath) { $ApplyScriptPath = Join-Path $OutputDir "Apply_${safeGroup}_$stamp.ps1" }

# --------------------------------------------------------------------------
# Logging-Helfer: schreibt gleichzeitig in Konsole (farbig) und Logdatei
# --------------------------------------------------------------------------
function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO', 'OK', 'WARN', 'ERROR', 'ADD', 'REMOVE')]
        [string]$Level = 'INFO'
    )
    $ts   = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[{0}] [{1,-6}] {2}" -f $ts, $Level, $Message

    # Farbcodierung fuer die Konsolenausgabe
    $color = switch ($Level) {
        'OK'     { 'Green' }
        'WARN'   { 'Yellow' }
        'ERROR'  { 'Red' }
        'ADD'    { 'Cyan' }
        'REMOVE' { 'Magenta' }
        default  { 'Gray' }
    }
    Write-Host $line -ForegroundColor $color

    # Bewusst tolerant: Logging darf den Sync-Lauf nie zum Absturz bringen
    try { Add-Content -Path $LogPath -Value $line -Encoding UTF8 } catch { }
}

# --------------------------------------------------------------------------
# Namens-Normalisierung: macht Umlaute/Diakritika/Sonderzeichen vergleichbar
#   "Grünberg"  -> "gruenberg"
#   "Jendroßek" -> "jendrossek"
#   "Marie-Christin" -> "marie christin"
# --------------------------------------------------------------------------
function ConvertTo-NormalizedName {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return '' }

    $n = $Name.Trim().ToLowerInvariant()

    # Deutsche Umlaute / ss vor der generischen Diakritika-Entfernung aufloesen,
    # weil "ue" != "u". Reihenfolge ist wichtig.
    $n = $n -replace 'ä', 'ae' -replace 'ö', 'oe' -replace 'ü', 'ue' -replace 'ß', 'ss'

    # Restliche Diakritika (é, è, ï, ...) ueber Unicode-Zerlegung entfernen
    $sb = [System.Text.StringBuilder]::new()
    foreach ($ch in $n.Normalize([Text.NormalizationForm]::FormD).ToCharArray()) {
        if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch) -ne
            [Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$sb.Append($ch)
        }
    }
    $n = $sb.ToString()

    # Alles ausser Buchstaben/Ziffern zu Leerzeichen, dann Mehrfach-Leerzeichen kappen
    $n = $n -replace '[^a-z0-9]', ' '
    $n = ($n -replace '\s+', ' ').Trim()
    return $n
}

# --------------------------------------------------------------------------
# Erzeugt fuer einen Nachnamen mehrere Vergleichsvarianten.
#   "Noack genannt Gräfe" -> @("noack genannt graefe", "noack", "graefe")
# So matchen wir egal, ob im AD der volle Name, der Geburts- oder Ehename steht.
# --------------------------------------------------------------------------
function Get-SurnameVariants {
    param([string]$Surname)
    $variants = New-Object System.Collections.Generic.List[string]
    $full = ConvertTo-NormalizedName $Surname
    if ($full) { $variants.Add($full) }

    if ($Surname -match '(?i)\bgenannt\b') {
        $parts = $Surname -split '(?i)\bgenannt\b'
        foreach ($p in $parts) {
            $v = ConvertTo-NormalizedName $p
            if ($v -and -not $variants.Contains($v)) { $variants.Add($v) }
        }
    }
    return $variants
}

# --------------------------------------------------------------------------
# Erzeugt fuer einen Vornamen Vergleichsvarianten.
#   "Leslie Jenny"   -> @("leslie jenny", "leslie")
#   "Marie-Christin" -> @("marie christin", "marie")
# --------------------------------------------------------------------------
function Get-GivenNameVariants {
    param([string]$GivenName)
    $variants = New-Object System.Collections.Generic.List[string]
    $full = ConvertTo-NormalizedName $GivenName
    if ($full) { $variants.Add($full) }

    $first = ($full -split ' ')[0]
    if ($first -and -not $variants.Contains($first)) { $variants.Add($first) }
    return $variants
}

# --------------------------------------------------------------------------
# Liefert die Namen aller Arbeitsblaetter der Datei (fuer beide Lese-Engines).
# --------------------------------------------------------------------------
function Get-WorksheetNames {
    param([string]$Path, [switch]$UseImportExcel)

    if ($UseImportExcel) {
        Import-Module ImportExcel -ErrorAction Stop
        return @(Get-ExcelSheetInfo -Path $Path | Select-Object -ExpandProperty Name)
    }

    # COM-Variante: Arbeitsblattnamen auslesen
    $excel = $null; $wb = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false; $excel.DisplayAlerts = $false
        $wb = $excel.Workbooks.Open((Resolve-Path $Path).Path, $null, $true)
        $names = @()
        foreach ($ws in $wb.Worksheets) { $names += [string]$ws.Name }
        return $names
    }
    finally {
        if ($wb)    { $wb.Close($false) | Out-Null }
        if ($excel) { $excel.Quit() | Out-Null }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}

# --------------------------------------------------------------------------
# Bestimmt das zu verwendende Arbeitsblatt:
#   * -WorksheetName gesetzt  -> exakt dieses (Fehler, falls nicht vorhanden)
#   * genau 1 Blatt           -> dieses
#   * mehrere Blaetter        -> interaktives Auswahlmenue (Standard: 1)
# --------------------------------------------------------------------------
function Select-Worksheet {
    param([string[]]$Sheets, [string]$Requested)

    if (-not $Sheets -or $Sheets.Count -eq 0) {
        throw "Die Excel-Datei enthaelt keine lesbaren Arbeitsblaetter."
    }

    # Explizit angefordertes Blatt (unabhaengig von Gross-/Kleinschreibung)
    if ($Requested) {
        $match = $Sheets | Where-Object { $_ -ieq $Requested } | Select-Object -First 1
        if (-not $match) {
            throw ("Arbeitsblatt '{0}' nicht gefunden. Vorhanden: {1}" -f $Requested, ($Sheets -join ', '))
        }
        return $match
    }

    if ($Sheets.Count -eq 1) { return $Sheets[0] }

    # Mehrere Blaetter, kein Parameter -> interaktiv fragen
    Write-Host ''
    Write-Host 'Die Datei enthaelt mehrere Arbeitsblaetter:' -ForegroundColor Cyan
    for ($i = 0; $i -lt $Sheets.Count; $i++) {
        Write-Host ('  [{0}] {1}' -f ($i + 1), $Sheets[$i])
    }
    $choice = Read-Host "Welches Blatt verwenden? (1-$($Sheets.Count), Enter = 1)"
    if ([string]::IsNullOrWhiteSpace($choice)) { return $Sheets[0] }

    $idx = 0
    if ([int]::TryParse($choice, [ref]$idx) -and $idx -ge 1 -and $idx -le $Sheets.Count) {
        return $Sheets[$idx - 1]
    }
    throw "Ungueltige Auswahl: '$choice'. Bitte eine Zahl zwischen 1 und $($Sheets.Count) angeben."
}

# --------------------------------------------------------------------------
# 1) Excel einlesen -> Liste von Personen (nur Zeilen MIT Vor-/Nachname).
#    Leere Namenszeilen gehoeren zum vorherigen MA (weitere Zustaendigkeits-
#    bereiche) und stellen KEINE neue Person dar -> werden uebersprungen.
# --------------------------------------------------------------------------
function Import-SachbearbeiterListe {
    param([string]$Path, [string]$WorksheetName)

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Excel-Datei nicht gefunden: $Path"
    }

    $rows = $null
    $useImportExcel = [bool](Get-Module -ListAvailable -Name ImportExcel)

    # 1) Verfuegbare Arbeitsblaetter ermitteln und passendes Blatt bestimmen
    $sheets = Get-WorksheetNames -Path $Path -UseImportExcel:$useImportExcel
    $sheet  = Select-Worksheet -Sheets $sheets -Requested $WorksheetName
    Write-Log "Verwende Arbeitsblatt: '$sheet' (von $($sheets.Count) verfuegbaren)." 'INFO'

    # 2) Zeilen aus dem gewaehlten Blatt lesen
    if ($useImportExcel) {
        Import-Module ImportExcel -ErrorAction Stop
        Write-Log "Lese Excel via Modul 'ImportExcel'." 'INFO'
        $rows = Import-Excel -Path $Path -WorksheetName $sheet
    }
    else {
        # Fallback: Excel-COM. Nur nutzbar, wenn Excel lokal installiert ist.
        Write-Log "Modul 'ImportExcel' nicht gefunden -> versuche Excel-COM-Fallback." 'WARN'
        $rows = Import-ExcelViaCom -Path $Path -WorksheetName $sheet
    }

    # Tolerant eine Spalte lesen (Header-Schreibweise/Whitespace kann variieren);
    # verhindert zugleich StrictMode-Fehler bei fehlender Property.
    function Get-Col {
        param($row, [string[]]$Names)
        foreach ($p in $row.PSObject.Properties) {
            $pn = ($p.Name).Trim()
            foreach ($n in $Names) {
                if ($pn -ieq $n) { return "$($p.Value)" }
            }
        }
        return ''
    }

    # Personen extrahieren + Duplikate (gleicher Name mehrfach) einmalig halten
    $persons = New-Object System.Collections.Generic.List[object]
    $seen    = @{}

    foreach ($r in $rows) {
        # Spaltennamen tolerant beziehen (Header-Schreibweise kann variieren)
        $vor  = (Get-Col $r @('Vorname'))  .Trim()
        $nach = (Get-Col $r @('Nachname')) .Trim()
        if (-not $vor -and -not $nach) { continue }   # leere Namenszeile -> skip

        $key = (ConvertTo-NormalizedName $vor) + '|' + (ConvertTo-NormalizedName $nach)
        if ($seen.ContainsKey($key)) { continue }
        $seen[$key] = $true

        $persons.Add([pscustomobject]@{
            Vorname  = $vor
            Nachname = $nach
            Anzeige  = "$vor $nach"
        })
    }

    return $persons
}

# --------------------------------------------------------------------------
# COM-Fallback zum Lesen der ersten Tabelle (Header in Zeile 1).
# --------------------------------------------------------------------------
function Import-ExcelViaCom {
    param([string]$Path, [string]$WorksheetName)

    # Alle in 'finally' referenzierten Variablen vorab initialisieren (StrictMode)
    $excel = $null; $wb = $null; $ws = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $wb = $excel.Workbooks.Open((Resolve-Path $Path).Path, $null, $true) # ReadOnly
        # Gewuenschtes Blatt oder ersatzweise das erste
        $ws = if ($WorksheetName) { $wb.Worksheets.Item($WorksheetName) } else { $wb.Worksheets.Item(1) }
        $used = $ws.UsedRange
        $data = $used.Value2   # 2D-Array [row, col]

        $rowCount = $used.Rows.Count
        $colCount = $used.Columns.Count

        # Header aus Zeile 1 lesen
        $headers = @{}
        for ($c = 1; $c -le $colCount; $c++) {
            $h = "$($data[1, $c])".Trim()
            if ($h) { $headers[$h] = $c }
        }

        $result = New-Object System.Collections.Generic.List[object]
        for ($rr = 2; $rr -le $rowCount; $rr++) {
            $obj = [ordered]@{}
            foreach ($h in $headers.Keys) {
                $obj[$h] = "$($data[$rr, $headers[$h]])"
            }
            $result.Add([pscustomobject]$obj)
        }
        return $result
    }
    finally {
        if ($wb)    { $wb.Close($false) | Out-Null }
        if ($excel) { $excel.Quit() | Out-Null }
        # COM-Objekte sauber freigeben
        foreach ($o in @($ws, $wb, $excel)) {
            if ($o) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) }
        }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}

# --------------------------------------------------------------------------
# 2) AD-Index aufbauen: alle AKTIVEN User einmal laden und normalisiert
#    indizieren. Das ist schneller als pro Person eine LDAP-Abfrage und
#    erlaubt Fuzzy-/Fallback-Matching im Speicher.
# --------------------------------------------------------------------------
function Build-AdUserIndex {
    param([hashtable]$AdParams)

    Write-Log "Lade aktive AD-Benutzer (givenName, sn, displayName, cn) ..." 'INFO'
    $users = Get-ADUser -Filter 'Enabled -eq $true' `
                        -Properties givenName, sn, displayName, cn @AdParams

    # Index: normalisierter Schluessel -> Liste von Usern (Kollisionen moeglich)
    $byGivenSn      = @{}   # "given|sn"
    $byDisplay      = @{}   # normalisierter displayName
    $byCn           = @{}   # normalisierter cn

    $addTo = {
        param($map, $key, $user)
        if (-not $key) { return }
        if (-not $map.ContainsKey($key)) {
            $map[$key] = New-Object System.Collections.Generic.List[object]
        }
        # gleichen SamAccountName nicht doppelt indexieren
        if (-not ($map[$key] | Where-Object { $_.SamAccountName -eq $user.SamAccountName })) {
            $map[$key].Add($user)
        }
    }

    foreach ($u in $users) {
        $gn = ConvertTo-NormalizedName $u.givenName
        $sn = ConvertTo-NormalizedName $u.sn
        if ($gn -and $sn) { & $addTo $byGivenSn "$gn|$sn" $u }
        & $addTo $byDisplay (ConvertTo-NormalizedName $u.displayName) $u
        & $addTo $byCn      (ConvertTo-NormalizedName $u.cn)          $u
    }

    Write-Log ("AD-Index aufgebaut: {0} aktive Benutzer." -f $users.Count) 'OK'
    return [pscustomobject]@{
        ByGivenSn = $byGivenSn
        ByDisplay = $byDisplay
        ByCn      = $byCn
    }
}

# --------------------------------------------------------------------------
# 3) Matching einer Person gegen den AD-Index.
#    Rueckgabe: PSObject mit Status = Matched | Ambiguous | NotFound
# --------------------------------------------------------------------------
function Resolve-AdUser {
    param(
        [pscustomobject]$Person,
        [pscustomobject]$Index
    )

    $givenVariants = Get-GivenNameVariants $Person.Vorname
    $snVariants    = Get-SurnameVariants   $Person.Nachname

    # -- Strategie A (primaer): givenName + sn, alle Varianten kombinieren --
    $candidates = New-Object System.Collections.Generic.List[object]
    $usedStrategy = $null

    foreach ($g in $givenVariants) {
        foreach ($s in $snVariants) {
            $key = "$g|$s"
            if ($Index.ByGivenSn.ContainsKey($key)) {
                foreach ($u in $Index.ByGivenSn[$key]) {
                    if (-not ($candidates | Where-Object { $_.SamAccountName -eq $u.SamAccountName })) {
                        $candidates.Add($u)
                    }
                }
            }
        }
    }
    if ($candidates.Count -gt 0) { $usedStrategy = 'givenName+sn' }

    # -- Strategie B (Fallback): displayName / cn == "Vorname Nachname" --
    if ($candidates.Count -eq 0) {
        $displayKeys = @(
            (ConvertTo-NormalizedName $Person.Anzeige),
            (ConvertTo-NormalizedName "$($Person.Nachname) $($Person.Vorname)")
        ) | Select-Object -Unique
        foreach ($dk in $displayKeys) {
            foreach ($map in @($Index.ByDisplay, $Index.ByCn)) {
                if ($map.ContainsKey($dk)) {
                    foreach ($u in $map[$dk]) {
                        if (-not ($candidates | Where-Object { $_.SamAccountName -eq $u.SamAccountName })) {
                            $candidates.Add($u)
                        }
                    }
                }
            }
        }
        if ($candidates.Count -gt 0) { $usedStrategy = 'displayName/cn' }
    }

    # -- Ergebnis bewerten --
    if ($candidates.Count -eq 1) {
        return [pscustomobject]@{
            Person = $Person; Status = 'Matched'
            User = $candidates[0]; Strategy = $usedStrategy; Candidates = $candidates
        }
    }
    elseif ($candidates.Count -gt 1) {
        return [pscustomobject]@{
            Person = $Person; Status = 'Ambiguous'
            User = $null; Strategy = $usedStrategy; Candidates = $candidates
        }
    }
    else {
        return [pscustomobject]@{
            Person = $Person; Status = 'NotFound'
            User = $null; Strategy = $null; Candidates = @()
        }
    }
}

# --------------------------------------------------------------------------
# 4) HTML-Report (Vorher/Nachher) erzeugen - rein optional, rein visuell.
# --------------------------------------------------------------------------
function New-HtmlReport {
    param(
        [string]$Path, [string]$GroupName, [bool]$DryRun,
        [object[]]$Keep, [object[]]$ToAdd, [object[]]$ToRemove,
        [object[]]$Ambiguous, [object[]]$NotFound,
        [string]$ApplyScriptPath
    )

    function _rows($items, $render) {
        if (-not $items -or $items.Count -eq 0) {
            return "<tr><td class='empty' colspan='3'>-- keine --</td></tr>"
        }
        ($items | ForEach-Object { & $render $_ }) -join "`n"
    }

    $enc = { param($s) [System.Web.HttpUtility]::HtmlEncode($s) }
    Add-Type -AssemblyName System.Web -ErrorAction SilentlyContinue

    $addRows = _rows $ToAdd    { param($x) "<tr><td class='badge add'>+ NEU</td><td>$(& $enc $x.Person.Anzeige)</td><td>$(& $enc $x.User.SamAccountName)</td></tr>" }
    $remRows = _rows $ToRemove { param($x) "<tr><td class='badge rem'>&minus; RAUS</td><td>$(& $enc $x.SamAccountName)</td><td>$(& $enc $x.Name)</td></tr>" }
    $keepRows = _rows $Keep    { param($x) "<tr><td class='badge keep'>= bleibt</td><td>$(& $enc $x.Person.Anzeige)</td><td>$(& $enc $x.User.SamAccountName)</td></tr>" }
    $ambRows = _rows $Ambiguous { param($x) "<tr><td class='badge warn'>? mehrdeutig</td><td>$(& $enc $x.Person.Anzeige)</td><td>$(& $enc (($x.Candidates | ForEach-Object { $_.SamAccountName }) -join ', '))</td></tr>" }
    $nfRows = _rows $NotFound  { param($x) "<tr><td class='badge miss'>x kein Treffer</td><td>$(& $enc $x.Person.Anzeige)</td><td>-</td></tr>" }

    $mode = "ANALYSE (nur gelesen - es wurde NICHTS im AD geaendert)"
    $now  = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

    $html = @"
<!DOCTYPE html><html lang='de'><head><meta charset='utf-8'>
<title>Sync-Report $GroupName</title>
<style>
 body{font-family:Segoe UI,Arial,sans-serif;margin:24px;color:#1f2933;background:#f5f7fa}
 h1{font-size:20px} h2{font-size:15px;margin-top:24px}
 .mode{display:inline-block;padding:4px 10px;border-radius:6px;background:#e3f2fd;color:#0b5cad;font-weight:600}
 .demo{background:#fff8e1;border:1px solid #ffe082;color:#8a6d00;padding:10px 14px;border-radius:8px;margin:14px 0;font-size:13px}
 code{background:#eef2f6;padding:1px 5px;border-radius:4px;font-family:Consolas,monospace}
 .cards{display:flex;gap:12px;flex-wrap:wrap;margin:16px 0}
 .card{flex:1;min-width:120px;background:#fff;border-radius:10px;padding:14px;box-shadow:0 1px 3px rgba(0,0,0,.08);text-align:center}
 .card .n{font-size:26px;font-weight:700} .card.add .n{color:#137333}.card.rem .n{color:#b3261e}
 .card.keep .n{color:#3c4043}.card.warn .n{color:#a56300}.card.miss .n{color:#8f4700}
 table{border-collapse:collapse;width:100%;background:#fff;border-radius:8px;overflow:hidden;margin-top:6px}
 th,td{padding:7px 10px;text-align:left;border-bottom:1px solid #eceff1;font-size:13px}
 th{background:#eef2f6}
 .badge{font-weight:600;white-space:nowrap} .add{color:#137333}.rem{color:#b3261e}
 .keep{color:#5f6368}.warn{color:#a56300}.miss{color:#8f4700}
 .empty{color:#9aa0a6;font-style:italic}
 footer{margin-top:20px;color:#80868b;font-size:12px}
</style></head><body>
<h1>AD-Gruppen-Sync &mdash; $(& $enc $GroupName)</h1>
<p><span class='mode'>$mode</span> &nbsp; erstellt: $now</p>
<div class='demo'><b>Naechster Schritt:</b> Diese Analyse hat nichts veraendert. Zum tatsaechlichen
Ergaenzen/Entfernen das generierte, pruefbare Skript ausfuehren:<br>
<code>$(& $enc $ApplyScriptPath)</code><br>
Es fragt vor dem Schreiben noch einmal nach und unterstuetzt <code>-WhatIf</code>.</div>
<div class='cards'>
 <div class='card add'><div class='n'>$($ToAdd.Count)</div>hinzufuegen</div>
 <div class='card rem'><div class='n'>$($ToRemove.Count)</div>entfernen</div>
 <div class='card keep'><div class='n'>$($Keep.Count)</div>bleibt</div>
 <div class='card warn'><div class='n'>$($Ambiguous.Count)</div>mehrdeutig</div>
 <div class='card miss'><div class='n'>$($NotFound.Count)</div>kein Treffer</div>
</div>
<h2>Hinzuzufuegen (in Excel, fehlt in Gruppe)</h2>
<table><tr><th>Status</th><th>Person (Excel)</th><th>sAMAccountName</th></tr>$addRows</table>
<h2>Zu entfernen (in Gruppe, nicht mehr in Excel)</h2>
<table><tr><th>Status</th><th>sAMAccountName</th><th>Name</th></tr>$remRows</table>
<h2>Mehrdeutige Treffer (manuell pruefen!)</h2>
<table><tr><th>Status</th><th>Person (Excel)</th><th>Kandidaten</th></tr>$ambRows</table>
<h2>Nicht gefundene Personen (manuell pruefen!)</h2>
<table><tr><th>Status</th><th>Person (Excel)</th><th></th></tr>$nfRows</table>
<h2>Bereits korrekt in der Gruppe</h2>
<table><tr><th>Status</th><th>Person (Excel)</th><th>sAMAccountName</th></tr>$keepRows</table>
<footer>Automatisch generiert von Sync-JobcenterGroup.ps1 (Analyse-Modus, nur lesend)</footer>
</body></html>
"@

    $html | Out-File -FilePath $Path -Encoding UTF8
    Write-Log "HTML-Report geschrieben: $Path" 'OK'
}

# --------------------------------------------------------------------------
# 4b) CSV-Export aller Kategorien (fuer Weitergabe an den Fachbereich).
# --------------------------------------------------------------------------
function Export-AnalyseCsv {
    param(
        [string]$Path,
        [object[]]$ToAdd, [object[]]$ToRemove, [object[]]$Keep,
        [object[]]$Ambiguous, [object[]]$NotFound
    )
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($m in $ToAdd)    { $rows.Add([pscustomobject]@{ Aktion='HINZUFUEGEN'; Person=$m.Person.Anzeige; sAMAccountName=$m.User.SamAccountName; Details=$m.Strategy }) }
    foreach ($u in $ToRemove) { $rows.Add([pscustomobject]@{ Aktion='ENTFERNEN';   Person=$u.Name;           sAMAccountName=$u.SamAccountName;      Details='nicht mehr in Excel' }) }
    foreach ($m in $Keep)     { $rows.Add([pscustomobject]@{ Aktion='BLEIBT';       Person=$m.Person.Anzeige; sAMAccountName=$m.User.SamAccountName; Details=$m.Strategy }) }
    foreach ($a in $Ambiguous){ $rows.Add([pscustomobject]@{ Aktion='MEHRDEUTIG';   Person=$a.Person.Anzeige; sAMAccountName='';                    Details=(($a.Candidates | ForEach-Object { $_.SamAccountName }) -join ' | ') }) }
    foreach ($n in $NotFound) { $rows.Add([pscustomobject]@{ Aktion='KEIN_TREFFER'; Person=$n.Person.Anzeige; sAMAccountName='';                    Details='im AD nicht gefunden' }) }

    $rows | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8 -Delimiter ';'
    Write-Log "CSV-Export geschrieben: $Path" 'OK'
}

# --------------------------------------------------------------------------
# 4c) GENERIERT das pruefbare Ausfuehrungsskript ("Apply-*.ps1").
#     Enthaelt je eine kommentierte Zeile pro Aenderung mit aufgeloestem DN.
#     Mehrdeutige/nicht gefundene Personen werden NICHT automatisiert,
#     sondern nur als Kommentar zur manuellen Pruefung ausgewiesen.
# --------------------------------------------------------------------------
function New-ApplyScript {
    param(
        [string]$Path, [string]$GroupName, [string]$Server, [string]$ExcelPath,
        [object[]]$ToAdd, [object[]]$ToRemove, [object[]]$Ambiguous, [object[]]$NotFound
    )

    # Kleiner Helfer: Strings fuer die Einbettung in einfachen Anfuehrungszeichen absichern
    function _q([string]$s) { "'" + ($s -replace "'", "''") + "'" }

    $sb = New-Object System.Text.StringBuilder
    $nl = "`r`n"
    [void]$sb.Append(@"
<#
  ============================================================================
  AUTOMATISCH GENERIERTES AUSFUEHRUNGSSKRIPT  -  BITTE VOR DEM START PRUEFEN!
  ----------------------------------------------------------------------------
  Erzeugt am : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
  Gruppe     : $GroupName
  Grundlage  : $ExcelPath
  ----------------------------------------------------------------------------
  ZUSAMMENFASSUNG:
    Hinzufuegen : $($ToAdd.Count)
    Entfernen   : $($ToRemove.Count)
    (Mehrdeutig : $($Ambiguous.Count),  Kein Treffer: $($NotFound.Count)
     -> NICHT automatisiert, siehe Kommentare weiter unten, bitte manuell klaeren)
  ----------------------------------------------------------------------------
  AUSFUEHREN:
    .\$(Split-Path $Path -Leaf) -WhatIf     # nochmaliger Trockenlauf (aendert nichts)
    .\$(Split-Path $Path -Leaf)             # echte Aenderung (fragt vorher nach)
    .\$(Split-Path $Path -Leaf) -Force      # echte Aenderung ohne Rueckfrage
  ============================================================================
#>
[CmdletBinding(SupportsShouldProcess = `$true, ConfirmImpact = 'High')]
param([switch]`$Force)

`$ErrorActionPreference = 'Stop'
Import-Module ActiveDirectory -ErrorAction Stop

`$Group    = $(_q $GroupName)
`$adParams = @{}
$(if ($Server) { "`$adParams['Server'] = $(_q $Server)" } else { "# (kein -Server gesetzt; Standard-DC wird verwendet)" })

# Sicherheitsabfrage vor echten Aenderungen (entfaellt bei -WhatIf oder -Force)
if (-not `$Force -and -not `$WhatIfPreference) {
    Write-Host ''
    Write-Host "Es werden $($ToAdd.Count) Benutzer HINZUGEFUEGT und $($ToRemove.Count) ENTFERNT." -ForegroundColor Yellow
    `$answer = Read-Host "Wirklich auf Gruppe '`$Group' anwenden? (ja/nein)"
    if (`$answer -ne 'ja') { Write-Host 'Abgebrochen - nichts geaendert.' -ForegroundColor Cyan; return }
}

"@)

    # --- HINZUFUEGEN ---
    [void]$sb.Append("# --------------------------- HINZUFUEGEN ---------------------------$nl")
    if ($ToAdd.Count -eq 0) {
        [void]$sb.Append("# (nichts hinzuzufuegen)$nl")
    } else {
        foreach ($m in $ToAdd) {
            $dn   = $m.User.DistinguishedName
            $note = "$($m.Person.Anzeige) ($($m.User.SamAccountName))"
            [void]$sb.Append("Add-ADGroupMember -Identity `$Group -Members $(_q $dn) @adParams   # + $note$nl")
        }
    }

    # --- ENTFERNEN ---
    [void]$sb.Append("$nl# ---------------------------- ENTFERNEN ----------------------------$nl")
    if ($ToRemove.Count -eq 0) {
        [void]$sb.Append("# (nichts zu entfernen)$nl")
    } else {
        foreach ($u in $ToRemove) {
            $note = "$($u.Name) ($($u.SamAccountName))"
            [void]$sb.Append("Remove-ADGroupMember -Identity `$Group -Members $(_q $u.DistinguishedName) -Confirm:`$false @adParams   # - $note$nl")
        }
    }

    # --- Nur zur Information: nicht automatisierte Faelle ---
    [void]$sb.Append("$nl# ------- MANUELL PRUEFEN (NICHT automatisiert) -------$nl")
    foreach ($a in $Ambiguous) {
        $cands = ($a.Candidates | ForEach-Object { $_.SamAccountName }) -join ', '
        [void]$sb.Append("# MEHRDEUTIG : $($a.Person.Anzeige) -> Kandidaten: $cands$nl")
    }
    foreach ($n in $NotFound) {
        [void]$sb.Append("# KEIN TREFFER: $($n.Person.Anzeige)$nl")
    }
    [void]$sb.Append("$nl Write-Host 'Fertig.' -ForegroundColor Green$nl")

    $sb.ToString() | Out-File -FilePath $Path -Encoding UTF8
    Write-Log "Ausfuehrungsskript generiert: $Path" 'OK'
}

# ==========================================================================
#  HAUPTPROGRAMM
# ==========================================================================
try {
    Write-Log "=== Start ANALYSE fuer Gruppe '$GroupName' ===" 'INFO'
    Write-Log "Excel: $ExcelPath" 'INFO'
    Write-Log "MODUS: ANALYSE (nur lesend) - es wird NICHTS im AD geaendert." 'WARN'

    # AD-Modul laden
    Import-Module ActiveDirectory -ErrorAction Stop

    # Optionalen Server-Parameter fuer alle AD-Cmdlets vorbereiten
    $adParams = @{}
    if ($Server) { $adParams['Server'] = $Server }

    # --- Excel einlesen (inkl. Auswahl des Arbeitsblatts) ---
    $persons = Import-SachbearbeiterListe -Path $ExcelPath -WorksheetName $WorksheetName
    Write-Log ("Excel gelesen: {0} eindeutige Personen." -f $persons.Count) 'OK'

    # --- Zielgruppe pruefen ---
    $group = Get-ADGroup -Identity $GroupName -Properties distinguishedName @adParams
    Write-Log "Zielgruppe gefunden: $($group.DistinguishedName)" 'OK'

    # --- AD-Index aufbauen & matchen ---
    $index    = Build-AdUserIndex -AdParams $adParams
    $matched  = New-Object System.Collections.Generic.List[object]
    $ambiguous= New-Object System.Collections.Generic.List[object]
    $notFound = New-Object System.Collections.Generic.List[object]

    foreach ($p in $persons) {
        $res = Resolve-AdUser -Person $p -Index $index
        switch ($res.Status) {
            'Matched' {
                $matched.Add($res)
                Write-Log ("MATCH  : {0,-28} -> {1} ({2})" -f $p.Anzeige, $res.User.SamAccountName, $res.Strategy) 'INFO'
            }
            'Ambiguous' {
                $ambiguous.Add($res)
                $cands = ($res.Candidates | ForEach-Object { $_.SamAccountName }) -join ', '
                Write-Log ("MEHRDEUTIG: {0} -> Kandidaten: {1}" -f $p.Anzeige, $cands) 'WARN'
            }
            'NotFound' {
                $notFound.Add($res)
                Write-Log ("KEIN TREFFER: {0}" -f $p.Anzeige) 'WARN'
            }
        }
    }

    # --- Soll-Zustand: eindeutige, aktive User aus der Excel ---
    #     (nach SamAccountName deduplizieren, falls zwei Excel-Zeilen denselben
    #      AD-User treffen)
    $desiredUsers = @{}
    foreach ($m in $matched) {
        $desiredUsers[$m.User.SamAccountName] = $m
    }

    # --- Ist-Zustand: NUR direkte Benutzer-Mitglieder der Gruppe ---
    #     Verschachtelte Gruppen bewusst ausschliessen (objectClass -eq 'user').
    $currentMembers = Get-ADGroupMember -Identity $GroupName @adParams |
                      Where-Object { $_.objectClass -eq 'user' }
    $currentSams = @{}
    foreach ($cm in $currentMembers) { $currentSams[$cm.SamAccountName] = $cm }

    Write-Log ("Direkte User-Mitglieder aktuell: {0}" -f $currentMembers.Count) 'INFO'

    # --- Differenzen bilden ---
    $toAdd    = New-Object System.Collections.Generic.List[object]  # match-Objekte
    $toKeep   = New-Object System.Collections.Generic.List[object]
    foreach ($sam in $desiredUsers.Keys) {
        if ($currentSams.ContainsKey($sam)) { $toKeep.Add($desiredUsers[$sam]) }
        else                                { $toAdd.Add($desiredUsers[$sam]) }
    }

    $toRemove = New-Object System.Collections.Generic.List[object]  # AD-Objekte
    foreach ($sam in $currentSams.Keys) {
        if (-not $desiredUsers.ContainsKey($sam)) { $toRemove.Add($currentSams[$sam]) }
    }

    # --- KEINE AD-Aenderung hier! Nur Artefakte fuer die Freigabe erzeugen. ---

    # Generiertes Ausfuehrungsskript (die eigentliche Schreiblogik zum "Absegnen")
    New-ApplyScript -Path $ApplyScriptPath -GroupName $GroupName -Server $Server -ExcelPath $ExcelPath `
        -ToAdd $toAdd.ToArray() -ToRemove $toRemove.ToArray() `
        -Ambiguous $ambiguous.ToArray() -NotFound $notFound.ToArray()

    # HTML-Report (Vorher/Nachher)
    New-HtmlReport -Path $HtmlReportPath -GroupName $GroupName -DryRun $true `
        -Keep $toKeep.ToArray() -ToAdd $toAdd.ToArray() -ToRemove $toRemove.ToArray() `
        -Ambiguous $ambiguous.ToArray() -NotFound $notFound.ToArray() `
        -ApplyScriptPath $ApplyScriptPath

    # CSV-Export (fuer Weitergabe / Ablage)
    Export-AnalyseCsv -Path $CsvPath `
        -ToAdd $toAdd.ToArray() -ToRemove $toRemove.ToArray() -Keep $toKeep.ToArray() `
        -Ambiguous $ambiguous.ToArray() -NotFound $notFound.ToArray()

    # --- Zusammenfassung ---
    Write-Log "==================== ZUSAMMENFASSUNG (ANALYSE) ====================" 'INFO'
    Write-Log ("Excel-Personen gesamt : {0}" -f $persons.Count)      'INFO'
    Write-Log ("Eindeutig gematcht    : {0}" -f $matched.Count)      'OK'
    Write-Log ("Wuerde hinzufuegen    : {0}" -f $toAdd.Count)        'ADD'
    Write-Log ("Wuerde entfernen      : {0}" -f $toRemove.Count)     'REMOVE'
    Write-Log ("Unveraendert (bleibt) : {0}" -f $toKeep.Count)       'INFO'
    Write-Log ("Mehrdeutig            : {0}" -f $ambiguous.Count)    'WARN'
    Write-Log ("Nicht gefunden        : {0}" -f $notFound.Count)     'WARN'
    Write-Log "==================================================================" 'INFO'
    Write-Log "Es wurde NICHTS im AD geaendert (reine Analyse)." 'WARN'
    Write-Log "" 'INFO'
    Write-Log "NAECHSTE SCHRITTE:" 'INFO'
    Write-Log ("  1) Report ansehen : {0}" -f $HtmlReportPath) 'INFO'
    Write-Log ("  2) Skript pruefen : {0}" -f $ApplyScriptPath) 'INFO'
    Write-Log ("  3) Ausfuehren     : `"{0}`" -WhatIf   (Test), danach ohne -WhatIf" -f $ApplyScriptPath) 'INFO'
    Write-Log ("CSV-Export          : {0}" -f $CsvPath) 'INFO'
    Write-Log ("Logdatei            : {0}" -f $LogPath) 'INFO'
    Write-Log "=== Ende Analyse ===" 'OK'
}
catch {
    Write-Log ("ABBRUCH: {0}" -f $_.Exception.Message) 'ERROR'
    Write-Log ($_.ScriptStackTrace) 'ERROR'
    exit 1
}
