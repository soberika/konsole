#Requires -Version 5.1
<#
.SYNOPSIS
    Einfachster Offline-Abgleich: du gibst zwei Listen, bekommst EINE Export-Datei.
    Kein Active Directory, keine Parameter noetig.

.DESCRIPTION
    Zwei Eingaben (Reihenfolge egal):
      * SOLL-Liste : Excel (.xlsx) ODER CSV mit Spalten "Vorname"/"Nachname"
                     = wer in der Gruppe sein SOLL.
      * IST-Liste  : Textdatei (.txt), eine Zeile je Mitglied im Format
                     "Nachname, Vorname" = wer AKTUELL in der Gruppe ist
                     (einfach aus eurem Benutzer-Tool hineinkopieren).

    Das Skript erkennt automatisch, welche Datei welche ist, ignoriert
    Nicht-Personen (verschachtelte Gruppen "Liste_*", "JC.*" usw.), behandelt
    Umlaute, "genannt"-Namen und Ziffern-Dubletten ("Schmidt1") und schreibt
    das Ergebnis als Datei (XLSX falls moeglich, sonst CSV) - farblich sortiert:
      HINZUFUEGEN | ENTFERNEN | BLEIBT | IGNORIERT

    NUTZUNG (drei Wege - alle ohne AD):
      1) Datei "Abgleich.cmd" doppelklicken: nimmt die zwei Dateien aus diesem
         Ordner automatisch.
      2) Beide Listen auf "Abgleich.cmd" ZIEHEN (Drag and Drop).
      3) In PowerShell:  .\Abgleich.ps1  Soll.xlsx  Ist.txt

.NOTES
    Rein namensbasiert. Es wird NICHTS veraendert - nur die Ergebnisdatei erzeugt.
#>

[CmdletBinding()]
param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$Files,

    # Muster fuer Nicht-Personen (verschachtelte Gruppen / Funktionskonten)
    [string[]]$IgnorePatterns = @('^Liste_', '^JC\.', '_Team', 'Springer', 'Teamleiter', 'Rechtsanwendung', 'Controlling')
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

# ==========================================================================
#  HILFSFUNKTIONEN
# ==========================================================================
function ConvertTo-NormalizedName {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return '' }
    $n = $Name.Trim().ToLowerInvariant()
    $n = $n -replace 'ä', 'ae' -replace 'ö', 'oe' -replace 'ü', 'ue' -replace 'ß', 'ss'
    $sb = [System.Text.StringBuilder]::new()
    foreach ($ch in $n.Normalize([Text.NormalizationForm]::FormD).ToCharArray()) {
        if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch) -ne
            [Globalization.UnicodeCategory]::NonSpacingMark) { [void]$sb.Append($ch) }
    }
    $n = $sb.ToString() -replace '[^a-z0-9]', ' '
    return ($n -replace '\s+', ' ').Trim()
}

# Vergleichsschluessel: normalisiert + angehaengte Ziffern je Token entfernt
function Get-MatchKey {
    param([string]$Vorname, [string]$Nachname)
    $strip = { param($s) (($s -split ' ') | ForEach-Object { $_ -replace '\d+$', '' }) -join ' ' }
    $g = (& $strip (ConvertTo-NormalizedName $Vorname)).Trim()
    $s = (& $strip (ConvertTo-NormalizedName $Nachname)).Trim()
    return "$g|$s"
}

function Get-Col {
    param($Row, [string[]]$Names)
    foreach ($p in $Row.PSObject.Properties) {
        $pn = ($p.Name).Trim()
        foreach ($n in $Names) { if ($pn -ieq $n) { return "$($p.Value)".Trim() } }
    }
    return ''
}

# Excel/CSV der SOLL-Liste einlesen
function Import-SollRows {
    param([string]$Path)
    $ext = [IO.Path]::GetExtension($Path).ToLowerInvariant()
    if ($ext -eq '.csv') {
        $first = Get-Content -LiteralPath $Path -TotalCount 1 -Encoding UTF8
        $delim = if ($first -match ';') { ';' } else { ',' }
        return Import-Csv -LiteralPath $Path -Delimiter $delim -Encoding UTF8
    }
    # XLSX: ImportExcel bevorzugt, sonst Excel-COM
    if (Get-Module -ListAvailable -Name ImportExcel) {
        Import-Module ImportExcel -ErrorAction Stop
        return Import-Excel -Path $Path
    }
    $excel = $null; $wb = $null; $ws = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false; $excel.DisplayAlerts = $false
        $wb = $excel.Workbooks.Open((Resolve-Path $Path).Path, $null, $true)
        $ws = $wb.Worksheets.Item(1); $used = $ws.UsedRange; $data = $used.Value2
        $rowCount = $used.Rows.Count; $colCount = $used.Columns.Count
        $headers = @{}
        for ($c = 1; $c -le $colCount; $c++) { $h = "$($data[1, $c])".Trim(); if ($h) { $headers[$h] = $c } }
        $out = New-Object System.Collections.Generic.List[object]
        for ($rr = 2; $rr -le $rowCount; $rr++) {
            $obj = [ordered]@{}
            foreach ($h in $headers.Keys) { $obj[$h] = "$($data[$rr, $headers[$h]])" }
            $out.Add([pscustomobject]$obj)
        }
        return $out
    }
    finally {
        if ($wb) { $wb.Close($false) | Out-Null }; if ($excel) { $excel.Quit() | Out-Null }
        foreach ($o in @($ws, $wb, $excel)) { if ($o) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}

# Erkennt anhand von Endung/Inhalt, ob eine Datei SOLL oder IST ist
function Get-FileRole {
    param([string]$Path)
    $ext = [IO.Path]::GetExtension($Path).ToLowerInvariant()
    if ($ext -in @('.xlsx', '.xlsm')) { return 'SOLL' }
    if ($ext -eq '.txt') { return 'IST' }
    if ($ext -eq '.csv') {
        $first = (Get-Content -LiteralPath $Path -TotalCount 1 -Encoding UTF8)
        if ($first -match '(?i)vorname' -or $first -match '(?i)nachname' -or
            $first -match '(?i)givenname' -or $first -match '(?i)surname') { return 'SOLL' }
        return 'IST'
    }
    return 'UNBEKANNT'
}

# ==========================================================================
#  1) Eingabedateien bestimmen (Argumente ODER Ordner-Scan)
# ==========================================================================
try {
    $baseDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }

    $candidates = @()
    if ($Files -and $Files.Count -gt 0) {
        $candidates = $Files | Where-Object { $_ -and (Test-Path -LiteralPath $_) }
    }
    else {
        # Ordner durchsuchen (eigene Ergebnisdateien ausschliessen)
        Write-Host "Keine Dateien uebergeben - durchsuche Ordner: $baseDir" -ForegroundColor Gray
        $candidates = Get-ChildItem -LiteralPath $baseDir -File |
            Where-Object { $_.Extension -in @('.xlsx', '.xlsm', '.csv', '.txt') -and $_.Name -notmatch '^Abgleich_Ergebnis' } |
            Select-Object -ExpandProperty FullName
    }

    if (-not $candidates -or @($candidates).Count -eq 0) {
        throw "Keine passenden Dateien gefunden. Bitte SOLL-Liste (xlsx/csv) und IST-Liste (txt) angeben."
    }

    # Rollen zuordnen
    $sollList = @($candidates | Where-Object { (Get-FileRole $_) -eq 'SOLL' })
    $istList  = @($candidates | Where-Object { (Get-FileRole $_) -eq 'IST'  })

    if ($sollList.Count -ne 1 -or $istList.Count -ne 1) {
        $liste = ($candidates | ForEach-Object { "  - $([IO.Path]::GetFileName($_))  [$(Get-FileRole $_)]" }) -join "`n"
        throw ("Konnte SOLL und IST nicht EINDEUTIG zuordnen (gefunden: SOLL=$($sollList.Count), IST=$($istList.Count)).`n" +
               "Dateien:`n$liste`n" +
               "Bitte genau ZWEI Dateien angeben/hierher ziehen: eine Excel/CSV (SOLL, mit Vorname/Nachname) " +
               "UND eine .txt (IST, 'Nachname, Vorname'). Tipp: beide Dateien direkt auf 'Abgleich.cmd' ziehen.")
    }
    $sollFile = $sollList[0]; $istFile = $istList[0]

    Write-Host "SOLL-Liste (Excel) : $([IO.Path]::GetFileName($sollFile))" -ForegroundColor Cyan
    Write-Host "IST-Liste  (Gruppe): $([IO.Path]::GetFileName($istFile))"  -ForegroundColor Cyan

    # ======================================================================
    #  2) Einlesen
    # ======================================================================
    # SOLL
    $soll = @{}
    foreach ($r in (Import-SollRows -Path $sollFile)) {
        $vn = Get-Col $r @('Vorname', 'GivenName', 'Rufname')
        $nn = Get-Col $r @('Nachname', 'Surname', 'Name')
        if (-not $vn -and -not $nn) { continue }
        $soll[(Get-MatchKey $vn $nn)] = [pscustomobject]@{ Vorname = $vn; Nachname = $nn }
    }

    # IST ("Nachname, Vorname" je Zeile; Nicht-Personen ignorieren)
    $ist = @{}; $ignored = New-Object System.Collections.Generic.List[string]
    foreach ($line in (Get-Content -LiteralPath $istFile -Encoding UTF8)) {
        $t = "$line".Trim()
        if (-not $t) { continue }
        $skip = ($t -notmatch ',')
        if (-not $skip) { foreach ($pat in $IgnorePatterns) { if ($t -match $pat) { $skip = $true; break } } }
        if ($skip) { $ignored.Add($t); continue }
        $parts = $t -split ',', 2
        $nn = $parts[0].Trim(); $vn = $parts[1].Trim()
        $ist[(Get-MatchKey $vn $nn)] = [pscustomobject]@{ Vorname = $vn; Nachname = $nn }
    }

    # ======================================================================
    #  3) Abgleich
    # ======================================================================
    $results = New-Object System.Collections.Generic.List[object]
    foreach ($k in $soll.Keys) {
        $aktion = if ($ist.ContainsKey($k)) { 'BLEIBT' } else { 'HINZUFUEGEN' }
        $results.Add([pscustomobject]@{ Aktion = $aktion; Vorname = $soll[$k].Vorname; Nachname = $soll[$k].Nachname })
    }
    foreach ($k in $ist.Keys) {
        if (-not $soll.ContainsKey($k)) {
            $results.Add([pscustomobject]@{ Aktion = 'ENTFERNEN'; Vorname = $ist[$k].Vorname; Nachname = $ist[$k].Nachname })
        }
    }
    foreach ($g in $ignored) { $results.Add([pscustomobject]@{ Aktion = 'IGNORIERT'; Vorname = $g; Nachname = '' }) }

    $nAdd  = @($results | Where-Object Aktion -eq 'HINZUFUEGEN').Count
    $nRem  = @($results | Where-Object Aktion -eq 'ENTFERNEN').Count
    $nKeep = @($results | Where-Object Aktion -eq 'BLEIBT').Count

    $order  = @{ 'HINZUFUEGEN' = 0; 'ENTFERNEN' = 1; 'BLEIBT' = 2; 'IGNORIERT' = 3 }
    $sorted = $results | Sort-Object @{ Expression = { $order[$_.Aktion] } }, Nachname, Vorname

    # ======================================================================
    #  4) Export-Datei schreiben (XLSX falls moeglich, sonst CSV) + oeffnen
    # ======================================================================
    $stamp   = Get-Date -Format 'yyyyMMdd_HHmmss'
    $csvOut  = Join-Path $baseDir "Abgleich_Ergebnis_$stamp.csv"
    $sorted | Export-Csv -LiteralPath $csvOut -NoTypeInformation -Encoding UTF8 -Delimiter ';'
    $finalOut = $csvOut

    if (Get-Module -ListAvailable -Name ImportExcel) {
        try {
            Import-Module ImportExcel -ErrorAction Stop
            $xlsxOut = Join-Path $baseDir "Abgleich_Ergebnis_$stamp.xlsx"
            # Farbige Hervorhebung je Aktion
            $cf = @(
                New-ConditionalText -Text 'HINZUFUEGEN' -BackgroundColor '#C6EFCE' -ConditionalTextColor '#006100'
                New-ConditionalText -Text 'ENTFERNEN'   -BackgroundColor '#FFC7CE' -ConditionalTextColor '#9C0006'
                New-ConditionalText -Text 'IGNORIERT'   -BackgroundColor '#FFEB9C' -ConditionalTextColor '#9C6500'
            )
            $sorted | Export-Excel -Path $xlsxOut -WorksheetName 'Abgleich' -AutoSize -FreezeTopRow -BoldTopRow -AutoFilter -ConditionalText $cf
            $finalOut = $xlsxOut
        }
        catch {
            Write-Host ("XLSX-Export uebersprungen ({0}) - CSV wurde geschrieben." -f $_.Exception.Message) -ForegroundColor Yellow
        }
    }

    # ======================================================================
    #  5) Zusammenfassung + Ergebnis oeffnen
    # ======================================================================
    Write-Host "`n==================== ERGEBNIS ====================" -ForegroundColor Cyan
    Write-Host ("Soll (Excel)      : {0}" -f $soll.Count)
    Write-Host ("Ist (in Gruppe)   : {0}   [ignoriert: {1}]" -f $ist.Count, $ignored.Count)
    Write-Host ("Hinzufuegen       : {0}" -f $nAdd)  -ForegroundColor Green
    Write-Host ("Entfernen         : {0}" -f $nRem)  -ForegroundColor Magenta
    Write-Host ("Bleibt            : {0}" -f $nKeep) -ForegroundColor Gray
    Write-Host "=================================================="
    Write-Host ("Ergebnisdatei     : {0}" -f $finalOut) -ForegroundColor Green

    try { Invoke-Item -LiteralPath $finalOut } catch { }
}
catch {
    Write-Host ("`nABBRUCH: {0}" -f $_.Exception.Message) -ForegroundColor Red
    exit 1
}
