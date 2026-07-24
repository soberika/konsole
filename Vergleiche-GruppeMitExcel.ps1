#Requires -Version 5.1
<#
.SYNOPSIS
    OFFLINE-Vergleich: Excel-Sollliste  <->  aktuelle Gruppen-Mitgliederliste.
    Ergebnis: wer HINZUGEFUEGT und wer ENTFERNT werden muss. Kein AD noetig.

.DESCRIPTION
    Zwei Eingaben, kein Active Directory:

      -ExcelPath    : Sachbearbeiter-Excel (Soll), Spalten "Vorname"/"Nachname".
      -MembersPath  : aktuelle Mitglieder der Gruppe (Ist) als Textdatei -
                      eine Zeile je Eintrag im Format "Nachname, Vorname"
                      (so wie es viele AD-Tools anzeigen).

    Nicht-Personen werden automatisch ignoriert:
      * Zeilen OHNE Komma (z. B. "JC.Controlling"),
      * verschachtelte Gruppen ("Liste_*", "*_Team*", konfigurierbar).
    Angehaengte Ziffern an Dubletten-Namen (z. B. "Schmidt1") werden fuer den
    Vergleich entfernt. Umlaute/"genannt"-Namen werden normalisiert.

    Ausgabe: eine CSV (Aktion; Vorname; Nachname) mit
      HINZUFUEGEN | ENTFERNEN | BLEIBT | IGNORIERT

.PARAMETER ExcelPath
    Pfad zur Soll-Excel (.xlsx).

.PARAMETER MembersPath
    Pfad zur Ist-Liste (.txt/.csv), eine Zeile "Nachname, Vorname".

.PARAMETER WorksheetName
    Optionales Arbeitsblatt in der Excel.

.PARAMETER OutputPath
    Ziel-CSV. Standard: neben der Excel mit Zeitstempel.

.PARAMETER IgnorePatterns
    Regex-Muster fuer zu ignorierende Nicht-Personen-Zeilen.
    Standard: '^Liste_', '^JC\.', '_Team', 'Springer', 'Teamleiter',
    'Rechtsanwendung', 'Controlling'.

.EXAMPLE
    .\Vergleiche-GruppeMitExcel.ps1 `
        -ExcelPath   .\Source\Liste_SachbearbeiterLeistung.xlsx `
        -MembersPath .\Source\Gruppe_Leistung_Ist.txt

.NOTES
    Rein namensbasiert. Bei ungewoehnlichen Schreibweisen Ergebnis stichprobenartig
    pruefen. Es wird NICHTS veraendert - nur die Ergebnisliste erzeugt.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ExcelPath,

    [Parameter(Mandatory = $true)]
    [string]$MembersPath,

    [Parameter(Mandatory = $false)]
    [string]$WorksheetName,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath,

    [Parameter(Mandatory = $false)]
    [string[]]$IgnorePatterns = @('^Liste_', '^JC\.', '_Team', 'Springer', 'Teamleiter', 'Rechtsanwendung', 'Controlling')
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

# --- Namens-Normalisierung (Umlaute/Diakritika/Sonderzeichen) ---
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

# --- Vergleichsschluessel: normalisiert + angehaengte Ziffern je Token entfernt ---
function Get-MatchKey {
    param([string]$Vorname, [string]$Nachname)
    $strip = { param($s) (($s -split ' ') | ForEach-Object { $_ -replace '\d+$', '' }) -join ' ' }
    $g = & $strip (ConvertTo-NormalizedName $Vorname)
    $s = & $strip (ConvertTo-NormalizedName $Nachname)
    return ($g.Trim() + '|' + $s.Trim())
}

# --- Excel lesen (ImportExcel bevorzugt, sonst COM) ---
function Import-XlsxRows {
    param([string]$Path, [string]$WorksheetName)
    if (Get-Module -ListAvailable -Name ImportExcel) {
        Import-Module ImportExcel -ErrorAction Stop
        if ($WorksheetName) { return Import-Excel -Path $Path -WorksheetName $WorksheetName }
        return Import-Excel -Path $Path
    }
    $excel = $null; $wb = $null; $ws = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false; $excel.DisplayAlerts = $false
        $wb = $excel.Workbooks.Open((Resolve-Path $Path).Path, $null, $true)
        $ws = if ($WorksheetName) { $wb.Worksheets.Item($WorksheetName) } else { $wb.Worksheets.Item(1) }
        $used = $ws.UsedRange; $data = $used.Value2
        $rowCount = $used.Rows.Count; $colCount = $used.Columns.Count
        $headers = @{}
        for ($c = 1; $c -le $colCount; $c++) { $h = "$($data[1, $c])".Trim(); if ($h) { $headers[$h] = $c } }
        $result = New-Object System.Collections.Generic.List[object]
        for ($rr = 2; $rr -le $rowCount; $rr++) {
            $obj = [ordered]@{}
            foreach ($h in $headers.Keys) { $obj[$h] = "$($data[$rr, $headers[$h]])" }
            $result.Add([pscustomobject]$obj)
        }
        return $result
    }
    finally {
        if ($wb) { $wb.Close($false) | Out-Null }; if ($excel) { $excel.Quit() | Out-Null }
        foreach ($o in @($ws, $wb, $excel)) { if ($o) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}

# --- Spalte tolerant lesen ---
function Get-Col {
    param($Row, [string[]]$Names)
    foreach ($p in $Row.PSObject.Properties) {
        $pn = ($p.Name).Trim()
        foreach ($n in $Names) { if ($pn -ieq $n) { return "$($p.Value)".Trim() } }
    }
    return ''
}

# ==========================================================================
#  HAUPTPROGRAMM
# ==========================================================================
try {
    Write-Host "=== OFFLINE-Vergleich  Excel (Soll)  <->  Gruppe (Ist) ===" -ForegroundColor Cyan
    if (-not (Test-Path -LiteralPath $ExcelPath))   { throw "Excel-Datei nicht gefunden: $ExcelPath" }
    if (-not (Test-Path -LiteralPath $MembersPath)) { throw "Mitgliederliste nicht gefunden: $MembersPath" }

    # --- SOLL: Personen aus Excel (nur Zeilen mit Namen, dedupliziert) ---
    $soll = @{}
    foreach ($r in (Import-XlsxRows -Path $ExcelPath -WorksheetName $WorksheetName)) {
        $vn = Get-Col $r @('Vorname', 'GivenName')
        $nn = Get-Col $r @('Nachname', 'Surname', 'Name')
        if (-not $vn -and -not $nn) { continue }
        $soll[(Get-MatchKey $vn $nn)] = [pscustomobject]@{ Vorname = $vn; Nachname = $nn }
    }

    # --- IST: Mitgliederliste "Nachname, Vorname" (Nicht-Personen ignorieren) ---
    $ist = @{}; $ignored = New-Object System.Collections.Generic.List[string]
    foreach ($line in (Get-Content -LiteralPath $MembersPath -Encoding UTF8)) {
        $t = $line.Trim()
        if (-not $t) { continue }

        # Nicht-Personen: kein Komma ODER Ignore-Muster
        $isIgnored = ($t -notmatch ',')
        if (-not $isIgnored) {
            foreach ($pat in $IgnorePatterns) { if ($t -match $pat) { $isIgnored = $true; break } }
        }
        if ($isIgnored) { $ignored.Add($t); continue }

        $parts = $t -split ',', 2
        $nn = $parts[0].Trim(); $vn = $parts[1].Trim()
        $ist[(Get-MatchKey $vn $nn)] = [pscustomobject]@{ Vorname = $vn; Nachname = $nn }
    }

    # --- Differenzen ---
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
    foreach ($g in $ignored) {
        $results.Add([pscustomobject]@{ Aktion = 'IGNORIERT'; Vorname = $g; Nachname = '' })
    }

    $nAdd  = @($results | Where-Object Aktion -eq 'HINZUFUEGEN').Count
    $nRem  = @($results | Where-Object Aktion -eq 'ENTFERNEN').Count
    $nKeep = @($results | Where-Object Aktion -eq 'BLEIBT').Count

    # Sortierung: HINZUFUEGEN, ENTFERNEN, BLEIBT, IGNORIERT; dann Nachname
    $order = @{ 'HINZUFUEGEN' = 0; 'ENTFERNEN' = 1; 'BLEIBT' = 2; 'IGNORIERT' = 3 }
    $sorted = $results | Sort-Object @{ Expression = { $order[$_.Aktion] } }, Nachname, Vorname

    # --- Ausgabe schreiben ---
    if (-not $OutputPath) {
        $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
        $dir = Split-Path -Parent (Resolve-Path $ExcelPath)
        $OutputPath = Join-Path $dir "Abgleich_Ergebnis_$stamp.csv"
    }
    $sorted | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding UTF8 -Delimiter ';'

    # --- Zusammenfassung ---
    Write-Host ("Soll (Excel)            : {0}" -f $soll.Count) -ForegroundColor Gray
    Write-Host ("Ist (Personen, Gruppe)  : {0}   [ignoriert: {1}]" -f $ist.Count, $ignored.Count) -ForegroundColor Gray
    Write-Host "`n==================== ERGEBNIS ====================" -ForegroundColor Cyan
    Write-Host ("Hinzufuegen : {0}" -f $nAdd)  -ForegroundColor Green
    Write-Host ("Entfernen   : {0}" -f $nRem)  -ForegroundColor Magenta
    Write-Host ("Bleibt      : {0}" -f $nKeep) -ForegroundColor Gray
    Write-Host ("Ignoriert   : {0}" -f $ignored.Count) -ForegroundColor Yellow
    Write-Host "=================================================="
    Write-Host ("Ergebnisliste: {0}" -f $OutputPath) -ForegroundColor Green
}
catch {
    Write-Host ("ABBRUCH: {0}" -f $_.Exception.Message) -ForegroundColor Red
    exit 1
}
