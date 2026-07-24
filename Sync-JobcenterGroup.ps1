#Requires -Version 5.1
<#
.SYNOPSIS
    Synchronisiert eine AD-Sicherheitsgruppe (z. B. "Liste_Jobcenter_Leistung")
    anhand einer Excel-Liste der aktuellen Sachbearbeiter.

.DESCRIPTION
    Die Excel-Datei ist die *einzige Quelle der Wahrheit* fuer den aktuellen
    Personalstand. Das Skript liest die Liste, matcht die Personen gegen das
    Active Directory (primaer ueber givenName + sn, mit Fallback-Strategien),
    und gleicht die *direkten* Benutzer-Mitglieder der Gruppe ab:
      - fehlende Personen  -> werden hinzugefuegt
      - ueberzaehlige User -> werden entfernt
    Verschachtelte Gruppen (nested groups) werden NIE angefasst.

    Das Skript unterstuetzt einen Dry-Run (-WhatIf) und erzeugt optional
    einen visuellen HTML-Report (Vorher/Nachher, wer kommt/geht).

.PARAMETER ExcelPath
    Pfad zur .xlsx-Datei mit den Sachbearbeitern.

.PARAMETER GroupName
    Name der Ziel-AD-Gruppe. Standard: "Liste_Jobcenter_Leistung".

.PARAMETER Server
    Optionaler Domain Controller / Domaenenname (z. B. "kreis-meissen.de").

.PARAMETER LogPath
    Optionaler Pfad zur Logdatei. Standard: neben dem Skript mit Zeitstempel.

.PARAMETER HtmlReportPath
    Optionaler Pfad fuer den visuellen HTML-Report (Vorher/Nachher).

.PARAMETER WhatIf
    Dry-Run: zeigt nur, was passieren wuerde (keine Aenderung im AD).

.EXAMPLE
    # 1) Trockenlauf inkl. HTML-Vorschau (nichts wird geaendert)
    .\Sync-JobcenterGroup.ps1 -ExcelPath .\Source\Liste_SachbearbeiterLeistung.xlsx `
        -HtmlReportPath .\Report.html -WhatIf

.EXAMPLE
    # 2) Echte Synchronisation
    .\Sync-JobcenterGroup.ps1 -ExcelPath .\Source\Liste_SachbearbeiterLeistung.xlsx

.EXAMPLE
    # 3) Andere Jobcenter-Gruppe (leicht erweiterbar)
    .\Sync-JobcenterGroup.ps1 -ExcelPath .\Source\Liste_Vermittlung.xlsx `
        -GroupName "Liste_Jobcenter_Vermittlung"

.NOTES
    ANLEITUNG / VORAUSSETZUNGEN
    ---------------------------
    * PowerShell 5.1+ (Windows). Ausfuehrung mit einem Konto, das die Gruppe
      aendern darf.
    * Modul "ActiveDirectory" (Teil der RSAT) muss installiert sein.
    * Zum Lesen der Excel-Datei wird EINES von beiden benoetigt:
        a) Modul "ImportExcel" (empfohlen, KEIN Excel noetig):
             Install-Module ImportExcel -Scope CurrentUser
        b) ODER lokal installiertes Microsoft Excel (COM-Fallback).
    * Empfehlung: IMMER zuerst mit -WhatIf und -HtmlReportPath laufen lassen,
      Report pruefen, danach ohne -WhatIf ausfuehren.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory = $true)]
    [string]$ExcelPath,

    [Parameter(Mandatory = $false)]
    [string]$GroupName = 'Liste_Jobcenter_Leistung',

    [Parameter(Mandatory = $false)]
    [string]$Server,

    [Parameter(Mandatory = $false)]
    [string]$LogPath,

    [Parameter(Mandatory = $false)]
    [string]$HtmlReportPath
)

# --------------------------------------------------------------------------
# 0) Grundgeruest: Fehler sollen hart stoppen, damit wir sie sauber fangen
# --------------------------------------------------------------------------
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

# Standard-Logpfad, falls keiner uebergeben wurde
if (-not $LogPath) {
    $stamp   = Get-Date -Format 'yyyyMMdd_HHmmss'
    $baseDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
    $LogPath = Join-Path $baseDir "Sync-Jobcenter_$stamp.log"
}

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
# 1) Excel einlesen -> Liste von Personen (nur Zeilen MIT Vor-/Nachname).
#    Leere Namenszeilen gehoeren zum vorherigen MA (weitere Zustaendigkeits-
#    bereiche) und stellen KEINE neue Person dar -> werden uebersprungen.
# --------------------------------------------------------------------------
function Import-SachbearbeiterListe {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Excel-Datei nicht gefunden: $Path"
    }

    $rows = $null

    # Bevorzugt: ImportExcel (kein Excel/COM noetig)
    if (Get-Module -ListAvailable -Name ImportExcel) {
        Import-Module ImportExcel -ErrorAction Stop
        Write-Log "Lese Excel via Modul 'ImportExcel'." 'INFO'
        $rows = Import-Excel -Path $Path
    }
    else {
        # Fallback: Excel-COM. Nur nutzbar, wenn Excel lokal installiert ist.
        Write-Log "Modul 'ImportExcel' nicht gefunden -> versuche Excel-COM-Fallback." 'WARN'
        $rows = Import-ExcelViaCom -Path $Path
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
    param([string]$Path)

    $excel = $null; $wb = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $wb = $excel.Workbooks.Open((Resolve-Path $Path).Path, $null, $true) # ReadOnly
        $ws = $wb.Worksheets.Item(1)
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
        [object[]]$Ambiguous, [object[]]$NotFound
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

    $mode = if ($DryRun) { "DRY-RUN (nichts wurde geaendert)" } else { "LIVE (Aenderungen wurden angewendet)" }
    $now  = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

    $html = @"
<!DOCTYPE html><html lang='de'><head><meta charset='utf-8'>
<title>Sync-Report $GroupName</title>
<style>
 body{font-family:Segoe UI,Arial,sans-serif;margin:24px;color:#1f2933;background:#f5f7fa}
 h1{font-size:20px} h2{font-size:15px;margin-top:24px}
 .mode{display:inline-block;padding:4px 10px;border-radius:6px;background:#e3f2fd;color:#0b5cad;font-weight:600}
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
<footer>Automatisch generiert von Sync-JobcenterGroup.ps1</footer>
</body></html>
"@

    $html | Out-File -FilePath $Path -Encoding UTF8
    Write-Log "HTML-Report geschrieben: $Path" 'OK'
}

# ==========================================================================
#  HAUPTPROGRAMM
# ==========================================================================
try {
    Write-Log "=== Start Sync fuer Gruppe '$GroupName' ===" 'INFO'
    Write-Log "Excel: $ExcelPath" 'INFO'
    if ($WhatIfPreference) { Write-Log "MODUS: DRY-RUN (-WhatIf) - es wird nichts geaendert." 'WARN' }

    # AD-Modul laden
    Import-Module ActiveDirectory -ErrorAction Stop

    # Optionalen Server-Parameter fuer alle AD-Cmdlets vorbereiten
    $adParams = @{}
    if ($Server) { $adParams['Server'] = $Server }

    # --- Excel einlesen ---
    $persons = Import-SachbearbeiterListe -Path $ExcelPath
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

    # --- Aenderungen anwenden (respektiert -WhatIf ueber ShouldProcess) ---
    foreach ($m in $toAdd) {
        $u = $m.User
        if ($PSCmdlet.ShouldProcess("$($u.SamAccountName) ($($m.Person.Anzeige))", "Zu Gruppe '$GroupName' HINZUFUEGEN")) {
            Add-ADGroupMember -Identity $GroupName -Members $u.DistinguishedName @adParams
            Write-Log ("HINZUGEFUEGT: {0} ({1})" -f $u.SamAccountName, $m.Person.Anzeige) 'ADD'
        } else {
            Write-Log ("[WhatIf] wuerde hinzufuegen: {0} ({1})" -f $u.SamAccountName, $m.Person.Anzeige) 'ADD'
        }
    }

    foreach ($u in $toRemove) {
        if ($PSCmdlet.ShouldProcess("$($u.SamAccountName) ($($u.Name))", "Aus Gruppe '$GroupName' ENTFERNEN")) {
            Remove-ADGroupMember -Identity $GroupName -Members $u.DistinguishedName -Confirm:$false @adParams
            Write-Log ("ENTFERNT: {0} ({1})" -f $u.SamAccountName, $u.Name) 'REMOVE'
        } else {
            Write-Log ("[WhatIf] wuerde entfernen: {0} ({1})" -f $u.SamAccountName, $u.Name) 'REMOVE'
        }
    }

    # --- Optionaler HTML-Report ---
    if ($HtmlReportPath) {
        New-HtmlReport -Path $HtmlReportPath -GroupName $GroupName -DryRun ([bool]$WhatIfPreference) `
            -Keep $toKeep.ToArray() -ToAdd $toAdd.ToArray() -ToRemove $toRemove.ToArray() `
            -Ambiguous $ambiguous.ToArray() -NotFound $notFound.ToArray()
    }

    # --- Zusammenfassung ---
    Write-Log "==================== ZUSAMMENFASSUNG ====================" 'INFO'
    Write-Log ("Excel-Personen gesamt : {0}" -f $persons.Count)      'INFO'
    Write-Log ("Eindeutig gematcht    : {0}" -f $matched.Count)      'OK'
    Write-Log ("Hinzugefuegt          : {0}" -f $toAdd.Count)        'ADD'
    Write-Log ("Entfernt              : {0}" -f $toRemove.Count)     'REMOVE'
    Write-Log ("Unveraendert (bleibt) : {0}" -f $toKeep.Count)       'INFO'
    Write-Log ("Mehrdeutig            : {0}" -f $ambiguous.Count)    'WARN'
    Write-Log ("Nicht gefunden        : {0}" -f $notFound.Count)     'WARN'
    Write-Log "=========================================================" 'INFO'
    if ($WhatIfPreference) { Write-Log "DRY-RUN beendet - es wurde NICHTS geaendert." 'WARN' }
    Write-Log "Logdatei: $LogPath" 'INFO'
    Write-Log "=== Ende ===" 'OK'
}
catch {
    Write-Log ("ABBRUCH: {0}" -f $_.Exception.Message) 'ERROR'
    Write-Log ($_.ScriptStackTrace) 'ERROR'
    exit 1
}
