# Jobcenter – AD-Abgleich

Zwei Wege, je nachdem ob du eine Live-Verbindung zum Active Directory hast:

| Skript | Wann | AD-Zugriff |
|--------|------|-----------|
| **`Abgleich-SachbearbeiterAD.ps1`** | Du hast **kein** RSAT/AD am Rechner. Du lieferst Excel + AD-Export als Datei, bekommst eine Ergebnisliste. | **keiner** (rein dateibasiert) |
| **`Sync-JobcenterGroup.ps1`** | Du hast RSAT/AD und willst die Gruppe direkt abgleichen (Analyse → Apply-Skript). | live |

> Wenn dein Rechner das Modul `ActiveDirectory` nicht laden kann
> (`... Modul "ActiveDirectory" wurde nicht geladen`), nimm **`Abgleich-SachbearbeiterAD.ps1`**.

---

## `Abgleich-SachbearbeiterAD.ps1` – Offline-Abgleich (kein AD nötig)

Du gibst zwei Dateien rein, bekommst **eine Ergebnisliste** heraus, deren Lücken du
noch manuell füllst.

**Eingaben**
1. Sachbearbeiter-Excel mit Spalten `Vorname`, `Nachname` (Quelle der Wahrheit).
2. AD-Export (**CSV oder XLSX**) aller AD-Benutzer mit `Vorname`, `Nachname`,
   `sAMAccountName`. Spaltennamen werden automatisch erkannt (Aliase wie
   `GivenName`/`Surname`/`SamAccountName` etc.), notfalls per Parameter überschreibbar.

**Ausgabe** – eine CSV (und, falls `ImportExcel` vorhanden, zusätzlich XLSX):

| Vorname | Nachname | Status | sAMAccountName | Kandidaten | Hinweis |
|---------|----------|--------|----------------|------------|---------|
| … | … | `GEFUNDEN` | `erika.mustermann` | | |
| … | … | `MEHRDEUTIG` | *(leer)* | `thomas.mueller \| thomas.mueller2` | bitte wählen |
| … | … | `NICHT_GEFUNDEN` | *(leer)* | | manuell ergänzen |

Die Zeilen mit leerem `sAMAccountName` sind genau die, die du noch ergänzt.
Offene Fälle stehen oben (Sortierung nach Status).

**Aufruf**

```powershell
.\Abgleich-SachbearbeiterAD.ps1 `
    -ExcelPath  .\Source\Liste_SachbearbeiterLeistung.xlsx `
    -AdListPath .\Source\AD_Export.csv
```

**Sofort testen** (mit mitgelieferten Beispieldateien, kein AD nötig):

```powershell
.\Abgleich-SachbearbeiterAD.ps1 `
    -ExcelPath  .\Source\Beispiel_SachbearbeiterListe.xlsx -WorksheetName "Leistung" `
    -AdListPath .\Source\Beispiel_AD_Export.csv
```

Erwartetes Ergebnis der Beispieldaten: **8 gefunden, 1 mehrdeutig** (Thomas Müller,
zwei Konten), **1 nicht gefunden** (Johanna Zäunig).

**Wie komme ich an den AD-Export?** Wenn jemand mit AD-Zugriff kurz Folgendes laufen
lässt (einmalig, nur Lesen), bekommst du die passende CSV:

```powershell
Get-ADUser -Filter 'Enabled -eq $true' -Properties GivenName,Surname,DisplayName |
    Select-Object GivenName,Surname,SamAccountName,DisplayName |
    Export-Csv .\AD_Export.csv -NoTypeInformation -Encoding UTF8 -Delimiter ';'
```

---

## `Sync-JobcenterGroup.ps1` – Live-Sync (mit AD/RSAT)

Synchronisiert eine AD-Sicherheitsgruppe (z. B. **`Liste_Jobcenter_Leistung`**) anhand
einer Excel-Liste der aktuellen Sachbearbeiter. Die Excel-Datei ist die **einzige
Quelle der Wahrheit** für den Personalstand.

Das Werkzeug arbeitet bewusst in **zwei getrennten Phasen**, damit niemals
versehentlich etwas im Active Directory geändert wird:

1. **Analyse** (nur lesend) → erzeugt Report + fertiges Ausführungsskript.
2. **Freigabe & Ausführung** (durch einen Menschen) → erst hier werden Namen
   ergänzt/entfernt.

---

## Voraussetzungen

| Was | Details |
|-----|---------|
| **PowerShell** | 5.1+ (Windows) |
| **RSAT / Modul `ActiveDirectory`** | Pflicht. Prüfen: `Get-Module -ListAvailable ActiveDirectory` |
| **Excel lesen** | Entweder Modul **`ImportExcel`** (empfohlen, kein Excel nötig) **oder** lokal installiertes Microsoft Excel (COM-Fallback) |
| **Berechtigung** | Das ausführende Konto muss Mitglieder der Zielgruppe ändern dürfen |

`ImportExcel` einmalig installieren:

```powershell
Install-Module ImportExcel -Scope CurrentUser
```

**RSAT / `ActiveDirectory`-Modul installieren** (falls `Das ... Modul "ActiveDirectory"
wurde nicht geladen` erscheint):

```powershell
# Windows 10/11 (als Administrator):
Add-WindowsCapability -Online -Name "Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0"

# Windows Server:
Install-WindowsFeature -Name RSAT-AD-PowerShell
```

Zum reinen Ausprobieren **ohne** AD/RSAT gibt es den Offline-Demomodus
`-DemoNoAd` (siehe unten).

---

## Ablauf

### Phase 1 – Analyse (ändert NICHTS)

```powershell
.\Sync-JobcenterGroup.ps1 -ExcelPath .\Source\Liste_SachbearbeiterLeistung.xlsx
```

Erzeugt im Ordner **`Sync-Output\`**:

| Datei | Zweck |
|-------|-------|
| `Report_*.html` | Visuelle Vorher/Nachher-Übersicht (dazu / raus / bleibt / mehrdeutig / kein Treffer) |
| `Apply_*.ps1` | **Generiertes Ausführungsskript** – eine kommentierte Zeile je Änderung |
| `Analyse_*.csv` | Dieselben Daten als CSV (Ablage / Weitergabe an den Fachbereich) |
| `Sync_*.log` | Vollständiges Protokoll des Laufs |

### Phase 2 – Prüfen & Ausführen

1. `Report_*.html` öffnen und prüfen.
2. `Apply_*.ps1` öffnen und die Zeilen kontrollieren („absegnen“).
3. Ausführen:

```powershell
# optionaler letzter Trockenlauf – ändert nichts:
.\Sync-Output\Apply_Liste_Jobcenter_Leistung_JJJJMMTT_HHMMSS.ps1 -WhatIf

# echte Änderung – fragt vorher "ja/nein":
.\Sync-Output\Apply_Liste_Jobcenter_Leistung_JJJJMMTT_HHMMSS.ps1

# echte Änderung ohne Rückfrage (z. B. Automatisierung):
.\Sync-Output\Apply_Liste_Jobcenter_Leistung_JJJJMMTT_HHMMSS.ps1 -Force
```

Nur das **Apply-Skript** verändert das AD.

---

## Parameter (Analyse-Skript)

| Parameter | Pflicht | Beschreibung |
|-----------|:-------:|--------------|
| `-ExcelPath` | ✔ | Pfad zur `.xlsx` |
| `-GroupName` | | Ziel-AD-Gruppe. Standard: `Liste_Jobcenter_Leistung` |
| `-WorksheetName` | | Arbeitsblatt in der xlsx. Fehlt der Parameter und gibt es mehrere Blätter, erscheint ein Auswahlmenü. **Für unbeaufsichtigte Läufe immer setzen.** |
| `-Server` | | Domain Controller / Domäne, z. B. `kreis-meissen.de` |
| `-OutputDir` | | Zielordner für Report/CSV/Apply/Log. Standard: `Sync-Output\` neben dem Skript |
| `-LogPath`, `-HtmlReportPath`, `-ApplyScriptPath`, `-CsvPath` | | Optional einzelne Ausgabepfade überschreiben |
| `-DemoNoAd` | | Offline-Demomodus ohne AD/RSAT (AD wird simuliert). Nur zum Ausprobieren, nicht produktiv. |

Volle Hilfe: `Get-Help .\Sync-JobcenterGroup.ps1 -Full`

---

## Matching-Logik (Kurzfassung)

Die Excel enthält **keinen** Windows-Benutzernamen, daher wird über Namen gematcht:

1. **Primär:** `givenName` + `sn`, jeweils mit Varianten
   (z. B. `Leslie Jenny` → auch `Leslie`; `Noack genannt Gräfe` → auch `Noack` / `Gräfe`).
2. **Fallback:** `displayName` / `cn` (auch „Nachname Vorname“).

Namen werden vorher normalisiert (Umlaute/ß aufgelöst: `Grünberg`→`gruenberg`,
`Jendroßek`→`jendrossek`; Bindestriche/Sonderzeichen vereinheitlicht).

- **1 Treffer** → wird verarbeitet.
- **Mehrere Treffer** → *mehrdeutig*: **nicht** automatisiert, nur zur Prüfung gelistet.
- **0 Treffer** → *kein Treffer*: ebenfalls nur zur Prüfung gelistet.

Berücksichtigt werden **nur aktive** (`Enabled`) Benutzer. Beim Abgleich werden
ausschließlich **direkte Benutzer-Mitglieder** betrachtet – verschachtelte Gruppen
werden nie angefasst.

---

## Erweiterbar auf weitere `Liste_Jobcenter_*`-Gruppen

Kein Code-Umbau nötig – nur andere Parameter:

```powershell
.\Sync-JobcenterGroup.ps1 -ExcelPath .\Source\Liste_Vermittlung.xlsx `
    -GroupName "Liste_Jobcenter_Vermittlung" -WorksheetName "Vermittlung"
```

---

## Gefahrlos testen

Die mitgelieferte Beispiel-Datei **`Source/Beispiel_SachbearbeiterListe.xlsx`**
(3 Blätter, erfundene Namen, alle Sonderfälle) eignet sich zum Durchspielen der
kompletten Pipeline:

```powershell
.\Sync-JobcenterGroup.ps1 -ExcelPath .\Source\Beispiel_SachbearbeiterListe.xlsx -WorksheetName "Leistung"
.\Sync-Output\Apply_*.ps1 -WhatIf
```

> Die erfundenen Namen erscheinen gegen ein echtes AD als „kein Treffer“ – das ist
> gewollt. Für einen echten Match-Test einige Namen durch reale Test-Benutzer ersetzen.

### Ganz ohne AD/RSAT testen (`-DemoNoAd`)

Wenn (noch) kein `ActiveDirectory`-Modul installiert ist, lässt sich die komplette
Pipeline offline durchspielen. Das AD wird dann **simuliert** (synthetischer Stand
aus der Excel), damit Report + Apply-Skript entstehen und angesehen werden können:

```powershell
.\Sync-JobcenterGroup.ps1 -ExcelPath .\Source\Beispiel_SachbearbeiterListe.xlsx `
    -WorksheetName "Leistung" -DemoNoAd
```

> **Nur zum Ausprobieren.** Das im Demomodus erzeugte `Apply_*.ps1` enthält
> synthetische DNs und trägt einen deutlichen Warnhinweis – niemals gegen ein
> echtes AD ausführen.

---

## Troubleshooting

| Symptom | Ursache / Lösung |
|---------|------------------|
| `Import-Excel : ... nicht erkannt` | Modul fehlt → `Install-Module ImportExcel -Scope CurrentUser` (oder Excel für COM-Fallback installieren) |
| `Get-ADGroup : ... not found` / kein AD | RSAT/`ActiveDirectory`-Modul fehlt, oder `-Server` angeben |
| Skript „hängt“ nach dem Start | Interaktives Blatt-Menü wartet auf Eingabe → `-WorksheetName` setzen |
| Viele „kein Treffer“ | Excel-Namen weichen von AD ab, oder falsches Blatt gewählt → Report/CSV prüfen |
| `Arbeitsblatt '…' nicht gefunden` | Blattname stimmt nicht → das Skript listet die vorhandenen Blätter in der Fehlermeldung |

---

## Verbesserungsideen

- **Eindeutiger Schlüssel** (Personalnummer → `employeeID`, oder Dienst-E-Mail →
  `mail`/`userPrincipalName`) als primäre Strategie ⇒ Mehrdeutigkeiten entfallen.
- **Sicherheitsnetz**: Abbruch/Rückfrage, wenn ein Lauf mehr als X % der Gruppe
  entfernen würde.
- **Fuzzy-Scoring** (Levenshtein) als dritte Stufe für Tippfehler.
- **Scheduling**: täglicher Analyse-Lauf, Diff-Report per Mail an den Fachbereich,
  Live-Lauf erst nach Freigabe.
