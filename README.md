---
created: 2026-07-24T11:18
updated: 2026-07-24T11:19
---
# Gruppen-Abgleich

Ein Werkzeug, das eine **Sachbearbeiter-Liste (Soll)** mit den **aktuellen
Mitgliedern einer AD-Gruppe (Ist)** vergleicht und zeigt, wer **hinzugefügt** und
wer **entfernt** werden muss.

Es ist eine einzelne HTML-Datei, die **komplett im Browser** läuft:
kein Active Directory, kein Server, keine Installation. **Es werden keine Daten
gesendet** – alles bleibt lokal (datenschutzfreundlich für Beschäftigtennamen).

## Dateien

| Datei | Zweck |
|-------|-------|
| `Gruppen-Abgleich.html` | Das Werkzeug. Standalone, per Doppelklick im Browser zu öffnen. |
| `Source/Liste_SachbearbeiterLeistung.xlsx` | Original-Sollliste (Referenz/Datenquelle). |

## Nutzung (Variante A – lokal)

1. `Gruppen-Abgleich.html` auf ein Netzlaufwerk legen und **doppelklicken**
   (öffnet im Standardbrowser).
2. **Oben (Soll):** die Sachbearbeiter-Tabelle einfügen – direkt aus der
   Quelle kopiert (mit Kopfzeile, Tabs, Mehrfachzeilen für
   Zuständigkeitsbereiche). Nur Zeilen mit Vor- und Nachname zählen.
3. **Unten (Ist):** die aktuelle Gruppenliste aus dem Benutzergruppen-Tool
   einfügen (Format „Nachname, Vorname“).
4. **„Vergleichen“** klicken → farbige Übersicht:
   **grün = hinzufügen**, **rot = entfernen**, plus „bleibt“ und „ignoriert“.
   Mit **„Namen kopieren“** die Add-/Remove-Liste übernehmen.

Über **„Beispiel laden“** lässt sich sofort ein kleiner Testfall ansehen.

## Was der Abgleich automatisch berücksichtigt

- Umlaute/ß (`Grünberg` ↔ `gruenberg`, `Jendroßek` ↔ `jendrossek`)
- „genannt“-Namen (`Noack genannt Gräfe`) und mehrteilige Vornamen (`Leslie Jenny`)
- Zähl-Dubletten der AD-Anzeige (`Schmidt1` → `Schmidt`)
- verschachtelte Gruppen / Funktionskonten (`Liste_*`, `JC.*` …) werden ignoriert

> **Hinweis:** rein namensbasierter Abgleich. Bei mehreren Änderungen lohnt eine
> kurze Sichtprüfung – taucht jemand zugleich bei „hinzufügen“ und „entfernen“
> auf, ist es meist eine abweichende Schreibweise.

## Weitere Verteilungswege (optional)

- **B – Intranet:** die HTML-Datei auf einen internen Webserver legen (feste URL,
  zentrale Updates). Weiterhin rein clientseitig.
- **C – Integration:** das bestehende Benutzergruppen-Tool könnte den Ist-Stand
  automatisch liefern, sodass nur noch die Soll-Liste eingefügt werden muss.
