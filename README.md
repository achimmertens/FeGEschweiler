# FeG Eschweiler — Generator für Ergebnisseiten und Präsentationen
#
# Autor: Achim Mertens
# Version 1.0 (2026-02-05)
#

Dieses Repository erzeugt aus CSV-Finanzdaten eine Reihe von Ergebnisseiten im Ordner `Daten/result` und (optional) eine PowerPoint-Präsentation.

Wichtig: README wurde aktualisiert — das Tool erstellt heute HTML/JSON/PNG-Ausgaben und einige Hilfsseiten. Die frühere automatische Excel-Erzeugung ist standardmäßig deaktiviert.

## Schneller Start

1. Node.js installieren (empfohlen >= 18)
2. Abhängigkeiten installieren:

```bash
npm install
```

3. Generierung starten:

```bash
node generate-presentation.js
```

Die Ausgaben landen in `Daten/result`.

## Was das Skript jetzt erzeugt

- HTML- und JSON-Ausgaben für Berichte (z.B. `Ausgaben_YYYY.html`, `Ausgaben_YYYY.json`)
- PNG-Vorschauen der Diagramme (wenn Playwright verfügbar)
- `index.html` in `Daten/result` mit Navigation und eingebetteten Seiten
- `Budget_YYYY.html` (Budget-Tabelle; die Plan-Spalte wird aus der letzten Spalte von `Budget_YYYY.csv` übernommen)
- `Sonderspenden.html` (aus `Daten/SonderspendenTermine.csv`)
- `Checkliste.html` (eine einfache jährliche Checkliste)

Außerdem werden Hilfsmodule verwendet:
- `lib/budget.js` — Budget-HTML-Generator
- `lib/sonderspenden.js` — Sonderspenden-Page-Generator
- `lib/checklist.js` — Checkliste-Generator

Wenn die Datei `Daten/Budget_<aktuellesJahr>.csv` existiert, wird sie beim Lauf automatisch nach `Daten/result/` kopiert und als Plan-Spalte verwendet.

## Hinweise zu Budget-HTML

- Die beiden Spalten `Verbraucht-Vorjahr` und `Geplant-Vorjahr` werden in der erzeugten HTML-Tabelle nicht angezeigt.
- Die Planwerte für das laufende Jahr werden aus der letzten Spalte der Budget-CSV übernommen.

## PPT / Excel Verhalten

- Die frühere automatische Erstellung von `Finanzlage_FeG_Eschweiler.xlsx` ist standardmäßig deaktiviert.
- Die PPT-Erzeugung wird nur ausgeführt, wenn ausreichend sortierte Daten (Einnahmen/Ausgaben/Pie) vorhanden sind. Andernfalls wird sie übersprungen.

Wenn du wieder eine Excel- oder PPT-Erzeugung aus Daten erzwingen willst, sag mir Bescheid — ich kann das wieder aktivieren und die Quellen (CSV/JSON) definieren.

# Was du jährlich tun musst, um die Präsentation zu erzeugen
## Erwartete Eingabedateien (Ordner `Daten/`)

- `gewinn-verlust-bericht_YYYY.csv` — Gewinn-/Verlustberichte
- `bilanzbericht_YYYY.csv` — Bilanzberichte
- `Entwicklung.csv` — historische Kontostände (wird aktualisiert)
- optional: `Budgets_YYYY.txt` oder `Budget_YYYY.csv` — Budgets
- optional: `SonderspendenTermine.csv` — Termine für Sonderspenden

## Überprüfe auch die Punkte in der Checklist
Siehe checklist.js — dort findest du eine einfache Schritt-für-Schritt-Anleitung, was du tun musst, um die Daten zu aktualisieren.
Ggf. musst du die Liste noch etwas anpassen.

## Ergebnisordner (`Daten/result`)

Nach Ausführung findest du dort mindestens:
- `index.html` — Hauptseite mit Navigation
- `Ausgaben_YYYY.html`, `Einnahmen_YYYY.html`, `Entwicklung_YYYY.html`
- `Budget_<aktuellesJahr>.html` und `Budget_<aktuellesJahr>.csv` (falls vorhanden)
- `Sonderspenden.html`, `Checkliste.html`
- ggf. `*.png` Vorschauen und `*.json` für Diagrammdaten

## Debugging

- Logs: `debug.log` im Projekt-Root enthält Laufinformationen.
- Wenn etwas fehlt, starte `node generate-presentation.js` und prüfe das Log auf Hinweise.

## Weiteres / Anpassungen

- Die HTML-Vorlagen und das Index-Menü werden in `generate-presentation.js` erzeugt — kleine Anpassungen (Labels, Reihenfolge) lassen sich dort schnell ändern.
- Wenn du interaktive Checklisten, sortierte/formatierte Sonderspenden-Tabellen oder weitere Exporte brauchst, implementiere ich das gern.

---

Wenn du möchtest, schreibe ich noch eine kurze Anleitung, welche Dateien du jährlich aktualisieren musst — oder mache die Checkliste interaktiv.
# FeG Eschweiler - Finanzpräsentation Generator

Dieses Projekt generiert automatisch eine PowerPoint-Präsentation zur Finanzlage der Freien evangelischen Gemeinde Eschweiler.

## Installation

1. Stelle sicher, dass Node.js installiert ist (Version 14 oder höher)
2. Installiere die Abhängigkeiten:

```bash
npm install
```

## Verwendung

Führe das Skript aus, um die Präsentation zu generieren:

```bash
npm start
```

oder

```bash
node generate-presentation.js
```

Das Skript:
1. Liest alle Gewinn-Verlust-Berichte aus dem Ordner `Daten/`
2. Liest alle Bilanzberichte aus dem Ordner `Daten/`
3. Aktualisiert die Datei `Entwicklung.csv` mit den neuesten Kontoständen
4. **Erstellt eine Excel-Datei** `Finanzlage_FeG_Eschweiler.xlsx` mit 4 Reitern:
   - **Einnahmen**: Sortierte Tabelle nach Gesamtsumme (absteigend)
   - **Ausgaben**: Sortierte Tabelle nach Gesamtsumme (absteigend)
   - **Ausgaben Kuchendiagramm**: Sortierte Tabelle nach Betrag (absteigend) mit Anteil in %
   - **Entwicklung**: Tabelle mit Kontoständen über die Jahre
5. Erstellt eine PowerPoint-Präsentation `Finanzlage_FeG_Eschweiler.pptx` mit 4 Slides:
   - **Einnahmen**: Gestapeltes Balkendiagramm über die Jahre (sortiert nach Gesamtsumme)
   - **Ausgaben**: Gestapeltes Balkendiagramm über die Jahre (sortiert nach Gesamtsumme)
   - **Ausgaben nach Kategorien**: Kuchendiagramm für das neueste Jahr (sortiert nach Betrag)
   - **Entwicklung der Kontostände**: Liniendiagramm und Tabelle der Kontostände über die Jahre

## Datenstruktur

Das Skript erwartet folgende Dateien im Ordner `Daten/`:

- `gewinn-verlust-bericht_YYYY.csv` - Gewinn-Verlust-Berichte für die Jahre YYYY
- `bilanzbericht_YYYY.csv` - Bilanzberichte für die Jahre YYYY
- `Entwicklung.csv` - Historische Entwicklung der Kontostände

## Automatische Aktualisierung

Das Skript kann jederzeit neu ausgeführt werden. Es:
- Erkennt automatisch neue Jahre in den Dateinamen
- Aktualisiert die `Entwicklung.csv` mit den neuesten Daten
- Generiert die Präsentation mit allen verfügbaren Daten neu

## Design

Die Präsentation verwendet ein professionelles Design mit:
- Blauen und grünen Farbtönen
- Klaren Diagrammen und Tabellen
- Gut lesbaren Schriftgrößen

## Abhängigkeiten

- `pptxgenjs` - PowerPoint-Generierung
- `exceljs` - Excel-Dateien erstellen
- `csv-parse` - CSV-Dateien lesen
- `csv-stringify` - CSV-Dateien schreiben

## Sortierung

Alle Tabellen werden automatisch nach der Größe der Ein- und Ausgaben sortiert:
- **Einnahmen**: Nach Gesamtsumme über alle Jahre (absteigend)
- **Ausgaben**: Nach Gesamtsumme über alle Jahre (absteigend)
- **Kuchendiagramm**: Nach Betrag des neuesten Jahres (absteigend)

Dies sorgt für eine übersichtlichere Darstellung der wichtigsten Posten.

----------------------------------------------------

# Wie man die Präsentation anhand dieses Tools jährlich erzeugt

Zuerst holt man sich die Daten von "Bilanzbericht", "Gewinn-verlust-Bericht" und speichert sie als csv Datei für das jeweilige Jahr im Ordner Daten.
Die Budgets.txt erstellt man durch Markieren und einem Copy-Befehl über die Seite "Bugdets".

Danach startet man das Script mit:
npm run start

# Präsentationsvorbereitung

Bitte kopiere die Excel Datei result/Finanzlage_FeG_Eschweiler.xlsx nach Kassenbericht_yyyy.ods und von dort weiter nach Documente/FeG.
Es ist sinnvoll, in der result/Kassenbericht_yyyy.ods weiter zu arbeiten, um zu vermeiden, dass bei einem nächsten Durchlauf diese Planung überschrieben wird.
Der Budget-Reiter aus dem Kassenbericht wird kopiert in die Datei Budget_yyyz.odp (wobei z das aktuelle Jahr, alys y+1 ist). Dort muss noch um die Planung für das neue Jahr erweitert werden.
Dieser Kassenbericht und das neue Budget dient dann wiederum als Vorlage um die Tabellen rauszukopieren, die dann in Powerpoint Datei Documente/FeG/Kassenbericht_yyyy.odp eingefügt werden. Dort kommen auch die Grafiken aus dem Result Ordner rein. Diese Datei wird dann für den Vortrag genutzt.
Kurz: 
- result/Finanzlage_FeG_Eschweiler.xlsx kopieren und umbenennen und kopieren -> Dokumente/FeG/Kassenbericht_yyyy.ods
- Budgets in result/Budgets_(aktuelles Jahr).ods und Dokumente/Feg/Kassenbericht.ods eintragen
- Budgets und resut/Bilder nach Documente/FeG/Kassenbericht_yyyy.odp überführen



