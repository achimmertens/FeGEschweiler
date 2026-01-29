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



