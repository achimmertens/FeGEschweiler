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
