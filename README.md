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


----------------------------------------------------

# Wie man die Präsentation anhand dieses Tools jährlich erzeugt

Zuerst hole die Daten von "Bilanzbericht", "Gewinn-verlust-Bericht" und speichere sie als csv Datei für das jeweilige Jahr im Ordner Daten.
Die Budgets.txt erstellt man durch Markieren und einem Copy-Befehl über die Seite "Bugdets".
Trage die Anzahl der Gemeindemitglieder in der Datei Daten/Mitgliederzahl.txt korrekt eingetragen ist, damit sie im PDF angezeigt wird. (Details dazu stehen in der Datei.)

Danach startet man das Script mit:
npm run start

Überprüfe die letzte Spalte in Entwicklung.csv und passe sie evtl. an. Starte npm run start erneut und überprüfe die Entwicklung.csv abermals.

# Schritte zum Starten der Webseite:

Generiere die HTML-Dateien (falls noch nicht geschehen):

- Führe npm run start aus (das läuft node generate-presentation.js und erstellt die Dateien in result).

## Starte einen lokalen Webserver:

Navigiere in den Ordner result:
Verwende einen einfachen HTTP-Server. Da Node.js installiert ist, kannst du npx nutzen:
> npx http-server . -p 8080

Das startet einen Server auf Port 8080.

Öffne die Webseite:
http://localhost:8080/Daten/result/index.html

Die Hauptseite ist index.html im Daten/result-Ordner.
Sie enthält Navigation zu den generierten Berichten (z. B. Ausgaben, Einnahmen, Budget usw.).

Kopiere die relevanten Inhalte der einzelnen Seiten in die neu zu erstellende Datei Kassenbericht_yyyy.odt.

# Jahresunterlagen PDF für die SKB Bank
Das PDF wird automatisch mit "npm run" mit erstellt im Ordner `Daten/result` als `JahresabschlussUnterlagen_YYYY.pdf` (wobei YYYY das aktuelle Jahr ist). 

