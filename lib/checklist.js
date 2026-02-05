import fs from 'fs';
import path from 'path';

function logDebug(msg) {
  try { fs.appendFileSync(path.join(process.cwd(), 'debug.log'), `[checklist] ${msg}\n`); } catch (e) {}
}

export function generateDefaultChecklist(currentYear) {
  try {
    const checklist = [
      'Neue Gewinn-Verlust-Berichte (gewinn-verlust-bericht_YYYY.csv) in /Daten ablegen',
      'Neue Bilanzberichte (bilanzbericht_YYYY.csv) in /Daten ablegen',
      'Budgets_YYYY.txt prüfen / aktualisieren und in /Daten ablegen',
      'node generate-presentation.js ausführen und Budget_yyyy.csv nach Daten/Budget_yyyy+1.csv kopieren',
      'Budgets für das laufende Jahr dort eintragen',
      'node generate-presentation.js ausführen und debug.log lesen',
      'Website/Index prüfen: /Daten/result/index.html öffnet alle Seiten',
      'Repository commit & push: git add -A && git commit -m "update data YYYY" && git push'
    ];

    const checklistHtml = `<!doctype html><html lang="de"><head><meta charset="utf-8"/><meta name="viewport" content="width=device-width,initial-scale=1"/><title>Checkliste</title><style>body{font-family:Arial,sans-serif;margin:18px}h1{color:#667eea}ol{font-size:16px;line-height:1.6}</style></head><body><h1>Checkliste: Jährliche Datenaktualisierung</h1><ol>${checklist.map(i=>`<li>${i}</li>`).join('')}</ol></body></html>`;
    const outPath = path.join(process.cwd(), 'Daten', 'result', 'Checkliste.html');
    fs.writeFileSync(outPath, checklistHtml, 'utf8');
    logDebug(`Checkliste HTML erstellt: ${outPath}`);
    return outPath;
  } catch (e) {
    logDebug('Fehler beim Erzeugen Checkliste-Seite: ' + (e && e.message ? e.message : String(e)));
    throw e;
  }
}
