import fs from 'fs';
import path from 'path';
import { readCSV } from './utils.js';

function logDebug(msg) {
  try { fs.appendFileSync(path.join(process.cwd(), 'debug.log'), `[sonderspenden] ${msg}\n`); } catch (e) {}
}

export async function generateSonderspendenPage() {
  try {
    const sFile = path.join(process.cwd(), 'Daten', 'SonderspendenTermine.csv');
    if (!fs.existsSync(sFile)) return null;
    const rows = readCSV(sFile);
    const escapeHtml = s => String(s === undefined || s === null ? '' : s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
    const tableRows = (rows||[]).map(r => '<tr>' + Object.values(r).map(v => `<td>${escapeHtml(v)}</td>`).join('') + '</tr>').join('\n');
    const headers = rows && rows.length ? Object.keys(rows[0]).map(h => `<th>${escapeHtml(h)}</th>`).join('\n') : '';
    const sHtml = `<!doctype html><html lang="de"><head><meta charset="utf-8"/><meta name="viewport" content="width=device-width,initial-scale=1"/><title>Sonderspenden</title><style>body{font-family:Arial,sans-serif;margin:12px}table{border-collapse:collapse;width:100%}th,td{border:1px solid #ddd;padding:6px;text-align:left;font-size:13px}th{background:#f3f4f6}</style></head><body><h1>Sonderspenden Termine</h1><table><thead><tr>${headers}</tr></thead><tbody>${tableRows}</tbody></table></body></html>`;
    const sOut = path.join(process.cwd(), 'Daten', 'result', 'Sonderspenden.html');
    fs.writeFileSync(sOut, sHtml, 'utf8');
    logDebug(`Sonderspenden HTML erstellt: ${sOut}`);
    return sOut;
  } catch (e) {
    logDebug('Fehler beim Erzeugen Sonderspenden-Seite: ' + (e && e.message ? e.message : String(e)));
    throw e;
  }
}
