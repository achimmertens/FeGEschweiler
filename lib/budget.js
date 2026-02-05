import fs from 'fs';
import path from 'path';
import { readCSV } from './utils.js';

function logDebug(msg) {
  try { fs.appendFileSync(path.join(process.cwd(), 'debug.log'), `[budget] ${msg}\n`); } catch (e) {}
}

export async function generateBudgetHtml(outCsv, bestYear, currentYear) {
  try {
    const csvRows = readCSV(outCsv);
    let headers = csvRows && csvRows.length ? Object.keys(csvRows[0]) : [];

    // remove previous-year columns from display
    headers = headers.filter(h => h !== 'Verbraucht-Vorjahr' && h !== 'Geplant-Vorjahr');

    // Determine plan column: prefer last column of Budget_<currentYear>.csv if present
    const planCandidates = [];
    const planSrcA = path.join(process.cwd(), 'Daten', `Budget_${currentYear}.csv`);
    const planSrcB = path.join(process.cwd(), 'Daten', 'result', `Budget_${currentYear}.csv`);
    if (fs.existsSync(planSrcA)) planCandidates.push(planSrcA);
    if (fs.existsSync(planSrcB) && planSrcB !== outCsv) planCandidates.push(planSrcB);

    let planHeader = null;
    const planMap = {};
    for (const p of planCandidates) {
      try {
        const prow = readCSV(p);
        if (!prow || !prow.length) continue;
        const pcols = Object.keys(prow[0] || {});
        if (!pcols || pcols.length === 0) continue;
        planHeader = pcols[pcols.length - 1];
        prow.forEach(r => {
          const key = (r['Nummer'] || r['Nummer'] === 0) ? String(r['Nummer']).trim() : String(r['Kostenstelle'] || '').trim();
          if (key) planMap[key] = r[planHeader];
        });
        if (planHeader) break;
      } catch (e) { /* ignore and try next */ }
    }

    const escapeHtml = s => String(s === undefined || s === null ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

    const tableRows = (csvRows || []).map(r => {
      const base = headers.map(h => `<td>${escapeHtml(r[h]===undefined? '': r[h])}</td>`).join('');
      const key = (r['Nummer'] || r['Nummer'] === 0) ? String(r['Nummer']).trim() : String(r['Kostenstelle'] || '').trim();
      const planCell = planHeader ? `<td>${escapeHtml(planMap[key] !== undefined ? planMap[key] : '')}</td>` : '';
      return '<tr>' + base + planCell + '</tr>';
    }).join('\n');

    const finalHeaders = headers.slice();
    if (planHeader) finalHeaders.push(`Plan ${currentYear}`);

    const html = `<!doctype html>\n<html lang="de">\n<head>\n<meta charset="utf-8"/>\n<meta name="viewport" content="width=device-width,initial-scale=1"/>\n<title>Budget ${bestYear}</title>\n<style>body{font-family:Arial,sans-serif;margin:12px}table{border-collapse:collapse;width:100%}th,td{border:1px solid #ddd;padding:6px;text-align:left;font-size:13px}th{background:#f3f4f6}</style>\n</head>\n<body>\n<h1>Budget ${bestYear} + ${bestYear+1}</h1>\n<table>\n<thead>\n<tr>\n${finalHeaders.map(h=>`<th>${escapeHtml(h)}</th>`).join('\n')}\n</tr>\n</thead>\n<tbody>\n${tableRows}\n</tbody>\n</table>\n</body>\n</html>`;

    const outHtml = path.join(process.cwd(), 'Daten', 'result', `Budget_${currentYear}.html`);
    fs.writeFileSync(outHtml, html, 'utf8');
    logDebug(`Budget HTML erstellt: ${outHtml}`);
    return outHtml;
  } catch (e) {
    logDebug('Fehler beim Erzeugen Budget-HTML: ' + (e && e.message ? e.message : String(e)));
    throw e;
  }
}
