import fs from 'fs';
import path from 'path';

export function findLatestAusgabenJson(dir) {
  if (!fs.existsSync(dir)) return null;
  const files = fs.readdirSync(dir);
  const re = /^(Ausgaben|Einnahmen)_(\d{4})\.json$/;
  let best = null;
  for (const f of files) {
    const m = f.match(re);
    if (m) {
      const year = Number(m[2]);
      if (!best || year > best.year) best = { file: f, year };
    }
  }
  return best;
}
