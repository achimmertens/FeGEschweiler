import * as ausgabenMod from '../lib/generate-ausgaben-chart.js';
import * as einnahmenMod from '../lib/generate-einnahmen-chart.js';

(async () => {
  try {
    const resA = await ausgabenMod.generateAusgabenChartPlaywright(undefined, { headless: false });
    console.log('Ausgaben Result:', resA);
    const resE = await einnahmenMod.generateEinnahmenChartPlaywright(undefined, { headless: false });
    console.log('Einnahmen Result:', resE);
  } catch (e) {
    console.error('Error running charts:', e);
    process.exit(1);
  }
})();
