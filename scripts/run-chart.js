import { generateAusgabenChartPlaywright } from '../lib/chart-playwright.js';

(async () => {
  try {
    const res = await generateAusgabenChartPlaywright(undefined, { headless: false });
    console.log('Result:', res);
  } catch (e) {
    console.error('Error running chart-playwright:', e);
    process.exit(1);
  }
})();
