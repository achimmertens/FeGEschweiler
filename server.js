import express from 'express';
import path from 'path';
import { fileURLToPath } from 'url';
import { chromium } from 'playwright';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const port = 3000;

// Add CORS headers
app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Methods', 'GET, OPTIONS');
  res.header('Access-Control-Allow-Headers', 'Content-Type');
  next();
});

// Route to generate PDF
app.get('/generate-pdf/:year', async (req, res) => {
  const year = req.params.year;
  const htmlPath = path.join(__dirname, 'Daten', 'result', `JahresabschlussUnterlagen_${year}.html`);

  try {
    const browser = await chromium.launch();
    const page = await browser.newPage();
    // Use http://localhost:3000 to load the page
    await page.goto(`http://localhost:3000/JahresabschlussUnterlagen_${year}.html`);
    const pdfBuffer = await page.pdf({
      format: 'A4',
      printBackground: true,
      margin: { top: '1cm', right: '1cm', bottom: '1cm', left: '1cm' }
    });
    await browser.close();

    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `attachment; filename=JahresabschlussUnterlagen_${year}.pdf`);
    res.send(pdfBuffer);
  } catch (error) {
    console.error('Error generating PDF:', error);
    res.status(500).send('Error generating PDF: ' + error.message);
  }
});

// Serve static files from Daten/result
app.use(express.static(path.join(__dirname, 'Daten', 'result')));

app.listen(port, () => {
  console.log(`Server running at http://localhost:${port}`);
  console.log(`Open http://localhost:${port}/JahresabschlussUnterlagen_2025.html in your browser`);
});