const express = require('express');
const cors = require('cors');
const bodyParser = require('body-parser');
const ExcelJS = require('exceljs');

const app = express(); // 👈 app initialization first
app.use(cors());
app.use(bodyParser.json());
const cors = require('cors');
app.use(cors());
const express = require('express');
const bodyParser = require('body-parser');
const ExcelJS = require('exceljs');

const app = express();
app.use(bodyParser.json());

// Auto-calc route
app.post('/auto-calc', (req, res) => {
  const manjur = parseFloat(req.body.manjur);
  const rate = 500; // Rs per unit area
  const depth = 0.2; // default depth

  const area = manjur / rate;
  const length = Math.sqrt(area).toFixed(2);
  const width = Math.sqrt(area).toFixed(2);

  res.json({ length, width, depth });
});

// Generate Excel route
app.post('/generate', async (req, res) => {
  try {
    const data = req.body;

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('skeleton.xlsx'); // load skeleton workbook

    const sheet = workbook.getWorksheet('Abstract');
    sheet.getCell('B2').value = data.yojna;
    sheet.getCell('B3').value = data.kam;
    sheet.getCell('B4').value = data.manjur;
    sheet.getCell('B5').value = data.aae;
    sheet.getCell('B6').value = data.ss;
    sheet.getCell('B7').value = data.srno;
    sheet.getCell('B8').value = data.length;
    sheet.getCell('B9').value = data.width;
    sheet.getCell('B10').value = data.depth;
    sheet.getCell('B11').value = data.total;

    const filename = `estimate_${Date.now()}.xlsx`;
    await workbook.xlsx.writeFile(filename);

    res.json({ file: filename, url: `/download/${filename}` });
  } catch (err) {
    res.status(500).json({ error: 'Failed to generate estimate' });
  }
});

// Route to serve generated files
app.get('/download/:filename', (req, res) => {
  const filename = req.params.filename;
  res.download(filename);
});

// Start server
app.listen(process.env.PORT || 3000, () => {
  console.log("✅ Server running");
});
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
