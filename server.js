const express = require('express');
const cors = require('cors');
const helmet = require('helmet');
const path = require('path');
const multer = require('multer');
const ExcelJS = require('exceljs');

const {
  toNumber,
  normalizeWeightToKg,
  parseDimensionsToLBH,
  normPartNo,
  mapMaterialClassificationToExcel,
  withinToleranceKG
} = require('./utils');
const { SiemensProductScraper, a2vUrl } = require('./scraper');

const app = express();
const PORT = process.env.PORT || 3000;
const SCRAPE_CONCURRENCY = Number(process.env.SCRAPE_CONCURRENCY || 6);
const WEIGHT_TOL_PCT = Number(process.env.WEIGHT_TOL_PCT || 0);
const COLS = { Z:'Z', E:'E', C:'C', S:'S', U:'U', V:'V', W:'W', P:'P', N:'N' };
const HEADER_ROW = 3;
const FIRST_DATA_ROW = 4;

app.use(helmet({ contentSecurityPolicy: false }));
app.use(cors());
app.use(express.json({ limit: '2mb' }));
app.use(express.static(__dirname));

const scraper = new SiemensProductScraper();

function statusCellFill(status) {
  if (status === 'GREEN') return { type:'pattern', pattern:'solid', fgColor:{ argb:'FFD5F4E6' } };
  if (status === 'RED')   return { type:'pattern', pattern:'solid', fgColor:{ argb:'FFFDEAEA' } };
  return { type:'pattern', pattern:'solid', fgColor:{ argb:'FFFFF3CD' } };
}

function compareText(excel, web) {
  if (!excel && !web) return { status:'ORANGE', comment:'Beide fehlen' };
  if (!excel) return { status:'ORANGE', comment:'Excel fehlt' };
  if (!web)   return { status:'ORANGE', comment:'Web fehlt' };
  const a = String(excel).trim().toLowerCase().replace(/\s+/g,' ');
  const b = String(web).trim().toLowerCase().replace(/\s+/g,' ');
  return a === b ? { status:'GREEN', comment:'identisch' } : { status:'RED', comment:'abweichend' };
}
function comparePartNo(excel, web) {
  if (!excel && !web) return { status:'ORANGE', comment:'Beide fehlen' };
  if (!excel) return { status:'ORANGE', comment:'Excel fehlt' };
  if (!web)   return { status:'ORANGE', comment:'Web fehlt' };
  return normPartNo(excel) === normPartNo(web)
    ? { status:'GREEN', comment:'identisch (normalisiert)' }
    : { status:'RED', comment:`abweichend: Excel ${excel} vs. Web ${web}` };
}
function compareWeight(excelVal, webVal) {
  const exKg = normalizeWeightToKg(excelVal);
  const wbKg = normalizeWeightToKg(webVal);
  if (exKg == null && wbKg == null) return { status:'ORANGE', comment:'Beide fehlen' };
  if (exKg == null) return { status:'ORANGE', comment:'Excel fehlt' };
  if (wbKg == null) return { status:'ORANGE', comment:'Web fehlt/unklar' };
  const ok = withinToleranceKG(exKg, wbKg, WEIGHT_TOL_PCT);
  const diffPct = ((wbKg - exKg) / Math.max(1e-9, Math.abs(exKg))) * 100;
  return ok ? { status:'GREEN', comment:`Δ ${diffPct.toFixed(1)}%` }
            : { status:'RED', comment:`Excel ${exKg.toFixed(3)} kg vs. Web ${wbKg.toFixed(3)} kg (${diffPct.toFixed(1)}%)` };
}
function compareDimensions(excelU, excelV, excelW, webDimText) {
  const L = toNumber(excelU); const B = toNumber(excelV); const H = toNumber(excelW);
  const fromWeb = parseDimensionsToLBH(webDimText);
  const anyExcel = L!=null || B!=null || H!=null;
  if (!anyExcel && !fromWeb.L && !fromWeb.B && !fromWeb.H) return { status:'ORANGE', comment:'Beide fehlen' };
  if (!anyExcel) return { status:'ORANGE', comment:'Excel fehlt' };
  if (!fromWeb.L && !fromWeb.B && !fromWeb.H) return { status:'ORANGE', comment:'Web fehlt/unklar' };
  const eq = (a,b)=> (a!=null && b!=null && Math.abs(a-b) < 1e-6);
  const match = eq(L, fromWeb.L) && eq(B, fromWeb.B) && eq(H, fromWeb.H);
  return match ? { status:'GREEN', comment:'L×B×H identisch (mm)' }
               : { status:'RED', comment:`Excel ${L||''}×${B||''}×${H||''} mm vs. Web ${fromWeb.L||''}×${fromWeb.B||''}×${fromWeb.H||''} mm` };
}

// Routes
app.get('/', (req, res) => res.sendFile(path.join(__dirname, 'index.html')));

// Single product (A2V only)
app.post('/api/scrape', async (req, res) => {
  try {
    const { articleNumber } = req.body || {};
    const a2v = String(articleNumber || '').trim().toUpperCase();
    if (!a2v.startsWith('A2V')) return res.status(400).json({ success:false, error:'Nur A2V-Nummern erlaubt' });
    const r = await scraper.scrapeOne(a2v);
    res.json({ success:true, data:r });
  } catch (e) {
    res.status(500).json({ success:false, error:e.message });
  }
});

const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 50 * 1024 * 1024 } });

app.post('/api/process-excel', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'Bitte Excel-Datei hochladen (file).' });

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(req.file.buffer);

    // 1) A2V-Nummern aus Spalte Z ab Zeile 4
    const a2vs = [];
    for (const ws of wb.worksheets) {
      const last = ws.lastRow?.number || 0;
      for (let r = FIRST_DATA_ROW; r <= last; r++) {
        const val = ws.getCell(`${COLS.Z}${r}`).value;
        if (val && String(val).trim()) {
          const a2v = String(val).trim().toUpperCase();
          if (a2v.startsWith('A2V')) a2vs.push(a2v);
        }
      }
    }

    // 2) Scrapen (A2V-only)
    const results = await scraper.scrapeMany(a2vs, SCRAPE_CONCURRENCY);

    // 3) Zwei Zusatzzeilen einfügen je Produkt
    for (const ws of wb.worksheets) {
      let lastRow = ws.lastRow?.number || 0;
      if (lastRow < HEADER_ROW) continue;
      const products = [];
      for (let r = FIRST_DATA_ROW; r <= lastRow; r++) {
        const anyValue = ['A','B','C','Z'].some(c => ws.getCell(`${c}${r}`).value && String(ws.getCell(`${c}${r}`).value).trim() !== '');
        if (anyValue) products.push(r);
      }

      for (let i = products.length - 1; i >= 0; i--) {
        const r = products[i];
        const A2V = String(ws.getCell(`${COLS.Z}${r}`).value || '').trim().toUpperCase();
        const manufNoExcel = ws.getCell(`${COLS.E}${r}`).value;
        const titleExcel = ws.getCell(`${COLS.C}${r}`).value;
        const weightExcel = ws.getCell(`${COLS.S}${r}`).value;
        const lenExcel = ws.getCell(`${COLS.U}${r}`).value;
        const widExcel = ws.getCell(`${COLS.V}${r}`).value;
        const heiExcel = ws.getCell(`${COLS.W}${r}`).value;
        const werkstoffExcel = ws.getCell(`${COLS.P}${r}`).value;
        const noteExcel = ws.getCell(`${COLS.N}${r}`).value;

        let webData = { A2V, URL: a2vUrl(A2V), Produkttitel:'', 'Weitere Artikelnummer':'', Gewicht:'', Abmessung:'', Werkstoff:'', Materialklassifizierung:'' };
        if (A2V.startsWith('A2V') && results.has(A2V)) webData = results.get(A2V);

        const insertAt = r + 1;
        ws.spliceRows(insertAt, 0, [null]);
        ws.spliceRows(insertAt + 1, 0, [null]);
        const webRow = insertAt;
        const cmpRow = insertAt + 1;

        // Web-Daten-Zeile
        ws.getCell(`${COLS.Z}${webRow}`).value = A2V || '';
        ws.getCell(`${COLS.E}${webRow}`).value = webData['Weitere Artikelnummer'] || '';
        ws.getCell(`${COLS.C}${webRow}`).value = webData.Produkttitel || '';
        ws.getCell(`${COLS.S}${webRow}`).value = webData.Gewicht || '';
        const dimParsed = parseDimensionsToLBH(webData.Abmessung);
        if (dimParsed.L != null) ws.getCell(`${COLS.U}${webRow}`).value = dimParsed.L;
        if (dimParsed.B != null) ws.getCell(`${COLS.V}${webRow}`).value = dimParsed.B;
        if (dimParsed.H != null) ws.getCell(`${COLS.W}${webRow}`).value = dimParsed.H;
        ws.getCell(`${COLS.P}${webRow}`).value = webData.Werkstoff || '';
        ws.getCell(`${COLS.N}${webRow}`).value = mapMaterialClassificationToExcel(webData.Materialklassifizierung) || '';

        // Vergleichs-Zeile
        const cmpZ = compareText(A2V || '', webData.A2V || A2V || '');
        ws.getCell(`${COLS.Z}${cmpRow}`).value = cmpZ.comment;
        ws.getCell(`${COLS.Z}${cmpRow}`).fill = statusCellFill(cmpZ.status);

        const cmpE = comparePartNo(manufNoExcel || '', webData['Weitere Artikelnummer'] || '');
        ws.getCell(`${COLS.E}${cmpRow}`).value = cmpE.comment;
        ws.getCell(`${COLS.E}${cmpRow}`).fill = statusCellFill(cmpE.status);

        const cmpC = compareText(titleExcel || '', webData.Produkttitel || '');
        ws.getCell(`${COLS.C}${cmpRow}`).value = cmpC.comment;
        ws.getCell(`${COLS.C}${cmpRow}`).fill = statusCellFill(cmpC.status);

        const cmpS = compareWeight(weightExcel, webData.Gewicht);
        ws.getCell(`${COLS.S}${cmpRow}`).value = cmpS.comment;
        ws.getCell(`${COLS.S}${cmpRow}`).fill = statusCellFill(cmpS.status);

        const cmpDim = compareDimensions(lenExcel, widExcel, heiExcel, webData.Abmessung);
        ws.getCell(`${COLS.U}${cmpRow}`).value = cmpDim.comment;
        ws.getCell(`${COLS.U}${cmpRow}`).fill = statusCellFill(cmpDim.status);

        const cmpP = compareText(werkstoffExcel || '', webData.Werkstoff || '');
        ws.getCell(`${COLS.P}${cmpRow}`).value = cmpP.comment;
        ws.getCell(`${COLS.P}${cmpRow}`).fill = statusCellFill(cmpP.status);

        const mappedN = mapMaterialClassificationToExcel(webData.Materialklassifizierung) || '';
        const cmpN = compareText((noteExcel||'').toString().toUpperCase(), mappedN.toString().toUpperCase());
        ws.getCell(`${COLS.N}${cmpRow}`).value = cmpN.comment;
        ws.getCell(`${COLS.N}${cmpRow}`).fill = statusCellFill(cmpN.status);
      }
    }

    const out = await wb.xlsx.writeBuffer();
    res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition','attachment; filename="DB_Produktvergleich_verarbeitet.xlsx"');
    res.send(Buffer.from(out));

  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

app.listen(PORT, () => console.log(`Server running at http://localhost:${PORT}`));