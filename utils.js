// utils.js - Normalisierung & Mapping

function a2vUrl(a2v) {
  const id = (a2v || '').toString().trim();
  return `https://www.mymobase.com/de/p/${id}`;
}

function cleanNumberString(s) {
  if (s == null) return null;
  const str = String(s).replace(/\s+/g, '').replace(',', '.'); // 12,3 -> 12.3
  return str;
}

function toNumber(val) {
  if (val == null || val === '') return null;
  const s = cleanNumberString(val);
  const m = s.match(/-?\d+(?:\.\d+)?/);
  if (!m) return null;
  return parseFloat(m[0]);
}

/**
 * Gewicht in kg normalisieren. Erlaubt mg, g, kg, t.
 */
function normalizeWeightToKg(value) {
  if (value == null || value === '') return null;
  const s = cleanNumberString(value).toLowerCase();
  const m = s.match(/-?\d+(?:\.\d+)?/);
  if (!m) return null;
  const num = parseFloat(m[0]);
  if (/mg/.test(s)) return num / 1e6;
  if (/[^k]g/.test(s) && !/kg/.test(s)) return num / 1000; // g
  if (/kg/.test(s)) return num;
  if (/\bt\b/.test(s)) return num * 1000;
  return num; // default: kg wenn keine Einheit
}

/**
 * Dimensions-Parser:
 * - akzeptiert "L×B×H", "LxBxH", "3X30X107,3X228", "30x20x10 mm", "0.3 x 0.2 x 0.1 m", etc.
 * - Ergebnis in mm (falls Einheiten erkennbar), sonst roh.
 */
function parseDimensionsToLBH(text) {
  if (!text) return { L:null, B:null, H:null };
  const raw = String(text).trim();
  let s = raw.toLowerCase().replace(/[×x]/g, 'x').replace(',', '.').replace(/\s+/g, '');

  // Einheit erkennen (mm, cm, m)
  let scale = 1; // default mm
  if (/[^a-z]cm\b/.test(s) || s.endsWith('cm')) scale = 10;
  if (/[^a-z]m\b/.test(s) || s.endsWith('m')) scale = 1000;
  // Zahlen extrahieren
  const nums = (s.match(/-?\d+(?:\.\d+)?/g) || []).map(parseFloat);

  // Heuristik: nimm die ersten drei Zahlen als L,B,H
  const L = nums.length > 0 ? Math.round(nums[0] * scale) : null;
  const B = nums.length > 1 ? Math.round(nums[1] * scale) : null;
  const H = nums.length > 2 ? Math.round(nums[2] * scale) : null;

  return { L, B, H };
}

/**
 * Artikelnummer-Normalisierung: Leerzeichen/Bindestriche/Slash entfernen, Uppercase.
 */
function normPartNo(s) {
  if (!s) return '';
  return String(s).toUpperCase().replace(/[\s\-\/_]+/g, '');
}

/**
 * Prozenttoleranz-Prüfung (für Gewicht). tolPct = 0 => streng gleich.
 */
function withinToleranceKG(exKg, wbKg, tolPct) {
  if (exKg == null || wbKg == null) return false;
  const diff = Math.abs(exKg - wbKg);
  if (!tolPct || tolPct <= 0) return diff < 1e-9; // streng
  const tol = Math.abs(exKg) * (tolPct / 100);
  return diff <= tol;
}

/**
 * Materialklassifizierung (Web-Text) → Excel-Code N
 * Beispiel: "Nicht Schweiss-/Guss-/Klebe-/Schmiede relevant" → "OHNE/N/N/N/N"
 */
function mapMaterialClassificationToExcel(text) {
  if (!text) return '';
  const s = String(text).toLowerCase();

  // häufige Schreibweisen tolerant prüfen
  const hasNicht = /nicht/.test(s);
  const hasSchweiss = /schwei|schweiß|schweiss/.test(s);
  const hasGuss = /guss/.test(s);
  const hasKlebe = /klebe/.test(s);
  const hasSchmiede = /schmiede/.test(s);
  const hasRelevant = /relev/.test(s);

  if (hasNicht && (hasSchweiss || hasGuss || hasKlebe || hasSchmiede) && hasRelevant) {
    return 'OHNE/N/N/N/N';
  }
  // sonst: keine sichere Zuordnung → leer lassen
  return '';
}

module.exports = {
  a2vUrl,
  cleanNumberString,
  toNumber,
  normalizeWeightToKg,
  parseDimensionsToLBH,
  normPartNo,
  withinToleranceKG,
  mapMaterialClassificationToExcel
};