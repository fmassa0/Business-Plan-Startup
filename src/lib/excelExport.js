import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { compute } from './compute.js';
import { Y } from './defaults.js';
import {
  injectNativeCharts,
  makeChartXml_RicaviEbitdaUtile,
  makeChartXml_FCF,
  makeChartXml_StatoPatrimoniale,
  makeChartXml_DSCRpfn
} from './xlsxCharts.js';

/* =====================================================================
   Stile Big4: blu Ingenia per header/section, input azzurrini, formule
   nere, totali in grassetto su sfondo verde chiaro, % e € formattati.
   Tutte le tabelle finanziarie sono in formula live: l'utente apre
   l'Excel, modifica le ipotesi e i prospetti si ricalcolano da soli.
   ===================================================================== */

const COL = {
  ingDeep:    'FF0A1F44',  // navy McKinsey
  ingDarker:  'FF061634',
  ingMid:     'FF1E3A8A',
  teal:       'FF00B0A0',  // accent
  tealLight:  'FFCFFAF4',
  ingLight:   'FFEEF4FF',
  inputBg:    'FFFFFBEB',
  inputFg:    'FF0A1F44',
  subTotal:   'FFEEF4FF',
  total:      'FFE6FAF6',  // teal tint per highlight
  section:    'FF0A1F44',
  sectionFg:  'FFFFFFFF',
  border:     'FFE2E7EE',
  borderSoft: 'FFEEF1F5',
  muted:      'FF64748B',
  white:      'FFFFFFFF',
  good:       'FFDCFCE7',
  goodFg:     'FF14532D',
  warn:       'FFFEF3C7',
  warnFg:     'FF78350F',
  bad:        'FFFEE2E2',
  badFg:      'FF7F1D1D'
};

/* Series colors for charts (must match Chart.js palette) */
export const CHART_PALETTE = {
  navy:  '0A1F44',
  navy2: '1E3A8A',
  teal:  '00B0A0',
  amber: 'F59E0B',
  red:   'EF4444',
  purple:'7C5BC9'
};

const FMT = {
  euro:    '#,##0;[Red](#,##0);"—"',
  euro2:   '#,##0.00;[Red](#,##0.00);"—"',
  pct:     '0.0%;[Red]-0.0%;"—"',
  pct1:    '0.0%',
  num2:    '0.00x',
  intGiorni: '0" gg"'
};

const yearCols = ['B', 'C', 'D', 'E', 'F'];
const yearCols0 = ['B', 'C', 'D', 'E', 'F', 'G']; // include Anno 0 per SP

const setBorder = (cell, sides = ['top', 'bottom', 'left', 'right']) => {
  cell.border = sides.reduce((acc, s) => ({ ...acc, [s]: { style: 'thin', color: { argb: COL.border } } }), cell.border || {});
};

const sectionRow = (ws, rowIdx, label, colspan) => {
  const row = ws.getRow(rowIdx);
  for (let c = 1; c <= colspan; c++) {
    const cell = row.getCell(c);
    cell.value = c === 1 ? label : '';
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COL.section } };
    cell.font = { name: 'Calibri', size: 10, bold: true, color: { argb: COL.sectionFg } };
    cell.alignment = { vertical: 'middle', horizontal: c === 1 ? 'left' : 'right' };
  }
  ws.mergeCells(rowIdx, 1, rowIdx, colspan);
  row.height = 18;
};

const headerRow = (ws, rowIdx, labels) => {
  const row = ws.getRow(rowIdx);
  labels.forEach((l, i) => {
    const cell = row.getCell(i + 1);
    cell.value = l;
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COL.ingLight } };
    cell.font = { name: 'Calibri', size: 10, bold: true, color: { argb: COL.ingDeep } };
    cell.alignment = { vertical: 'middle', horizontal: i === 0 ? 'left' : 'right' };
    setBorder(cell, ['bottom']);
  });
  row.height = 22;
};

const dataRow = (ws, rowIdx, label, valuesOrFormulas, opts = {}) => {
  const row = ws.getRow(rowIdx);
  const lblCell = row.getCell(1);
  lblCell.value = label;
  lblCell.font = {
    name: 'Calibri',
    size: 10,
    bold: !!opts.bold,
    italic: !!opts.italic,
    color: { argb: opts.muted ? COL.muted : 'FF0F172A' }
  };
  if (opts.indent) lblCell.alignment = { indent: 1 };

  valuesOrFormulas.forEach((v, i) => {
    const cell = row.getCell(i + 2);
    if (v === null || v === undefined) {
      cell.value = null;
    } else if (typeof v === 'object' && 'formula' in v) {
      cell.value = { formula: v.formula };
    } else {
      cell.value = v;
    }
    cell.numFmt = opts.fmt || FMT.euro;
    cell.font = {
      name: 'Calibri',
      size: 10,
      bold: !!opts.bold,
      color: { argb: opts.formula ? 'FF0F172A' : COL.inputFg }
    };
    if (opts.bold) {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COL.total } };
    } else if (opts.subtotal) {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COL.subTotal } };
    }
    cell.alignment = { vertical: 'middle', horizontal: 'right' };
  });

  if (opts.subtotal) lblCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COL.subTotal } };
  if (opts.bold) lblCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COL.total } };
  return row;
};

/* ------------------- COVER SHEET ------------------- */
function buildCover(wb, state, R, logoBase64) {
  const ws = wb.addWorksheet('Copertina', {
    properties: { tabColor: { argb: COL.ingDeep } },
    views: [{ showGridLines: false }]
  });
  ws.getColumn('A').width = 4;
  ws.getColumn('B').width = 32;
  ws.getColumn('C').width = 32;
  ws.getColumn('D').width = 32;

  // Banner (rows 2..7, cols B..D)
  for (let r = 2; r <= 7; r++) {
    for (let c = 2; c <= 4; c++) {
      const cell = ws.getCell(r, c);
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COL.ingDarker } };
    }
    ws.getRow(r).height = 22;
  }
  ws.mergeCells('B2:D2');
  ws.mergeCells('B3:D3');
  ws.mergeCells('B4:D4');
  ws.mergeCells('B5:D5');
  ws.mergeCells('B6:D6');
  ws.mergeCells('B7:D7');

  // Embed Ingenia logo PNG (top-right of banner) if provided
  if (logoBase64) {
    const imgId = wb.addImage({ base64: logoBase64, extension: 'png' });
    ws.addImage(imgId, {
      tl: { col: 3.2, row: 1.4 },
      ext: { width: 240, height: 73 }
    });
  }

  const t = ws.getCell('B4');
  t.value = 'Business Plan a 5 anni';
  t.font = { name: 'Calibri', size: 24, bold: true, color: { argb: COL.white } };
  t.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
  const sub = ws.getCell('B5');
  sub.value = 'Simulatore Start Up — CE · SP · RF · Indici';
  sub.font = { name: 'Calibri', size: 11, color: { argb: 'FFCBD5E1' } };
  sub.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
  const brand = ws.getCell('B6');
  brand.value = 'Ingenia · Powering Future · ingenia.cloud';
  brand.font = { name: 'Calibri', size: 11, bold: true, color: { argb: COL.white } };
  brand.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };

  // Sintesi scenario
  ws.getCell('B9').value = 'Sintesi dello scenario';
  ws.getCell('B9').font = { name: 'Calibri', size: 14, bold: true, color: { argb: COL.ingDeep } };

  const summary = [
    ['Ricavi 5° anno',          { v: R.CE[Y - 1].ricavi, fmt: FMT.euro }],
    ['CAGR ricavi',             { v: R.CE[0].ricavi > 0 ? Math.pow(R.CE[Y - 1].ricavi / R.CE[0].ricavi, 1 / (Y - 1)) - 1 : 0, fmt: FMT.pct1 }],
    ['EBITDA cumulato',         { v: R.CE.reduce((a, c) => a + c.ebitda, 0), fmt: FMT.euro }],
    ['Margine EBITDA medio',    { v: R.CE.reduce((a, c) => a + c.ebitdaMar, 0) / Y, fmt: FMT.pct1 }],
    ['Utile cumulato',          { v: R.CE.reduce((a, c) => a + c.utile, 0), fmt: FMT.euro }],
    ['VAN @ WACC',              { v: R.NPV, fmt: FMT.euro }],
    ['TIR (IRR)',               { v: isFinite(R.IRR) ? R.IRR : null, fmt: FMT.pct1 }],
    ['Payback',                 { v: R.payback ? `Anno ${R.payback}` : 'oltre Y5', fmt: '@' }],
    ['PFN / EBITDA Y5',         { v: R.ID[Y - 1].pfnEbitda, fmt: FMT.num2 }],
    ['DSCR medio',              { v: R.ID.reduce((a, d) => a + d.DSCR, 0) / Y, fmt: FMT.num2 }]
  ];

  summary.forEach(([lbl, val], i) => {
    const r = 12 + i;
    const lc = ws.getCell(`B${r}`);
    lc.value = lbl;
    lc.font = { name: 'Calibri', size: 10, color: { argb: COL.muted } };
    setBorder(lc, ['bottom']);
    const vc = ws.getCell(`C${r}`);
    vc.value = val.v;
    vc.numFmt = val.fmt;
    vc.font = { name: 'Calibri', size: 11, bold: true, color: { argb: COL.ingDeep } };
    vc.alignment = { horizontal: 'right' };
    setBorder(vc, ['bottom']);
  });

  // Sezione "Indice del workbook"
  ws.getCell('B10').value = 'Sezione';
  ws.getCell('B10').font = { name: 'Calibri', size: 10, bold: true, color: { argb: COL.muted } };
  ws.getCell('C10').value = 'Contenuto';
  ws.getCell('C10').font = { name: 'Calibri', size: 10, bold: true, color: { argb: COL.muted } };

  // Date + disclaimer
  const today = new Date().toLocaleDateString('it-IT');
  ws.getCell('B24').value = `Generato il ${today} con il Simulatore BP Ingenia`;
  ws.getCell('B24').font = { name: 'Calibri', size: 9, italic: true, color: { argb: COL.muted } };
  ws.getCell('B26').value = 'Modello didattico — non sostituisce un parere professionale. IRR e Payback sono valori statici al momento dell\'esportazione: per ricalcolarli dopo modifiche agli input, riesporta dal simulatore.';
  ws.getCell('B26').alignment = { wrapText: true };
  ws.getCell('B26').font = { name: 'Calibri', size: 9, italic: true, color: { argb: COL.muted } };
  ws.mergeCells('B26:D28');
}

/* ------------------- IPOTESI SHEET ------------------- */
/* Tutti gli input con named ranges; celle azzurrine = editabili dall'utente.
   Le formule in altri fogli puntano a questi nomi. */
function buildIpotesi(wb, state) {
  const ws = wb.addWorksheet('Ipotesi', { properties: { tabColor: { argb: '32CD32' } } });
  ws.getColumn('A').width = 36;
  ['B', 'C', 'D', 'E', 'F'].forEach((c) => (ws.getColumn(c).width = 14));
  ws.views = [{ state: 'frozen', xSplit: 1, ySplit: 3 }];

  // Title
  ws.getCell('A1').value = 'Ipotesi & Parametri';
  ws.getCell('A1').font = { name: 'Calibri', size: 14, bold: true, color: { argb: COL.ingDeep } };
  ws.getCell('A2').value = 'Le celle in giallo chiaro sono editabili. Modifica i valori e i prospetti si ricalcolano.';
  ws.getCell('A2').font = { name: 'Calibri', size: 9, italic: true, color: { argb: COL.muted } };
  ws.mergeCells('A1:F1');
  ws.mergeCells('A2:F2');

  let r = 3;
  const setInput = (rowIdx, col, name, value, fmt) => {
    const cell = ws.getCell(rowIdx, col);
    cell.value = value;
    cell.numFmt = fmt;
    cell.font = { name: 'Calibri', size: 10, bold: true, color: { argb: COL.inputFg } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COL.inputBg } };
    cell.alignment = { horizontal: 'right' };
    setBorder(cell);
    if (name) {
      const colLetter = String.fromCharCode(64 + col);
      wb.definedNames.add(`Ipotesi!$${colLetter}$${rowIdx}`, name);
    }
  };
  const setLabel = (rowIdx, label) => {
    const cell = ws.getCell(rowIdx, 1);
    cell.value = label;
    cell.font = { name: 'Calibri', size: 10 };
  };
  const sectionTitle = (rowIdx, label) => {
    const cell = ws.getCell(rowIdx, 1);
    cell.value = label;
    cell.font = { name: 'Calibri', size: 11, bold: true, color: { argb: COL.ingDeep } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COL.ingLight } };
    for (let c = 2; c <= 6; c++) {
      ws.getCell(rowIdx, c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COL.ingLight } };
    }
  };

  r = 4; sectionTitle(r, 'Ricavi — Volumi e Prezzi');
  r++; setLabel(r, 'Volumi venduti Y1 (pz)'); setInput(r, 2, 'vol', +state.i_vol || 0, '#,##0');
  r++; setLabel(r, 'Prezzo unitario Y1 (€)'); setInput(r, 2, 'prc', +state.i_prc || 0, '#,##0.00');
  r++; setLabel(r, 'Crescita volumi % / anno'); setInput(r, 2, 'g_vol', (+state.g_vol || 0) / 100, FMT.pct1);
  r++; setLabel(r, 'Crescita prezzo % / anno'); setInput(r, 2, 'g_prc', (+state.g_prc || 0) / 100, FMT.pct1);

  r += 2; sectionTitle(r, 'Costi Variabili');
  r++; setLabel(r, 'Costo unit. variab. (€/pz)'); setInput(r, 2, 'cuv', +state.i_cuv || 0, '#,##0.00');
  r++; setLabel(r, 'Crescita costo unit. % / anno'); setInput(r, 2, 'g_cuv', (+state.g_cuv || 0) / 100, FMT.pct1);

  r += 2; sectionTitle(r, 'Costi Fissi Operativi (Y1)');
  r++; setLabel(r, 'Personale (€)'); setInput(r, 2, 'pers', +state.i_pers || 0, '#,##0');
  r++; setLabel(r, 'Affitto / utenze (€)'); setInput(r, 2, 'aff', +state.i_aff || 0, '#,##0');
  r++; setLabel(r, 'Marketing (€)'); setInput(r, 2, 'mkt', +state.i_mkt || 0, '#,##0');
  r++; setLabel(r, 'Servizi/consulenze (€)'); setInput(r, 2, 'serv', +state.i_serv || 0, '#,##0');
  r++; setLabel(r, 'Altri costi fissi (€)'); setInput(r, 2, 'alt', +state.i_alt || 0, '#,##0');
  r++; setLabel(r, 'Crescita costi fissi % / anno'); setInput(r, 2, 'g_cf', (+state.g_cf || 0) / 100, FMT.pct1);

  r += 2; sectionTitle(r, 'Investimenti & Ammortamenti');
  r++;
  ws.getCell(r, 1).value = 'Anno';
  ws.getCell(r, 1).font = { name: 'Calibri', size: 9, color: { argb: COL.muted } };
  for (let y = 0; y < Y; y++) {
    const c = ws.getCell(r, 2 + y);
    c.value = `Y${y + 1}`;
    c.font = { name: 'Calibri', size: 9, color: { argb: COL.muted } };
    c.alignment = { horizontal: 'right' };
  }
  r++;
  setLabel(r, 'Inv. materiali (€)');
  ['im1', 'im2', 'im3', 'im4', 'im5'].forEach((k, i) => setInput(r, 2 + i, k, +state[k] || 0, '#,##0'));
  r++;
  setLabel(r, 'Inv. immateriali (€)');
  ['ii1', 'ii2', 'ii3', 'ii4', 'ii5'].forEach((k, i) => setInput(r, 2 + i, k, +state[k] || 0, '#,##0'));
  r++;
  setLabel(r, 'Aliquota amm. materiali %');
  setInput(r, 2, 'aliq_m', (+state.aliq_m || 0) / 100, FMT.pct1);
  r++;
  setLabel(r, 'Aliquota amm. immateriali %');
  setInput(r, 2, 'aliq_i', (+state.aliq_i || 0) / 100, FMT.pct1);

  r += 2; sectionTitle(r, 'Capitale Circolante (giorni)');
  r++; setLabel(r, 'DSO clienti (giorni)'); setInput(r, 2, 'dso', +state.dso || 0, '0');
  r++; setLabel(r, 'DPO fornitori (giorni)'); setInput(r, 2, 'dpo', +state.dpo || 0, '0');
  r++; setLabel(r, 'DIO scorte (giorni)'); setInput(r, 2, 'dio', +state.dio || 0, '0');

  r += 2; sectionTitle(r, 'IVA');
  r++; setLabel(r, 'IVA vendite %'); setInput(r, 2, 'iva_v', (+state.iva_v || 0) / 100, FMT.pct1);
  r++; setLabel(r, 'IVA acquisti %'); setInput(r, 2, 'iva_a', (+state.iva_a || 0) / 100, FMT.pct1);
  r++; setLabel(r, 'Periodicità (mesi)'); setInput(r, 2, 'iva_p', +state.iva_p || 0, '0');

  r += 2; sectionTitle(r, 'Finanziamenti & Equity');
  r++;
  ws.getCell(r, 1).value = 'Anno';
  ws.getCell(r, 1).font = { name: 'Calibri', size: 9, color: { argb: COL.muted } };
  for (let y = 0; y < Y; y++) {
    const c = ws.getCell(r, 2 + y);
    c.value = `Y${y + 1}`;
    c.font = { name: 'Calibri', size: 9, color: { argb: COL.muted } };
    c.alignment = { horizontal: 'right' };
  }
  r++;
  setLabel(r, 'Apporto equity (€)');
  ['eq1', 'eq2', 'eq3', 'eq4', 'eq5'].forEach((k, i) => setInput(r, 2 + i, k, +state[k] || 0, '#,##0'));
  r++;
  setLabel(r, 'Nuovo debito bancario (€)');
  ['db1', 'db2', 'db3', 'db4', 'db5'].forEach((k, i) => setInput(r, 2 + i, k, +state[k] || 0, '#,##0'));
  r++;
  setLabel(r, 'Rimborso quota capitale (€)');
  ['rb1', 'rb2', 'rb3', 'rb4', 'rb5'].forEach((k, i) => setInput(r, 2 + i, k, +state[k] || 0, '#,##0'));
  r++;
  setLabel(r, 'Tasso interesse sul debito %');
  setInput(r, 2, 'tasso', (+state.tasso || 0) / 100, FMT.pct1);

  r += 2; sectionTitle(r, 'Fiscalità & WACC');
  r++; setLabel(r, 'Aliquota IRES %'); setInput(r, 2, 'ires', (+state.ires || 0) / 100, FMT.pct1);
  r++; setLabel(r, 'Aliquota IRAP %'); setInput(r, 2, 'irap', (+state.irap || 0) / 100, FMT.pct1);
  r++; setLabel(r, 'WACC %'); setInput(r, 2, 'wacc', (+state.wacc || 0) / 100, FMT.pct1);
}

/* ------------------- CONTO ECONOMICO SHEET ------------------- */
function buildCE(wb) {
  const ws = wb.addWorksheet('Conto Economico', { properties: { tabColor: { argb: COL.ingDeep } } });
  ws.getColumn('A').width = 38;
  yearCols.forEach((c) => (ws.getColumn(c).width = 15));
  ws.views = [{ state: 'frozen', xSplit: 1, ySplit: 4 }];

  ws.getCell('A1').value = 'Conto Economico Previsionale';
  ws.getCell('A1').font = { name: 'Calibri', size: 14, bold: true, color: { argb: COL.ingDeep } };
  ws.mergeCells('A1:F1');

  headerRow(ws, 3, ['Voce', 'Anno 1', 'Anno 2', 'Anno 3', 'Anno 4', 'Anno 5']);

  let r = 4;
  // Volumi (riga di servizio)
  dataRow(ws, r, 'Volumi venduti (pz)', [
    { formula: 'vol' },
    { formula: 'B' + r + '*(1+g_vol)' },
    { formula: 'C' + r + '*(1+g_vol)' },
    { formula: 'D' + r + '*(1+g_vol)' },
    { formula: 'E' + r + '*(1+g_vol)' }
  ], { fmt: '#,##0', muted: true, italic: true, formula: true });
  const rowVol = r;
  r++;
  dataRow(ws, r, 'Prezzo unitario (€)', [
    { formula: 'prc' },
    { formula: 'B' + r + '*(1+g_prc)' },
    { formula: 'C' + r + '*(1+g_prc)' },
    { formula: 'D' + r + '*(1+g_prc)' },
    { formula: 'E' + r + '*(1+g_prc)' }
  ], { fmt: '#,##0.00', muted: true, italic: true, formula: true });
  const rowPrc = r;
  r++;
  dataRow(ws, r, 'Costo variabile unitario (€)', [
    { formula: 'cuv' },
    { formula: 'B' + r + '*(1+g_cuv)' },
    { formula: 'C' + r + '*(1+g_cuv)' },
    { formula: 'D' + r + '*(1+g_cuv)' },
    { formula: 'E' + r + '*(1+g_cuv)' }
  ], { fmt: '#,##0.00', muted: true, italic: true, formula: true });
  const rowCuv = r;
  r++;

  // Ricavi
  r++;
  dataRow(ws, r, 'Ricavi delle vendite', yearCols.map((c) => ({ formula: `${c}${rowVol}*${c}${rowPrc}` })), { formula: true });
  const rowRic = r;
  r++;
  dataRow(ws, r, '(-) Costi variabili', yearCols.map((c) => ({ formula: `-${c}${rowVol}*${c}${rowCuv}` })), { formula: true });
  const rowCVar = r;
  r++;
  dataRow(ws, r, 'MARGINE DI CONTRIBUZIONE', yearCols.map((c) => ({ formula: `${c}${rowRic}+${c}${rowCVar}` })), { subtotal: true, formula: true });
  const rowMdc = r;
  r++;
  dataRow(ws, r, 'MdC %', yearCols.map((c) => ({ formula: `IFERROR(${c}${rowMdc}/${c}${rowRic},0)` })), { fmt: FMT.pct, muted: true, formula: true });
  r++;

  sectionRow(ws, r, 'Costi Fissi', 6);
  r++;
  // Costi fissi con crescita: B = -pers, C = B*(1+g_cf), ...
  const fixedRows = [
    ['(-) Personale',           'pers'],
    ['(-) Affitto / utenze',    'aff'],
    ['(-) Marketing',           'mkt'],
    ['(-) Servizi e consulenze','serv'],
    ['(-) Altri costi fissi',   'alt']
  ];
  const fixedRowIdx = [];
  fixedRows.forEach(([lbl, name]) => {
    dataRow(ws, r, lbl, [
      { formula: `-${name}` },
      { formula: `B${r}*(1+g_cf)` },
      { formula: `C${r}*(1+g_cf)` },
      { formula: `D${r}*(1+g_cf)` },
      { formula: `E${r}*(1+g_cf)` }
    ], { indent: true, formula: true });
    fixedRowIdx.push(r);
    r++;
  });
  dataRow(ws, r, 'Totale costi fissi',
    yearCols.map((c) => ({ formula: fixedRowIdx.map((rr) => `${c}${rr}`).join('+') })),
    { subtotal: true, formula: true });
  const rowTotCF = r;
  r++;

  // EBITDA
  dataRow(ws, r, 'EBITDA (MOL)', yearCols.map((c) => ({ formula: `${c}${rowMdc}+${c}${rowTotCF}` })), { bold: true, formula: true });
  const rowEbitda = r;
  r++;
  dataRow(ws, r, 'EBITDA margin %', yearCols.map((c) => ({ formula: `IFERROR(${c}${rowEbitda}/${c}${rowRic},0)` })), { fmt: FMT.pct, muted: true, formula: true });
  r++;

  // Ammortamenti cumulativi
  // ammMat[y] = sum(im_yk * aliq_m, k=1..y)
  dataRow(ws, r, '(-) Ammortamenti materiali', [
    { formula: '-im1*aliq_m' },
    { formula: `B${r}-im2*aliq_m` },
    { formula: `C${r}-im3*aliq_m` },
    { formula: `D${r}-im4*aliq_m` },
    { formula: `E${r}-im5*aliq_m` }
  ], { formula: true });
  const rowAmmM = r;
  r++;
  dataRow(ws, r, '(-) Ammortamenti immateriali', [
    { formula: '-ii1*aliq_i' },
    { formula: `B${r}-ii2*aliq_i` },
    { formula: `C${r}-ii3*aliq_i` },
    { formula: `D${r}-ii4*aliq_i` },
    { formula: `E${r}-ii5*aliq_i` }
  ], { formula: true });
  const rowAmmI = r;
  r++;

  // EBIT
  dataRow(ws, r, 'EBIT (Reddito Operativo)',
    yearCols.map((c) => ({ formula: `${c}${rowEbitda}+${c}${rowAmmM}+${c}${rowAmmI}` })),
    { bold: true, formula: true });
  const rowEbit = r;
  r++;
  dataRow(ws, r, 'EBIT margin %',
    yearCols.map((c) => ({ formula: `IFERROR(${c}${rowEbit}/${c}${rowRic},0)` })),
    { fmt: FMT.pct, muted: true, formula: true });
  r++;

  // Oneri finanziari = -openDebt * tasso. openDebt[y] = SP debFin[y-1]
  // Anno 1: openDebt = 0 → oneriFin = 0
  // Anno 2..5: oneriFin = -SP!B(debFin)[anno y-1] * tasso
  // Riferimento al foglio SP — sarà compilato dopo. Uso placeholder con SP!B<rowDebFin>
  const SP_DEBFIN_PLACEHOLDER = '__DEBFIN_ROW__';
  dataRow(ws, r, '(-) Oneri finanziari', [
    0,
    { formula: `-'Stato Patrimoniale'!B${SP_DEBFIN_PLACEHOLDER}*tasso` },
    { formula: `-'Stato Patrimoniale'!C${SP_DEBFIN_PLACEHOLDER}*tasso` },
    { formula: `-'Stato Patrimoniale'!D${SP_DEBFIN_PLACEHOLDER}*tasso` },
    { formula: `-'Stato Patrimoniale'!E${SP_DEBFIN_PLACEHOLDER}*tasso` }
  ], { formula: true });
  const rowOneri = r;
  r++;

  dataRow(ws, r, 'Risultato ante imposte (RAI)',
    yearCols.map((c) => ({ formula: `${c}${rowEbit}+${c}${rowOneri}` })),
    { subtotal: true, formula: true });
  const rowRai = r;
  r++;
  dataRow(ws, r, '(-) Imposte (IRES + IRAP)',
    yearCols.map((c) => ({ formula: `-MAX(${c}${rowRai},0)*(ires+irap)` })),
    { formula: true });
  const rowImposte = r;
  r++;
  dataRow(ws, r, 'UTILE / (PERDITA) NETTA',
    yearCols.map((c) => ({ formula: `${c}${rowRai}+${c}${rowImposte}` })),
    { bold: true, formula: true });
  const rowUtile = r;

  return {
    rows: { rowRic, rowCVar, rowMdc, rowTotCF, rowEbitda, rowAmmM, rowAmmI, rowEbit, rowOneri, rowRai, rowImposte, rowUtile },
    placeholders: { SP_DEBFIN_PLACEHOLDER }
  };
}

/* ------------------- STATO PATRIMONIALE ------------------- */
function buildSP(wb, ceMeta) {
  const ws = wb.addWorksheet('Stato Patrimoniale', { properties: { tabColor: { argb: COL.ingDeep } } });
  ws.getColumn('A').width = 40;
  yearCols0.forEach((c) => (ws.getColumn(c).width = 15));
  ws.views = [{ state: 'frozen', xSplit: 1, ySplit: 4 }];

  ws.getCell('A1').value = 'Stato Patrimoniale Previsionale';
  ws.getCell('A1').font = { name: 'Calibri', size: 14, bold: true, color: { argb: COL.ingDeep } };
  ws.mergeCells('A1:G1');

  headerRow(ws, 3, ['Voce', 'Anno 0', 'Anno 1', 'Anno 2', 'Anno 3', 'Anno 4', 'Anno 5']);

  // Helpers per colonne (Anno 0 = B, Anno 1 = C, ..., Anno 5 = G)
  const Y_COLS = ['C', 'D', 'E', 'F', 'G'];   // anni 1..5
  const PREV  = { C: 'B', D: 'C', E: 'D', F: 'E', G: 'F' }; // colonna anno precedente
  const INV_M = { C: 'im1', D: 'im2', E: 'im3', F: 'im4', G: 'im5' };
  const INV_I = { C: 'ii1', D: 'ii2', E: 'ii3', F: 'ii4', G: 'ii5' };
  const EQ    = { C: 'eq1', D: 'eq2', E: 'eq3', F: 'eq4', G: 'eq5' };
  const DB    = { C: 'db1', D: 'db2', E: 'db3', F: 'db4', G: 'db5' };
  const RB    = { C: 'rb1', D: 'rb2', E: 'rb3', F: 'rb4', G: 'rb5' };

  let r = 4;
  sectionRow(ws, r, 'ATTIVO', 7);
  r++;
  // Imm. materiali lordo: Anno 0 = 0, Anno y = anno_prev + im_y
  dataRow(ws, r, 'Imm. materiali (lordo)',
    [0, ...Y_COLS.map((c) => ({ formula: `${PREV[c]}${r}+${INV_M[c]}` }))],
    { formula: true });
  const rowImL = r;
  r++;
  // F.do amm. materiali: Anno 0 = 0, Anno y = anno_prev + |CE!ammMat|
  dataRow(ws, r, 'F.do amm. materiali',
    [0, ...Y_COLS.map((c, i) => {
      const ceCol = yearCols[i];
      return { formula: `${PREV[c]}${r}+(-'Conto Economico'!${ceCol}${ceMeta.rows.rowAmmM})` };
    })],
    { formula: true });
  const rowFAmmM = r;
  r++;
  dataRow(ws, r, 'Imm. materiali nette',
    [0, ...Y_COLS.map((c) => ({ formula: `${c}${rowImL}-${c}${rowFAmmM}` }))],
    { subtotal: true, formula: true });
  const rowImN = r;
  r++;

  dataRow(ws, r, 'Imm. immateriali (lordo)',
    [0, ...Y_COLS.map((c) => ({ formula: `${PREV[c]}${r}+${INV_I[c]}` }))],
    { formula: true });
  const rowIiL = r;
  r++;
  dataRow(ws, r, 'F.do amm. immateriali',
    [0, ...Y_COLS.map((c, i) => {
      const ceCol = yearCols[i];
      return { formula: `${PREV[c]}${r}+(-'Conto Economico'!${ceCol}${ceMeta.rows.rowAmmI})` };
    })],
    { formula: true });
  const rowFAmmI = r;
  r++;
  dataRow(ws, r, 'Imm. immateriali nette',
    [0, ...Y_COLS.map((c) => ({ formula: `${c}${rowIiL}-${c}${rowFAmmI}` }))],
    { subtotal: true, formula: true });
  const rowIiN = r;
  r++;

  dataRow(ws, r, 'TOTALE ATTIVO IMMOBILIZZATO',
    [0, ...Y_COLS.map((c) => ({ formula: `${c}${rowImN}+${c}${rowIiN}` }))],
    { bold: true, formula: true });
  const rowTotImm = r;
  r++;

  // Rimanenze = -CE.cVar * dio / 365  (cVar è negativo nel CE → -CE.cVar = positivo)
  dataRow(ws, r, 'Rimanenze',
    [0, ...Y_COLS.map((c, i) => {
      const ceCol = yearCols[i];
      return { formula: `(-'Conto Economico'!${ceCol}${ceMeta.rows.rowCVar})*dio/365` };
    })],
    { formula: true });
  const rowRim = r;
  r++;

  // Crediti = ricavi * (1+iva_v) * dso / 365
  dataRow(ws, r, 'Crediti commerciali (IVA incl.)',
    [0, ...Y_COLS.map((c, i) => {
      const ceCol = yearCols[i];
      return { formula: `'Conto Economico'!${ceCol}${ceMeta.rows.rowRic}*(1+iva_v)*dso/365` };
    })],
    { formula: true });
  const rowCre = r;
  r++;

  // Cassa: la calcoliamo dopo via RF; placeholder per ora (formula al ciclo)
  // Per ottenerla, calcoliamo prima totAttivoCirc senza cassa, poi recuperiamo cassa da quadratura?
  // Più pulito: la cassa la mettiamo come SP_calculated = openCash + RF.varCassa per ogni anno.
  // Anno 0 = 0; Anno y = Anno_prev + RF!flussoTotale
  // RF non è ancora stato costruito → uso placeholder
  const RF_VARCASSA_PLACEHOLDER = '__RF_VARCASSA_ROW__';
  dataRow(ws, r, 'Liquidità (cassa e banche)',
    [0, ...Y_COLS.map((c, i) => {
      const rfCol = yearCols[i];
      return { formula: `${PREV[c]}${r}+'Rendiconto Finanziario'!${rfCol}${RF_VARCASSA_PLACEHOLDER}` };
    })],
    { formula: true });
  const rowCassa = r;
  r++;

  dataRow(ws, r, 'TOTALE ATTIVO CIRCOLANTE',
    [0, ...Y_COLS.map((c) => ({ formula: `${c}${rowRim}+${c}${rowCre}+${c}${rowCassa}` }))],
    { subtotal: true, formula: true });
  const rowTotAC = r;
  r++;

  dataRow(ws, r, 'TOTALE ATTIVO',
    [0, ...Y_COLS.map((c) => ({ formula: `${c}${rowTotImm}+${c}${rowTotAC}` }))],
    { bold: true, formula: true });
  const rowTotAtt = r;
  r++;

  // PASSIVO
  sectionRow(ws, r, 'PASSIVO E PATRIMONIO NETTO', 7);
  r++;

  // Capitale sociale: Anno 0 = 0, Anno y = anno_prev + eq_y
  dataRow(ws, r, 'Capitale sociale',
    [0, ...Y_COLS.map((c) => ({ formula: `${PREV[c]}${r}+${EQ[c]}` }))],
    { formula: true });
  const rowCap = r;
  r++;

  // Utili a nuovo: Anno y = utNuovo[y-1] + utile_esercizio[y-1]; Anno 0 = 0
  // Quindi Anno 1 = 0; Anno y >= 2: = anno_prev_utNuovo + anno_prev_utile_esercizio
  dataRow(ws, r, 'Utili / (perdite) a nuovo',
    [0, 0, ...['D', 'E', 'F', 'G'].map((c, i) => {
      const prevCol = PREV[c];
      const prevUtileRow = r + 1; // riga sotto = utile esercizio
      return { formula: `${prevCol}${r}+${prevCol}${prevUtileRow}` };
    })],
    { formula: true });
  const rowUtNuovo = r;
  r++;

  // Utile esercizio: linkato al CE
  dataRow(ws, r, 'Utile / (perdita) d\'esercizio',
    [0, ...Y_COLS.map((c, i) => {
      const ceCol = yearCols[i];
      return { formula: `'Conto Economico'!${ceCol}${ceMeta.rows.rowUtile}` };
    })],
    { formula: true });
  const rowUtEs = r;
  r++;

  dataRow(ws, r, 'TOTALE PATRIMONIO NETTO',
    [0, ...Y_COLS.map((c) => ({ formula: `${c}${rowCap}+${c}${rowUtNuovo}+${c}${rowUtEs}` }))],
    { bold: true, formula: true });
  const rowPN = r;
  r++;

  // Debiti finanziari: Anno y = anno_prev + db_y - rb_y
  dataRow(ws, r, 'Debiti finanziari',
    [0, ...Y_COLS.map((c) => ({ formula: `${PREV[c]}${r}+${DB[c]}-${RB[c]}` }))],
    { formula: true });
  const rowDebFin = r;
  r++;

  // Debiti commerciali = (-cVar) * (1+iva_a) * dpo / 365
  dataRow(ws, r, 'Debiti commerciali (IVA incl.)',
    [0, ...Y_COLS.map((c, i) => {
      const ceCol = yearCols[i];
      return { formula: `(-'Conto Economico'!${ceCol}${ceMeta.rows.rowCVar})*(1+iva_a)*dpo/365` };
    })],
    { formula: true });
  const rowDebComm = r;
  r++;

  // Debito IVA = (ricavi*iva_v - cVar_pos*iva_a) * iva_p / 12
  dataRow(ws, r, 'Debito IVA verso Erario',
    [0, ...Y_COLS.map((c, i) => {
      const ceCol = yearCols[i];
      return { formula: `('Conto Economico'!${ceCol}${ceMeta.rows.rowRic}*iva_v-(-'Conto Economico'!${ceCol}${ceMeta.rows.rowCVar})*iva_a)*iva_p/12` };
    })],
    { formula: true });
  const rowDebIVA = r;
  r++;

  // Debiti tributari = -CE.imposte (sono negative → diventa positivo)
  dataRow(ws, r, 'Debiti tributari',
    [0, ...Y_COLS.map((c, i) => {
      const ceCol = yearCols[i];
      return { formula: `-'Conto Economico'!${ceCol}${ceMeta.rows.rowImposte}` };
    })],
    { formula: true });
  const rowDebTrib = r;
  r++;

  dataRow(ws, r, 'TOTALE PASSIVO',
    [0, ...Y_COLS.map((c) => ({ formula: `${c}${rowDebFin}+${c}${rowDebComm}+${c}${rowDebIVA}+${c}${rowDebTrib}` }))],
    { subtotal: true, formula: true });
  const rowTotPass = r;
  r++;

  dataRow(ws, r, 'TOTALE P + PN',
    [0, ...Y_COLS.map((c) => ({ formula: `${c}${rowPN}+${c}${rowTotPass}` }))],
    { bold: true, formula: true });
  const rowTotPP = r;
  r++;

  dataRow(ws, r, 'Quadratura (Att - P&PN)',
    [0, ...Y_COLS.map((c) => ({ formula: `${c}${rowTotAtt}-${c}${rowTotPP}` }))],
    { fmt: FMT.euro, muted: true, formula: true });
  r++;

  // Helper: Altri debiti (commerc.+IVA+trib.) — usato per grafico SP stacked
  dataRow(ws, r, 'Altri debiti (commerc.+IVA+trib.)',
    [0, ...Y_COLS.map((c) => ({ formula: `${c}${rowDebComm}+${c}${rowDebIVA}+${c}${rowDebTrib}` }))],
    { fmt: FMT.euro, muted: true, italic: true, formula: true });
  const rowDebOther = r;

  return {
    rows: {
      rowImL, rowFAmmM, rowImN, rowIiL, rowFAmmI, rowIiN,
      rowTotImm, rowRim, rowCre, rowCassa, rowTotAC, rowTotAtt,
      rowCap, rowUtNuovo, rowUtEs, rowPN,
      rowDebFin, rowDebComm, rowDebIVA, rowDebTrib, rowDebOther, rowTotPass, rowTotPP
    },
    placeholders: { RF_VARCASSA_PLACEHOLDER }
  };
}

/* ------------------- RENDICONTO FINANZIARIO ------------------- */
function buildRF(wb, ceMeta, spMeta) {
  const ws = wb.addWorksheet('Rendiconto Finanziario', { properties: { tabColor: { argb: COL.ingDeep } } });
  ws.getColumn('A').width = 40;
  yearCols.forEach((c) => (ws.getColumn(c).width = 15));
  ws.views = [{ state: 'frozen', xSplit: 1, ySplit: 4 }];

  ws.getCell('A1').value = 'Rendiconto Finanziario (metodo indiretto)';
  ws.getCell('A1').font = { name: 'Calibri', size: 14, bold: true, color: { argb: COL.ingDeep } };
  ws.mergeCells('A1:F1');

  headerRow(ws, 3, ['Voce', 'Anno 1', 'Anno 2', 'Anno 3', 'Anno 4', 'Anno 5']);

  // SP usa colonne C-G per anni 1-5; CE usa B-F. Mapping CE → RF: identico (B-F).
  const SP_C = ['C', 'D', 'E', 'F', 'G']; // colonne SP per anni 1..5
  const SP_PREV = { C: 'B', D: 'C', E: 'D', F: 'E', G: 'F' };

  let r = 4;
  sectionRow(ws, r, 'A) Gestione Operativa', 6);
  r++;

  dataRow(ws, r, "Utile / (perdita) d'esercizio",
    yearCols.map((c) => ({ formula: `'Conto Economico'!${c}${ceMeta.rows.rowUtile}` })),
    { formula: true });
  const rfUtile = r;
  r++;
  dataRow(ws, r, '(+) Ammortamenti (sterilizzati)',
    yearCols.map((c) => ({ formula: `-'Conto Economico'!${c}${ceMeta.rows.rowAmmM}-'Conto Economico'!${c}${ceMeta.rows.rowAmmI}` })),
    { formula: true });
  const rfAmm = r;
  r++;
  dataRow(ws, r, '(+) Oneri finanziari (sterilizzati)',
    yearCols.map((c) => ({ formula: `-'Conto Economico'!${c}${ceMeta.rows.rowOneri}` })),
    { formula: true });
  const rfOneri = r;
  r++;
  dataRow(ws, r, 'Flusso prima del CCN',
    yearCols.map((c) => ({ formula: `${c}${rfUtile}+${c}${rfAmm}+${c}${rfOneri}` })),
    { subtotal: true, formula: true });
  const rfPreCCN = r;
  r++;
  // Δ Rimanenze = -(Rim_y - Rim_y-1) → SP rimanenze, anno y in colonna SP_C[i]; precedente in SP_PREV
  dataRow(ws, r, '(-) Δ Rimanenze',
    yearCols.map((c, i) => {
      const cur = SP_C[i], prev = SP_PREV[cur];
      return { formula: `-('Stato Patrimoniale'!${cur}${spMeta.rows.rowRim}-'Stato Patrimoniale'!${prev}${spMeta.rows.rowRim})` };
    }),
    { formula: true });
  const rfDRim = r;
  r++;
  dataRow(ws, r, '(-) Δ Crediti commerciali',
    yearCols.map((c, i) => {
      const cur = SP_C[i], prev = SP_PREV[cur];
      return { formula: `-('Stato Patrimoniale'!${cur}${spMeta.rows.rowCre}-'Stato Patrimoniale'!${prev}${spMeta.rows.rowCre})` };
    }),
    { formula: true });
  const rfDCre = r;
  r++;
  dataRow(ws, r, '(+) Δ Debiti commerciali',
    yearCols.map((c, i) => {
      const cur = SP_C[i], prev = SP_PREV[cur];
      return { formula: `'Stato Patrimoniale'!${cur}${spMeta.rows.rowDebComm}-'Stato Patrimoniale'!${prev}${spMeta.rows.rowDebComm}` };
    }),
    { formula: true });
  const rfDDebC = r;
  r++;
  dataRow(ws, r, '(+) Δ Debito IVA',
    yearCols.map((c, i) => {
      const cur = SP_C[i], prev = SP_PREV[cur];
      return { formula: `'Stato Patrimoniale'!${cur}${spMeta.rows.rowDebIVA}-'Stato Patrimoniale'!${prev}${spMeta.rows.rowDebIVA}` };
    }),
    { formula: true });
  const rfDDebIVA = r;
  r++;
  dataRow(ws, r, '(-) Pagamento oneri finanziari',
    yearCols.map((c) => ({ formula: `'Conto Economico'!${c}${ceMeta.rows.rowOneri}` })),
    { formula: true });
  const rfPagOneri = r;
  r++;
  dataRow(ws, r, '(+) Δ Debiti tributari',
    yearCols.map((c, i) => {
      const cur = SP_C[i], prev = SP_PREV[cur];
      return { formula: `'Stato Patrimoniale'!${cur}${spMeta.rows.rowDebTrib}-'Stato Patrimoniale'!${prev}${spMeta.rows.rowDebTrib}` };
    }),
    { formula: true });
  const rfDDebTrib = r;
  r++;
  dataRow(ws, r, 'FLUSSO GESTIONE OPERATIVA',
    yearCols.map((c) => ({ formula: `${c}${rfPreCCN}+${c}${rfDRim}+${c}${rfDCre}+${c}${rfDDebC}+${c}${rfDDebIVA}+${c}${rfPagOneri}+${c}${rfDDebTrib}` })),
    { bold: true, formula: true });
  const rfFlussoOp = r;
  r++;

  sectionRow(ws, r, 'B) Gestione Investimenti', 6);
  r++;
  const INV_TOT = { B: '(im1+ii1)', C: '(im2+ii2)', D: '(im3+ii3)', E: '(im4+ii4)', F: '(im5+ii5)' };
  dataRow(ws, r, '(-) Investimenti mat. e immat.',
    yearCols.map((c) => ({ formula: `-${INV_TOT[c]}` })),
    { formula: true });
  const rfFlussoInv = r;
  r++;
  dataRow(ws, r, 'FLUSSO INVESTIMENTI',
    yearCols.map((c) => ({ formula: `${c}${rfFlussoInv}` })),
    { subtotal: true, formula: true });
  r++;

  sectionRow(ws, r, 'C) Gestione Finanziaria', 6);
  r++;
  const EQ_DBRB = { B: 'eq1+db1-rb1', C: 'eq2+db2-rb2', D: 'eq3+db3-rb3', E: 'eq4+db4-rb4', F: 'eq5+db5-rb5' };
  dataRow(ws, r, 'Apporti equity / debito - rimborsi',
    yearCols.map((c) => ({ formula: EQ_DBRB[c] })),
    { formula: true });
  const rfFlussoFin = r;
  r++;
  dataRow(ws, r, 'FLUSSO FINANZIARIO',
    yearCols.map((c) => ({ formula: `${c}${rfFlussoFin}` })),
    { subtotal: true, formula: true });
  r++;

  dataRow(ws, r, 'VARIAZIONE NETTA DI CASSA',
    yearCols.map((c) => ({ formula: `${c}${rfFlussoOp}+${c}${rfFlussoInv}+${c}${rfFlussoFin}` })),
    { bold: true, formula: true });
  const rfVarCassa = r;
  r++;

  // Cassa iniziale: Anno 1 = 0, Anno y = Anno_prev_finale
  dataRow(ws, r, 'Cassa iniziale', [
    0,
    { formula: `B${r + 1}` },
    { formula: `C${r + 1}` },
    { formula: `D${r + 1}` },
    { formula: `E${r + 1}` }
  ], { formula: true });
  const rfCassaIn = r;
  r++;
  dataRow(ws, r, 'CASSA FINALE',
    yearCols.map((c) => ({ formula: `${c}${rfCassaIn}+${c}${rfVarCassa}` })),
    { bold: true, formula: true });

  return { rows: { rfFlussoOp, rfFlussoInv, rfFlussoFin, rfVarCassa } };
}

/* ------------------- INDICI ------------------- */
function buildIndici(wb, R, ceMeta, spMeta, rfMeta) {
  const ws = wb.addWorksheet('Indici', { properties: { tabColor: { argb: COL.ingDeep } } });
  ws.getColumn('A').width = 42;
  yearCols.forEach((c) => (ws.getColumn(c).width = 15));
  ws.views = [{ state: 'frozen', xSplit: 1, ySplit: 4 }];

  ws.getCell('A1').value = 'Indici di Bilancio & Valutazione del Progetto';
  ws.getCell('A1').font = { name: 'Calibri', size: 14, bold: true, color: { argb: COL.ingDeep } };
  ws.mergeCells('A1:F1');

  headerRow(ws, 3, ['Indice', 'Anno 1', 'Anno 2', 'Anno 3', 'Anno 4', 'Anno 5']);

  // SP colonne anni 1-5 → C..G
  const SP_C = ['C', 'D', 'E', 'F', 'G'];

  let r = 4;
  sectionRow(ws, r, 'Redditività', 6); r++;
  dataRow(ws, r, 'EBITDA margin',
    yearCols.map((c) => ({ formula: `IFERROR('Conto Economico'!${c}${ceMeta.rows.rowEbitda}/'Conto Economico'!${c}${ceMeta.rows.rowRic},0)` })),
    { fmt: FMT.pct, formula: true }); r++;
  dataRow(ws, r, 'EBIT margin',
    yearCols.map((c) => ({ formula: `IFERROR('Conto Economico'!${c}${ceMeta.rows.rowEbit}/'Conto Economico'!${c}${ceMeta.rows.rowRic},0)` })),
    { fmt: FMT.pct, formula: true }); r++;
  dataRow(ws, r, 'ROS (Utile/Ricavi)',
    yearCols.map((c) => ({ formula: `IFERROR('Conto Economico'!${c}${ceMeta.rows.rowUtile}/'Conto Economico'!${c}${ceMeta.rows.rowRic},0)` })),
    { fmt: FMT.pct, formula: true }); r++;
  dataRow(ws, r, 'ROI (EBIT/CapInv.)',
    yearCols.map((c, i) => {
      const sp = SP_C[i];
      return { formula: `IFERROR('Conto Economico'!${c}${ceMeta.rows.rowEbit}/('Stato Patrimoniale'!${sp}${spMeta.rows.rowTotImm}+'Stato Patrimoniale'!${sp}${spMeta.rows.rowRim}+'Stato Patrimoniale'!${sp}${spMeta.rows.rowCre}-'Stato Patrimoniale'!${sp}${spMeta.rows.rowDebComm}),0)` };
    }),
    { fmt: FMT.pct, formula: true }); r++;
  dataRow(ws, r, 'ROE (Utile/PN)',
    yearCols.map((c, i) => {
      const sp = SP_C[i];
      return { formula: `IFERROR('Conto Economico'!${c}${ceMeta.rows.rowUtile}/'Stato Patrimoniale'!${sp}${spMeta.rows.rowPN},0)` };
    }),
    { fmt: FMT.pct, formula: true }); r++;

  sectionRow(ws, r, 'Solidità', 6); r++;
  dataRow(ws, r, 'PFN (Debito - Cassa)',
    yearCols.map((c, i) => {
      const sp = SP_C[i];
      return { formula: `'Stato Patrimoniale'!${sp}${spMeta.rows.rowDebFin}-'Stato Patrimoniale'!${sp}${spMeta.rows.rowCassa}` };
    }),
    { formula: true }); r++;
  const pfnEbitdaRow = r;
  dataRow(ws, r, 'PFN / EBITDA',
    yearCols.map((c, i) => {
      const sp = SP_C[i];
      return { formula: `IFERROR(('Stato Patrimoniale'!${sp}${spMeta.rows.rowDebFin}-'Stato Patrimoniale'!${sp}${spMeta.rows.rowCassa})/'Conto Economico'!${c}${ceMeta.rows.rowEbitda},0)` };
    }),
    { fmt: FMT.num2, formula: true }); r++;
  dataRow(ws, r, 'Leva finanziaria (Deb/PN)',
    yearCols.map((c, i) => {
      const sp = SP_C[i];
      return { formula: `IFERROR('Stato Patrimoniale'!${sp}${spMeta.rows.rowDebFin}/'Stato Patrimoniale'!${sp}${spMeta.rows.rowPN},0)` };
    }),
    { fmt: FMT.num2, formula: true }); r++;
  dataRow(ws, r, 'Indipendenza finanziaria (PN/Att)',
    yearCols.map((c, i) => {
      const sp = SP_C[i];
      return { formula: `IFERROR('Stato Patrimoniale'!${sp}${spMeta.rows.rowPN}/'Stato Patrimoniale'!${sp}${spMeta.rows.rowTotAtt},0)` };
    }),
    { fmt: FMT.pct, formula: true }); r++;

  sectionRow(ws, r, 'Liquidità', 6); r++;
  dataRow(ws, r, 'Current ratio',
    yearCols.map((c, i) => {
      const sp = SP_C[i];
      return { formula: `IFERROR('Stato Patrimoniale'!${sp}${spMeta.rows.rowTotAC}/('Stato Patrimoniale'!${sp}${spMeta.rows.rowDebComm}+'Stato Patrimoniale'!${sp}${spMeta.rows.rowDebIVA}+'Stato Patrimoniale'!${sp}${spMeta.rows.rowDebTrib}),0)` };
    }),
    { fmt: FMT.num2, formula: true }); r++;
  dataRow(ws, r, 'Quick ratio',
    yearCols.map((c, i) => {
      const sp = SP_C[i];
      return { formula: `IFERROR(('Stato Patrimoniale'!${sp}${spMeta.rows.rowTotAC}-'Stato Patrimoniale'!${sp}${spMeta.rows.rowRim})/('Stato Patrimoniale'!${sp}${spMeta.rows.rowDebComm}+'Stato Patrimoniale'!${sp}${spMeta.rows.rowDebIVA}+'Stato Patrimoniale'!${sp}${spMeta.rows.rowDebTrib}),0)` };
    }),
    { fmt: FMT.num2, formula: true }); r++;
  dataRow(ws, r, 'Ciclo del circolante (gg)',
    yearCols.map(() => ({ formula: 'dso+dio-dpo' })),
    { fmt: FMT.intGiorni, formula: true }); r++;

  sectionRow(ws, r, 'Sostenibilità', 6); r++;
  const dscrRow = r;
  const RB_Y = { B: 'rb1', C: 'rb2', D: 'rb3', E: 'rb4', F: 'rb5' };
  dataRow(ws, r, 'DSCR',
    yearCols.map((c) => ({ formula: `IFERROR('Rendiconto Finanziario'!${c}${rfMeta.rows.rfFlussoOp}/(${RB_Y[c]}+(-'Conto Economico'!${c}${ceMeta.rows.rowOneri})),0)` })),
    { fmt: FMT.num2, formula: true }); r++;
  dataRow(ws, r, 'Copertura Oneri Finanziari',
    yearCols.map((c) => ({ formula: `IFERROR('Conto Economico'!${c}${ceMeta.rows.rowEbitda}/(-'Conto Economico'!${c}${ceMeta.rows.rowOneri}),0)` })),
    { fmt: FMT.num2, formula: true }); r++;

  sectionRow(ws, r, 'Break-Even', 6); r++;
  dataRow(ws, r, 'Punto di pareggio (€)',
    yearCols.map((c) => ({ formula: `IFERROR(-'Conto Economico'!${c}${ceMeta.rows.rowTotCF}/('Conto Economico'!${c}${ceMeta.rows.rowMdc}/'Conto Economico'!${c}${ceMeta.rows.rowRic}),0)` })),
    { formula: true }); r++;
  dataRow(ws, r, 'Margine di sicurezza',
    yearCols.map((c) => ({ formula: `IFERROR(('Conto Economico'!${c}${ceMeta.rows.rowRic}-IFERROR(-'Conto Economico'!${c}${ceMeta.rows.rowTotCF}/('Conto Economico'!${c}${ceMeta.rows.rowMdc}/'Conto Economico'!${c}${ceMeta.rows.rowRic}),0))/'Conto Economico'!${c}${ceMeta.rows.rowRic},0)` })),
    { fmt: FMT.pct, formula: true }); r++;

  sectionRow(ws, r, 'Valutazione del Progetto', 6); r++;
  // FCF = flussoOp + flussoInv (entrambi presi da RF)
  dataRow(ws, r, 'Free Cash Flow',
    yearCols.map((c) => ({ formula: `'Rendiconto Finanziario'!${c}${rfMeta.rows.rfFlussoOp}+'Rendiconto Finanziario'!${c}${rfMeta.rows.rfFlussoInv}` })),
    { formula: true });
  const rowFCF = r;
  r++;
  dataRow(ws, r, 'FCF cumulato', [
    { formula: `B${rowFCF}` },
    { formula: `B${r}+C${rowFCF}` },
    { formula: `C${r}+D${rowFCF}` },
    { formula: `D${r}+E${rowFCF}` },
    { formula: `E${r}+F${rowFCF}` }
  ], { formula: true });
  const rowFCFcum = r;
  r++;
  // VAN: NPV(wacc, FCF range)
  ws.getCell(r, 1).value = 'VAN @ WACC';
  ws.getCell(r, 1).font = { name: 'Calibri', size: 10, bold: true };
  ws.getCell(r, 1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COL.total } };
  const cellVAN = ws.getCell(r, 2);
  cellVAN.value = { formula: `NPV(wacc,B${rowFCF}:F${rowFCF})` };
  cellVAN.numFmt = FMT.euro;
  cellVAN.font = { name: 'Calibri', size: 10, bold: true };
  cellVAN.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COL.total } };
  cellVAN.alignment = { horizontal: 'right' };
  ws.mergeCells(r, 2, r, 6);
  r++;

  // TIR (statico)
  ws.getCell(r, 1).value = 'TIR (IRR)';
  ws.getCell(r, 1).font = { name: 'Calibri', size: 10, bold: true };
  const cellIRR = ws.getCell(r, 2);
  cellIRR.value = isFinite(R.IRR) ? R.IRR : null;
  cellIRR.numFmt = FMT.pct1;
  cellIRR.font = { name: 'Calibri', size: 10, bold: true };
  cellIRR.alignment = { horizontal: 'right' };
  ws.mergeCells(r, 2, r, 6);
  r++;

  // Payback (statico)
  ws.getCell(r, 1).value = 'Payback';
  ws.getCell(r, 1).font = { name: 'Calibri', size: 10, bold: true };
  const cellPB = ws.getCell(r, 2);
  cellPB.value = R.payback ? `Anno ${R.payback}` : 'oltre Y5';
  cellPB.font = { name: 'Calibri', size: 10, bold: true };
  cellPB.alignment = { horizontal: 'right' };
  ws.mergeCells(r, 2, r, 6);
  r++;

  // Conditional formatting per DSCR e PFN/EBITDA
  ws.addConditionalFormatting({
    ref: `B${dscrRow}:F${dscrRow}`,
    rules: [
      { type: 'cellIs', operator: 'greaterThanOrEqual', formulae: [1.3], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: COL.good } }, font: { color: { argb: COL.goodFg } } } },
      { type: 'cellIs', operator: 'between', formulae: [1.0, 1.3], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: COL.warn } }, font: { color: { argb: COL.warnFg } } } },
      { type: 'cellIs', operator: 'lessThan', formulae: [1.0], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: COL.bad } }, font: { color: { argb: COL.badFg } } } }
    ]
  });
  ws.addConditionalFormatting({
    ref: `B${pfnEbitdaRow}:F${pfnEbitdaRow}`,
    rules: [
      { type: 'cellIs', operator: 'lessThanOrEqual', formulae: [3], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: COL.good } }, font: { color: { argb: COL.goodFg } } } },
      { type: 'cellIs', operator: 'between', formulae: [3, 5], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: COL.warn } }, font: { color: { argb: COL.warnFg } } } },
      { type: 'cellIs', operator: 'greaterThan', formulae: [5], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: COL.bad } }, font: { color: { argb: COL.badFg } } } }
    ]
  });

  return { rows: { rowFCF, rowFCFcum, dscrRow, pfnEbitdaRow } };
}

/* ------------------- GRAFICI (placeholder sheet for native chart injection) ---- */
function buildGraficiPlaceholder(wb) {
  const ws = wb.addWorksheet('Grafici', { properties: { tabColor: { argb: COL.teal } } });
  ws.views = [{ showGridLines: false }];
  ws.getColumn('A').width = 2;
  for (let c = 2; c <= 18; c++) ws.getColumn(c).width = 11;
  // Force min row count so cell anchors land on existing cells
  for (let row = 1; row <= 50; row++) ws.getRow(row).height = 18;

  ws.getCell('B2').value = 'Grafici riepilogativi';
  ws.getCell('B2').font = { name: 'Calibri', size: 14, bold: true, color: { argb: COL.ingDeep } };
  ws.getCell('B3').value = 'Grafici nativi Excel — collegati ai dati di CE / SP / Indici. Modifica le ipotesi e i grafici si aggiornano automaticamente.';
  ws.getCell('B3').font = { name: 'Calibri', size: 9, italic: true, color: { argb: COL.muted } };
  return ws;
}

/* ------------------- ENTRYPOINT ------------------- */
export async function exportToExcel(state, _chartImages = [], logoBase64 = null) {
  const R = compute(state);

  const wb = new ExcelJS.Workbook();
  wb.creator = 'Ingenia · Simulatore BP';
  wb.lastModifiedBy = 'Ingenia';
  wb.created = new Date();
  wb.modified = new Date();
  wb.title = 'Business Plan Startup — 5 anni';

  buildCover(wb, state, R, logoBase64);
  buildIpotesi(wb, state);

  const ceMeta = buildCE(wb);
  const spMeta = buildSP(wb, ceMeta);
  const rfMeta = buildRF(wb, ceMeta, spMeta);

  // Risolvi placeholder cross-sheet
  const ceWs = wb.getWorksheet('Conto Economico');
  for (let i = 1; i < Y; i++) {
    const colLetter = yearCols[i];
    const cell = ceWs.getCell(`${colLetter}${ceMeta.rows.rowOneri}`);
    if (cell.value && cell.value.formula) {
      cell.value = { formula: cell.value.formula.replace('__DEBFIN_ROW__', String(spMeta.rows.rowDebFin)) };
    }
  }
  const spWs = wb.getWorksheet('Stato Patrimoniale');
  ['C', 'D', 'E', 'F', 'G'].forEach((c) => {
    const cell = spWs.getCell(`${c}${spMeta.rows.rowCassa}`);
    if (cell.value && cell.value.formula) {
      cell.value = { formula: cell.value.formula.replace('__RF_VARCASSA_ROW__', String(rfMeta.rows.rfVarCassa)) };
    }
  });

  const idMeta = buildIndici(wb, R, ceMeta, spMeta, rfMeta);
  buildGraficiPlaceholder(wb);

  const buffer = await wb.xlsx.writeBuffer();

  // -------- Native chart injection ----------
  const yearLabels = ['Anno 1', 'Anno 2', 'Anno 3', 'Anno 4', 'Anno 5'];
  const HEADER_ROW = 3;
  const palette = CHART_PALETTE;

  const chartXmls = [
    {
      xml: makeChartXml_RicaviEbitdaUtile({
        ceSheetName: 'Conto Economico',
        headerRow: HEADER_ROW,
        rowRic: ceMeta.rows.rowRic,
        rowEbitda: ceMeta.rows.rowEbitda,
        rowUtile: ceMeta.rows.rowUtile,
        yearLabels,
        ricavi: R.CE.map((c) => c.ricavi),
        ebitda: R.CE.map((c) => c.ebitda),
        utile:  R.CE.map((c) => c.utile),
        palette
      }),
      position: { fromCol: 1, fromRow: 5,  toCol: 9,  toRow: 22 }
    },
    {
      xml: makeChartXml_FCF({
        indSheetName: 'Indici',
        headerRow: HEADER_ROW,
        rowFCF: idMeta.rows.rowFCF,
        rowFCFcum: idMeta.rows.rowFCFcum,
        yearLabels,
        fcf: R.FCF,
        fcfcum: R.FCFcum,
        palette
      }),
      position: { fromCol: 9, fromRow: 5,  toCol: 17, toRow: 22 }
    },
    {
      xml: makeChartXml_StatoPatrimoniale({
        spSheetName: 'Stato Patrimoniale',
        headerRow: HEADER_ROW,
        rowPN:    spMeta.rows.rowPN,
        rowDebFin: spMeta.rows.rowDebFin,
        rowDebTot: spMeta.rows.rowDebOther,
        rowCassa: spMeta.rows.rowCassa,
        yearLabels,
        pn:     R.SP.map((s) => s.PN),
        debFin: R.SP.map((s) => s.debFin),
        debTot: R.SP.map((s) => s.debComm + s.debIVA + s.debTrib),
        cassa:  R.SP.map((s) => s.cassa),
        palette
      }),
      position: { fromCol: 1, fromRow: 24, toCol: 9,  toRow: 41 }
    },
    {
      xml: makeChartXml_DSCRpfn({
        indSheetName: 'Indici',
        headerRow: HEADER_ROW,
        rowDSCR: idMeta.rows.dscrRow,
        rowPFNebitda: idMeta.rows.pfnEbitdaRow,
        yearLabels,
        dscr: R.ID.map((d) => d.DSCR),
        pfnEbitda: R.ID.map((d) => d.pfnEbitda),
        palette
      }),
      position: { fromCol: 9, fromRow: 24, toCol: 17, toRow: 41 }
    }
  ];

  const blob = await injectNativeCharts(buffer, {
    targetSheetName: 'Grafici',
    charts: chartXmls
  });

  const today = new Date().toISOString().slice(0, 10);
  saveAs(blob, `BusinessPlan_Ingenia_${today}.xlsx`);
}
