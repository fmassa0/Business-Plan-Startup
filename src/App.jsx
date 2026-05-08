import { useMemo, useState } from 'react';
import { DEFAULTS, ZEROES, PRESETS } from './lib/defaults.js';
import { compute } from './lib/compute.js';
import Sidebar from './components/Sidebar.jsx';
import WhatIf from './components/WhatIf.jsx';
import KPIRow from './components/KPIRow.jsx';
import ChartGrid from './components/ChartGrid.jsx';
import { TableCE, TableSP, TableRF, TableID } from './components/Tables.jsx';

const BASE = import.meta.env.BASE_URL;

async function svgUrlToPngBase64(svgUrl, width, height) {
  const res = await fetch(svgUrl);
  const svgText = await res.text();
  const blob = new Blob([svgText], { type: 'image/svg+xml;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  try {
    const img = new Image();
    img.crossOrigin = 'anonymous';
    await new Promise((resolve, reject) => {
      img.onload = resolve;
      img.onerror = reject;
      img.src = url;
    });
    const canvas = document.createElement('canvas');
    const dpr = 2;
    canvas.width = width * dpr;
    canvas.height = height * dpr;
    const ctx = canvas.getContext('2d');
    ctx.scale(dpr, dpr);
    ctx.drawImage(img, 0, 0, width, height);
    return canvas.toDataURL('image/png').split(',')[1];
  } finally {
    URL.revokeObjectURL(url);
  }
}

function captureChartImages() {
  const nodes = document.querySelectorAll('.chart-box canvas');
  const images = [];
  nodes.forEach((c) => {
    try {
      images.push(c.toDataURL('image/png').split(',')[1]);
    } catch {
      images.push(null);
    }
  });
  return images;
}

export default function App() {
  const [inputs, setInputs] = useState(DEFAULTS);
  const [exporting, setExporting] = useState(false);

  const onField = (name, value) => {
    setInputs((s) => ({ ...s, [name]: value }));
  };

  const onPreset = (key) => {
    const p = PRESETS[key];
    setInputs((s) => ({ ...s, ...p }));
  };

  const onReset = () => setInputs(DEFAULTS);

  const onClear = () => {
    const ok = window.confirm(
      'Vuoi azzerare TUTTI gli input?\n\n' +
        'Tutti i valori della colonna di sinistra verranno portati a 0 ' +
        '(volumi, prezzi, costi, investimenti, equity, debiti, aliquote, ecc.) ' +
        'e gli slider What-if torneranno a 0%.\n\n' +
        "L'operazione non può essere annullata. Per ripristinare i valori dell'esempio iniziale usa 'Reset esempio'."
    );
    if (!ok) return;
    setInputs(ZEROES);
  };

  const onPrint = () => window.print();
  const onRecalc = () => setInputs((s) => ({ ...s }));

  const onExportExcel = async () => {
    if (exporting) return;
    setExporting(true);
    try {
      const charts = captureChartImages();
      const logo = await svgUrlToPngBase64(`${BASE}ingenia-logo-light.svg`, 360, 110).catch(() => null);
      const { exportToExcel } = await import('./lib/excelExport.js');
      await exportToExcel(inputs, charts, logo);
    } catch (err) {
      console.error('Errore export Excel:', err);
      window.alert('Errore durante l\'export Excel. Controlla la console del browser.');
    } finally {
      setExporting(false);
    }
  };

  const R = useMemo(() => compute(inputs), [inputs]);

  return (
    <>
      <header className="app-header">
        <a href="https://ingenia.cloud" target="_blank" rel="noopener" className="brand" title="Ingenia · ingenia.cloud">
          <img src={`${BASE}ingenia-logo-light.svg`} alt="Ingenia · Powering Future" className="brand-logo" />
        </a>
        <h1>Simulatore Start Up — Dashboard Interattiva (5 anni)</h1>
        <p>
          Inserisci i valori base e usa gli slider per simulare scenari what-if in tempo reale.
          Conto Economico, Stato Patrimoniale, Rendiconto Finanziario e Indici si ricalcolano automaticamente.
        </p>
      </header>

      <div className="container">
        <Sidebar
          inputs={inputs}
          onField={onField}
          onRecalc={onRecalc}
          onReset={onReset}
          onClear={onClear}
          onPrint={onPrint}
          onExportExcel={onExportExcel}
          exporting={exporting}
        />

        <main>
          <WhatIf inputs={inputs} onField={onField} onPreset={onPreset} />
          <KPIRow R={R} />
          <ChartGrid R={R} />

          <h2 className="section-title">Conto Economico Previsionale</h2>
          <div className="table-wrap"><TableCE R={R} /></div>

          <h2 className="section-title">Stato Patrimoniale Previsionale</h2>
          <div className="table-wrap"><TableSP R={R} /></div>

          <h2 className="section-title">Rendiconto Finanziario (metodo indiretto)</h2>
          <div className="table-wrap"><TableRF R={R} /></div>

          <h2 className="section-title">Indici di Bilancio &amp; Valutazione del Progetto</h2>
          <div className="table-wrap"><TableID R={R} /></div>
        </main>
      </div>

      <div className="footer no-print">
        Modello didattico basato sulla struttura "Simulatore Start Up". Logica di calcolo: ipotesi → CE → SP → RF → indici.
        Per banche, bandi (PNRR/Invitalia/Simest), investitori — sempre da validare con un consulente. © Ingenia.
      </div>
    </>
  );
}
