# Business Plan Startup — Ingenia

Dashboard interattiva (React + Vite) per la costruzione di un Business Plan a 5 anni: Conto Economico, Stato Patrimoniale, Rendiconto Finanziario e Indici di bilancio. 

## Quick start

```powershell
npm install
npm run dev
```

Apre http://localhost:5173 con hot-reload.

## Build di produzione

```powershell
npm run build
npm run preview
```

L'output statico è in `dist/` ed è pubblicabile su qualsiasi static host (Netlify, Vercel, GitHub Pages, S3).

## Struttura

```
src/
  App.jsx               composizione pagina
  main.jsx              entry React
  styles.css            palette Ingenia + layout
  lib/
    compute.js          motore di calcolo (CE/SP/RF/Indici, NPV, IRR)
    defaults.js         valori di default + preset what-if
    format.js           formatter Intl per €/% /numeri
  components/
    Sidebar.jsx         input ipotesi (collapsible)
    WhatIf.jsx          slider scenari + preset
    KPIRow.jsx          card metriche di sintesi
    ChartGrid.jsx       4 grafici Chart.js
    Tables.jsx          tabelle CE / SP / RF / Indici
    NumberField.jsx     input numerico riusabile
```

## Note

Modello didattico — non sostituisce un parere professionale. Il file `reference_original.html` (incluso solo per riferimento) è la versione standalone su cui si basa la riscrittura.
