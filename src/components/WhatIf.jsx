import { PRESETS } from '../lib/defaults.js';

const at = (v) => (v === 0 ? '0' : (v > 0 ? '+' : '') + v.toFixed(0) + '%');

function Slider({ label, name, value, onField }) {
  return (
    <div className="slider-card">
      <div className="slider-head">
        <span>{label}</span>
        <span className="val">{at(Number(value) || 0)}</span>
      </div>
      <input
        type="range"
        min={-50}
        max={50}
        step={1}
        value={value ?? 0}
        onChange={(e) => onField(name, e.target.value)}
      />
      <div className="slider-foot"><span>-50%</span><span>0</span><span>+50%</span></div>
    </div>
  );
}

export default function WhatIf({ inputs, onField, onPreset }) {
  const tag = inputs.s_prc || inputs.s_vol || inputs.s_cuv || inputs.s_cf
    ? `Prezzo ${at(+inputs.s_prc)} · Volumi ${at(+inputs.s_vol)} · Cuv ${at(+inputs.s_cuv)} · CF ${at(+inputs.s_cf)}`
    : 'Caso Base';

  return (
    <section className="panel" style={{ marginBottom: 18 }}>
      <h2>
        <span className="num">2</span>What-if — Scenari Rapidi <span className="scenario-tag">{tag}</span>
      </h2>
      <div className="grid2">
        <Slider label="Prezzo di vendita" name="s_prc" value={inputs.s_prc} onField={onField} />
        <Slider label="Volumi venduti" name="s_vol" value={inputs.s_vol} onField={onField} />
        <Slider label="Costo variabile unitario" name="s_cuv" value={inputs.s_cuv} onField={onField} />
        <Slider label="Costi fissi totali" name="s_cf" value={inputs.s_cf} onField={onField} />
      </div>
      <div className="btn-row">
        {Object.keys(PRESETS).map((k) => (
          <button key={k} className="btn btn-ghost" onClick={() => onPreset(k)}>
            {k === 'base' && 'Caso Base'}
            {k === 'best' && 'Best Case (+15% prezzo, +20% vol)'}
            {k === 'worst' && 'Worst Case (-15% prezzo, -20% vol)'}
            {k === 'cost' && 'Stress Costi (+15% costi)'}
          </button>
        ))}
      </div>
    </section>
  );
}
