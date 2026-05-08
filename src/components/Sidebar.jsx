import NumberField from './NumberField.jsx';

export default function Sidebar({ inputs, onField, onRecalc, onReset, onClear, onPrint, onExportExcel, exporting }) {
  const f = (name, label, step = 1) => (
    <NumberField name={name} label={label} step={step} value={inputs[name]} onChange={onField} />
  );

  return (
    <aside className="panel controls" id="controls">
      <h2><span className="num">1</span>Ipotesi base &amp; Crescita</h2>

      <details open>
        <summary>Ricavi — Volumi e Prezzi (Anno 1)</summary>
        <div className="body">
          <div className="grid2">
            {f('i_vol', 'Volumi venduti (pz)', 100)}
            {f('i_prc', 'Prezzo unitario (€)', 1)}
            {f('g_vol', 'Crescita volumi %/anno', 1)}
            {f('g_prc', 'Crescita prezzo %/anno', 0.5)}
          </div>
        </div>
      </details>

      <details open>
        <summary>Costi Variabili</summary>
        <div className="body">
          <div className="grid2">
            {f('i_cuv', 'Costo unit. variab. (€/pz)', 1)}
            {f('g_cuv', 'Crescita costo unit. %/anno', 0.5)}
          </div>
        </div>
      </details>

      <details>
        <summary>Costi Fissi Operativi (Anno 1)</summary>
        <div className="body">
          <div className="grid2">
            {f('i_pers', 'Personale (€)', 1000)}
            {f('i_aff', 'Affitto / utenze (€)', 500)}
            {f('i_mkt', 'Marketing (€)', 500)}
            {f('i_serv', 'Servizi/consulenze (€)', 500)}
            {f('i_alt', 'Altri costi fissi (€)', 500)}
            {f('g_cf', 'Crescita costi fissi %/anno', 0.5)}
          </div>
        </div>
      </details>

      <details>
        <summary>Investimenti &amp; Ammortamenti</summary>
        <div className="body">
          <h3>Materiali</h3>
          <div className="grid5">
            {f('im1', 'Y1', 1000)}
            {f('im2', 'Y2', 1000)}
            {f('im3', 'Y3', 1000)}
            {f('im4', 'Y4', 1000)}
            {f('im5', 'Y5', 1000)}
          </div>
          <label style={{ marginTop: 6 }}>Aliquota amm. materiali %</label>
          <input
            type="number"
            value={inputs.aliq_m}
            step={1}
            onChange={(e) => onField('aliq_m', e.target.value)}
          />
          <h3>Immateriali</h3>
          <div className="grid5">
            {f('ii1', 'Y1', 500)}
            {f('ii2', 'Y2', 500)}
            {f('ii3', 'Y3', 500)}
            {f('ii4', 'Y4', 500)}
            {f('ii5', 'Y5', 500)}
          </div>
          <label style={{ marginTop: 6 }}>Aliquota amm. immateriali %</label>
          <input
            type="number"
            value={inputs.aliq_i}
            step={1}
            onChange={(e) => onField('aliq_i', e.target.value)}
          />
        </div>
      </details>

      <details>
        <summary>Capitale Circolante (giorni)</summary>
        <div className="body">
          <div className="grid3">
            {f('dso', 'DSO clienti', 5)}
            {f('dpo', 'DPO fornitori', 5)}
            {f('dio', 'DIO scorte', 5)}
          </div>
        </div>
      </details>

      <details>
        <summary>IVA</summary>
        <div className="body">
          <div className="grid3">
            {f('iva_v', 'IVA vendite %', 1)}
            {f('iva_a', 'IVA acquisti %', 1)}
            {f('iva_p', 'Periodicità (mese)', 1)}
          </div>
        </div>
      </details>

      <details>
        <summary>Finanziamenti &amp; Equity</summary>
        <div className="body">
          <h3>Apporto Equity (€)</h3>
          <div className="grid5">
            {f('eq1', 'Y1', 1000)}{f('eq2', 'Y2')}{f('eq3', 'Y3')}{f('eq4', 'Y4')}{f('eq5', 'Y5')}
          </div>
          <h3>Nuovo Debito Bancario (€)</h3>
          <div className="grid5">
            {f('db1', 'Y1', 1000)}{f('db2', 'Y2')}{f('db3', 'Y3')}{f('db4', 'Y4')}{f('db5', 'Y5')}
          </div>
          <h3>Rimborso Quota Capitale (€)</h3>
          <div className="grid5">
            {f('rb1', 'Y1')}{f('rb2', 'Y2')}{f('rb3', 'Y3')}{f('rb4', 'Y4')}{f('rb5', 'Y5')}
          </div>
          <label style={{ marginTop: 6 }}>Tasso interesse sul debito %</label>
          <input
            type="number"
            value={inputs.tasso}
            step={0.1}
            onChange={(e) => onField('tasso', e.target.value)}
          />
        </div>
      </details>

      <details>
        <summary>Fiscalità &amp; WACC</summary>
        <div className="body">
          <div className="grid3">
            {f('ires', 'Aliquota IRES %', 0.5)}
            {f('irap', 'Aliquota IRAP %', 0.1)}
            {f('wacc', 'WACC %', 0.5)}
          </div>
        </div>
      </details>

      <div className="btn-row">
        <button className="btn btn-primary" onClick={onRecalc}>Ricalcola</button>
        <button className="btn btn-ghost" onClick={onReset}>Reset esempio</button>
        <button className="btn btn-danger" title="Azzera tutti i campi" onClick={onClear}>Azzera tutti gli input</button>
        <button className="btn btn-ghost no-print" onClick={onPrint}>Stampa / PDF</button>
        <button
          className="btn btn-export no-print"
          onClick={onExportExcel}
          disabled={exporting}
          title="Esporta workbook Excel con formule live, formattazione professionale, grafici e indici"
        >
          {exporting ? 'Esportazione…' : 'Esporta Excel'}
        </button>
      </div>
      <p className="help" style={{ marginTop: 10 }}>
        I calcoli si aggiornano automaticamente ad ogni modifica. Tutti i valori sono in € a parità di potere d&apos;acquisto. Modello didattico — non sostituisce un parere professionale.
      </p>
      <p className="help" style={{ marginTop: 6 }}>
        <strong>Reset esempio</strong>: ripristina i valori dell&apos;esempio iniziale. <strong>Azzera tutti gli input</strong>: porta a 0 tutti i campi per partire da una pagina bianca con i tuoi dati.
      </p>
    </aside>
  );
}
