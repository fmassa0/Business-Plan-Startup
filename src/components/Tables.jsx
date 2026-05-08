import { Y } from '../lib/defaults.js';
import { fmt, fmtE, fmtP } from '../lib/format.js';

const yLabels  = ['Anno 1', 'Anno 2', 'Anno 3', 'Anno 4', 'Anno 5'];
const yLabels0 = ['Anno 0', ...yLabels];

function Header({ labels }) {
  return (
    <thead>
      <tr>
        <th>Voce</th>
        {labels.map((l) => <th key={l}>{l}</th>)}
      </tr>
    </thead>
  );
}

function valCell(v, fmtType) {
  if (v && typeof v === 'object' && 'txt' in v) return v.txt;
  if (fmtType === 'pct') return fmtP(v);
  if (fmtType === 'num') return fmt(v, 2);
  return fmtE(v);
}

function Row({ label, vals, cls = '', fmtType }) {
  return (
    <tr className={cls}>
      <td>{label}</td>
      {vals.map((v, i) => {
        const isNeg = typeof v === 'number' && v < 0 && fmtType !== 'pct';
        return <td key={i} className={isNeg ? 'neg' : ''}>{valCell(v, fmtType)}</td>;
      })}
    </tr>
  );
}

function flagPill(value, type) {
  if (type === 'dscr') {
    if (value >= 1.3) return <span className="pill pill-ok">Ok</span>;
    if (value >= 1.0) return <span className="pill pill-warn">Marg.</span>;
    return <span className="pill pill-bad">Critico</span>;
  }
  if (type === 'pfnEbitda') {
    if (value <= 3) return <span className="pill pill-ok">Sost.</span>;
    if (value <= 5) return <span className="pill pill-warn">Tirato</span>;
    return <span className="pill pill-bad">Eccessivo</span>;
  }
  return null;
}

export function TableCE({ R }) {
  return (
    <table>
      <Header labels={yLabels} />
      <tbody>
        <Row label="Ricavi delle vendite" vals={R.CE.map((c) => c.ricavi)} />
        <Row label="(-) Costi variabili" vals={R.CE.map((c) => c.cVar)} />
        <Row label="MARGINE DI CONTRIBUZIONE" vals={R.CE.map((c) => c.mdc)} cls="subtotal" />
        <Row label="MdC %" vals={R.CE.map((c) => c.mdcPct)} fmtType="pct" cls="muted" />
        <tr className="section"><td colSpan={Y + 1}>Costi Fissi</td></tr>
        <Row label="(-) Personale"          vals={R.CE.map((c) => c.pers)} cls="indent" />
        <Row label="(-) Affitto / utenze"   vals={R.CE.map((c) => c.aff)}  cls="indent" />
        <Row label="(-) Marketing / vendite" vals={R.CE.map((c) => c.mkt)}  cls="indent" />
        <Row label="(-) Servizi e consulenze" vals={R.CE.map((c) => c.serv)} cls="indent" />
        <Row label="(-) Altri costi fissi"  vals={R.CE.map((c) => c.alt)}  cls="indent" />
        <Row label="Totale costi fissi"     vals={R.CE.map((c) => c.totCF)} cls="subtotal" />
        <Row label="EBITDA (MOL)"           vals={R.CE.map((c) => c.ebitda)} cls="bold" />
        <Row label="EBITDA margin %"        vals={R.CE.map((c) => c.ebitdaMar)} fmtType="pct" cls="muted" />
        <Row label="(-) Ammortamenti materiali"   vals={R.CE.map((c) => c.ammMat)} />
        <Row label="(-) Ammortamenti immateriali" vals={R.CE.map((c) => c.ammImm)} />
        <Row label="EBIT (Reddito Operativo)" vals={R.CE.map((c) => c.ebit)} cls="bold" />
        <Row label="EBIT margin %"          vals={R.CE.map((c) => c.ebitMar)} fmtType="pct" cls="muted" />
        <Row label="(-) Oneri finanziari"   vals={R.CE.map((c) => c.oneriFin)} />
        <Row label="Risultato ante imposte" vals={R.CE.map((c) => c.rai)} cls="subtotal" />
        <Row label="(-) Imposte (IRES+IRAP)" vals={R.CE.map((c) => c.imposte)} />
        <Row label="UTILE / (PERDITA) NETTA" vals={R.CE.map((c) => c.utile)} cls="bold" />
      </tbody>
    </table>
  );
}

export function TableSP({ R }) {
  const z = fmtE(0);
  const cellsAfterZero = (key) => (
    <>
      <td>{z}</td>
      {R.SP.map((s, i) => <td key={i}>{fmtE(s[key])}</td>)}
    </>
  );
  return (
    <table>
      <Header labels={yLabels0} />
      <tbody>
        <tr className="section"><td colSpan={Y + 2}>ATTIVO</td></tr>
        <tr><td>Imm. materiali (lordo)</td>{cellsAfterZero('immMatLordo')}</tr>
        <tr><td>F.do amm. materiali</td>{cellsAfterZero('fAmmMat')}</tr>
        <tr className="subtotal"><td>Imm. materiali nette</td>{cellsAfterZero('immMatNetto')}</tr>
        <tr><td>Imm. immateriali (lordo)</td>{cellsAfterZero('immImmatLordo')}</tr>
        <tr><td>F.do amm. immateriali</td>{cellsAfterZero('fAmmImm')}</tr>
        <tr className="subtotal"><td>Imm. immateriali nette</td>{cellsAfterZero('immImmatNetto')}</tr>
        <tr className="bold"><td>TOTALE ATTIVO IMMOBILIZZATO</td>{cellsAfterZero('totImm')}</tr>
        <tr><td>Rimanenze</td>{cellsAfterZero('rimanenze')}</tr>
        <tr><td>Crediti commerciali (IVA incl.)</td>{cellsAfterZero('crediti')}</tr>
        <tr><td>Liquidità (cassa e banche)</td>{cellsAfterZero('cassa')}</tr>
        <tr className="subtotal"><td>TOTALE ATTIVO CIRCOLANTE</td>{cellsAfterZero('totAttivoCirc')}</tr>
        <tr className="bold"><td>TOTALE ATTIVO</td>{cellsAfterZero('totAttivo')}</tr>
        <tr className="section"><td colSpan={Y + 2}>PASSIVO E PATRIMONIO NETTO</td></tr>
        <tr><td>Capitale sociale</td>{cellsAfterZero('cap')}</tr>
        <tr><td>Utili / (perdite) a nuovo</td>{cellsAfterZero('utNuovo')}</tr>
        <tr><td>Utile / (perdita) d&apos;esercizio</td>{cellsAfterZero('utile')}</tr>
        <tr className="bold"><td>TOTALE PATRIMONIO NETTO</td>{cellsAfterZero('PN')}</tr>
        <tr><td>Debiti finanziari</td>{cellsAfterZero('debFin')}</tr>
        <tr><td>Debiti commerciali (IVA incl.)</td>{cellsAfterZero('debComm')}</tr>
        <tr><td>Debito IVA verso Erario</td>{cellsAfterZero('debIVA')}</tr>
        <tr><td>Debiti tributari</td>{cellsAfterZero('debTrib')}</tr>
        <tr className="subtotal"><td>TOTALE PASSIVO</td>{cellsAfterZero('totPassivo')}</tr>
        <tr className="bold"><td>TOTALE P + PN</td>{cellsAfterZero('totPP')}</tr>
        <tr className="muted"><td>Quadratura (Att - P&amp;PN)</td>{cellsAfterZero('quad')}</tr>
      </tbody>
    </table>
  );
}

export function TableRF({ R }) {
  return (
    <table>
      <Header labels={yLabels} />
      <tbody>
        <tr className="section"><td colSpan={Y + 1}>A) Gestione Operativa</td></tr>
        <Row label="Utile / (perdita) d'esercizio" vals={R.RF.map((r) => r.utile)} />
        <Row label="(+) Ammortamenti (sterilizzati)" vals={R.RF.map((r) => r.amm)} />
        <Row label="(+) Oneri finanziari (sterilizzati)" vals={R.RF.map((r) => r.oneriSter)} />
        <Row label="Flusso prima del CCN" vals={R.RF.map((r) => r.preCCN)} cls="subtotal" />
        <Row label="(-) Δ Rimanenze" vals={R.RF.map((r) => r.varRim)} />
        <Row label="(-) Δ Crediti commerciali" vals={R.RF.map((r) => r.varCre)} />
        <Row label="(+) Δ Debiti commerciali" vals={R.RF.map((r) => r.varDebC)} />
        <Row label="(+) Δ Debito IVA" vals={R.RF.map((r) => r.varDebIVA)} />
        <Row label="(-) Pagamento oneri finanziari" vals={R.RF.map((r) => r.pagOneri)} />
        <Row label="(+) Δ Debiti tributari" vals={R.RF.map((r) => r.varDebTrib)} />
        <Row label="FLUSSO GESTIONE OPERATIVA" vals={R.RF.map((r) => r.flussoOp)} cls="bold" />
        <tr className="section"><td colSpan={Y + 1}>B) Gestione Investimenti</td></tr>
        <Row label="(-) Investimenti mat. e immat." vals={R.RF.map((r) => r.flussoInv)} />
        <Row label="FLUSSO INVESTIMENTI" vals={R.RF.map((r) => r.flussoInv)} cls="subtotal" />
        <tr className="section"><td colSpan={Y + 1}>C) Gestione Finanziaria</td></tr>
        <Row label="Apporti equity / debito - rimborsi" vals={R.RF.map((r) => r.flussoFin)} />
        <Row label="FLUSSO FINANZIARIO" vals={R.RF.map((r) => r.flussoFin)} cls="subtotal" />
        <Row label="VARIAZIONE NETTA DI CASSA" vals={R.RF.map((r) => r.varCassa)} cls="bold" />
        <Row label="Cassa iniziale" vals={R.RF.map((r) => r.cassaIn)} />
        <Row label="CASSA FINALE" vals={R.RF.map((r) => r.cassaFin)} cls="bold" />
      </tbody>
    </table>
  );
}

export function TableID({ R }) {
  const irrTxt = isFinite(R.IRR) ? (R.IRR * 100).toFixed(1) + '%' : 'n.c.';
  const pbTxt = R.payback ? `Anno ${R.payback}` : 'oltre Y5';
  return (
    <table>
      <Header labels={yLabels} />
      <tbody>
        <tr className="section"><td colSpan={Y + 1}>Redditività</td></tr>
        <Row label="EBITDA margin"      vals={R.CE.map((c) => c.ebitdaMar)} fmtType="pct" />
        <Row label="EBIT margin"        vals={R.CE.map((c) => c.ebitMar)}   fmtType="pct" />
        <Row label="ROS (Utile/Ricavi)" vals={R.ID.map((d) => d.ROS)} fmtType="pct" />
        <Row label="ROI (EBIT/CapInvest.)" vals={R.ID.map((d) => d.ROI)} fmtType="pct" />
        <Row label="ROE (Utile/PN)"     vals={R.ID.map((d) => d.ROE)} fmtType="pct" />
        <tr className="section"><td colSpan={Y + 1}>Solidità</td></tr>
        <Row label="PFN (Debito - Cassa)" vals={R.ID.map((d) => d.PFN)} />
        <tr>
          <td>PFN / EBITDA</td>
          {R.ID.map((d, i) => (
            <td key={i}>{fmt(d.pfnEbitda, 2)}x {flagPill(d.pfnEbitda, 'pfnEbitda')}</td>
          ))}
        </tr>
        <Row label="Leva finanziaria"        vals={R.ID.map((d) => d.leva)}  fmtType="num" />
        <Row label="Indipendenza finanziaria" vals={R.ID.map((d) => d.indip)} fmtType="pct" />
        <tr className="section"><td colSpan={Y + 1}>Liquidità</td></tr>
        <Row label="Current ratio" vals={R.ID.map((d) => d.currR)}  fmtType="num" />
        <Row label="Quick ratio"   vals={R.ID.map((d) => d.quickR)} fmtType="num" />
        <Row label="Ciclo del circolante (giorni)" vals={R.ID.map((d) => ({ txt: fmt(d.ciclo, 0) + ' gg' }))} />
        <tr className="section"><td colSpan={Y + 1}>Sostenibilità</td></tr>
        <tr>
          <td>DSCR</td>
          {R.ID.map((d, i) => (
            <td key={i}>{fmt(d.DSCR, 2)}x {flagPill(d.DSCR, 'dscr')}</td>
          ))}
        </tr>
        <Row label="Copertura Oneri Finanz. (EBITDA/OF)" vals={R.ID.map((d) => d.copOF)} fmtType="num" />
        <tr className="section"><td colSpan={Y + 1}>Break-Even</td></tr>
        <Row label="Punto di pareggio (€)" vals={R.ID.map((d) => d.BEP)} />
        <Row label="Margine di sicurezza"  vals={R.ID.map((d) => d.margSic)} fmtType="pct" />
        <tr className="section"><td colSpan={Y + 1}>Valutazione del Progetto</td></tr>
        <Row label="Free Cash Flow" vals={R.FCF} />
        <Row label="FCF cumulato"   vals={R.FCFcum} />
        <tr className="bold">
          <td>VAN @ {(R.inp.wacc * 100).toFixed(1)}%</td>
          <td colSpan={Y} style={{ textAlign: 'left' }}>
            {fmtE(R.NPV)} &nbsp;·&nbsp; <strong>TIR:</strong> {irrTxt} &nbsp;·&nbsp; <strong>Payback:</strong> {pbTxt}
          </td>
        </tr>
      </tbody>
    </table>
  );
}
