import { Y } from '../lib/defaults.js';
import { fmt, fmtE } from '../lib/format.js';
import NumberTicker from './NumberTicker.jsx';

function pfnPill(v) {
  if (v <= 3) return <span className="pill pill-ok">Sostenibile</span>;
  if (v <= 5) return <span className="pill pill-warn">Tirato</span>;
  return <span className="pill pill-bad">Eccessivo</span>;
}

const fmtEur = (n) => (isFinite(n) ? fmtE(Math.round(n)) : '—');
const fmtPct = (n) => (isFinite(n) ? (n * 100).toFixed(1) + '%' : '—');
const fmtNumX = (n) => (isFinite(n) ? fmt(n, 2) + 'x' : '—');

export default function KPIRow({ R }) {
  const ricY1 = R.CE[0].ricavi;
  const ricY5 = R.CE[Y - 1].ricavi;
  const cagr = ricY1 > 0 && ricY5 > 0 ? Math.pow(ricY5 / ricY1, 1 / (Y - 1)) - 1 : 0;
  const ebiCum = R.CE.reduce((a, c) => a + c.ebitda, 0);
  const ebiMarMed = R.CE.reduce((a, c) => a + c.ebitdaMar, 0) / Y;
  const irr = isFinite(R.IRR) ? R.IRR : null;
  const pfn5 = R.ID[Y - 1].pfnEbitda;

  return (
    <div className="kpi-row">
      <div className="kpi blue fade-up">
        <div className="lbl">Ricavi 5° anno</div>
        <div className="v"><NumberTicker value={ricY5} format={fmtEur} /></div>
        <div className="sub">CAGR <NumberTicker value={cagr * 100} format={(n) => (isFinite(n) ? n.toFixed(1) + '%' : '—')} /></div>
      </div>
      <div className="kpi green fade-up-delay-1">
        <div className="lbl">EBITDA cumulato</div>
        <div className="v"><NumberTicker value={ebiCum} format={fmtEur} /></div>
        <div className="sub">Margine medio <NumberTicker value={ebiMarMed * 100} format={(n) => (isFinite(n) ? n.toFixed(1) + '%' : '—')} /></div>
      </div>
      <div className="kpi amber fade-up-delay-2">
        <div className="lbl">VAN @ WACC</div>
        <div className="v"><NumberTicker value={R.NPV} format={fmtEur} /></div>
        <div className="sub">TIR {irr === null ? 'n.c.' : <NumberTicker value={irr * 100} format={(n) => n.toFixed(1) + '%'} />}</div>
      </div>
      <div className="kpi red fade-up-delay-3">
        <div className="lbl">PFN/EBITDA Y5</div>
        <div className="v"><NumberTicker value={pfn5} format={fmtNumX} /></div>
        <div className="sub">{pfnPill(pfn5)}</div>
      </div>
    </div>
  );
}
