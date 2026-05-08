import { Y } from '../lib/defaults.js';
import { fmt, fmtE } from '../lib/format.js';

function pfnPill(v) {
  if (v <= 3) return <span className="pill pill-ok">Sostenibile</span>;
  if (v <= 5) return <span className="pill pill-warn">Tirato</span>;
  return <span className="pill pill-bad">Eccessivo</span>;
}

export default function KPIRow({ R }) {
  const ricY1 = R.CE[0].ricavi;
  const ricY5 = R.CE[Y - 1].ricavi;
  const cagr = ricY1 > 0 && ricY5 > 0 ? Math.pow(ricY5 / ricY1, 1 / (Y - 1)) - 1 : 0;
  const ebiCum = R.CE.reduce((a, c) => a + c.ebitda, 0);
  const ebiMarMed = R.CE.reduce((a, c) => a + c.ebitdaMar, 0) / Y;
  const irrTxt = isFinite(R.IRR) ? (R.IRR * 100).toFixed(1) + '%' : 'n.c.';
  const pfn5 = R.ID[Y - 1].pfnEbitda;

  return (
    <div className="kpi-row">
      <div className="kpi blue">
        <div className="lbl">Ricavi 5° anno</div>
        <div className="v">{fmtE(ricY5)}</div>
        <div className="sub">CAGR {(cagr * 100).toFixed(1)}%</div>
      </div>
      <div className="kpi green">
        <div className="lbl">EBITDA cumulato</div>
        <div className="v">{fmtE(ebiCum)}</div>
        <div className="sub">Margine medio {(ebiMarMed * 100).toFixed(1)}%</div>
      </div>
      <div className="kpi amber">
        <div className="lbl">VAN @ WACC</div>
        <div className="v">{fmtE(R.NPV)}</div>
        <div className="sub">TIR {irrTxt}</div>
      </div>
      <div className="kpi red">
        <div className="lbl">PFN/EBITDA Y5</div>
        <div className="v">{fmt(pfn5, 2)}x</div>
        <div className="sub">{pfnPill(pfn5)}</div>
      </div>
    </div>
  );
}
