import { Bar, Line } from 'react-chartjs-2';
import {
  Chart, BarController, BarElement, LineController, LineElement, PointElement,
  CategoryScale, LinearScale, Tooltip, Legend, Title, Filler
} from 'chart.js';

Chart.register(
  BarController, BarElement, LineController, LineElement, PointElement,
  CategoryScale, LinearScale, Tooltip, Legend, Title, Filler
);
Chart.defaults.font.family = "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto";
Chart.defaults.font.size = 11;

const yLabels = ['Anno 1', 'Anno 2', 'Anno 3', 'Anno 4', 'Anno 5'];
const eurTick = { callback: (v) => '€' + (v / 1000).toFixed(0) + 'k' };

export default function ChartGrid({ R }) {
  const revData = {
    labels: yLabels,
    datasets: [
      { type: 'bar',  label: 'Ricavi',      data: R.CE.map((c) => c.ricavi),  backgroundColor: '#0b3d91' },
      { type: 'bar',  label: 'EBITDA',      data: R.CE.map((c) => c.ebitda),  backgroundColor: '#10b981' },
      { type: 'line', label: 'Utile Netto', data: R.CE.map((c) => c.utile),   borderColor: '#f59e0b', backgroundColor: '#f59e0b', tension: 0.25 }
    ]
  };

  const fcfData = {
    labels: yLabels,
    datasets: [
      { type: 'bar',  label: 'FCF',          data: R.FCF,    backgroundColor: R.FCF.map((v) => (v >= 0 ? '#10b981' : '#ef4444')) },
      { type: 'line', label: 'FCF cumulato', data: R.FCFcum, borderColor: '#0b3d91', backgroundColor: '#0b3d91', tension: 0.25, fill: false }
    ]
  };

  const spData = {
    labels: yLabels,
    datasets: [
      { type: 'bar',  label: 'Patrimonio Netto',          data: R.SP.map((s) => s.PN),                       backgroundColor: '#0b3d91', stack: 'a' },
      { type: 'bar',  label: 'Debiti finanziari',         data: R.SP.map((s) => s.debFin),                   backgroundColor: '#ef4444', stack: 'a' },
      { type: 'bar',  label: 'Debiti commerciali+IVA+trib.', data: R.SP.map((s) => s.debComm + s.debIVA + s.debTrib), backgroundColor: '#f59e0b', stack: 'a' },
      { type: 'line', label: 'Cassa',                      data: R.SP.map((s) => s.cassa),                    borderColor: '#10b981', backgroundColor: '#10b981', tension: 0.25 }
    ]
  };

  const indData = {
    labels: yLabels,
    datasets: [
      { label: 'DSCR',       data: R.ID.map((d) => d.DSCR),      borderColor: '#10b981', backgroundColor: '#10b981', tension: 0.25, yAxisID: 'y'  },
      { label: 'PFN/EBITDA', data: R.ID.map((d) => d.pfnEbitda), borderColor: '#ef4444', backgroundColor: '#ef4444', tension: 0.25, yAxisID: 'y1' }
    ]
  };

  const baseOpts = (extra = {}) => ({
    responsive: true,
    maintainAspectRatio: false,
    plugins: { legend: { position: 'bottom' } },
    ...extra
  });

  return (
    <div className="chart-grid">
      <div className="chart-box">
        <h4>Ricavi vs EBITDA vs Utile Netto</h4>
        <div style={{ height: 240 }}>
          <Bar data={revData} options={baseOpts({ scales: { y: { ticks: eurTick } } })} />
        </div>
      </div>
      <div className="chart-box">
        <h4>Free Cash Flow cumulato &amp; Payback</h4>
        <div style={{ height: 240 }}>
          <Bar data={fcfData} options={baseOpts({ scales: { y: { ticks: eurTick } } })} />
        </div>
      </div>
      <div className="chart-box">
        <h4>Struttura Stato Patrimoniale</h4>
        <div style={{ height: 240 }}>
          <Bar data={spData} options={baseOpts({ scales: { x: { stacked: true }, y: { stacked: true, ticks: eurTick } } })} />
        </div>
      </div>
      <div className="chart-box">
        <h4>Indici di Sostenibilità (DSCR, PFN/EBITDA)</h4>
        <div style={{ height: 240 }}>
          <Line data={indData} options={baseOpts({
            scales: {
              y:  { position: 'left',  title: { display: true, text: 'DSCR (x)' } },
              y1: { position: 'right', title: { display: true, text: 'PFN/EBITDA (x)' }, grid: { drawOnChartArea: false } }
            }
          })} />
        </div>
      </div>
    </div>
  );
}
