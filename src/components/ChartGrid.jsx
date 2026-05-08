import { Bar, Line } from 'react-chartjs-2';
import {
  Chart, BarController, BarElement, LineController, LineElement, PointElement,
  CategoryScale, LinearScale, Tooltip, Legend, Title, Filler
} from 'chart.js';

Chart.register(
  BarController, BarElement, LineController, LineElement, PointElement,
  CategoryScale, LinearScale, Tooltip, Legend, Title, Filler
);
Chart.defaults.font.family = "'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto";
Chart.defaults.font.size = 11;
Chart.defaults.color = '#64748B';
Chart.defaults.borderColor = '#EEF1F5';
Chart.defaults.plugins.legend.labels.boxWidth = 10;
Chart.defaults.plugins.legend.labels.boxHeight = 10;
Chart.defaults.plugins.legend.labels.padding = 16;
Chart.defaults.plugins.tooltip.backgroundColor = '#0A1F44';
Chart.defaults.plugins.tooltip.titleFont = { weight: '600', size: 12 };
Chart.defaults.plugins.tooltip.bodyFont = { size: 12 };
Chart.defaults.plugins.tooltip.padding = 10;
Chart.defaults.plugins.tooltip.cornerRadius = 8;

const PALETTE = {
  navy:   '#0A1F44',
  navy2:  '#1E3A8A',
  teal:   '#00B0A0',
  teal2:  '#00C9B7',
  amber:  '#F59E0B',
  red:    '#EF4444',
  purple: '#7C5BC9'
};

const yLabels = ['Anno 1', 'Anno 2', 'Anno 3', 'Anno 4', 'Anno 5'];
const eurTick = { callback: (v) => '€' + (v / 1000).toFixed(0) + 'k' };

export default function ChartGrid({ R }) {
  const revData = {
    labels: yLabels,
    datasets: [
      { type: 'bar',  label: 'Ricavi',      data: R.CE.map((c) => c.ricavi),  backgroundColor: PALETTE.navy,  borderRadius: 4 },
      { type: 'bar',  label: 'EBITDA',      data: R.CE.map((c) => c.ebitda),  backgroundColor: PALETTE.teal,  borderRadius: 4 },
      { type: 'line', label: 'Utile Netto', data: R.CE.map((c) => c.utile),   borderColor: PALETTE.amber, backgroundColor: PALETTE.amber, tension: 0.3, pointRadius: 3, pointHoverRadius: 5, borderWidth: 2 }
    ]
  };

  const fcfData = {
    labels: yLabels,
    datasets: [
      { type: 'bar',  label: 'FCF',          data: R.FCF,    backgroundColor: R.FCF.map((v) => (v >= 0 ? PALETTE.teal : PALETTE.red)), borderRadius: 4 },
      { type: 'line', label: 'FCF cumulato', data: R.FCFcum, borderColor: PALETTE.navy, backgroundColor: PALETTE.navy, tension: 0.3, fill: false, pointRadius: 3, pointHoverRadius: 5, borderWidth: 2 }
    ]
  };

  const spData = {
    labels: yLabels,
    datasets: [
      { type: 'bar',  label: 'Patrimonio Netto',          data: R.SP.map((s) => s.PN),                       backgroundColor: PALETTE.navy,   stack: 'a', borderRadius: 4 },
      { type: 'bar',  label: 'Debiti finanziari',         data: R.SP.map((s) => s.debFin),                   backgroundColor: PALETTE.red,    stack: 'a', borderRadius: 4 },
      { type: 'bar',  label: 'Debiti commerciali+IVA+trib.', data: R.SP.map((s) => s.debComm + s.debIVA + s.debTrib), backgroundColor: PALETTE.amber,  stack: 'a', borderRadius: 4 },
      { type: 'line', label: 'Cassa',                      data: R.SP.map((s) => s.cassa),                    borderColor: PALETTE.teal, backgroundColor: PALETTE.teal, tension: 0.3, pointRadius: 3, borderWidth: 2 }
    ]
  };

  const indData = {
    labels: yLabels,
    datasets: [
      { label: 'DSCR',       data: R.ID.map((d) => d.DSCR),      borderColor: PALETTE.teal, backgroundColor: PALETTE.teal, tension: 0.3, yAxisID: 'y',  pointRadius: 3, borderWidth: 2 },
      { label: 'PFN/EBITDA', data: R.ID.map((d) => d.pfnEbitda), borderColor: PALETTE.red,  backgroundColor: PALETTE.red,  tension: 0.3, yAxisID: 'y1', pointRadius: 3, borderWidth: 2 }
    ]
  };

  const baseOpts = (extra = {}) => ({
    responsive: true,
    maintainAspectRatio: false,
    interaction: { mode: 'index', intersect: false },
    plugins: { legend: { position: 'bottom', labels: { usePointStyle: true } } },
    animation: { duration: 450, easing: 'easeOutCubic' },
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
