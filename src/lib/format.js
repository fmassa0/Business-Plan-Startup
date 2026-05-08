const nf0 = new Intl.NumberFormat('it-IT', { minimumFractionDigits: 0, maximumFractionDigits: 0 });
const nf2 = new Intl.NumberFormat('it-IT', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

export const fmt = (n, d = 0) => {
  if (!isFinite(n)) return '—';
  return d === 2 ? nf2.format(n) : nf0.format(n);
};

export const fmtE = (n, d = 0) => '€ ' + fmt(n, d);

export const fmtP = (n) => (isFinite(n) ? (n * 100).toFixed(1) + '%' : '—');
