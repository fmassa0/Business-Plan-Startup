export const Y = 5;

export const DEFAULTS = {
  i_vol: 1000, i_prc: 100, i_cuv: 40,
  g_vol: 20, g_prc: 2, g_cuv: 1,
  i_pers: 40000, i_aff: 12000, i_mkt: 10000, i_serv: 6000, i_alt: 3000, g_cf: 3,
  im1: 50000, im2: 0, im3: 10000, im4: 0, im5: 0, aliq_m: 20,
  ii1: 15000, ii2: 0, ii3: 0, ii4: 0, ii5: 0, aliq_i: 33,
  dso: 60, dpo: 60, dio: 30,
  iva_v: 22, iva_a: 22, iva_p: 3,
  eq1: 50000, eq2: 0, eq3: 0, eq4: 0, eq5: 0,
  db1: 40000, db2: 0, db3: 0, db4: 0, db5: 0,
  rb1: 8000, rb2: 8000, rb3: 8000, rb4: 8000, rb5: 8000,
  tasso: 6, ires: 24, irap: 3.9, wacc: 8,
  s_prc: 0, s_vol: 0, s_cuv: 0, s_cf: 0
};

export const ZEROES = Object.fromEntries(Object.keys(DEFAULTS).map((k) => [k, 0]));

export const PRESETS = {
  base:  { s_prc:   0, s_vol:   0, s_cuv:  0, s_cf:  0 },
  best:  { s_prc:  15, s_vol:  20, s_cuv: -5, s_cf:  0 },
  worst: { s_prc: -15, s_vol: -20, s_cuv:  5, s_cf:  5 },
  cost:  { s_prc:   0, s_vol:   0, s_cuv: 15, s_cf: 15 }
};
