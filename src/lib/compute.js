import { Y } from './defaults.js';

const num = (raw) => {
  const n = parseFloat(raw);
  return isFinite(n) ? n : 0;
};

export function normalizeInputs(state) {
  return {
    vol: num(state.i_vol), prc: num(state.i_prc), cuv: num(state.i_cuv),
    g_vol: num(state.g_vol) / 100, g_prc: num(state.g_prc) / 100, g_cuv: num(state.g_cuv) / 100,
    pers: num(state.i_pers), aff: num(state.i_aff), mkt: num(state.i_mkt), serv: num(state.i_serv), alt: num(state.i_alt),
    g_cf: num(state.g_cf) / 100,
    inv_m: [num(state.im1), num(state.im2), num(state.im3), num(state.im4), num(state.im5)],
    inv_i: [num(state.ii1), num(state.ii2), num(state.ii3), num(state.ii4), num(state.ii5)],
    aliq_m: num(state.aliq_m) / 100, aliq_i: num(state.aliq_i) / 100,
    dso: num(state.dso), dpo: num(state.dpo), dio: num(state.dio),
    iva_v: num(state.iva_v) / 100, iva_a: num(state.iva_a) / 100, iva_p: num(state.iva_p),
    eq: [num(state.eq1), num(state.eq2), num(state.eq3), num(state.eq4), num(state.eq5)],
    db: [num(state.db1), num(state.db2), num(state.db3), num(state.db4), num(state.db5)],
    rb: [num(state.rb1), num(state.rb2), num(state.rb3), num(state.rb4), num(state.rb5)],
    tasso: num(state.tasso) / 100,
    ires: num(state.ires) / 100, irap: num(state.irap) / 100,
    wacc: num(state.wacc) / 100,
    s_prc: num(state.s_prc) / 100,
    s_vol: num(state.s_vol) / 100,
    s_cuv: num(state.s_cuv) / 100,
    s_cf:  num(state.s_cf)  / 100
  };
}

export function compute(state) {
  const inp = normalizeInputs(state);

  const anni = [];
  for (let y = 0; y < Y; y++) {
    const vol = inp.vol * Math.pow(1 + inp.g_vol, y) * (1 + inp.s_vol);
    const prc = inp.prc * Math.pow(1 + inp.g_prc, y) * (1 + inp.s_prc);
    const cuv = inp.cuv * Math.pow(1 + inp.g_cuv, y) * (1 + inp.s_cuv);
    const cf_mult = Math.pow(1 + inp.g_cf, y) * (1 + inp.s_cf);
    const pers = inp.pers * cf_mult, aff = inp.aff * cf_mult, mkt = inp.mkt * cf_mult, serv = inp.serv * cf_mult, alt = inp.alt * cf_mult;
    const ricavi = vol * prc;
    const cVar = vol * cuv;
    const totCF = pers + aff + mkt + serv + alt;
    anni.push({ vol, prc, cuv, ricavi, cVar, pers, aff, mkt, serv, alt, totCF });
  }

  const ammMat = [], ammImm = [];
  for (let y = 0; y < Y; y++) {
    let am = 0, ai = 0;
    for (let k = 0; k <= y; k++) {
      am += inp.inv_m[k] * inp.aliq_m;
      ai += inp.inv_i[k] * inp.aliq_i;
    }
    ammMat.push(am); ammImm.push(ai);
  }
  const aliq_tot = inp.ires + inp.irap;

  const CE = [], SP = [], RF = [];
  let openDebt = 0, openCash = 0;
  let openImmMat = 0, openImmImmat = 0, openFAmmMat = 0, openFAmmImm = 0;
  let openCap = 0, openUtNuovo = 0;
  let openRimanenze = 0, openCrediti = 0, openDebComm = 0, openDebIVA = 0, openDebTrib = 0;

  for (let y = 0; y < Y; y++) {
    const a = anni[y];

    const mdc = a.ricavi - a.cVar;
    const mdcPct = a.ricavi ? mdc / a.ricavi : 0;
    const ebitda = mdc - a.totCF;
    const ebitdaMar = a.ricavi ? ebitda / a.ricavi : 0;
    const ammM = ammMat[y], ammI = ammImm[y];
    const ebit = ebitda - ammM - ammI;
    const ebitMar = a.ricavi ? ebit / a.ricavi : 0;
    const oneriFin = -openDebt * inp.tasso;
    const rai = ebit + oneriFin;
    const imposte = -Math.max(rai, 0) * aliq_tot;
    const utile = rai + imposte;

    CE.push({
      ricavi: a.ricavi, cVar: -a.cVar, mdc, mdcPct,
      pers: -a.pers, aff: -a.aff, mkt: -a.mkt, serv: -a.serv, alt: -a.alt,
      totCF: -a.totCF, ebitda, ebitdaMar, ammMat: -ammM, ammImm: -ammI,
      ebit, ebitMar, oneriFin, rai, imposte, utile
    });

    const immMatLordo = openImmMat + inp.inv_m[y];
    const fAmmMat = openFAmmMat + ammM;
    const immMatNetto = immMatLordo - fAmmMat;
    const immImmatLordo = openImmImmat + inp.inv_i[y];
    const fAmmImm = openFAmmImm + ammI;
    const immImmatNetto = immImmatLordo - fAmmImm;
    const totImm = immMatNetto + immImmatNetto;

    const rimanenze = a.cVar * inp.dio / 365;
    const crediti = a.ricavi * (1 + inp.iva_v) * inp.dso / 365;
    const credIVA = 0;

    const debFin = openDebt + inp.db[y] - inp.rb[y];
    const debComm = a.cVar * (1 + inp.iva_a) * inp.dpo / 365;
    const debIVA = (a.ricavi * inp.iva_v - a.cVar * inp.iva_a) * inp.iva_p / 12;
    const debTrib = -imposte;
    const altriDeb = 0;

    const varRim = -(rimanenze - openRimanenze);
    const varCre = -(crediti - openCrediti);
    const varDebC = (debComm - openDebComm);
    const varDebIVA = (debIVA - openDebIVA);
    const varDebTrib = (debTrib - openDebTrib);
    const pagOneri = oneriFin;
    const flussoPreCCN = utile + (ammM + ammI) + (-oneriFin);
    const flussoOp = flussoPreCCN + varRim + varCre + varDebC + varDebIVA + pagOneri + varDebTrib;
    const flussoInv = -(inp.inv_m[y] + inp.inv_i[y]);
    const flussoFin = inp.eq[y] + inp.db[y] - inp.rb[y];
    const varCassa = flussoOp + flussoInv + flussoFin;
    const cassa = openCash + varCassa;

    const cap = openCap + inp.eq[y];
    const utNuovo = openUtNuovo;
    const PN = cap + utNuovo + utile;

    const totAttivoCirc = rimanenze + crediti + credIVA + cassa;
    const totAttivo = totImm + totAttivoCirc;
    const totPassCorr = debComm + debIVA + debTrib + altriDeb;
    const totPassivo = debFin + debComm + debIVA + debTrib + altriDeb;
    const totPP = PN + totPassivo;
    const quad = totAttivo - totPP;

    SP.push({
      immMatLordo, fAmmMat, immMatNetto, immImmatLordo, fAmmImm, immImmatNetto, totImm,
      rimanenze, crediti, credIVA, cassa, totAttivoCirc, totAttivo,
      cap, utNuovo, utile, PN, debFin, debComm, debIVA, debTrib, altriDeb,
      totPassivo, totPassCorr, totPP, quad
    });

    RF.push({
      utile, amm: ammM + ammI, oneriSter: -oneriFin, preCCN: flussoPreCCN,
      varRim, varCre, varDebC, varDebIVA, pagOneri, varDebTrib,
      flussoOp, flussoInv, flussoFin, varCassa, cassaIn: openCash, cassaFin: cassa
    });

    openDebt = debFin; openCash = cassa;
    openImmMat = immMatLordo; openImmImmat = immImmatLordo;
    openFAmmMat = fAmmMat; openFAmmImm = fAmmImm;
    openCap = cap;
    openUtNuovo = utNuovo + utile;
    openRimanenze = rimanenze; openCrediti = crediti;
    openDebComm = debComm; openDebIVA = debIVA; openDebTrib = debTrib;
  }

  const ID = [];
  for (let y = 0; y < Y; y++) {
    const ce = CE[y], sp = SP[y];
    const ebitda = ce.ebitda, utile = ce.utile, PN = sp.PN;
    const capInv = sp.totImm + sp.rimanenze + sp.crediti - sp.debComm;
    const ROI = capInv ? ce.ebit / capInv : 0;
    const ROE = PN ? utile / PN : 0;
    const ROS = ce.ricavi ? utile / ce.ricavi : 0;
    const PFN = sp.debFin - sp.cassa;
    const pfnEbitda = ebitda ? PFN / ebitda : 0;
    const leva = PN ? sp.debFin / PN : 0;
    const indip = sp.totAttivo ? PN / sp.totAttivo : 0;
    const currR = sp.totPassCorr ? sp.totAttivoCirc / sp.totPassCorr : 0;
    const quickR = sp.totPassCorr ? (sp.totAttivoCirc - sp.rimanenze) / sp.totPassCorr : 0;
    const ciclo = inp.dso + inp.dio - inp.dpo;
    const dscrSrv = inp.rb[y] + (-ce.oneriFin);
    const DSCR = dscrSrv ? RF[y].flussoOp / dscrSrv : 0;
    const copOF = ce.oneriFin ? ebitda / (-ce.oneriFin) : 0;
    const BEP = ce.mdcPct ? (-ce.totCF) / ce.mdcPct : 0;
    const margSic = ce.ricavi ? (ce.ricavi - BEP) / ce.ricavi : 0;
    ID.push({ ROI, ROE, ROS, PFN, pfnEbitda, leva, indip, currR, quickR, ciclo, DSCR, copOF, BEP, margSic });
  }

  const FCF = CE.map((_, y) => RF[y].flussoOp + RF[y].flussoInv);
  let NPV = 0;
  for (let y = 0; y < Y; y++) NPV += FCF[y] / Math.pow(1 + inp.wacc, y + 1);

  let IRR = NaN;
  const npvAt = (r) => {
    let s = 0;
    for (let y = 0; y < Y; y++) s += FCF[y] / Math.pow(1 + r, y + 1);
    return s;
  };
  const hasNeg = FCF.some((v) => v < 0), hasPos = FCF.some((v) => v > 0);
  if (hasNeg && hasPos) {
    let lo = -0.99, hi = 5, fLo = npvAt(lo), fHi = npvAt(hi);
    if (fLo * fHi < 0) {
      for (let i = 0; i < 80; i++) {
        const m = (lo + hi) / 2, fm = npvAt(m);
        if (fm === 0) { IRR = m; break; }
        if (fLo * fm < 0) { hi = m; fHi = fm; } else { lo = m; fLo = fm; }
      }
      IRR = (lo + hi) / 2;
    }
  }

  const FCFcum = [];
  let cum = 0;
  FCF.forEach((v) => { cum += v; FCFcum.push(cum); });
  let payback = null;
  for (let i = 0; i < FCFcum.length; i++) if (FCFcum[i] >= 0) { payback = i + 1; break; }

  return { anni, CE, SP, RF, ID, FCF, FCFcum, NPV, IRR, payback, inp };
}
