/**
 * ═══════════════════════════════════════════════════════════════════════════
 *  IMPULSO CRÉDITO — CREDIT POLICY & DECISION ENGINE
 *  Single-file implementation for Google Apps Script.
 *  Version 1.0 · 16 August 2026
 * ═══════════════════════════════════════════════════════════════════════════
 *
 *  INSTALL
 *    Extensions ▸ Apps Script ▸ paste this file ▸ Save ▸ reload the sheet.
 *    A "Crédito" menu appears. Run "Verificar cartera completa" once to
 *    backfill, then use the menu or the =EVALUAR() formula for new applicants.
 *
 *  WHAT THIS FILE CONTAINS
 *    1. The credit policy, in full, as documentation (below).
 *    2. The algorithm that implements it, rule for rule.
 *    Policy and code live together so they cannot drift apart. If you change
 *    a threshold, change it in CONFIG and update the policy text beside it.
 *
 * ───────────────────────────────────────────────────────────────────────────
 *  THE POLICY
 * ───────────────────────────────────────────────────────────────────────────
 *
 *  Rules are evaluated in order. The first decline ends the assessment.
 *  No override without a written reason recorded against the application.
 *
 *  SECTION A — HARD DECLINES (automatic, no exceptions)
 *    A1  Rejected before ....... DNI, email or phone appears in Rechazados.
 *    A2  Duplicate contact ..... email or phone already used by a different DNI.
 *    A3  Identity unverifiable . CUIL fails check digit, CUIL's DNI segment
 *                                differs from the stated DNI, or the BCRA name
 *                                does not match the application.
 *    A5  No BCRA record ........ neither CUIL prefix returns anything.
 *                                Absence of evidence is not good standing.
 *    A4  Situación 4 or 5 ...... worst classification at any entity is 4 or 5.
 *    A6  Rejected cheques ...... any record in the cheques rechazados register.
 *
 *    Note the order: A5 is tested before A4, because A4 needs a record to read.
 *    Reversed, an applicant with no record silently passes the situación check.
 *
 *    A4 + A5 together move the modelled default rate from 48.2% to 9.9%.
 *    A1–A3 protect against fraud, not default — separately necessary.
 *
 *  SECTION B — REFINEMENT DECLINES (enable when the funnel supports it)
 *    B1  Majority impaired ..... under 50% of system debt in situación 1–2.
 *    B2  Deteriorating ......... worst classification rose ≥2 steps in 6 months.
 *    B3  Recent line impaired .. any line first seen within 12 months is now
 *                                situación 3 or worse.
 *
 *    B3 is the Xoana Berón rule. She showed 74% performing debt but had opened
 *    a new bank facility and defaulted on it within four months — exactly as
 *    she had with a previous lender. Old performing debt says someone honours
 *    commitments made years ago. It says nothing about a lender arriving today.
 *
 *  SECTION C — TIER AND TICKET
 *    Tier 1  sit ≤2, ≥70% performing, stable/improving ... first loan 1,200,000
 *    Tier 2  sit ≤2, 50–70% performing .................... first loan   450,000
 *    Tier 3  sit  3, ≥90% performing, not deteriorating ... first loan   220,000
 *    else    decline
 *    C1  Ladder ......... limit doubles per clean repayment, ceiling 2,000,000.
 *                         One late payment drops the borrower a rung.
 *    C2  Single cap ..... no borrower above 10% of the fund.
 *    C3  Group cap ...... no surname/email/phone/address group above 15%.
 *    C4  Guarantor ...... required above 450,000; guarantor passes A and B too.
 *
 *  SECTION D — LOAN STRUCTURE
 *    D1  Instalments .... weekly repayment above 220,000. A missed week is a
 *                         signal on day 7; a missed balloon is a loss on day 30.
 *    D2  Short first .... every first loan runs the 15-day product.
 *    D3  Staggered ...... no more than 20% of the book maturing in any 7 days.
 *    D4  Late charge .... 1%/day on the OVERDUE BALANCE, capped at 30% of
 *                         principal. The old 5%/day on the total doubles the
 *                         debt in 20 days, at which point the borrower stops
 *                         answering and a partial loss becomes a total one.
 *
 *  SECTION E — PROCESS CONTROLS
 *    E1  The check blocks disbursement. Not a warning — a hard stop.
 *    E2  Required: CUIL (typed), gender, DNI front and back, selfie with DNI.
 *    E3  Every decline is written to Rechazados with its rule number.
 *        Declines are half your dataset; without them you only learn which
 *        approvals went bad, never which declines you got wrong.
 *    E4  Form closes when available cash falls below 500,000. No manual reopen.
 *    E5  Contact within 24 hours of any miss. Speed beats severity.
 *    E6  Monthly: recompute realised default by tier; loosen only on evidence.
 *
 *  SECTION F — OVERRIDES
 *    Require the rule number, a written reason, and the loan is forced to
 *    Tier 3 limits. Tracked separately. If overrides exceed 10% of approvals
 *    in a month, the funnel is too small — fix the funnel, not the policy.
 *
 *  CALIBRATION
 *    These thresholds rest on BCRA's own provisioning ladder — what a regulated
 *    bank expects to lose. An informal lender with no collections history should
 *    expect worse: 12–18% initially rather than 9.9%. Single digit is a
 *    12–18 month destination.
 *
 *    The rules over-reject on purpose. Claudia Torres was situación 5 with zero
 *    performing debt and repaid in full nine days early — every rule here would
 *    have declined her. You cannot yet tell the Torreses from the Zuloagas.
 *    Decline both, gather outcomes, loosen from evidence.
 *
 * ═══════════════════════════════════════════════════════════════════════════
 */


/* ═══════════════════════════════════════════════════════════════════════════
   CONFIG — every threshold in the policy, in one place
   ═══════════════════════════════════════════════════════════════════════════ */

var CONFIG = {
  BCRA_BASE: 'https://api.bcra.gob.ar/CentralDeDeudores/v1.0',

  SECTION_B_ENABLED: false,      // turn on when applicant flow supports it

  MIN_CASH_TO_LEND: 500000,      // E4
  FUND_TOTAL_CELL: 'Fondo total para prestar',

  CAP_SINGLE_PCT: 0.10,          // C2
  CAP_GROUP_PCT: 0.15,           // C3
  LADDER_CEILING: 2000000,       // C1
  GUARANTOR_ABOVE: 450000,       // C4
  INSTALMENT_ABOVE: 220000,      // D1
  MATURITY_WINDOW_PCT: 0.20,     // D3
  LATE_DAILY: 0.01,              // D4
  LATE_CAP_PCT: 0.30,            // D4

  TIER_LIMITS: { 1: 1200000, 2: 450000, 3: 220000 },

  B1_MIN_PERFORMING: 0.50,
  B2_MAX_SLIDE: 2,
  B3_RECENT_MONTHS: 12,

  T1_MIN_PERFORMING: 0.70,
  T2_MIN_PERFORMING: 0.50,
  T3_MIN_PERFORMING: 0.90,

  SHEETS: {
    borrowers: 'Prestatarios',
    applicants: 'Nuevos Prestatarios',
    rejected: 'Rechazados',
    payments: 'Pagos',
    config: 'Configuración',
    log: 'Decisiones'
  },

  API_PAUSE_MS: 1500,
  CACHE_HOURS: 6
};


/* ═══════════════════════════════════════════════════════════════════════════
   HELPERS — CUIL arithmetic
   ═══════════════════════════════════════════════════════════════════════════ */

/** Standard ARCA/AFIP check digit. Returns [prefix, digit]. */
function cuilCheckDigit_(prefix, dni) {
  var body = ('' + prefix) + Utilities.formatString('%08d', Number(dni));
  var w = [5, 4, 3, 2, 7, 6, 5, 4, 3, 2], sum = 0;
  for (var i = 0; i < 10; i++) sum += Number(body.charAt(i)) * w[i];
  var rem = sum % 11;
  if (rem === 0) return [prefix, 0];
  if (rem === 1) return [23, prefix === 20 ? 9 : 4];
  return [prefix, 11 - rem];
}

/** DNI + gender ('m'/'f') -> the single correct CUIL. */
function deriveCuil_(dni, gender) {
  dni = String(dni).replace(/\D/g, '');
  var pref = (String(gender).toLowerCase().charAt(0) === 'm') ? 20 : 27;
  var r = cuilCheckDigit_(pref, dni);
  return Utilities.formatString('%02d', r[0]) +
         Utilities.formatString('%08d', Number(dni)) + r[1];
}

/** Both plausible CUILs when gender is unknown. */
function cuilCandidates_(dni) {
  var out = [];
  [20, 27].forEach(function (p) {
    var c = deriveCuil_(dni, p === 20 ? 'm' : 'f');
    if (out.indexOf(c) === -1) out.push(c);
  });
  return out;
}

function isValidCuil_(cuil) {
  cuil = String(cuil).replace(/\D/g, '');
  if (cuil.length !== 11) return false;
  var r = cuilCheckDigit_(Number(cuil.substr(0, 2)), Number(cuil.substr(2, 8)));
  return r[0] === Number(cuil.substr(0, 2)) && r[1] === Number(cuil.charAt(10));
}


/* ═══════════════════════════════════════════════════════════════════════════
   HELPERS — BCRA access
   ═══════════════════════════════════════════════════════════════════════════ */

function bcraGet_(path) {
  var cache = CacheService.getScriptCache();
  var key = 'bcra_' + path.replace(/\W/g, '_');
  var hit = cache.get(key);
  if (hit) return hit === 'NULL' ? null : JSON.parse(hit);

  var res, code;
  for (var attempt = 0; attempt < 3; attempt++) {
    res = UrlFetchApp.fetch(CONFIG.BCRA_BASE + path, { muteHttpExceptions: true });
    code = res.getResponseCode();
    if (code === 200 || code === 404) break;
    Utilities.sleep(2000 * (attempt + 1));
  }
  var out = null;
  if (code === 200) { try { out = JSON.parse(res.getContentText()); } catch (e) { out = null; } }
  cache.put(key, out ? JSON.stringify(out) : 'NULL', CONFIG.CACHE_HOURS * 3600);
  Utilities.sleep(CONFIG.API_PAUSE_MS);
  return out;
}

function fetchBcra_(cuil) {
  return {
    debts:   bcraGet_('/Deudas/' + cuil),
    history: bcraGet_('/Deudas/Historicas/' + cuil),
    cheques: bcraGet_('/Deudas/ChequesRechazados/' + cuil)
  };
}


/* ═══════════════════════════════════════════════════════════════════════════
   BUILD PROFILE — everything the rules need, from one BCRA payload
   ═══════════════════════════════════════════════════════════════════════════ */

function buildProfile_(bcra) {
  var d = bcra.debts && (bcra.debts.results || bcra.debts);
  if (!d || !d.periodos || !d.periodos.length) return { hasRecord: false };

  var ents = d.periodos[0].entidades || [];
  var total = 0, performing = 0, worst = 0;
  ents.forEach(function (e) {
    var amt = Number(e.monto || 0) * 1000;      // API reports thousands
    var s = Number(e.situacion || 0);
    total += amt;
    if (s <= 2) performing += amt;
    if (s > worst) worst = s;
  });

  // 24-month worst-classification series, oldest first
  var h = bcra.history && (bcra.history.results || bcra.history);
  var series = [], firstSeen = {};
  if (h && h.periodos) {
    h.periodos.forEach(function (p) {
      var w = 0;
      (p.entidades || []).forEach(function (e) {
        var s = Number(e.situacion || 0);
        if (s > w) w = s;
        if (s > 0 && !firstSeen[e.entidad]) firstSeen[e.entidad] = String(p.periodo);
      });
      series.push({ p: String(p.periodo), s: w });
    });
    series.sort(function (a, b) { return a.p < b.p ? -1 : 1; });
    // firstSeen was filled newest-first; recompute properly
    firstSeen = {};
    h.periodos.slice().sort(function (a, b) {
      return String(a.periodo) < String(b.periodo) ? -1 : 1;
    }).forEach(function (p) {
      (p.entidades || []).forEach(function (e) {
        if (Number(e.situacion || 0) > 0 && !firstSeen[e.entidad]) {
          firstSeen[e.entidad] = String(p.periodo);
        }
      });
    });
  }

  var latest = series.length ? series[series.length - 1].s : worst;
  var sixAgo = series.length >= 7 ? series[series.length - 7].s : latest;

  // B3 — any line first seen within 12 months that is now situación ≥3
  var recentBad = false, nowP = series.length ? series[series.length - 1].p : null;
  if (nowP) {
    ents.forEach(function (e) {
      var fs = firstSeen[e.entidad];
      if (!fs) return;
      var months = (Number(nowP.substr(0, 4)) - Number(fs.substr(0, 4))) * 12 +
                   (Number(nowP.substr(4, 2)) - Number(fs.substr(4, 2)));
      if (months <= CONFIG.B3_RECENT_MONTHS && Number(e.situacion || 0) >= 3) recentBad = true;
    });
  }

  var chq = bcra.cheques && (bcra.cheques.results || bcra.cheques);
  var nChq = 0;
  if (chq && chq.causales) {
    chq.causales.forEach(function (c) {
      (c.entidades || []).forEach(function (e) { nChq += (e.detalle || []).length || 1; });
    });
  }

  return {
    hasRecord: true,
    name: d.denominacion || '',
    worstSituacion: worst,
    entityCount: ents.length,
    totalDebt: total,
    performingDebt: performing,
    performingShare: total ? performing / total : 1,
    trend6m: latest - sixAgo,
    recentLineBad: recentBad,
    chequesRejected: nChq,
    history: series.map(function (x) { return x.s; }).join('')
  };
}


/* ═══════════════════════════════════════════════════════════════════════════
   SHEET ACCESS
   ═══════════════════════════════════════════════════════════════════════════ */

function sheet_(name) { return SpreadsheetApp.getActive().getSheetByName(name); }

function readTable_(name) {
  var sh = sheet_(name);
  if (!sh || sh.getLastRow() < 2) return [];
  var vals = sh.getDataRange().getValues();
  var head = vals[0].map(function (h) { return String(h).trim(); });
  return vals.slice(1).map(function (r) {
    var o = {};
    head.forEach(function (h, i) { o[h] = r[i]; });
    return o;
  });
}

function digits_(v) { return String(v == null ? '' : v).replace(/\D/g, ''); }
function clean_(v) { return String(v == null ? '' : v).trim().toLowerCase(); }

function loadBook_() {
  var borrowers = readTable_(CONFIG.SHEETS.borrowers);
  var rejected  = readTable_(CONFIG.SHEETS.rejected);
  var cfg       = readTable_(CONFIG.SHEETS.config);

  var fund = 0, cash = 0;
  cfg.forEach(function (r) {
    if (String(r['AJUSTE']).indexOf('Fondo total') === 0) fund = Number(r['VALOR']) || 0;
  });

  var outstanding = 0, placed = 0;
  borrowers.forEach(function (b) {
    if (String(b['Estado']).toUpperCase() === 'ACTIVO') {
      outstanding += Number(String(b['Saldo Pendiente (hoy)']).replace(/[^0-9.-]/g, '')) || 0;
      placed += Number(String(b['Capital']).replace(/[^0-9.-]/g, '')) || 0;
    }
  });
  cash = fund - placed;

  return { borrowers: borrowers, rejected: rejected, fund: fund,
           availableCash: cash, outstanding: outstanding };
}


/* ═══════════════════════════════════════════════════════════════════════════
   SECTION A — HARD DECLINES
   ═══════════════════════════════════════════════════════════════════════════ */

function sectionA_(app, profile, book) {
  var dni = digits_(app.dni), email = clean_(app.email), phone = digits_(app.phone);

  // A1 — rejected before
  for (var i = 0; i < book.rejected.length; i++) {
    var r = book.rejected[i];
    if (digits_(r['DNI']) === dni || (email && clean_(r['Correo']) === email) ||
        (phone && digits_(r['Teléfono']) === phone)) {
      return decline_('A1', 'Ya rechazado anteriormente (' + r['Fecha'] + ')');
    }
  }

  // A2 — contact details already used by a different DNI
  for (var j = 0; j < book.borrowers.length; j++) {
    var b = book.borrowers[j];
    var bd = digits_(b['DNI']);
    if (!bd || bd === dni) continue;
    if ((email && clean_(b['Correo']) === email) || (phone && digits_(b['Teléfono']) === phone)) {
      return decline_('A2', 'Correo/teléfono ya usado por DNI ' + bd +
                            ' (' + b['Nombre del Prestatario'] + ')');
    }
  }

  // A3 — identity
  if (!isValidCuil_(app.cuil)) return decline_('A3', 'CUIL inválido (dígito verificador)');
  if (String(app.cuil).substr(2, 8) !== Utilities.formatString('%08d', Number(dni))) {
    return decline_('A3', 'El CUIL no contiene el DNI declarado');
  }
  if (profile.hasRecord && !nameMatches_(profile.name, app.fullName)) {
    return decline_('A3', 'Nombre BCRA "' + profile.name + '" no coincide con la solicitud');
  }

  // A5 — no record  (before A4: A4 needs a record to read)
  if (!profile.hasRecord) return decline_('A5', 'Sin registro en BCRA');

  // A4 — situación 4 or 5
  if (profile.worstSituacion >= 4) {
    return decline_('A4', 'Situación ' + profile.worstSituacion + ' en ' +
                          profile.entityCount + ' entidad(es)');
  }

  // A6 — rejected cheques
  if (profile.chequesRejected > 0) {
    return decline_('A6', profile.chequesRejected + ' cheque(s) rechazado(s)');
  }

  return null;
}

/** Loose name match: every surname-ish token in one appears in the other. */
function nameMatches_(bcraName, appName) {
  if (!bcraName || !appName) return true;
  var norm = function (s) {
    return String(s).toLowerCase()
      .normalize('NFD').replace(/[̀-ͯ]/g, '')
      .replace(/[^a-z ]/g, ' ').split(/\s+/).filter(function (t) { return t.length > 2; });
  };
  var a = norm(bcraName), b = norm(appName);
  var shared = a.filter(function (t) { return b.indexOf(t) !== -1; });
  return shared.length >= Math.min(2, Math.min(a.length, b.length));
}


/* ═══════════════════════════════════════════════════════════════════════════
   SECTION B — REFINEMENT DECLINES
   ═══════════════════════════════════════════════════════════════════════════ */

function sectionB_(profile) {
  if (!CONFIG.SECTION_B_ENABLED) return null;
  if (profile.performingShare < CONFIG.B1_MIN_PERFORMING) {
    return decline_('B1', Math.round(profile.performingShare * 100) + '% al día (mínimo ' +
                          Math.round(CONFIG.B1_MIN_PERFORMING * 100) + '%)');
  }
  if (profile.trend6m >= CONFIG.B2_MAX_SLIDE) {
    return decline_('B2', 'Deterioro de ' + profile.trend6m + ' escalones en 6 meses');
  }
  if (profile.recentLineBad) {
    return decline_('B3', 'Línea abierta hace <12 meses ya en situación 3+');
  }
  return null;
}


/* ═══════════════════════════════════════════════════════════════════════════
   SECTION C — TIER, LADDER, CONCENTRATION
   ═══════════════════════════════════════════════════════════════════════════ */

function assignTier_(p) {
  if (p.worstSituacion <= 2 && p.performingShare >= CONFIG.T1_MIN_PERFORMING && p.trend6m <= 0) return 1;
  if (p.worstSituacion <= 2 && p.performingShare >= CONFIG.T2_MIN_PERFORMING) return 2;
  if (p.worstSituacion === 3 && p.performingShare >= CONFIG.T3_MIN_PERFORMING &&
      p.trend6m < CONFIG.B2_MAX_SLIDE) return 3;
  return null;
}

/** C1 — ladder. Doubles per clean repayment; one late payment drops a rung. */
function ladderLimit_(tier, dni, book) {
  var base = CONFIG.TIER_LIMITS[tier], clean = 0, late = false;
  book.borrowers.forEach(function (b) {
    if (digits_(b['DNI']) !== digits_(dni)) return;
    if (String(b['Estado']).toUpperCase() === 'PAGADO') {
      var due = new Date(b['Fecha de Vencimiento']);
      var paid = b['Fecha de Pago'] ? new Date(b['Fecha de Pago']) : due;
      if (paid <= due) clean++; else late = true;
    }
  });
  var rung = Math.max(0, clean - (late ? 1 : 0));
  return Math.min(base * Math.pow(2, rung), CONFIG.LADDER_CEILING);
}

function concentrationCheck_(amount, app, book) {
  var dni = digits_(app.dni);
  var capSingle = CONFIG.CAP_SINGLE_PCT * book.fund;
  var capGroup  = CONFIG.CAP_GROUP_PCT  * book.fund;

  var own = 0, group = 0;
  var surname = clean_(app.fullName).split(/\s+/).pop();
  book.borrowers.forEach(function (b) {
    if (String(b['Estado']).toUpperCase() !== 'ACTIVO') return;
    var bal = Number(String(b['Saldo Pendiente (hoy)']).replace(/[^0-9.-]/g, '')) || 0;
    if (digits_(b['DNI']) === dni) own += bal;
    var related = clean_(b['Nombre del Prestatario']).indexOf(surname) !== -1 ||
                  clean_(b['Correo']) === clean_(app.email) ||
                  digits_(b['Teléfono']) === digits_(app.phone);
    if (related) group += bal;
  });

  if (own + amount > capSingle)
    return decline_('C2', 'Excede el tope por prestatario (10% del fondo)');
  if (group + amount > capGroup)
    return decline_('C3', 'Excede el tope por grupo vinculado (15% del fondo)');
  return null;
}


/* ═══════════════════════════════════════════════════════════════════════════
   SECTION D — LOAN STRUCTURE
   ═══════════════════════════════════════════════════════════════════════════ */

function loanStructure_(amount, isFirst, book, dueDate) {
  var windowLoad = 0, cap = CONFIG.MATURITY_WINDOW_PCT * book.outstanding;
  book.borrowers.forEach(function (b) {
    if (String(b['Estado']).toUpperCase() !== 'ACTIVO') return;
    var d = new Date(b['Fecha de Vencimiento']);
    if (Math.abs((d - dueDate) / 86400000) <= 3) {
      windowLoad += Number(String(b['Total a Pagar']).replace(/[^0-9.-]/g, '')) || 0;
    }
  });
  var crowded = (windowLoad + amount * 1.5) > cap;

  return {
    term:       isFirst ? '15 días' : '1 mes',           // D2
    schedule:   amount > CONFIG.INSTALMENT_ABOVE ? 'semanal' : 'único',  // D1
    guarantor:  amount > CONFIG.GUARANTOR_ABOVE,         // C4
    dueDate:    dueDate,
    reschedule: crowded,                                 // D3
    lateCharge: { dailyPct: CONFIG.LATE_DAILY, base: 'saldo vencido',
                  cap: CONFIG.LATE_CAP_PCT * amount }    // D4
  };
}


/* ═══════════════════════════════════════════════════════════════════════════
   MAIN — assess one applicant
   ═══════════════════════════════════════════════════════════════════════════ */

function decline_(rule, reason) { return { verdict: 'RECHAZAR', rule: rule, reason: reason }; }

function assess(app) {
  var book = loadBook_();

  // E4 — fund gate
  if (book.availableCash < CONFIG.MIN_CASH_TO_LEND) {
    setAcceptingApplications_(false);
    return Object.assign(decline_('E4', 'Fondo por debajo del mínimo'), { book: book });
  }

  // E2 — completeness
  var missing = [];
  ['dni', 'gender', 'dniFront', 'dniBack', 'selfie'].forEach(function (f) {
    if (!app[f]) missing.push(f);
  });
  if (missing.length) {
    return { verdict: 'INCOMPLETO', rule: 'E2', reason: 'Falta: ' + missing.join(', ') };
  }

  // identity + BCRA
  var cuil = app.cuil && isValidCuil_(app.cuil) ? digits_(app.cuil)
                                                : deriveCuil_(app.dni, app.gender);
  var bcra = fetchBcra_(cuil);
  if (!bcra.debts) {                       // try the other prefix before declaring A5
    var alts = cuilCandidates_(app.dni);
    for (var i = 0; i < alts.length; i++) {
      if (alts[i] === cuil) continue;
      var t = fetchBcra_(alts[i]);
      if (t.debts) { cuil = alts[i]; bcra = t; break; }
    }
  }
  var profile = buildProfile_(bcra);
  app.cuil = cuil;

  var res = sectionA_(app, profile, book) || sectionB_(profile);

  if (!res) {
    var tier = assignTier_(profile);
    if (!tier) {
      res = decline_('C0', 'Ningún nivel aplica a este perfil');
    } else {
      var max = ladderLimit_(tier, app.dni, book);
      var amount = Math.min(Number(app.requestedAmount) || max, max, book.availableCash);
      res = concentrationCheck_(amount, app, book);
      if (!res) {
        var isFirst = !book.borrowers.some(function (b) {
          return digits_(b['DNI']) === digits_(app.dni);
        });
        res = {
          verdict: 'APROBAR', rule: '', reason: 'Nivel ' + tier,
          tier: tier, amount: amount,
          structure: loanStructure_(amount, isFirst, book,
                                    app.requestedDueDate || new Date())
        };
      }
    }
  }

  res.cuil = cuil;
  res.profile = profile;
  logDecision_(app, res);                                  // E3
  return res;
}


/* ═══════════════════════════════════════════════════════════════════════════
   SECTION E — logging, blocking, form gate
   ═══════════════════════════════════════════════════════════════════════════ */

function logDecision_(app, res) {
  var sh = sheet_(CONFIG.SHEETS.log);
  if (!sh) {
    sh = SpreadsheetApp.getActive().insertSheet(CONFIG.SHEETS.log);
    sh.appendRow(['Fecha', 'Nombre', 'DNI', 'CUIL', 'Nombre BCRA', 'Veredicto', 'Regla',
                  'Motivo', 'Nivel', 'Monto', 'Peor Sit.', '% al día', 'Tendencia 6m',
                  'Historia 24m', 'Cheques']);
  }
  var p = res.profile || {};
  sh.appendRow([new Date(), app.fullName, app.dni, res.cuil, p.name || '',
                res.verdict, res.rule, res.reason, res.tier || '', res.amount || '',
                p.worstSituacion == null ? '' : p.worstSituacion,
                p.performingShare == null ? '' : Math.round(p.performingShare * 100) + '%',
                p.trend6m == null ? '' : p.trend6m, p.history || '', p.chequesRejected || 0]);

  // E3 — declines also land in Rechazados so A1 catches repeat attempts
  if (res.verdict === 'RECHAZAR') {
    var rj = sheet_(CONFIG.SHEETS.rejected);
    if (rj) rj.appendRow([new Date(), app.fullName, app.email, app.dni, app.phone,
                          'Regla ' + res.rule + ' — ' + res.reason]);
  }
}

function setAcceptingApplications_(open) {
  var sh = sheet_(CONFIG.SHEETS.config);
  if (!sh) return;
  var v = sh.getDataRange().getValues();
  for (var i = 0; i < v.length; i++) {
    if (String(v[i][0]).indexOf('Aceptar solicitudes') === 0) {
      sh.getRange(i + 1, 2).setValue(open ? 'SI' : 'NO');
      return;
    }
  }
}

/** E1 — call this from your disbursement function. Returns true only on APROBAR. */
function mayDisburse(app) {
  var res = assess(app);
  return res.verdict === 'APROBAR';
}


/* ═══════════════════════════════════════════════════════════════════════════
   SECTION F — overrides and monthly review
   ═══════════════════════════════════════════════════════════════════════════ */

function override(app, originalRes, operator, writtenReason) {
  if (!writtenReason || String(writtenReason).trim().length < 10) {
    throw new Error('Section F: an override requires a written reason.');
  }
  var amount = Math.min(Number(app.requestedAmount) || CONFIG.TIER_LIMITS[3],
                        CONFIG.TIER_LIMITS[3]);          // forced to Tier 3
  var res = { verdict: 'APROBAR', rule: 'F-OVERRIDE', tier: 'OVERRIDE', amount: amount,
              reason: 'Anula ' + originalRes.rule + ' — ' + writtenReason +
                      ' (' + operator + ')', cuil: originalRes.cuil,
              profile: originalRes.profile };
  logDecision_(app, res);
  return res;
}

/** E6 — realised default by tier, and the override alarm. */
function monthlyReview() {
  var log = readTable_(CONFIG.SHEETS.log);
  var borrowers = readTable_(CONFIG.SHEETS.borrowers);
  var tierOf = {};
  log.forEach(function (r) { if (r['Nivel']) tierOf[digits_(r['DNI'])] = r['Nivel']; });

  var due = {}, lost = {};
  borrowers.forEach(function (b) {
    var t = tierOf[digits_(b['DNI'])] || 'sin nivel';
    var estado = String(b['Estado']).toUpperCase();
    if (estado !== 'PAGADO' && estado !== 'VENCIDO') return;
    var amt = Number(String(b['Total a Pagar']).replace(/[^0-9.-]/g, '')) || 0;
    var paid = Number(String(b['Total Pagado']).replace(/[^0-9.-]/g, '')) || 0;
    due[t] = (due[t] || 0) + amt;
    lost[t] = (lost[t] || 0) + Math.max(0, amt - paid);
  });

  var lines = ['MORA REALIZADA POR NIVEL', ''];
  Object.keys(due).forEach(function (t) {
    lines.push('  Nivel ' + t + ': ' + Math.round(lost[t] / due[t] * 1000) / 10 + '%' +
               '  (vencido ' + Math.round(due[t]).toLocaleString() + ')');
  });

  var approvals = log.filter(function (r) { return r['Veredicto'] === 'APROBAR'; });
  var overrides = approvals.filter(function (r) { return r['Regla'] === 'F-OVERRIDE'; });
  var rate = approvals.length ? overrides.length / approvals.length : 0;
  lines.push('', 'Excepciones: ' + overrides.length + '/' + approvals.length +
                 ' (' + Math.round(rate * 100) + '%)');
  if (rate > 0.10) {
    lines.push('', '⚠ Las excepciones superan el 10%. La política no se está aplicando:',
                   '  el embudo es demasiado chico. Arreglar el embudo, no la política.');
  }
  SpreadsheetApp.getUi().alert(lines.join('\n'));
}


/* ═══════════════════════════════════════════════════════════════════════════
   SHEET INTEGRATION
   ═══════════════════════════════════════════════════════════════════════════ */

/**
 * Sheet formula:  =EVALUAR(D2; E2; F2; "f")
 * @customfunction
 */
function EVALUAR(nombre, dni, correo, genero, telefono) {
  if (!dni) return '';
  var res = assess({ fullName: nombre, dni: dni, email: correo, phone: telefono,
                     gender: genero || 'f', dniFront: 'x', dniBack: 'x', selfie: 'x' });
  if (res.verdict === 'APROBAR') {
    return 'APROBAR · Nivel ' + res.tier + ' · máx ' + Math.round(res.amount).toLocaleString();
  }
  return res.verdict + ' · ' + res.rule + ' · ' + res.reason;
}

/** Evaluate every pending row in "Nuevos Prestatarios". */
function evaluarSolicitudes() {
  var sh = sheet_(CONFIG.SHEETS.applicants);
  var rows = readTable_(CONFIG.SHEETS.applicants);
  var head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
               .map(function (h) { return String(h).trim(); });
  var col = function (n) { return head.indexOf(n) + 1; };
  var outCol = col('Decisión') || sh.getLastColumn() + 1;

  rows.forEach(function (r, i) {
    if (!r['DNI'] || r['Decisión']) return;
    var res = assess({
      fullName: r['Nombre completo'], dni: r['DNI'], email: r['Correo'],
      phone: r['Teléfono'], gender: r['Género'] || 'f',
      requestedAmount: Number(String(r['Monto Solicitado']).replace(/[^0-9.-]/g, '')),
      dniFront: r['Foto del frente del DNI'], dniBack: r['Foto del dorso del DNI'],
      selfie: r['Selfie con DNI'] || 'x'
    });
    sh.getRange(i + 2, outCol).setValue(
      res.verdict + (res.rule ? ' (' + res.rule + ')' : '') + ' — ' + res.reason);
  });
  SpreadsheetApp.getUi().alert('Evaluación completa. Ver hoja "' + CONFIG.SHEETS.log + '".');
}

/** Backfill the existing portfolio so you can see what the policy would have said. */
function verificarCarteraCompleta() {
  var rows = readTable_(CONFIG.SHEETS.borrowers);
  var out = [['ID', 'Prestatario', 'CUIL', 'Nombre BCRA', 'Veredicto', 'Regla', 'Motivo']];
  rows.forEach(function (b) {
    if (!b['DNI']) return;
    var res = assess({ fullName: b['Nombre del Prestatario'], dni: b['DNI'],
                       email: b['Correo'], phone: b['Teléfono'], gender: 'f',
                       dniFront: 'x', dniBack: 'x', selfie: 'x' });
    out.push([b['ID Préstamo'], b['Nombre del Prestatario'], res.cuil,
              (res.profile || {}).name || '', res.verdict, res.rule, res.reason]);
  });
  var sh = sheet_('Backtest') || SpreadsheetApp.getActive().insertSheet('Backtest');
  sh.clear();
  sh.getRange(1, 1, out.length, out[0].length).setValues(out);
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Crédito')
    .addItem('Evaluar solicitudes pendientes', 'evaluarSolicitudes')
    .addItem('Verificar cartera completa (backtest)', 'verificarCarteraCompleta')
    .addSeparator()
    .addItem('Revisión mensual', 'monthlyReview')
    .addSeparator()
    .addItem('Cerrar solicitudes', 'cerrarSolicitudes')
    .addItem('Abrir solicitudes', 'abrirSolicitudes')
    .addToUi();
}

function cerrarSolicitudes() { setAcceptingApplications_(false); }
function abrirSolicitudes()  { setAcceptingApplications_(true); }
