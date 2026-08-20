/************************************************************
 * GESTOR DE PRÉSTAMOS v2  —  Google Sheets + Apps Script
 * ----------------------------------------------------------
 * Versión mejorada y en español. Reemplaza a LoanManager.gs
 * (pegue SOLO este archivo para evitar funciones duplicadas).
 *
 * MEJORAS
 *  - Interfaz 100% en español (incluye estados PAGADO/VENCIDO/ACTIVO).
 *  - Panel lateral (sidebar) con botones reales — sin dibujos.
 *  - Confirmación antes de "Configurar" (evita borrar datos).
 *  - Correos automáticos: solicitud recibida, aprobación (con contrato),
 *    recibo de pago, y recordatorios de vencimiento / mora.
 *  - Hoja "Errores" con registro visible de fallos.
 *  - Hoja "Panel" con indicadores del prestamista.
 *  - Ruta de rechazo/archivo para solicitantes.
 *  - Detección de DNI duplicado y validaciones a prueba de errores.
 *  - Formulario web reforzado + página de consulta de saldo.
 *  - Columnas de fórmula protegidas (aviso al editar).
 *
 * PUESTA EN MARCHA
 *  1) Extensions ▸ Apps Script. Pegue TODO este archivo. Guarde.
 *  2) Ejecute  ->  onOpen  (autorice: Sheets, Gmail, Drive, Forms/ScriptApp).
 *  3) Recargue la hoja. Menú "Gestor de Préstamos" ▸ ① Configurar.
 *  4) Complete "Configuración". Menú ⑥ para publicar el formulario web.
 ************************************************************/

const CFG = {
  CURRENCY_FMT: '"$"#,##0.00',
  LOCALE: 'es-AR', CURRENCY_CODE: 'ARS',
  MAX_ROWS: 300,
  MAIN_FOLDER: 'Documentos de Prestatarios',
  SHEETS: {
    BORROWERS: 'Prestatarios', PAYMENTS: 'Pagos', SUMMARY: 'Resumen', SETTINGS: 'Configuración',
    AGREEMENT: 'Estudio de Contratos', STATEMENTS: 'Estudio de Estados', REMINDER: 'Estudio de Recordatorios', NEW: 'Nuevos Prestatarios',
    PANEL: 'Panel', STATS: 'Estadísticas', LATE: 'Pagos Atrasados', CLEARED: 'Saldados', REJECTED: 'Rechazados', ERRORS: 'Errores', HELP: 'Instrucciones',
    SIGN: 'Firmas',
  },
};
const ST = { ACTIVE: 'ACTIVO', OVERDUE: 'VENCIDO', PAID: 'PAGADO', CLEARED: 'SALDADO' };
// Estado de firma del contrato (columna S de "Prestatarios").
const SIGN = { PENDING: 'PENDIENTE', SIGNED: 'FIRMADO' };
// Columnas de firma en "Prestatarios" (agregadas al final, no desplazan nada).
const COL_SIGN_STATUS = 19, COL_SIGN_DATE = 20, COL_SIGN_PDF = 21;
// Hojas de operaciones diarias: siempre visibles. El resto son de configuración/consulta
// y se pueden ocultar/mostrar desde el menú ▸ 🗂 Pestañas.
const MAIN_SHEETS = [
  CFG.SHEETS.BORROWERS,   // Prestatarios — cartera
  CFG.SHEETS.NEW,         // Nuevos Prestatarios — solicitudes a revisar
  CFG.SHEETS.PAYMENTS,    // Pagos — registrar cobros
  CFG.SHEETS.LATE,        // Pagos Atrasados — mora / avisos
  CFG.SHEETS.PANEL,       // Panel — indicadores
];
// Encabezados de "Nuevos Prestatarios" (el formulario agrega filas aquí, en este orden).
const NB = ['Fecha de Envío', 'Nombre completo', 'Correo', 'DNI', 'Teléfono', 'Monto Solicitado',
  'Plazo (meses)', 'Notas / Motivo', 'Foto del frente del DNI', 'Foto del dorso del DNI', 'Verificado?', 'Rechazar?', 'Resultado',
  'Verificar BCRA?', 'Parámetro BCRA', 'CUIL', 'Nombre BCRA', 'Peor Situación', '% al día', 'Decisión', 'Puntaje BCRA', 'Calificación de Riesgo', 'Resumen BCRA'];

/* ============================ MENÚ ============================ */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Gestor de Préstamos')
    .addItem('▶ Abrir panel del prestamista', 'openSidebar')
    .addSeparator()
    .addItem('① Configurar / reconstruir hojas', 'setupConfirm')
    .addItem('⑥ Formulario web — publicar / ver enlace', 'showWebFormLink')
    .addItem('Configurar marca (nombre + logo)', 'setBranding')
    .addItem('④ Estadísticas y gráficos', 'showStats')
    .addItem('⑤ Actualizar saldos y resumen', 'refreshAll')
    .addItem('➕ Crear hoja de Recordatorios de pago', 'createRemindersSheet_')
    .addItem('Agregar columnas BCRA a Nuevos Prestatarios', 'setupNewOnly')
    .addItem('Verificar cartera en BCRA', 'backfillPrestatarios')
    .addSeparator()
    .addItem('✍ Actualizar Prestatarios: columnas de firma', 'addSigningColumns_')
    .addItem('✉ Reenviar enlace de firma (fila seleccionada)', 'resendSigningLink_')
    .addSeparator()
    .addItem('Registrar un pago', 'openSidebar')
    .addItem('✔ Mover préstamos saldados → hoja Saldados', 'moveClearedBorrowersConfirm')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('🗂 Pestañas')
      .addItem('Ver solo operaciones diarias', 'showDailyOnly')
      .addItem('Mostrar todas las hojas', 'showAllSheets'))
    .addSubMenu(SpreadsheetApp.getUi().createMenu('🧪 Datos de prueba')
      .addItem('Crear 40 prestatarios de prueba', 'seedTestBorrowersConfirm')
      .addItem('Eliminar prestatarios de prueba', 'removeTestBorrowersConfirm'))
    .addToUi();
}

function openSidebar() {
  const html = HtmlService.createHtmlOutput(sidebarHtml_()).setTitle('Panel del Prestamista');
  SpreadsheetApp.getUi().showSidebar(html);
}

/* ===================== REGISTRO DE ERRORES ===================== */
function logError_(where, err) {
  try {
    const sh = getSS_().getSheetByName(CFG.SHEETS.ERRORS) || getSS_().insertSheet(CFG.SHEETS.ERRORS);
    if (sh.getLastRow() === 0)
      sh.getRange(1, 1, 1, 3).setValues([['Fecha', 'Función', 'Error']]).setFontWeight('bold');
    sh.appendRow([new Date(), where, (err && err.message) ? err.message : String(err)]);
  } catch (e) { /* nunca fallar por el log */ }
}
function guard_(where, fn) { try { return fn(); } catch (err) { logError_(where, err); throw err; } }

/* ===================== MARCA (nombre + logo) ===================== */
function setBranding() {
  const ui = SpreadsheetApp.getUi();
  const n = ui.prompt('Marca', 'Nombre de la empresa:', ui.ButtonSet.OK_CANCEL);
  if (n.getSelectedButton() !== ui.Button.OK) return;
  setSetting_('Nombre del Prestamista', n.getResponseText().trim() || 'Impulso Crédito');
  const l = ui.prompt('Marca', 'Enlace de Drive del logo (pegue el enlace para compartir; opcional):', ui.ButtonSet.OK_CANCEL);
  if (l.getSelectedButton() === ui.Button.OK) setSetting_('Logo (enlace de Drive)', l.getResponseText().trim());
  _logoBlob = null; _logoLoaded = false; _logoUri = null;
  ui.alert('Marca actualizada. El logo aparecerá en correos, contratos y páginas web.');
}
function setSetting_(key, value) {
  const sh = getSS_().getSheetByName(CFG.SHEETS.SETTINGS);
  const data = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1).getValues();
  for (let i = 0; i < data.length; i++) if (String(data[i][0]).trim() === key) { sh.getRange(i + 2, 2).setValue(value); return; }
  sh.appendRow([key, value]);
}

var _logoBlob = null, _logoLoaded = false, _logoUri = null;
function getLogoBlob_() {
  if (_logoLoaded) return _logoBlob;
  _logoLoaded = true; _logoBlob = null;
  const ref = String(getSetting_('Logo (enlace de Drive)') || '').trim();
  const ids = extractDriveIds_(ref);
  if (ids.length) { try { const b = DriveApp.getFileById(ids[0]).getBlob(); b.setName('logo'); _logoBlob = b; } catch (e) { logError_('getLogoBlob_', e); } }
  return _logoBlob;
}
function logoDataUri_() {
  if (_logoUri !== null) return _logoUri;
  const b = getLogoBlob_();
  _logoUri = b ? ('data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes())) : '';
  return _logoUri;
}
/**
 * Firma del prestamista (dueño de Impulso Crédito) como data URI, para estampar
 * en TODOS los contratos. Prioridad: 1) "Firma del Prestamista (enlace de Drive)"
 * en Configuración; 2) firma embebida por defecto (OWNER_SIGNATURE_DATAURI,
 * en el archivo SignatureData.gs). Devuelve '' si no hay ninguna.
 */
var _ownerSigUri = null;
function ownerSignatureDataUri_() {
  if (_ownerSigUri !== null) return _ownerSigUri;
  _ownerSigUri = '';
  const ref = String(getSetting_('Firma del Prestamista (enlace de Drive)') || '').trim();
  const ids = extractDriveIds_(ref);
  if (ids.length) {
    try {
      const b = DriveApp.getFileById(ids[0]).getBlob();
      _ownerSigUri = 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes());
      return _ownerSigUri;
    } catch (e) { logError_('ownerSignatureDataUri_', e); }
  }
  if (typeof OWNER_SIGNATURE_DATAURI === 'string' && OWNER_SIGNATURE_DATAURI) _ownerSigUri = OWNER_SIGNATURE_DATAURI;
  return _ownerSigUri;
}
/**
 * Firma del segundo titular de Impulso Crédito (Xoana Elizabeth Beron) como data URI.
 * Se estampa, junto con la del titular principal, en TODOS los contratos.
 * Prioridad: 1) "Firma del Prestamista 2 (enlace de Drive)" en Configuración;
 * 2) COOWNER_SIGNATURE_DATAURI (en SignatureData.gs). Devuelve '' si no hay ninguna.
 */
var _coOwnerSigUri = null;
function coOwnerSignatureDataUri_() {
  if (_coOwnerSigUri !== null) return _coOwnerSigUri;
  _coOwnerSigUri = '';
  const ref = String(getSetting_('Firma del Prestamista 2 (enlace de Drive)') || '').trim();
  const ids = extractDriveIds_(ref);
  if (ids.length) {
    try {
      const b = DriveApp.getFileById(ids[0]).getBlob();
      _coOwnerSigUri = 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes());
      return _coOwnerSigUri;
    } catch (e) { logError_('coOwnerSignatureDataUri_', e); }
  }
  if (typeof COOWNER_SIGNATURE_DATAURI === 'string' && COOWNER_SIGNATURE_DATAURI) _coOwnerSigUri = COOWNER_SIGNATURE_DATAURI;
  return _coOwnerSigUri;
}
/** Nombres completos de los dos firmantes titulares (configurables). */
function ownerSignatories_() {
  return [
    { name: String(getSetting_('Firmante 1 (nombre completo)') || 'Denis Delroy Bell').trim(), sig: ownerSignatureDataUri_() },
    { name: String(getSetting_('Firmante 2 (nombre completo)') || 'Xoana Elizabeth Beron').trim(), sig: coOwnerSignatureDataUri_() },
  ];
}
function companyName_() {
  const v = String(getSetting_('Nombre del Prestamista') || '').trim();
  return (!v || v === 'Su Nombre / Empresa' || v === 'Su Nombre / Company') ? 'Impulso Crédito' : v;
}
/** Recargo por mora como fracción (5% => 0.05). Predeterminado 5% si no está configurado. */
function lateFeeRate_() {
  const raw = String(getSetting_('Recargo por Mora diario (%)') || getSetting_('Recargo por Mora (%)') || '').trim();
  const v = raw === '' ? 5 : Number(raw.replace(',', '.'));
  return (isFinite(v) && v > 0) ? v / 100 : 0;
}
function lateFeePctText_() { return round2_(lateFeeRate_() * 100) + '%'; }

// Colores de marca Impulso Crédito
var BRAND = { navy: '#1c4587', amber: '#e8a33d', tag: 'Préstamos rápidos y confiables' };
function brandInitials_() {
  const ini = companyName_().trim().split(/\s+/).filter(Boolean).map(function (w) { return w.charAt(0); }).join('').slice(0, 2).toUpperCase();
  return ini || 'IC';
}
function brandWordmark_(size) {
  const parts = companyName_().trim().split(/\s+/);
  const first = esc_(parts.shift() || 'Impulso'), rest = esc_(parts.join(' '));
  return '<span style="font-size:' + size + 'px;font-weight:bold;color:' + BRAND.navy + '">' + first + '</span>' +
    (rest ? '<span style="font-size:' + size + 'px;font-weight:bold;color:' + BRAND.amber + '"> ' + rest + '</span>' : '');
}
function brandBadge_(px, cid) {
  if (cid && getLogoBlob_()) return '<img src="cid:logo" style="max-height:' + px + 'px;max-width:' + (px + 8) + 'px">';
  if (!cid) { const u = logoDataUri_(); if (u) return '<img src="' + u + '" style="max-height:' + px + 'px">'; }
  const s = Math.round(px * 0.92);
  return '<span style="display:inline-block;width:' + s + 'px;height:' + s + 'px;line-height:' + s +
    'px;text-align:center;background:' + BRAND.navy + ';color:#fff;border-radius:' + Math.round(s / 4.5) +
    'px;font-size:' + Math.round(s / 2.2) + 'px;font-weight:bold;font-family:Arial,Helvetica,sans-serif">' + esc_(brandInitials_()) + '</span>';
}

/** Cabecera de marca para correos (logo en línea vía cid, o monograma CSS). */
function brandHeaderEmail_() {
  return '<table style="border-collapse:collapse;width:100%;border-bottom:3px solid ' + BRAND.navy + ';margin-bottom:14px"><tr>' +
    '<td style="width:62px;padding:0 10px 8px 0;vertical-align:middle">' + brandBadge_(50, true) + '</td>' +
    '<td style="padding-bottom:8px;vertical-align:middle">' + brandWordmark_(21) +
    '<div style="font-size:11px;color:#888">' + esc_(BRAND.tag) + '</div></td></tr></table>';
}
function brandFooterEmail_() {
  const c = esc_(companyName_()), phone = esc_(getSetting_('Teléfono del Prestamista')),
    email = esc_(getSetting_('Correo del Prestamista')), addr = esc_(getSetting_('Dirección del Prestamista'));
  const bits = [c, addr, phone, email].filter(function (x) { return x; }).join(' · ');
  return '<p style="margin-top:22px;font-size:12px;color:#888;border-top:1px solid #ddd;padding-top:8px">' + bits + '</p>';
}
/** Envía un correo con la cabecera/pie de marca y el logo en línea. opts.attachments opcional. */
function sendBrandedEmail_(to, subject, plainBody, htmlInner, opts) {
  opts = opts || {};
  const html = '<div style="font-family:Arial,Helvetica,sans-serif;color:#222;max-width:640px">' +
    brandHeaderEmail_() + htmlInner + brandFooterEmail_() + '</div>';
  const o = { name: companyName_(), htmlBody: html };
  const logo = getLogoBlob_(); if (logo) o.inlineImages = { logo: logo };
  if (opts.attachments) o.attachments = opts.attachments;
  GmailApp.sendEmail(to, subject, plainBody, o);
}
/** Cabecera de marca para PDF y páginas web (data URI, o monograma CSS). */
function brandHeaderHtml_() {
  return '<div style="text-align:center;margin-bottom:12px">' + brandBadge_(64, false) +
    '<div style="margin-top:6px">' + brandWordmark_(24) + '</div>' +
    '<div style="font-size:12px;color:#888">' + esc_(BRAND.tag) + '</div></div>';
}

/* ============================ SETUP ============================ */
function setupConfirm() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.alert('Configurar', 'Esto RECONSTRUYE las hojas Prestatarios y Pagos y BORRA sus datos. ' +
    '¿Desea continuar?', ui.ButtonSet.YES_NO);
  if (r === ui.Button.YES) setup();
}

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  PropertiesService.getScriptProperties().setProperty('SS_ID', ss.getId());
  setupSettings_(ss); setupBorrowers_(ss); setupPayments_(ss); setupSummary_(ss);
  setupAgreement_(ss); setupStatements_(ss); setupReminders_(ss); setupNew_(ss); setupPanel_(ss); setupStats_(ss); setupLate_(ss); setupCleared_(ss); setupRejected_(ss);
  setupErrors_(ss); setupHelp_(ss); setupFirmas_(ss);
  rebuildLateSheet_(ss);
  ensureTriggers_();
  const s1 = ss.getSheetByName('Sheet1') || ss.getSheetByName('Hoja 1');
  if (s1 && ss.getSheets().length > 1) { try { ss.deleteSheet(s1); } catch (e) {} }
  applyDailyVisibility_(ss);
  SpreadsheetApp.getUi().alert('Listo. Se dejaron visibles las pestañas de operaciones diarias; ' +
    'las demás (Configuración, Estadísticas, etc.) están ocultas — use el menú ▸ 🗂 Pestañas para mostrarlas. ' +
    'Complete "Configuración" y publique el formulario con el menú ⑥.');
}

function ensureTriggers_() {
  const fns = ScriptApp.getProjectTriggers().map(t => t.getHandlerFunction());
  if (fns.indexOf('onEditInstallable') === -1)
    ScriptApp.newTrigger('onEditInstallable').forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet()).onEdit().create();
  if (fns.indexOf('dailyTasks') === -1)
    ScriptApp.newTrigger('dailyTasks').timeBased().everyDays(1).atHour(7).create();
}

function setupSettings_(ss) {
  const sh = getOrCreate_(ss, CFG.SHEETS.SETTINGS); sh.clear();
  const rows = [
    ['AJUSTE', 'VALOR'],
    ['Nombre del Prestamista', 'Impulso Crédito'],
    ['Logo (enlace de Drive)', ''],
    ['Firmante 1 (nombre completo)', 'Denis Delroy Bell'],
    ['Firma del Prestamista (enlace de Drive)', ''],
    ['Firmante 2 (nombre completo)', 'Xoana Elizabeth Beron'],
    ['Firma del Prestamista 2 (enlace de Drive)', ''],
    ['URL de la app web (enlaces a clientes)', ''],
    ['Correo del Prestamista', (Session.getActiveUser().getEmail() || '')],
    ['Correos de aviso adicionales (separados por coma)', ''],
    ['Teléfono del Prestamista', ''],
    ['Dirección del Prestamista', ''],
    ['Jurisdicción', 'Buenos Aires, Argentina'],
    ['Cláusula de Mora', 'Se aplica un recargo por mora del 5% diario sobre el total a pagar por cada día de atraso posterior a la fecha de vencimiento.'],
    ['Recargo por Mora diario (%)', '5'],
    ['Pie del Contrato', 'Este acuerdo es legalmente vinculante desde la firma de ambas partes.'],
    ['Días de aviso antes del vencimiento', '3'],
    ['Fondo total para prestar', '0'],
    ['Aceptar solicitudes de préstamo (SÍ/NO)', 'SÍ'],
  ];
  sh.getRange(1, 1, rows.length, 2).setValues(rows);
  sh.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#1c4587').setFontColor('#fff');
  sh.getRange(2, 1, rows.length - 1, 1).setFontWeight('bold');
  sh.setColumnWidth(1, 260); sh.setColumnWidth(2, 500);
  sh.getRange(1, 1, rows.length, 2).setWrap(true); sh.setFrozenRows(1);
}

function setupBorrowers_(ss) {
  const sh = getOrCreate_(ss, CFG.SHEETS.BORROWERS); sh.clear();
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearDataValidations();
  const H = ['ID Préstamo', 'Nombre del Prestatario', 'Correo', 'DNI', 'Teléfono', 'Capital',
    'Plazo (meses)', 'Tasa', 'Fecha del Préstamo', 'Fecha de Vencimiento', 'Interés', 'Total a Pagar',
    'Total Pagado', 'Saldo Pendiente (hoy)', 'Estado', 'Contrato PDF', 'Contrato Enviado', 'Último Aviso',
    'Estado de Firma', 'Fecha de Firma', 'Contrato Firmado (PDF)'];
  sh.getRange(1, 1, 1, H.length).setValues([H]).setFontWeight('bold').setBackground('#1c4587').setFontColor('#fff').setWrap(true);
  sh.setFrozenRows(1);
  const N = CFG.MAX_ROWS, P = CFG.SHEETS.PAYMENTS;
  const fH = [], fJ = [], fKLM = [], fO = [];
  for (let r = 2; r <= N + 1; r++) {
    fH.push([`=IF($G${r}="","",IFS($G${r}=15,0.25,$G${r}=1,0.5,$G${r}=2,1,TRUE,"⚠ PLAZO INVÁLIDO"))`]);
    fJ.push([`=IF(OR($I${r}="",$G${r}=""),"",IF($G${r}=15,$I${r}+15,EDATE($I${r},$G${r})))`]);
    fKLM.push([
      `=IF(OR($F${r}="",NOT(ISNUMBER($H${r}))),"",$F${r}*$H${r})`,
      `=IF($K${r}="","",$F${r}+$K${r})`,
      `=IF($A${r}="","",SUMIF('${P}'!$B:$B,$A${r},'${P}'!$G:$G))`,
    ]);
    fO.push([`=IF($A${r}="","",IF($N${r}="","…",IF($N${r}<=0.009,"${ST.PAID}",IF(AND($J${r}<>"",TODAY()>$J${r}),"${ST.OVERDUE}","${ST.ACTIVE}"))))`]);
  }
  sh.getRange(2, 8, N, 1).setFormulas(fH);
  sh.getRange(2, 10, N, 1).setFormulas(fJ);
  sh.getRange(2, 11, N, 3).setFormulas(fKLM);
  sh.getRange(2, 15, N, 1).setFormulas(fO);
  sh.getRange(2, 6, N, 1).setNumberFormat(CFG.CURRENCY_FMT);
  sh.getRange(2, 8, N, 1).setNumberFormat('0%');
  sh.getRange(2, 9, N, 2).setNumberFormat('yyyy-mm-dd');
  sh.getRange(2, 11, N, 4).setNumberFormat(CFG.CURRENCY_FMT);
  sh.getRange(2, 18, N, 1).setNumberFormat('yyyy-mm-dd');
  sh.getRange(2, COL_SIGN_DATE, N, 1).setNumberFormat('yyyy-mm-dd hh:mm');
  sh.getRange(2, 7, N, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['15', '1', '2'], true).setAllowInvalid(false)
    .setHelpText('El plazo debe ser 15 días (25%), 1 mes (50%) o 2 meses (100%).').build());
  sh.getRange(2, 9, N, 1).setDataValidation(datePicker_());
  const widths = [90, 170, 200, 120, 110, 110, 90, 60, 110, 110, 110, 120, 110, 130, 90, 130, 130, 110, 120, 140, 150];
  widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));
  const stRange = sh.getRange(2, 15, N, 1);
  const signRange = sh.getRange(2, COL_SIGN_STATUS, N, 1);
  sh.setConditionalFormatRules([
    cc_(stRange, ST.PAID, '#b6d7a8'), cc_(stRange, ST.OVERDUE, '#ea9999'), cc_(stRange, ST.ACTIVE, '#fff2cc'),
    cc_(signRange, SIGN.PENDING, '#f4cccc'), cc_(signRange, SIGN.SIGNED, '#b6d7a8'),
  ]);
  sh.getRange(2, 16, N, 1).setFontColor('#1155cc');
  sh.getRange(2, COL_SIGN_PDF, N, 1).setFontColor('#1155cc');
  protectFormulas_(sh, N);
}

/**
 * Migración NO destructiva: agrega las columnas de firma (S, T, U) a
 * "Prestatarios" si faltan, sin borrar datos. Segura de correr sobre una hoja
 * con préstamos reales. Se ejecuta desde el menú.
 */
function addSigningColumns_() {
  const ss = getSS_(), sh = ss.getSheetByName(CFG.SHEETS.BORROWERS);
  if (!sh) { SpreadsheetApp.getUi().alert('No existe la hoja "Prestatarios". Ejecute ① Configurar primero.'); return; }
  const N = CFG.MAX_ROWS;
  const head = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), COL_SIGN_PDF)).getValues()[0]
    .map(h => String(h).trim());
  const labels = { [COL_SIGN_STATUS]: 'Estado de Firma', [COL_SIGN_DATE]: 'Fecha de Firma', [COL_SIGN_PDF]: 'Contrato Firmado (PDF)' };
  [COL_SIGN_STATUS, COL_SIGN_DATE, COL_SIGN_PDF].forEach(c => {
    if (head[c - 1] !== labels[c]) {
      sh.getRange(1, c).setValue(labels[c]).setFontWeight('bold').setBackground('#1c4587').setFontColor('#fff').setWrap(true);
    }
  });
  sh.setColumnWidth(COL_SIGN_STATUS, 120); sh.setColumnWidth(COL_SIGN_DATE, 140); sh.setColumnWidth(COL_SIGN_PDF, 150);
  sh.getRange(2, COL_SIGN_DATE, N, 1).setNumberFormat('yyyy-mm-dd hh:mm');
  sh.getRange(2, COL_SIGN_PDF, N, 1).setFontColor('#1155cc');
  // Marca como PENDIENTE los préstamos existentes sin estado de firma (activos/vencidos).
  const last = sh.getLastRow();
  if (last >= 2) {
    const ids = sh.getRange(2, 1, last - 1, 1).getValues();
    const est = sh.getRange(2, 15, last - 1, 1).getValues();
    const cur = sh.getRange(2, COL_SIGN_STATUS, last - 1, 1).getValues();
    const out = cur.map((r, i) => {
      const has = String(r[0]).trim();
      if (has) return [has];
      if (!String(ids[i][0]).trim()) return [''];
      const e = String(est[i][0]).toUpperCase();
      return [(e === ST.PAID || e === ST.CLEARED) ? SIGN.SIGNED : SIGN.PENDING];
    });
    sh.getRange(2, COL_SIGN_STATUS, out.length, 1).setValues(out);
  }
  // Reglas de formato condicional para la columna de firma (preservando las de Estado).
  const stRange = sh.getRange(2, 15, N, 1), signRange = sh.getRange(2, COL_SIGN_STATUS, N, 1);
  sh.setConditionalFormatRules([
    cc_(stRange, ST.PAID, '#b6d7a8'), cc_(stRange, ST.OVERDUE, '#ea9999'), cc_(stRange, ST.ACTIVE, '#fff2cc'),
    cc_(signRange, SIGN.PENDING, '#f4cccc'), cc_(signRange, SIGN.SIGNED, '#b6d7a8'),
  ]);
  setupFirmas_(ss);
  // Asegura (sin sobrescribir) las nuevas filas de Configuración para hojas existentes.
  if (ss.getSheetByName(CFG.SHEETS.SETTINGS)) {
    setSetting_('URL de la app web (enlaces a clientes)', getSetting_('URL de la app web (enlaces a clientes)') || '');
    setSetting_('Firmante 1 (nombre completo)', getSetting_('Firmante 1 (nombre completo)') || 'Denis Delroy Bell');
    setSetting_('Firma del Prestamista (enlace de Drive)', getSetting_('Firma del Prestamista (enlace de Drive)') || '');
    setSetting_('Firmante 2 (nombre completo)', getSetting_('Firmante 2 (nombre completo)') || 'Xoana Elizabeth Beron');
    setSetting_('Firma del Prestamista 2 (enlace de Drive)', getSetting_('Firma del Prestamista 2 (enlace de Drive)') || '');
  }
  try { ss.toast('Columnas de firma agregadas. Complete la URL de la app web en Configuración.', 'Listo', 7); } catch (e) {}
}

/** Hoja de auditoría de firmas. */
function setupFirmas_(ss) {
  const sh = getOrCreate_(ss, CFG.SHEETS.SIGN);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, 8).setValues([['Fecha y hora', 'ID Préstamo', 'Nombre', 'DNI', 'Correo', 'Método', 'Dispositivo', 'Contrato Firmado (PDF)']])
      .setFontWeight('bold').setBackground('#38761d').setFontColor('#fff').setWrap(true);
    [150, 100, 180, 110, 200, 90, 260, 150].forEach((w, i) => sh.setColumnWidth(i + 1, w));
    sh.getRange(2, 1, CFG.MAX_ROWS, 1).setNumberFormat('yyyy-mm-dd hh:mm');
  }
  sh.setFrozenRows(1);
  return sh;
}

/** Aviso (no bloqueo) al editar columnas calculadas. */
function protectFormulas_(sh, N) {
  try {
    ['H', 'J', 'K', 'L', 'M', 'N', 'O'].forEach(col => {
      const p = sh.getRange(col + '2:' + col + (N + 1)).protect()
        .setDescription('Columna calculada — no editar');
      p.setWarningOnly(true);
    });
  } catch (e) { /* ignora si no hay permisos */ }
}

function setupPayments_(ss) {
  const sh = getOrCreate_(ss, CFG.SHEETS.PAYMENTS); sh.clear();
  const H = ['ID Pago', 'ID Préstamo', 'Nombre', 'Apellido', 'DNI', 'Fecha de Pago', 'Monto Pagado', 'Saldo Posterior', 'Recibo Enviado'];
  sh.getRange(1, 1, 1, H.length).setValues([H]).setFontWeight('bold').setBackground('#38761d').setFontColor('#fff');
  sh.setFrozenRows(1);
  sh.getRange(2, 6, CFG.MAX_ROWS, 1).setNumberFormat('yyyy-mm-dd');
  sh.getRange(2, 7, CFG.MAX_ROWS, 2).setNumberFormat(CFG.CURRENCY_FMT);
  [140, 90, 120, 120, 120, 110, 120, 130, 150].forEach((w, i) => sh.setColumnWidth(i + 1, w));
}

function setupSummary_(ss) {
  const sh = getOrCreate_(ss, CFG.SHEETS.SUMMARY); sh.clear();
  const B = CFG.SHEETS.BORROWERS;
  sh.getRange(1, 1, 1, 7).setValues([['Prestatario', 'Préstamos', 'Capital Total', 'Interés Total',
    'Total Pagado', 'Saldo Total Pendiente', 'Estado']]).setFontWeight('bold').setBackground('#674ea7').setFontColor('#fff');
  sh.setFrozenRows(1);
  sh.getRange('A2').setFormula(`=IFERROR(UNIQUE(FILTER('${B}'!B2:B,'${B}'!B2:B<>"")),"")`);
  const N = CFG.MAX_ROWS, f = [];
  for (let r = 2; r <= N + 1; r++) {
    f.push([
      `=IF($A${r}="","",COUNTIF('${B}'!$B:$B,$A${r}))`,
      `=IF($A${r}="","",SUMIF('${B}'!$B:$B,$A${r},'${B}'!$F:$F))`,
      `=IF($A${r}="","",SUMIF('${B}'!$B:$B,$A${r},'${B}'!$K:$K))`,
      `=IF($A${r}="","",SUMIF('${B}'!$B:$B,$A${r},'${B}'!$M:$M))`,
      `=IF($A${r}="","",SUMIF('${B}'!$B:$B,$A${r},'${B}'!$N:$N))`,
      `=IF($A${r}="","",IF($F${r}<=0.009,"${ST.CLEARED}",IF(COUNTIFS('${B}'!$B:$B,$A${r},'${B}'!$O:$O,"${ST.OVERDUE}")>0,"${ST.OVERDUE}","${ST.ACTIVE}")))`,
    ]);
  }
  sh.getRange(2, 2, N, 6).setFormulas(f);
  sh.getRange(2, 3, N, 4).setNumberFormat(CFG.CURRENCY_FMT);
  [200, 80, 130, 130, 130, 150, 100].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  const range = sh.getRange(2, 7, N, 1);
  sh.setConditionalFormatRules([cc_(range, ST.CLEARED, '#b6d7a8'), cc_(range, ST.OVERDUE, '#ea9999'), cc_(range, ST.ACTIVE, '#fff2cc')]);
}

/** Estudio de Contratos — casillas para vista previa y envío del contrato. */
function setupAgreement_(ss) {
  const sh = getOrCreate_(ss, CFG.SHEETS.AGREEMENT); sh.clear();
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearDataValidations();
  const B = CFG.SHEETS.BORROWERS;
  sh.getRange('A1').setValue('ESTUDIO DE CONTRATOS').setFontSize(16).setFontWeight('bold').setFontColor('#1c4587');
  sh.getRange('A2').setValue('Elija un ID de Préstamo, luego tilde ① Vista previa o ② Enviar contrato.').setFontColor('#666');
  sh.getRange('A4').setValue('ID Préstamo:').setFontWeight('bold');
  sh.getRange('B4').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInRange(ss.getSheetByName(B).getRange('A2:A' + (CFG.MAX_ROWS + 1)), true).setAllowInvalid(false).build())
    .setBackground('#fff2cc').setFontWeight('bold');
  const rows = [
    ['Prestatario', `=IFERROR(VLOOKUP($B$4,'${B}'!$A:$B,2,FALSE),"")`],
    ['DNI', `=IFERROR(VLOOKUP($B$4,'${B}'!$A:$D,4,FALSE),"")`],
    ['Correo', `=IFERROR(VLOOKUP($B$4,'${B}'!$A:$C,3,FALSE),"")`],
    ['Capital', `=IFERROR(VLOOKUP($B$4,'${B}'!$A:$F,6,FALSE),"")`],
    ['Plazo (meses)', `=IFERROR(VLOOKUP($B$4,'${B}'!$A:$G,7,FALSE),"")`],
    ['Interés', `=IFERROR(VLOOKUP($B$4,'${B}'!$A:$K,11,FALSE),"")`],
    ['Total a Pagar', `=IFERROR(VLOOKUP($B$4,'${B}'!$A:$L,12,FALSE),"")`],
    ['Fecha de Vencimiento', `=IFERROR(VLOOKUP($B$4,'${B}'!$A:$J,10,FALSE),"")`],
  ];
  sh.getRange(6, 1, rows.length, 1).setValues(rows.map(r => [r[0]])).setFontWeight('bold');
  sh.getRange(6, 2, rows.length, 1).setFormulas(rows.map(r => [r[1]]));
  [9, 11, 12].forEach(r => sh.getRange(r, 2).setNumberFormat(CFG.CURRENCY_FMT));
  sh.getRange(13, 2).setNumberFormat('yyyy-mm-dd');
  const cb = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  sh.getRange('A15').setValue('①  Vista previa del contrato  👁').setFontWeight('bold').setBackground('#cfe2f3');
  sh.getRange('B15').setDataValidation(cb);
  sh.getRange('A16').setValue('②  Enviar contrato por correo  ✉').setFontWeight('bold').setBackground('#d9ead3');
  sh.getRange('B16').setDataValidation(cb);
  sh.getRange('A18').setValue('Último PDF:').setFontWeight('bold');
  sh.getRange('A19').setValue('Estado:').setFontWeight('bold');
  sh.setColumnWidth(1, 200); sh.setColumnWidth(2, 420);
  sh.getRange('A1:B2').setWrap(true);
}

function setupStatements_(ss) {
  const sh = getOrCreate_(ss, CFG.SHEETS.STATEMENTS); sh.clear();
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearDataValidations();
  const B = CFG.SHEETS.BORROWERS, P = CFG.SHEETS.PAYMENTS;
  sh.getRange('A1').setValue('ESTUDIO DE ESTADOS DE CUENTA').setFontSize(16).setFontWeight('bold').setFontColor('#674ea7');
  sh.getRange('A2').setValue('Elija un ID de Préstamo; tilde ① para generar el PDF o ② para enviarlo por correo.').setFontColor('#666');
  sh.getRange('A4').setValue('ID Préstamo:').setFontWeight('bold');
  sh.getRange('B4').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInRange(ss.getSheetByName(B).getRange('A2:A' + (CFG.MAX_ROWS + 1)), true).setAllowInvalid(false).build())
    .setBackground('#fff2cc').setFontWeight('bold');
  const rows = [
    ['Prestatario', `=IFERROR(VLOOKUP($B$4,'${B}'!$A:$B,2,FALSE),"")`],       // 6
    ['DNI', `=IFERROR(VLOOKUP($B$4,'${B}'!$A:$D,4,FALSE),"")`],               // 7
    ['Correo', `=IFERROR(VLOOKUP($B$4,'${B}'!$A:$C,3,FALSE),"")`],            // 8
    ['Capital', `=IFERROR(VLOOKUP($B$4,'${B}'!$A:$F,6,FALSE),"")`],           // 9
    ['Total a Pagar', `=IFERROR(VLOOKUP($B$4,'${B}'!$A:$L,12,FALSE),"")`],    // 10
    ['Total Pagado', `=IFERROR(VLOOKUP($B$4,'${B}'!$A:$M,13,FALSE),"")`],     // 11
    ['Días de Atraso', `=IFERROR(IF(VLOOKUP($B$4,'${B}'!$A:$J,10,FALSE)="","",MAX(0,TODAY()-VLOOKUP($B$4,'${B}'!$A:$J,10,FALSE))),"")`], // 12
    ['Recargo por Mora (acum.)', `=IFERROR(IF(OR($B$12="",$B$12<=0),0,$B$12*$B$10*IFERROR(VLOOKUP("Recargo por Mora diario (%)",'${CFG.SHEETS.SETTINGS}'!$A:$B,2,FALSE)/100,0.05)),"")`], // 13
    ['Saldo Pendiente', `=IFERROR(VLOOKUP($B$4,'${B}'!$A:$N,14,FALSE),"")`],  // 14
    ['Fecha de Vencimiento', `=IFERROR(VLOOKUP($B$4,'${B}'!$A:$J,10,FALSE),"")`], // 15
    ['ESTADO', `=IFERROR(VLOOKUP($B$4,'${B}'!$A:$O,15,FALSE),"")`],           // 16
  ];
  sh.getRange(6, 1, rows.length, 1).setValues(rows.map(r => [r[0]])).setFontWeight('bold');
  sh.getRange(6, 2, rows.length, 1).setFormulas(rows.map(r => [r[1]]));
  [9, 10, 11, 13, 14].forEach(r => sh.getRange(r, 2).setNumberFormat(CFG.CURRENCY_FMT));
  sh.getRange(12, 2).setNumberFormat('0');
  sh.getRange(15, 2).setNumberFormat('yyyy-mm-dd');
  sh.getRange(13, 2).setFontColor('#900');
  const st = sh.getRange(16, 2); st.setFontWeight('bold').setFontSize(12).setHorizontalAlignment('center');
  sh.setConditionalFormatRules([cc_(st, ST.PAID, '#b6d7a8'), cc_(st, ST.OVERDUE, '#ea9999'), cc_(st, ST.ACTIVE, '#fff2cc')]);
  const cbSt = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  sh.getRange('A18').setValue('①  Generar PDF del estado  👁').setFontWeight('bold').setBackground('#cfe2f3');
  sh.getRange('B18').setDataValidation(cbSt);
  sh.getRange('A19').setValue('②  Enviar estado por correo  ✉').setFontWeight('bold').setBackground('#d9ead3');
  sh.getRange('B19').setDataValidation(cbSt);
  sh.getRange('A20').setValue('Último PDF:').setFontWeight('bold');
  sh.getRange('A21').setValue('Estado:').setFontWeight('bold');
  sh.getRange('A23:C23').setValues([['Fecha de Pago', 'Monto Pagado', 'Saldo Posterior']])
    .setFontWeight('bold').setBackground('#674ea7').setFontColor('#fff');
  sh.getRange('A24').setFormula(`=IFERROR(FILTER({'${P}'!$F$2:$F,'${P}'!$G$2:$G,'${P}'!$H$2:$H},'${P}'!$B$2:$B=$B$4),"— sin pagos registrados —")`);
  sh.getRange('A24:A' + (CFG.MAX_ROWS + 23)).setNumberFormat('yyyy-mm-dd');
  sh.getRange('B24:C' + (CFG.MAX_ROWS + 23)).setNumberFormat(CFG.CURRENCY_FMT);
  [170, 200, 160].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.setFrozenRows(23);
}

/* ---- Acciones por casilla en los estudios ---- */
function contractStudioPreview_(sh) {
  const status = m => sh.getRange('B19').setValue(m + ' (' + fmtDate_(new Date()) + ')');
  const loanId = String(sh.getRange('B4').getValue()).trim();
  if (!loanId) { status('Elija un ID en B4'); return; }
  try {
    const loan = findLoanById_(loanId);
    if (!loan) { status('No se encontró el préstamo'); return; }
    const file = makeAgreementFile_(loan);
    sh.getRange('B18').setFormula('=HYPERLINK("' + file.getUrl() + '","👁 Abrir PDF")');
    status('Vista previa lista — abra el PDF, luego tilde ②');
  } catch (err) { status('Error: ' + err.message); logError_('contractStudioPreview_', err); }
}
function contractStudioSend_(sh) {
  const status = m => sh.getRange('B19').setValue(m + ' (' + fmtDate_(new Date()) + ')');
  const loanId = String(sh.getRange('B4').getValue()).trim();
  if (!loanId) { status('Elija un ID en B4'); return; }
  try {
    const loan = findLoanById_(loanId);
    if (!loan) { status('No se encontró el préstamo'); return; }
    if (!loan.email) { status('El prestatario no tiene correo'); return; }
    const file = makeAgreementFile_(loan);
    emailAgreement_(loan, file); markAgreementSent_(loan.loanId, file);
    sh.getRange('B18').setFormula('=HYPERLINK("' + file.getUrl() + '","Ver contrato")');
    status('Contrato enviado a ' + loan.email);
  } catch (err) { status('Error: ' + err.message); logError_('contractStudioSend_', err); }
}
function statementStudioPreview_(sh) {
  const status = m => sh.getRange('B21').setValue(m + ' (' + fmtDate_(new Date()) + ')');
  const loanId = String(sh.getRange('B4').getValue()).trim();
  if (!loanId) { status('Elija un ID en B4'); return; }
  try {
    const loan = findLoanById_(loanId);
    if (!loan) { status('No se encontró el préstamo'); return; }
    const file = makeStatementFile_(loan);
    sh.getRange('B20').setFormula('=HYPERLINK("' + file.getUrl() + '","👁 Abrir PDF")');
    status('PDF del estado generado — ábralo, luego tilde ② para enviarlo');
  } catch (err) { status('Error: ' + err.message); logError_('statementStudioPreview_', err); }
}
function statementStudioSend_(sh) {
  const status = m => sh.getRange('B21').setValue(m + ' (' + fmtDate_(new Date()) + ')');
  const loanId = String(sh.getRange('B4').getValue()).trim();
  if (!loanId) { status('Elija un ID en B4'); return; }
  try {
    const loan = findLoanById_(loanId);
    if (!loan) { status('No se encontró el préstamo'); return; }
    if (!loan.email) { status('El prestatario no tiene correo'); return; }
    const file = makeStatementFile_(loan);
    const msg = sendStatementEmail_(loan, file);
    sh.getRange('B20').setFormula('=HYPERLINK("' + file.getUrl() + '","Ver estado de cuenta")');
    status(msg);
  } catch (err) { status('Error: ' + err.message); logError_('statementStudioSend_', err); }
}

/* ===================== ESTUDIO DE RECORDATORIOS (lista) ===================== */
const REM_HEADERS = ['ID Préstamo', 'Prestatario', 'Fecha de Vencimiento', 'Días (− = atraso)',
  'Estado', 'Total a Pagar', 'Saldo Pendiente', 'Recargo por Mora diario', 'Correo',
  'Generar PDF 👁', 'Enviar ✉', 'Último PDF'];
const REM_GEN_COL = 10, REM_SEND_COL = 11, REM_LINK_COL = 12;
const REM_HEADER_ROW = 3, REM_DATA_START = 4; // fila 1 título, fila 2 control, fila 3 encabezados

/** Construye la estructura (título, casilla de refresco y encabezados) y llena la lista. */
function setupReminders_(ss) {
  const sh = getOrCreate_(ss, CFG.SHEETS.REMINDER); sh.clear();
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearDataValidations();
  sh.getRange('A1').setValue('ESTUDIO DE RECORDATORIOS DE PAGO').setFontSize(16).setFontWeight('bold').setFontColor('#b45f06');
  sh.getRange('C1').setValue('Préstamos vencidos (rojo) o próximos a vencer (amarillo). Tilde 👁 para generar el PDF o ✉ para enviarlo por correo.').setFontColor('#666').setWrap(true);
  sh.getRange('A2').setValue('🔄 Actualizar lista →').setFontWeight('bold').setBackground('#cfe2f3');
  sh.getRange('B2').setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
  sh.getRange('A1:A2').setWrap(true);
  sh.getRange(REM_HEADER_ROW, 1, 1, REM_HEADERS.length).setValues([REM_HEADERS])
    .setFontWeight('bold').setBackground('#b45f06').setFontColor('#fff').setWrap(true);
  sh.setFrozenRows(REM_HEADER_ROW);
  [110, 180, 130, 120, 110, 120, 130, 150, 210, 110, 90, 130].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.getRange(REM_HEADER_ROW, 1).setNote('Se actualiza con "⑤ Actualizar", con la casilla 🔄 y automáticamente cada día. Incluye préstamos vencidos o dentro de los "Días de aviso antes del vencimiento".');
  rebuildRemindersSheet_(ss);
  return sh;
}

/** Reconstruye la lista: un renglón por préstamo vencido o próximo a vencer, con colores. */
function rebuildRemindersSheet_(ss) {
  ss = ss || getSS_();
  let sh = ss.getSheetByName(CFG.SHEETS.REMINDER);
  if (!sh) { setupReminders_(ss); return; } // setupReminders_ ya llama a rebuild
  const bs = ss.getSheetByName(CFG.SHEETS.BORROWERS), last = bs ? bs.getLastRow() : 0;
  // limpia datos, casillas y colores previos (conserva encabezados)
  const clearH = Math.max(sh.getMaxRows() - REM_HEADER_ROW, 1);
  sh.getRange(REM_DATA_START, 1, clearH, REM_HEADERS.length).clearContent().clearDataValidations().setBackground(null);
  if (!bs || last < 2) return;
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const noticeDays = Number(getSetting_('Días de aviso antes del vencimiento')) || 3;
  const feePct = lateFeeRate_(), rows = [], overdueFlags = [];
  for (let row = 2; row <= last; row++) {
    const loan = readLoan_(bs, row);
    if (!loan.loanId || !(loan.dueDate instanceof Date)) continue;
    const dd = new Date(loan.dueDate); dd.setHours(0, 0, 0, 0);
    const daysUntil = Math.round((dd - today) / 86400000);
    if (daysUntil > noticeDays) continue; // ni vencido ni próximo
    const out = computeOutstanding_(loan.principal, termRate_(loan.term), loan.loanDate, loanPayments_(loan.loanId), today, new Date(9999, 0, 1));
    if (out <= 0) continue; // saldado
    const overdue = daysUntil < 0;
    const feeDay = round2_(loan.totalDue * feePct);
    rows.push([loan.loanId, loan.name, loan.dueDate, daysUntil, overdue ? ST.OVERDUE : 'PRÓXIMO',
      loan.totalDue, out, feeDay, loan.email, false, false, '']);
    overdueFlags.push(overdue);
  }
  // ordena: más atrasados primero (daysUntil ascendente)
  const order = rows.map((r, i) => i).sort((a, b) => rows[a][3] - rows[b][3]);
  const sorted = order.map(i => rows[i]), sortedFlags = order.map(i => overdueFlags[i]);
  if (!sorted.length) return;
  const n = sorted.length;
  sh.getRange(REM_DATA_START, 1, n, REM_HEADERS.length).setValues(sorted);
  sh.getRange(REM_DATA_START, 3, n, 1).setNumberFormat('yyyy-mm-dd');
  sh.getRange(REM_DATA_START, 4, n, 1).setNumberFormat('0');
  sh.getRange(REM_DATA_START, 6, n, 3).setNumberFormat(CFG.CURRENCY_FMT);
  sh.getRange(REM_DATA_START, REM_GEN_COL, n, 1).insertCheckboxes();
  sh.getRange(REM_DATA_START, REM_SEND_COL, n, 1).insertCheckboxes();
  // colores por renglón: vencido = rojo, próximo = amarillo (columnas de datos 1..9)
  for (let i = 0; i < n; i++) {
    sh.getRange(REM_DATA_START + i, 1, 1, 9).setBackground(sortedFlags[i] ? '#f4cccc' : '#fff2cc');
  }
}

/* ---- Acciones por renglón en el Estudio de Recordatorios ---- */
function reminderRowGenerate_(sh, row) { // 👁 genera el PDF y deja el enlace (sin correo)
  try {
    const loan = findLoanById_(String(sh.getRange(row, 1).getValue()).trim());
    if (!loan) { getSS_().toast('No se encontró el préstamo', '⚠', 5); return; }
    const file = makeReminderFile_(loan);
    sh.getRange(row, REM_LINK_COL).setFormula('=HYPERLINK("' + file.getUrl() + '","👁 Abrir PDF")');
    getSS_().toast('PDF generado para ' + loan.name, 'Recordatorio', 4);
  } catch (err) { logError_('reminderRowGenerate_', err); getSS_().toast(err.message || String(err), '⚠ No se pudo generar', 8); }
}
function reminderRowSend_(sh, row) { // ✉ genera el PDF y lo envía por correo
  try {
    const loan = findLoanById_(String(sh.getRange(row, 1).getValue()).trim());
    if (!loan) { getSS_().toast('No se encontró el préstamo', '⚠', 5); return; }
    if (!loan.email) { getSS_().toast('El prestatario no tiene correo', '⚠', 6); return; }
    const file = makeReminderFile_(loan); emailReminder_(loan, file);
    sh.getRange(row, REM_LINK_COL).setFormula('=HYPERLINK("' + file.getUrl() + '","👁 Abrir PDF")');
    getSS_().toast('Recordatorio enviado a ' + loan.email, 'Recordatorio', 5);
  } catch (err) { logError_('reminderRowSend_', err); getSS_().toast(err.message || String(err), '⚠ No se pudo enviar', 8); }
}

/** Crea únicamente la hoja "Estudio de Recordatorios" sin tocar Prestatarios ni Pagos. */
function createRemindersSheet_() {
  const ss = getSS_();
  setupReminders_(ss);
  ensureTriggers_();
  ss.setActiveSheet(ss.getSheetByName(CFG.SHEETS.REMINDER));
  ss.toast('Hoja "Estudio de Recordatorios" creada. Sus datos de Prestatarios y Pagos no fueron modificados.', 'Listo', 6);
}

/**
 * Instala/actualiza SOLO la hoja "Nuevos Prestatarios" (encabezados, casillas,
 * lista DNI/CUIL, formatos y columnas BCRA) SIN borrar datos de otras hojas ni
 * de las filas existentes. Seguro para correr sobre una hoja con datos reales.
 */
function setupNewOnly() {
  setupNew_(getSS_());
  try { getSS_().toast('Columnas BCRA agregadas a "Nuevos Prestatarios"', 'Listo', 5); } catch (e) {}
}

function setupNew_(ss) {
  let sh = ss.getSheetByName(CFG.SHEETS.NEW) || ss.getSheetByName('New Borrowers');
  if (!sh) sh = ss.insertSheet(CFG.SHEETS.NEW);
  else if (sh.getName() !== CFG.SHEETS.NEW && !ss.getSheetByName(CFG.SHEETS.NEW)) sh.setName(CFG.SHEETS.NEW);
  sh.getRange(1, 1, 1, NB.length).setValues([NB]).setFontWeight('bold').setBackground('#b45f06').setFontColor('#fff').setWrap(true);
  sh.setFrozenRows(1);
  const cb = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  sh.getRange(2, 11, CFG.MAX_ROWS, 1).setDataValidation(cb); // Verificado?
  sh.getRange(2, 12, CFG.MAX_ROWS, 1).setDataValidation(cb); // Rechazar?
  sh.getRange(2, 14, CFG.MAX_ROWS, 1).setDataValidation(cb); // Verificar BCRA?
  const paramDv = SpreadsheetApp.newDataValidation().requireValueInList(['DNI', 'CUIL'], true).build();
  const paramRange = sh.getRange(2, 15, CFG.MAX_ROWS, 1); // Parámetro BCRA
  paramRange.setDataValidation(paramDv);
  paramRange.setValue('DNI');
  sh.getRange(2, 1, CFG.MAX_ROWS, 1).setNumberFormat('yyyy-mm-dd hh:mm');
  sh.getRange(2, 6, CFG.MAX_ROWS, 1).setNumberFormat(CFG.CURRENCY_FMT);
  [150, 170, 200, 130, 120, 130, 100, 200, 190, 190, 90, 90, 220,
    110, 120, 120, 170, 110, 90, 130, 100, 170, 460].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.getRange(2, 23, CFG.MAX_ROWS, 1).setWrap(true); // Resumen BCRA
  sh.getRange('K1').setNote('Tilde "Verificado?" para aprobar (pasa a Prestatarios) o "Rechazar?" para archivar en Rechazados.');
  sh.getRange('N1').setNote('Tilde "Verificar BCRA?" para consultar la Central de Deudores del BCRA. En "Parámetro BCRA" elegí DNI o CUIL: con CUIL se usa el valor de la columna CUIL (si está vacío, se deriva del DNI).');
}

function setupRejected_(ss) {
  const sh = getOrCreate_(ss, CFG.SHEETS.REJECTED);
  if (sh.getLastRow() === 0)
    sh.getRange(1, 1, 1, 6).setValues([['Fecha', 'Nombre completo', 'Correo', 'DNI', 'Teléfono', 'Motivo']])
      .setFontWeight('bold').setBackground('#990000').setFontColor('#fff');
  sh.setFrozenRows(1);
}

function setupErrors_(ss) {
  const sh = getOrCreate_(ss, CFG.SHEETS.ERRORS);
  if (sh.getLastRow() === 0)
    sh.getRange(1, 1, 1, 3).setValues([['Fecha', 'Función', 'Error']]).setFontWeight('bold').setBackground('#7f6000').setFontColor('#fff');
  sh.setFrozenRows(1);
}

function setupPanel_(ss) {
  const sh = getOrCreate_(ss, CFG.SHEETS.PANEL); sh.clear();
  const B = CFG.SHEETS.BORROWERS, fondo = fondoFormula_();
  sh.getRange('A1').setValue('PANEL DEL PRESTAMISTA').setFontSize(18).setFontWeight('bold').setFontColor('#1c4587');
  const kpis = [
    ['Préstamos activos', `=COUNTIF('${B}'!$O:$O,"${ST.ACTIVE}")`],           // 3  int
    ['Préstamos vencidos', `=COUNTIF('${B}'!$O:$O,"${ST.OVERDUE}")`],         // 4  int
    ['Préstamos pagados', `=COUNTIF('${B}'!$O:$O,"${ST.PAID}")`],             // 5  int
    ['Solicitudes pendientes', `=COUNTA('${CFG.SHEETS.NEW}'!$B$2:$B)`],       // 6  int
    ['Contratos sin firmar', `=COUNTIF('${B}'!$S:$S,"${SIGN.PENDING}")`],     // 7  int
    ['Fondo total para prestar', `=${fondo}`],                                // 8  money
    ['Capital prestado', `=SUM('${B}'!$F:$F)`],                              // 9  money
    ['Total cobrado', `=SUM('${B}'!$M:$M)`],                                 // 10 money
    ['Efectivo disponible para prestar', `=${fondo}-SUM('${B}'!$F:$F)+SUM('${B}'!$M:$M)`], // 11 money
    ['Interés contratado', `=SUM('${B}'!$K:$K)`],                            // 12 money
    ['Saldo pendiente total', `=SUM('${B}'!$N:$N)`],                         // 13 money
  ];
  sh.getRange(3, 1, kpis.length, 1).setValues(kpis.map(k => [k[0]])).setFontWeight('bold');
  sh.getRange(3, 2, kpis.length, 1).setFormulas(kpis.map(k => [k[1]]));
  sh.getRange(3, 2, 5, 1).setNumberFormat('0');                 // filas 3–7 enteros
  sh.getRange(8, 2, 6, 1).setNumberFormat(CFG.CURRENCY_FMT);    // filas 8–13 moneda
  sh.getRange('A7').setFontColor('#990000'); sh.getRange('B7').setFontColor('#990000').setFontWeight('bold');
  sh.getRange('A7').setNote('Préstamos aprobados cuyo contrato aún no fue firmado por el prestatario. No desembolsar hasta la firma.');
  sh.getRange('A11').setFontColor('#38761d'); sh.getRange('B11').setFontColor('#38761d').setFontWeight('bold');
  sh.getRange('A11').setNote('Efectivo disponible = Fondo total − Capital prestado + Total cobrado. Baja al prestar y sube al cobrar.');
  sh.getRange('A15').setValue('Próximos vencimientos (7 días):').setFontWeight('bold');
  sh.getRange('A16').setFormula(
    `=IFERROR(SORT(FILTER({'${B}'!$A2:$A,'${B}'!$B2:$B,'${B}'!$J2:$J,'${B}'!$N2:$N},` +
    `('${B}'!$O2:$O<>"${ST.PAID}")*('${B}'!$J2:$J>=TODAY())*('${B}'!$J2:$J<=TODAY()+7)),3,TRUE),"— sin vencimientos próximos —")`);
  sh.getRange('A15').setNote('Muestra préstamos no pagados que vencen dentro de 7 días.');
  sh.setColumnWidth(1, 240); sh.setColumnWidth(2, 200); sh.setColumnWidth(3, 130); sh.setColumnWidth(4, 130);
}

/* ===================== FONDO (capital disponible) ===================== */
/** Fórmula que lee el "Fondo total para prestar" de Configuración como número. */
function fondoFormula_() {
  return `IFERROR(VALUE(REGEXREPLACE(TO_TEXT(VLOOKUP("Fondo total para prestar",'${CFG.SHEETS.SETTINGS}'!$A:$B,2,FALSE)),"[^0-9\\.\\-]","")),0)`;
}
/** ¿El formulario web acepta nuevas solicitudes? Predeterminado: SÍ. */
function acceptingApplications_() {
  const raw = String(getSetting_('Aceptar solicitudes de préstamo (SÍ/NO)') || '').trim();
  if (raw === '') return true; // por defecto abierto si no está configurado
  return /^(s[íi]|si|true|verdadero|1|x|on|abierto)$/i.test(raw);
}
/**
 * Destinatarios de los avisos de nuevas solicitudes: "Correo del Prestamista"
 * más "Correos de aviso adicionales" (separados por coma, punto y coma o espacio).
 * Devuelve una lista de correos válidos, sin duplicados. Ej: ["a@x.com","b@y.com"].
 */
function noticeRecipients_() {
  const raw = String(getSetting_('Correo del Prestamista') || '') + ',' +
    String(getSetting_('Correos de aviso adicionales (separados por coma)') || '');
  const seen = {}, out = [];
  raw.split(/[,;\s]+/).forEach(function (e) {
    e = String(e).trim();
    const key = e.toLowerCase();
    if (!e || seen[key]) return;
    if (/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(e)) { seen[key] = true; out.push(e); }
  });
  return out;
}
/** Fondo total configurado (número). */
function fondoTotal_() {
  return Number(String(getSetting_('Fondo total para prestar') || '0').replace(/[^0-9.\-]/g, '')) || 0;
}
/**
 * Efectivo disponible para prestar = Fondo total − Capital prestado + Total cobrado.
 * Baja cuando se presta (capital, col F) y sube cuando se cobra (pagos, col M).
 */
function fundStats_() {
  const bs = getSS_().getSheetByName(CFG.SHEETS.BORROWERS), last = bs ? bs.getLastRow() : 0;
  let lent = 0, collected = 0;
  if (last >= 2) {
    const vals = bs.getRange(2, 6, last - 1, 8).getValues(); // F(6)…M(13): F=capital idx0, M=total pagado idx7
    vals.forEach(r => { lent += Number(r[0]) || 0; collected += Number(r[7]) || 0; });
  }
  const total = fondoTotal_();
  return { total: total, lent: round2_(lent), collected: round2_(collected), available: round2_(total - lent + collected) };
}
function fundAvailable_() { return fundStats_().available; }

/* ===================== ESTADÍSTICAS + GRÁFICOS ===================== */
function showStats() {
  const ss = getSS_(); setupStats_(ss);
  ss.setActiveSheet(ss.getSheetByName(CFG.SHEETS.STATS));
}

/** Hoja de estadísticas de la cartera con indicadores y gráficos típicos de una agencia de préstamos. */
function setupStats_(ss) {
  ss = ss || getSS_();
  const sh = getOrCreate_(ss, CFG.SHEETS.STATS); sh.clear();
  sh.getCharts().forEach(c => sh.removeChart(c));
  const B = CFG.SHEETS.BORROWERS, fondo = fondoFormula_();
  const money = CFG.CURRENCY_FMT;
  sh.getRange('A1').setValue('ESTADÍSTICAS DE LA CARTERA').setFontSize(18).setFontWeight('bold').setFontColor('#1c4587');
  sh.getRange('A2').setValue('Se actualiza con "⑤ Actualizar" y automáticamente cada día. Los gráficos se recalculan solos.').setFontColor('#666');

  // ---- Indicadores clave (A3:B…) ----
  // type: m=moneda, i=entero, p=porcentaje
  const kpis = [
    ['Fondo total para prestar', `=${fondo}`, 'm'],
    ['Capital prestado (colocado)', `=SUM('${B}'!$F:$F)`, 'm'],
    ['Total cobrado', `=SUM('${B}'!$M:$M)`, 'm'],
    ['Efectivo disponible para prestar', `=${fondo}-SUM('${B}'!$F:$F)+SUM('${B}'!$M:$M)`, 'm'],
    ['Interés contratado (ganancia proyectada)', `=SUM('${B}'!$K:$K)`, 'm'],
    ['Saldo pendiente por cobrar', `=SUM('${B}'!$N:$N)`, 'm'],
    ['Cartera en mora (saldo vencido)', `=SUMIF('${B}'!$O:$O,"${ST.OVERDUE}",'${B}'!$N:$N)`, 'm'],
    ['Ticket promedio', `=IFERROR(AVERAGEIF('${B}'!$A:$A,"<>",'${B}'!$F:$F),0)`, 'm'],
    ['Préstamos activos', `=COUNTIF('${B}'!$O:$O,"${ST.ACTIVE}")`, 'i'],
    ['Préstamos vencidos', `=COUNTIF('${B}'!$O:$O,"${ST.OVERDUE}")`, 'i'],
    ['Préstamos pagados', `=COUNTIF('${B}'!$O:$O,"${ST.PAID}")`, 'i'],
    ['Tasa de morosidad (préstamos vencidos)', `=IFERROR(COUNTIF('${B}'!$O:$O,"${ST.OVERDUE}")/(COUNTIF('${B}'!$O:$O,"${ST.ACTIVE}")+COUNTIF('${B}'!$O:$O,"${ST.OVERDUE}")),0)`, 'p'],
    ['Tasa de recuperación (cobrado / a pagar)', `=IFERROR(SUM('${B}'!$M:$M)/SUM('${B}'!$L:$L),0)`, 'p'],
    ['Rendimiento sobre capital (ROI)', `=IFERROR(SUM('${B}'!$K:$K)/SUM('${B}'!$F:$F),0)`, 'p'],
    ['Utilización del fondo', `=IFERROR((SUM('${B}'!$F:$F)-SUM('${B}'!$M:$M))/${fondo},0)`, 'p'],
  ];
  const r0 = 4;
  sh.getRange(r0, 1, kpis.length, 1).setValues(kpis.map(k => [k[0]])).setFontWeight('bold');
  sh.getRange(r0, 2, kpis.length, 1).setFormulas(kpis.map(k => [k[1]]));
  kpis.forEach((k, i) => {
    const cell = sh.getRange(r0 + i, 2);
    cell.setNumberFormat(k[2] === 'm' ? money : k[2] === 'p' ? '0.0%' : '0');
  });
  // resaltar el efectivo disponible
  sh.getRange(r0 + 3, 1, 1, 2).setFontColor('#38761d').setFontWeight('bold');
  sh.getRange(r0 + 3, 1).setNote('Baja al prestar (capital) y sube al cobrar (pagos).');

  // ---- Tablas de datos para los gráficos (columnas D:E) ----
  // 1) Estado de la cartera (torta)
  sh.getRange('D3').setValue('Estado de la cartera').setFontWeight('bold').setFontColor('#1c4587');
  sh.getRange('D4:E7').setValues([
    ['Estado', 'Cantidad'],
    ['Activos', 0], ['Vencidos', 0], ['Pagados', 0],
  ]);
  sh.getRange('E5').setFormula(`=COUNTIF('${B}'!$O:$O,"${ST.ACTIVE}")`);
  sh.getRange('E6').setFormula(`=COUNTIF('${B}'!$O:$O,"${ST.OVERDUE}")`);
  sh.getRange('E7').setFormula(`=COUNTIF('${B}'!$O:$O,"${ST.PAID}")`);

  // 2) Flujo de dinero (barras)
  sh.getRange('D10').setValue('Flujo de dinero').setFontWeight('bold').setFontColor('#1c4587');
  sh.getRange('D11:E15').setValues([
    ['Concepto', 'Monto'],
    ['Capital prestado', 0], ['Interés contratado', 0], ['Total cobrado', 0], ['Saldo pendiente', 0],
  ]);
  sh.getRange('E12').setFormula(`=SUM('${B}'!$F:$F)`);
  sh.getRange('E13').setFormula(`=SUM('${B}'!$K:$K)`);
  sh.getRange('E14').setFormula(`=SUM('${B}'!$M:$M)`);
  sh.getRange('E15').setFormula(`=SUM('${B}'!$N:$N)`);
  sh.getRange('E12:E15').setNumberFormat(money);

  // 3) Liquidez del fondo (torta)
  sh.getRange('D18').setValue('Liquidez del fondo').setFontWeight('bold').setFontColor('#1c4587');
  sh.getRange('D19:E21').setValues([
    ['Composición', 'Monto'],
    ['Efectivo disponible', 0], ['Por cobrar (saldo)', 0],
  ]);
  sh.getRange('E20').setFormula(`=MAX(0,${fondo}-SUM('${B}'!$F:$F)+SUM('${B}'!$M:$M))`);
  sh.getRange('E21').setFormula(`=SUM('${B}'!$N:$N)`);
  sh.getRange('E20:E21').setNumberFormat(money);

  sh.setColumnWidth(1, 280); sh.setColumnWidth(2, 150);
  sh.setColumnWidth(4, 170); sh.setColumnWidth(5, 120);

  // ---- Gráficos ----
  const pieStatus = sh.newChart().setChartType(Charts.ChartType.PIE)
    .addRange(sh.getRange('D4:E7'))
    .setPosition(4, 7, 0, 0)
    .setOption('title', 'Distribución de préstamos por estado')
    .setOption('width', 460).setOption('height', 260)
    .setOption('pieHole', 0.4)
    .setOption('colors', ['#93c47d', '#e06666', '#f6b26b'])
    .build();
  sh.insertChart(pieStatus);

  const colFlow = sh.newChart().setChartType(Charts.ChartType.COLUMN)
    .addRange(sh.getRange('D11:E15'))
    .setPosition(19, 7, 0, 0)
    .setOption('title', 'Flujo de dinero de la cartera')
    .setOption('width', 460).setOption('height', 260)
    .setOption('legend', 'none')
    .setOption('colors', ['#1c4587'])
    .build();
  sh.insertChart(colFlow);

  const pieFund = sh.newChart().setChartType(Charts.ChartType.PIE)
    .addRange(sh.getRange('D19:E21'))
    .setPosition(34, 7, 0, 0)
    .setOption('title', 'Liquidez del fondo: disponible vs. por cobrar')
    .setOption('width', 460).setOption('height', 260)
    .setOption('colors', ['#38761d', '#e8a33d'])
    .build();
  sh.insertChart(pieFund);

  return sh;
}

const LATE_HEADERS = ['ID Préstamo', 'Prestatario', 'DNI', 'Correo', 'Teléfono', 'Fecha de Vencimiento',
  'Días de Atraso', 'Total a Pagar', 'Recargo por Mora (acum.)', 'Total Pagado', 'Saldo Pendiente',
  'Enviar Aviso ✉', 'Último Aviso Enviado'];
// Columnas (1-based) de control en "Pagos Atrasados".
const LATE_SEND_COL = 12, LATE_SENT_COL = 13;
function setupLate_(ss) {
  const sh = getOrCreate_(ss, CFG.SHEETS.LATE); sh.clear();
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearDataValidations();
  sh.getRange(1, 1, 1, LATE_HEADERS.length).setValues([LATE_HEADERS])
    .setFontWeight('bold').setBackground('#990000').setFontColor('#fff').setWrap(true);
  sh.setFrozenRows(1);
  [110, 170, 120, 200, 120, 130, 100, 120, 160, 120, 130, 110, 150].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.getRange('A1').setNote('Se actualiza con "⑤ Actualizar" y automáticamente cada día. Lista los préstamos vencidos con saldo.');
  sh.getRange('L1').setNote('Tilde la casilla para enviar un aviso de mora por correo (en español). La casilla se destilda sola y la fecha de envío aparece en "Último Aviso Enviado".');
  return sh;
}

/** Reconstruye "Pagos Atrasados": un renglón por préstamo vencido con saldo, con recargo acumulado. */
function rebuildLateSheet_(ss) {
  ss = ss || getSS_();
  const sh = ss.getSheetByName(CFG.SHEETS.LATE) || setupLate_(ss) || ss.getSheetByName(CFG.SHEETS.LATE);
  const bs = ss.getSheetByName(CFG.SHEETS.BORROWERS), last = bs.getLastRow();
  // limpia datos y casillas previas (conserva encabezados)
  const clearH = Math.max(sh.getMaxRows() - 1, 1);
  sh.getRange(2, 1, clearH, LATE_HEADERS.length).clearContent();
  sh.getRange(2, LATE_SEND_COL, clearH, 1).clearDataValidations();
  if (last < 2) return;
  const today = new Date(), feePct = lateFeeRate_(), rows = [];
  for (let row = 2; row <= last; row++) {
    const loan = readLoan_(bs, row);
    if (!loan.loanId || !(loan.dueDate instanceof Date)) continue;
    const days = daysLate_(loan.dueDate, today);
    if (days <= 0) continue;
    const rate = termRate_(loan.term);
    const pays = loanPayments_(loan.loanId);
    const out = computeOutstanding_(loan.principal, rate, loan.loanDate, pays, today, new Date(9999, 0, 1));
    if (out <= 0) continue; // vencido pero saldado
    const totalPaid = pays.reduce((s, p) => s + p.amount, 0);
    const feeAccum = round2_(loan.totalDue * feePct * days);
    const lastNotice = bs.getRange(row, 18).getValue(); // "Último Aviso" del prestatario
    rows.push([loan.loanId, loan.name, loan.dni, loan.email, loan.phone, loan.dueDate, days,
      loan.totalDue, feeAccum, totalPaid, out, false, (lastNotice instanceof Date) ? lastNotice : '']);
  }
  rows.sort((a, b) => b[6] - a[6]); // más atrasados primero
  if (rows.length) {
    sh.getRange(2, 1, rows.length, LATE_HEADERS.length).setValues(rows);
    sh.getRange(2, 6, rows.length, 1).setNumberFormat('yyyy-mm-dd');
    sh.getRange(2, 8, rows.length, 1).setNumberFormat(CFG.CURRENCY_FMT);
    sh.getRange(2, 9, rows.length, 3).setNumberFormat(CFG.CURRENCY_FMT);
    sh.getRange(2, 7, rows.length, 1).setNumberFormat('0');
    sh.getRange(2, LATE_SEND_COL, rows.length, 1).insertCheckboxes();
    sh.getRange(2, LATE_SENT_COL, rows.length, 1).setNumberFormat('yyyy-mm-dd');
  }
}

function setupHelp_(ss) {
  const sh = getOrCreate_(ss, CFG.SHEETS.HELP); sh.clear();
  const txt = [
    ['INSTRUCCIONES — GESTOR DE PRÉSTAMOS'],
    [''],
    ['1) Configuración: complete la hoja "Configuración" (nombre, contacto, jurisdicción).'],
    ['2) Formulario web: menú ⑥ para publicar y obtener el enlace. Compártalo con solicitantes.'],
    ['   (Si cambia el código, vuelva a implementar: Deploy ▸ Administrar implementaciones ▸ Nueva versión.)'],
    ['3) Las solicitudes llegan a "Nuevos Prestatarios" con fotos de DNI y CUIL.'],
    ['4) Verificar: tilde "Verificado?" para aprobar (pasa a "Prestatarios", crea carpeta y envía contrato).'],
    ['   Tilde "Rechazar?" para archivar en "Rechazados" (envía correo de rechazo).'],
    ['5) Pagos: use el Panel del prestamista (menú ▸ Abrir panel) para registrar pagos y enviar recibos.'],
    ['6) Intereses: 15 días = 25%, 1 mes = 50%, 2 meses = 100%. Tras el vencimiento: recargo por mora del 5% diario sobre el total a pagar.'],
    ['   Los préstamos vencidos con saldo aparecen en la hoja "Pagos Atrasados".'],
    ['   En "Pagos Atrasados", tilde "Enviar Aviso ✉" para mandar un aviso de mora por correo; la fecha queda en "Último Aviso Enviado".'],
    ['7) Estados: ACTIVO, VENCIDO, PAGADO. El Panel muestra indicadores y próximos vencimientos.'],
    ['   Saldados: use "✔ Mover préstamos saldados" (menú o Panel) para archivar los préstamos pagados en la hoja "Saldados".'],
    ['8) Recordatorios: se envían automáticamente antes del vencimiento y en mora.'],
    ['9) Fondo: complete "Fondo total para prestar" en "Configuración". El efectivo disponible baja al aprobar un préstamo (capital) y sube al registrar un pago. Se muestra en el Panel y en "Estadísticas".'],
    ['10) Cada nueva solicitud del formulario también avisa por correo al "Correo del Prestamista" con los datos y el efectivo disponible.'],
    ['11) Estadísticas: menú ④ abre la hoja "Estadísticas" con indicadores (morosidad, recuperación, ROI, utilización) y gráficos de la cartera.'],
    ['12) Abrir/cerrar solicitudes: en "Configuración", ponga "Aceptar solicitudes de préstamo (SÍ/NO)" en SÍ para recibir solicitudes o NO para mostrar "no hay préstamos disponibles" en el formulario.'],
    ['13) Errores: revise la hoja "Errores" si algo no funciona.'],
  ];
  sh.getRange(1, 1, txt.length, 1).setValues(txt);
  sh.getRange('A1').setFontSize(16).setFontWeight('bold').setFontColor('#1c4587');
  sh.setColumnWidth(1, 760); sh.getRange(1, 1, txt.length, 1).setWrap(true);
}

/* ===================== MOTOR DE SALDO ===================== */
/**
 * Saldo con recargo por mora SIMPLE del 5% por día sobre el Total a Pagar.
 * Recargo diario = feePct × totalContrato. Acumula por cada día de atraso,
 * hasta la fecha "asOf" o hasta que el préstamo quede saldado. Sin duplicación mensual.
 */
function computeOutstanding_(principal, rate, loanDate, payments, asOf, paymentCutoff, feeRate) {
  const DAY = 86400000;
  const cutoff = (paymentCutoff instanceof Date) ? paymentCutoff : asOf;
  const dueDate = rate === 0.25 ? addDays_(loanDate, 15) : addMonths_(loanDate, rate === 0.5 ? 1 : 2);
  const feePct = (typeof feeRate === 'number') ? feeRate : lateFeeRate_();
  const total = round2_(principal * (1 + rate));
  const feePerDay = round2_(total * feePct);
  let out = total, cursor = dueDate.getTime();
  const accrueTo = t => {
    if (out <= 0 || feePct <= 0 || t <= dueDate.getTime()) { cursor = Math.max(cursor, t); return; }
    const from = Math.max(cursor, dueDate.getTime());
    const days = Math.floor((t - from) / DAY);
    if (days > 0) out = round2_(out + feePerDay * days);
    cursor = t;
  };
  const pays = payments.filter(p => new Date(p.date).getTime() <= cutoff.getTime())
    .map(p => ({ t: new Date(p.date).getTime(), amount: p.amount })).sort((a, b) => a.t - b.t);
  for (const p of pays) {
    accrueTo(p.t);
    if (out <= 0) { out = 0; break; }
    out = round2_(out - p.amount);
    if (out <= 0) { out = 0; break; }
  }
  accrueTo(asOf.getTime());
  return Math.max(0, round2_(out));
}
/** Días completos de atraso a hoy (0 si aún no vence). */
function daysLate_(dueDate, asOf) {
  if (!(dueDate instanceof Date)) return 0;
  const DAY = 86400000, a = (asOf || new Date());
  const d0 = new Date(dueDate.getFullYear(), dueDate.getMonth(), dueDate.getDate()).getTime();
  const a0 = new Date(a.getFullYear(), a.getMonth(), a.getDate()).getTime();
  return Math.max(0, Math.floor((a0 - d0) / DAY));
}

function updateLoanOutstanding_(loanId) {
  const bs = getSS_().getSheetByName(CFG.SHEETS.BORROWERS), last = bs.getLastRow();
  if (last < 2) return;
  const ids = bs.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]).trim() === loanId) { writeOutstandingRow_(bs, i + 2); return; }
}
function updateAllOutstanding_() {
  const bs = getSS_().getSheetByName(CFG.SHEETS.BORROWERS), last = bs.getLastRow();
  for (let row = 2; row <= last; row++) writeOutstandingRow_(bs, row);
}
function writeOutstandingRow_(bs, row) {
  const loanId = String(bs.getRange(row, 1).getValue()).trim(), cell = bs.getRange(row, 14);
  if (!loanId) { cell.clearContent(); return; }
  const loan = readLoan_(bs, row);
  if (!loan.principal || !loan.term || !(loan.loanDate instanceof Date)) { cell.clearContent(); return; }
  const rate = termRate_(loan.term);
  if (rate === null) { cell.setValue(''); return; }
  cell.setValue(computeOutstanding_(loan.principal, rate, loan.loanDate, loanPayments_(loanId), new Date(), new Date(9999, 0, 1)));
}

/* ===================== CONTRATO (PDF + CORREO) ===================== */
function makeAgreementFile_(loan) {
  const pdf = Utilities.newBlob(agreementHtml_(loan), 'text/html', 'contrato.html').getAs('application/pdf');
  const folder = borrowerFolder_(borrowerFolderName_(loan.name, loan.dni));
  const f = folder.createFile(pdf); f.setName('Contrato ' + loan.loanId + ' - ' + loan.name + '.pdf');
  return f;
}
/** PDF del contrato YA FIRMADO (incluye la imagen de la firma y el bloque de auditoría). */
function makeSignedAgreementFile_(loan, sig) {
  const pdf = Utilities.newBlob(agreementHtml_(loan, sig), 'text/html', 'contrato-firmado.html').getAs('application/pdf');
  const folder = borrowerFolder_(borrowerFolderName_(loan.name, loan.dni));
  const f = folder.createFile(pdf); f.setName('Contrato Firmado ' + loan.loanId + ' - ' + loan.name + '.pdf');
  return f;
}
/** @param {object=} sig  Firma opcional: { dataUrl, signedAt (Date), method, userAgent }. */
function agreementHtml_(loan, sig) {
  const lender = companyName_(), lPhone = getSetting_('Teléfono del Prestamista'),
    lAddr = getSetting_('Dirección del Prestamista'), juris = getSetting_('Jurisdicción'),
    penalty = getSetting_('Cláusula de Mora'), footer = getSetting_('Pie del Contrato');
  const rate = termRatePct_(loan.term);
  // Firmantes titulares de Impulso Crédito (Prestamista): ambos firman TODOS los contratos.
  const signatories = ownerSignatories_();
  // Firma del prestatario: solo la capturada al firmar (si aún no firmó, va una línea en blanco).
  const borrowerSig = (sig && sig.dataUrl) ? sig.dataUrl : '';
  const sigCell = (img, name, role) =>
    `<td>${img ? `<img src="${img}" alt="firma" style="max-height:64px;max-width:92%;display:block;margin:0 auto 2px">` : '<div style="height:64px"></div>'}` +
    `<div style="border-top:1px solid #333;padding-top:4px">${esc_(name)}<br><span style="font-size:9pt;color:#555">${esc_(role)}</span></div></td>`;
  return `<!DOCTYPE html><html lang="es"><head><meta charset="utf-8"><style>
    body{font-family:Georgia,serif;color:#222;margin:48px;line-height:1.5;font-size:12pt}
    h1{text-align:center;font-size:20pt;border-bottom:2px solid #1c4587;padding-bottom:8px}
    h2{font-size:13pt;color:#1c4587;margin-top:24px;border-bottom:1px solid #ccc;padding-bottom:3px}
    table{width:100%;border-collapse:collapse;margin-top:8px} td{padding:6px 8px;border:1px solid #ccc}
    td.k{background:#f0f4fa;font-weight:bold;width:40%}
    .sign{margin-top:28px;width:100%} .sign td{border:none;text-align:center;padding:6px 12px;width:50%;vertical-align:bottom}
    .pagebreak{page-break-before:always}
    .foot{margin-top:40px;font-size:10pt;color:#666;text-align:center}</style></head><body>
    ${brandHeaderHtml_()}
    <h1>CONTRATO DE PRÉSTAMO</h1>
    <p>Este contrato se celebra el <b>${fmtDate_(new Date())}</b> entre <b>${esc_(lender)}</b> ("Prestamista")
       y <b>${esc_(loan.name)}</b> ("Prestatario").</p>
    <h2>1. Partes</h2>
    <table><tr><td class="k">Prestamista</td><td>${esc_(lender)} ${lPhone ? '· ' + esc_(lPhone) : ''}<br>${esc_(lAddr)}</td></tr>
    <tr><td class="k">Prestatario</td><td>${esc_(loan.name)}<br>DNI: ${esc_(loan.dni)} · ${esc_(loan.phone)}<br>${esc_(loan.email)}</td></tr></table>
    <h2>2. Términos del Préstamo</h2>
    <table>
      <tr><td class="k">ID Préstamo</td><td>${esc_(loan.loanId)}</td></tr>
      <tr><td class="k">Capital</td><td>${fmtMoney_(loan.principal)}</td></tr>
      <tr><td class="k">Plazo</td><td>${termLabel_(loan.term)}</td></tr>
      <tr><td class="k">Tasa de interés (fija por el plazo)</td><td>${rate}</td></tr>
      <tr><td class="k">Interés</td><td>${fmtMoney_(loan.interest)}</td></tr>
      <tr><td class="k">Total a pagar</td><td><b>${fmtMoney_(loan.totalDue)}</b></td></tr>
      <tr><td class="k">Fecha de desembolso</td><td>${fmtDate_(loan.loanDate)}</td></tr>
      <tr><td class="k">Fecha de vencimiento</td><td><b>${fmtDate_(loan.dueDate)}</b></td></tr>
    </table>
    <h2>3. Pago</h2><p>El Prestatario se compromete a pagar <b>${fmtMoney_(loan.totalDue)}</b> a más tardar el <b>${fmtDate_(loan.dueDate)}</b>.</p>
    <h2>4. Mora</h2><p>${esc_(penalty)}</p>
    <p><b>Recargo por mora:</b> se aplicará un recargo del ${lateFeePctText_()} por día sobre el total a pagar
       (${fmtMoney_(loan.totalDue)}), es decir <b>${fmtMoney_(round2_(loan.totalDue * lateFeeRate_()))} por día</b>
       de atraso posterior a la fecha de vencimiento.</p>
    <h2>5. Ley Aplicable</h2><p>Este contrato se rige por las leyes de ${esc_(juris)}.</p>

    <div class="pagebreak"></div>
    <h2>Firmas</h2>
    <p style="margin:6px 0">Por el <b>Prestamista</b> — ${esc_(lender)} (ambos titulares):</p>
    <table class="sign"><tr>
      ${sigCell(signatories[0].sig, signatories[0].name, 'Por ' + lender + ' (Prestamista)')}
      ${sigCell(signatories[1].sig, signatories[1].name, 'Por ' + lender + ' (Prestamista)')}
    </tr></table>
    <p style="margin:28px 0 6px">El <b>Prestatario</b>:</p>
    <table class="sign"><tr>
      ${sigCell(borrowerSig, loan.name, 'Prestatario · DNI ' + loan.dni)}
      <td style="border:none"></td>
    </tr></table>
    ${sig && sig.dataUrl ? `<p style="margin-top:22px">El Prestatario declara haber leído y aceptado los términos de este
       contrato y lo suscribe electrónicamente conforme a la Ley 25.506.</p>
    <table style="width:100%;margin-top:8px;font-size:9pt;color:#555">
      <tr><td class="k" style="background:#f0f4fa;font-weight:bold;width:40%">Firmado electrónicamente por</td><td>${esc_(loan.name)} · DNI ${esc_(loan.dni)}</td></tr>
      <tr><td class="k" style="background:#f0f4fa;font-weight:bold">Correo</td><td>${esc_(loan.email)}</td></tr>
      <tr><td class="k" style="background:#f0f4fa;font-weight:bold">Fecha y hora</td><td>${fmtDateTime_(sig.signedAt || new Date())}</td></tr>
      <tr><td class="k" style="background:#f0f4fa;font-weight:bold">Método</td><td>${sig.method === 'upload' ? 'Imagen de firma cargada' : 'Firma dibujada en pantalla'}</td></tr>
      <tr><td class="k" style="background:#f0f4fa;font-weight:bold">Dispositivo</td><td>${esc_(String(sig.userAgent || '').slice(0, 180))}</td></tr>
    </table>` : ''}
    <p class="foot">${esc_(footer)}</p></body></html>`;
}
function emailAgreement_(loan, file) {
  sendBrandedEmail_(loan.email, 'Contrato de Préstamo — ' + loan.loanId,
    'Estimado/a ' + loan.name + ', adjuntamos su contrato de préstamo ' + loan.loanId + '.',
    '<p>Estimado/a ' + esc_(loan.name) + ',</p><p>Adjuntamos su contrato de préstamo <b>' + esc_(loan.loanId) + '</b>.</p>' +
    '<ul><li>Capital: <b>' + fmtMoney_(loan.principal) + '</b></li><li>Plazo: ' + termLabel_(loan.term) + ' (' + termRatePct_(loan.term) + ')</li>' +
    '<li>Total a pagar: <b>' + fmtMoney_(loan.totalDue) + '</b></li><li>Vencimiento: <b>' + fmtDate_(loan.dueDate) + '</b></li></ul>',
    { attachments: [file.getAs('application/pdf')] });
}

/* ===================== RECORDATORIO DE PAGO (PDF) ===================== */
function makeReminderFile_(loan) {
  const pdf = Utilities.newBlob(reminderHtml_(loan), 'text/html', 'recordatorio.html').getAs('application/pdf');
  const folder = borrowerFolder_(borrowerFolderName_(loan.name, loan.dni));
  const f = folder.createFile(pdf); f.setName('Recordatorio ' + loan.loanId + ' - ' + loan.name + '.pdf');
  return f;
}
function reminderHtml_(loan) {
  const lender = companyName_(), lPhone = getSetting_('Teléfono del Prestamista'),
    lAddr = getSetting_('Dirección del Prestamista'), penalty = getSetting_('Cláusula de Mora'),
    footer = getSetting_('Pie del Contrato');
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const due = loan.dueDate instanceof Date ? new Date(loan.dueDate) : null;
  if (due) due.setHours(0, 0, 0, 0);
  const days = due ? Math.round((due - today) / 86400000) : null;
  const late = due ? daysLate_(loan.dueDate, today) : 0;
  const feeDay = round2_(loan.totalDue * lateFeeRate_());
  const outstanding = computeOutstanding_(loan.principal, termRate_(loan.term), loan.loanDate,
    loanPayments_(loan.loanId), today, new Date(9999, 0, 1));
  // Encabezado y bloque que cambian según falte tiempo o esté en mora.
  let banner, timeLine, feeLine;
  if (days === null) {
    banner = { bg: '#fff2cc', color: '#7f6000', text: 'RECORDATORIO DE PAGO' };
    timeLine = '';
  } else if (days > 0) {
    banner = { bg: '#cfe2f3', color: '#1c4587', text: 'FALTAN ' + days + ' DÍA(S) PARA EL VENCIMIENTO' };
    timeLine = '<p>Su pago vence el <b>' + fmtDate_(loan.dueDate) + '</b>. Le quedan <b>' + days +
      ' día(s)</b> para pagar antes de que se aplique el recargo por mora.</p>';
  } else if (days === 0) {
    banner = { bg: '#fce5cd', color: '#b45f06', text: 'SU PAGO VENCE HOY' };
    timeLine = '<p>Su pago vence <b>hoy, ' + fmtDate_(loan.dueDate) + '</b>. Pague hoy para evitar el recargo por mora.</p>';
  } else {
    banner = { bg: '#f4cccc', color: '#990000', text: 'PAGO VENCIDO — ' + late + ' DÍA(S) DE ATRASO' };
    timeLine = '<p>Su pago venció el <b>' + fmtDate_(loan.dueDate) + '</b>, hace <b>' + late +
      ' día(s)</b>. Ya se está aplicando el recargo por mora indicado abajo.</p>';
  }
  if (late > 0) {
    feeLine = '<p><b>Recargo por mora acumulado:</b> ' + fmtMoney_(round2_(feeDay * late)) +
      ' (' + lateFeePctText_() + ' por día × ' + late + ' día(s)). El recargo sigue creciendo ' +
      fmtMoney_(feeDay) + ' por cada día adicional de atraso.</p>';
  } else {
    feeLine = '<p><b>Recargo por mora:</b> si no paga a tiempo se aplicará un recargo del ' +
      lateFeePctText_() + ' por día sobre el total a pagar (' + fmtMoney_(loan.totalDue) + '), es decir <b>' +
      fmtMoney_(feeDay) + ' por día</b> de atraso.</p>';
  }
  return `<!DOCTYPE html><html lang="es"><head><meta charset="utf-8"><style>
    body{font-family:Georgia,serif;color:#222;margin:48px;line-height:1.5;font-size:12pt}
    h1{text-align:center;font-size:20pt;border-bottom:2px solid #b45f06;padding-bottom:8px}
    h2{font-size:13pt;color:#b45f06;margin-top:24px;border-bottom:1px solid #ccc;padding-bottom:3px}
    table{width:100%;border-collapse:collapse;margin-top:8px} td{padding:6px 8px;border:1px solid #ccc}
    td.k{background:#faf0e6;font-weight:bold;width:45%}
    .banner{margin:16px 0;padding:12px;text-align:center;font-weight:bold;font-size:14pt;border-radius:6px;
      background:${banner.bg};color:${banner.color}}
    .foot{margin-top:40px;font-size:10pt;color:#666;text-align:center}</style></head><body>
    ${brandHeaderHtml_()}
    <h1>RECORDATORIO DE PAGO</h1>
    <p>Estimado/a <b>${esc_(loan.name)}</b>, le escribimos para recordarle el pago de su préstamo
       <b>${esc_(loan.loanId)}</b> con ${esc_(lender)}.</p>
    <div class="banner">${banner.text}</div>
    ${timeLine}
    <h2>Detalle del préstamo</h2>
    <table>
      <tr><td class="k">ID Préstamo</td><td>${esc_(loan.loanId)}</td></tr>
      <tr><td class="k">Total a pagar</td><td><b>${fmtMoney_(loan.totalDue)}</b></td></tr>
      <tr><td class="k">Saldo pendiente (hoy)</td><td><b>${fmtMoney_(outstanding)}</b></td></tr>
      <tr><td class="k">Fecha de vencimiento</td><td><b>${fmtDate_(loan.dueDate)}</b></td></tr>
    </table>
    <h2>Mora</h2>
    ${feeLine}
    ${penalty ? '<p>' + esc_(penalty) + '</p>' : ''}
    <p style="margin-top:20px">Si ya realizó el pago, por favor ignore este mensaje. Ante cualquier duda,
       contáctenos${lPhone ? ' al ' + esc_(lPhone) : ''}.</p>
    <p style="margin-top:24px">${esc_(lender)}${lAddr ? '<br>' + esc_(lAddr) : ''}</p>
    <p class="foot">${esc_(footer)}</p></body></html>`;
}
function emailReminder_(loan, file) {
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const due = loan.dueDate instanceof Date ? new Date(loan.dueDate) : null;
  if (due) due.setHours(0, 0, 0, 0);
  const days = due ? Math.round((due - today) / 86400000) : null;
  const when = days === null ? '' : days > 0 ? ' Faltan ' + days + ' día(s) para el vencimiento.'
    : days === 0 ? ' Su pago vence hoy.' : ' Su pago está vencido hace ' + daysLate_(loan.dueDate, today) + ' día(s).';
  sendBrandedEmail_(loan.email, 'Recordatorio de pago — ' + loan.loanId,
    'Estimado/a ' + loan.name + ', le recordamos el pago de su préstamo ' + loan.loanId + '.' + when,
    '<p>Estimado/a ' + esc_(loan.name) + ',</p><p>Le recordamos el pago de su préstamo <b>' + esc_(loan.loanId) + '</b>.' + esc_(when) + '</p>' +
    '<ul><li>Total a pagar: <b>' + fmtMoney_(loan.totalDue) + '</b></li>' +
    '<li>Vencimiento: <b>' + fmtDate_(loan.dueDate) + '</b></li>' +
    '<li>Recargo por mora: <b>' + fmtMoney_(round2_(loan.totalDue * lateFeeRate_())) + ' por día</b> de atraso (' + lateFeePctText_() + ')</li></ul>' +
    '<p>Adjuntamos el recordatorio en PDF con el detalle. Si ya realizó el pago, ignore este mensaje.</p>',
    { attachments: [file.getAs('application/pdf')] });
}

/* ===================== PAGOS + RECIBO ===================== */
function recordPayment(loanId, amountStr, dateStr) {
  return guard_('recordPayment', function () {
    loanId = String(loanId || '').trim();
    const loan = findLoanById_(loanId);
    if (!loan) throw new Error('No se encontró el préstamo "' + loanId + '".');
    const amount = Number(String(amountStr || '').replace(/[^0-9.\-]/g, ''));
    if (!amount || amount <= 0) throw new Error('Monto inválido.');
    const payDate = dateStr ? new Date(dateStr) : new Date();
    if (isNaN(payDate.getTime())) throw new Error('Fecha inválida.');
    const rate = termRate_(loan.term);
    const existing = loanPayments_(loanId); existing.push({ date: payDate, amount: amount });
    const outAfter = computeOutstanding_(loan.principal, rate, loan.loanDate, existing, payDate, new Date(9999, 0, 1));
    const psh = getSS_().getSheetByName(CFG.SHEETS.PAYMENTS);
    const payId = 'P-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyMMdd-HHmmss');
    psh.appendRow([payId, loanId, firstName_(loan.name), lastName_(loan.name), loan.dni, payDate, amount, outAfter, '']);
    const row = psh.getLastRow();
    psh.getRange(row, 6).setNumberFormat('yyyy-mm-dd'); psh.getRange(row, 7, 1, 2).setNumberFormat(CFG.CURRENCY_FMT);
    updateLoanOutstanding_(loanId);
    if (loan.email) {
      sendBrandedEmail_(loan.email, 'Pago recibido — ' + loanId,
        'Estimado/a ' + loan.name + ', confirmamos su pago de ' + fmtMoney_(amount) + '. Saldo: ' + fmtMoney_(outAfter) + '.',
        receiptHtml_(loan, payDate, amount, outAfter));
      psh.getRange(row, 9).setValue(new Date());
    }
    return 'Pago registrado: ' + fmtMoney_(amount) + '. Saldo posterior: ' + fmtMoney_(outAfter) +
      (loan.email ? '. Recibo enviado a ' + loan.email : '');
  });
}
function receiptHtml_(loan, date, amount, out) {
  const cleared = out <= 0.009;
  return '<p>Estimado/a ' + esc_(loan.name) + ',</p><p>Confirmamos su pago del préstamo <b>' + esc_(loan.loanId) + '</b>.</p>' +
    '<table style="border-collapse:collapse">' +
    row_('Fecha de pago', fmtDate_(date)) + row_('Monto pagado', fmtMoney_(amount)) + row_('Saldo pendiente', '<b>' + fmtMoney_(out) + '</b>') +
    '</table>' + (cleared ? '<p style="color:#38761d"><b>¡Su préstamo está totalmente pagado. Gracias!</b></p>' : '') +
    '<p>Saludos,<br>' + esc_(companyName_()) + '</p>';
}
function row_(k, v) { return '<tr><td style="padding:4px 10px;border:1px solid #ccc;background:#f0f4fa"><b>' + k + '</b></td><td style="padding:4px 10px;border:1px solid #ccc">' + v + '</td></tr>'; }

/** Datos derivados del estado de cuenta (compartidos por el PDF y el correo). */
function statementData_(loan) {
  const pays = loanPayments_(loan.loanId).sort((a, b) => a.date - b.date);
  const rate = termRate_(loan.term);
  const out = computeOutstanding_(loan.principal, rate, loan.loanDate, pays, new Date(), new Date(9999, 0, 1));
  const totalPaid = pays.reduce((s, p) => s + p.amount, 0);
  const paid = out <= 0.009;
  const overdue = !paid && (loan.dueDate instanceof Date) && (new Date().getTime() > loan.dueDate.getTime());
  const dLate = daysLate_(loan.dueDate);
  const feeAccum = overdue ? round2_(loan.totalDue * lateFeeRate_() * dLate) : 0;
  return { pays, out, totalPaid, paid, overdue, dLate, feeAccum };
}

/** PDF del estado de cuenta, guardado en la carpeta del prestatario. */
function makeStatementFile_(loan) {
  const pdf = Utilities.newBlob(statementHtml_(loan), 'text/html', 'estado.html').getAs('application/pdf');
  const folder = borrowerFolder_(borrowerFolderName_(loan.name, loan.dni));
  const f = folder.createFile(pdf); f.setName('Estado de Cuenta ' + loan.loanId + ' - ' + loan.name + '.pdf');
  return f;
}

/** HTML del estado de cuenta: banner de estado + detalle + pagos. Mensaje amable si está saldado. */
function statementHtml_(loan) {
  const lender = companyName_(), lPhone = getSetting_('Teléfono del Prestamista'),
    lAddr = getSetting_('Dirección del Prestamista'), footer = getSetting_('Pie del Contrato');
  const d = statementData_(loan);
  let banner, message;
  if (d.paid) {
    banner = { bg: '#d9ead3', color: '#38761d', text: 'PRÉSTAMO SALDADO ✓' };
    message = `<p style="background:#eaf5e6;border-left:4px solid #38761d;padding:12px 16px;border-radius:6px;font-size:12.5pt">
      🎉 <b>¡Felicitaciones, ${esc_(loan.name)}!</b> Su préstamo está <b>totalmente saldado</b>.
      Muchas gracias por su cumplimiento y su confianza — fue un placer acompañarle.
      Esperamos poder ayudarle nuevamente cuando lo necesite.</p>`;
  } else if (d.overdue) {
    banner = { bg: '#f4cccc', color: '#990000', text: 'PAGO VENCIDO — ' + d.dLate + ' DÍA(S) DE ATRASO' };
    message = `<p>Su préstamo presenta un saldo pendiente vencido. Le pedimos regularizar el pago a la brevedad
      para evitar que el recargo por mora siga creciendo.</p>`;
  } else {
    banner = { bg: '#cfe2f3', color: '#1c4587', text: 'PRÉSTAMO ACTIVO' };
    message = `<p>A continuación encontrará el detalle de su préstamo y de los pagos registrados a la fecha.</p>`;
  }
  const cell = 'padding:6px 8px;border:1px solid #ccc';
  let payRows = d.pays.map(p => `<tr><td style="${cell}">${fmtDate_(p.date)}</td><td style="${cell};text-align:right">${fmtMoney_(p.amount)}</td></tr>`).join('');
  if (!payRows) payRows = `<tr><td colspan="2" style="${cell}">— Sin pagos registrados —</td></tr>`;
  const feeRows = d.overdue ?
    `<tr><td class="k">Días de atraso</td><td>${d.dLate}</td></tr>
     <tr><td class="k" style="color:#900">Recargo por mora (${lateFeePctText_()}/día)</td><td style="color:#900"><b>${fmtMoney_(d.feeAccum)}</b></td></tr>` : '';
  return `<!DOCTYPE html><html lang="es"><head><meta charset="utf-8"><style>
    body{font-family:Georgia,serif;color:#222;margin:48px;line-height:1.5;font-size:12pt}
    h1{text-align:center;font-size:20pt;border-bottom:2px solid #674ea7;padding-bottom:8px}
    h2{font-size:13pt;color:#674ea7;margin-top:24px;border-bottom:1px solid #ccc;padding-bottom:3px}
    table{width:100%;border-collapse:collapse;margin-top:8px} td{padding:6px 8px;border:1px solid #ccc}
    td.k{background:#f3eefb;font-weight:bold;width:45%}
    .banner{margin:16px 0;padding:12px;text-align:center;font-weight:bold;font-size:14pt;border-radius:6px;
      background:${banner.bg};color:${banner.color}}
    .foot{margin-top:40px;font-size:10pt;color:#666;text-align:center}</style></head><body>
    ${brandHeaderHtml_()}
    <h1>ESTADO DE CUENTA</h1>
    <p>Estimado/a <b>${esc_(loan.name)}</b>, este es el estado de cuenta de su préstamo
       <b>${esc_(loan.loanId)}</b> con ${esc_(lender)}, al <b>${fmtDate_(new Date())}</b>.</p>
    <div class="banner">${banner.text}</div>
    ${message}
    <h2>Detalle del préstamo</h2>
    <table>
      <tr><td class="k">ID Préstamo</td><td>${esc_(loan.loanId)}</td></tr>
      <tr><td class="k">DNI</td><td>${esc_(loan.dni)}</td></tr>
      <tr><td class="k">Capital</td><td>${fmtMoney_(loan.principal)}</td></tr>
      <tr><td class="k">Total a pagar</td><td><b>${fmtMoney_(loan.totalDue)}</b></td></tr>
      <tr><td class="k">Fecha de vencimiento</td><td>${fmtDate_(loan.dueDate)}</td></tr>
      <tr><td class="k">Total pagado</td><td><b>${fmtMoney_(d.totalPaid)}</b></td></tr>
      ${feeRows}
      <tr><td class="k">Saldo pendiente</td><td><b>${fmtMoney_(d.out)}</b></td></tr>
    </table>
    <h2>Pagos registrados</h2>
    <table><tr><td class="k" style="width:50%">Fecha</td><td class="k" style="text-align:right">Monto</td></tr>
      ${payRows}
      <tr><td style="${cell}"><b>Total pagado</b></td><td style="${cell};text-align:right"><b>${fmtMoney_(d.totalPaid)}</b></td></tr>
    </table>
    <p style="margin-top:20px">Ante cualquier duda, contáctenos${lPhone ? ' al ' + esc_(lPhone) : ''}.</p>
    <p style="margin-top:14px">${esc_(lender)}${lAddr ? '<br>' + esc_(lAddr) : ''}</p>
    <p class="foot">${esc_(footer)}</p></body></html>`;
}

/** Envía el estado de cuenta con el PDF adjunto. Mensaje amable si el préstamo está saldado. */
function sendStatementEmail_(loan, file) {
  const d = statementData_(loan);
  const cell = 'padding:4px 10px;border:1px solid #ccc';
  let rows = d.pays.map(p => '<tr><td style="' + cell + '">' + fmtDate_(p.date) + '</td><td style="' + cell + ';text-align:right">' + fmtMoney_(p.amount) + '</td></tr>').join('');
  if (!rows) rows = '<tr><td colspan="2" style="' + cell + '">Sin pagos registrados.</td></tr>';
  const feeRows = d.overdue ?
    '<tr><td style="' + cell + '"><b>Días de atraso</b></td><td style="' + cell + ';text-align:right">' + d.dLate + '</td></tr>' +
    '<tr><td style="' + cell + ';color:#900"><b>Recargo por mora (' + lateFeePctText_() + '/día)</b></td><td style="' + cell + ';text-align:right;color:#900"><b>' + fmtMoney_(d.feeAccum) + '</b></td></tr>' : '';
  const intro = d.paid
    ? '<p>Estimado/a ' + esc_(loan.name) + ',</p>' +
      '<p style="background:#eaf5e6;border-left:4px solid #38761d;padding:10px 14px;border-radius:4px">' +
      '🎉 <b>¡Felicitaciones!</b> Su préstamo <b>' + esc_(loan.loanId) + '</b> está <b>totalmente saldado</b>. ' +
      'Muchas gracias por su cumplimiento y su confianza — fue un placer acompañarle. ' +
      'Esperamos poder ayudarle nuevamente cuando lo necesite.</p>'
    : '<p>Estimado/a ' + esc_(loan.name) + ',</p><p>Le enviamos el estado de cuenta de su préstamo <b>' + esc_(loan.loanId) + '</b>:</p>';
  const html = intro +
    '<p>Capital ' + fmtMoney_(loan.principal) + ' · Total a pagar ' + fmtMoney_(loan.totalDue) + ' · Vence ' + fmtDate_(loan.dueDate) + '</p>' +
    '<table style="border-collapse:collapse"><tr><th style="' + cell + ';background:#f0f4fa">Fecha</th><th style="' + cell + ';background:#f0f4fa">Monto</th></tr>' +
    rows + '<tr><td style="' + cell + '"><b>Total pagado</b></td><td style="' + cell + ';text-align:right"><b>' + fmtMoney_(d.totalPaid) + '</b></td></tr>' +
    feeRows +
    '<tr><td style="' + cell + '"><b>Saldo pendiente</b></td><td style="' + cell + ';text-align:right"><b>' + fmtMoney_(d.out) + '</b></td></tr></table>' +
    '<p>Adjuntamos el estado de cuenta en PDF.</p>' +
    '<p>Saludos,<br>' + esc_(companyName_()) + '</p>';
  const subject = d.paid ? '¡Préstamo saldado! — ' + loan.loanId : 'Estado de cuenta — ' + loan.loanId;
  const plain = d.paid ? 'Su préstamo ' + loan.loanId + ' está totalmente saldado. ¡Muchas gracias!'
    : 'Estado del préstamo ' + loan.loanId + '. Total pagado ' + fmtMoney_(d.totalPaid) + '. Saldo ' + fmtMoney_(d.out) + '.';
  sendBrandedEmail_(loan.email, subject, plain, html, { attachments: [file.getAs('application/pdf')] });
  return (d.paid ? 'Estado (préstamo saldado) enviado a ' : 'Estado de cuenta enviado a ') + loan.email + '.';
}

function emailStatement(loanId) {
  return guard_('emailStatement', function () {
    const loan = findLoanById_(String(loanId || '').trim());
    if (!loan) throw new Error('No se encontró el préstamo.');
    if (!loan.email) throw new Error('El prestatario no tiene correo.');
    return sendStatementEmail_(loan, makeStatementFile_(loan));
  });
}

/* ===================== SIDEBAR (PANEL) ===================== */
function sidebarHtml_() {
  const ctx = activeLoanId_();
  return `<!DOCTYPE html><html><head><base target="_top"><style>
    body{font-family:Arial,sans-serif;margin:0;padding:12px;font-size:13px;color:#222}
    h3{color:#1c4587;margin:14px 0 6px;border-bottom:1px solid #ddd;padding-bottom:3px}
    input{width:100%;box-sizing:border-box;padding:7px;margin:3px 0 8px;border:1px solid #ccc;border-radius:5px}
    button{width:100%;padding:9px;margin:4px 0;border:0;border-radius:5px;color:#fff;font-weight:bold;cursor:pointer}
    .blue{background:#1c4587}.green{background:#38761d}.purple{background:#674ea7}.grey{background:#666}
    #msg{margin-top:10px;padding:8px;border-radius:5px;display:none;white-space:pre-wrap}
    .ok{background:#d9ead3}.err{background:#f4cccc}</style></head><body>
    <h3>Préstamo</h3>
    <label>ID Préstamo</label><input id="loan" value="${esc_(ctx || '')}" placeholder="L-0001">
    <button class="blue" onclick="run('previewAgreement','Generando vista previa…')">👁 Vista previa del contrato</button>
    <button class="blue" onclick="run('sendAgreement','Enviando contrato…')">✉ Enviar contrato por correo</button>
    <button class="purple" onclick="run('sendStatement','Enviando estado…')">✉ Enviar estado de cuenta</button>
    <h3>Registrar pago</h3>
    <label>Monto (ARS)</label><input id="amount" type="number" min="1" step="any">
    <label>Fecha (opcional)</label><input id="date" type="date">
    <button class="green" onclick="pay()">💵 Registrar pago</button>
    <h3>General</h3>
    <button class="grey" onclick="run('refresh','Actualizando…')">🔄 Actualizar saldos y resumen</button>
    <button class="green" onclick="run('moveCleared','Moviendo saldados…')">✔ Mover saldados → hoja Saldados</button>
    <h3>Pestañas</h3>
    <button class="blue" onclick="run('dailyOnly','Ajustando pestañas…')">🗂 Ver solo operaciones diarias</button>
    <button class="grey" onclick="run('allSheets','Mostrando hojas…')">👁 Mostrar todas las hojas</button>
    <div id="msg"></div>
    <script>
      var msg=document.getElementById('msg');
      function show(t,ok){msg.style.display='block';msg.className=ok?'ok':'err';msg.textContent=t;}
      function loanId(){return document.getElementById('loan').value.trim();}
      function run(kind,busy){
        show(busy,true);
        var f={previewAgreement:'sbPreviewAgreement',sendAgreement:'sbSendAgreement',sendStatement:'sbSendStatement',refresh:'sbRefresh',moveCleared:'sbMoveCleared',dailyOnly:'sbDailyOnly',allSheets:'sbAllSheets'}[kind];
        google.script.run.withSuccessHandler(function(r){show(r,true);}).withFailureHandler(function(e){show(e.message||e,false);})[f](loanId());
      }
      function pay(){
        show('Registrando pago…',true);
        google.script.run.withSuccessHandler(function(r){show(r,true);}).withFailureHandler(function(e){show(e.message||e,false);})
          .recordPayment(loanId(),document.getElementById('amount').value,document.getElementById('date').value);
      }
    </script></body></html>`;
}
function activeLoanId_() {
  try {
    const sh = SpreadsheetApp.getActiveSheet();
    if (sh.getName() !== CFG.SHEETS.BORROWERS) return '';
    const row = sh.getActiveRange().getRow();
    return row >= 2 ? String(sh.getRange(row, 1).getValue()).trim() : '';
  } catch (e) { return ''; }
}
function sbPreviewAgreement(loanId) {
  return guard_('sbPreviewAgreement', function () {
    const loan = findLoanById_(String(loanId).trim());
    if (!loan) throw new Error('No se encontró el préstamo.');
    const file = makeAgreementFile_(loan);
    return 'Vista previa creada. Ábrala aquí:\n' + file.getUrl();
  });
}
function sbSendAgreement(loanId) {
  return guard_('sbSendAgreement', function () {
    const loan = findLoanById_(String(loanId).trim());
    if (!loan) throw new Error('No se encontró el préstamo.');
    if (!loan.email) throw new Error('El prestatario no tiene correo.');
    const file = makeAgreementFile_(loan); emailAgreement_(loan, file);
    markAgreementSent_(loan.loanId, file);
    return 'Contrato enviado a ' + loan.email + '.';
  });
}
function sbSendStatement(loanId) { return emailStatement(loanId); }
function sbRefresh() { refreshAll(true); return 'Saldos y resumen actualizados.'; }
function sbMoveCleared() {
  const n = doMoveCleared_();
  return n ? ('Se movieron ' + n + ' préstamo(s) saldado(s) a la hoja "Saldados".')
    : 'No hay préstamos saldados (pagados en su totalidad) para mover.';
}
function sbDailyOnly() {
  const n = applyDailyVisibility_(getSS_());
  return 'Vista de operaciones diarias: ' + n + ' hoja(s) secundaria(s) oculta(s).';
}
function sbAllSheets() {
  const ss = getSS_(); let shown = 0;
  ss.getSheets().forEach(function (sh) { try { if (sh.isSheetHidden()) shown++; sh.showSheet(); } catch (e) {} });
  return shown + ' hoja(s) mostrada(s). Todas las pestañas están visibles.';
}

function markAgreementSent_(loanId, file) {
  const bs = getSS_().getSheetByName(CFG.SHEETS.BORROWERS), last = bs.getLastRow();
  const ids = bs.getRange(2, 1, Math.max(last - 1, 1), 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]).trim() === loanId) {
    bs.getRange(i + 2, 16).setFormula('=HYPERLINK("' + file.getUrl() + '","Ver contrato")');
    bs.getRange(i + 2, 17).setValue(new Date()); return;
  }
}

/* ===================== ACTUALIZAR ===================== */
function refreshAll(silent) {
  const ss = getSS_(), N = CFG.MAX_ROWS;
  cleanupLegacySheets_(ss);
  updateAllOutstanding_();
  setupSummary_(ss); setupPanel_(ss); setupStats_(ss); setupStatements_(ss); rebuildLateSheet_(ss); rebuildRemindersSheet_(ss);
  SpreadsheetApp.flush();
  if (!silent) SpreadsheetApp.getUi().alert('Actualizado. Saldos, resumen, panel, estadísticas y pagos atrasados recalculados.');
}
function cleanupLegacySheets_(ss) {
  ss.getSheets().forEach(s => {
    const n = s.getName();
    if (/^New Borrowers \(old/i.test(n) || n === 'New Borrowers (pre-form)') { try { ss.deleteSheet(s); } catch (e) {} }
  });
}

/* ===================== FORMULARIO WEB ===================== */
function showWebFormLink() {
  setupNew_(getSS_());
  let url = '';
  try { url = ScriptApp.getService().getUrl(); } catch (e) { url = ''; }
  if (url) {
    PropertiesService.getScriptProperties().setProperty('WEBAPP_URL', url);
    const cfg = String(getSetting_('URL de la app web (enlaces a clientes)') || '').trim();
    SpreadsheetApp.getUi().alert('Formulario de solicitud (compartir con solicitantes):\n' + url +
      '\n\nConsulta de saldo (para prestatarios):\n' + url + '?page=saldo' +
      '\n\nIMPORTANTE: para que los enlaces de FIRMA que se envían a los clientes usen la URL correcta, ' +
      'pegue la URL de su implementación (la que termina en /exec) en:\n' +
      'Configuración ▸ "URL de la app web (enlaces a clientes)".' +
      (cfg ? '\n\nActualmente configurada:\n' + cfg : '\n\n(Actualmente vacía — se usa la URL detectada automáticamente.)'));
  } else {
    SpreadsheetApp.getUi().alert('Publique el formulario (una vez):\n\n1. Extensions ▸ Apps Script.\n' +
      '2. Deploy ▸ Nueva implementación ▸ ⚙ ▸ Aplicación web.\n3. Ejecutar como: Yo.  Acceso: Cualquiera.\n' +
      '4. Implementar ▸ autorizar ▸ copiar la URL.\n\nLuego ejecute ⑥ otra vez para ver el enlace.');
  }
}

function doGet(e) {
  const page = e && e.parameter && e.parameter.page;
  if (page === 'saldo') return HtmlService.createHtmlOutput(balancePageHtml_()).setTitle('Consulta de Saldo')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  if (page === 'firmar') return HtmlService.createHtmlOutput(signingPageHtml_(String((e.parameter.loan) || '').trim()))
    .setTitle('Firmar Contrato').addMetaTag('viewport', 'width=device-width, initial-scale=1');
  const html = acceptingApplications_() ? intakeHtml_() : closedHtml_();
  return HtmlService.createHtmlOutput(html).setTitle('Solicitud de Préstamo')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Página que se muestra cuando las solicitudes están cerradas. */
function closedHtml_() {
  return `<!DOCTYPE html><html lang="es"><head><base target="_top"><style>
    body{font-family:Arial,sans-serif;background:#f4f6fa;margin:0;padding:24px;color:#222}
    .card{max-width:560px;margin:0 auto;background:#fff;border-radius:10px;padding:28px;box-shadow:0 2px 12px rgba(0,0,0,.08);text-align:center}
    h1{font-size:20px;color:#1c4587;margin:8px 0}
    .badge{display:inline-block;background:#fce8e6;color:#a50e0e;font-weight:bold;padding:8px 16px;border-radius:999px;margin:8px 0 4px}
    p{color:#555;line-height:1.5}</style></head><body>
    <div class="card">${brandHeaderHtml_()}
      <div class="badge">Solicitudes cerradas</div>
      <h1>Por el momento no hay préstamos disponibles</h1>
      <p>Gracias por su interés. En este momento <b>no estamos aceptando nuevas solicitudes de préstamo</b>.
         Le invitamos a intentarlo nuevamente más adelante.</p>
      <p style="font-size:13px;color:#888">${esc_(companyName_())}</p>
    </div></body></html>`;
}

function intakeHtml_() {
  const lender = esc_(companyName_());
  const feePct = lateFeePctText_();
  const juris = esc_(getSetting_('Jurisdicción') || 'Buenos Aires, Argentina');
  const moraClause = esc_(getSetting_('Cláusula de Mora') ||
    ('Se aplica un recargo por mora del ' + feePct + ' diario sobre el total a pagar por cada día de atraso posterior a la fecha de vencimiento.'));
  const contractFooter = esc_(getSetting_('Pie del Contrato') ||
    'Este acuerdo es legalmente vinculante desde la firma de ambas partes.');
  return `<!DOCTYPE html><html lang="es"><head><base target="_top"><style>
    body{font-family:Arial,sans-serif;background:#f4f6fa;margin:0;padding:24px;color:#222}
    .card{max-width:560px;margin:0 auto;background:#fff;border-radius:10px;padding:24px 28px;box-shadow:0 2px 12px rgba(0,0,0,.08)}
    h1{font-size:20px;color:#1c4587;margin:0 0 4px}.sub{color:#666;margin:0 0 18px}
    label{display:block;font-weight:bold;margin:14px 0 4px;font-size:14px}.req{color:#c00}
    input,textarea,select{width:100%;box-sizing:border-box;padding:9px 10px;border:1px solid #ccc;border-radius:6px;font-size:14px}
    .hint{font-weight:normal;color:#888;font-size:12px}
    #calc{margin-top:8px;padding:10px;background:#f0f4fa;border-radius:6px;font-size:13px;display:none}
    .amount-words{display:none;margin-top:5px;font-size:12.5px;font-style:italic;color:#1c4587;line-height:1.35}
    button{margin-top:22px;width:100%;background:#1c4587;color:#fff;border:0;border-radius:6px;padding:12px;font-size:15px;font-weight:bold;cursor:pointer}
    button:disabled{background:#9db4d6}
    #msg{margin-top:16px;padding:12px;border-radius:6px;display:none}.ok{background:#d9ead3;color:#274e13}.err{background:#f4cccc;color:#900}
    /* Cargando */
    #loadingView{display:none;text-align:center;padding:34px 0}
    .spinner{width:56px;height:56px;margin:0 auto 18px;border:6px solid #e3e9f4;border-top-color:#1c4587;border-radius:50%;animation:spin .9s linear infinite}
    @keyframes spin{to{transform:rotate(360deg)}}
    /* Éxito */
    #successView{display:none}
    .check{width:72px;height:72px;line-height:72px;margin:6px auto 10px;border-radius:50%;background:#38761d;color:#fff;font-size:40px;font-weight:bold;text-align:center}
    .summary{width:100%;border-collapse:collapse;margin:18px 0}
    .summary td{padding:9px 12px;border:1px solid #e2e6ee;font-size:14px}
    .summary td.k{background:#f0f4fa;font-weight:bold;width:46%}
    .note{background:#e8f0fe;border-radius:8px;padding:12px 14px;font-size:13px;color:#274e6b;margin-top:6px}
    /* Términos y Condiciones */
    .terms{margin-top:18px;border:1px solid #ccd3e0;border-radius:8px;background:#fafbfe}
    .terms h2{font-size:14px;color:#1c4587;margin:0;padding:10px 12px;border-bottom:1px solid #e2e6ee;background:#f0f4fa;border-radius:8px 8px 0 0}
    .terms .body{max-height:160px;overflow-y:auto;padding:8px 14px;font-size:12.5px;color:#444;line-height:1.5}
    .terms ul{margin:6px 0 0;padding-left:18px}.terms li{margin:5px 0}
    .agree{display:flex;align-items:flex-start;gap:9px;margin:14px 0 0;font-weight:normal;font-size:13.5px;cursor:pointer}
    .agree input{width:18px;height:18px;flex:0 0 auto;margin-top:1px}</style></head><body>
    <div class="card" id="card">${brandHeaderHtml_()}
    <div id="formView">
      <h1>Solicitud de Préstamo</h1>
      <p class="sub">${lender} — todos los campos son obligatorios. Su solicitud será revisada y verificada.</p>
      <form id="f">
        <label>Nombre completo <span class="req">*</span><input name="fullName" required></label>
        <label>Correo electrónico <span class="req">*</span><input name="email" type="email" required></label>
        <label>DNI <span class="req">*</span><input name="dni" required placeholder="XX.XXX.XXX"></label>
        <label>Teléfono <span class="req">*</span><input name="phone" required></label>
        <label>Monto solicitado (ARS) <span class="req">*</span><input name="amount" id="amount" type="text" inputmode="numeric" autocomplete="off" placeholder="Ej: 100.000" required>
          <span id="amountWords" class="amount-words" aria-live="polite"></span></label>
        <label>Plazo <span class="req">*</span><select name="term" id="term" required>
          <option value="">— elegir —</option><option value="15">15 días (25% de interés)</option><option value="1">1 mes (50% de interés)</option><option value="2">2 meses (100% de interés)</option></select></label>
        <div id="calc"></div>
        <label>Notas / Motivo <span class="req">*</span><textarea name="notes" rows="2" required></textarea></label>
        <label>Foto del frente del DNI <span class="req">*</span> <span class="hint">— imagen o PDF</span><input name="dniPhoto" type="file" accept="image/*,.pdf" capture="environment" required></label>
        <label>Foto del dorso del DNI <span class="req">*</span> <span class="hint">— imagen o PDF</span><input name="cuilPhoto" type="file" accept="image/*,.pdf" capture="environment" required></label>
        <div class="terms">
          <h2>Términos y Condiciones del Préstamo</h2>
          <div class="body">
            <p>Al enviar esta solicitud usted declara haber leído, comprendido y aceptado los siguientes términos con ${lender}:</p>
            <ul>
              <li><b>Intereses según el plazo:</b> 15 días = 25%, 1 mes = 50%, 2 meses = 100% sobre el capital solicitado.</li>
              <li><b>Devolución:</b> el capital más los intereses se devuelven en su totalidad en la fecha de vencimiento.</li>
              <li><b>Recargo por mora:</b> ${moraClause}</li>
              <li><b>Verificación:</b> la solicitud está sujeta a revisión y verificación de identidad (DNI y CUIL). La aprobación no está garantizada.</li>
              <li><b>Declaración:</b> declaro que los datos y documentos aportados son verídicos y de mi titularidad.</li>
              <li><b>Jurisdicción:</b> ${juris}.</li>
              <li>${contractFooter}</li>
            </ul>
          </div>
        </div>
        <label class="agree"><input type="checkbox" name="agree" id="agree" required>
          <span>He leído y acepto los <b>Términos y Condiciones</b> del préstamo. <span class="req">*</span></span></label>
        <button type="submit" id="btn">Enviar solicitud</button></form>
      <div id="msg"></div>
    </div>
    <div id="loadingView">
      <div class="spinner"></div>
      <p style="font-weight:bold;color:#1c4587;margin:0">Enviando su solicitud…</p>
      <p style="color:#888;font-size:13px;margin:6px 0 0">Subiendo sus documentos. Esto puede tardar unos segundos — por favor no cierre esta ventana.</p>
    </div>
    <div id="successView"></div>
    </div>
    <script>
      var f=document.getElementById('f'),btn=document.getElementById('btn'),msg=document.getElementById('msg'),calc=document.getElementById('calc');
      var formView=document.getElementById('formView'),loadingView=document.getElementById('loadingView'),successView=document.getElementById('successView');
      function esc(s){return String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');}
      // Mensajes de validación en español, sin importar el idioma del navegador.
      Array.prototype.forEach.call(f.elements,function(el){
        if(!el.name)return;
        el.addEventListener('invalid',function(){
          if(el.validity.valueMissing) el.setCustomValidity('Este campo es obligatorio.');
          else if(el.validity.typeMismatch) el.setCustomValidity('Ingrese un '+(el.type==='email'?'correo electrónico válido':'valor válido')+'.');
          else if(el.validity.rangeUnderflow||el.validity.badInput) el.setCustomValidity('Ingrese un monto válido.');
          else el.setCustomValidity('');
        });
        el.addEventListener('input',function(){el.setCustomValidity('');});
        el.addEventListener('change',function(){el.setCustomValidity('');});
      });
      // Mensaje específico para la casilla de Términos y Condiciones.
      var agree=document.getElementById('agree');
      agree.addEventListener('invalid',function(){if(agree.validity.valueMissing)agree.setCustomValidity('Debe aceptar los Términos y Condiciones para enviar la solicitud.');});
      agree.addEventListener('change',function(){agree.setCustomValidity('');});
      function fmt(n){try{return n.toLocaleString('es-AR',{style:'currency',currency:'ARS'});}catch(e){return '$'+n.toFixed(2);}}
      var amountEl=document.getElementById('amount'),amountWords=document.getElementById('amountWords');
      // Convierte solo dígitos a valor numérico (los puntos son separadores de miles).
      function parseAmount(v){var d=String(v==null?'':v).replace(/[^0-9]/g,'');return d?parseInt(d,10):0;}
      // Agrupa los dígitos con puntos de miles al estilo argentino: 1000000 -> 1.000.000
      // (agrupación manual, sin regex con backslash: se pierden dentro de la plantilla de texto).
      function groupNum(d){var out='',c=0;for(var i=d.length-1;i>=0;i--){out=d.charAt(i)+out;c++;if(c%3===0&&i>0){out='.'+out;}}return out;}
      // Convierte un entero a su expresión en palabras (español).
      function numeroALetras(num){
        num=Math.floor(Math.abs(num));if(num===0)return 'cero';
        var U=['','uno','dos','tres','cuatro','cinco','seis','siete','ocho','nueve','diez','once','doce','trece','catorce','quince','dieciséis','diecisiete','dieciocho','diecinueve','veinte','veintiuno','veintidós','veintitrés','veinticuatro','veinticinco','veintiséis','veintisiete','veintiocho','veintinueve'];
        var D=['','','','treinta','cuarenta','cincuenta','sesenta','setenta','ochenta','noventa'];
        var C=['','ciento','doscientos','trescientos','cuatrocientos','quinientos','seiscientos','setecientos','ochocientos','novecientos'];
        function seccion(n){ // n: 0..999, apocopa uno->un (siempre precede a un sustantivo)
          if(n===0)return '';if(n===100)return 'cien';
          var c=Math.floor(n/100),r=n%100,t=c>0?C[c]:'';
          if(r>0){if(t)t+=' ';
            if(r<30){t+=(r===1)?'un':(r===21)?'veintiún':U[r];}
            else{var dd=Math.floor(r/10),u=r%10;t+=D[dd];if(u>0)t+=' y '+((u===1)?'un':U[u]);}}
          return t;
        }
        var p=[],mill=Math.floor(num/1000000),mil=Math.floor((num%1000000)/1000),r=num%1000;
        if(mill>0)p.push(mill===1?'un millón':seccion(mill)+' millones');
        if(mil>0)p.push(mil===1?'mil':seccion(mil)+' mil');
        if(r>0)p.push(seccion(r));
        return p.join(' ').trim();
      }
      function capitalizar(s){return s?s.charAt(0).toUpperCase()+s.slice(1):s;}
      // Reformatea el campo mientras se escribe, preservando la posición del cursor, y muestra el monto en palabras.
      function formatAmountInput(){
        var start=amountEl.selectionStart||0;
        var digitsBefore=amountEl.value.slice(0,start).replace(/[^0-9]/g,'').length;
        var digits=amountEl.value.replace(/[^0-9]/g,'').replace(/^0+/,'');
        var formatted=digits?groupNum(digits):'';
        amountEl.value=formatted;
        var pos=formatted.length,count=0;
        if(digitsBefore>0){for(var i=0;i<formatted.length;i++){if(formatted.charAt(i)!=='.')count++;if(count>=digitsBefore){pos=i+1;break;}}}
        else pos=0;
        try{amountEl.setSelectionRange(pos,pos);}catch(e){}
        var n=digits?parseInt(digits,10):0;
        if(n>0){var w=numeroALetras(n),noun=/(millón|millones)$/.test(w)?' de pesos':(n===1?' peso':' pesos');
          amountWords.textContent=capitalizar(w)+noun;amountWords.style.display='block';}
        else{amountWords.textContent='';amountWords.style.display='none';}
      }
      function recalc(){var a=parseAmount(amountEl.value),t=document.getElementById('term').value;
        if(a>0&&(t==='15'||t==='1'||t==='2')){var r=t==='15'?0.25:t==='1'?0.5:1;calc.style.display='block';
          calc.innerHTML='Interés: <b>'+fmt(a*r)+'</b><br>Total a devolver: <b>'+fmt(a*(1+r))+'</b>';}else{calc.style.display='none';}}
      amountEl.addEventListener('input',function(){formatAmountInput();recalc();});
      document.getElementById('term').addEventListener('change',recalc);
      function onSuccess(r){
        loadingView.style.display='none';
        if(!r||!r.ok){ // respaldo por si el servidor devolviera texto
          successView.innerHTML='<div style="text-align:center"><div class="check">✓</div><h1 style="color:#38761d">¡Solicitud recibida!</h1><p class="sub">'+esc(r&&r.message?r.message:r)+'</p></div>';
          successView.style.display='block';return;
        }
        successView.innerHTML=
          '<div style="text-align:center"><div class="check">&#10003;</div>'+
          '<h1 style="color:#38761d">¡Solicitud recibida!</h1>'+
          '<p class="sub">Gracias, <b>'+esc(r.name)+'</b>. Su solicitud fue registrada correctamente y está pendiente de revisión.</p></div>'+
          '<table class="summary">'+
          '<tr><td class="k">Solicitante</td><td>'+esc(r.name)+'</td></tr>'+
          '<tr><td class="k">Correo</td><td>'+esc(r.email)+'</td></tr>'+
          '<tr><td class="k">DNI</td><td>'+esc(r.dni)+'</td></tr>'+
          '<tr><td class="k">Teléfono</td><td>'+esc(r.phone)+'</td></tr>'+
          '<tr><td class="k">Monto solicitado</td><td>'+esc(r.amountFmt)+'</td></tr>'+
          '<tr><td class="k">Plazo</td><td>'+esc(r.termLabel)+' ('+esc(r.ratePct)+' de interés)</td></tr>'+
          '<tr><td class="k">Interés</td><td>'+esc(r.interestFmt)+'</td></tr>'+
          '<tr><td class="k">Total a devolver</td><td><b>'+esc(r.totalFmt)+'</b></td></tr>'+
          '</table>'+
          '<div class="note">Enviamos un correo de confirmación a <b>'+esc(r.email)+'</b>. Revisaremos su solicitud y nos pondremos en contacto a la brevedad. Puede cerrar esta ventana.</div>';
        successView.style.display='block';
      }
      function onFail(err){
        loadingView.style.display='none';
        formView.style.display='block';
        btn.disabled=false;btn.textContent='Enviar solicitud';
        msg.className='err';msg.style.display='block';msg.textContent=(err.message||err);
      }
      f.addEventListener('submit',function(e){
        e.preventDefault();
        msg.style.display='none';
        btn.disabled=true;btn.textContent='Subiendo…';
        formView.style.display='none';
        loadingView.style.display='block';
        google.script.run.withSuccessHandler(onSuccess).withFailureHandler(onFail).submitIntake(f);
      });
    </script></body></html>`;
}

function submitIntake(form) {
  return guard_('submitIntake', function () {
    if (!acceptingApplications_()) throw new Error('En este momento no estamos aceptando nuevas solicitudes de préstamo.');
    const ss = getSS_(), nb = ss.getSheetByName(CFG.SHEETS.NEW) || setupNew_(ss) || ss.getSheetByName(CFG.SHEETS.NEW);
    const name = String(form.fullName || '').trim(), email = String(form.email || '').trim(),
      dni = String(form.dni || '').trim(), phone = String(form.phone || '').trim(),
      amount = Number(String(form.amount || '').replace(/\D/g, '')) || '',
      term = Number(form.term) || '', notes = String(form.notes || '').trim();
    if (!name) throw new Error('El nombre completo es obligatorio.');
    if (!email) throw new Error('El correo electrónico es obligatorio.');
    if (!dni) throw new Error('El DNI es obligatorio.');
    if (!phone) throw new Error('El teléfono es obligatorio.');
    if (!amount) throw new Error('Ingrese un monto de préstamo válido.');
    if (term !== 1 && term !== 2 && term !== 15) throw new Error('Elija un plazo de 15 días, 1 mes o 2 meses.');
    if (!notes) throw new Error('Las notas / motivo son obligatorias.');
    if (!form.agree) throw new Error('Debe aceptar los Términos y Condiciones para enviar la solicitud.');
    if (!hasFile_(form.dniPhoto)) throw new Error('Adjunte la foto del DNI.');
    if (!hasFile_(form.cuilPhoto)) throw new Error('Adjunte la foto del CUIL.');
    const rate = termRate_(term), interest = round2_(amount * rate), total = round2_(amount * (1 + rate));
    const folder = borrowerFolder_(borrowerFolderName_(name, dni));
    const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');
    const dniUrl = savePhoto_(folder, form.dniPhoto, name + ' DNI ' + stamp);
    const cuilUrl = savePhoto_(folder, form.cuilPhoto, name + ' CUIL ' + stamp);
    const row = nbAppend_(nb, [new Date(), name, email, dni, phone, amount, term, notes, dniUrl, cuilUrl, false, false, 'Pendiente de verificación']);
    nb.getRange(row, 1).setNumberFormat('yyyy-mm-dd hh:mm'); nb.getRange(row, 6).setNumberFormat(CFG.CURRENCY_FMT);
    try {
      sendBrandedEmail_(email, 'Solicitud recibida',
        'Estimado/a ' + name + ', recibimos su solicitud de préstamo por ' + fmtMoney_(amount) + '. La revisaremos y nos pondremos en contacto.',
        '<p>Estimado/a ' + esc_(name) + ',</p><p>Recibimos su solicitud de préstamo por <b>' + fmtMoney_(amount) +
        '</b> a ' + termLabel_(term) + '. La revisaremos y nos pondremos en contacto a la brevedad.</p>');
    } catch (e) { logError_('submitIntake:email', e); }
    // Aviso al prestamista (Correo del Prestamista) con los datos de la nueva solicitud.
    try {
      const recipients = noticeRecipients_();
      if (recipients.length) {
        const avail = fundAvailable_();
        sendBrandedEmail_(recipients.join(','), 'Nueva solicitud de préstamo — ' + name,
          'Nueva solicitud de ' + name + ' por ' + fmtMoney_(amount) + ' a ' + termLabel_(term) + '. Efectivo disponible: ' + fmtMoney_(avail) + '.',
          '<p>Se recibió una <b>nueva solicitud de préstamo</b>.</p>' +
          '<table style="border-collapse:collapse">' +
          row_('Nombre completo', esc_(name)) + row_('Correo', esc_(email)) + row_('DNI', esc_(dni)) +
          row_('Teléfono', esc_(phone)) + row_('Monto solicitado', '<b>' + fmtMoney_(amount) + '</b>') +
          row_('Plazo', termLabel_(term) + ' (' + termRatePct_(term) + ')') +
          row_('Interés', fmtMoney_(interest)) + row_('Total a devolver', fmtMoney_(total)) +
          row_('Notas / Motivo', esc_(notes)) +
          row_('Foto del frente del DNI', dniUrl ? '<a href="' + dniUrl + '">Ver documento</a>' : '—') +
          row_('Foto del dorso del DNI', cuilUrl ? '<a href="' + cuilUrl + '">Ver documento</a>' : '—') +
          row_('Efectivo disponible para prestar', (avail < amount ? '<b style="color:#900">' : '<b>') + fmtMoney_(avail) + '</b>' + (avail < amount ? ' ⚠ fondo insuficiente' : '')) +
          '</table>' +
          '<p>Revísela en la hoja <b>Nuevos Prestatarios</b> y tilde <b>Verificado?</b> para aprobar o <b>Rechazar?</b> para archivar.</p>');
      }
    } catch (e) { logError_('submitIntake:lenderEmail', e); }
    return {
      ok: true,
      name: name,
      email: email,
      dni: dni,
      phone: phone,
      term: term,
      termLabel: termLabel_(term),
      ratePct: termRatePct_(term),
      amountFmt: fmtMoney_(amount),
      interestFmt: fmtMoney_(interest),
      totalFmt: fmtMoney_(total),
      lenderName: companyName_(),
      message: '¡Solicitud recibida! Gracias, ' + name + '. Le enviamos un correo de confirmación.',
    };
  });
}

function balancePageHtml_() {
  return `<!DOCTYPE html><html lang="es"><head><base target="_top"><style>
    body{font-family:Arial,sans-serif;background:#f4f6fa;margin:0;padding:24px}
    .card{max-width:460px;margin:0 auto;background:#fff;border-radius:10px;padding:24px 28px;box-shadow:0 2px 12px rgba(0,0,0,.08)}
    h1{font-size:20px;color:#674ea7}label{font-weight:bold;display:block;margin:12px 0 4px}
    input{width:100%;box-sizing:border-box;padding:9px;border:1px solid #ccc;border-radius:6px}
    button{margin-top:16px;width:100%;background:#674ea7;color:#fff;border:0;border-radius:6px;padding:12px;font-weight:bold;cursor:pointer}
    #out{margin-top:16px;padding:12px;border-radius:6px;display:none}.ok{background:#e9e2f5}.err{background:#f4cccc}</style></head><body>
    <div class="card">${brandHeaderHtml_()}<h1>Consulta de Saldo</h1>
    <label>ID Préstamo</label><input id="loan" placeholder="L-0001">
    <label>DNI</label><input id="dni">
    <button onclick="go()">Consultar</button><div id="out"></div>
    <script>
      var out=document.getElementById('out');
      function go(){out.style.display='block';out.className='ok';out.textContent='Consultando…';
        google.script.run.withSuccessHandler(function(r){out.className=r.ok?'ok':'err';out.innerHTML=r.html;})
          .withFailureHandler(function(e){out.className='err';out.textContent=(e.message||e);})
          .lookupBalance(document.getElementById('loan').value,document.getElementById('dni').value);}
    </script></div></body></html>`;
}
function lookupBalance(loanId, dni) {
  return guard_('lookupBalance', function () {
    const loan = findLoanById_(String(loanId || '').trim());
    if (!loan || String(loan.dni).replace(/\D/g, '') !== String(dni || '').replace(/\D/g, ''))
      return { ok: false, html: 'Datos no encontrados. Verifique el ID y el DNI.' };
    const rate = termRate_(loan.term);
    const out = computeOutstanding_(loan.principal, rate, loan.loanDate, loanPayments_(loan.loanId), new Date(), new Date(9999, 0, 1));
    return {
      ok: true, html: '<b>' + esc_(loan.name) + '</b><br>Préstamo: ' + esc_(loan.loanId) +
        '<br>Total a pagar: ' + fmtMoney_(loan.totalDue) + '<br>Vencimiento: ' + fmtDate_(loan.dueDate) +
        '<br><b>Saldo pendiente: ' + fmtMoney_(out) + '</b>'
    };
  });
}

/* ===================== FIRMA VIRTUAL DEL CONTRATO ===================== */
/** Ubica la fila de un préstamo en "Prestatarios". Devuelve {sh,row} o null. */
function signStatusRow_(loanId) {
  const sh = getSS_().getSheetByName(CFG.SHEETS.BORROWERS), last = sh.getLastRow();
  if (last < 2) return null;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]).trim() === loanId) return { sh: sh, row: i + 2 };
  return null;
}

/** Resumen de términos del contrato (HTML seguro) para la página de firma. */
function contractTermsHtml_(loan) {
  return '<table class="terms">' +
    '<tr><td class="k">Prestatario</td><td>' + esc_(loan.name) + ' · DNI ' + esc_(loan.dni) + '</td></tr>' +
    '<tr><td class="k">Préstamo</td><td>' + esc_(loan.loanId) + '</td></tr>' +
    '<tr><td class="k">Capital</td><td>' + fmtMoney_(loan.principal) + '</td></tr>' +
    '<tr><td class="k">Plazo</td><td>' + termLabel_(loan.term) + ' (' + termRatePct_(loan.term) + ')</td></tr>' +
    '<tr><td class="k">Interés</td><td>' + fmtMoney_(loan.interest) + '</td></tr>' +
    '<tr><td class="k">Total a pagar</td><td><b>' + fmtMoney_(loan.totalDue) + '</b></td></tr>' +
    '<tr><td class="k">Vencimiento</td><td><b>' + fmtDate_(loan.dueDate) + '</b></td></tr>' +
    '</table>';
}

/** Verifica identidad (ID + DNI) y devuelve los términos a firmar, o el estado "ya firmado". */
function getContractForSigning(loanId, dni) {
  return guard_('getContractForSigning', function () {
    loanId = String(loanId || '').trim();
    const loan = findLoanById_(loanId);
    if (!loan || String(loan.dni).replace(/\D/g, '') !== String(dni || '').replace(/\D/g, ''))
      return { ok: false, msg: 'Datos no encontrados. Verifique el ID del préstamo y su DNI.' };
    const loc = signStatusRow_(loanId);
    const status = loc ? String(loc.sh.getRange(loc.row, COL_SIGN_STATUS).getValue() || '').toUpperCase() : '';
    if (status === SIGN.SIGNED)
      return { ok: true, signed: true, msg: 'Este contrato ya fue firmado. Le enviamos una copia por correo.' };
    return { ok: true, signed: false, name: loan.name, contractHtml: contractTermsHtml_(loan) };
  });
}

/** Recibe la firma (dibujada o imagen), genera el PDF firmado, guarda y avisa a ambas partes. */
function submitSignature(loanId, dni, signatureDataUrl, agree, meta) {
  return guard_('submitSignature', function () {
    loanId = String(loanId || '').trim();
    const loan = findLoanById_(loanId);
    if (!loan || String(loan.dni).replace(/\D/g, '') !== String(dni || '').replace(/\D/g, ''))
      throw new Error('Datos no encontrados. Verifique el ID del préstamo y su DNI.');
    if (agree !== true) throw new Error('Debe aceptar los términos del contrato para firmar.');
    const m = /^data:(image\/[a-z0-9.+-]+);base64,(.+)$/i.exec(String(signatureDataUrl || ''));
    if (!m) throw new Error('Falta la firma. Dibuje su firma o cargue una imagen.');
    const loc = signStatusRow_(loanId);
    if (!loc) throw new Error('No se encontró el préstamo.');
    if (String(loc.sh.getRange(loc.row, COL_SIGN_STATUS).getValue() || '').toUpperCase() === SIGN.SIGNED)
      return { ok: true, message: 'Este contrato ya estaba firmado. Le enviamos una copia por correo.' };

    const mime = m[1], ext = mime.split('/')[1].replace('+xml', '').replace('jpeg', 'jpg').replace('svg', 'png');
    const bytes = Utilities.base64Decode(m[2]);
    const method = (meta && meta.method === 'upload') ? 'upload' : 'draw';
    const now = new Date();
    const folder = borrowerFolder_(borrowerFolderName_(loan.name, loan.dni));
    folder.createFile(Utilities.newBlob(bytes, mime, 'Firma ' + loanId + '.' + ext));

    // El reloj de interés arranca al firmar (Fecha del Préstamo = hoy).
    loc.sh.getRange(loc.row, 9).setValue(now).setNumberFormat('yyyy-mm-dd');
    writeOutstandingRow_(loc.sh, loc.row);
    const freshLoan = readLoan_(loc.sh, loc.row);

    const sig = { dataUrl: signatureDataUrl, signedAt: now, method: method, userAgent: (meta && meta.userAgent) || '' };
    const signedFile = makeSignedAgreementFile_(freshLoan, sig);

    loc.sh.getRange(loc.row, COL_SIGN_STATUS).setValue(SIGN.SIGNED);
    loc.sh.getRange(loc.row, COL_SIGN_DATE).setValue(now).setNumberFormat('yyyy-mm-dd hh:mm');
    loc.sh.getRange(loc.row, COL_SIGN_PDF).setFormula('=HYPERLINK("' + signedFile.getUrl() + '","Ver contrato firmado")');

    try {
      setupFirmas_(getSS_()).appendRow([now, loanId, loan.name, loan.dni, loan.email,
        method === 'upload' ? 'Imagen cargada' : 'Dibujada en pantalla',
        String((meta && meta.userAgent) || '').slice(0, 250), signedFile.getUrl()]);
    } catch (e) { logError_('submitSignature:firmas', e); }
    try { emailSignedContract_(freshLoan, signedFile); } catch (e) { logError_('submitSignature:email', e); }
    try { getSS_().toast('Contrato firmado: ' + loan.name + ' → ' + loanId, 'Firma recibida ✍', 6); } catch (e) {}
    return { ok: true, message: '¡Contrato firmado! Gracias, ' + loan.name + '. Le enviamos una copia por correo.' };
  });
}

/** Enlace personal de firma para un préstamo. */
/**
 * URL base de la app web para los enlaces que se envían a clientes.
 * Prioridad: 1) "URL de la app web (enlaces a clientes)" en Configuración,
 * 2) propiedad WEBAPP_URL (menú ⑥), 3) ScriptApp.getService().getUrl().
 * Se limpian parámetros/fragmentos para poder anexar los propios.
 */
function webAppBaseUrl_() {
  let url = String(getSetting_('URL de la app web (enlaces a clientes)') || '').trim();
  if (!url) url = String(PropertiesService.getScriptProperties().getProperty('WEBAPP_URL') || '');
  if (!url) { try { url = ScriptApp.getService().getUrl(); } catch (e) { url = ''; } }
  return url ? url.replace(/[?#].*$/, '') : '';
}
function signingLink_(loanId) {
  const base = webAppBaseUrl_();
  return base ? base + '?page=firmar&loan=' + encodeURIComponent(loanId) : '';
}

/** Correo al prestatario pidiéndole que firme (con enlace) + copia sin firmar para revisión. */
function emailSigningRequest_(loan, file) {
  if (!loan.email) return;
  const link = signingLink_(loan.loanId);
  const btn = link
    ? '<p style="margin:18px 0"><a href="' + link + '" style="background:#1c4587;color:#fff;text-decoration:none;padding:12px 22px;border-radius:6px;font-weight:bold">Revisar y firmar mi contrato</a></p>' +
      '<p style="font-size:12px;color:#888">O copie este enlace: ' + esc_(link) + '</p>'
    : '<p><b>Su contrato está listo para firmar.</b> Solicite el enlace de firma al prestamista.</p>';
  sendBrandedEmail_(loan.email, 'Revise y firme su contrato — ' + loan.loanId,
    'Estimado/a ' + loan.name + ', su contrato ' + loan.loanId + ' está listo para firmar. ' + (link ? ('Ingrese a: ' + link) : ''),
    '<p>Estimado/a ' + esc_(loan.name) + ',</p><p>Su solicitud fue aprobada. Antes de recibir el préstamo debe <b>firmar el contrato</b> <b>' + esc_(loan.loanId) + '</b>.</p>' +
    '<ul><li>Capital: <b>' + fmtMoney_(loan.principal) + '</b></li><li>Total a pagar: <b>' + fmtMoney_(loan.totalDue) + '</b></li><li>Vencimiento: <b>' + fmtDate_(loan.dueDate) + '</b></li></ul>' +
    btn + '<p style="font-size:12px;color:#888">Para firmar necesitará su DNI (' + esc_(loan.dni) + '). Se adjunta una copia del contrato para su revisión.</p>',
    file ? { attachments: [file.getAs('application/pdf')] } : {});
}

/** Correo con el PDF firmado, al prestatario y a Impulso Crédito (prestamista). */
function emailSignedContract_(loan, file) {
  const pdf = file.getAs('application/pdf');
  if (loan.email) {
    sendBrandedEmail_(loan.email, 'Contrato firmado — ' + loan.loanId,
      'Estimado/a ' + loan.name + ', adjuntamos su contrato de préstamo ' + loan.loanId + ' ya firmado.',
      '<p>Estimado/a ' + esc_(loan.name) + ',</p><p>Gracias. Adjuntamos su <b>contrato firmado</b> <b>' + esc_(loan.loanId) + '</b>.</p>' +
      '<ul><li>Capital: <b>' + fmtMoney_(loan.principal) + '</b></li><li>Total a pagar: <b>' + fmtMoney_(loan.totalDue) + '</b></li>' +
      '<li>Vencimiento: <b>' + fmtDate_(loan.dueDate) + '</b></li></ul>', { attachments: [pdf] });
  }
  const rec = noticeRecipients_();
  if (rec.length) {
    sendBrandedEmail_(rec.join(','), 'Contrato firmado — ' + loan.loanId + ' — ' + loan.name,
      loan.name + ' firmó el contrato ' + loan.loanId + '. Ya puede desembolsar.',
      '<p>El prestatario <b>' + esc_(loan.name) + '</b> (DNI ' + esc_(loan.dni) + ') <b>firmó</b> el contrato <b>' + esc_(loan.loanId) + '</b>.</p>' +
      '<p>Fecha de firma: <b>' + fmtDateTime_(new Date()) + '</b>. <b>Ya puede desembolsar.</b></p>' +
      '<p>Se adjunta el contrato firmado para su archivo.</p>', { attachments: [pdf] });
  }
}

/** Menú/panel: reenvía el enlace de firma para el préstamo de la fila seleccionada en "Prestatarios". */
function resendSigningLink_() {
  const ss = getSS_(), sh = ss.getActiveSheet(), ui = SpreadsheetApp.getUi();
  if (sh.getName() !== CFG.SHEETS.BORROWERS) { ui.alert('Abra la hoja "Prestatarios", seleccione la fila del préstamo y vuelva a usar esta opción.'); return; }
  const row = sh.getActiveRange().getRow();
  if (row < 2) { ui.alert('Seleccione la fila de un préstamo.'); return; }
  const loanId = String(sh.getRange(row, 1).getValue()).trim();
  if (!loanId) { ui.alert('La fila seleccionada no tiene un préstamo.'); return; }
  const loan = readLoan_(sh, row);
  if (!loan.email) { ui.alert('El préstamo no tiene un correo cargado.'); return; }
  const st = String(sh.getRange(row, COL_SIGN_STATUS).getValue() || '').toUpperCase();
  if (st === SIGN.SIGNED && ui.alert('Este contrato ya figura FIRMADO. ¿Reenviar el enlace igualmente?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  try {
    if (st !== SIGN.SIGNED) sh.getRange(row, COL_SIGN_STATUS).setValue(SIGN.PENDING);
    let file = null; try { file = makeAgreementFile_(loan); } catch (e) { logError_('resendSigningLink_:pdf', e); }
    emailSigningRequest_(loan, file);
    ss.toast('Enlace de firma reenviado a ' + loan.email, 'Listo', 5);
  } catch (err) { logError_('resendSigningLink_', err); ui.alert('No se pudo reenviar: ' + (err.message || err)); }
}

/** Página web independiente para firmar el contrato (?page=firmar&loan=L-0001). */
function signingPageHtml_(loanId) {
  return `<!DOCTYPE html><html lang="es"><head><base target="_top"><meta charset="utf-8"><style>
    body{font-family:Arial,sans-serif;background:#f4f6fa;margin:0;padding:24px;color:#222}
    .card{max-width:560px;margin:0 auto;background:#fff;border-radius:10px;padding:24px 28px;box-shadow:0 2px 12px rgba(0,0,0,.08)}
    h1{font-size:20px;color:#1c4587;margin:6px 0}
    label{font-weight:bold;display:block;margin:12px 0 4px}
    input[type=text]{width:100%;box-sizing:border-box;padding:9px;border:1px solid #ccc;border-radius:6px}
    button{margin-top:8px;background:#1c4587;color:#fff;border:0;border-radius:6px;padding:12px 18px;font-weight:bold;cursor:pointer}
    button.sec{background:#eee;color:#333}
    .terms{width:100%;border-collapse:collapse;margin:8px 0 4px}
    .terms td{padding:6px 8px;border:1px solid #ddd;font-size:14px}
    .terms td.k{background:#f0f4fa;font-weight:bold;width:42%}
    .tabs{margin:14px 0 6px}.tabs button{margin-right:6px;background:#eee;color:#333}
    .tabs button.on{background:#1c4587;color:#fff}
    canvas{border:1px dashed #999;border-radius:6px;width:100%;height:180px;touch-action:none;background:#fff}
    #prev{max-width:100%;max-height:180px;border:1px solid #ddd;border-radius:6px;margin-top:8px;display:none}
    .hint{font-size:12px;color:#888}.consent{margin:14px 0}.consent input{transform:scale(1.3);margin-right:8px}
    #out{margin-top:14px;padding:12px;border-radius:6px;display:none}.ok{background:#e6f4ea}.err{background:#f4cccc}
    #loadingView{text-align:center;padding:34px 0}
    .spinner{width:56px;height:56px;margin:0 auto 18px;border:6px solid #e3e9f4;border-top-color:#1c4587;border-radius:50%;animation:spin .9s linear infinite}
    @keyframes spin{to{transform:rotate(360deg)}}
    .check{width:72px;height:72px;line-height:72px;margin:6px auto 10px;border-radius:50%;background:#38761d;color:#fff;font-size:40px;font-weight:bold;text-align:center}
    .sub{color:#555;margin:0 0 8px;line-height:1.45}
    .hide{display:none}</style></head><body>
    <div class="card">${brandHeaderHtml_()}<h1>Firmar Contrato</h1>

    <div id="gate">
      <p class="hint">Ingrese su DNI para ver y firmar el contrato <b>${esc_(loanId)}</b>.</p>
      <label>ID Préstamo</label><input type="text" id="loan" value="${esc_(loanId)}">
      <label>DNI</label><input type="text" id="dni" placeholder="Su DNI">
      <button onclick="unlock()">Ver contrato</button>
      <div id="gerr" class="err" style="display:none;margin-top:12px;padding:12px;border-radius:6px"></div>
    </div>

    <div id="signview" class="hide">
      <div id="terms"></div>
      <div class="tabs">
        <button id="tabDraw" class="on" onclick="showTab('draw')">✍ Dibujar firma</button>
        <button id="tabUp" onclick="showTab('upload')">🖼 Subir imagen</button>
      </div>
      <div id="paneDraw">
        <canvas id="cv"></canvas>
        <button class="sec" onclick="clearCanvas()">Borrar</button>
      </div>
      <div id="paneUp" class="hide">
        <input type="file" id="file" accept="image/*">
        <div class="hint">Tome una foto o cargue una imagen de su firma (fondo claro).</div>
        <img id="prev">
      </div>
      <label class="consent"><input type="checkbox" id="agree">He leído y acepto los términos de este contrato y lo firmo electrónicamente.</label>
      <button id="go" onclick="sign()">Firmar y enviar</button>
      <div id="out"></div>
    </div>

    <div id="loadingView" class="hide">
      <div class="spinner"></div>
      <p style="font-weight:bold;color:#1c4587;margin:0">Firmando su contrato…</p>
      <p class="hint" style="margin-top:6px">Esto puede tardar unos segundos. No cierre esta ventana.</p>
    </div>

    <div id="done" class="hide" style="text-align:center;padding:8px 0"></div>

    <script>
      var LOAN_ID = ${JSON.stringify(String(loanId))};
      var DNI = '', TAB = 'draw', uploaded = '', drew = false, cv, ctx, drawing = false;
      function $(id){return document.getElementById(id);}
      function unlock(){
        DNI = $('dni').value; $('gerr').style.display='none';
        if(!$('loan').value || !DNI){ gerr('Ingrese el ID y el DNI.'); return; }
        google.script.run.withSuccessHandler(function(r){
          if(!r || !r.ok){ gerr((r&&r.msg)||'No se pudo verificar.'); return; }
          if(r.signed){ $('gate').classList.add('hide'); showDone(r.msg); return; }
          $('terms').innerHTML = r.contractHtml;
          $('gate').classList.add('hide'); $('signview').classList.remove('hide'); setupCanvas();
        }).withFailureHandler(function(e){ gerr(e.message||e); }).getContractForSigning($('loan').value, DNI);
      }
      function gerr(m){ var g=$('gerr'); g.textContent=m; g.style.display='block'; }
      function showTab(t){ TAB=t;
        $('tabDraw').className = t==='draw'?'on':''; $('tabUp').className = t==='upload'?'on':'';
        $('paneDraw').className = t==='draw'?'':'hide'; $('paneUp').className = t==='upload'?'':'hide';
        if(t==='draw') setupCanvas();
      }
      function setupCanvas(){
        cv=$('cv'); if(!cv) return; cv.width=cv.clientWidth; cv.height=180;
        ctx=cv.getContext('2d'); ctx.lineWidth=2.5; ctx.lineCap='round'; ctx.strokeStyle='#111';
        cv.onpointerdown=function(e){ drawing=true; drew=true; var p=pos(e); ctx.beginPath(); ctx.moveTo(p.x,p.y); cv.setPointerCapture(e.pointerId); };
        cv.onpointermove=function(e){ if(!drawing) return; var p=pos(e); ctx.lineTo(p.x,p.y); ctx.stroke(); };
        cv.onpointerup=function(){ drawing=false; }; cv.onpointerleave=function(){ drawing=false; };
      }
      function pos(e){ var r=cv.getBoundingClientRect(); return {x:e.clientX-r.left, y:e.clientY-r.top}; }
      function clearCanvas(){ if(ctx){ ctx.clearRect(0,0,cv.width,cv.height); drew=false; } }
      $('file') && $('file').addEventListener('change', function(ev){
        var f=ev.target.files[0]; if(!f) return; var rd=new FileReader();
        rd.onload=function(){ var img=new Image(); img.onload=function(){
          var max=1000, s=Math.min(1, max/img.width), w=Math.round(img.width*s), h=Math.round(img.height*s);
          var c=document.createElement('canvas'); c.width=w; c.height=h;
          c.getContext('2d').drawImage(img,0,0,w,h);
          uploaded=c.toDataURL('image/jpeg',0.85);
          var pv=$('prev'); pv.src=uploaded; pv.style.display='block';
        }; img.src=rd.result; };
        rd.readAsDataURL(f);
      });
      function sign(){
        var out=$('out'); out.style.display='none';
        if(!$('agree').checked){ msg('Debe aceptar los términos para firmar.', true); return; }
        var data='';
        if(TAB==='draw'){ if(!drew){ msg('Dibuje su firma primero.', true); return; } data=cv.toDataURL('image/png'); }
        else { if(!uploaded){ msg('Cargue una imagen de su firma primero.', true); return; } data=uploaded; }
        $('go').disabled=true;
        $('signview').classList.add('hide'); $('loadingView').classList.remove('hide');
        google.script.run.withSuccessHandler(function(r){
          $('loadingView').classList.add('hide');
          if(r && r.ok){ showDone(r.message); }
          else { $('signview').classList.remove('hide'); $('go').disabled=false; msg((r&&r.message)||'No se pudo firmar.', true); }
        }).withFailureHandler(function(e){
          $('loadingView').classList.add('hide'); $('signview').classList.remove('hide'); $('go').disabled=false; msg(e.message||e, true);
        }).submitSignature(LOAN_ID, DNI, data, true, { userAgent: navigator.userAgent, method: TAB });
      }
      function msg(m, isErr){ var o=$('out'); o.className=isErr?'err':'ok'; o.textContent=m; o.style.display='block'; }
      function esc(s){ return String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }
      function showDone(m){
        var d=$('done'); d.classList.remove('hide');
        d.innerHTML='<div class="check">✓</div>'+
          '<h1 style="color:#38761d">¡Contrato firmado!</h1>'+
          '<p class="sub">'+esc(m||'Su contrato fue firmado correctamente.')+'</p>'+
          '<p class="sub"><b>Revise su correo:</b> le enviamos una copia de su contrato firmado (PDF).</p>'+
          '<p class="hint">Ya puede cerrar esta ventana.</p>';
      }
    </script></div></body></html>`;
}

/* ===================== VERIFICAR / RECHAZAR (onEdit) ===================== */
function onEditInstallable(e) {
  try {
    const sh = e.range.getSheet(), name = sh.getName();
    const c0 = e.range.getColumn(), cN = c0 + e.range.getNumColumns() - 1;
    const r0 = Math.max(e.range.getRow(), 2), rN = e.range.getRow() + e.range.getNumRows() - 1;
    if (name === CFG.SHEETS.PAYMENTS) {
      if ((c0 <= 2 && cN >= 2) || (c0 <= 6 && cN >= 6) || (c0 <= 7 && cN >= 7)) {
        const ids = {};
        for (let row = r0; row <= rN; row++) { fillPaymentLookup_(sh, row); const id = String(sh.getRange(row, 2).getValue()).trim(); if (id) ids[id] = true; }
        Object.keys(ids).forEach(id => { recalcLoanPayments_(id); updateLoanOutstanding_(id); });
      }
    } else if (name === CFG.SHEETS.BORROWERS) {
      if ([1, 6, 7, 9].some(c => c >= c0 && c <= cN)) for (let row = r0; row <= rN; row++) writeOutstandingRow_(sh, row);
    } else if (name === CFG.SHEETS.AGREEMENT) {
      const a1 = e.range.getA1Notation();
      if (a1 === 'B15' && sh.getRange('B15').getValue() === true) { sh.getRange('B15').setValue(false); contractStudioPreview_(sh); }
      else if (a1 === 'B16' && sh.getRange('B16').getValue() === true) { sh.getRange('B16').setValue(false); contractStudioSend_(sh); }
    } else if (name === CFG.SHEETS.STATEMENTS) {
      const a1 = e.range.getA1Notation();
      if (a1 === 'B18' && sh.getRange('B18').getValue() === true) { sh.getRange('B18').setValue(false); statementStudioPreview_(sh); }
      else if (a1 === 'B19' && sh.getRange('B19').getValue() === true) { sh.getRange('B19').setValue(false); statementStudioSend_(sh); }
    } else if (name === CFG.SHEETS.REMINDER) {
      if (e.range.getA1Notation() === 'B2' && sh.getRange('B2').getValue() === true) {
        sh.getRange('B2').setValue(false); rebuildRemindersSheet_(getSS_());
      } else if ((REM_GEN_COL >= c0 && REM_GEN_COL <= cN) || (REM_SEND_COL >= c0 && REM_SEND_COL <= cN)) {
        for (let row = rN; row >= Math.max(r0, REM_DATA_START); row--) {
          if (sh.getRange(row, 1).getValue() === '') continue;
          if (sh.getRange(row, REM_GEN_COL).getValue() === true) { sh.getRange(row, REM_GEN_COL).setValue(false); reminderRowGenerate_(sh, row); }
          else if (sh.getRange(row, REM_SEND_COL).getValue() === true) { sh.getRange(row, REM_SEND_COL).setValue(false); reminderRowSend_(sh, row); }
        }
      }
    } else if (name === CFG.SHEETS.LATE) {
      if (LATE_SEND_COL >= c0 && LATE_SEND_COL <= cN) {
        for (let row = rN; row >= r0; row--) {
          if (sh.getRange(row, LATE_SEND_COL).getValue() !== true) continue;
          sh.getRange(row, LATE_SEND_COL).setValue(false);
          const loanId = String(sh.getRange(row, 1).getValue()).trim();
          if (!loanId) continue;
          try {
            const res = sendOverdueNotice_(loanId);
            sh.getRange(row, LATE_SENT_COL).setValue(res.date).setNumberFormat('yyyy-mm-dd');
            getSS_().toast(res.message, 'Aviso de mora', 5);
          } catch (err) { logError_('onEdit:LATE', err); getSS_().toast(err.message || String(err), '⚠ No se pudo enviar', 8); }
        }
      }
    } else if (name === CFG.SHEETS.NEW) {
      const m = headerMap_(sh), verCol = m['Verificado?'], rejCol = m['Rechazar?'], bcraCol = m['Verificar BCRA?'];
      const acts = [];
      for (let row = r0; row <= rN; row++) {
        if (bcraCol && bcraCol >= c0 && bcraCol <= cN && sh.getRange(row, bcraCol).getValue() === true) { bcraCheckRow_(sh, row); continue; }
        if (verCol && verCol >= c0 && verCol <= cN && sh.getRange(row, verCol).getValue() === true) acts.push({ row: row, kind: 'v' });
        else if (rejCol && rejCol >= c0 && rejCol <= cN && sh.getRange(row, rejCol).getValue() === true) acts.push({ row: row, kind: 'r' });
      }
      acts.sort((a, b) => b.row - a.row).forEach(a => a.kind === 'v' ? verifyApplicant_(sh, a.row) : rejectApplicant_(sh, a.row));
    }
  } catch (err) { logError_('onEditInstallable', err); }
}

function headerMap_(sh) {
  const h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0], m = {};
  h.forEach((x, i) => { const k = String(x).trim(); if (k) m[k] = i + 1; });
  return m;
}

function verifyApplicant_(nb, row) {
  const m = headerMap_(nb);
  const g = t => m[t] ? nb.getRange(row, m[t]).getValue() : '';
  const resCol = m['Resultado'], verCol = m['Verificado?'];
  const fail = msg => { if (resCol) nb.getRange(row, resCol).setValue(msg); if (verCol) nb.getRange(row, verCol).setValue(false); };
  const name = String(g('Nombre completo')).trim(), email = g('Correo'), dni = String(g('DNI')).trim(),
    phone = g('Teléfono'), amount = Number(g('Monto Solicitado')) || 0, term = Number(g('Plazo (meses)')) || 0,
    natId = g('Foto del frente del DNI') || g('Foto DNI'), workId = g('Foto del dorso del DNI') || g('Foto CUIL');
  if (!name || !dni) { fail('⚠ Falta nombre o DNI'); return; }
  if (term !== 1 && term !== 2 && term !== 15) { fail('⚠ Plazo debe ser 15, 1 o 2'); return; }
  if (!amount) { fail('⚠ Monto inválido'); return; }
  if (!natId || !workId) { fail('⚠ Faltan fotos (DNI/CUIL)'); return; }
  // Efectivo disponible ANTES de colocar este préstamo (baja al prestar el capital).
  const availBefore = fundAvailable_();
  // Duplicado por DNI
  const bs = getSS_().getSheetByName(CFG.SHEETS.BORROWERS);
  const existing = bs.getRange(2, 4, Math.max(bs.getLastRow() - 1, 1), 1).getValues()
    .some(r => String(r[0]).replace(/\D/g, '') === dni.replace(/\D/g, '') && dni.replace(/\D/g, '') !== '');
  if (existing) fail('⚠ DNI ya existe (préstamo previo) — revisar antes de aprobar');
  const loanId = nextLoanId_(), tRow = firstEmptyBorrowerRow_(bs);
  bs.getRange(tRow, 1, 1, 7).setValues([[loanId, name, email, dni, phone, amount, term]]);
  bs.getRange(tRow, 9).setValue(new Date()); bs.getRange(tRow, 9).setNumberFormat('yyyy-mm-dd');
  writeOutstandingRow_(bs, tRow);
  const folder = borrowerFolder_(borrowerFolderName_(name, dni));
  [natId, workId].forEach(u => extractDriveIds_(u).forEach(id => { try { DriveApp.getFileById(id).moveTo(folder); } catch (e) {} }));
  bs.getRange(tRow, 4).setNote('📁 Carpeta: ' + folder.getUrl() + (natId ? '\nDNI: ' + natId : '') + (workId ? '\nCUIL: ' + workId : ''));
  // Firma obligatoria: el préstamo queda PENDIENTE DE FIRMA hasta que el prestatario firme.
  bs.getRange(tRow, COL_SIGN_STATUS).setValue(SIGN.PENDING);
  // Contrato (copia de revisión) + correo pidiendo la FIRMA con enlace.
  try {
    const loan = readLoan_(bs, tRow);
    const file = makeAgreementFile_(loan);
    markAgreementSent_(loanId, file);
    if (email) emailSigningRequest_(loan, file);
  } catch (e) { logError_('verifyApplicant_:contrato', e); }
  nb.deleteRow(row);
  const shortfall = amount > availBefore;
  const availAfter = round2_(availBefore - amount);
  try {
    getSS_().toast(
      (shortfall ? '⚠ Fondo insuficiente. ' : '') + 'Pendiente de firma. Disponible tras prestar: ' + fmtMoney_(availAfter),
      (shortfall ? '⚠ ' : '') + 'Aprobado ' + name + ' → ' + loanId + ' (firma pendiente)', shortfall ? 8 : 6);
  } catch (e) {}
}

function rejectApplicant_(nb, row) {
  const m = headerMap_(nb);
  const g = t => m[t] ? nb.getRange(row, m[t]).getValue() : '';
  const name = String(g('Nombre completo')).trim(), email = g('Correo'), dni = g('DNI'), phone = g('Teléfono');
  const arch = getSS_().getSheetByName(CFG.SHEETS.REJECTED) || setupRejected_(getSS_());
  arch.appendRow([new Date(), name, email, dni, phone, 'Rechazado desde Nuevos Prestatarios']);
  if (email) {
    try {
      sendBrandedEmail_(email, 'Actualización de su solicitud',
        'Estimado/a ' + name + ', gracias por su solicitud. Lamentamos informarle que no podemos avanzar en este momento.',
        '<p>Estimado/a ' + esc_(name) + ',</p><p>Gracias por su solicitud. Lamentamos informarle que no podemos avanzar en este momento.</p>');
    } catch (e) { logError_('rejectApplicant_:email', e); }
  }
  nb.deleteRow(row);
  try { getSS_().toast('Rechazado ' + name, 'Archivado', 5); } catch (e) {}
}

/* ================== BCRA Central de Deudores ==================
 * Consulta la Central de Deudores del BCRA en el momento de la carga,
 * antes de desembolsar. Acepta DNI o CUIL. Los montos que devuelve la
 * API vienen en MILES de pesos; acá se multiplican por 1000 para mostrar
 * pesos reales. También disponible como fórmula: =BCRA_SITUACION(D2).
 */
var BCRA_BASE = 'https://api.bcra.gob.ar/CentralDeDeudores/v1.0';

/** Dígito verificador AFIP/ARCA estándar. Devuelve [prefijo, dígito]. */
function cuilCheckDigit_(prefix, dni) {
  var body = ('' + prefix) + Utilities.formatString('%08d', Number(dni));
  var w = [5, 4, 3, 2, 7, 6, 5, 4, 3, 2], sum = 0;
  for (var i = 0; i < 10; i++) sum += Number(body.charAt(i)) * w[i];
  var rem = sum % 11;
  if (rem === 0) return [prefix, 0];
  if (rem === 1) return [23, prefix === 20 ? 9 : 4];
  return [prefix, 11 - rem];
}

/** Genera candidatos de CUIL a partir de un DNI (prefijos 20 y 27). */
function cuilCandidates_(dni) {
  dni = String(dni).replace(/\D/g, '');
  var out = [];
  [20, 27].forEach(function (p) {
    var r = cuilCheckDigit_(p, dni);
    var c = Utilities.formatString('%02d', r[0]) +
            Utilities.formatString('%08d', Number(dni)) + r[1];
    if (out.indexOf(c) === -1) out.push(c);
  });
  return out;
}

/** GET a un endpoint del BCRA. Devuelve el objeto parseado, o null en 404. */
function bcraGet_(path) {
  var res = UrlFetchApp.fetch(BCRA_BASE + path, {
    muteHttpExceptions: true,
    followRedirects: true
  });
  var code = res.getResponseCode();
  if (code === 404) return null;
  if (code === 429 || code >= 500) {       // throttled — un reintento lento
    Utilities.sleep(3000);
    res = UrlFetchApp.fetch(BCRA_BASE + path, { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return null;
  } else if (code !== 200) {
    return null;
  }
  try { return JSON.parse(res.getContentText()); } catch (e) { return null; }
}

/** Consulta principal. Acepta un DNI o un CUIL completo. */
function consultarBcra_(dniOrCuil) {
  var raw = String(dniOrCuil).replace(/\D/g, '');
  var cands = raw.length === 11 ? [raw] : cuilCandidates_(raw);

  for (var i = 0; i < cands.length; i++) {
    var d = bcraGet_('/Deudas/' + cands[i]);
    var r = d && (d.results || d);
    if (!r || !r.periodos || !r.periodos.length) continue;

    var ents = r.periodos[0].entidades || [];
    var total = 0, perf = 0, worst = 0, deudas = [];
    ents.forEach(function (e) {
      var m = Number(e.monto || 0) * 1000, s = Number(e.situacion || 0);
      total += m;
      if (s <= 2) perf += m;
      if (s > worst) worst = s;
      deudas.push({ entidad: String(e.entidad || '').trim(), situacion: s, monto: m });
    });
    deudas.sort(function (a, b) { return b.situacion - a.situacion; }); // peor primero

    // Tendencia 24 meses
    var h = bcraGet_('/Deudas/Historicas/' + cands[i]);
    var hr = h && (h.results || h), serie = [];
    if (hr && hr.periodos) {
      hr.periodos.forEach(function (p) {
        var w = 0;
        (p.entidades || []).forEach(function (e) {
          w = Math.max(w, Number(e.situacion || 0));
        });
        serie.push({ p: String(p.periodo), s: w });
      });
      serie.sort(function (a, b) { return a.p < b.p ? -1 : 1; });
    }
    var latest = serie.length ? serie[serie.length - 1].s : worst;
    var sixAgo = serie.length >= 7 ? serie[serie.length - 7].s : latest;

    var share = total ? perf / total : 1;
    var decision =
      (worst >= 4 && share < 0.5) ? 'RECHAZAR' :
      (latest - sixAgo >= 2 || worst >= 4) ? 'REVISAR' : 'OK';

    return {
      cuil: cands[i],
      nombre: r.denominacion || '',
      peorSituacion: worst,
      entidades: ents.length,
      deudaTotal: total,
      deudaPerformnte: perf,
      pctPerformante: Math.round(share * 1000) / 10,
      delta6m: latest - sixAgo,
      historia: serie.map(function (x) { return x.s; }).join(''),
      deudas: deudas,
      decision: decision
    };
  }
  return { cuil: '', nombre: '', peorSituacion: 0, entidades: 0, deudaTotal: 0,
           deudaPerformnte: 0, pctPerformante: 100, delta6m: 0, historia: '',
           deudas: [], decision: 'SIN REGISTRO' };
}

/**
 * Fórmula de hoja:  =BCRA_SITUACION(D2)
 * Devuelve un veredicto de una línea para la celda.
 * @customfunction
 */
function BCRA_SITUACION(dniOrCuil) {
  if (!dniOrCuil) return '';
  var r = consultarBcra_(dniOrCuil);
  if (r.decision === 'SIN REGISTRO') return 'SIN REGISTRO';
  return Utilities.formatString('%s · sit %d · %s%% al día · %s',
    r.decision, r.peorSituacion, r.pctPerformante,
    r.delta6m >= 1 ? 'DETERIORANDO' : (r.delta6m <= -1 ? 'MEJORANDO' : 'ESTABLE'));
}

/**
 * Puntaje de desempeño (0–100) y calificación de riesgo según estándares de
 * clasificación de deudores del BCRA (situación 1–6):
 *   1 Normal · 2 Riesgo bajo/seguimiento especial · 3 Con problemas (medio)
 *   4 Alto riesgo de insolvencia · 5 Irrecuperable · 6 Irrec. por disp. técnica.
 * Devuelve { puntaje, rating }.
 */
function bcraScore_(r) {
  if (!r || r.decision === 'SIN REGISTRO') return { puntaje: 85, rating: 'Sin historial' };
  var worst = Number(r.peorSituacion) || 0;
  var sitBase = ({ 1: 100, 2: 75, 3: 50, 4: 25, 5: 5, 6: 0 })[worst];
  if (sitBase === undefined) sitBase = worst > 6 ? 0 : 100; // sin deuda reportada → normal
  var pct = Number(r.pctPerformante); if (isNaN(pct)) pct = 100;
  var puntaje = Math.round(0.6 * sitBase + 0.4 * pct);
  var delta = Number(r.delta6m) || 0;
  if (delta >= 2) puntaje -= 10;        // deteriorando
  else if (delta <= -2) puntaje += 5;   // mejorando
  puntaje = Math.max(0, Math.min(100, puntaje));

  var rating =
    puntaje >= 80 ? 'Riesgo Bajo (A)' :
    puntaje >= 60 ? 'Riesgo Medio-Bajo (B)' :
    puntaje >= 40 ? 'Riesgo Medio (C)' :
    puntaje >= 20 ? 'Riesgo Alto (D)' : 'Riesgo Muy Alto (E)';
  // Piso por situación: la calificación nunca contradice al BCRA.
  var order = ['Riesgo Bajo (A)', 'Riesgo Medio-Bajo (B)', 'Riesgo Medio (C)', 'Riesgo Alto (D)', 'Riesgo Muy Alto (E)'];
  if (worst >= 5 && order.indexOf(rating) < 4) rating = 'Riesgo Muy Alto (E)';
  else if (worst === 4 && order.indexOf(rating) < 3) rating = 'Riesgo Alto (D)';
  return { puntaje: puntaje, rating: rating };
}

/** Etiqueta de la clasificación de deudores del BCRA (situación 1–6). */
function bcraSituLabel_(s) {
  return ({
    1: 'Normal', 2: 'Riesgo bajo (seguimiento especial)', 3: 'Con problemas (riesgo medio)',
    4: 'Alto riesgo de insolvencia', 5: 'Irrecuperable', 6: 'Irrecuperable por disposición técnica'
  })[Number(s)] || 'Sin clasificación';
}

/**
 * Resumen en lenguaje llano que interpreta cada columna para esta persona:
 * situación, % al día, tendencia, puntaje/calificación y la decisión sugerida.
 */
function bcraResumen_(r, s) {
  if (!r || r.decision === 'SIN REGISTRO') {
    return 'Sin registro en la Central de Deudores del BCRA (no se hallaron deudas). ' +
      'Puntaje ' + s.puntaje + '/100 → ' + s.rating + '. ' +
      'Historial crediticio limitado: verificar identidad y CUIL antes de avanzar.';
  }
  var tend = r.delta6m >= 1 ? 'deteriorando' : (r.delta6m <= -1 ? 'mejorando' : 'estable');
  var decTxt = ({
    'OK': 'situación favorable, puede avanzar',
    'REVISAR': 'revisar manualmente antes de prestar',
    'RECHAZAR': 'no se recomienda prestar'
  })[r.decision] || r.decision;
  var nombre = r.nombre ? r.nombre + ' ' : '';
  var ents = Number(r.entidades) || 0;
  var detalle = (r.deudas && r.deudas.length)
    ? ' Detalle por préstamo: ' + r.deudas.map(function (d, i) {
        return (i + 1) + ') ' + (d.entidad || 'Entidad s/nombre') + ' — sit ' + d.situacion +
          ' (' + bcraSituLabel_(d.situacion) + ')';
      }).join('; ') + '.'
    : '';
  return nombre + '(CUIL ' + r.cuil + '). ' +
    'Tiene ' + ents + ' préstamo' + (ents === 1 ? '' : 's') + ' reportado' + (ents === 1 ? '' : 's') + ' en BCRA. ' +
    'Peor situación ' + r.peorSituacion + ' — ' + bcraSituLabel_(r.peorSituacion) + '. ' +
    r.pctPerformante + '% de la deuda al día' +
    (Number(r.pctPerformante) < 100 ? '; el resto en categorías de mayor riesgo' : '') + '. ' +
    'Tendencia 6 meses: ' + tend + '. ' +
    'Puntaje ' + s.puntaje + '/100 → ' + s.rating + '. ' +
    'Decisión: ' + r.decision + ' (' + decTxt + ').' + detalle;
}

/**
 * Ejecuta la consulta BCRA para una fila de "Nuevos Prestatarios" cuando se
 * tilda "Verificar BCRA?". Usa DNI o CUIL según "Parámetro BCRA".
 */
function bcraCheckRow_(sh, row) {
  const m = headerMap_(sh), bcraCol = m['Verificar BCRA?'];
  const uncheck = () => { if (bcraCol) sh.getRange(row, bcraCol).setValue(false); };
  try {
    const param = String((m['Parámetro BCRA'] ? sh.getRange(row, m['Parámetro BCRA']).getValue() : '') || 'DNI').toUpperCase();
    const dni = String((m['DNI'] ? sh.getRange(row, m['DNI']).getValue() : '') || '').trim();
    const cuil = String((m['CUIL'] ? sh.getRange(row, m['CUIL']).getValue() : '') || '').trim();
    const key = (param === 'CUIL' && cuil) ? cuil : dni;
    if (!String(key).replace(/\D/g, '')) { uncheck(); getSS_().toast('Falta DNI o CUIL en la fila ' + row, '⚠ BCRA', 6); return; }
    const r = consultarBcra_(key);
    const s = bcraScore_(r);
    sh.getRange(row, m['CUIL'], 1, 8).setValues([[
      r.cuil, r.nombre, r.peorSituacion, r.pctPerformante, r.decision, s.puntaje, s.rating, bcraResumen_(r, s)
    ]]);
    uncheck();
    getSS_().toast((r.nombre || key) + ' → ' + r.decision + ' · ' + s.rating + ' · ' + s.puntaje + '/100', 'BCRA', 6);
  } catch (err) {
    logError_('bcraCheckRow_', err);
    uncheck();
    try { getSS_().toast(err.message || String(err), '⚠ BCRA', 8); } catch (e) {}
  }
}

/**
 * Escribe el resultado BCRA en una fila y devuelve true solo si es 'OK'.
 * Bloque de 7 columnas: CUIL, Nombre BCRA, Peor Situación, % al día, Decisión, Puntaje, Calificación.
 */
function verificarSolicitante_(sheet, row, dniCol, outCol) {
  var dni = sheet.getRange(row, dniCol).getValue();
  var r = consultarBcra_(dni);
  var s = bcraScore_(r);
  sheet.getRange(row, outCol, 1, 8).setValues([[
    r.cuil, r.nombre, r.peorSituacion, r.pctPerformante, r.decision, s.puntaje, s.rating, bcraResumen_(r, s)
  ]]);
  return r.decision === 'OK';
}

/** Menú ▸ verifica en BCRA todas las filas de "Prestatarios". */
function backfillPrestatarios() {
  var sh = getSS_().getSheetByName(CFG.SHEETS.BORROWERS);
  if (!sh) return;
  var last = sh.getLastRow();
  var head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var dniCol = head.indexOf('DNI') + 1;
  if (dniCol < 1) { getSS_().toast('No se encontró la columna DNI', '⚠ BCRA', 6); return; }
  var outCol = sh.getLastColumn() + 1;
  sh.getRange(1, outCol, 1, 8)
    .setValues([['CUIL', 'Nombre BCRA', 'Peor Situación', '% al día', 'Decisión', 'Puntaje BCRA', 'Calificación de Riesgo', 'Resumen BCRA']]);
  for (var r = 2; r <= last; r++) {
    if (!sh.getRange(r, dniCol).getValue()) continue;
    verificarSolicitante_(sh, r, dniCol, outCol);
    Utilities.sleep(1500);   // respetar el límite por IP del BCRA
  }
  try { getSS_().toast('Verificación BCRA completada', 'BCRA', 5); } catch (e) {}
}

/* ===================== TAREAS DIARIAS (recordatorios) ===================== */
function dailyTasks() {
  guard_('dailyTasks', function () {
    updateAllOutstanding_();
    rebuildLateSheet_(getSS_());
    rebuildRemindersSheet_(getSS_());
    const bs = getSS_().getSheetByName(CFG.SHEETS.BORROWERS), last = bs.getLastRow();
    if (last < 2) return;
    const noticeDays = Number(getSetting_('Días de aviso antes del vencimiento')) || 3;
    const today = new Date(); today.setHours(0, 0, 0, 0);
    for (let row = 2; row <= last; row++) {
      const loan = readLoan_(bs, row); if (!loan.loanId || !loan.email) continue;
      const status = bs.getRange(row, 15).getValue(), outstanding = Number(bs.getRange(row, 14).getValue()) || 0;
      if (status === ST.PAID || outstanding <= 0) continue;
      const due = loan.dueDate instanceof Date ? loan.dueDate : null; if (!due) continue;
      const dd = new Date(due); dd.setHours(0, 0, 0, 0);
      const days = Math.round((dd - today) / 86400000);
      const lastNotice = bs.getRange(row, 18).getValue();
      const lastStr = (lastNotice instanceof Date) ? fmtDate_(lastNotice) : '';
      let send = false, subject = '', intro = '';
      if (days >= 0 && days <= noticeDays && lastStr !== fmtDate_(today)) {
        send = true; subject = 'Recordatorio: su préstamo vence pronto';
        intro = 'Le recordamos que su préstamo ' + loan.loanId + ' vence el ' + fmtDate_(due) + '.';
      } else if (days < 0) {
        // en mora: avisar como máximo cada 7 días
        const daysSince = lastNotice instanceof Date ? Math.round((today - new Date(lastNotice.getFullYear(), lastNotice.getMonth(), lastNotice.getDate())) / 86400000) : 999;
        if (daysSince >= 7) { send = true; subject = 'Aviso de mora — préstamo ' + loan.loanId; intro = 'Su préstamo ' + loan.loanId + ' venció el ' + fmtDate_(due) + ' (' + daysLate_(due, today) + ' día(s) de atraso). Se aplica un recargo por mora del ' + lateFeePctText_() + ' por día (' + fmtMoney_(round2_(loan.totalDue * lateFeeRate_())) + ' por día) sobre el total a pagar.'; }
      }
      if (send) {
        try {
          sendBrandedEmail_(loan.email, subject, intro + ' Saldo actual: ' + fmtMoney_(outstanding) + '.',
            '<p>Estimado/a ' + esc_(loan.name) + ',</p><p>' + esc_(intro) + '</p><p><b>Saldo pendiente: ' + fmtMoney_(outstanding) + '</b></p>');
          bs.getRange(row, 18).setValue(today);
        } catch (e) { logError_('dailyTasks:email', e); }
      }
    }
  });
}

/* ===================== HELPERS ===================== */
function fillPaymentLookup_(sh, row) {
  const loanId = String(sh.getRange(row, 2).getValue()).trim();
  if (!loanId) { sh.getRange(row, 3, 1, 3).clearContent(); return; }
  const loan = findLoanById_(loanId);
  if (!loan) { sh.getRange(row, 3).setValue('?'); sh.getRange(row, 4, 1, 2).clearContent(); return; }
  sh.getRange(row, 3).setValue(firstName_(loan.name));
  sh.getRange(row, 4).setValue(lastName_(loan.name));
  sh.getRange(row, 5).setValue(loan.dni);
}
function recalcLoanPayments_(loanId) {
  const sh = getSS_().getSheetByName(CFG.SHEETS.PAYMENTS), last = sh.getLastRow();
  if (last < 2) return;
  const loan = findLoanById_(loanId); if (!loan) return;
  const rate = termRate_(loan.term);
  const data = sh.getRange(2, 2, last - 1, 6).getValues(), rows = [];
  data.forEach((r, i) => { if (String(r[0]).trim() === loanId && typeof r[5] === 'number' && r[5] !== 0) rows.push({ row: i + 2, date: r[4] instanceof Date ? r[4] : new Date(), amount: r[5] }); });
  const all = rows.map(x => ({ date: x.date, amount: x.amount }));
  rows.forEach(x => sh.getRange(x.row, 8).setValue(computeOutstanding_(loan.principal, rate, loan.loanDate, all, x.date)));
}
function nbAppend_(nb, values) {
  const last = Math.max(nb.getLastRow(), 1), dataRows = Math.max(last - 1, 1);
  const colB = nb.getRange(2, 2, dataRows, 1).getValues(); let row = last + 1;
  for (let i = 0; i < colB.length; i++) if (String(colB[i][0]).trim() === '') { row = i + 2; break; }
  nb.getRange(row, 1, 1, values.length).setValues([values]); return row;
}
function readLoan_(sh, row) {
  const v = sh.getRange(row, 1, 1, 15).getValues()[0];
  return {
    loanId: v[0], name: v[1], email: v[2], dni: v[3], phone: v[4], principal: Number(v[5]) || 0, term: Number(v[6]) || 0,
    loanDate: v[8] instanceof Date ? v[8] : new Date(v[8]),
    dueDate: v[9] instanceof Date ? v[9] : dueDateForTerm_(Number(v[6]) || 1, new Date(v[8])),
    interest: Number(v[10]) || 0, totalDue: Number(v[11]) || 0,
  };
}
function findLoanById_(loanId) {
  const sh = getSS_().getSheetByName(CFG.SHEETS.BORROWERS), last = sh.getLastRow();
  if (last < 2) return null;
  const data = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < data.length; i++) if (String(data[i][0]).trim() === loanId) return readLoan_(sh, i + 2);
  return null;
}
function loanPayments_(loanId) {
  const sh = getSS_().getSheetByName(CFG.SHEETS.PAYMENTS); if (sh.getLastRow() < 2) return [];
  const data = sh.getRange(2, 2, sh.getLastRow() - 1, 6).getValues(), out = [];
  data.forEach(r => { if (String(r[0]).trim() === loanId && typeof r[5] === 'number' && r[5] !== 0) out.push({ date: r[4] instanceof Date ? r[4] : new Date(), amount: r[5] }); });
  return out;
}
function nextLoanId_() {
  const bs = getSS_().getSheetByName(CFG.SHEETS.BORROWERS), last = bs.getLastRow(); let max = 0;
  if (last >= 2) bs.getRange(2, 1, last - 1, 1).getValues().forEach(r => { const m = /^L-(\d+)$/.exec(String(r[0]).trim()); if (m) max = Math.max(max, parseInt(m[1], 10)); });
  return 'L-' + String(max + 1).padStart(4, '0');
}
function firstEmptyBorrowerRow_(bs) {
  const vals = bs.getRange(2, 1, CFG.MAX_ROWS, 1).getValues();
  for (let i = 0; i < vals.length; i++) if (String(vals[i][0]).trim() === '') return i + 2;
  return bs.getLastRow() + 1;
}
function getSetting_(key) {
  const sh = getSS_().getSheetByName(CFG.SHEETS.SETTINGS);
  if (!sh || sh.getLastRow() < 2) return '';
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  for (const r of data) if (String(r[0]).trim() === key) return r[1];
  return '';
}
function spreadsheetParentFolder_() {
  const p = DriveApp.getFileById(getSS_().getId()).getParents();
  return p.hasNext() ? p.next() : DriveApp.getRootFolder();
}
function borrowerMainFolder_() {
  const parent = spreadsheetParentFolder_(), it = parent.getFoldersByName(CFG.MAIN_FOLDER);
  return it.hasNext() ? it.next() : parent.createFolder(CFG.MAIN_FOLDER);
}
function borrowerFolder_(folderName) {
  const main = borrowerMainFolder_(), it = main.getFoldersByName(folderName);
  return it.hasNext() ? it.next() : main.createFolder(folderName);
}
function borrowerFolderName_(fullName, dni) {
  return sanitizeName_([firstName_(fullName), lastName_(fullName), String(dni || '').trim()].filter(String).join(' ')) || sanitizeName_(fullName);
}
function savePhoto_(folder, blob, baseName) {
  try {
    if (!blob || !blob.getBytes || blob.getBytes().length === 0) return '';
    const ct = blob.getContentType() || '', ext = ct.indexOf('pdf') >= 0 ? 'pdf' : (ct.split('/')[1] || 'jpg');
    return folder.createFile(blob.setName(sanitizeName_(baseName) + '.' + ext)).getUrl();
  } catch (e) { logError_('savePhoto_', e); return ''; }
}
function extractDriveIds_(cell) {
  const s = String(cell == null ? '' : cell), ids = []; let m;
  const re = /(?:id=|\/d\/)([-\w]{20,})/g; while ((m = re.exec(s))) ids.push(m[1]);
  if (!ids.length) { const re2 = /[-\w]{25,}/g; while ((m = re2.exec(s))) ids.push(m[0]); }
  return ids;
}
function hasFile_(b) { return !!(b && b.getBytes && b.getBytes().length > 0); }
function getSS_() { return SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SS_ID')); }
function getOrCreate_(ss, name) { return ss.getSheetByName(name) || ss.insertSheet(name); }
function datePicker_() { return SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).setHelpText('Elija una fecha.').build(); }
function cc_(range, text, color) { return SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(text).setBackground(color).setRanges([range]).build(); }
function addMonths_(date, n) { const d = new Date(date.getFullYear(), date.getMonth(), date.getDate()), day = d.getDate(); d.setMonth(d.getMonth() + n); if (d.getDate() < day) d.setDate(0); return d; }
function round2_(n) { return Math.round((n + Number.EPSILON) * 100) / 100; }
/* ---- Plazos (términos del préstamo) ----
 * El "Plazo" se codifica como número entero: 15 = 15 días (25%), 1 = 1 mes (50%),
 * 2 = 2 meses (100%). El interés es fijo por el plazo. */
/** Interés fijo del plazo como fracción, o null si el plazo es inválido. */
function termRate_(term) { return term === 15 ? 0.25 : term === 1 ? 0.5 : term === 2 ? 1 : null; }
/** Etiqueta legible del plazo (p. ej. "15 días", "1 mes(es)"). */
function termLabel_(term) { return term === 15 ? '15 días' : term + ' mes(es)'; }
/** Porcentaje de interés del plazo como texto (p. ej. "25%"), o '' si es inválido. */
function termRatePct_(term) { const r = termRate_(term); return r === null ? '' : round2_(r * 100) + '%'; }
/** Fecha de vencimiento del plazo a partir de la fecha del préstamo. */
function dueDateForTerm_(term, loanDate) { return term === 15 ? addDays_(loanDate, 15) : addMonths_(loanDate, term); }
function sanitizeName_(s) { return String(s == null ? '' : s).replace(/[\/\\:*?"<>|]/g, '-').replace(/\s+/g, ' ').trim(); }
function firstName_(f) { return String(f == null ? '' : f).trim().split(/\s+/)[0] || ''; }
function lastName_(f) { const p = String(f == null ? '' : f).trim().split(/\s+/); return p.length > 1 ? p.slice(1).join(' ') : ''; }
function fmtMoney_(n) { try { return Number(n).toLocaleString(CFG.LOCALE, { style: 'currency', currency: CFG.CURRENCY_CODE }); } catch (e) { return '$ ' + round2_(Number(n)).toFixed(2); } }
const MESES_ES = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];
// Fechas para contrato y comunicaciones al cliente: año, mes (en palabra) y día. Ej.: "2026 Agosto 11".
function fmtDate_(d) {
  if (!(d instanceof Date)) d = new Date(d);
  const tz = Session.getScriptTimeZone();
  const y = Utilities.formatDate(d, tz, 'yyyy');
  const m = Number(Utilities.formatDate(d, tz, 'MM'));
  const day = Utilities.formatDate(d, tz, 'd');
  return y + ' ' + MESES_ES[m - 1] + ' ' + day;
}
function fmtDateTime_(d) {
  if (!(d instanceof Date)) d = new Date(d);
  return fmtDate_(d) + ' · ' + Utilities.formatDate(d, Session.getScriptTimeZone(), 'HH:mm');
}
function esc_(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }

/* ==================================================================
 *  DATOS DE PRUEBA  —  crear / eliminar (create & tear down)
 * ------------------------------------------------------------------
 *  Agrega 40 prestatarios de PRUEBA en distintos estados y casos
 *  límite (ACTIVO, VENCIDO, PAGADO). Los préstamos de prueba usan IDs
 *  con prefijo "T-XXXX" (los reales usan "L-XXXX"), de modo que la
 *  eliminación NUNCA toca datos reales. No se envían correos.
 * ================================================================== */
const TEST_PREFIX = 'T-';               // IDs de préstamos de prueba
const TEST_EMAIL_DOMAIN = 'prueba.test'; // dominio ficticio (no envía correos)

/** ¿Es un ID de préstamo de prueba? (T-0001, etc.) */
function isTestId_(v) { return /^T-\d+/i.test(String(v == null ? '' : v).trim()); }

function countTestBorrowers_(bs) {
  if (!bs || bs.getLastRow() < 2) return 0;
  return bs.getRange(2, 1, bs.getLastRow() - 1, 1).getValues().filter(r => isTestId_(r[0])).length;
}

function startOfDay_(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
function addDays_(d, n) { const x = startOfDay_(d); x.setDate(x.getDate() + n); return x; }

/* ---------- CREAR ---------- */
function seedTestBorrowersConfirm() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.alert('Datos de prueba',
    'Se agregarán 40 prestatarios de PRUEBA (IDs "T-XXXX") en distintos estados: '
    + 'ACTIVO, VENCIDO y PAGADO, incluyendo casos límite (vence hoy, vencido ayer, '
    + 'pago parcial, sobrepago, capital mínimo/máximo, nombres con acentos).\n\n'
    + 'No se envían correos. ¿Continuar?', ui.ButtonSet.YES_NO);
  if (r === ui.Button.YES) seedTestBorrowers();
}

function seedTestBorrowers() {
  return guard_('seedTestBorrowers', function () {
    const ss = getSS_();
    const bs = ss.getSheetByName(CFG.SHEETS.BORROWERS);
    const ps = ss.getSheetByName(CFG.SHEETS.PAYMENTS);
    if (!bs || !ps) throw new Error('Faltan las hojas base. Ejecute ① Configurar primero.');

    const existing = countTestBorrowers_(bs);
    if (existing > 0) throw new Error('Ya existen ' + existing + ' prestatarios de prueba. '
      + 'Use "Eliminar prestatarios de prueba" antes de volver a crearlos.');

    let specs = buildTestSpecs_();
    const startRow = firstEmptyBorrowerRow_(bs);
    const room = (CFG.MAX_ROWS + 1) - startRow + 1;
    if (room <= 0) throw new Error('No hay filas libres en "Prestatarios".');
    if (specs.length > room) specs = specs.slice(0, room);

    const rowsAG = [];   // columnas de entrada A..G
    const rowsI = [];    // columna I (Fecha del Préstamo)
    const payRows = [];  // filas para la hoja de Pagos
    const today = startOfDay_(new Date());

    specs.forEach(function (s, i) {
      const loanId = TEST_PREFIX + String(i + 1).padStart(4, '0');
      const rate = termRate_(s.term);
      const totalDue = round2_(s.capital * (1 + rate));
      // A..G  (G como texto "1"/"2" para respetar la validación de la columna)
      rowsAG.push([loanId, s.name, s.email, s.dni, s.phone, s.capital, String(s.term)]);
      rowsI.push([s.loanDate]);
      if (s.payFrac > 0) {
        const amount = round2_(totalDue * s.payFrac);
        let payDate = addDays_(s.loanDate, 5);
        if (payDate.getTime() > today.getTime()) payDate = today; // nunca pagar en el futuro
        payRows.push(['TP-' + String(i + 1).padStart(4, '0'), loanId,
          firstName_(s.name), lastName_(s.name), s.dni, payDate, amount,
          Math.max(0, round2_(totalDue - amount)), '']);
      }
    });

    bs.getRange(startRow, 1, rowsAG.length, 7).setValues(rowsAG);
    bs.getRange(startRow, 9, rowsI.length, 1).setValues(rowsI);
    bs.getRange(startRow, 6, rowsAG.length, 1).setNumberFormat(CFG.CURRENCY_FMT);
    bs.getRange(startRow, 9, rowsI.length, 1).setNumberFormat('yyyy-mm-dd');

    if (payRows.length) {
      const pstart = ps.getLastRow() + 1;
      ps.getRange(pstart, 1, payRows.length, 9).setValues(payRows);
      ps.getRange(pstart, 6, payRows.length, 1).setNumberFormat('yyyy-mm-dd');
      ps.getRange(pstart, 7, payRows.length, 2).setNumberFormat(CFG.CURRENCY_FMT);
    }

    SpreadsheetApp.flush();
    try { refreshAll(true); } catch (e) { logError_('seedTestBorrowers/refreshAll', e); }

    // Resumen por estado para la confirmación.
    const counts = summarizeTestStates_(bs);
    SpreadsheetApp.getUi().alert('Datos de prueba creados',
      'Se agregaron ' + rowsAG.length + ' prestatarios de prueba y ' + payRows.length + ' pagos.\n\n'
      + 'ACTIVO: ' + counts.ACTIVO + '\nVENCIDO: ' + counts.VENCIDO + '\nPAGADO: ' + counts.PAGADO,
      SpreadsheetApp.getUi().ButtonSet.OK);
  });
}

/** Cuenta los estados actuales de las filas de prueba (columna O = Estado). */
function summarizeTestStates_(bs) {
  const out = { ACTIVO: 0, VENCIDO: 0, PAGADO: 0, otros: 0 };
  if (!bs || bs.getLastRow() < 2) return out;
  const last = bs.getLastRow();
  const rows = bs.getRange(2, 1, last - 1, 15).getValues();
  rows.forEach(function (r) {
    if (!isTestId_(r[0])) return;
    const st = String(r[14]).trim();
    if (out[st] === undefined) out.otros++; else out[st]++;
  });
  return out;
}

/* ---------- ELIMINAR (tear down) ---------- */
function removeTestBorrowersConfirm() {
  const ui = SpreadsheetApp.getUi();
  const bs = getSS_().getSheetByName(CFG.SHEETS.BORROWERS);
  const n = countTestBorrowers_(bs);
  if (n === 0) { ui.alert('Datos de prueba', 'No hay prestatarios de prueba para eliminar.', ui.ButtonSet.OK); return; }
  const r = ui.alert('Datos de prueba',
    'Se eliminarán ' + n + ' prestatarios de prueba (IDs "T-XXXX") y todos sus pagos. '
    + 'Los datos reales ("L-XXXX") no se tocan.\n\n¿Continuar?', ui.ButtonSet.YES_NO);
  if (r === ui.Button.YES) removeTestBorrowers();
}

function removeTestBorrowers() {
  return guard_('removeTestBorrowers', function () {
    const ss = getSS_();
    const bs = ss.getSheetByName(CFG.SHEETS.BORROWERS);
    const ps = ss.getSheetByName(CFG.SHEETS.PAYMENTS);
    let removed = 0, payRemoved = 0;

    // Prestatarios: limpiar solo las celdas de entrada (conserva la plantilla de fórmulas).
    if (bs && bs.getLastRow() >= 2) {
      const last = bs.getLastRow();
      const ids = bs.getRange(2, 1, last - 1, 1).getValues();
      for (let i = 0; i < ids.length; i++) {
        if (!isTestId_(ids[i][0])) continue;
        const row = i + 2;
        bs.getRange(row, 1, 1, 7).clearContent();  // A..G (entradas)
        bs.getRange(row, 9).clearContent();         // I (fecha del préstamo)
        bs.getRange(row, 16, 1, 3).clearContent();  // P..R (contrato/avisos, por si acaso)
        removed++;
      }
    }

    // Pagos: eliminar filas de prueba (de abajo hacia arriba para no desplazar índices).
    if (ps && ps.getLastRow() >= 2) {
      const pids = ps.getRange(2, 2, ps.getLastRow() - 1, 1).getValues();
      for (let i = pids.length - 1; i >= 0; i--) {
        if (isTestId_(pids[i][0])) { ps.deleteRow(i + 2); payRemoved++; }
      }
    }

    SpreadsheetApp.flush();
    try { refreshAll(true); } catch (e) { logError_('removeTestBorrowers/refreshAll', e); }
    SpreadsheetApp.getUi().alert('Datos de prueba eliminados',
      'Se eliminaron ' + removed + ' prestatarios de prueba y ' + payRemoved + ' pagos.',
      SpreadsheetApp.getUi().ButtonSet.OK);
  });
}

/* ---------- CATÁLOGO DE 40 CASOS ---------- */
/**
 * Devuelve 40 especificaciones de prueba con estados y casos límite variados.
 * Cada spec: { name, email, dni, phone, capital, term(1|2), loanDate, payFrac }.
 *  - payFrac 0 = sin pagos; <1 = pago parcial; 1 = saldado; >1 = sobrepago.
 *  - Estado resultante (col O): PAGADO si saldo≤0; VENCIDO si venció; si no ACTIVO.
 *  - Los PAGADO se mantienen NO vencidos para que el recargo por mora no los reactive.
 */
function buildTestSpecs_() {
  const today = startOfDay_(new Date());
  const dueToday = t => addMonths_(today, -t);            // vence exactamente hoy
  const dueYest = t => addDays_(addMonths_(today, -t), -1); // venció ayer
  const ago = d => addDays_(today, -d);                    // fecha del préstamo = hoy - d

  // [nombre, capital, plazo, fechaPréstamo, payFrac, etiqueta]
  const raw = [
    /* ---- ACTIVO (vigente): vence en el futuro o justo hoy ---- */
    ['José Martínez',                 15000,  1, ago(0),        0,    'nuevo hoy'],
    ['Lucía Fernández',               30000,  2, ago(0),        0,    'nuevo hoy (2m)'],
    ['Carlos Gómez',                  12000,  1, ago(5),        0,    'vigente'],
    ['Ana María Suárez',              20000,  1, ago(10),       0.30, 'vigente c/pago parcial'],
    ['Diego Romero',                  18000,  1, ago(20),       0,    'vigente, vence pronto'],
    ['Sofía Ramírez',                 45000,  2, ago(15),       0.50, 'vigente 2m parcial'],
    ['Martín Acosta',                 60000,  2, ago(40),       0,    'vigente 2m'],
    ['Valentina Torres',              25000,  1, dueToday(1),   0,    'límite: vence HOY'],
    ['Nicolás Herrera',               50000,  2, dueToday(2),   0,    'límite: vence HOY (2m)'],
    ['Micaela Rossi',                     1,  1, ago(0),        0,    'capital mínimo ($1)'],
    ['Fernando Del Río',            5000000,  2, ago(3),        0,    'capital máximo'],
    ['María José Rodríguez de la Fuente', 22000, 1, ago(7),     0.10, 'nombre largo c/acentos'],

    /* ---- VENCIDO (en mora): la fecha de vencimiento ya pasó ---- */
    ['Roberto Silva',                 15000,  1, dueYest(1),    0,    'límite: venció AYER'],
    ['Paula Cabrera',                 28000,  2, dueYest(2),    0,    'límite: venció AYER (2m)'],
    ['Gonzalo Medina',                17000,  1, ago(45),       0,    'vencido ~15 días'],
    ['Camila Ortiz',                  19000,  1, ago(60),       0,    'vencido ~30 días'],
    ['Andrés Molina',                 21000,  1, ago(90),       0,    'vencido ~60 días'],
    ['Florencia Vega',                40000,  1, ago(120),      0.50, 'vencido c/pago parcial'],
    ['Sebastián Ríos',                55000,  2, ago(100),      0,    'vencido 2m'],
    ['Julieta Sosa',                  70000,  2, ago(150),      0.25, 'vencido 2m parcial'],
    ['Emiliano Castro',               16000,  1, ago(200),      0,    'vencido, muy atrasado'],
    ['Agustina Peña',                 80000,  2, ago(365),      0,    'vencido ~1 año'],
    ['Tomás Aguirre',                   500,  1, ago(50),       0,    'capital chico vencido'],
    ['Brenda Ledesma',              2000000,  2, ago(80),       0,    'capital grande vencido'],
    ['Ramón Ñáñez',                   14000,  1, ago(70),       0,    'vencido c/acentos'],
    ['Lucas Benítez',                 26000,  1, ago(55),       0.90, 'vencido, casi saldado'],

    /* ---- PAGADO (saldado): no vencido + pago total o sobrepago ---- */
    ['Patricia Núñez',                15000,  1, ago(10),       1.00, 'saldado'],
    ['Hernán Ibáñez',                 30000,  2, ago(5),        1.00, 'saldado 2m'],
    ['Rocío Franco',                  12000,  1, ago(2),        1.00, 'saldado reciente'],
    ['Federico Luna',                 18000,  1, ago(0),        1.00, 'saldado el mismo día'],
    ['Daniela Vera',                  20000,  1, ago(15),       1.05, 'saldado c/sobrepago'],
    ['Ignacio Paz',                   24000,  1, ago(25),       1.00, 'saldado, vencía pronto'],
    ['Antonella Bruno',               35000,  2, ago(20),       1.00, 'saldado 2m'],
    ['Mateo Correa',                      1,  1, ago(1),        1.00, 'saldado capital $1'],
    ['Guadalupe Ferreyra',          3000000,  2, ago(4),        1.00, 'saldado capital grande'],
    ['Ángela Muñoz',                  21000,  1, ago(8),        1.00, 'saldado c/acentos'],
    ['Bautista Godoy',                33000,  2, ago(12),       1.00, 'saldado 2m'],
    ['Renata Ponce',                  16000,  1, ago(6),        1.00, 'saldado'],
    ['Thiago Cardozo',                27000,  2, ago(3),        1.02, 'saldado c/sobrepago'],
    ['Delfina Ojeda',                 19000,  1, ago(9),        1.00, 'saldado'],
  ];

  return raw.map(function (r, i) {
    const name = r[0];
    return {
      name: name,
      email: slugEmail_(name, i),
      dni: 90000001 + i,
      phone: '+54 9 11 ' + String(40000000 + i),
      capital: r[1],
      term: r[2],
      loanDate: r[3],
      payFrac: r[4],
    };
  });
}

/** Correo ficticio único y determinístico a partir del nombre (sin acentos). */
function slugEmail_(name, i) {
  let base = String(name).toLowerCase();
  try { base = base.normalize('NFD').replace(/[\u0300-\u036f]/g, ''); } catch (e) { /* normalize no disponible */ }
  const slug = base.replace(/[^a-z0-9]+/g, '.').replace(/^\.+|\.+$/g, '');
  return slug + '.' + (i + 1) + '@' + TEST_EMAIL_DOMAIN;
}

/* ==================================================================
 *  AVISO DE MORA (correo desde "Pagos Atrasados")
 * ================================================================== */
/** Envía un aviso de mora en español para un préstamo y marca la fecha de envío. */
function sendOverdueNotice_(loanId) {
  return guard_('sendOverdueNotice_', function () {
    loanId = String(loanId || '').trim();
    const loan = findLoanById_(loanId);
    if (!loan) throw new Error('No se encontró el préstamo ' + loanId + '.');
    if (!loan.email) throw new Error('El préstamo ' + loanId + ' no tiene correo cargado.');
    const today = new Date();
    const rate = termRate_(loan.term);
    const pays = loanPayments_(loanId);
    const outstanding = computeOutstanding_(loan.principal, rate, loan.loanDate, pays, today, new Date(9999, 0, 1));
    const days = daysLate_(loan.dueDate, today);
    const feeDay = round2_(loan.totalDue * lateFeeRate_());
    const subject = 'Aviso de mora — préstamo ' + loan.loanId;
    const intro = 'Su préstamo ' + loan.loanId + ' venció el ' + fmtDate_(loan.dueDate) + ' (' + days +
      ' día(s) de atraso). Se aplica un recargo por mora del ' + lateFeePctText_() + ' por día (' +
      fmtMoney_(feeDay) + ' por día) sobre el total a pagar.';
    sendBrandedEmail_(loan.email, subject, intro + ' Saldo actual: ' + fmtMoney_(outstanding) + '.',
      '<p>Estimado/a ' + esc_(loan.name) + ',</p>' +
      '<p>' + esc_(intro) + '</p>' +
      '<p><b>Saldo pendiente: ' + fmtMoney_(outstanding) + '</b></p>' +
      '<p>Le solicitamos regularizar su situación a la brevedad para evitar mayores recargos. ' +
      'Si ya realizó el pago, por favor ignore este mensaje.</p>');
    // Registrar la fecha del último aviso en "Prestatarios" (col 18) para que persista.
    const bs = getSS_().getSheetByName(CFG.SHEETS.BORROWERS), last = bs.getLastRow();
    if (last >= 2) {
      const ids = bs.getRange(2, 1, last - 1, 1).getValues();
      for (let i = 0; i < ids.length; i++) {
        if (String(ids[i][0]).trim() === loanId) { bs.getRange(i + 2, 18).setValue(today).setNumberFormat('yyyy-mm-dd'); break; }
      }
    }
    return { date: today, message: 'Aviso de mora enviado a ' + loan.email + ' (' + days + ' día(s) de atraso).' };
  });
}

/* ==================================================================
 *  HOJA "SALDADOS" + MOVER PRÉSTAMOS PAGADOS
 * ================================================================== */
const CLEARED_HEADERS = ['ID Préstamo', 'Prestatario', 'DNI', 'Correo', 'Teléfono', 'Capital',
  'Plazo (meses)', 'Interés', 'Total a Pagar', 'Total Pagado', 'Fecha del Préstamo',
  'Fecha de Vencimiento', 'Fecha de Saldado'];

/** Crea (una vez) la hoja "Saldados". No borra el archivo si ya existe. */
function setupCleared_(ss) {
  const sh = getOrCreate_(ss, CFG.SHEETS.CLEARED);
  // Si ya está inicializada con encabezados, conservar el archivo existente.
  if (sh.getLastRow() >= 1 && String(sh.getRange(1, 1).getValue()).trim() === CLEARED_HEADERS[0]) return sh;
  sh.clear();
  sh.getRange(1, 1, 1, CLEARED_HEADERS.length).setValues([CLEARED_HEADERS])
    .setFontWeight('bold').setBackground('#38761d').setFontColor('#fff').setWrap(true);
  sh.setFrozenRows(1);
  [110, 170, 120, 200, 120, 110, 90, 110, 120, 120, 120, 130, 120].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.getRange('A1').setNote('Préstamos saldados (pagados en su totalidad), archivados desde "Prestatarios". ' +
    'Se completa con el botón "Mover saldados" o el menú. El historial en "Pagos" se conserva.');
  return sh;
}

/** Devuelve los números de fila de "Prestatarios" cuyos préstamos están PAGADOS. */
function clearedLoanRows_(bs) {
  const out = [];
  if (!bs || bs.getLastRow() < 2) return out;
  const data = bs.getRange(2, 1, bs.getLastRow() - 1, 15).getValues();
  for (let i = 0; i < data.length; i++) {
    if (!String(data[i][0]).trim()) continue;         // sin ID
    if (String(data[i][14]).trim() === ST.PAID) out.push(i + 2); // col O = Estado
  }
  return out;
}

/** Mueve todos los préstamos PAGADOS a "Saldados" y libera sus filas. Devuelve la cantidad movida. */
function doMoveCleared_() {
  return guard_('doMoveCleared_', function () {
    const ss = getSS_();
    const bs = ss.getSheetByName(CFG.SHEETS.BORROWERS);
    const cs = ss.getSheetByName(CFG.SHEETS.CLEARED) || setupCleared_(ss);
    if (!bs) throw new Error('Falta la hoja "Prestatarios". Ejecute ① Configurar.');
    const rowsToMove = clearedLoanRows_(bs);
    if (!rowsToMove.length) return 0;
    const now = new Date();
    const archive = [];
    rowsToMove.forEach(function (row) {
      // v = A..O (0..14): ID,Nombre,Correo,DNI,Teléfono,Capital,Plazo,Tasa,FechaPréstamo,FechaVenc,Interés,Total,Pagado,Saldo,Estado
      const v = bs.getRange(row, 1, 1, 15).getValues()[0];
      archive.push([v[0], v[1], v[3], v[2], v[4], v[5], v[6], v[10], v[11], v[12], v[8], v[9], now]);
      // Limpiar solo las celdas de entrada (conserva la plantilla de fórmulas).
      bs.getRange(row, 1, 1, 7).clearContent();  // A..G
      bs.getRange(row, 9).clearContent();         // I (fecha del préstamo)
      bs.getRange(row, 16, 1, 3).clearContent();  // P..R (contrato/avisos)
    });
    const start = cs.getLastRow() + 1;
    cs.getRange(start, 1, archive.length, CLEARED_HEADERS.length).setValues(archive);
    cs.getRange(start, 6, archive.length, 1).setNumberFormat(CFG.CURRENCY_FMT);   // Capital
    cs.getRange(start, 8, archive.length, 3).setNumberFormat(CFG.CURRENCY_FMT);   // Interés, Total, Pagado
    cs.getRange(start, 11, archive.length, 3).setNumberFormat('yyyy-mm-dd');      // 3 fechas
    SpreadsheetApp.flush();
    try { refreshAll(true); } catch (e) { logError_('doMoveCleared_/refreshAll', e); }
    return archive.length;
  });
}

/** Menú: confirma y mueve los préstamos saldados. */
function moveClearedBorrowersConfirm() {
  const ui = SpreadsheetApp.getUi();
  const bs = getSS_().getSheetByName(CFG.SHEETS.BORROWERS);
  const n = clearedLoanRows_(bs).length;
  if (n === 0) { ui.alert('Saldados', 'No hay préstamos saldados (pagados en su totalidad) para mover.', ui.ButtonSet.OK); return; }
  const r = ui.alert('Mover saldados',
    'Se moverán ' + n + ' préstamo(s) saldado(s) de "Prestatarios" a la hoja "Saldados". ' +
    'El historial de pagos se conserva. ¿Continuar?', ui.ButtonSet.YES_NO);
  if (r === ui.Button.YES) moveClearedBorrowers();
}

function moveClearedBorrowers() {
  const n = doMoveCleared_();
  SpreadsheetApp.getUi().alert('Saldados',
    n ? ('Se movieron ' + n + ' préstamo(s) saldado(s) a la hoja "Saldados".')
      : 'No hay préstamos saldados para mover.',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

/* ==================================================================
 *  VISIBILIDAD DE PESTAÑAS (operaciones diarias vs. secundarias)
 * ================================================================== */
function isMainSheet_(name) { return MAIN_SHEETS.indexOf(name) !== -1; }

/** Deja visibles las hojas de operaciones diarias y oculta las secundarias. Devuelve cuántas ocultó. */
function applyDailyVisibility_(ss) {
  ss = ss || getSS_();
  // Activar una hoja principal para no dejar activa una que se va a ocultar.
  const home = ss.getSheetByName(CFG.SHEETS.BORROWERS) || ss.getSheets()[0];
  if (home) { try { home.showSheet(); ss.setActiveSheet(home); } catch (e) {} }
  let hidden = 0;
  ss.getSheets().forEach(function (sh) {
    if (isMainSheet_(sh.getName())) { try { sh.showSheet(); } catch (e) {} }
    else { try { sh.hideSheet(); hidden++; } catch (e) { /* no se puede ocultar la última visible */ } }
  });
  return hidden;
}

/** Menú: ver solo las pestañas de operaciones diarias (oculta el resto). */
function showDailyOnly() {
  return guard_('showDailyOnly', function () {
    const ss = getSS_();
    const n = applyDailyVisibility_(ss);
    ss.toast('Vista de operaciones diarias: ' + n + ' hoja(s) ocultada(s). ' +
      'Menú ▸ 🗂 Pestañas ▸ "Mostrar todas las hojas" para verlas.', 'Pestañas', 6);
  });
}

/** Menú: mostrar todas las hojas (incluidas las secundarias). */
function showAllSheets() {
  return guard_('showAllSheets', function () {
    const ss = getSS_();
    let shown = 0;
    ss.getSheets().forEach(function (sh) {
      try { if (sh.isSheetHidden()) shown++; sh.showSheet(); } catch (e) {}
    });
    ss.toast(shown + ' hoja(s) mostrada(s). Todas las pestañas están visibles.', 'Pestañas', 6);
  });
}
