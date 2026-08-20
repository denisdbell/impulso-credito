/************************************************************
 * LOAN MANAGER  —  Google Sheets + Apps Script
 * ----------------------------------------------------------
 * Spreadsheet: "Loan Manager — Agreements, Payments & Balances"
 *
 * WHAT IT DOES
 *  - Loan register. Only 2 terms allowed: 1 month = 50%, 2 months = 100%.
 *  - Penalty after term: outstanding balance DOUBLES for every overdue
 *    month (compounding: +100% of the current outstanding per month).
 *  - Generate a Loan Agreement PDF from the row values and EMAIL it to
 *    the borrower. A link to each borrower's agreement is stored in the row.
 *  - Record payments: logs the date + amount, computes amount outstanding,
 *    and emails a receipt/statement to the borrower.
 *  - Per-borrower Summary: loans, principal, interest, total paid,
 *    total outstanding, status.
 *
 * SETUP
 *  1) Extensions > Apps Script. Delete any code, paste ALL of this. Save.
 *  2) Run  ->  onOpen  (grant permissions when asked: Sheets, Gmail, Drive).
 *  3) Reload the spreadsheet. Use the new "Loan Manager" menu > "Setup / Rebuild".
 *  4) Fill the Settings sheet (lender name, etc.), then add loans in Borrowers.
 *
 * Currency: ARS (Argentine Peso), es-AR formatting.
 ************************************************************/

const CFG = {
  CURRENCY_FMT: '"$"#,##0.00',          // ARS style
  LOCALE: 'es-AR',
  CURRENCY_CODE: 'ARS',
  PDF_FOLDER: 'Loan Agreements',
  MAX_ROWS: 300,                        // pre-filled formula rows in Borrowers
  SHEETS: { BORROWERS: 'Prestatarios', PAYMENTS: 'Pagos', SUMMARY: 'Resumen', SETTINGS: 'Configuración', AGREEMENT: 'Estudio de Contratos', STATEMENTS: 'Estudio de Estados', NEW_BORROWERS: 'Nuevos Prestatarios' },
  // Legacy (English) tab names → migrated in place to the Spanish names above.
  OLD_SHEETS: { 'Borrowers': 'Prestatarios', 'Payments': 'Pagos', 'Summary': 'Resumen', 'Settings': 'Configuración', 'Agreement Studio': 'Estudio de Contratos', 'Statements Studio': 'Estudio de Estados', 'New Borrowers': 'Nuevos Prestatarios' },
};

/* ============================ MENU ============================ */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Loan Manager')
    .addItem('① Setup / Rebuild sheets', 'setup')
    .addSeparator()
    .addItem('② Generate Agreement PDF + Email (selected loan row)', 'generateAgreementForActiveRow')
    .addItem('③ Record a payment (+ email receipt)', 'recordPaymentDialog')
    .addItem('④ Email statement to borrower (selected loan row)', 'emailStatementForActiveRow')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Agreement Studio')
      .addItem('Preview selected loan', 'previewAgreementStudio')
      .addItem('Send selected loan by email', 'sendAgreementStudio'))
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Statements Studio')
      .addItem('Email statement for selected loan', 'emailStatementStudio'))
    .addSeparator()
    .addItem('⑥ Borrower intake web form — set up / show link', 'createBorrowerForm')
    .addItem('⑤ Refresh formatting & summary', 'refreshAll')
    .addToUi();
}

/** Rename any legacy English sheet tabs to their Spanish names (in place — Google
 *  auto-updates formula references; data is preserved). */
function migrateSheetNames_(ss) {
  ss = ss || getSS_();
  Object.keys(CFG.OLD_SHEETS).forEach(oldName => {
    const target = CFG.OLD_SHEETS[oldName];
    const sh = ss.getSheetByName(oldName);
    if (sh && !ss.getSheetByName(target)) sh.setName(target);
  });
}

/* ============================ SETUP ============================ */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  PropertiesService.getScriptProperties().setProperty('SS_ID', ss.getId());
  migrateSheetNames_(ss);
  cleanupLegacySheets_(ss);
  setupSettings_(ss);
  setupBorrowers_(ss);
  setupPayments_(ss);
  setupSummary_(ss);
  setupAgreement_(ss);
  setupStatements_(ss);
  setupNewBorrowersPlain_(ss);
  ensureEditTrigger_();
  // Deploy the intake web form and get its link via menu ⑥.
  // remove default "Sheet1" if empty
  const s1 = ss.getSheetByName('Sheet1');
  if (s1 && ss.getSheets().length > 1) { try { ss.deleteSheet(s1); } catch (e) {} }
  ss.setActiveSheet(ss.getSheetByName(CFG.SHEETS.BORROWERS));
  SpreadsheetApp.getUi().alert('Setup complete. Fill the Settings sheet, then add loans in "Borrowers". ' +
    'Use the "Agreement Studio" sheet to preview & send agreements.');
}

/** Installable triggers: onEdit powers auto lookups/balances; a daily timer advances compounding. */
function ensureEditTrigger_() {
  const fns = ScriptApp.getProjectTriggers().map(t => t.getHandlerFunction());
  if (fns.indexOf('onEditInstallable') === -1) {
    ScriptApp.newTrigger('onEditInstallable')
      .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet()).onEdit().create();
  }
  if (fns.indexOf('dailyOutstandingRecalc') === -1) {
    ScriptApp.newTrigger('dailyOutstandingRecalc').timeBased().everyDays(1).atHour(1).create();
  }
}

/** Daily job: advance overdue compounding by recomputing all balances. */
function dailyOutstandingRecalc() { updateAllOutstanding_(); }

function setupSettings_(ss) {
  const sh = getOrCreate_(ss, CFG.SHEETS.SETTINGS);
  sh.clear();
  const rows = [
    ['AJUSTE', 'VALOR'],
    ['Lender Name', 'Your Name / Company'],
    ['Lender Email', Session.getActiveUser().getEmail() || ''],
    ['Lender Phone', ''],
    ['Lender Address', ''],
    ['Governing Jurisdiction', 'Buenos Aires, Argentina'],
    ['Late Penalty Clause', '100% of the outstanding balance is added for each month the loan remains unpaid after the due date (compounding monthly).'],
    ['Agreement Footer', 'This agreement is legally binding upon signature by both parties.'],
  ];
  sh.getRange(1, 1, rows.length, 2).setValues(rows);
  sh.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#1c4587').setFontColor('#ffffff');
  sh.getRange(2, 1, rows.length - 1, 1).setFontWeight('bold');
  sh.setColumnWidth(1, 200); sh.setColumnWidth(2, 480);
  sh.getRange(1, 1, rows.length, 2).setWrap(true);
  sh.setFrozenRows(1);
}

function setupBorrowers_(ss) {
  const sh = getOrCreate_(ss, CFG.SHEETS.BORROWERS);
  sh.clear();
  sh.getDataValidations && sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearDataValidations();
  const headers = ['ID Préstamo', 'Nombre del Prestatario', 'Correo', 'DNI', 'Teléfono',
    'Capital', 'Plazo (meses)', 'Tasa', 'Fecha del Préstamo', 'Fecha de Vencimiento',
    'Interés', 'Total a Pagar', 'Total Pagado', 'Saldo Pendiente (hoy)', 'Estado',
    'Contrato PDF', 'Contrato Enviado'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#1c4587').setFontColor('#ffffff').setWrap(true);
  sh.setFrozenRows(1);

  const N = CFG.MAX_ROWS;
  // Per-row formulas. NOTE: Outstanding (N) is NOT a formula — it is written by the
  // script (updateLoanOutstanding_) so it updates reliably the moment a payment changes.
  const fH = [], fJ = [], fKLM = [], fO = [];
  for (let r = 2; r <= N + 1; r++) {
    fH.push([`=IF($G${r}="","",IFS($G${r}=1,0.5,$G${r}=2,1,TRUE,"⚠ INVALID TERM"))`]);      // H Rate
    fJ.push([`=IF(OR($I${r}="",$G${r}=""),"",EDATE($I${r},$G${r}))`]);                        // J Due Date
    fKLM.push([
      `=IF(OR($F${r}="",NOT(ISNUMBER($H${r}))),"",$F${r}*$H${r})`,                            // K Interest
      `=IF($K${r}="","",$F${r}+$K${r})`,                                                       // L Total Due
      `=IF($A${r}="","",SUMIF('${CFG.SHEETS.PAYMENTS}'!$B:$B,$A${r},'${CFG.SHEETS.PAYMENTS}'!$G:$G))`, // M Total Paid
    ]);
    // O Status — blank N means "not computed yet" (shown as …)
    fO.push([`=IF($A${r}="","",IF($N${r}="","…",IF($N${r}<=0.009,"PAID",IF(AND($J${r}<>"",TODAY()>$J${r}),"OVERDUE","ACTIVE"))))`]);
  }
  sh.getRange(2, 8, N, 1).setFormulas(fH);    // H
  sh.getRange(2, 10, N, 1).setFormulas(fJ);   // J
  sh.getRange(2, 11, N, 3).setFormulas(fKLM); // K,L,M
  sh.getRange(2, 15, N, 1).setFormulas(fO);   // O (N=14 left for the script)

  // Formats
  sh.getRange(2, 6, N, 1).setNumberFormat(CFG.CURRENCY_FMT);   // F Principal
  sh.getRange(2, 8, N, 1).setNumberFormat('0%');               // H Rate
  sh.getRange(2, 9, N, 2).setNumberFormat('yyyy-mm-dd');       // I,J dates
  sh.getRange(2, 11, N, 4).setNumberFormat(CFG.CURRENCY_FMT);  // K,L,M,N money
  // Term validation: only 1 or 2
  const termRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['1', '2'], true).setAllowInvalid(false)
    .setHelpText('Term must be 1 month (50%) or 2 months (100%).').build();
  sh.getRange(2, 7, N, 1).setDataValidation(termRule);
  // Date picker on Loan Date (I)
  sh.getRange(2, 9, N, 1).setDataValidation(datePickerRule_());

  // Column widths
  const widths = [90, 160, 200, 100, 110, 110, 80, 60, 100, 100, 110, 130, 110, 130, 90, 220, 140];
  widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.setColumnWidth(16, 220);

  // Status conditional colors
  const range = sh.getRange(2, 15, N, 1);
  const rules = [
    condColor_(range, 'PAID', '#b6d7a8'),
    condColor_(range, 'OVERDUE', '#ea9999'),
    condColor_(range, 'ACTIVE', '#fff2cc'),
  ];
  sh.setConditionalFormatRules(rules);
  sh.getRange(2, 16, N, 1).setFontColor('#1155cc');

  // Example row (deleteable) to show the shape.
  // Cols A-G are entered by hand; H (Rate) is a formula so it is skipped;
  // I (Loan Date) is written separately.
  if (sh.getRange('A2').getValue() === '') {
    sh.getRange(2, 1, 1, 7).setValues([[
      'L-0001', 'Juan Pérez', 'juan@example.com', '20-12345678-9', '+54 11 5555-5555',
      100000, 1]]);
    sh.getRange(2, 9).setValue(new Date()); // I: Loan Date
  }
}

function setupPayments_(ss) {
  const sh = getOrCreate_(ss, CFG.SHEETS.PAYMENTS);
  sh.clear();
  // First Name, Last Name & DNI are pulled from Borrowers by Loan ID.
  const headers = ['ID Pago', 'ID Préstamo', 'Nombre', 'Apellido', 'DNI', 'Fecha de Pago',
    'Monto Pagado', 'Saldo Posterior', 'Recibo Enviado'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#38761d').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  sh.getRange(2, 6, CFG.MAX_ROWS, 1).setNumberFormat('yyyy-mm-dd');   // F Payment Date
  sh.getRange(2, 7, CFG.MAX_ROWS, 2).setNumberFormat(CFG.CURRENCY_FMT); // G,H money
  const w = [140, 90, 120, 120, 120, 110, 120, 140, 160];
  w.forEach((x, i) => sh.setColumnWidth(i + 1, x));
  // Loan ID dropdown so manual entries stay valid
  const b = ss.getSheetByName(CFG.SHEETS.BORROWERS);
  if (b) {
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(b.getRange('A2:A' + (CFG.MAX_ROWS + 1)), true)
      .setAllowInvalid(true).build();
    sh.getRange(2, 2, CFG.MAX_ROWS, 1).setDataValidation(rule);
  }
  // Date picker on Payment Date (F)
  sh.getRange(2, 6, CFG.MAX_ROWS, 1).setDataValidation(datePickerRule_());
  // Grey-out auto-filled columns to signal they are computed
  sh.getRange(2, 3, CFG.MAX_ROWS, 3).setBackground('#f3f3f3');  // First,Last,DNI
  sh.getRange(2, 8, CFG.MAX_ROWS, 1).setBackground('#f3f3f3');  // Outstanding After
  sh.getRange('C1').setNote('First Name, Last Name & DNI auto-fill from Borrowers, and Outstanding After ' +
    'is calculated automatically, whenever you set the Loan ID, Payment Date, or Amount.');
}

function setupSummary_(ss) {
  const sh = getOrCreate_(ss, CFG.SHEETS.SUMMARY);
  sh.clear();
  const B = CFG.SHEETS.BORROWERS;
  sh.getRange(1, 1, 1, 7).setValues([[
    'Prestatario', 'Préstamos', 'Capital Total', 'Interés Total', 'Total Pagado',
    'Saldo Total Pendiente', 'Estado']])
    .setFontWeight('bold').setBackground('#674ea7').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  sh.getRange('A2').setFormula(`=IFERROR(UNIQUE(FILTER(${B}!B2:B, ${B}!B2:B<>"")),"")`);
  const N = CFG.MAX_ROWS;
  const f = [];
  for (let r = 2; r <= N + 1; r++) {
    f.push([
      `=IF($A${r}="","",COUNTIF(${B}!$B:$B,$A${r}))`,
      `=IF($A${r}="","",SUMIF(${B}!$B:$B,$A${r},${B}!$F:$F))`,
      `=IF($A${r}="","",SUMIF(${B}!$B:$B,$A${r},${B}!$K:$K))`,
      `=IF($A${r}="","",SUMIF(${B}!$B:$B,$A${r},${B}!$M:$M))`,
      `=IF($A${r}="","",SUMIF(${B}!$B:$B,$A${r},${B}!$N:$N))`,
      `=IF($A${r}="","",IF($F${r}<=0.009,"CLEARED",IF(COUNTIFS(${B}!$B:$B,$A${r},${B}!$O:$O,"OVERDUE")>0,"OVERDUE","ACTIVE")))`,
    ]);
  }
  sh.getRange(2, 2, N, 6).setFormulas(f);
  sh.getRange(2, 3, N, 4).setNumberFormat(CFG.CURRENCY_FMT);
  const w = [180, 70, 130, 130, 130, 140, 100];
  w.forEach((x, i) => sh.setColumnWidth(i + 1, x));
  const range = sh.getRange(2, 7, N, 1);
  sh.setConditionalFormatRules([
    condColor_(range, 'CLEARED', '#b6d7a8'),
    condColor_(range, 'OVERDUE', '#ea9999'),
    condColor_(range, 'ACTIVE', '#fff2cc'),
  ]);
}

/* ==================== AGREEMENT STUDIO SHEET ==================== */
function setupAgreement_(ss) {
  const sh = getOrCreate_(ss, CFG.SHEETS.AGREEMENT);
  sh.clear();
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearDataValidations();
  const B = CFG.SHEETS.BORROWERS;

  sh.getRange('A1').setValue('ESTUDIO DE CONTRATOS')
    .setFontSize(16).setFontWeight('bold').setFontColor('#1c4587');
  sh.getRange('A2').setValue('Elija un ID de Préstamo, ① tilde Vista previa para revisar, luego ② tilde Enviar correo.').setFontColor('#666');

  sh.getRange('A4').setValue('ID Préstamo:').setFontWeight('bold');
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(ss.getSheetByName(B).getRange('A2:A' + (CFG.MAX_ROWS + 1)), true)
    .setAllowInvalid(false).build();
  sh.getRange('B4').setDataValidation(rule).setBackground('#fff2cc').setFontWeight('bold');

  // Live field preview (reads the chosen loan from Borrowers)
  const rows = [
    ['Prestatario', `=IFERROR(VLOOKUP($B$4,${B}!$A:$B,2,FALSE),"")`],
    ['DNI',         `=IFERROR(VLOOKUP($B$4,${B}!$A:$D,4,FALSE),"")`],
    ['Correo',      `=IFERROR(VLOOKUP($B$4,${B}!$A:$C,3,FALSE),"")`],
    ['Capital',     `=IFERROR(VLOOKUP($B$4,${B}!$A:$F,6,FALSE),"")`],
    ['Plazo (meses)',`=IFERROR(VLOOKUP($B$4,${B}!$A:$G,7,FALSE),"")`],
    ['Interés',     `=IFERROR(VLOOKUP($B$4,${B}!$A:$K,11,FALSE),"")`],
    ['Total a Pagar',`=IFERROR(VLOOKUP($B$4,${B}!$A:$L,12,FALSE),"")`],
    ['Fecha de Vencimiento', `=IFERROR(VLOOKUP($B$4,${B}!$A:$J,10,FALSE),"")`],
  ];
  sh.getRange(6, 1, rows.length, 1).setValues(rows.map(r => [r[0]])).setFontWeight('bold');
  sh.getRange(6, 2, rows.length, 1).setFormulas(rows.map(r => [r[1]]));
  sh.getRange(9, 2).setNumberFormat(CFG.CURRENCY_FMT);   // Principal
  sh.getRange(11, 2).setNumberFormat(CFG.CURRENCY_FMT);  // Interest
  sh.getRange(12, 2).setNumberFormat(CFG.CURRENCY_FMT);  // Total Due
  sh.getRange(13, 2).setNumberFormat('yyyy-mm-dd');      // Due Date

  // Button placeholders (assign a drawing to these — see note in A21).
  buttonCell_(sh, 'A15:B15', '①  ▶  VISTA PREVIA DEL CONTRATO', '#1c4587');
  buttonCell_(sh, 'A16:B16', '②  ✉  ENVIAR CORREO', '#38761d');

  sh.getRange('A18').setValue('Última vista previa PDF:').setFontWeight('bold');
  // B18 will receive a "Open PDF preview" hyperlink after a preview
  sh.getRange('A19').setValue('Estado:').setFontWeight('bold');

  sh.getRange('A21').setValue(
    'PARA ACTIVAR LOS BOTONES (una vez): Insertar ▸ Dibujo, dibuje un recuadro, Guardar y cerrar, ' +
    'colóquelo sobre la celda de arriba, haga clic ▸ ⋮ ▸ Asignar secuencia de comandos, y escriba:  ' +
    'previewAgreementStudio  (Vista previa)  /  sendAgreementStudio  (Enviar). ' +
    'O use el menú: Loan Manager ▸ Agreement Studio.')
    .setFontColor('#7f6000').setWrap(true);
  sh.getRange('A21:B24').merge();

  sh.setColumnWidth(1, 160); sh.setColumnWidth(2, 380);
  sh.getRange('A1:B2').setWrap(true);
}

/** Style a merged range to look like a button. */
function buttonCell_(sh, a1, label, color) {
  const r = sh.getRange(a1);
  r.merge().setValue(label).setBackground(color).setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(12).setHorizontalAlignment('center')
    .setVerticalAlignment('middle').setBorder(true, true, true, true, false, false);
  sh.setRowHeight(r.getRow(), 34);
}

/* ==================== STATEMENTS STUDIO SHEET ==================== */
function setupStatements_(ss) {
  const sh = getOrCreate_(ss, CFG.SHEETS.STATEMENTS);
  sh.clear();
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearDataValidations();
  const B = CFG.SHEETS.BORROWERS, P = CFG.SHEETS.PAYMENTS;

  sh.getRange('A1').setValue('ESTUDIO DE ESTADOS DE CUENTA')
    .setFontSize(16).setFontWeight('bold').setFontColor('#674ea7');
  sh.getRange('A2').setValue('Elija un ID de Préstamo para ver el prestatario, cada pago y el estado del préstamo.').setFontColor('#666');

  sh.getRange('A4').setValue('ID Préstamo:').setFontWeight('bold');
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(ss.getSheetByName(B).getRange('A2:A' + (CFG.MAX_ROWS + 1)), true)
    .setAllowInvalid(false).build();
  sh.getRange('B4').setDataValidation(rule).setBackground('#fff2cc').setFontWeight('bold');

  // Loan summary block (VLOOKUP into Borrowers)
  const rows = [
    ['Prestatario',  `=IFERROR(VLOOKUP($B$4,${B}!$A:$B,2,FALSE),"")`],
    ['DNI',          `=IFERROR(VLOOKUP($B$4,${B}!$A:$D,4,FALSE),"")`],
    ['Correo',       `=IFERROR(VLOOKUP($B$4,${B}!$A:$C,3,FALSE),"")`],
    ['Capital',      `=IFERROR(VLOOKUP($B$4,${B}!$A:$F,6,FALSE),"")`],
    ['Total a Pagar',`=IFERROR(VLOOKUP($B$4,${B}!$A:$L,12,FALSE),"")`],
    ['Total Pagado', `=IFERROR(VLOOKUP($B$4,${B}!$A:$M,13,FALSE),"")`],
    ['Saldo Pendiente',`=IFERROR(VLOOKUP($B$4,${B}!$A:$N,14,FALSE),"")`],
    ['Fecha de Vencimiento', `=IFERROR(VLOOKUP($B$4,${B}!$A:$J,10,FALSE),"")`],
    ['ESTADO',       `=IFERROR(VLOOKUP($B$4,${B}!$A:$O,15,FALSE),"")`],
  ];
  sh.getRange(6, 1, rows.length, 1).setValues(rows.map(r => [r[0]])).setFontWeight('bold');
  sh.getRange(6, 2, rows.length, 1).setFormulas(rows.map(r => [r[1]]));
  sh.getRange(9, 2).setNumberFormat(CFG.CURRENCY_FMT);    // Principal
  sh.getRange(10, 2).setNumberFormat(CFG.CURRENCY_FMT);   // Total Due
  sh.getRange(11, 2).setNumberFormat(CFG.CURRENCY_FMT);   // Total Paid
  sh.getRange(12, 2).setNumberFormat(CFG.CURRENCY_FMT);   // Outstanding
  sh.getRange(13, 2).setNumberFormat('yyyy-mm-dd');       // Due Date
  const statusCell = sh.getRange(14, 2);                  // STATUS
  statusCell.setFontWeight('bold').setFontSize(12).setHorizontalAlignment('center');
  sh.setConditionalFormatRules([
    condColor_(statusCell, 'PAID', '#b6d7a8'),
    condColor_(statusCell, 'OVERDUE', '#ea9999'),
    condColor_(statusCell, 'ACTIVE', '#fff2cc'),
  ]);

  // "Email statement" button placeholder
  buttonCell_(sh, 'A16:B16', '✉  ENVIAR ESTADO AL PRESTATARIO', '#674ea7');
  sh.getRange('A18').setValue(
    'Asigne un dibujo al botón de arriba → nombre de la función:  emailStatementStudio  ' +
    '(o use el menú: Loan Manager ▸ Statements Studio).')
    .setFontColor('#7f6000').setWrap(true);
  sh.getRange('A18:D19').merge();

  // Payments table — all payments for the selected loan (spills automatically)
  sh.getRange('A21:C21').setValues([['Fecha de Pago', 'Monto Pagado', 'Saldo Posterior']])
    .setFontWeight('bold').setBackground('#674ea7').setFontColor('#ffffff');
  sh.getRange('A22').setFormula(
    `=IFERROR(FILTER({${P}!$F$2:$F, ${P}!$G$2:$G, ${P}!$H$2:$H}, ${P}!$B$2:$B=$B$4), "— sin pagos registrados —")`);
  sh.getRange('A22:A' + (CFG.MAX_ROWS + 21)).setNumberFormat('yyyy-mm-dd');
  sh.getRange('B22:C' + (CFG.MAX_ROWS + 21)).setNumberFormat(CFG.CURRENCY_FMT);

  const w = [150, 150, 160, 150];
  w.forEach((x, i) => sh.setColumnWidth(i + 1, x));
  sh.getRange('A1:D2').setWrap(true);
  sh.setFrozenRows(21);
}

function statementStatus_(msg) {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEETS.STATEMENTS)
    .getRange('D16').setValue(msg + '  (' + fmtDate_(new Date()) + ')');
}
function emailStatementStudio() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEETS.STATEMENTS);
  const loanId = String(sh.getRange('B4').getValue()).trim();
  if (!loanId) { statementStatus_('Pick a Loan ID in B4 first.'); return; }
  const loan = findLoanById_(loanId);
  if (!loan) { statementStatus_(`Loan "${loanId}" not found.`); return; }
  emailStatementForLoan_(loan, statementStatus_);
}

/** Read the Loan ID currently selected in Agreement Studio. */
function studioLoan_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEETS.AGREEMENT);
  const loanId = String(sh.getRange('B4').getValue()).trim();
  if (!loanId) { studioStatus_('Pick a Loan ID in B4 first.'); return null; }
  const loan = findLoanById_(loanId);
  if (!loan) { studioStatus_(`Loan "${loanId}" not found.`); return null; }
  return loan;
}
function studioStatus_(msg) {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEETS.AGREEMENT)
    .getRange('B19').setValue(msg + '  (' + fmtDate_(new Date()) + ')');
}

/** ① Preview: build the PDF, drop a link to it, and try to show the rendered doc in a dialog. */
function previewAgreementStudio() {
  const loan = studioLoan_();
  if (!loan) return;
  const file = makeAgreementFile_(loan);
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEETS.AGREEMENT);
  sh.getRange('B18').setFormula(`=HYPERLINK("${file.getUrl()}","👁 Open PDF preview")`);
  writeAgreementLinkByLoanId_(loan.loanId, file);
  studioStatus_('Preview ready — open the PDF link above, then tick ② to send.');
  // Nice-to-have: render inline. Wrapped so a trigger context that blocks UI never errors.
  try {
    const html = HtmlService.createHtmlOutput(agreementDocHtml_(loan)).setWidth(820).setHeight(600);
    SpreadsheetApp.getUi().showModalDialog(html, `Agreement preview — ${loan.loanId}`);
  } catch (err) { /* dialog not available from this context; PDF link is the fallback */ }
}

/** ② Send: email the agreement to the borrower. */
function sendAgreementStudio() {
  const loan = studioLoan_();
  if (!loan) return;
  if (!loan.email) { studioStatus_('Borrower has no email on file.'); return; }
  const file = makeAgreementFile_(loan);
  writeAgreementLinkByLoanId_(loan.loanId, file);
  emailAgreement_(loan, file);
  studioStatus_(`Agreement emailed to ${loan.email}.`);
}

/* ==================== NEW BORROWERS (custom web form + photo upload) ==================== */
// Fixed column layout for the New Borrowers sheet (the web app appends rows here).
const NB_HEADERS = ['Fecha de Envío', 'Nombre completo', 'Correo', 'DNI', 'Teléfono', 'Monto Solicitado',
  'Plazo (meses)', 'Notas / Motivo', 'Foto DNI', 'Foto CUIL',
  'Verificado?', 'Resultado'];

/** Append a row to New Borrowers at the first row whose Nombre (col B) is empty.
 *  Avoids appendRow's reliance on getLastRow (which stray checkbox values can inflate). */
function nbAppend_(nb, values) {
  const last = Math.max(nb.getLastRow(), 1);
  const dataRows = Math.max(last - 1, 1);
  const colB = nb.getRange(2, 2, dataRows, 1).getValues();
  let row = last + 1;
  for (let i = 0; i < colB.length; i++) { if (String(colB[i][0]).trim() === '') { row = i + 2; break; } }
  nb.getRange(row, 1, 1, values.length).setValues([values]);
  return row;
}

/** Delete leftover intake sheets like "New Borrowers (old …)" and "New Borrowers (pre-form)". */
function cleanupLegacySheets_(ss) {
  ss = ss || getSS_();
  ss.getSheets().forEach(s => {
    const n = s.getName();
    if (/^New Borrowers \(old/i.test(n) || n === 'New Borrowers (pre-form)') {
      try { ss.deleteSheet(s); } catch (e) {}
    }
  });
}

/** Creates / normalizes the New Borrowers sheet used by the web form (non-destructive). */
function setupNewBorrowersPlain_(ss) {
  ss = ss || getSS_();
  // Reuse the Spanish sheet, or the legacy English one (renaming it) — never lose rows.
  let sh = ss.getSheetByName(CFG.SHEETS.NEW_BORROWERS) || ss.getSheetByName('New Borrowers');
  if (!sh) sh = ss.insertSheet(CFG.SHEETS.NEW_BORROWERS);
  else if (sh.getName() !== CFG.SHEETS.NEW_BORROWERS && !ss.getSheetByName(CFG.SHEETS.NEW_BORROWERS))
    sh.setName(CFG.SHEETS.NEW_BORROWERS);
  // Column order is identical across versions, so overwriting the header row is safe.
  sh.getRange(1, 1, 1, NB_HEADERS.length).setValues([NB_HEADERS])
    .setFontWeight('bold').setBackground('#b45f06').setFontColor('#ffffff').setWrap(true);
  sh.setFrozenRows(1);
  sh.getRange(2, 11, CFG.MAX_ROWS, 1)  // K Verified?
    .setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
  sh.getRange(2, 1, CFG.MAX_ROWS, 1).setNumberFormat('yyyy-mm-dd hh:mm');
  sh.getRange(2, 6, CFG.MAX_ROWS, 1).setNumberFormat(CFG.CURRENCY_FMT);
  const w = [150, 170, 210, 130, 120, 140, 120, 220, 210, 210, 90, 230];
  w.forEach((x, i) => sh.setColumnWidth(i + 1, x));
  sh.getRange('K1').setNote('Tilde "Verificado?" para aprobar un solicitante — se mueve automáticamente ' +
    'a la hoja Prestatarios y sus fotos quedan en la carpeta "Nombre Apellido DNI".');
  return sh;
}

/** Menu ⑥ — ensure the sheet, then show the web-form link (or deployment steps). */
function createBorrowerForm() {
  setupNewBorrowersPlain_(getSS_());
  let url = '';
  try { url = ScriptApp.getService().getUrl(); } catch (e) { url = ''; }
  if (url) {
    PropertiesService.getScriptProperties().setProperty('INTAKE_WEBAPP_URL', url);
    getSS_().getSheetByName(CFG.SHEETS.NEW_BORROWERS).getRange('A1')
      .setNote('BORROWER INTAKE FORM (share with applicants):\n' + url);
    alertMsg_('Borrower intake web form is live.\n\nShare this link with applicants:\n' + url +
      '\n\nIt collects their details AND photo uploads (DNI + CUIL), which appear on "New Borrowers". ' +
      'Tick "Verified?" to move an applicant to Borrowers and file their photos.');
  } else {
    alertMsg_('The "New Borrowers" sheet is ready. Now DEPLOY the web form (one-time):\n\n' +
      '1. Extensions ▸ Apps Script.\n' +
      '2. Top-right "Deploy" ▸ New deployment.\n' +
      '3. Click the gear ▸ select "Web app".\n' +
      '4. Execute as: Me.   Who has access: Anyone.\n' +
      '5. Deploy ▸ authorize ▸ copy the "Web app URL".\n\n' +
      'Then run ⑥ again — it will display the shareable link.');
  }
}

/** Web app entry point — serves the borrower intake page (with photo upload). */
function doGet() {
  return HtmlService.createHtmlOutput(intakePageHtml_())
    .setTitle('Borrower Intake Form')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function intakePageHtml_() {
  const lender = esc_(getSetting_('Lender Name') || 'Gestor de Préstamos');
  return `<!DOCTYPE html><html lang="es"><head><base target="_top"><style>
    body{font-family:Arial,Helvetica,sans-serif;background:#f4f6fa;margin:0;padding:24px;color:#222}
    .card{max-width:560px;margin:0 auto;background:#fff;border-radius:10px;padding:24px 28px;box-shadow:0 2px 12px rgba(0,0,0,.08)}
    h1{font-size:20px;color:#1c4587;margin:0 0 4px}
    p.sub{color:#666;margin:0 0 18px}
    label{display:block;font-weight:bold;margin:14px 0 4px;font-size:14px}
    input,textarea,select{width:100%;box-sizing:border-box;padding:9px 10px;border:1px solid #ccc;border-radius:6px;font-size:14px}
    .hint{font-weight:normal;color:#888;font-size:12px}
    .req{color:#c00}
    button{margin-top:22px;width:100%;background:#1c4587;color:#fff;border:0;border-radius:6px;padding:12px;font-size:15px;font-weight:bold;cursor:pointer}
    button:disabled{background:#9db4d6}
    #msg{margin-top:16px;padding:12px;border-radius:6px;display:none}
    .ok{background:#d9ead3;color:#274e13}.err{background:#f4cccc;color:#990000}
  </style></head><body>
  <div class="card">
    <h1>Formulario de Solicitud de Préstamo</h1>
    <p class="sub">${lender} — solicite un préstamo. Su solicitud será revisada y verificada. Todos los campos son obligatorios.</p>
    <form id="f">
      <label>Nombre completo <span class="req">*</span><input name="fullName" required></label>
      <label>Correo electrónico <span class="req">*</span><input name="email" type="email" required></label>
      <label>DNI <span class="req">*</span><input name="dni" required></label>
      <label>Teléfono <span class="req">*</span><input name="phone" required></label>
      <label>Monto del préstamo solicitado (ARS) <span class="req">*</span><input name="amount" type="number" min="1" step="any" required></label>
      <label>Plazo del préstamo <span class="req">*</span><select name="term" required>
        <option value="">— elegir —</option>
        <option value="1">1 mes (50% de interés)</option>
        <option value="2">2 meses (100% de interés)</option></select></label>
      <label>Notas / Motivo <span class="req">*</span><textarea name="notes" rows="2" required></textarea></label>
      <label>Foto del DNI <span class="req">*</span> <span class="hint">— imagen o PDF</span>
        <input name="dniPhoto" type="file" accept="image/*,.pdf" required></label>
      <label>Foto del CUIL <span class="req">*</span> <span class="hint">— imagen o PDF</span>
        <input name="cuilPhoto" type="file" accept="image/*,.pdf" required></label>
      <button type="submit" id="btn">Enviar solicitud</button>
    </form>
    <div id="msg"></div>
  </div>
  <script>
    var f=document.getElementById('f'), btn=document.getElementById('btn'), msg=document.getElementById('msg');
    f.addEventListener('submit', function(e){
      e.preventDefault();
      btn.disabled=true; btn.textContent='Subiendo…'; msg.style.display='none';
      google.script.run
        .withSuccessHandler(function(r){ msg.className='ok'; msg.style.display='block'; msg.textContent=r; f.reset(); btn.textContent='Enviado ✓'; })
        .withFailureHandler(function(err){ msg.className='err'; msg.style.display='block'; msg.textContent=(err.message||err); btn.disabled=false; btn.textContent='Enviar solicitud'; })
        .submitBorrowerIntake(f);
    });
  </script></body></html>`;
}

/** Save one uploaded blob to a folder, returning its Drive URL (or '' if empty). */
function savePhoto_(folder, blob, baseName) {
  try {
    if (!blob || !blob.getBytes || blob.getBytes().length === 0) return '';
    const ct = blob.getContentType() || '';
    const ext = ct.indexOf('pdf') >= 0 ? 'pdf' : (ct.split('/')[1] || 'jpg');
    return folder.createFile(blob.setName(sanitizeName_(baseName) + '.' + ext)).getUrl();
  } catch (e) { return ''; }
}

function hasFile_(blob) { return !!(blob && blob.getBytes && blob.getBytes().length > 0); }

/** Diagnostic (run from the editor): append a test row to prove the sheet path works. */
function testAddNewBorrower() {
  const ss = getSS_();
  cleanupLegacySheets_(ss);
  const nb = ss.getSheetByName(CFG.SHEETS.NEW_BORROWERS) || setupNewBorrowersPlain_(ss);
  const row = nbAppend_(nb, [new Date(), 'Prueba Apellido', 'prueba@example.com', '20-11111111-1',
    '+54 11 1111-1111', 50000, 1, 'Fila de prueba', '', '', false, 'Pendiente de verificación']);
  nb.getRange(row, 1).setNumberFormat('yyyy-mm-dd hh:mm');
  nb.getRange(row, 6).setNumberFormat(CFG.CURRENCY_FMT);
  try { ss.setActiveSheet(nb); nb.setActiveRange(nb.getRange(row, 1)); } catch (e) {}
  SpreadsheetApp.getUi().alert('Fila de prueba agregada en la hoja "' + nb.getName() + '", fila ' + row +
    '. Si la ve, la hoja funciona y el problema es la implementación (redeploy) del formulario web.');
}

/** Called by the web form. Saves photos + appends a Pending row to New Borrowers. All fields required. */
function submitBorrowerIntake(form) {
  const ss = getSS_();
  const nb = ss.getSheetByName(CFG.SHEETS.NEW_BORROWERS) || setupNewBorrowersPlain_(ss);
  const name = String(form.fullName || '').trim();
  const email = String(form.email || '').trim();
  const dni = String(form.dni || '').trim();
  const phone = String(form.phone || '').trim();
  const amount = Number(String(form.amount || '').replace(/[^0-9.\-]/g, '')) || '';
  const term = Number(form.term) || '';
  const notes = String(form.notes || '').trim();

  // Every field is mandatory.
  if (!name) throw new Error('El nombre completo es obligatorio.');
  if (!email) throw new Error('El correo electrónico es obligatorio.');
  if (!dni) throw new Error('El DNI es obligatorio.');
  if (!phone) throw new Error('El teléfono es obligatorio.');
  if (!amount) throw new Error('El monto del préstamo es obligatorio y debe ser válido.');
  if (term !== 1 && term !== 2) throw new Error('Debe elegir un plazo de 1 o 2 meses.');
  if (!notes) throw new Error('El campo Notas / Motivo es obligatorio.');
  if (!hasFile_(form.dniPhoto)) throw new Error('Debe adjuntar la foto del DNI.');
  if (!hasFile_(form.cuilPhoto)) throw new Error('Debe adjuntar la foto del CUIL.');

  // Each borrower gets their own subfolder (Nombre Apellido DNI) inside the main folder.
  const folder = borrowerFolder_(borrowerFolderName_(name, dni));
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');
  const dniUrl = savePhoto_(folder, form.dniPhoto, name + ' DNI ' + stamp);
  const cuilUrl = savePhoto_(folder, form.cuilPhoto, name + ' CUIL ' + stamp);

  const row = nbAppend_(nb, [new Date(), name, email, dni, phone, amount, term, notes, dniUrl, cuilUrl, false, 'Pendiente de verificación']);
  nb.getRange(row, 1).setNumberFormat('yyyy-mm-dd hh:mm');
  nb.getRange(row, 6).setNumberFormat(CFG.CURRENCY_FMT);
  return '¡Solicitud recibida! Gracias, ' + name + '. Revisaremos su solicitud y nos pondremos en contacto.';
}

/** Header title → 1-based column index for the New Borrowers sheet (robust to column shifts). */
function nbHeaderMap_(nb) {
  const hdr = nb.getRange(1, 1, 1, nb.getLastColumn()).getValues()[0];
  const m = {};
  hdr.forEach((h, i) => { const k = String(h).trim(); if (k) m[k] = i + 1; });
  return m;
}

/** Move a verified applicant row into the Borrowers sheet, then remove it from New Borrowers. */
function moveVerifiedBorrower_(nb, row) {
  const m = nbHeaderMap_(nb);
  // Tolerant of both Spanish and legacy English headers.
  const colOf = (...titles) => { for (const t of titles) if (m[t]) return m[t]; return 0; };
  const pick = (...titles) => { const c = colOf.apply(null, titles); return c ? nb.getRange(row, c).getValue() : ''; };
  const verCol = colOf('Verificado?', 'Verified?');
  const resCol = colOf('Resultado', 'Result') || (verCol ? verCol + 1 : nb.getLastColumn());
  const fail = msg => { nb.getRange(row, resCol).setValue(msg); if (verCol) nb.getRange(row, verCol).setValue(false); };

  const name = String(pick('Nombre completo', 'Full Name')).trim();
  const email = pick('Correo', 'Email'), dni = pick('DNI', 'ID / DNI'), phone = pick('Teléfono', 'Phone');
  const amount = Number(pick('Monto Solicitado', 'Loan Amount Requested')) || 0;
  const term = Number(pick('Plazo (meses)', 'Loan Term (months)')) || 0;
  const natId = pick('Foto DNI', 'National ID (DNI) — upload'), workId = pick('Foto CUIL', 'Work ID (CUIL) — upload'), cuil = '';
  if (!name) { fail('⚠ Falta el nombre'); return; }
  if (term !== 1 && term !== 2) { fail('⚠ El plazo debe ser 1 o 2'); return; }
  if (!amount) { fail('⚠ Monto faltante/ inválido'); return; }

  const bs = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEETS.BORROWERS);
  const loanId = nextLoanId_();
  const tRow = firstEmptyBorrowerRow_(bs);
  bs.getRange(tRow, 1, 1, 7).setValues([[loanId, name, email, dni, phone, amount, term]]);
  bs.getRange(tRow, 9).setValue(new Date()); // Loan Date = verification date
  bs.getRange(tRow, 9).setNumberFormat('yyyy-mm-dd');
  writeOutstandingForRow_(bs, tRow);

  // The borrower's subfolder (FirstName LastName DNI) was created at upload; reuse it.
  // Ensure any photos are filed there (no-op if already inside).
  const folder = borrowerFolder_(borrowerFolderName_(name, dni) || loanId);
  let moved = 0;
  extractDriveIds_(natId).concat(extractDriveIds_(workId)).forEach(id => {
    try { DriveApp.getFileById(id).moveTo(folder); moved++; } catch (e) {}
  });

  // Preserve folder link + uploaded/entered IDs (the row is about to be deleted) as a note on the DNI cell.
  const idParts = ['📁 Folder: ' + folder.getUrl()];
  if (natId) idParts.push('National ID (DNI): ' + natId);
  if (workId) idParts.push('Work ID (CUIL): ' + workId);
  if (cuil) idParts.push('CUIL number: ' + cuil);
  bs.getRange(tRow, 4).setNote(idParts.join('\n'));

  nb.deleteRow(row); // "move" = remove from New Borrowers
  try { SpreadsheetApp.getActiveSpreadsheet().toast(`Verified ${name} → ${loanId} · ${moved} photo(s) filed`, 'Borrower added', 5); } catch (e) {}
}

/** Next sequential Loan ID like L-0007, based on existing Borrowers. */
function nextLoanId_() {
  const bs = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEETS.BORROWERS);
  const last = bs.getLastRow();
  let max = 0;
  if (last >= 2) {
    bs.getRange(2, 1, last - 1, 1).getValues().forEach(r => {
      const m = /^L-(\d+)$/.exec(String(r[0]).trim());
      if (m) max = Math.max(max, parseInt(m[1], 10));
    });
  }
  return 'L-' + String(max + 1).padStart(4, '0');
}

/** First blank Loan-ID row in Borrowers (which already has its formulas pre-filled). */
function firstEmptyBorrowerRow_(bs) {
  const vals = bs.getRange(2, 1, CFG.MAX_ROWS, 1).getValues();
  for (let i = 0; i < vals.length; i++) if (String(vals[i][0]).trim() === '') return i + 2;
  return bs.getLastRow() + 1;
}

/* ==================== INSTALLABLE onEdit (checkbox buttons + payment lookup) ==================== */
function onEditInstallable(e) {
  try {
    const sh = e.range.getSheet();
    const name = sh.getName();
    const c0 = e.range.getColumn(), cN = c0 + e.range.getNumColumns() - 1;
    const r0 = Math.max(e.range.getRow(), 2), rN = e.range.getRow() + e.range.getNumRows() - 1;

    if (name === CFG.SHEETS.PAYMENTS) {
      // React when Loan ID (B=2), Payment Date (F=6), or Amount (G=7) change.
      const touched = (c0 <= 2 && cN >= 2) || (c0 <= 6 && cN >= 6) || (c0 <= 7 && cN >= 7);
      if (!touched) return;
      const ids = {};
      for (let row = r0; row <= rN; row++) {
        fillPaymentLookup_(sh, row);                          // First, Last, DNI
        const id = String(sh.getRange(row, 2).getValue()).trim();
        if (id) ids[id] = true;
      }
      Object.keys(ids).forEach(id => { recalcLoanPayments_(id); updateLoanOutstanding_(id); });

    } else if (name === CFG.SHEETS.BORROWERS) {
      // Recompute balance when a loan's key fields change: A,F,G,I (1,6,7,9).
      const touched = [1, 6, 7, 9].some(c => c >= c0 && c <= cN);
      if (!touched) return;
      for (let row = r0; row <= rN; row++) writeOutstandingForRow_(sh, row);

    } else if (name === CFG.SHEETS.NEW_BORROWERS || name === 'New Borrowers') {
      // Move each newly-ticked applicant (Verificado?/Verified? column located by header).
      const hm = nbHeaderMap_(sh);
      const verCol = hm['Verificado?'] || hm['Verified?'];
      if (!verCol || verCol < c0 || verCol > cN) return;
      const rows = [];
      for (let row = r0; row <= rN; row++) if (sh.getRange(row, verCol).getValue() === true) rows.push(row);
      rows.sort((a, b) => b - a).forEach(r => moveVerifiedBorrower_(sh, r)); // descending so deletes don't shift
    }
  } catch (err) {
    try { SpreadsheetApp.getActiveSpreadsheet().toast('Loan Manager: ' + err.message, 'Error', 6); } catch (e2) {}
  }
}

/** Fill First Name (C), Last Name (D) + DNI (E) on a Payments row from the Loan ID (B). */
function fillPaymentLookup_(sh, row) {
  const loanId = String(sh.getRange(row, 2).getValue()).trim();
  if (!loanId) { sh.getRange(row, 3, 1, 3).clearContent(); return; }
  const loan = findLoanById_(loanId);
  if (!loan) { sh.getRange(row, 3).setValue('?'); sh.getRange(row, 4, 1, 2).clearContent(); return; }
  sh.getRange(row, 3).setValue(firstNameOf_(loan.name));
  sh.getRange(row, 4).setValue(lastNameOf_(loan.name));
  sh.getRange(row, 5).setValue(loan.dni);
}

/** Recompute "Outstanding After" (H) for every payment of a loan, in date order. */
function recalcLoanPayments_(loanId) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEETS.PAYMENTS);
  const last = sh.getLastRow();
  if (last < 2) return;
  const loan = findLoanById_(loanId);
  if (!loan) return;
  const rate = loan.term === 1 ? 0.5 : 1;
  const data = sh.getRange(2, 2, last - 1, 6).getValues(); // B..G
  const rows = [];
  data.forEach((r, i) => {
    if (String(r[0]).trim() === loanId && typeof r[5] === 'number' && r[5] !== 0)
      rows.push({ row: i + 2, date: r[4] instanceof Date ? r[4] : new Date(), amount: r[5] });
  });
  const allPays = rows.map(x => ({ date: x.date, amount: x.amount }));
  rows.forEach(x => {
    const out = computeOutstanding_(loan.principal, rate, loan.loanDate, allPays, x.date);
    sh.getRange(x.row, 8).setValue(out); // H Outstanding After
  });
}

/** Recompute the Borrowers "Outstanding (today)" value (col N) for one loan. */
function updateLoanOutstanding_(loanId) {
  const bs = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEETS.BORROWERS);
  const last = bs.getLastRow();
  if (last < 2) return;
  const ids = bs.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === loanId) { writeOutstandingForRow_(bs, i + 2); return; }
  }
}

/** Recompute Outstanding (col N) for every loan. Used by Refresh and the daily trigger. */
function updateAllOutstanding_() {
  const bs = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEETS.BORROWERS);
  const last = bs.getLastRow();
  for (let row = 2; row <= last; row++) writeOutstandingForRow_(bs, row);
}

function writeOutstandingForRow_(bs, row) {
  const loanId = String(bs.getRange(row, 1).getValue()).trim();
  const cell = bs.getRange(row, 14); // N
  if (!loanId) { cell.clearContent(); return; }
  const loan = readLoan_(bs, row);
  if (!loan.principal || !loan.term || !(loan.loanDate instanceof Date)) { cell.clearContent(); return; }
  const rate = loan.term === 1 ? 0.5 : loan.term === 2 ? 1 : null;
  if (rate === null) { cell.setValue(''); return; }
  // Penalties up to today, but ALL recorded payments count (far-future cutoff).
  const out = computeOutstanding_(loan.principal, rate, loan.loanDate,
    loanPayments_(loanId), new Date(), new Date(9999, 0, 1));
  cell.setValue(out);
}

/* ==================== CUSTOM SHEET FUNCTION ==================== */
/**
 * Current outstanding balance for a loan, with compounding overdue penalty.
 * @param {number} principal   Loan principal.
 * @param {number} term        Term in months (1 or 2).
 * @param {Date}   loanDate     Disbursement date.
 * @param {string} loanId       Loan ID (matched in payments range).
 * @param {Array}  payRange     Payments!B2:G => [LoanID, First, Last, DNI, Date, Amount]
 * @param {Date}   today        Pass TODAY() so it recalculates daily.
 * @return {number} Outstanding balance as of today.
 * @customfunction
 */
function OUTSTANDING(principal, term, loanDate, loanId, payRange, today) {
  if (!principal || !term || !loanDate) return '';
  const rate = Number(term) === 1 ? 0.5 : Number(term) === 2 ? 1 : null;
  if (rate === null) return 'INVALID TERM';
  const payments = [];
  if (payRange && payRange.length) {
    for (const row of payRange) {
      if (row[0] === loanId && row[4] instanceof Date && typeof row[5] === 'number') {
        payments.push({ date: row[4], amount: row[5] });
      }
    }
  }
  const asOf = today instanceof Date ? today : new Date();
  return computeOutstanding_(Number(principal), rate, new Date(loanDate), payments, asOf);
}

/**
 * Core ledger. Obligation starts at principal*(1+rate). After the due date,
 * each completed overdue month doubles whatever is still outstanding.
 * Payments are applied on their dates, chronologically interleaved.
 */
/**
 * @param asOf         Date up to which overdue penalties accrue.
 * @param paymentCutoff (optional) Only payments on/before this date are applied.
 *        Defaults to asOf. Pass a far-future date to apply ALL recorded payments
 *        (used for the "current outstanding" so future-dated entries still count).
 */
function computeOutstanding_(principal, rate, loanDate, payments, asOf, paymentCutoff) {
  const cutoff = (paymentCutoff instanceof Date) ? paymentCutoff : asOf;
  const dueDate = addMonths_(loanDate, rate === 0.5 ? 1 : 2);
  let outstanding = round2_(principal * (1 + rate));

  const events = [];
  // penalty (doubling) events: 1 month after due date, monthly, up to asOf
  let d = addMonths_(dueDate, 1);
  while (d.getTime() <= asOf.getTime()) {
    events.push({ date: new Date(d), type: 'penalty' });
    d = addMonths_(d, 1);
  }
  // payments dated on/before the cutoff count toward this balance
  payments.forEach(p => {
    if (new Date(p.date).getTime() <= cutoff.getTime())
      events.push({ date: new Date(p.date), type: 'payment', amount: p.amount });
  });
  // sort by date; on ties, apply payment before penalty
  events.sort((a, b) => (a.date - b.date) || (a.type === 'payment' ? -1 : 1));

  for (const e of events) {
    if (outstanding <= 0) { outstanding = 0; break; }
    if (e.type === 'payment') outstanding = round2_(outstanding - e.amount);
    else outstanding = round2_(outstanding * 2); // +100% of outstanding
  }
  return Math.max(0, round2_(outstanding));
}

/* ==================== AGREEMENT PDF + EMAIL ==================== */
function generateAgreementForActiveRow() {
  const info = activeLoanRow_();
  if (!info) return;
  const { sh, row, loan } = info;
  if (!loan.email) { alertMsg_('This loan row has no borrower email (column C).'); return; }

  const file = makeAgreementFile_(loan);
  writeAgreementLink_(sh, row, file);           // link in the Borrowers row + timestamp
  emailAgreement_(loan, file);
  alertMsg_(`Agreement PDF generated and emailed to ${loan.email}.`);
}

/** Build the PDF, save it to the Drive folder, return the Drive file. */
function makeAgreementFile_(loan) {
  const pdf = Utilities.newBlob(agreementDocHtml_(loan), 'text/html', 'agreement.html').getAs('application/pdf');
  const file = getFolder_(CFG.PDF_FOLDER).createFile(pdf);
  file.setName(`Loan Agreement ${loan.loanId} - ${loan.name}.pdf`);
  return file;
}

/** Record a "View Agreement" hyperlink + timestamp in the Borrowers row for this loan. */
function writeAgreementLink_(sh, row, file) {
  sh.getRange(row, 16).setFormula(`=HYPERLINK("${file.getUrl()}","View Agreement")`);
  sh.getRange(row, 17).setValue(new Date());
}
function writeAgreementLinkByLoanId_(loanId, file) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEETS.BORROWERS);
  const ids = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === loanId) { writeAgreementLink_(sh, i + 2, file); return; }
  }
}

function emailAgreement_(loan, file) {
  GmailApp.sendEmail(loan.email,
    `Loan Agreement — ${loan.loanId}`,
    agreementEmailBody_(loan),
    { attachments: [file.getAs('application/pdf')], name: getSetting_('Lender Name'), htmlBody: agreementEmailHtml_(loan) });
}

function agreementDocHtml_(loan) {
  const lender = getSetting_('Lender Name'), lPhone = getSetting_('Lender Phone'),
    lAddr = getSetting_('Lender Address'), juris = getSetting_('Governing Jurisdiction'),
    penalty = getSetting_('Late Penalty Clause'), footer = getSetting_('Agreement Footer');
  const rate = loan.term === 1 ? '50%' : '100%';
  return `<!DOCTYPE html><html><head><meta charset="utf-8"><style>
    body{font-family:Georgia,serif;color:#222;margin:48px;line-height:1.5;font-size:12pt}
    h1{text-align:center;font-size:20pt;letter-spacing:1px;border-bottom:2px solid #1c4587;padding-bottom:8px}
    h2{font-size:13pt;color:#1c4587;margin-top:26px;border-bottom:1px solid #ccc;padding-bottom:3px}
    table{width:100%;border-collapse:collapse;margin-top:8px}
    td{padding:6px 8px;border:1px solid #ccc;vertical-align:top}
    td.k{background:#f0f4fa;font-weight:bold;width:38%}
    .sign{margin-top:60px;width:100%}
    .sign td{border:none;border-top:1px solid #333;text-align:center;padding-top:6px;width:45%}
    .foot{margin-top:40px;font-size:10pt;color:#666;text-align:center}
  </style></head><body>
  <h1>LOAN AGREEMENT</h1>
  <p>This Loan Agreement ("Agreement") is entered into on <b>${fmtDate_(new Date())}</b> between
     <b>${esc_(lender)}</b> ("Lender") and <b>${esc_(loan.name)}</b> ("Borrower").</p>

  <h2>1. Parties</h2>
  <table>
    <tr><td class="k">Lender</td><td>${esc_(lender)} ${lPhone ? '· ' + esc_(lPhone) : ''}<br>${esc_(lAddr)}</td></tr>
    <tr><td class="k">Borrower</td><td>${esc_(loan.name)}<br>ID/DNI: ${esc_(loan.dni)} · ${esc_(loan.phone)}<br>${esc_(loan.email)}</td></tr>
  </table>

  <h2>2. Loan Terms</h2>
  <table>
    <tr><td class="k">Loan ID</td><td>${esc_(loan.loanId)}</td></tr>
    <tr><td class="k">Principal Amount</td><td>${fmtMoney_(loan.principal)}</td></tr>
    <tr><td class="k">Term</td><td>${loan.term} month(s)</td></tr>
    <tr><td class="k">Interest Rate (fixed for term)</td><td>${rate}</td></tr>
    <tr><td class="k">Interest Amount</td><td>${fmtMoney_(loan.interest)}</td></tr>
    <tr><td class="k">Total Amount Due</td><td><b>${fmtMoney_(loan.totalDue)}</b></td></tr>
    <tr><td class="k">Disbursement Date</td><td>${fmtDate_(loan.loanDate)}</td></tr>
    <tr><td class="k">Due Date</td><td><b>${fmtDate_(loan.dueDate)}</b></td></tr>
  </table>

  <h2>3. Repayment</h2>
  <p>The Borrower agrees to repay the Total Amount Due of <b>${fmtMoney_(loan.totalDue)}</b>
     on or before the Due Date of <b>${fmtDate_(loan.dueDate)}</b>.</p>

  <h2>4. Late Payment Penalty</h2>
  <p>${esc_(penalty)}</p>

  <h2>5. Governing Law</h2>
  <p>This Agreement shall be governed by the laws of ${esc_(juris)}.</p>

  <table class="sign"><tr>
    <td>_______________________<br>${esc_(lender)}<br>Lender</td>
    <td>_______________________<br>${esc_(loan.name)}<br>Borrower</td>
  </tr></table>
  <p class="foot">${esc_(footer)}</p>
  </body></html>`;
}

function agreementEmailBody_(loan) {
  return `Dear ${loan.name},\n\nPlease find attached your loan agreement (${loan.loanId}).\n\n` +
    `Principal: ${fmtMoney_(loan.principal)}\nTerm: ${loan.term} month(s) (${loan.term === 1 ? '50%' : '100%'})\n` +
    `Total due: ${fmtMoney_(loan.totalDue)}\nDue date: ${fmtDate_(loan.dueDate)}\n\n` +
    `Regards,\n${getSetting_('Lender Name')}`;
}
function agreementEmailHtml_(loan) {
  return `<p>Dear ${esc_(loan.name)},</p><p>Please find attached your loan agreement (<b>${esc_(loan.loanId)}</b>).</p>
  <ul><li>Principal: <b>${fmtMoney_(loan.principal)}</b></li>
  <li>Term: ${loan.term} month(s) (${loan.term === 1 ? '50%' : '100%'})</li>
  <li>Total due: <b>${fmtMoney_(loan.totalDue)}</b></li>
  <li>Due date: <b>${fmtDate_(loan.dueDate)}</b></li></ul>
  <p>Regards,<br>${esc_(getSetting_('Lender Name'))}</p>`;
}

/* ==================== RECORD PAYMENT + RECEIPT ==================== */
function recordPaymentDialog() {
  const ui = SpreadsheetApp.getUi();
  const idR = ui.prompt('Record payment', 'Loan ID (e.g. L-0001):', ui.ButtonSet.OK_CANCEL);
  if (idR.getSelectedButton() !== ui.Button.OK) return;
  const loanId = idR.getResponseText().trim();
  const loan = findLoanById_(loanId);
  if (!loan) { alertMsg_(`Loan ID "${loanId}" not found in Borrowers.`); return; }

  const amtR = ui.prompt('Record payment', `Borrower: ${loan.name}\nAmount paid (numbers only):`, ui.ButtonSet.OK_CANCEL);
  if (amtR.getSelectedButton() !== ui.Button.OK) return;
  const amount = Number(String(amtR.getResponseText()).replace(/[^0-9.\-]/g, ''));
  if (!amount || amount <= 0) { alertMsg_('Invalid amount.'); return; }

  const dateR = ui.prompt('Record payment', 'Payment date YYYY-MM-DD (blank = today):', ui.ButtonSet.OK_CANCEL);
  if (dateR.getSelectedButton() !== ui.Button.OK) return;
  const payDate = dateR.getResponseText().trim() ? new Date(dateR.getResponseText().trim()) : new Date();
  if (isNaN(payDate.getTime())) { alertMsg_('Invalid date.'); return; }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const psh = ss.getSheetByName(CFG.SHEETS.PAYMENTS);

  // gather existing payments for this loan (incl. this one) to compute outstanding after
  const existing = loanPayments_(loanId);
  existing.push({ date: payDate, amount: amount });
  const rate = loan.term === 1 ? 0.5 : 1;
  const outstandingAfter = computeOutstanding_(loan.principal, rate, loan.loanDate, existing, payDate);

  const payId = 'P-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyMMdd-HHmmss');
  const rowVals = [payId, loanId, firstNameOf_(loan.name), lastNameOf_(loan.name), loan.dni,
    payDate, amount, outstandingAfter, ''];
  psh.appendRow(rowVals);
  const newRow = psh.getLastRow();
  psh.getRange(newRow, 6).setNumberFormat('yyyy-mm-dd');
  psh.getRange(newRow, 7, 1, 2).setNumberFormat(CFG.CURRENCY_FMT);
  updateLoanOutstanding_(loanId); // refresh Borrowers/Summary/Statements balance immediately

  // email receipt
  if (loan.email) {
    GmailApp.sendEmail(loan.email, `Payment received — ${loanId}`,
      paymentReceiptText_(loan, payDate, amount, outstandingAfter),
      { name: getSetting_('Lender Name'), htmlBody: paymentReceiptHtml_(loan, payDate, amount, outstandingAfter) });
    psh.getRange(newRow, 9).setValue(new Date());  // I Receipt Emailed
  }
  alertMsg_(`Payment logged.\nAmount: ${fmtMoney_(amount)}\nOutstanding after: ${fmtMoney_(outstandingAfter)}` +
    (loan.email ? `\nReceipt emailed to ${loan.email}.` : '\n(No email on file — receipt not sent.)'));
}

function paymentReceiptText_(loan, date, amount, outstanding) {
  return `Dear ${loan.name},\n\nWe confirm your payment.\n\n` +
    `Loan: ${loan.loanId}\nPayment date: ${fmtDate_(date)}\nAmount paid: ${fmtMoney_(amount)}\n` +
    `Outstanding balance: ${fmtMoney_(outstanding)}\n\n` +
    (outstanding <= 0.009 ? 'Your loan is now fully paid. Thank you!\n\n' : '') +
    `Regards,\n${getSetting_('Lender Name')}`;
}
function paymentReceiptHtml_(loan, date, amount, outstanding) {
  const cleared = outstanding <= 0.009;
  return `<p>Dear ${esc_(loan.name)},</p><p>We confirm your payment for loan <b>${esc_(loan.loanId)}</b>.</p>
  <table style="border-collapse:collapse">
   <tr><td style="padding:4px 10px;border:1px solid #ccc;background:#f0f4fa"><b>Payment date</b></td><td style="padding:4px 10px;border:1px solid #ccc">${fmtDate_(date)}</td></tr>
   <tr><td style="padding:4px 10px;border:1px solid #ccc;background:#f0f4fa"><b>Amount paid</b></td><td style="padding:4px 10px;border:1px solid #ccc">${fmtMoney_(amount)}</td></tr>
   <tr><td style="padding:4px 10px;border:1px solid #ccc;background:#f0f4fa"><b>Outstanding balance</b></td><td style="padding:4px 10px;border:1px solid #ccc"><b>${fmtMoney_(outstanding)}</b></td></tr>
  </table>
  ${cleared ? '<p style="color:#38761d"><b>Your loan is now fully paid. Thank you!</b></p>' : ''}
  <p>Regards,<br>${esc_(getSetting_('Lender Name'))}</p>`;
}

/* ==================== STATEMENT (full history) ==================== */
function emailStatementForActiveRow() {
  const info = activeLoanRow_();
  if (!info) return;
  emailStatementForLoan_(info.loan, alertMsg_);
}

/** Build + send a full statement (payment history + outstanding) for one loan. */
function emailStatementForLoan_(loan, notify) {
  notify = notify || function () {};
  if (!loan.email) { notify('No borrower email on file for this loan.'); return; }
  const pays = loanPayments_(loan.loanId).sort((a, b) => a.date - b.date);
  const rate = loan.term === 1 ? 0.5 : 1;
  const outstanding = computeOutstanding_(loan.principal, rate, loan.loanDate, pays, new Date());

  let rows = pays.map(p =>
    `<tr><td style="padding:4px 10px;border:1px solid #ccc">${fmtDate_(p.date)}</td>
     <td style="padding:4px 10px;border:1px solid #ccc;text-align:right">${fmtMoney_(p.amount)}</td></tr>`).join('');
  if (!rows) rows = `<tr><td colspan="2" style="padding:4px 10px;border:1px solid #ccc">No payments recorded yet.</td></tr>`;
  const totalPaid = pays.reduce((s, p) => s + p.amount, 0);

  const html = `<p>Dear ${esc_(loan.name)},</p><p>Statement for loan <b>${esc_(loan.loanId)}</b>:</p>
   <p>Principal ${fmtMoney_(loan.principal)} · Total due ${fmtMoney_(loan.totalDue)} · Due date ${fmtDate_(loan.dueDate)}</p>
   <table style="border-collapse:collapse"><tr><th style="padding:4px 10px;border:1px solid #ccc;background:#f0f4fa">Date paid</th>
   <th style="padding:4px 10px;border:1px solid #ccc;background:#f0f4fa">Amount</th></tr>${rows}
   <tr><td style="padding:4px 10px;border:1px solid #ccc"><b>Total paid</b></td><td style="padding:4px 10px;border:1px solid #ccc;text-align:right"><b>${fmtMoney_(totalPaid)}</b></td></tr>
   <tr><td style="padding:4px 10px;border:1px solid #ccc"><b>Outstanding</b></td><td style="padding:4px 10px;border:1px solid #ccc;text-align:right"><b>${fmtMoney_(outstanding)}</b></td></tr>
   </table><p>Regards,<br>${esc_(getSetting_('Lender Name'))}</p>`;

  GmailApp.sendEmail(loan.email, `Loan statement — ${loan.loanId}`,
    `Statement for ${loan.loanId}. Total paid ${fmtMoney_(totalPaid)}. Outstanding ${fmtMoney_(outstanding)}.`,
    { name: getSetting_('Lender Name'), htmlBody: html });
  notify(`Statement emailed to ${loan.email}.`);
}

/* ==================== REFRESH (non-destructive) ==================== */
function refreshAll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  migrateSheetNames_(ss);   // rename English tabs → Spanish in place (no data loss)
  cleanupLegacySheets_(ss); // remove "New Borrowers (old …)" / "(pre-form)" leftovers
  const N = CFG.MAX_ROWS;
  const bs = ss.getSheetByName(CFG.SHEETS.BORROWERS);
  if (bs) {
    // reapply money/date/rate number formats only — never clears cell values
    bs.getRange(2, 6, N, 1).setNumberFormat(CFG.CURRENCY_FMT);
    bs.getRange(2, 8, N, 1).setNumberFormat('0%');
    bs.getRange(2, 9, N, 2).setNumberFormat('yyyy-mm-dd');
    bs.getRange(2, 11, N, 4).setNumberFormat(CFG.CURRENCY_FMT);
    // Re-apply Spanish headers (row 1 only — does not touch data).
    bs.getRange(1, 1, 1, 17).setValues([['ID Préstamo', 'Nombre del Prestatario', 'Correo', 'DNI', 'Teléfono',
      'Capital', 'Plazo (meses)', 'Tasa', 'Fecha del Préstamo', 'Fecha de Vencimiento', 'Interés',
      'Total a Pagar', 'Total Pagado', 'Saldo Pendiente (hoy)', 'Estado', 'Contrato PDF', 'Contrato Enviado']]);
  }
  const ps = ss.getSheetByName(CFG.SHEETS.PAYMENTS);
  if (ps) {
    ps.getRange(1, 1, 1, 9).setValues([['ID Pago', 'ID Préstamo', 'Nombre', 'Apellido', 'DNI',
      'Fecha de Pago', 'Monto Pagado', 'Saldo Posterior', 'Recibo Enviado']]);
  }
  updateAllOutstanding_();   // recompute every loan's Outstanding value
  // Summary is 100% formulas, so it is always safe to rebuild.
  setupSummary_(ss);
  SpreadsheetApp.flush();
  alertMsg_('Actualizado. Encabezados en español aplicados y saldos recalculados para cada préstamo.');
}

/* ============================ HELPERS ============================ */
function activeLoanRow_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getActiveSheet();
  if (sh.getName() !== CFG.SHEETS.BORROWERS) { alertMsg_('Select a loan row in the "Borrowers" sheet first.'); return null; }
  const row = sh.getActiveRange().getRow();
  if (row < 2) { alertMsg_('Select a data row (not the header).'); return null; }
  const loan = readLoan_(sh, row);
  if (!loan.loanId || !loan.name) { alertMsg_('This row is missing a Loan ID or Borrower Name.'); return null; }
  return { sh, row, loan };
}

function readLoan_(sh, row) {
  const v = sh.getRange(row, 1, 1, 15).getValues()[0];
  return {
    loanId: v[0], name: v[1], email: v[2], dni: v[3], phone: v[4],
    principal: Number(v[5]) || 0, term: Number(v[6]) || 0,
    loanDate: v[8] instanceof Date ? v[8] : new Date(v[8]),
    dueDate: v[9] instanceof Date ? v[9] : addMonths_(new Date(v[8]), Number(v[6]) || 1),
    interest: Number(v[10]) || 0, totalDue: Number(v[11]) || 0,
  };
}

function findLoanById_(loanId) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEETS.BORROWERS);
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 15).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === loanId) return readLoan_(sh, i + 2);
  }
  return null;
}

function loanPayments_(loanId) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEETS.PAYMENTS);
  if (sh.getLastRow() < 2) return [];
  const data = sh.getRange(2, 2, sh.getLastRow() - 1, 6).getValues(); // B..G: LoanID,First,Last,DNI,Date,Amount
  const out = [];
  data.forEach(r => {
    if (String(r[0]).trim() === loanId && typeof r[5] === 'number' && r[5] !== 0) {
      // A payment with no valid date still counts — assume it was made today.
      out.push({ date: r[4] instanceof Date ? r[4] : new Date(), amount: r[5] });
    }
  });
  return out;
}

function getSetting_(key) {
  const sh = getSS_().getSheetByName(CFG.SHEETS.SETTINGS);
  if (!sh || sh.getLastRow() < 2) return '';
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  for (const r of data) if (String(r[0]).trim() === key) return r[1];
  return '';
}

/** A validation rule that turns a cell into a date field with a calendar picker. */
function datePickerRule_() {
  return SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false)
    .setHelpText('Pick a date (double-click for the calendar).').build();
}

function getFolder_(name) {
  const it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}

/** The Drive folder that contains this spreadsheet (its "directory"). */
function spreadsheetParentFolder_() {
  const parents = DriveApp.getFileById(getSS_().getId()).getParents();
  return parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
}

/** The single MAIN folder (in the spreadsheet's directory) that holds all borrower subfolders. */
function borrowerMainFolder_() {
  const parent = spreadsheetParentFolder_();
  const name = 'Borrower Files';
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

/** Get (or create) a per-borrower SUBfolder inside the main folder. */
function borrowerFolder_(folderName) {
  const main = borrowerMainFolder_();
  const it = main.getFoldersByName(folderName);
  return it.hasNext() ? it.next() : main.createFolder(folderName);
}

/** Folder name from the borrower's first name, last name and DNI. */
function borrowerFolderName_(fullName, dni) {
  return sanitizeName_([firstNameOf_(fullName), lastNameOf_(fullName), String(dni || '').trim()]
    .filter(String).join(' ')) || sanitizeName_(fullName);
}

function sanitizeName_(s) { return String(s == null ? '' : s).replace(/[\/\\:*?"<>|]/g, '-').replace(/\s+/g, ' ').trim(); }

/** Pull Drive file IDs out of a Forms file-upload cell (one or more URLs). */
function extractDriveIds_(cellValue) {
  const s = String(cellValue == null ? '' : cellValue);
  const ids = [];
  let m; const re = /(?:id=|\/d\/)([-\w]{20,})/g;
  while ((m = re.exec(s))) ids.push(m[1]);
  if (!ids.length) { const re2 = /[-\w]{25,}/g; while ((m = re2.exec(s))) ids.push(m[0]); }
  return ids;
}

function getOrCreate_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

/** The bound spreadsheet — works even inside a form-submit trigger. */
function getSS_() {
  return SpreadsheetApp.getActiveSpreadsheet() ||
    SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SS_ID'));
}

function condColor_(range, text, color) {
  return SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(text).setBackground(color).setRanges([range]).build();
}

function addMonths_(date, n) {
  const d = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  const day = d.getDate();
  d.setMonth(d.getMonth() + n);
  if (d.getDate() < day) d.setDate(0); // clamp to end of month
  return d;
}

function round2_(n) { return Math.round((n + Number.EPSILON) * 100) / 100; }

function firstNameOf_(fullName) {
  return String(fullName == null ? '' : fullName).trim().split(/\s+/)[0] || '';
}
function lastNameOf_(fullName) {
  const parts = String(fullName == null ? '' : fullName).trim().split(/\s+/);
  return parts.length > 1 ? parts.slice(1).join(' ') : ''; // everything after the first name
}

function fmtMoney_(n) {
  try { return Number(n).toLocaleString(CFG.LOCALE, { style: 'currency', currency: CFG.CURRENCY_CODE }); }
  catch (e) { return '$ ' + round2_(Number(n)).toFixed(2); }
}
function fmtDate_(d) {
  if (!(d instanceof Date)) d = new Date(d);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}
function esc_(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
function alertMsg_(m) { SpreadsheetApp.getUi().alert(m); }
