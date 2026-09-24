// ============================================================
//  SplitEase 4-Person — Google Apps Script
//  Rename people in the PEOPLE array below — everything else
//  (Sheets, balances, emails) follows automatically.
//
//  Deploy: Extensions → Apps Script → paste this
//          → Deploy → New deployment → Web app
//          → Execute as: Me | Who has access: Anyone
//          → Copy the /exec URL into EXEC_URL at the top of index.html
// ============================================================

var SHEET_ID = '11go52W0cRlwYDnPlJ6XVDk7WFr2Z8pFGRhIJmL5_jf8'; // your Google Sheet — script always writes here explicitly
var PEOPLE = ['Ali Asgar', 'Hakim', 'Taher', 'Yusuf'];

function getSheet_() {
  return SpreadsheetApp.openById(SHEET_ID);
}

var COLORS = {
  expHeader: '#E8A895', expHeaderFont: '#5C3A2E',
  expRowA: '#FCEEE6', expRowB: '#E6F5EE', expRowC: '#F0F1FC', expRowD: '#FFF6DD',
  settleHeader: '#A8D8BE', settleHeaderFont: '#2E5C43',
  settleRow1: '#EEF8F1', settleRow2: '#DEF0E4'
};

function doGet(e) {
  var action = e.parameter.action;
  if (action === 'data') return jsonOut({ status: 'ok', data: getAllData() });
  return jsonOut({ status: 'error', message: 'Unknown or missing action. Use ?action=data' });
}

function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var result;
    if (payload.type === 'expense') result = saveExpense(payload.data);
    else if (payload.type === 'settlement') result = saveSettlement(payload.data);
    else throw new Error('Unknown payload type: ' + payload.type);
    return jsonOut({ status: 'ok', data: result });
  } catch (err) {
    return jsonOut({ status: 'error', message: err.toString() });
  }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function getAllData() {
  var ss = getSheet_();
  var expenses = readExpenses(ss);
  var settlements = readSettlements(ss);
  return {
    people: PEOPLE,
    expenses: expenses,
    settlements: settlements,
    balances: computeBalances(expenses, settlements)
  };
}

function normalizeDateCell(val) {
  if (val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return val;
}
function normalizeTimeCell(val) {
  if (val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), 'HH:mm');
  return val;
}
function forceTextColumn(sheet, col) {
  sheet.getRange(2, col, Math.max(sheet.getMaxRows() - 1, 1), 1).setNumberFormat('@');
}

// Expenses columns: ID, Date, Time, Place, Amount, PaidBy, Share_A..D, Amount_A..D, Notes, CreatedAt
function readExpenses(ss) {
  var sheet = ss.getSheetByName('Expenses');
  if (!sheet || sheet.getLastRow() < 2) return [];
  var n = PEOPLE.length;
  var totalCols = 6 + n + n + 2; // id,date,time,place,amount,payer + shares + amounts + notes + createdAt
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, totalCols).getValues();
  return rows.map(function (r) {
    var shares = {}, amounts = {};
    for (var i = 0; i < n; i++) shares[PEOPLE[i]] = r[6 + i];
    for (var i = 0; i < n; i++) amounts[PEOPLE[i]] = r[6 + n + i];
    return {
      id: r[0], date: normalizeDateCell(r[1]), time: normalizeTimeCell(r[2]),
      place: r[3], amount: r[4], payer: r[5],
      shares: shares, amounts: amounts,
      note: r[6 + 2 * n], createdAt: r[7 + 2 * n]
    };
  }).sort(function (a, b) {
    return new Date(b.date + ' ' + (b.time || '00:00')) - new Date(a.date + ' ' + (a.time || '00:00'));
  });
}

function readSettlements(ss) {
  var sheet = ss.getSheetByName('Settlements');
  if (!sheet || sheet.getLastRow() < 2) return [];
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();
  return rows.map(function (r) {
    return {
      id: r[0], date: normalizeDateCell(r[1]), time: normalizeTimeCell(r[2]),
      from: r[3], to: r[4], amount: r[5], mode: r[6], note: r[7], createdAt: r[8]
    };
  }).sort(function (a, b) {
    return new Date(b.date + ' ' + (b.time || '00:00')) - new Date(a.date + ' ' + (a.time || '00:00'));
  });
}

// net[person] > 0 => person is owed that much overall; < 0 => person owes that much overall
function computeBalances(expenses, settlements) {
  var net = {};
  PEOPLE.forEach(function (p) { net[p] = 0; });

  expenses.forEach(function (x) {
    net[x.payer] = (net[x.payer] || 0) + Number(x.amount);
    PEOPLE.forEach(function (p) {
      net[p] = (net[p] || 0) - (Number(x.amounts[p]) || 0);
    });
  });

  settlements.forEach(function (s) {
    net[s.from] = (net[s.from] || 0) + Number(s.amount);
    net[s.to] = (net[s.to] || 0) - Number(s.amount);
  });

  Object.keys(net).forEach(function (p) { net[p] = Math.round(net[p] * 100) / 100; });

  return { net: net, suggestions: simplifyDebts(net) };
}

// Greedy debt simplification: minimum number of settle-up transactions
// that would bring everyone's net balance to zero.
function simplifyDebts(net) {
  var creditors = [], debtors = [];
  Object.keys(net).forEach(function (p) {
    if (net[p] > 0.01) creditors.push({ p: p, amt: net[p] });
    else if (net[p] < -0.01) debtors.push({ p: p, amt: -net[p] });
  });
  creditors.sort(function (a, b) { return b.amt - a.amt; });
  debtors.sort(function (a, b) { return b.amt - a.amt; });

  var suggestions = [];
  var i = 0, j = 0;
  while (i < debtors.length && j < creditors.length) {
    var pay = Math.min(debtors[i].amt, creditors[j].amt);
    pay = Math.round(pay * 100) / 100;
    if (pay > 0.01) {
      suggestions.push({ from: debtors[i].p, to: creditors[j].p, amount: pay });
    }
    debtors[i].amt -= pay;
    creditors[j].amt -= pay;
    if (debtors[i].amt <= 0.01) i++;
    if (creditors[j].amt <= 0.01) j++;
  }
  return suggestions;
}

function saveExpense(d) {
  var ss = getSheet_();
  var sheet = ss.getSheetByName('Expenses');
  var n = PEOPLE.length;

  if (!sheet) {
    sheet = ss.insertSheet('Expenses');
    var headers = ['ID', 'Date', 'Time', 'Place / Description', 'Total Amount', 'Paid By'];
    PEOPLE.forEach(function (p) { headers.push(p + ' Share'); });
    PEOPLE.forEach(function (p) { headers.push(p + ' Amount'); });
    headers.push('Notes', 'Created At');
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setBackground(COLORS.expHeader).setFontColor(COLORS.expHeaderFont).setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(4, 200);
  }

  if (!d.id) d.id = 'E' + new Date().getTime();
  if (!d.createdAt) d.createdAt = new Date().toISOString();
  if (!d.time) throw new Error('Local time is required');
  if (PEOPLE.indexOf(d.payer) < 0) throw new Error('Unknown payer: ' + d.payer);

  forceTextColumn(sheet, 2); forceTextColumn(sheet, 3);

  var totalShares = 0;
  PEOPLE.forEach(function (p) { totalShares += Number(d.shares[p]) || 0; });
  if (totalShares <= 0) throw new Error('Total shares must be greater than 0');

  var amounts = {}, runningTotal = 0;
  PEOPLE.forEach(function (p, idx) {
    if (idx === PEOPLE.length - 1) {
      amounts[p] = Math.round((Number(d.amount) - runningTotal) * 100) / 100; // last person absorbs rounding remainder
    } else {
      var amt = Math.round((Number(d.amount) * (Number(d.shares[p]) || 0) / totalShares) * 100) / 100;
      amounts[p] = amt;
      runningTotal += amt;
    }
  });

  var row = [d.id, d.date, d.time, d.place, Number(d.amount), d.payer];
  PEOPLE.forEach(function (p) { row.push(Number(d.shares[p]) || 0); });
  PEOPLE.forEach(function (p) { row.push(amounts[p]); });
  row.push(d.note || '', d.createdAt);

  sheet.appendRow(row);
  var lastRow = sheet.getLastRow();
  var payerIdx = PEOPLE.indexOf(d.payer);
  var rowColors = [COLORS.expRowA, COLORS.expRowB, COLORS.expRowC, COLORS.expRowD];
  sheet.getRange(lastRow, 1, 1, row.length).setBackground(rowColors[payerIdx % rowColors.length]).setFontSize(10);
  sheet.getRange(lastRow, 5).setFontWeight('bold');

  return getAllData();
}

function saveSettlement(d) {
  var ss = getSheet_();
  var sheet = ss.getSheetByName('Settlements');

  if (!sheet) {
    sheet = ss.insertSheet('Settlements');
    var headers = ['ID', 'Date', 'Time', 'From', 'To', 'Amount', 'Mode of Payment', 'Notes', 'Created At'];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setBackground(COLORS.settleHeader).setFontColor(COLORS.settleHeaderFont).setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1, 9, 140);
  }

  if (!d.id) d.id = 'S' + new Date().getTime();
  if (!d.createdAt) d.createdAt = new Date().toISOString();
  if (PEOPLE.indexOf(d.from) < 0 || PEOPLE.indexOf(d.to) < 0) throw new Error('Unknown person in settlement');

  forceTextColumn(sheet, 2); forceTextColumn(sheet, 3);

  sheet.appendRow([d.id, d.date, d.time, d.from, d.to, Number(d.amount), d.mode, d.note || '', d.createdAt]);
  var lastRow = sheet.getLastRow();
  var isEven = (lastRow % 2 === 0);
  sheet.getRange(lastRow, 1, 1, 9).setBackground(isEven ? COLORS.settleRow2 : COLORS.settleRow1).setFontSize(10);
  sheet.getRange(lastRow, 6).setFontWeight('bold');

  return getAllData();
}

function testSetup() {
  Logger.log('Sheet: ' + getSheet_().getName());
  Logger.log('People: ' + PEOPLE.join(', '));
}
