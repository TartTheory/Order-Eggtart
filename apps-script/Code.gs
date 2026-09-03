var HEADERS = ['Timestamp', 'Order #', 'Items', 'Total', 'Paid via Venmo', 'Amount Received', 'Venmo Name', 'Pickup Day', 'Pickup Time', 'Pickup Place', 'Language'];
var COL = { TIMESTAMP: 1, ORDER_NUM: 2, ITEMS: 3, TOTAL: 4, PAID_VENMO: 5, AMOUNT_RECEIVED: 6, VENMO_NAME: 7, PICKUP_DAY: 8, PICKUP_TIME: 9, PICKUP_PLACE: 10, LANGUAGE: 11 };

// Keep in sync with the FLAVORS array in index.html.
var FLAVORS = [
  { id: 'original', name: 'Original Portuguese Style', price: 4 },
  { id: 'lychee', name: 'Lychee Cream Cheese Eggtart', price: 6 },
  { id: 'matcha', name: 'Matcha Red Bean Tart', price: 6 },
  { id: 'porkfloss', name: 'Pork Floss Salted Egg Tart', price: 6 },
  { id: 'oreo', name: 'Oreo Brownie Egg Tart', price: 6 },
  { id: 'taro', name: 'Purple Rice Taro Egg Tart', price: 6 },
  { id: 'durian', name: 'Roasted Durian Egg Tart', price: 6 },
  { id: 'blacksesame', name: 'Black Sesame Mochi Tart', price: 6 },
  { id: 'soymilk', name: 'Soy Milk Mochi Egg Tart', price: 6 }
];
var PICKUP_PLACE = 'Bellevue T&T Supermarket parking lot';
var PICKUP_TIME = '3:00–3:30 PM';

function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  appendOrder({
    weekLabel: data.weekLabel,
    orderNum: data.orderNum,
    items: data.items,
    total: data.total,
    paidVenmo: false,
    amountReceived: '',
    venmoName: '',
    pickupDay: data.pickupDay,
    pickupTime: data.pickupTime,
    pickupPlace: data.pickupPlace,
    lang: data.lang
  });
  return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
}

// Read-only endpoint the website polls to enforce a weekly baking cap
// (index.html's PICKUP_CAPS). ?action=remaining&week=<label> returns how many
// tarts are already committed for that week's sheet, counting every submitted
// order regardless of paid status -- a checkout reserves baking capacity the
// moment it's placed, not once payment is confirmed.
function doGet(e) {
  var action = e.parameter.action;
  if (action === 'remaining') {
    var week = e.parameter.week || '';
    return ContentService.createTextOutput(JSON.stringify({ ok: true, week: week, orderedQty: orderedQtyForWeek(week) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'unknown action' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function orderedQtyForWeek(weekLabel) {
  if (!weekLabel) return 0;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(weekLabel);
  if (!sheet) return 0;

  var summaryStart = findSummaryStart(sheet);
  var dataEndRow = summaryStart ? summaryStart - 1 : sheet.getLastRow();
  if (dataEndRow < 2) return 0;

  var rows = sheet.getRange(2, 1, dataEndRow - 1, HEADERS.length).getValues();
  var totalQty = 0;
  rows.forEach(function (r) {
    if (!r[COL.ORDER_NUM - 1]) return;
    String(r[COL.ITEMS - 1]).split('\n').forEach(function (part) {
      var m = part.trim().match(/^(\d+)x\s+/);
      if (m) totalQty += parseInt(m[1], 10);
    });
  });
  return totalQty;
}

// Shared by doPost (website checkout) and addManualOrder (Tart Theory > Add
// Manual Order menu) so every order - online or hand-entered - is written in
// the same format and the summary is rebuilt in the same call, instead of
// relying on someone remembering to check a box or match a text format.
function appendOrder(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = data.weekLabel || 'Orders';
  var sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  ensureHeaders(sheet);

  var summaryStart = findSummaryStart(sheet);
  var insertAt;
  if (summaryStart) {
    sheet.insertRowBefore(summaryStart);
    insertAt = summaryStart;
  } else {
    insertAt = sheet.getLastRow() + 1;
  }

  var newRowRange = sheet.getRange(insertAt, 1, 1, HEADERS.length);
  newRowRange.setValues([[
    new Date(), data.orderNum, data.items, data.total, !!data.paidVenmo, data.amountReceived || '', data.venmoName || '', data.pickupDay, data.pickupTime, data.pickupPlace, data.lang
  ]]);
  // insertRowBefore can inherit the SUMMARY block's bold/border styling since the
  // new row is inserted directly above it — reset to plain data-row formatting.
  newRowRange.setFontWeight('normal');
  newRowRange.setBorder(false, false, false, false, false, false);
  sheet.getRange(insertAt, COL.PAID_VENMO).insertCheckboxes();
  sheet.getRange(insertAt, COL.PAID_VENMO).setValue(!!data.paidVenmo);

  formatSheet(sheet);
  rebuildSummary(sheet);

  return sheet;
}

// Menu entry point — see onOpen(). Opens the manual-order form.
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Tart Theory')
    .addItem('Add Manual Order', 'showAddOrderDialog')
    .addToUi();
}

function showAddOrderDialog() {
  var tpl = HtmlService.createTemplateFromFile('AddOrderDialog');
  tpl.flavors = FLAVORS;
  tpl.pickupPlace = PICKUP_PLACE;
  tpl.pickupTime = PICKUP_TIME;
  var html = tpl.evaluate().setWidth(420).setHeight(560);
  SpreadsheetApp.getUi().showModalDialog(html, 'Add Manual Order');
}

// Called from AddOrderDialog.html via google.script.run.
function submitManualOrder(form) {
  var lines = [];
  var total = 0;
  FLAVORS.forEach(function (f) {
    var qty = parseInt(form['qty_' + f.id], 10) || 0;
    if (qty <= 0) return;
    lines.push(qty + 'x ' + f.name + ' ($' + f.price + ')');
    total += qty * f.price;
  });
  if (lines.length === 0) throw new Error('Add at least one item.');

  var overrideTotal = parseFloat(form.total);
  var finalTotal = !isNaN(overrideTotal) && form.total !== '' ? overrideTotal : total;

  var pickupDate = new Date(form.pickupDate + 'T00:00:00');
  var weekday = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'][pickupDate.getDay()];
  var pickupDayLabel = weekday + ' ' + (pickupDate.getMonth() + 1) + '/' + pickupDate.getDate();
  var orderNum = 'TT-' + Math.floor(1000 + Math.random() * 9000);

  appendOrder({
    weekLabel: weekLabelFor(pickupDate),
    orderNum: orderNum,
    items: lines.join(';\n'),
    total: '$' + finalTotal.toFixed(2),
    paidVenmo: !!form.paidVenmo,
    amountReceived: form.amountReceived ? '$' + parseFloat(form.amountReceived).toFixed(2) : '',
    venmoName: form.venmoName || '',
    pickupDay: pickupDayLabel,
    pickupTime: PICKUP_TIME,
    pickupPlace: PICKUP_PLACE,
    lang: form.lang || 'en'
  });

  return orderNum;
}

function weekLabelFor(date) {
  var months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  // Anchor to the Saturday that ends this date's pickup week, then label by
  // Friday-Saturday (the actual pickup days) rather than Saturday-Sunday --
  // both Friday and Saturday orders land on the same Saturday and end up
  // in the same weekly sheet, named to match. Keep in sync with
  // weekLabelFor() in index.html.
  var sat = new Date(date);
  sat.setDate(sat.getDate() + ((6 - sat.getDay() + 7) % 7));
  var fri = new Date(sat); fri.setDate(sat.getDate() - 1);
  return months[sat.getMonth()] + ' ' + fri.getDate() + '-' + sat.getDate();
}

// Installable "On change" trigger — catches row deletions (and other structural
// edits) that doPost never sees, so the summary doesn't go stale when someone
// deletes an order row, or fills in Amount Received / checks Paid via Venmo, by hand.
//
// One-time setup (per spreadsheet): open Extensions > Apps Script > Triggers
// (clock icon) > + Add Trigger > choose function "onSheetChange", event source
// "From spreadsheet", event type "On change" > Save.
function onSheetChange(e) {
  var sheet = SpreadsheetApp.getActiveSheet();
  if (!sheet) return;
  rebuildSummary(sheet);
}

function ensureHeaders(sheet) {
  var existing = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  var matches = existing.every(function (v, i) { return v === HEADERS[i]; });
  if (matches) return;
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
}

function findSummaryStart(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow === 0) return null;
  var values = sheet.getRange(1, 1, lastRow, 1).getValues();
  for (var i = 0; i < values.length; i++) {
    if (values[i][0] === 'SUMMARY') return i + 1;
  }
  return null;
}

function formatSheet(sheet) {
  var numCols = HEADERS.length;
  var lastRow = Math.max(sheet.getLastRow(), 1);
  var range = sheet.getRange(1, 1, lastRow, numCols);
  range.setHorizontalAlignment('left');
  range.setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW);
  sheet.getRange(1, COL.ITEMS, lastRow, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  sheet.getRange(1, COL.TIMESTAMP, lastRow, 1).setNumberFormat('M/d/yyyy h:mm am/pm');
  sheet.getRange(1, COL.TOTAL, lastRow, 1).setNumberFormat('@');
  sheet.getRange(1, COL.AMOUNT_RECEIVED, lastRow, 1).setNumberFormat('$#,##0.00');
  for (var c = 1; c <= numCols; c++) {
    if (c === COL.ITEMS) continue; // wrapped column — auto-width would collapse it to fit one line
    sheet.autoResizeColumn(c);
  }
}

function rebuildSummary(sheet) {
  var numCols = HEADERS.length;
  var summaryStart = findSummaryStart(sheet);
  var dataEndRow;
  var existingNotes = {};
  if (summaryStart) {
    dataEndRow = summaryStart - 1;
    var oldRange = sheet.getRange(summaryStart, 1, sheet.getMaxRows() - summaryStart + 1, numCols);
    // Preserve notes by label before wiping the block, since row positions
    // shift as flavors/orders come and go and can't be trusted to stay put.
    oldRange.getValues().forEach(function (r) {
      var label = r[0];
      var note = r[2];
      if (label && note) existingNotes[label] = note;
    });
    oldRange.clearContent();
    oldRange.clearDataValidations();
    oldRange.setBorder(false, false, false, false, false, false);
  } else {
    dataEndRow = sheet.getLastRow();
  }

  // No summary to maintain and no order rows yet — nothing to do.
  if (!summaryStart && dataEndRow < 2) return;

  var dataRows = dataEndRow >= 2 ? sheet.getRange(2, 1, dataEndRow - 1, numCols).getValues() : [];
  var flavorCounts = {};
  var flavorIncome = {};
  var flavorOrder = [];
  var totalSales = 0;
  var totalReceived = 0;
  var orderCount = 0;
  var totalTarts = 0;

  dataRows.forEach(function (r) {
    if (!r[COL.ORDER_NUM - 1]) return;
    if (r[COL.PAID_VENMO - 1] !== true) return; // unpaid orders don't count toward any summary total
    orderCount++;
    totalSales += parseFloat(String(r[COL.TOTAL - 1]).replace('$', '')) || 0;
    var received = r[COL.AMOUNT_RECEIVED - 1];
    if (received !== '' && received !== null) {
      totalReceived += parseFloat(String(received).replace('$', '')) || 0;
    }
    String(r[COL.ITEMS - 1]).split('\n').forEach(function (part) {
      part = part.trim();
      var m = part.match(/^(\d+)x\s+(.+?)\s+\(\$([\d.]+)\)/);
      if (m) {
        var qty = parseInt(m[1], 10);
        var name = m[2];
        var unitPrice = parseFloat(m[3]);
        if (!(name in flavorCounts)) flavorOrder.push(name);
        flavorCounts[name] = (flavorCounts[name] || 0) + qty;
        flavorIncome[name] = (flavorIncome[name] || 0) + qty * unitPrice;
        totalTarts += qty;
      }
    });
  });

  // Recompute even down to zero — otherwise deleting the last order rows
  // leaves a stale non-zero summary behind.
  var newSummaryStart = dataEndRow + 1;
  var restBlank = new Array(numCols - 3).fill('');
  function summaryRow(label, value) {
    return [label, value, existingNotes[label] || ''].concat(restBlank);
  }

  var rows = [];
  rows.push(['SUMMARY', '', 'Note'].concat(restBlank));
  rows.push(summaryRow('Total Orders', orderCount));
  rows.push(summaryRow('Total Sales', '$' + totalSales.toFixed(2)));
  rows.push(summaryRow('Total Received (After Venmo Fees)', '$' + totalReceived.toFixed(2)));
  rows.push(summaryRow('Total Tarts', totalTarts));
  flavorOrder.forEach(function (name) {
    rows.push(summaryRow(name, flavorCounts[name] + ' ($' + flavorIncome[name].toFixed(2) + ')'));
  });

  sheet.getRange(newSummaryStart, 1, rows.length, numCols).setValues(rows);
  sheet.getRange(newSummaryStart, 1, 1, numCols).setFontWeight('bold');
  sheet.getRange(newSummaryStart, 1, 1, numCols).setBorder(true, false, false, false, false, false);

  formatSheet(sheet);
}
