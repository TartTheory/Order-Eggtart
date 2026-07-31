var HEADERS = ['Timestamp', 'Order #', 'Items', 'Total', 'Paid via Venmo', 'Amount Received', 'Venmo Name', 'Pickup Day', 'Pickup Time', 'Pickup Place', 'Language'];
var COL = { TIMESTAMP: 1, ORDER_NUM: 2, ITEMS: 3, TOTAL: 4, PAID_VENMO: 5, AMOUNT_RECEIVED: 6, VENMO_NAME: 7, PICKUP_DAY: 8, PICKUP_TIME: 9, PICKUP_PLACE: 10, LANGUAGE: 11 };

function doPost(e) {
  var data = JSON.parse(e.postData.contents);
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
    new Date(), data.orderNum, data.items, data.total, false, '', '', data.pickupDay, data.pickupTime, data.pickupPlace, data.lang
  ]]);
  // insertRowBefore can inherit the SUMMARY block's bold/border styling since the
  // new row is inserted directly above it — reset to plain data-row formatting.
  newRowRange.setFontWeight('normal');
  newRowRange.setBorder(false, false, false, false, false, false);
  sheet.getRange(insertAt, COL.PAID_VENMO).insertCheckboxes();

  formatSheet(sheet);
  rebuildSummary(sheet);

  return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
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
