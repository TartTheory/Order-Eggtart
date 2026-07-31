var HEADERS = ['Timestamp', 'Order #', 'Items', 'Total', 'Pickup Day', 'Pickup Time', 'Pickup Place', 'Language'];

function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = data.weekLabel || 'Orders';
  var sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  }

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
    new Date(), data.orderNum, data.items, data.total, data.pickupDay, data.pickupTime, data.pickupPlace, data.lang
  ]]);
  // insertRowBefore can inherit the SUMMARY block's bold/border styling since the
  // new row is inserted directly above it — reset to plain data-row formatting.
  newRowRange.setFontWeight('normal');
  newRowRange.setBorder(false, false, false, false, false, false);

  formatSheet(sheet);
  rebuildSummary(sheet);

  return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
}

// Installable "On change" trigger — catches row deletions (and other structural
// edits) that doPost never sees, so the summary doesn't go stale when someone
// deletes an order row by hand.
//
// One-time setup (per spreadsheet): open Extensions > Apps Script > Triggers
// (clock icon) > + Add Trigger > choose function "onSheetChange", event source
// "From spreadsheet", event type "On change" > Save.
function onSheetChange(e) {
  var sheet = SpreadsheetApp.getActiveSheet();
  if (!sheet) return;
  rebuildSummary(sheet);
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
  var range = sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 1), numCols);
  range.setHorizontalAlignment('left');
  range.setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW);
  for (var c = 1; c <= numCols; c++) {
    sheet.autoResizeColumn(c);
  }
}

function rebuildSummary(sheet) {
  var numCols = HEADERS.length;
  var summaryStart = findSummaryStart(sheet);
  var dataEndRow;
  if (summaryStart) {
    dataEndRow = summaryStart - 1;
    sheet.getRange(summaryStart, 1, sheet.getMaxRows() - summaryStart + 1, numCols).clearContent();
    sheet.getRange(summaryStart, 1, sheet.getMaxRows() - summaryStart + 1, numCols).setBorder(false, false, false, false, false, false);
  } else {
    dataEndRow = sheet.getLastRow();
  }

  // No summary to maintain and no order rows yet — nothing to do.
  if (!summaryStart && dataEndRow < 2) return;

  var dataRows = dataEndRow >= 2 ? sheet.getRange(2, 1, dataEndRow - 1, numCols).getValues() : [];
  var flavorCounts = {};
  var flavorOrder = [];
  var totalRevenue = 0;
  var orderCount = 0;

  dataRows.forEach(function (row) {
    if (!row[1]) return;
    orderCount++;
    totalRevenue += parseFloat(String(row[3]).replace('$', '')) || 0;
    String(row[2]).split(';').forEach(function (part) {
      part = part.trim();
      var m = part.match(/^(\d+)x\s+(.+?)\s+\(\$/);
      if (m) {
        var qty = parseInt(m[1], 10);
        var name = m[2];
        if (!(name in flavorCounts)) flavorOrder.push(name);
        flavorCounts[name] = (flavorCounts[name] || 0) + qty;
      }
    });
  });

  // Recompute even down to zero — otherwise deleting the last order rows
  // leaves a stale non-zero summary behind.
  var newSummaryStart = dataEndRow + 1;
  var rows = [];
  rows.push(['SUMMARY', '', '', '', '', '', '', '']);
  rows.push(['Total Orders', orderCount, '', '', '', '', '', '']);
  rows.push(['Total Revenue', '$' + totalRevenue.toFixed(2), '', '', '', '', '', '']);
  flavorOrder.forEach(function (name) {
    rows.push([name, flavorCounts[name], '', '', '', '', '', '']);
  });

  sheet.getRange(newSummaryStart, 1, rows.length, numCols).setValues(rows);
  sheet.getRange(newSummaryStart, 1, 1, numCols).setFontWeight('bold');
  sheet.getRange(newSummaryStart, 1, 1, numCols).setBorder(true, false, false, false, false, false);

  formatSheet(sheet);
}
