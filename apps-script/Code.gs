var HEADERS = ['Timestamp', 'Order #', 'Items', 'Total', 'Paid via Venmo', 'Amount Received', 'Venmo Name', 'Pickup Day', 'Pickup Time', 'Pickup Place', 'Language'];
var COL = { TIMESTAMP: 1, ORDER_NUM: 2, ITEMS: 3, TOTAL: 4, PAID_VENMO: 5, AMOUNT_RECEIVED: 6, VENMO_NAME: 7, PICKUP_DAY: 8, PICKUP_TIME: 9, PICKUP_PLACE: 10, LANGUAGE: 11 };

// Keep in sync with the FLAVORS array in index.html.
var FLAVORS = [
  { id: 'original', name: 'Original Portuguese Style', price: 4 },
  { id: 'lychee', name: 'Lychee Cream Cheese Eggtart', price: 6 },
  { id: 'persimmon', name: 'Persimmon Egg Tart', price: 6 },
  { id: 'matcha', name: 'Matcha Red Bean Tart', price: 6 },
  { id: 'tiramisu', name: 'Tiramisu Egg Tart', price: 6 },
  { id: 'porkfloss', name: 'Pork Floss Salted Egg Tart', price: 6 },
  { id: 'oreo', name: 'Oreo Brownie Egg Tart', price: 6 },
  { id: 'taro', name: 'Purple Rice Taro Egg Tart', price: 6 },
  { id: 'durian', name: 'Roasted Durian Egg Tart', price: 6 },
  { id: 'blacksesame', name: 'Black Sesame Mochi Tart', price: 6 },
  { id: 'soymilk', name: 'Soy Milk Mochi Egg Tart', price: 6 },
  { id: 'osmanthus', name: 'Osmanthus Rice Wine Egg Tart', price: 6 },
  { id: 'lavachoc', name: 'Lava Chocolate Egg Tart', price: 6 },
  { id: 'chestnut', name: 'Chestnut Egg Tart', price: 6 },
  { id: 'lavacheese', name: 'Lava Cheese Tart', price: 6 },
  { id: 'corn', name: 'Corn Egg Tart', price: 6 },
  // Halloween gift boxes (limited, ends 10/31). Remove after the holiday.
  { id: 'hbox4', name: 'Halloween Box of 4', price: 30 },
  { id: 'hbox6', name: 'Halloween Box of 6', price: 40 }
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

// A row is an order if it has items in it. It used to require an Order # as
// well, which silently dropped any order typed straight into the sheet by hand
// (no Order # in column B) from both the summary and the weekly cap count.
function isOrderRow(r) {
  return String(r[COL.ITEMS - 1] || '').trim() !== '';
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
    if (!isOrderRow(r)) return;
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
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    return appendOrderLocked_(data);
  } finally {
    lock.releaseLock();
  }
}

function appendOrderLocked_(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = data.weekLabel || 'Orders';
  var sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName, 0); // newest week is the first tab
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
  rebuildSummaryLocked_(sheet);

  return sheet;
}

// Menu entry point — see onOpen(). Opens the manual-order form.
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Tart Theory')
    .addItem('Add Manual Order', 'showAddOrderDialog')
    .addItem('Refresh Summary (this sheet)', 'refreshActiveSummary')
    .addToUi();
}

// Rebuilds the summary on the sheet you're looking at, e.g. to clean up old formatting.
function refreshActiveSummary() {
  rebuildSummary(SpreadsheetApp.getActiveSheet());
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
  // A hand-typed note: just remember it. Rebuilding here would rewrite the very
  // cells being edited and could drop what was typed.
  if (e && e.changeType === 'EDIT' && isNoteEdit(sheet)) {
    var lock = LockService.getDocumentLock();
    try { lock.waitLock(30000); } catch (err) { return; }
    try {
      var start = findSummaryStart(sheet);
      var store = loadNotesStore(sheet);
      readNotesFromBlock(sheet.getRange(start, 1, sheet.getMaxRows() - start + 1, HEADERS.length).getValues(), store);
      saveNotesStore(sheet, store);
    } finally { lock.releaseLock(); }
    return;
  }
  rebuildSummary(sheet);
}

function isNoteEdit(sheet) {
  var r = sheet.getActiveRange();
  var start = findSummaryStart(sheet);
  return !!(r && start && r.getRow() >= start && r.getColumn() >= 3);
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

// Pickup days that share a weekly sheet (see weekLabelFor) but should get
// their own tart breakdown, since Friday and Saturday are baked/packed as
// separate batches. Matched against the start of the Pickup Day cell (e.g.
// "Fri 9/26", "Sat 9/27").
var PICKUP_DAYS = ['Friday', 'Saturday'];

function pickupDayBucket(pickupDayLabel) {
  var s = String(pickupDayLabel || '');
  if (s.indexOf('Fri') === 0) return 'Friday';
  if (s.indexOf('Sat') === 0) return 'Saturday';
  return null;
}

// ---- Hand-typed notes on the summary block ----
// Anything typed to the right of a summary row's value (the Note column and
// beyond) is kept by row key and written back after every rebuild. Flavor
// rows are keyed by pickup day ("Friday|Name") because Friday and Saturday
// list the same flavors. Notes are also copied into the document properties,
// so a note on a flavor row survives that flavor dropping out of the summary
// for a while (an order unchecked as paid, a row deleted) and comes back when
// the row does.
function notesStoreKey(sheet) { return 'notes:' + sheet.getName(); }

function loadNotesStore(sheet) {
  try {
    var raw = PropertiesService.getDocumentProperties().getProperty(notesStoreKey(sheet));
    return raw ? JSON.parse(raw) : {};
  } catch (err) { return {}; }
}

function saveNotesStore(sheet, store) {
  try {
    PropertiesService.getDocumentProperties().setProperty(notesStoreKey(sheet), JSON.stringify(store));
  } catch (err) { /* notes still live in the sheet itself */ }
}

function dayOfHeaderLabel(label) {
  for (var i = 0; i < PICKUP_DAYS.length; i++) {
    if (label === PICKUP_DAYS[i].toUpperCase() + ' TARTS') return PICKUP_DAYS[i];
  }
  return null;
}

function hasText(values) {
  return values.some(function (v) { return String(v === null || v === undefined ? '' : v) !== ''; });
}

// Reads hand-typed notes (Note column onward) out of the summary block into
// `store`, keyed by row (flavor rows by "Day|Name"). A row whose note was
// cleared on purpose drops out of the store.
function readNotesFromBlock(block, store) {
  var section = '';
  block.forEach(function (r) {
    var label = r[0];
    if (!label || label === 'SUMMARY') return;
    var header = dayOfHeaderLabel(label);
    if (header) section = header;
    var key = (!header && section) ? section + '|' + label : label;
    var tail = r.slice(2);
    if (hasText(tail)) store[key] = tail;
    else delete store[key];
  });
  return store;
}

// Only one rebuild at a time: an order, its row-insert change trigger and an
// edit can all fire together, and overlapping rebuilds used to overwrite each
// other (and the notes).
function rebuildSummary(sheet) {
  var lock = LockService.getDocumentLock();
  try { lock.waitLock(30000); } catch (err) { return; }
  try {
    rebuildSummaryLocked_(sheet);
  } finally {
    lock.releaseLock();
  }
}

function rebuildSummaryLocked_(sheet) {
  var numCols = HEADERS.length;
  var summaryStart = findSummaryStart(sheet);
  var dataEndRow;
  var existingNotes = loadNotesStore(sheet); // key -> cells from the Note column onward
  var existingCost = '';
  var oldBlock = [];
  if (summaryStart) {
    dataEndRow = summaryStart - 1;
    var oldRange = sheet.getRange(summaryStart, 1, sheet.getMaxRows() - summaryStart + 1, numCols);
    // Preserve notes by row key, and the hand-entered Cost value -- row
    // positions shift as flavors/orders come and go and can't be trusted to
    // stay put. Nothing is cleared yet: if the layout is unchanged only the
    // numbers are rewritten, so notes are never touched.
    oldBlock = oldRange.getValues();
    readNotesFromBlock(oldBlock, existingNotes);
    oldBlock.forEach(function (r) { if (r[0] === 'Cost') existingCost = r[1]; });
  } else {
    dataEndRow = sheet.getLastRow();
  }

  // No summary to maintain and no order rows yet — nothing to do.
  if (!summaryStart && dataEndRow < 2) return;

  var dataRows = dataEndRow >= 2 ? sheet.getRange(2, 1, dataEndRow - 1, numCols).getValues() : [];
  var totalSales = 0;
  var totalReceived = 0;
  var orderCount = 0;
  var totalTarts = 0;

  var byDay = {};
  PICKUP_DAYS.forEach(function (d) { byDay[d] = { counts: {}, income: {}, order: [], tarts: 0 }; });

  dataRows.forEach(function (r) {
    if (!isOrderRow(r)) return;
    if (r[COL.PAID_VENMO - 1] !== true) return; // unpaid orders don't count toward any summary total
    orderCount++;
    totalSales += parseFloat(String(r[COL.TOTAL - 1]).replace('$', '')) || 0;
    var received = r[COL.AMOUNT_RECEIVED - 1];
    if (received !== '' && received !== null) {
      totalReceived += parseFloat(String(received).replace('$', '')) || 0;
    }
    var day = byDay[pickupDayBucket(r[COL.PICKUP_DAY - 1])]; // undefined for a day outside PICKUP_DAYS
    String(r[COL.ITEMS - 1]).split('\n').forEach(function (part) {
      part = part.trim();
      var m = part.match(/^(\d+)x\s+(.+?)\s+\(\$([\d.]+)\)/);
      if (m) {
        var qty = parseInt(m[1], 10);
        var name = m[2];
        var unitPrice = parseFloat(m[3]);
        totalTarts += qty;
        if (day) {
          if (!(name in day.counts)) day.order.push(name);
          day.counts[name] = (day.counts[name] || 0) + qty;
          day.income[name] = (day.income[name] || 0) + qty * unitPrice;
          day.tarts += qty;
        }
      }
    });
  });

  // Recompute even down to zero — otherwise deleting the last order rows
  // leaves a stale non-zero summary behind.
  var newSummaryStart = dataEndRow + 1;
  var noteCols = numCols - 2;
  var restBlank = new Array(noteCols - 1).fill('');
  function notesFor(key) {
    var t = existingNotes[key];
    return t ? t.concat(new Array(Math.max(0, noteCols - t.length)).fill('')).slice(0, noteCols) : new Array(noteCols).fill('');
  }
  function summaryRow(label, value, noteKey) {
    return [label, value].concat(notesFor(noteKey || label));
  }

  // Cost is hand-entered (ingredient cost is a per-batch expense, not
  // something tied to any one order), so it's carried over as-is rather than
  // recomputed. Profit is Total Received (i.e. after Venmo fees) minus Cost.
  var cost = parseFloat(String(existingCost).replace('$', '')) || 0;
  var profit = totalReceived - cost;

  var rows = [];
  var headerRowIndexes = []; // indexes into rows[] that get the same bold+border as the SUMMARY row
  rows.push(['SUMMARY', '', 'Note'].concat(restBlank));
  rows.push(summaryRow('Total Orders', orderCount));
  rows.push(summaryRow('Total Sales', '$' + totalSales.toFixed(2)));
  rows.push(summaryRow('Total Received (After Venmo Fees)', '$' + totalReceived.toFixed(2)));
  rows.push(summaryRow('Total Tarts', totalTarts));
  rows.push(summaryRow('Cost', existingCost));
  rows.push(summaryRow('Profit', '$' + profit.toFixed(2)));

  // Per-day tart breakdown -- Friday and Saturday orders share this sheet
  // (weekLabelFor) but are baked/packed as separate batches, so each gets
  // its own flavor breakdown instead of one merged list.
  PICKUP_DAYS.forEach(function (d) {
    var bucket = byDay[d];
    headerRowIndexes.push(rows.length);
    rows.push(summaryRow(d.toUpperCase() + ' TARTS', bucket.tarts));
    bucket.order.forEach(function (name) {
      rows.push(summaryRow(name, bucket.counts[name], d + '|' + name));
    });
  });

  var sameLayout = !!summaryStart && summaryStart === newSummaryStart
    && rows.every(function (r, i) { return oldBlock[i] && String(oldBlock[i][0]) === String(r[0]); })
    && oldBlock.slice(rows.length).every(function (r) { return String(r[0]) === ''; });
  var block = sheet.getRange(newSummaryStart, 1, rows.length, numCols);
  if (sameLayout) {
    // Same rows as before: refresh the numbers only, leave the Note cells alone.
    sheet.getRange(newSummaryStart, 2, rows.length, 1).setValues(rows.map(function (r) { return [r[1]]; }));
  } else {
    if (summaryStart) {
      var oldRange2 = sheet.getRange(summaryStart, 1, sheet.getMaxRows() - summaryStart + 1, numCols);
      oldRange2.clearContent();
      oldRange2.clearDataValidations();
      oldRange2.setBorder(false, false, false, false, false, false);
    }
    block.setValues(rows);
  }
  saveNotesStore(sheet, existingNotes);
  // Reset to regular first: rows shift, so cells that used to hold a bold header
  // would otherwise leave flavor names bold. Only the headers are bold.
  block.setFontWeight('normal');
  sheet.getRange(newSummaryStart, 1, 1, numCols).setFontWeight('bold');
  sheet.getRange(newSummaryStart, 1, 1, numCols).setBorder(true, false, false, false, false, false);
  headerRowIndexes.forEach(function (idx) {
    var r = sheet.getRange(newSummaryStart + idx, 1, 1, numCols);
    r.setFontWeight('bold');
    r.setBorder(true, false, false, false, false, false);
  });

  formatSheet(sheet);
}
