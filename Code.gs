/**
 * fruitVeggieComparator — Google Apps Script
 *
 * Rebuilt from the Python Streamlit project at:
 *   https://github.com/amirvolinsky/fruitVeggieComparator
 *
 * Weekly pricing flow:
 *   1. Each week: dealer sends a pricing Excel (מקט, פריט, מחירון)
 *   2. User uploads via menu → tagged with the relevant week (YYYY-WW)
 *   3. Pricing sheet accumulates all weeks
 *   4. End of month: dealer sends expense report → uploaded via menu
 *   5. Comparison joins on (שבוע, מקט) — one price per SKU per week
 *
 * Status logic (from Python compare_prices / determine_status_and_diff):
 *   |diff| < 1e-6        → ✅ תואם
 *   otherwise             → ❌ לא תואם  (min diff across price columns)
 *   no pricing entry      → 🟡 חסר במחירון
 */

// ──────────────────────────────────────────────
// CONFIGURATION
// ──────────────────────────────────────────────

var CONFIG = {
  SHEET_EXPENSES:    'Expenses',
  SHEET_PRICING:     'Pricing',
  SHEET_COMPARISON:  'Comparison',
  EXACT_TOLERANCE:   1e-6
};

// ──────────────────────────────────────────────
// MENU
// ──────────────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🍎 השוואת מחירים')
    .addItem('העלה מחירון שבועי', 'showUploadPricingDialog')
    .addItem('העלה דוח הוצאות', 'showUploadExpensesDialog')
    .addItem('הרץ השוואה', 'compareWeeklyPrices')
    .addSeparator()
    .addItem('מדריך שימוש', 'showGuide')
    .addToUi();
}

// ──────────────────────────────────────────────
// SIDEBAR — Hebrew guide
// ──────────────────────────────────────────────

function showGuide() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('מדריך שימוש');
  SpreadsheetApp.getUi().showSidebar(html);
}

// ──────────────────────────────────────────────
// UPLOAD PRICING — dialog
// ──────────────────────────────────────────────

function showUploadPricingDialog() {
  var html = HtmlService.createHtmlOutputFromFile('UploadPricing')
    .setWidth(480)
    .setHeight(360);
  SpreadsheetApp.getUi().showModalDialog(html, 'העלאת מחירון שבועי');
}

// ──────────────────────────────────────────────
// UPLOAD EXPENSES — dialog
// ──────────────────────────────────────────────

function showUploadExpensesDialog() {
  var html = HtmlService.createHtmlOutputFromFile('UploadExpenses')
    .setWidth(480)
    .setHeight(320);
  SpreadsheetApp.getUi().showModalDialog(html, 'העלאת דוח הוצאות');
}

/**
 * Called from UploadExpenses.html after client-side Excel parsing.
 * Receives a 2D array (including header row) and writes to Expenses sheet.
 * Replaces existing data each time (expense report is a full monthly file).
 */
function importExpensesData(rows) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_EXPENSES);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_EXPENSES);
  }

  if (!rows || rows.length === 0) {
    return { success: false, message: 'לא נמצאו שורות בקובץ.' };
  }

  sheet.clearContents();
  sheet.clearFormats();

  // Write all rows including header
  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);

  // Format header row
  sheet.getRange(1, 1, 1, rows[0].length)
    .setFontWeight('bold')
    .setBackground('#F1F5F9');
  sheet.setFrozenRows(1);
  sheet.setRightToLeft(true);

  for (var col = 1; col <= rows[0].length; col++) {
    sheet.autoResizeColumn(col);
  }

  return {
    success: true,
    message: 'נטענו ' + (rows.length - 1) + ' שורות הוצאות.'
  };
}

/**
 * Called from UploadPricing.html after client-side Excel parsing.
 * Receives the week string and a 2D array of rows [מקט, פריט, מחירון].
 * Appends to the Pricing sheet with the week tag.
 */
function importPricingData(week, rows) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_PRICING);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_PRICING);
    sheet.appendRow(['שבוע', 'מקט', 'פריט', 'מחירון']);
    sheet.getRange(1, 1, 1, 4)
      .setFontWeight('bold')
      .setBackground('#F1F5F9');
    sheet.setFrozenRows(1);
  }

  if (!rows || rows.length === 0) {
    return { success: false, message: 'לא נמצאו שורות בקובץ.' };
  }

  // Detect header row — skip it if present
  var startIdx = 0;
  var firstCell = String(rows[0][0]).replace(/[^A-Za-z\u0590-\u05FF0-9]/g, '');
  if (firstCell === 'מקט' || firstCell === 'מקט') {
    startIdx = 1;
  }

  var output = [];
  var skipped = 0;
  for (var i = startIdx; i < rows.length; i++) {
    var row = rows[i];
    var sku = row[0];
    if (sku === null || sku === undefined || String(sku).trim() === '') {
      skipped++;
      continue;
    }
    var item = row.length > 1 ? row[1] : '';
    var price = row.length > 2 ? row[2] : '';
    output.push([week, sku, item, price]);
  }

  if (output.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, output.length, 4).setValues(output);
  }

  return {
    success: true,
    message: 'נוספו ' + output.length + ' פריטים לשבוע ' + week +
             (skipped > 0 ? ' (' + skipped + ' שורות ללא מקט דולגו)' : '')
  };
}

// ──────────────────────────────────────────────
// HELPER — ISO week string YYYY-WW
// ──────────────────────────────────────────────

function getISOWeek(date) {
  if (!(date instanceof Date) || isNaN(date.getTime())) return null;
  var d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay() || 7));
  var yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  var weekNo = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
  var year = d.getUTCFullYear();
  return year + '-' + (weekNo < 10 ? '0' : '') + weekNo;
}

/**
 * Returns current ISO week string. Called from UploadPricing.html.
 */
function getCurrentWeek() {
  return getISOWeek(new Date());
}

// ──────────────────────────────────────────────
// HELPER — read sheet as array-of-objects
// ──────────────────────────────────────────────

function sheetToObjects(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      obj[String(headers[j]).trim()] = data[i][j];
    }
    rows.push(obj);
  }
  return rows;
}

// ──────────────────────────────────────────────
// HELPER — find column by keywords
// Mirrors Python _find_column() / _sanitize_column_name()
// ──────────────────────────────────────────────

function sanitizeColName(name) {
  return String(name).replace(/[^A-Za-z\u0590-\u05FF0-9]/g, '');
}

function findColumn(headers, keywords) {
  var sanitizedKeywords = keywords.map(function(k) { return sanitizeColName(k); });
  for (var i = 0; i < headers.length; i++) {
    var sanitized = sanitizeColName(headers[i]);
    for (var j = 0; j < sanitizedKeywords.length; j++) {
      if (sanitized.indexOf(sanitizedKeywords[j]) !== -1) {
        return headers[i];
      }
    }
  }
  return null;
}

// ──────────────────────────────────────────────
// HELPER — safe number
// ──────────────────────────────────────────────

function toNum(val) {
  if (val === '' || val === null || val === undefined) return NaN;
  return Number(val);
}

// ──────────────────────────────────────────────
// CORE — Weekly Comparison
// ──────────────────────────────────────────────
//
// Adapted from Python compare_prices():
//
// Python:  merge on (מקט, תאריך) with left join
// Sheets:  merge on (שבוע, מקט) — one price per SKU per week
//
// For each expense row:
//   1. Convert תאריך → שבוע (YYYY-WW)
//   2. Look up dealerMap[שבוע|מקט] → single expected price
//   3. Compare actual vs expected (both מחיר לפני/אחרי מע"מ checked)
//   4. Status: ✅ תואם / ❌ לא תואם / 🟡 חסר במחירון

function compareWeeklyPrices() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  // ── 1. Read pricing into map: "week|sku" → מחירון ──────
  var pricingRows = sheetToObjects(CONFIG.SHEET_PRICING);
  if (pricingRows.length === 0) {
    ui.alert('גיליון Pricing ריק או חסר.\nהעלה מחירון דרך התפריט.');
    return;
  }

  var pricingHeaders = Object.keys(pricingRows[0]);
  var priceCol = findColumn(pricingHeaders, ['מחירון', 'מחיר']);
  if (!priceCol) {
    ui.alert('לא נמצאה עמודת מחירון בגיליון Pricing.');
    return;
  }

  var dealerMap = {};  // "week|sku" → price
  for (var p = 0; p < pricingRows.length; p++) {
    var pr = pricingRows[p];
    var pWeek = String(pr['שבוע']).trim();
    var pSku = String(pr['מקט']).replace(/\.0$/, '').trim();
    if (!pWeek || !pSku || pSku === '' || pSku === 'undefined') continue;
    var key = pWeek + '|' + pSku;
    dealerMap[key] = toNum(pr[priceCol]);
  }

  // ── 2. Read expenses ──────────────────────────────────
  var expRows = sheetToObjects(CONFIG.SHEET_EXPENSES);
  if (expRows.length === 0) {
    ui.alert('גיליון Expenses ריק או חסר.');
    return;
  }

  var expHeaders = Object.keys(expRows[0]);
  var dateCol = findColumn(expHeaders, ['תאריך', 'date']);
  var skuCol = findColumn(expHeaders, ['מקט']);
  if (!dateCol || !skuCol) {
    ui.alert('לא נמצאה עמודת מקט או תאריך בגיליון Expenses.');
    return;
  }

  // Find actual price columns (mirrors Python _find_price_columns)
  var priceBeforeVAT = findColumn(expHeaders, ['מחירלפנימעמ']);
  var priceAfterVAT = findColumn(expHeaders, ['מחירלאחרמעמ']);
  var actualPriceCols = [];
  if (priceBeforeVAT) actualPriceCols.push(priceBeforeVAT);
  if (priceAfterVAT) actualPriceCols.push(priceAfterVAT);
  if (actualPriceCols.length === 0) {
    for (var h = 0; h < expHeaders.length; h++) {
      if (expHeaders[h].indexOf('מחיר') !== -1) {
        actualPriceCols.push(expHeaders[h]);
      }
    }
  }

  // ── 3. Build comparison rows ─────────────────────────
  var resultRows = [];
  var matchedKeys = {};

  for (var i = 0; i < expRows.length; i++) {
    var row = expRows[i];
    var rawDate = row[dateCol];
    var date = (rawDate instanceof Date) ? rawDate : new Date(rawDate);
    var week = getISOWeek(date);
    var sku = String(row[skuCol]).replace(/\.0$/, '').trim();

    // Collect actual price values
    var numericActuals = [];
    for (var ac = 0; ac < actualPriceCols.length; ac++) {
      numericActuals.push(toNum(row[actualPriceCols[ac]]));
    }

    // Base row data
    var baseRow = [
      week,
      row['#'] !== undefined ? row['#'] : '',
      row['כרטיס'] !== undefined ? row['כרטיס'] : '',
      row['לקוח'] !== undefined ? row['לקוח'] : '',
      row['תעודה'] !== undefined ? row['תעודה'] : '',
      date,
      row[skuCol],
      row['פריט'] !== undefined ? row['פריט'] : '',
      row['כמות'] !== undefined ? row['כמות'] : '',
      actualPriceCols.length > 0 ? row[actualPriceCols[0]] : '',
      actualPriceCols.length > 1 ? row[actualPriceCols[1]] : '',
      row['הנחה'] !== undefined ? row['הנחה'] : '',
      row['סכ"ה שורה'] !== undefined ? row['סכ"ה שורה'] : ''
    ];

    var lookupKey = week + '|' + sku;
    matchedKeys[lookupKey] = true;

    if (!dealerMap.hasOwnProperty(lookupKey)) {
      // Python: pd.isna(expected) → 🟡 חסר במחירון
      resultRows.push(baseRow.concat(['', '🟡 חסר במחירון', '']));
      continue;
    }

    var expected = dealerMap[lookupKey];

    if (isNaN(expected)) {
      resultRows.push(baseRow.concat([expected, '🟡 חסר במחירון', '']));
      continue;
    }

    // Python logic: check if ANY actual price column matches exactly
    var matched = false;
    var matchDiff = 0;
    for (var na = 0; na < numericActuals.length; na++) {
      if (!isNaN(numericActuals[na]) && Math.abs(numericActuals[na] - expected) < CONFIG.EXACT_TOLERANCE) {
        matched = true;
        matchDiff = numericActuals[na] - expected;
        break;
      }
    }

    if (matched) {
      resultRows.push(baseRow.concat([expected, '✅ תואם', matchDiff]));
    } else {
      // Find minimum absolute difference
      var minDiff = null;
      for (var na2 = 0; na2 < numericActuals.length; na2++) {
        if (!isNaN(numericActuals[na2])) {
          var d = numericActuals[na2] - expected;
          if (minDiff === null || Math.abs(d) < Math.abs(minDiff)) {
            minDiff = d;
          }
        }
      }
      resultRows.push(baseRow.concat([
        expected,
        '❌ לא תואם',
        minDiff !== null ? minDiff : ''
      ]));
    }
  }

  // ── 4. Flag pricing entries with no matching expense ──
  var dealerKeys = Object.keys(dealerMap);
  for (var m = 0; m < dealerKeys.length; m++) {
    if (!matchedKeys[dealerKeys[m]]) {
      var parts = dealerKeys[m].split('|');
      resultRows.push([
        parts[0], '', '', '', '', '', parts[1], '', '', '', '', '', '',
        dealerMap[dealerKeys[m]],
        '🔵 אין רכישה',
        ''
      ]);
    }
  }

  // ── 5. Sort: latest week first, then by row # ────────
  resultRows.sort(function(a, b) {
    if (a[0] > b[0]) return -1;
    if (a[0] < b[0]) return 1;
    var na = toNum(a[1]);
    var nb = toNum(b[1]);
    if (!isNaN(na) && !isNaN(nb)) return na - nb;
    return 0;
  });

  // ── 6. Write results ─────────────────────────────────
  var compSheet = ss.getSheetByName(CONFIG.SHEET_COMPARISON);
  if (!compSheet) {
    compSheet = ss.insertSheet(CONFIG.SHEET_COMPARISON);
  }
  compSheet.clearContents();
  compSheet.clearFormats();

  var headers = [
    'שבוע', '#', 'כרטיס', 'לקוח', 'תעודה', 'תאריך', 'מקט', 'פריט', 'כמות',
    'מחיר לפני מע"מ', 'מחיר לאחר מע"מ', 'הנחה', 'סכ"ה שורה',
    'מחיר מחירון', 'סטאטוס', 'שוני במחיר'
  ];
  var numCols = headers.length;

  compSheet.getRange(1, 1, 1, numCols).setValues([headers]);

  if (resultRows.length > 0) {
    compSheet.getRange(2, 1, resultRows.length, numCols).setValues(resultRows);
  }

  var totalRows = resultRows.length + 1;

  // ── 7. Formatting (mirrors Python create_downloadable_excel) ──
  var headerRange = compSheet.getRange(1, 1, 1, numCols);
  headerRange.setFontWeight('bold')
    .setBackground('#F1F5F9')
    .setBorder(true, true, true, true, true, true)
    .setHorizontalAlignment('center');
  compSheet.setFrozenRows(1);

  // Set RTL for the sheet
  compSheet.setRightToLeft(true);

  for (var col = 1; col <= numCols; col++) {
    compSheet.autoResizeColumn(col);
  }

  // ── 8. Conditional coloring (batch) ───────────────────
  if (resultRows.length > 0) {
    var backgrounds = [];
    for (var r = 0; r < resultRows.length; r++) {
      var statusVal = String(resultRows[r][14]);
      var color;
      if (statusVal.indexOf('✅') !== -1) {
        color = '#DCFCE7';
      } else if (statusVal.indexOf('❌') !== -1) {
        color = '#FEE2E2';
      } else if (statusVal.indexOf('🟡') !== -1) {
        color = '#FEF9C3';
      } else {
        color = '#DBEAFE';
      }
      var rowColors = [];
      for (var c = 0; c < numCols; c++) rowColors.push(color);
      backgrounds.push(rowColors);
    }
    compSheet.getRange(2, 1, resultRows.length, numCols).setBackgrounds(backgrounds);
  }

  // ── 9. Weekly summary ─────────────────────────────────
  var summaryStart = totalRows + 2;
  compSheet.getRange(summaryStart, 1).setValue('סיכום שבועי').setFontWeight('bold').setFontSize(12);

  var weekStats = {};
  for (var s = 0; s < resultRows.length; s++) {
    var w = resultRows[s][0];
    var st = String(resultRows[s][14]);
    if (!weekStats[w]) {
      weekStats[w] = { match: 0, mismatch: 0, missing: 0, noPurchase: 0, total: 0 };
    }
    weekStats[w].total++;
    if (st.indexOf('✅') !== -1) weekStats[w].match++;
    else if (st.indexOf('❌') !== -1) weekStats[w].mismatch++;
    else if (st.indexOf('🟡') !== -1) weekStats[w].missing++;
    else weekStats[w].noPurchase++;
  }

  var summaryHeaders = ['שבוע', 'סה"כ', '✅ תואם', '❌ לא תואם', '🟡 חסר במחירון', '🔵 אין רכישה'];
  var summaryRow = summaryStart + 1;
  compSheet.getRange(summaryRow, 1, 1, 6).setValues([summaryHeaders])
    .setFontWeight('bold')
    .setBackground('#E2E8F0')
    .setBorder(true, true, true, true, true, true);

  var weekList = Object.keys(weekStats).sort().reverse();
  var summaryData = [];
  for (var wi = 0; wi < weekList.length; wi++) {
    var ws = weekStats[weekList[wi]];
    summaryData.push([weekList[wi], ws.total, ws.match, ws.mismatch, ws.missing, ws.noPurchase]);
  }
  if (summaryData.length > 0) {
    compSheet.getRange(summaryRow + 1, 1, summaryData.length, 6).setValues(summaryData);
  }

  // Grand total
  var grandTotal = { match: 0, mismatch: 0, missing: 0, noPurchase: 0, total: 0 };
  for (var gt = 0; gt < weekList.length; gt++) {
    var wgt = weekStats[weekList[gt]];
    grandTotal.total += wgt.total;
    grandTotal.match += wgt.match;
    grandTotal.mismatch += wgt.mismatch;
    grandTotal.missing += wgt.missing;
    grandTotal.noPurchase += wgt.noPurchase;
  }
  var grandRow = summaryRow + 1 + summaryData.length;
  compSheet.getRange(grandRow, 1, 1, 6)
    .setValues([['סה"כ', grandTotal.total, grandTotal.match, grandTotal.mismatch, grandTotal.missing, grandTotal.noPurchase]])
    .setFontWeight('bold')
    .setBorder(true, true, true, true, true, true);

  ui.alert(
    'השוואה הושלמה!\n\n' +
    resultRows.length + ' שורות הושוו.\n' +
    grandTotal.match + ' תואם ✅\n' +
    grandTotal.mismatch + ' לא תואם ❌\n' +
    grandTotal.missing + ' חסר במחירון 🟡\n' +
    grandTotal.noPurchase + ' אין רכישה 🔵'
  );
}
