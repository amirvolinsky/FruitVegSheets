/** @OnlyCurrentDoc */

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
 *   5. Comparison joins on (תקופה, מקט); תקופה יכולה להיות יום / שבוע (ראשון) / מפתח ישן
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
    .createMenu('🥕 השוואת מחירים')
    .addItem('🚀 התחל השוואה', 'showWizard')
    .addItem('📘 מדריך שימוש', 'showGuide')
    .addSeparator()
    .addItem('📋 צור גיליון בעיות מהגיליון הפעיל', 'createIssuesFromActiveComparison')
    .addToUi();
}

/**
 * Single wizard dialog — pricing file + date + expenses file + compare.
 */
function showWizard() {
  var html = HtmlService.createHtmlOutputFromFile('Wizard')
    .setWidth(560)
    .setHeight(760);
  SpreadsheetApp.getUi().showModalDialog(html, '🥕 השוואת מחירים');
}

/**
 * Wizard entry point (single pricing file) kept for backward compatibility.
 */
function runFullComparison(isoDateStr, pricingRows, expenseRows) {
  return runFullComparisonBatch(
    [{ isoDateStr: isoDateStr, pricingRows: pricingRows }],
    expenseRows
  );
}

/**
 * Wizard entry point (multiple pricing files):
 * pricingEntries: [{ isoDateStr: "YYYY-MM-DD", pricingRows: [...] }, ...]
 * expenseRows: parsed expense sheet rows.
 */
function runFullComparisonBatch(pricingEntries, expenseRows) {
  if (!pricingEntries || pricingEntries.length === 0) {
    return { success: false, message: 'נא להוסיף לפחות מחירון אחד.' };
  }

  var importedWeeks = [];
  var i;
  var ki;
  for (i = 0; i < pricingEntries.length; i++) {
    var entry = pricingEntries[i];
    var keys = resolvePricingPeriodKeys(entry);
    if (keys.length === 0) {
      return {
        success: false,
        message:
          'מחירון ' + (i + 1) + ': לא ניתן לחשב טווח (בדקו תאריכים / טווח / חודש).'
      };
    }
    for (ki = 0; ki < keys.length; ki++) {
      var pr = importPricingData(keys[ki], entry.pricingRows);
      if (!pr.success) {
        return {
          success: false,
          message: 'מחירון ' + (i + 1) + ': ' + pr.message
        };
      }
      importedWeeks.push(keys[ki]);
    }
  }

  expenseRows = normalizeSheetRows(expenseRows);
  var er = importExpensesData(expenseRows);
  if (!er.success) {
    return er;
  }

  var cmp = executeComparison();
  if (!cmp.success) return cmp;

  // Add upload summary for non-technical users.
  var uniqueWeeks = {};
  for (i = 0; i < importedWeeks.length; i++) uniqueWeeks[importedWeeks[i]] = true;
  var weeksList = Object.keys(uniqueWeeks).sort().map(getWeekRangeLabelFromIsoWeekFull).join(', ');
  return {
    success: true,
    message:
      '✅ נטענו ' + pricingEntries.length + ' מחירונים לתאריכים: ' + weeksList + '.\n' +
      er.message + '\n\n' + cmp.message
  };
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

  rows = normalizeSheetRows(rows);

  if (rows.length < 2) {
    return {
      success: false,
      message: 'קובץ ההוצאות חייב לכלול שורת כותרות ולפחות שורת נתונים אחת.'
    };
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

  // Detect header row — skip it if first row looks like headers
  var startIdx = 0;
  if (rows[0] && rows[0].length > 0) {
    var head = sanitizeColName(String(rows[0][0]));
    if (head.indexOf('מקט') !== -1 || head === 'מקט') {
      startIdx = 1;
    }
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

  if (output.length === 0) {
    return { success: false, message: 'לא נמצאו שורות תקינות במחירון (חסר מקט).' };
  }

  sheet.getRange(sheet.getLastRow() + 1, 1, output.length, 4).setValues(output);

  return {
    success: true,
    message: 'נוספו ' + output.length + ' פריטים לשבוע ' + week +
             (skipped > 0 ? ' (' + skipped + ' שורות ללא מקט דולגו)' : '')
  };
}

// ──────────────────────────────────────────────
// HELPER — week key (Sunday-Saturday)
// ──────────────────────────────────────────────

/**
 * Returns Sunday of the week for a given date (local calendar).
 */
function getWeekStartSunday(date) {
  if (!(date instanceof Date) || isNaN(date.getTime())) return null;
  var d = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  var day = d.getDay(); // Sunday = 0
  d.setDate(d.getDate() - day);
  return d;
}

/**
 * Current business week key format: YYYY-MM-DD (Sunday start).
 * Kept under the same function name to avoid breaking callers.
 */
function getISOWeek(date) {
  var sunday = getWeekStartSunday(date);
  if (!sunday) return null;
  var y = sunday.getFullYear();
  var m = String(sunday.getMonth() + 1).padStart(2, '0');
  var d = String(sunday.getDate()).padStart(2, '0');
  return y + '-' + m + '-' + d;
}

/**
 * Legacy ISO week key (YYYY-WW), used for backward compatibility with old Pricing rows.
 */
function getLegacyIsoWeekKey(date) {
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

/**
 * Formats Date to Israeli short format: dd/MM/yy
 */
function formatDateIsraeliShort(dateObj) {
  // Use local calendar fields (not UTC) to avoid day-shift issues.
  var dd = String(dateObj.getDate()).padStart(2, '0');
  var mm = String(dateObj.getMonth() + 1).padStart(2, '0');
  var yyyy = String(dateObj.getFullYear());
  return dd + '/' + mm + '/' + yyyy;
}

/**
 * Converts week key to readable range label.
 * Supported formats:
 * - New format: YYYY-MM-DD (Sunday start)
 * - Legacy ISO: YYYY-WW
 */
function getWeekRangeLabelFromIsoWeek(weekKey) {
  if (!weekKey) return '';
  var key = String(weekKey).trim();

  // New Sunday-based key: YYYY-MM-DD
  if (/^\d{4}-\d{2}-\d{2}$/.test(key)) {
    var p = key.split('-');
    var sy = parseInt(p[0], 10);
    var sm = parseInt(p[1], 10) - 1;
    var sd = parseInt(p[2], 10);
    var start = new Date(sy, sm, sd);
    if (isNaN(start.getTime())) return key;
    var end = new Date(start.getFullYear(), start.getMonth(), start.getDate());
    end.setDate(end.getDate() + 6);
    return formatDateIsraeliShort(start) + ' - ' + formatDateIsraeliShort(end);
  }

  // Legacy ISO key: YYYY-WW
  var parts = key.split('-');
  if (parts.length !== 2) return key;
  var year = parseInt(parts[0], 10);
  var week = parseInt(parts[1], 10);
  if (isNaN(year) || isNaN(week) || week < 1 || week > 53) return key;

  var jan4 = new Date(Date.UTC(year, 0, 4));
  var jan4Day = jan4.getUTCDay() || 7;
  var week1Monday = new Date(jan4);
  week1Monday.setUTCDate(jan4.getUTCDate() - (jan4Day - 1));

  var monday = new Date(week1Monday);
  monday.setUTCDate(week1Monday.getUTCDate() + (week - 1) * 7);
  var sunday = new Date(monday);
  sunday.setUTCDate(monday.getUTCDate() + 6);

  return formatDateIsraeliShort(monday) + ' - ' + formatDateIsraeliShort(sunday);
}

/**
 * Full readable week range label from ISO week key.
 * Example: 2025-22 -> 26/05/2025 - 01/06/2025
 */
function getWeekRangeLabelFromIsoWeekFull(weekKey) {
  return getWeekRangeLabelFromIsoWeek(weekKey);
}

/**
 * סיכום שבועי: המפתחות ב-weekStats הם לעיתים כבר טווח תאריכים "dd/mm/yyyy - dd/mm/yyyy"
 * (עמודת שבוע בשורות התוצאה). לא להעביר דרך getWeekRangeLabelFromIsoWeek — זה יפצל לפי "-" וייצור שנות 1900.
 */
function formatWeekLabelForSummary_(weekKey) {
  var s = String(weekKey || '').trim();
  if (s.indexOf('/') !== -1 && s.indexOf(' - ') !== -1) return s;
  return getWeekRangeLabelFromIsoWeek(s);
}

/** סכום פער כספי משוער לשורה ❌: שוני ליחידה × כמות (אם אין כמות — רק השוני). */
function linePriceGapShekels_(row) {
  var diff = toNum(row[15]);
  if (isNaN(diff)) return 0;
  var qty = toNum(row[8]);
  if (!isNaN(qty) && qty !== 0) return diff * qty;
  return diff;
}

/** רוחב עמודת «שבוע» בגיליון השוואה — לא לפי המקרא הממוזג (שהורס autoResize לעמודה A). */
function estimateComparisonColAPixels_(maxTextLen) {
  var n = Math.max(8, Number(maxTextLen) || 12);
  var px = Math.round(n * 6.9 + 40);
  if (px < 118) px = 118;
  if (px > 440) px = 440;
  return px;
}

function maxLenWeekColumnFromSummaryAndRows_(summaryData, resultRows) {
  var m = 12;
  var i;
  for (i = 0; summaryData && i < summaryData.length; i++) {
    var a = String(summaryData[i][0] || '').length;
    if (a > m) m = a;
  }
  for (i = 0; resultRows && i < resultRows.length; i++) {
    var b = String(resultRows[i][0] || '').length;
    if (b > m) m = b;
  }
  return m;
}

/**
 * עמודת עזר לפילטר: רק לשורות ❌ — חיובי / שלילי בלי תלות במספר המדויק.
 */
function diffSignFilterLabel_(statusStr, diffVal) {
  var st = String(statusStr || '');
  if (st.indexOf('❌') === -1) return '';
  var d = toNum(diffVal);
  if (isNaN(d)) return '';
  if (d > 0) return 'מעל המחירון (+)';
  if (d < 0) return 'מתחת למחירון (−)';
  return 'פער 0';
}

/**
 * ISO calendar date string "YYYY-MM-DD" → ISO week string. Used by Wizard.html.
 */
function getISOWeekFromIsoDateString(isoDateStr) {
  if (!isoDateStr || typeof isoDateStr !== 'string') return null;
  var parts = isoDateStr.split('-');
  if (parts.length !== 3) return null;
  var y = parseInt(parts[0], 10);
  var m = parseInt(parts[1], 10) - 1;
  var d = parseInt(parts[2], 10);
  if (isNaN(y) || isNaN(m) || isNaN(d)) return null;
  var date = new Date(y, m, d);
  if (isNaN(date.getTime())) return null;
  return getISOWeek(date);
}

/**
 * Returns full week range label from ISO date string.
 * Example: 2025-06-01 -> 26/05/2025 - 01/06/2025
 */
function getWeekRangeLabelFromIsoDateStringFull(isoDateStr) {
  var wk = getISOWeekFromIsoDateString(isoDateStr);
  if (!wk) return '';
  return getWeekRangeLabelFromIsoWeekFull(wk);
}

function parseIsoLocal(isoStr) {
  if (!isoStr || typeof isoStr !== 'string') return null;
  var parts = isoStr.split('-');
  if (parts.length !== 3) return null;
  var y = parseInt(parts[0], 10);
  var m = parseInt(parts[1], 10) - 1;
  var d = parseInt(parts[2], 10);
  if (isNaN(y) || isNaN(m) || isNaN(d)) return null;
  var dt = new Date(y, m, d);
  if (isNaN(dt.getTime())) return null;
  return dt;
}

function formatIsoDateLocal(d) {
  if (!(d instanceof Date) || isNaN(d.getTime())) return '';
  var y = d.getFullYear();
  var mo = String(d.getMonth() + 1).padStart(2, '0');
  var day = String(d.getDate()).padStart(2, '0');
  return y + '-' + mo + '-' + day;
}

function getCalendarDayKey(date) {
  if (!(date instanceof Date) || isNaN(date.getTime())) return null;
  return formatIsoDateLocal(date);
}

function getMonthBoundsIsoStrings(isoMonthStr) {
  var p = String(isoMonthStr || '').split('-');
  if (p.length < 2) return null;
  var y = parseInt(p[0], 10);
  var mo = parseInt(p[1], 10);
  if (isNaN(y) || isNaN(mo) || mo < 1 || mo > 12) return null;
  var first = new Date(y, mo - 1, 1);
  var last = new Date(y, mo, 0);
  return [formatIsoDateLocal(first), formatIsoDateLocal(last)];
}

function getSundayWeekKeysOverlappingInclusive(isoStart, isoEnd) {
  var start = parseIsoLocal(isoStart);
  var end = parseIsoLocal(isoEnd);
  if (!start || !end || start > end) return [];

  var sun = getWeekStartSunday(start);
  var keys = [];
  var seen = {};

  while (sun <= end) {
    var weekEnd = new Date(sun.getFullYear(), sun.getMonth(), sun.getDate());
    weekEnd.setDate(weekEnd.getDate() + 6);
    if (weekEnd >= start && sun <= end) {
      var k = getISOWeek(sun);
      if (!seen[k]) {
        seen[k] = true;
        keys.push(k);
      }
    }
    sun.setDate(sun.getDate() + 7);
  }
  return keys;
}

/**
 * Wizard entry: { scope, isoDateStr, isoMonthStr, isoRangeStart, isoRangeEnd, pricingRows }
 * scope: 'week' | 'month' | 'day' | 'range'
 */
function resolvePricingPeriodKeys(entry) {
  if (!entry) return [];
  var scope = entry.scope || 'week';

  if (scope === 'day') {
    var dDay = parseIsoLocal(entry.isoDateStr);
    if (!dDay) return [];
    var dk = getCalendarDayKey(dDay);
    return dk ? [dk] : [];
  }

  if (scope === 'week') {
    var w = getISOWeekFromIsoDateString(entry.isoDateStr);
    return w ? [w] : [];
  }

  if (scope === 'month') {
    var bounds = getMonthBoundsIsoStrings(entry.isoMonthStr);
    if (!bounds) return [];
    return getSundayWeekKeysOverlappingInclusive(bounds[0], bounds[1]);
  }

  if (scope === 'range') {
    return getSundayWeekKeysOverlappingInclusive(entry.isoRangeStart, entry.isoRangeEnd);
  }

  return [];
}

function formatPricingPeriodDisplay(storedKey) {
  if (!storedKey) return '';
  var s = String(storedKey).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
    var d = parseIsoLocal(s);
    if (!d) return s;
    var sun = getWeekStartSunday(d);
    var sunKey = formatIsoDateLocal(sun);
    if (s === sunKey) return getWeekRangeLabelFromIsoWeek(s);
    return formatDateIsraeliShort(d) + ' (מחיר יומי)';
  }
  return getWeekRangeLabelFromIsoWeek(s);
}

/**
 * מנרמל את עמודת "שבוע" בגיליון Pricing למפתח זהה לזה שמשמש בהוצאות:
 * ראשון–שבת כ־YYYY-MM-DD (או מפתח legacy YYYY-WW).
 * תא שגוגל שיטס שומר כ־Date נקרא כ־Date — ללא נרמול יוצאים מפתחות כמו
 * "Sun Jun 01 2025..." שלא תואמים ל־2025-06-01|מקט.
 */
function normalizePricingWeekCell(val) {
  if (val === null || val === undefined || val === '') return '';

  if (val instanceof Date && !isNaN(val.getTime())) {
    return getISOWeek(val);
  }

  // מספר סידורי של תאריך בגיליון (נדיר אחרי getValues)
  if (typeof val === 'number' && val > 40000 && val < 80000) {
    var utc = new Date(Date.UTC(1899, 11, 30) + Math.round(val * 86400000));
    if (!isNaN(utc.getTime())) {
      var loc = new Date(utc.getUTCFullYear(), utc.getUTCMonth(), utc.getUTCDate());
      return getISOWeek(loc);
    }
  }

  var s = String(val).trim();
  if (!s) return '';

  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
    var p = parseIsoLocal(s);
    return p ? getISOWeek(p) : s;
  }

  var parts = s.split('-');
  if (parts.length === 2 && /^\d{4}$/.test(parts[0]) && /^\d{1,2}$/.test(parts[1])) {
    var yLeg = parseInt(parts[0], 10);
    var wLeg = parseInt(parts[1], 10);
    if (!isNaN(yLeg) && !isNaN(wLeg) && wLeg >= 1 && wLeg <= 53) {
      return yLeg + '-' + (wLeg < 10 ? '0' + wLeg : String(wLeg));
    }
  }

  if (/GMT|Israel Daylight|^\w{3},\s+\w{3}\s+\d/.test(s)) {
    var fallback = new Date(s);
    if (!isNaN(fallback.getTime())) return getISOWeek(fallback);
  }

  return s;
}

function escapeHtml(text) {
  if (text === null || text === undefined) return '';
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/**
 * תצוגת תקציר לוויזארד (HTML) — רשימת שבועות מוצגת ב־LTR כדי שלא יבלבלו סדר התאריכים ב־RTL.
 */
function describePricingImport(opts) {
  opts = opts || {};
  var entry = {
    scope: opts.scope || 'week',
    isoDateStr: opts.isoDateStr || '',
    isoMonthStr: opts.isoMonthStr || '',
    isoRangeStart: opts.isoRangeStart || '',
    isoRangeEnd: opts.isoRangeEnd || ''
  };
  var keys = resolvePricingPeriodKeys(entry);
  if (keys.length === 0) {
    return '<div class="pricing-badge-wrap" dir="rtl"><p class="pricing-badge-title">' +
      escapeHtml('לא ניתן לחשב טווח — בדוק תאריכים.') + '</p></div>';
  }

  var sc = entry.scope || 'week';
  var i;

  if (sc === 'day') {
    return (
      '<div class="pricing-badge-wrap" dir="rtl">' +
      '<p class="pricing-badge-title">' +
      escapeHtml('מחיר יומי לתאריך: ') +
      '<span dir="ltr" class="pricing-badge-ltr">' +
      escapeHtml(formatPricingPeriodDisplay(keys[0])) +
      '</span></p></div>'
    );
  }

  if (keys.length === 1) {
    return (
      '<div class="pricing-badge-wrap" dir="rtl">' +
      '<p class="pricing-badge-title">' +
      escapeHtml('משויך לשבוע אחד: ') +
      '<span dir="ltr" class="pricing-badge-ltr">' +
      escapeHtml(getWeekRangeLabelFromIsoWeekFull(keys[0])) +
      '</span></p></div>'
    );
  }

  var title =
    'הקובץ יישמר ' +
    keys.length +
    ' פעמים בגיליון המחירונים — פעם אחת לכל שבוע קלנדרי. לכל מק״ט יופיע אותו מחיר בכל השבועות האלה:';

  var html =
    '<div class="pricing-badge-wrap" dir="rtl">' +
    '<p class="pricing-badge-title">' +
    escapeHtml(title) +
    '</p>' +
    '<ol class="pricing-week-list" dir="ltr">';

  for (i = 0; i < keys.length; i++) {
    html +=
      '<li><span class="pricing-badge-ltr">' +
      escapeHtml(getWeekRangeLabelFromIsoWeekFull(keys[i])) +
      '</span></li>';
  }

  html += '</ol></div>';
  return html;
}

/**
 * Pad jagged 2D arrays so setValues does not fail.
 */
function normalizeSheetRows(rows) {
  if (!rows || rows.length === 0) return rows;
  var maxCols = 0;
  var i;
  for (i = 0; i < rows.length; i++) {
    if (rows[i] && rows[i].length > maxCols) maxCols = rows[i].length;
  }
  if (maxCols === 0) return rows;
  for (i = 0; i < rows.length; i++) {
    if (!rows[i]) rows[i] = [];
    while (rows[i].length < maxCols) rows[i].push('');
  }
  return rows;
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
  if (typeof val === 'string') {
    var t = val.trim();
    if (
      t === '#NUM!' ||
      t === '#VALUE!' ||
      t === '#REF!' ||
      t === '#N/A' ||
      t === '#DIV/0!'
    ) {
      return NaN;
    }
  }
  return Number(val);
}

/**
 * Build a readable, unique comparison sheet name from date range.
 */
function buildComparisonSheetName(startDate, endDate) {
  var tz = Session.getScriptTimeZone() || 'Asia/Jerusalem';
  var baseName = CONFIG.SHEET_COMPARISON;

  if (startDate instanceof Date && !isNaN(startDate.getTime()) &&
      endDate instanceof Date && !isNaN(endDate.getTime())) {
    var fromStr = Utilities.formatDate(startDate, tz, 'dd.MM.yy');
    var toStr = Utilities.formatDate(endDate, tz, 'dd.MM.yy');
    baseName = 'Comparison ' + fromStr + '-' + toStr;
  }

  // Keep only safe sheet-name characters and length.
  baseName = baseName.replace(/[\\/?*[\]:]/g, '-').trim();
  if (!baseName) baseName = CONFIG.SHEET_COMPARISON;
  if (baseName.length > 90) baseName = baseName.substring(0, 90);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var name = baseName;
  var suffix = 2;
  while (ss.getSheetByName(name)) {
    var candidate = baseName + ' (' + suffix + ')';
    if (candidate.length > 99) {
      candidate = baseName.substring(0, Math.max(1, 99 - (' (' + suffix + ')').length)) + ' (' + suffix + ')';
    }
    name = candidate;
    suffix++;
  }
  return name;
}

/**
 * שם גיליון ייחודי (בטוח לתווים ולאורך).
 */
function ensureUniqueSheetName_(baseName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var name = String(baseName || 'גיליון')
    .replace(/[\\/?*[\]:]/g, '-')
    .trim();
  if (name.length > 99) name = name.substring(0, 99);
  if (!name) name = 'גיליון';
  var candidate = name;
  var suffix = 2;
  while (ss.getSheetByName(candidate)) {
    var suf = ' (' + suffix + ')';
    candidate = name.substring(0, Math.max(1, 99 - suf.length)) + suf;
    suffix++;
  }
  return candidate;
}

function headerColumnIndex_(headerRow, substr) {
  var s = String(substr);
  for (var i = 0; i < headerRow.length; i++) {
    if (String(headerRow[i]).indexOf(s) !== -1) return i;
  }
  return -1;
}

/** עמודת סטטוס — גם כותרת ישנה «סטאטוס» בגיליונות שנוצרו לפני התיקון */
function statusColumnIndex_(headerRow) {
  var ix = headerColumnIndex_(headerRow, 'סטטוס');
  if (ix >= 0) return ix;
  return headerColumnIndex_(headerRow, 'סטאטוס');
}

/**
 * מוצא את שורת כותרות טבלת ההשוואה הגדולה (שבוע | # | כרטיס…), לא את טבלת הסיכום העליונה.
 */
function findMainComparisonHeaderRowIndex_(data) {
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (String(row[0]).trim() !== 'שבוע') continue;
    var second = String(row[1] != null ? row[1] : '').trim();
    if (second === '#' || second === '\u0023') return i;
  }
  return -1;
}

/**
 * תפריט: בונה גיליון בעיות מגיליון השוואה פתוח (נוצר רק בלחיצה כאן).
 */
function createIssuesFromActiveComparison() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getActiveSheet();
  var data = sh.getDataRange().getValues();
  if (data.length < 2) {
    SpreadsheetApp.getUi().alert('אין מספיק שורות בגיליון.');
    return;
  }

  var hdrIx = findMainComparisonHeaderRowIndex_(data);
  if (hdrIx < 0) {
    SpreadsheetApp.getUi().alert('לא נמצאה שורת כותרות של טבלת ההשוואה (שבוע / #). בחרו גיליון Comparison.');
    return;
  }

  var headerRow = data[hdrIx];
  var ixStatus = statusColumnIndex_(headerRow);
  if (ixStatus < 0) {
    SpreadsheetApp.getUi().alert('לא נמצאה עמודת "סטטוס".');
    return;
  }

  var ixWeek = headerColumnIndex_(headerRow, 'שבוע');
  var ixSku = headerColumnIndex_(headerRow, 'מקט');
  var ixItem = headerColumnIndex_(headerRow, 'פריט');
  var ixDoc = headerColumnIndex_(headerRow, 'תעודה');
  var ixDate = headerColumnIndex_(headerRow, 'תאריך');
  var ixList = headerColumnIndex_(headerRow, 'מחיר מחירון');
  var ixDiff = headerColumnIndex_(headerRow, 'שוני');
  var ixSign = headerColumnIndex_(headerRow, 'מעל / מתחת');
  if (ixSign < 0) ixSign = headerColumnIndex_(headerRow, 'מעל');

  function cell(line, idx) {
    if (idx < 0) return '';
    return line[idx];
  }

  var issueRows = [];
  var r;
  for (r = hdrIx + 1; r < data.length; r++) {
    var line = data[r];
    var c0 = String(line[0] || '');
    if (c0.indexOf('סיכום שבועי') !== -1) break;

    var stCell = String(cell(line, ixStatus) || '');
    if (stCell.indexOf('✅') !== -1) continue;
    if (!stCell.trim()) continue;

    issueRows.push([
      cell(line, ixWeek),
      cell(line, ixSku),
      cell(line, ixItem),
      cell(line, ixDoc),
      cell(line, ixDate),
      stCell,
      cell(line, ixList),
      cell(line, ixDiff),
      ixSign >= 0 ? cell(line, ixSign) : ''
    ]);
  }

  if (issueRows.length === 0) {
    SpreadsheetApp.getUi().alert('לא נמצאו שורות בעיה (או שהכל תואם).');
    return;
  }

  var shortLabel = sh.getName().length > 72 ? sh.getName().substring(0, 72) + '…' : sh.getName();
  var sheetName = ensureUniqueSheetName_('בעיות — ' + shortLabel);

  var allSheets = ss.getSheets();
  var insertBefore = allSheets.length + 1;
  for (var si = 0; si < allSheets.length; si++) {
    if (allSheets[si].getSheetId() === sh.getSheetId()) {
      insertBefore = si + 2;
      break;
    }
  }

  var issueSheet = ss.insertSheet(sheetName, insertBefore);
  var ih = ['שבוע', 'מקט', 'פריט', 'תעודה', 'תאריך', 'סטטוס', 'מחיר מחירון', 'שוני במחיר', 'מעל / מתחת'];
  issueSheet.getRange(1, 1, 1, ih.length).setValues([ih]).setFontWeight('bold').setBackground('#E2E8F0');
  issueSheet.getRange(2, 1, issueRows.length, ih.length).setValues(issueRows);
  issueSheet.setRightToLeft(true);
  for (var ic = 1; ic <= ih.length; ic++) issueSheet.autoResizeColumn(ic);
  issueSheet.setFrozenRows(1);

  try {
    issueSheet.getRange(1, 1, issueRows.length + 1, ih.length).createFilter();
  } catch (fe) {}

  SpreadsheetApp.getUi().alert('נוצר גיליון: ' + sheetName + '\n(' + issueRows.length + ' שורות)');
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

/**
 * Menu entry / programmatic re-run: compares sheets already loaded (no upload).
 */
function compareWeeklyPrices() {
  var r = executeComparison();
  SpreadsheetApp.getUi().alert(r.message);
}

/**
 * Core comparison: reads Pricing + Expenses, writes Comparison. Returns { success, message }.
 */
function executeComparison() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var minExpenseDate = null;
  var maxExpenseDate = null;

  // ── 1. Read pricing into map: "week|sku" → מחירון ──────
  var pricingRows = sheetToObjects(CONFIG.SHEET_PRICING);
  if (pricingRows.length === 0) {
    return {
      success: false,
      message: 'גיליון Pricing ריק או חסר.\nהשתמשו בתפריט: התחל השוואה.'
    };
  }

  var pricingHeaders = Object.keys(pricingRows[0]);
  var priceCol = findColumn(pricingHeaders, ['מחירון', 'מחיר']);
  if (!priceCol) {
    return { success: false, message: 'לא נמצאה עמודת מחירון בגיליון Pricing.' };
  }

  var dealerMap = {};  // "week|sku" → price
  for (var p = 0; p < pricingRows.length; p++) {
    var pr = pricingRows[p];
    var pWeek = normalizePricingWeekCell(pr['שבוע']);
    var pSku = String(pr['מקט']).replace(/\.0$/, '').trim();
    if (!pWeek || !pSku || pSku === '' || pSku === 'undefined') continue;
    var key = pWeek + '|' + pSku;
    dealerMap[key] = toNum(pr[priceCol]);
  }

  // ── 2. Read expenses ──────────────────────────────────
  var expRows = sheetToObjects(CONFIG.SHEET_EXPENSES);
  if (expRows.length === 0) {
    return { success: false, message: 'גיליון Expenses ריק או חסר.' };
  }

  var expHeaders = Object.keys(expRows[0]);
  var dateCol = findColumn(expHeaders, ['תאריך', 'date']);
  var skuCol = findColumn(expHeaders, ['מקט']);
  if (!dateCol || !skuCol) {
    return { success: false, message: 'לא נמצאה עמודת מקט או תאריך בגיליון Expenses.' };
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
    if (date instanceof Date && !isNaN(date.getTime())) {
      if (!minExpenseDate || date < minExpenseDate) minExpenseDate = new Date(date);
      if (!maxExpenseDate || date > maxExpenseDate) maxExpenseDate = new Date(date);
    }
    var week = getISOWeek(date);
    var legacyWeek = getLegacyIsoWeekKey(date);
    var weekDisplay = getWeekRangeLabelFromIsoWeek(week);
    var sku = String(row[skuCol]).replace(/\.0$/, '').trim();

    // Collect actual price values
    var numericActuals = [];
    for (var ac = 0; ac < actualPriceCols.length; ac++) {
      numericActuals.push(toNum(row[actualPriceCols[ac]]));
    }

    // Base row data
    var baseRow = [
      weekDisplay,
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

    var dayKey = getCalendarDayKey(date);
    var lookupKeyWeek = week + '|' + sku;
    var lookupKeyLegacy = legacyWeek + '|' + sku;
    var lookupKeyDay = dayKey ? dayKey + '|' + sku : null;

    var lookupOrder = [];
    if (lookupKeyDay) lookupOrder.push(lookupKeyDay);
    lookupOrder.push(lookupKeyWeek);
    lookupOrder.push(lookupKeyLegacy);

    var matchedLookupKey = null;
    var li;
    for (li = 0; li < lookupOrder.length; li++) {
      if (dealerMap.hasOwnProperty(lookupOrder[li])) {
        matchedLookupKey = lookupOrder[li];
        break;
      }
    }

    if (matchedLookupKey === null) {
      resultRows.push(baseRow.concat(['', '🟡 חסר במחירון', '', '']));
      continue;
    }

    matchedKeys[matchedLookupKey] = true;
    var expected = dealerMap[matchedLookupKey];

    if (isNaN(expected)) {
      resultRows.push(baseRow.concat([expected, '🟡 חסר במחירון', '', '']));
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
      resultRows.push(baseRow.concat([expected, '✅ תואם', matchDiff, '']));
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
        minDiff !== null ? minDiff : '',
        diffSignFilterLabel_('❌ לא תואם', minDiff)
      ]));
    }
  }

  // ── 4. Flag pricing entries with no matching expense ──
  var dealerKeys = Object.keys(dealerMap);
  for (var m = 0; m < dealerKeys.length; m++) {
    if (!matchedKeys[dealerKeys[m]]) {
      var parts = dealerKeys[m].split('|');
      resultRows.push([
        formatPricingPeriodDisplay(parts[0]), '', '', '', '', '', parts[1], '', '', '', '', '', '',
        dealerMap[dealerKeys[m]],
        '🔵 אין רכישה',
        '',
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

  // ── 6. Weekly stats (לפני כתיבה לגיליון) ───────────────
  var weekStats = {};
  for (var s = 0; s < resultRows.length; s++) {
    var w = resultRows[s][0];
    var st = String(resultRows[s][14]);
    if (!weekStats[w]) {
      weekStats[w] = { match: 0, mismatch: 0, missing: 0, noPurchase: 0, total: 0, gapSum: 0 };
    }
    weekStats[w].total++;
    if (st.indexOf('✅') !== -1) weekStats[w].match++;
    else if (st.indexOf('❌') !== -1) {
      weekStats[w].mismatch++;
      weekStats[w].gapSum += linePriceGapShekels_(resultRows[s]);
    } else if (st.indexOf('🟡') !== -1) weekStats[w].missing++;
    else weekStats[w].noPurchase++;
  }

  var weekList = Object.keys(weekStats).sort().reverse();
  var summaryData = [];
  var wi;
  for (wi = 0; wi < weekList.length; wi++) {
    var ws = weekStats[weekList[wi]];
    summaryData.push([
      formatWeekLabelForSummary_(weekList[wi]),
      ws.total,
      ws.match,
      ws.mismatch,
      ws.missing,
      ws.noPurchase,
      ws.gapSum
    ]);
  }

  var grandTotal = { match: 0, mismatch: 0, missing: 0, noPurchase: 0, total: 0, gapSum: 0 };
  for (var gt = 0; gt < weekList.length; gt++) {
    var wgt = weekStats[weekList[gt]];
    grandTotal.total += wgt.total;
    grandTotal.match += wgt.match;
    grandTotal.mismatch += wgt.mismatch;
    grandTotal.missing += wgt.missing;
    grandTotal.noPurchase += wgt.noPurchase;
    grandTotal.gapSum += wgt.gapSum;
  }

  var maxWeekColLen = maxLenWeekColumnFromSummaryAndRows_(summaryData, resultRows);

  // ── 7. Write results — סיכום למעלה, אחר כך טבלת ההשוואה ──
  var comparisonSheetName = buildComparisonSheetName(minExpenseDate, maxExpenseDate);
  var compSheet = ss.insertSheet(comparisonSheetName);

  var headers = [
    'שבוע', '#', 'כרטיס', 'לקוח', 'תעודה', 'תאריך', 'מקט', 'פריט', 'כמות',
    'מחיר לפני מע"מ', 'מחיר לאחר מע"מ', 'הנחה', 'סכ"ה שורה',
    'מחיר מחירון', 'סטטוס', 'שוני במחיר', 'מעל / מתחת'
  ];
  var numCols = headers.length;

  var COMPARISON_LEGEND_TEXT =
    'מקרא: ✅ תואם — המחיר תואם למחירון | ❌ לא תואם — פער מול המחירון | 🟡 חסר במחירון — אין מחיר לשבוע+מקט בגיליון Pricing | 🔵 אין רכישה — במחירון יש מחיר לשבוע זה ולא נמצאה רכישה';

  var FILTER_HINT_TEXT =
    'סינון מהיר (שורת כותרות מתחת): בעמודת «סטטוס» — ✅ בדיוק במחיר | ❌ לא תואם | 🟡 חסר במחירון | 🔵 אין רכישה. ' +
    'בעמודת «מעל / מתחת» אפשר לסנן במהירות חיובי מול שלילי מול שורות ❌. בעמודת «שוני במחיר» — ערך חיובי / שלילי הוא ההפרש ליחידה.';

  var summaryHeaders = [
    'שבוע',
    'סה"כ',
    '✅ תואם',
    '❌ לא תואם',
    '🟡 חסר במחירון',
    '🔵 אין רכישה',
    'פער במחיר (₪)'
  ];
  var summaryNumCols = summaryHeaders.length;

  var row = 1;

  compSheet.getRange(row, 1, 1, numCols).merge();
  compSheet.getRange(row, 1)
    .setValue(COMPARISON_LEGEND_TEXT)
    .setWrap(true)
    .setFontSize(10)
    .setBackground('#f8fafc')
    .setVerticalAlignment('middle');
  compSheet.setRowHeight(1, 96);
  row++;

  compSheet.getRange(row, 1).setValue('סיכום שבועי').setFontWeight('bold').setFontSize(12);
  row++;

  compSheet.getRange(row, 1, 1, summaryNumCols)
    .setValues([summaryHeaders])
    .setFontWeight('bold')
    .setBackground('#E2E8F0')
    .setBorder(true, true, true, true, true, true);
  row++;

  if (summaryData.length > 0) {
    compSheet.getRange(row, 1, summaryData.length, summaryNumCols).setValues(summaryData);
    compSheet.getRange(row, summaryNumCols, summaryData.length, 1).setNumberFormat('#,##0.00');
    row += summaryData.length;
  }

  compSheet.getRange(row, 1, 1, summaryNumCols)
    .setValues([['סה"כ', grandTotal.total, grandTotal.match, grandTotal.mismatch, grandTotal.missing, grandTotal.noPurchase, grandTotal.gapSum]])
    .setFontWeight('bold')
    .setBorder(true, true, true, true, true, true);
  compSheet.getRange(row, summaryNumCols).setNumberFormat('#,##0.00');
  row++;

  row++;

  var FILTER_HINT_ROW = row;
  compSheet.getRange(row, 1, 1, numCols).merge();
  compSheet.getRange(row, 1)
    .setValue(FILTER_HINT_TEXT)
    .setWrap(true)
    .setFontSize(10)
    .setBackground('#fffbeb')
    .setVerticalAlignment('middle');
  compSheet.setRowHeight(FILTER_HINT_ROW, 72);
  row++;

  var MAIN_HEADER_ROW = row;
  compSheet.getRange(row, 1, 1, numCols).setValues([headers]);
  row++;

  var DATA_FIRST_ROW = row;
  var lastDataRow = DATA_FIRST_ROW - 1;
  if (resultRows.length > 0) {
    lastDataRow = DATA_FIRST_ROW + resultRows.length - 1;
    compSheet.getRange(DATA_FIRST_ROW, 1, resultRows.length, numCols).setValues(resultRows);
    compSheet.getRange('A' + DATA_FIRST_ROW + ':A' + lastDataRow).setNumberFormat('@');
    compSheet.getRange('F' + DATA_FIRST_ROW + ':F' + lastDataRow).setNumberFormat('dd/mm/yy');
  }

  compSheet.getRange(MAIN_HEADER_ROW, 1, 1, numCols)
    .setFontWeight('bold')
    .setBackground('#F1F5F9')
    .setBorder(true, true, true, true, true, true)
    .setHorizontalAlignment('center');

  compSheet.setFrozenRows(MAIN_HEADER_ROW);
  compSheet.setRightToLeft(true);

  // לא קוראים ל-autoResizeColumn(1): תא ממוזג עם המקרא גורם לגוגל שיטס להרחיב את עמודה A לכל רוחב הטקסט.
  compSheet.setColumnWidth(1, estimateComparisonColAPixels_(maxWeekColLen));
  for (var col = 2; col <= numCols; col++) {
    compSheet.autoResizeColumn(col);
  }

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
    compSheet.getRange(DATA_FIRST_ROW, 1, resultRows.length, numCols).setBackgrounds(backgrounds);
  }

  if (resultRows.length > 0) {
    try {
      compSheet.getRange(MAIN_HEADER_ROW, 1, resultRows.length + 1, numCols).createFilter();
    } catch (filterErr) {
      // גיליון חדש — לרוב יעבוד; אם לא, מתעלמים בשקט
    }
  }

  var msg =
    'השוואה הושלמה!\n\n' +
    resultRows.length + ' שורות הושוו.\n' +
    grandTotal.match + ' תואם ✅\n' +
    grandTotal.mismatch + ' לא תואם ❌\n' +
    grandTotal.missing + ' חסר במחירון 🟡\n' +
    grandTotal.noPurchase + ' אין רכישה 🔵\n\n' +
    'התוצאות בגיליון: ' + comparisonSheetName + '.\n' +
    'לגיליון בעיות (ללא שורות תואמות): תפריט ← צור גיליון בעיות מהגיליון הפעיל.';

  return { success: true, message: msg };
}
