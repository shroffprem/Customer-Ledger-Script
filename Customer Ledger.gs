// BridgeLine Ledger - consolidated script
//
// What this does, in one run:
//   1. Removes all pre-June-2026 rows from "Accounts" and "M Coll" (archiving
//      them first into "Accounts_Archive_PreJune" / "MColl_Archive_PreJune").
//      One-time historical cleanup -- not part of the recurring monthly flow.
//   2. Every 10 minutes, rebuilds the "Customer Ledger" tab as a Tally-style
//      accounting ledger: one section per CASE (Disbursement ID), columns
//      Date | Particulars | Vch Type | Vch No. | Debit | Credit | Balance,
//      each showing that case's FULL chronological history with a true
//      running balance ending in a bold double-ruled Closing Balance row.
//      A customer with multiple loans gets one separate section per loan
//      (labeled "Name — Disbursement ID") -- never merged into one
//      combined balance. A case is included if it has any activity this
//      month, OR it's still open (regardless of how many earlier months
//      back it was disbursed) -- closed cases with no current-month
//      activity drop out entirely.
//      Anything disbursed before LEDGER_CUTOFF_DATE (March/April) is
//      excluded permanently, no matter its status. Reference text (Debit
//      Note / Credit Note / M-Coll note) is shown exactly as entered --
//      nothing here re-parses or can silently blank it.
//   3. On the 1st of every month, monthEndRollover() automatically:
//        a. Moves every Accounts row disbursed in the month that just ended
//           into a new dated tab (e.g. "June 26"), in the same column
//           layout as Accounts -- mirroring the existing "Apr/May26"
//           follow-up pattern, but automated and named per month.
//        b. Registers that new tab in the Config sheet's
//           active_archive_tabs list, so the widget (Python side) picks it
//           up automatically with no code changes.
//        c. Rebuilds the live "Customer Ledger" so it reflects the new
//           month immediately (no separate frozen snapshot needed -- the
//           inclusion rule above always reflects the correct picture).
//   4. Installs time-based triggers (every 10 minutes for the ledger, once
//      a month for the rollover) so both stay current automatically. No
//      onChange trigger is used, because onChange fires on the script's own
//      writes too and causes an endless rebuild loop.
//
// Setup (run once):
//   1. Delete any old "Customer Ledger", "Accounts_Archive_PreJune",
//      "MColl_Archive_PreJune" tabs and any old triggers (Apps Script editor,
//      left sidebar clock icon -> delete any existing triggers).
//   2. Keep "Accounts" and "M Coll" - those are fed by your Forms and are the
//      only source data. Do not delete them.
//   3. Open "BridgeLine Accounts" then Extensions then Apps Script.
//   4. Delete any existing script content, paste this whole file in, Save.
//   5. Select freshSetup from the function dropdown at the top, then Run.
//      Approve the permission prompt when asked.
//   6. Reload the spreadsheet tab in your browser once it finishes.
//
// After that, a "Ledger Tools" menu appears in the spreadsheet itself with
// manual options, the ledger refreshes on its own every 10 minutes, and the
// monthly rollover runs itself automatically on the 1st of each month.

var ACCOUNTS_SHEET = "Accounts";
var MCOLL_SHEET = "M Coll";
var LEDGER_SHEET = "Customer Ledger";
var ACC_ARCHIVE_SHEET = "Accounts_Archive_PreJune";
var MC_ARCHIVE_SHEET = "MColl_Archive_PreJune";
var CONFIG_SHEET = "Config";
var CUTOFF = new Date(2026, 5, 1); // June 1, 2026
var LEDGER_CUTOFF_DATE = new Date(2026, 4, 1); // May 1, 2026 -- March/April are
  // permanently done; nothing dated before this ever enters the Customer
  // Ledger again, regardless of open/closed status.

function toDate_(v) {
  if (v instanceof Date) return v;
  if (!v) return null;
  var s = String(v).trim();
  var m = s.match(/^(\d{1,2})-(\d{1,2})-(\d{4})/);
  if (m) return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
  var d = new Date(s);
  if (isNaN(d.getTime())) return null;
  return d;
}

// Some Amount/Collected Amount cells are stored as genuine numbers, others
// as text with Indian thousands separators (e.g. "13,51,000.00") -- plain
// Number() returns NaN on the comma-text ones, silently zeroing them out.
// Strip commas first so both forms parse correctly.
function parseNum_(v) {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return v;
  var n = Number(String(v).replace(/,/g, ''));
  return isNaN(n) ? 0 : n;
}

function getOrCreateArchive_(ss, name, headerRow) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow]);
    sheet.getRange(1, 1, 1, headerRow.length).setFontWeight("bold");
  }
  return sheet;
}

// ── Config sheet helpers (same key/value row format the Python widget uses) ──

function getConfigValue_(key) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cfgSheet = ss.getSheetByName(CONFIG_SHEET);
  if (!cfgSheet) return null;
  var vals = cfgSheet.getDataRange().getValues();
  for (var i = 0; i < vals.length; i++) {
    if (vals[i][0] === key) {
      var raw = vals[i][1];
      try {
        return JSON.parse(raw);
      } catch (e) {
        return raw;
      }
    }
  }
  return null;
}

function setConfigValue_(key, value) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cfgSheet = ss.getSheetByName(CONFIG_SHEET);
  if (!cfgSheet) {
    cfgSheet = ss.insertSheet(CONFIG_SHEET);
    cfgSheet.appendRow(["key", "value"]);
  }
  var serialized = (typeof value === "object") ? JSON.stringify(value) : String(value);
  var vals = cfgSheet.getDataRange().getValues();
  for (var i = 0; i < vals.length; i++) {
    if (vals[i][0] === key) {
      cfgSheet.getRange(i + 1, 2).setValue(serialized);
      return;
    }
  }
  cfgSheet.appendRow([key, serialized]);
}

function getArchiveTabNames_() {
  var tabs = getConfigValue_('active_archive_tabs');
  if (!tabs || !tabs.length) return ['Apr/May26'];
  return tabs;
}

function removePreJuneData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var accSheet = ss.getSheetByName(ACCOUNTS_SHEET);
  var mcSheet = ss.getSheetByName(MCOLL_SHEET);
  if (!accSheet || !mcSheet) {
    throw new Error("Could not find Accounts or M Coll tabs.");
  }

  var accData = accSheet.getDataRange().getValues();
  var keepDisbIds = {};
  var accRemoveRowNums = [];
  var accRemovedRows = [];

  for (var i = 2; i < accData.length; i++) {
    var row = accData[i];
    if (!row[0]) continue;
    var disbDate = toDate_(row[1]);
    if (disbDate && disbDate < CUTOFF) {
      accRemoveRowNums.push(i + 1);
      accRemovedRows.push(row);
    } else {
      keepDisbIds[row[0]] = true;
    }
  }

  var mcData = mcSheet.getDataRange().getValues();
  var mcRemoveRowNums = [];
  var mcRemovedRows = [];

  for (var j = 1; j < mcData.length; j++) {
    var mrow = mcData[j];
    if (!mrow[0]) continue;
    if (!keepDisbIds[mrow[0]]) {
      mcRemoveRowNums.push(j + 1);
      mcRemovedRows.push(mrow);
    }
  }

  if (accRemovedRows.length > 0) {
    var accArchive = getOrCreateArchive_(ss, ACC_ARCHIVE_SHEET, accData[1]);
    accArchive.getRange(accArchive.getLastRow() + 1, 1, accRemovedRows.length, accRemovedRows[0].length)
      .setValues(accRemovedRows);
  }
  if (mcRemovedRows.length > 0) {
    var mcArchive = getOrCreateArchive_(ss, MC_ARCHIVE_SHEET, mcData[0]);
    mcArchive.getRange(mcArchive.getLastRow() + 1, 1, mcRemovedRows.length, mcRemovedRows[0].length)
      .setValues(mcRemovedRows);
  }

  accRemoveRowNums.sort(function (a, b) { return b - a; }).forEach(function (r) { accSheet.deleteRow(r); });
  mcRemoveRowNums.sort(function (a, b) { return b - a; }).forEach(function (r) { mcSheet.deleteRow(r); });

  return { accRemoved: accRemovedRows.length, mcRemoved: mcRemovedRows.length };
}

// ── Ledger data model ─────────────────────────────────────────────────────

// Reads Accounts (live, current month) + every registered archive tab, and
// builds one flat, FULL chronological event list per customer (Disbursement
// /Charges/GST/Collection) -- no Opening Balance abstraction, no period
// splitting. A case is included in full (every event it has, ever) if
// either:
//   (a) it has at least one event (disbursement or collection) dated in the
//       current calendar month, or
//   (b) it is still open (Overdue Status != Closed), regardless of how many
//       months back it was disbursed.
// Closed cases with no current-month activity are dropped entirely. Anything
// disbursed before LEDGER_CUTOFF_DATE (March/April) is excluded no matter
// what -- permanently retired.
function buildLedgerData_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var accSheet = ss.getSheetByName(ACCOUNTS_SHEET);
  var mcSheet = ss.getSheetByName(MCOLL_SHEET);
  if (!accSheet || !mcSheet) {
    throw new Error("Could not find Accounts or M Coll tabs.");
  }

  var now = new Date();
  var monthStart = new Date(now.getFullYear(), now.getMonth(), 1);
  var monthEndExclusive = new Date(now.getFullYear(), now.getMonth() + 1, 1);

  var accData = accSheet.getDataRange().getValues();
  var accRows = accData.slice(2).filter(function (r) { return r[0]; });

  getArchiveTabNames_().forEach(function (tabName) {
    var tab = ss.getSheetByName(tabName);
    if (!tab) return;
    var tabData = tab.getDataRange().getValues();
    // Tabs created by monthEndRollover have the same 2-header-row layout as
    // Accounts; the legacy Apr/May26 tab has no header row at all -- detect
    // which by checking whether row 0 already looks like a data row.
    var startIdx = (tabData[0] && String(tabData[0][0]).indexOf('BLP-') === 0) ? 0 : 2;
    accRows = accRows.concat(tabData.slice(startIdx).filter(function (r) { return r[0]; }));
  });

  var mcData = mcSheet.getDataRange().getValues();
  var mcRows = mcData.slice(1).filter(function (r) { return r[0]; });

  var mcByDisb = {};
  mcRows.forEach(function (r) {
    var disbId = r[0];
    if (!mcByDisb[disbId]) mcByDisb[disbId] = [];
    mcByDisb[disbId].push({ date: r[1], amount: r[2], note: r[3] });
  });

  // One flat list of ledger events (Debit = disbursement/charges/GST,
  // Credit = collection), grouped by CASE (Disbursement ID), not by
  // customer -- the same person taking multiple loans gets one separate
  // section per loan, each with its own running balance, never merged
  // together. Reference text (Debit Note / Credit Note / M-Coll note)
  // flows straight through unparsed -- there is no regex here that can
  // silently blank it.
  var eventsByCase = {};
  var caseLabels = {}; // disbId -> "Customer Name — BLP-XXX" section header text

  accRows.forEach(function (r) {
    var disbId = r[0], disbDate = r[1];
    var disbDateParsed = toDate_(disbDate);
    // March/April are permanently retired -- never enter the ledger again.
    if (disbDateParsed && disbDateParsed < LEDGER_CUTOFF_DATE) return;

    var status = (r[18] || '').toString().trim().toUpperCase();
    var customer = (r[2] || '').toString().trim() || '(No Name)';
    var amount = r[7], charges = r[8], gst = r[9];
    var collDate = r[11], collAmt = r[12];
    var debitNote = r[21], creditNote = r[22];

    // Build this case's full event list first, then decide whether to
    // include it at all.
    var caseEvents = [];
    caseEvents.push({
      date: disbDate,
      bold: "Disbursed to " + customer,
      note: debitNote ? String(debitNote) : "",
      vchType: "Disbursement",
      vchNo: disbId,
      debit: parseNum_(amount),
      credit: 0
    });

    // Charges and GST are owed by the customer too (Total Payable =
    // Amount + Charges + GST) -- shown as their own debit rows rather
    // than folded into the disbursement figure, so each is independently
    // traceable in the running balance.
    if (parseNum_(charges) > 0) {
      caseEvents.push({
        date: disbDate,
        bold: "Processing Charges",
        note: "",
        vchType: "Charges",
        vchNo: disbId,
        debit: parseNum_(charges),
        credit: 0
      });
    }
    if (parseNum_(gst) > 0) {
      caseEvents.push({
        date: disbDate,
        bold: "GST on Charges",
        note: "",
        vchType: "GST",
        vchNo: disbId,
        debit: parseNum_(gst),
        credit: 0
      });
    }

    // M Coll's instalments can understate the truth -- e.g. a final top-up
    // payment recorded directly in Accounts' own Collected Amount column
    // but never logged as a separate M Coll row. Never let the M-Coll sum
    // fall short of what Accounts itself shows; add the gap as one more
    // collection event so the case can actually reach its real balance.
    var colls = [];
    var mcSum = 0;
    if (mcByDisb[disbId] && mcByDisb[disbId].length > 0) {
      mcByDisb[disbId].forEach(function (c) {
        colls.push({ date: c.date, amount: c.amount, note: c.note });
        mcSum += parseNum_(c.amount);
      });
    }
    var accountsCollAmt = parseNum_(collAmt);
    if (colls.length === 0 && accountsCollAmt > 0) {
      colls.push({ date: collDate, amount: accountsCollAmt, note: creditNote });
    } else if (accountsCollAmt > mcSum + 0.5) {
      // Credit Note often accumulates every ref ever recorded for this case
      // (comma-joined), so pull out only the piece(s) NOT already shown via
      // an M Coll row -- otherwise the gap event duplicates a ref that's
      // already visible above it.
      var existingNotes = colls.map(function (c) { return c.note || ''; }).join(' | ');
      var newRefParts = String(creditNote || '').split(',')
        .map(function (p) { return p.trim(); })
        .filter(function (p) { return p && existingNotes.indexOf(p) === -1; });
      colls.push({
        date: collDate,
        amount: accountsCollAmt - mcSum,
        note: newRefParts.join(', ') || '(Additional collection recorded in Accounts, not logged in M Coll)'
      });
    }

    colls.forEach(function (c) {
      caseEvents.push({
        date: c.date,
        bold: "Collection received",
        note: c.note ? String(c.note) : "",
        vchType: "Collection",
        vchNo: disbId,
        debit: 0,
        credit: parseNum_(c.amount)
      });
    });

    // Inclusion rule: this case appears (with its FULL history) if it has
    // any activity this month, OR it's still open -- regardless of which
    // earlier month (after the cutoff) it originated in.
    var hasActivityThisMonth = caseEvents.some(function (ev) {
      var d = toDate_(ev.date);
      return d && d >= monthStart && d < monthEndExclusive;
    });
    var stillOpen = status !== 'CLOSED';
    if (!hasActivityThisMonth && !stillOpen) return;

    eventsByCase[disbId] = caseEvents;
    caseLabels[disbId] = customer + " — " + disbId;
  });

  // Order sections chronologically by each case's earliest transaction date
  // (not alphabetically by name) -- each Disbursement ID is its own
  // section, even when the same customer has multiple loans.
  var caseIds = Object.keys(eventsByCase).sort(function (a, b) {
    var da = eventsByCase[a].reduce(function (min, ev) {
      var d = toDate_(ev.date);
      return (d && (!min || d < min)) ? d : min;
    }, null);
    var db = eventsByCase[b].reduce(function (min, ev) {
      var d = toDate_(ev.date);
      return (d && (!min || d < min)) ? d : min;
    }, null);
    if (!da && !db) return 0;
    if (!da) return 1;
    if (!db) return -1;
    return da - db;
  });

  var headers = ["Date", "Particulars", "Vch Type", "Vch No.", "Debit", "Credit", "Balance"];
  var out = [headers];
  var sectionHeaderRows = [];    // customer-name rows -> bold + merged
  var sectionTotalRows = [];     // Closing Balance rows -> bold + top border
  var particularsRichText = [];  // {row, boldLen, hasNote} for bold/italic split

  caseIds.forEach(function (disbId) {
    var events = eventsByCase[disbId].slice();
    events.sort(function (a, b) {
      var da = toDate_(a.date), db = toDate_(b.date);
      if (!da && !db) return 0;
      if (!da) return 1;
      if (!db) return -1;
      return da - db;
    });

    sectionHeaderRows.push(out.length);
    out.push([caseLabels[disbId], "", "", "", "", "", ""]);

    var runDebit = 0, runCredit = 0;
    events.forEach(function (ev) {
      runDebit += ev.debit;
      runCredit += ev.credit;
      var bal = runDebit - runCredit;
      if (Math.abs(bal) < 1) bal = 0; // sub-rupee residue rounds to settled

      var rowIdx = out.length;
      var particulars = ev.note ? (ev.bold + "\n" + ev.note) : ev.bold;
      out.push([
        ev.date, particulars, ev.vchType, ev.vchNo,
        ev.debit || "", ev.credit || "", bal
      ]);
      particularsRichText.push({ row: rowIdx, boldLen: ev.bold.length, hasNote: !!ev.note });
    });

    var closingBal = runDebit - runCredit;
    if (Math.abs(closingBal) < 1) closingBal = 0;

    sectionTotalRows.push(out.length);
    out.push(["", "Closing Balance", "", "", runDebit, runCredit, closingBal]);
    out.push(["", "", "", "", "", "", ""]); // spacer row between case sections
  });

  return {
    headers: headers,
    out: out,
    sectionHeaderRows: sectionHeaderRows,
    sectionTotalRows: sectionTotalRows,
    particularsRichText: particularsRichText
  };
}

// Writes a buildLedgerData_() result into the given sheet name (creating it
// if needed), with all the Tally-style formatting.
function writeLedgerToSheet_(sheetName, data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ledgerSheet = ss.getSheetByName(sheetName);
  if (!ledgerSheet) {
    ledgerSheet = ss.insertSheet(sheetName);
  } else {
    ledgerSheet.clear();
    ledgerSheet.clearFormats();
    ledgerSheet.getBandings().forEach(function (b) { b.remove(); });
    ledgerSheet.clearConditionalFormatRules();
  }

  var out = data.out;
  var numRows = out.length;
  var numCols = data.headers.length;
  ledgerSheet.getRange(1, 1, numRows, numCols).setValues(out);

  ledgerSheet.getRange(1, 1, 1, numCols).setFontWeight("bold");
  ledgerSheet.setFrozenRows(1);

  data.sectionHeaderRows.forEach(function (r) {
    var rng = ledgerSheet.getRange(r + 1, 1, 1, numCols);
    rng.merge();
    rng.setFontWeight("bold").setFontSize(11).setHorizontalAlignment("left");
  });

  data.sectionTotalRows.forEach(function (r) {
    ledgerSheet.getRange(r + 1, 1, 1, numCols).setFontWeight("bold");
    ledgerSheet.getRange(r + 1, 1, 1, numCols)
      .setBorder(true, null, null, null, null, null, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  });

  var particularsCol = 2;
  ledgerSheet.getRange(2, particularsCol, numRows - 1, 1)
    .setWrap(true).setHorizontalAlignment("left").setVerticalAlignment("top");

  data.particularsRichText.forEach(function (info) {
    var cell = ledgerSheet.getRange(info.row + 1, particularsCol);
    var text = cell.getValue();
    if (!text) return;
    var builder = SpreadsheetApp.newRichTextValue().setText(text);
    builder.setTextStyle(0, info.boldLen, SpreadsheetApp.newTextStyle().setBold(true).build());
    if (info.hasNote) {
      builder.setTextStyle(info.boldLen + 1, text.length,
        SpreadsheetApp.newTextStyle().setItalic(true).setBold(false).build());
    }
    cell.setRichTextValue(builder.build());
  });

  ledgerSheet.getRange(2, 1, numRows - 1, 1).setNumberFormat("d-mmm-yy");

  var moneyCols = [5, 6, 7]; // Debit, Credit, Balance
  moneyCols.forEach(function (c) {
    ledgerSheet.getRange(2, c, numRows - 1, 1)
      .setNumberFormat("#,##0.00;(#,##0.00)")
      .setHorizontalAlignment("right");
  });

  ledgerSheet.getRange(1, 1, numRows, numCols)
    .setBorder(true, true, true, true, true, true, "#cccccc", SpreadsheetApp.BorderStyle.SOLID);

  // Narrow column widths so the whole ledger fits on screen without
  // horizontal scrolling.
  ledgerSheet.setColumnWidth(1, 90);   // Date
  ledgerSheet.setColumnWidth(2, 320);  // Particulars (wide, wraps to 2 lines)
  ledgerSheet.setColumnWidth(3, 100);  // Vch Type
  ledgerSheet.setColumnWidth(4, 120);  // Vch No.
  ledgerSheet.setColumnWidth(5, 100);  // Debit
  ledgerSheet.setColumnWidth(6, 100);  // Credit
  ledgerSheet.setColumnWidth(7, 100);  // Balance
}

// Rebuilds the live "Customer Ledger" tab: every case with activity this
// month, plus every still-open case from any earlier month after the
// cutoff -- each shown with its full history (see buildLedgerData_).
function rebuildLedger_impl_() {
  var data = buildLedgerData_();
  writeLedgerToSheet_(LEDGER_SHEET, data);
}

function rebuildLedger() {
  var lock = LockService.getScriptLock();
  var gotLock = lock.tryLock(10000);
  if (!gotLock) return; // another run is already in progress, skip this one
  try {
    rebuildLedger_impl_();
  } finally {
    lock.releaseLock();
  }
}

// ── Monthly rollover ──────────────────────────────────────────────────────

// Runs automatically on the 1st of every month (see setupTriggers_). Moves
// the month that just ended out of Accounts into a dated archive tab,
// registers it for the widget to keep reading, then rebuilds the live
// ledger so it immediately reflects the new month.
function monthEndRollover() {
  var lock = LockService.getScriptLock();
  var gotLock = lock.tryLock(30000);
  if (!gotLock) return;

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var accSheet = ss.getSheetByName(ACCOUNTS_SHEET);
    if (!accSheet) {
      throw new Error("Could not find Accounts tab.");
    }

    var now = new Date();
    var lastMonthDate = new Date(now.getFullYear(), now.getMonth() - 1, 1);
    var monthLabel = Utilities.formatDate(lastMonthDate, Session.getScriptTimeZone(), "MMMM yy"); // e.g. "June 26"

    var accData = accSheet.getDataRange().getValues();
    var headerRow1 = accData[0]; // title row
    var headerRow2 = accData[1]; // actual column headers

    var moveRows = [];
    var moveRowNums = [];

    for (var i = 2; i < accData.length; i++) {
      var row = accData[i];
      if (!row[0]) continue;
      var disbDate = toDate_(row[1]);
      if (disbDate && disbDate.getFullYear() === lastMonthDate.getFullYear() &&
          disbDate.getMonth() === lastMonthDate.getMonth()) {
        moveRows.push(row);
        moveRowNums.push(i + 1);
      }
    }

    if (moveRows.length === 0) {
      Logger.log("monthEndRollover: no Accounts rows found for " + monthLabel + " -- nothing to archive.");
    } else {
      var archiveSheet = ss.getSheetByName(monthLabel);
      if (!archiveSheet) archiveSheet = ss.insertSheet(monthLabel);
      archiveSheet.getRange(1, 1, 1, headerRow1.length).setValues([headerRow1]);
      archiveSheet.getRange(2, 1, 1, headerRow2.length).setValues([headerRow2]);
      archiveSheet.getRange(1, 1, 2, headerRow1.length).setFontWeight("bold");
      archiveSheet.getRange(3, 1, moveRows.length, moveRows[0].length).setValues(moveRows);

      // Remove the moved rows from Accounts, highest row number first so
      // earlier deletions don't shift the indices of rows still to delete.
      moveRowNums.sort(function (a, b) { return b - a; }).forEach(function (r) { accSheet.deleteRow(r); });

      var tabs = getArchiveTabNames_();
      if (tabs.indexOf(monthLabel) === -1) {
        tabs.push(monthLabel);
        setConfigValue_('active_archive_tabs', tabs);
      }

      Logger.log("monthEndRollover: archived " + moveRows.length + " row(s) into '" + monthLabel + "'.");
    }

    // Refresh the live ledger now that Accounts has shifted to the new
    // month (no separate frozen snapshot -- the live ledger's own
    // current-month-or-still-open inclusion rule always reflects the
    // correct picture, so there's nothing extra to freeze).
    rebuildLedger_impl_();

    Logger.log("monthEndRollover: rollover complete for '" + monthLabel + "', live ledger rebuilt.");
  } finally {
    lock.releaseLock();
  }
}

function setupTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function (t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger("rebuildLedger")
    .timeBased()
    .everyMinutes(10)
    .create();
  ScriptApp.newTrigger("monthEndRollover")
    .timeBased()
    .onMonthDay(1)
    .atHour(1)
    .create();
}

// Run this once after deleting old sheets/triggers.
function freshSetup() {
  var result = removePreJuneData();
  rebuildLedger();
  setupTriggers_();
  SpreadsheetApp.getUi().alert(
    "Setup complete. Removed " + result.accRemoved + " pre-June disbursement row(s) and " +
    result.mcRemoved + " pre-June M Coll row(s). Customer Ledger rebuilt as a Tally-style " +
    "per-customer ledger (this month's activity + all still-open cases). It will auto-refresh " +
    "every 10 minutes, and the monthly Accounts archival will run itself on the 1st of each month."
  );
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Ledger Tools")
    .addItem("Run Fresh Setup (remove pre-June + rebuild)", "freshSetup")
    .addItem("Rebuild Ledger Now", "rebuildLedger")
    .addItem("Run Month-End Rollover Now", "monthEndRollover")
    .addItem("Remove Pre-June Data Only", "removePreJuneData")
    .addToUi();
}
// Web App endpoint - lets the widget trigger an instant ledger rebuild.
//
// Append this whole block to the bottom of the same consolidated script
// (the one with rebuildLedger / removePreJuneData / freshSetup already in it).
// Do not create a second script project - this must live in the SAME script
// so it can see the same rebuildLedger function and the same spreadsheet.
//
// One-time setup:
//   1. Paste this block into the script, Save.
//   2. Run setupWebhookToken once from the function dropdown. Approve the
//      permission prompt if asked. Check the log (View -> Logs) for the
//      generated token - copy it somewhere safe, you will need it.
//   3. Deploy -> New deployment -> select type "Web app".
//        Description: anything, e.g. "Ledger webhook"
//        Execute as: Me
//        Who has access: Anyone
//      Click Deploy, authorize again if asked, then copy the Web app URL
//      it gives you (ends in /exec).
//   4. Paste that URL into ledger_webhook_url and the token from step 2 into
//      ledger_webhook_token in bridgeline_config.json (already wired up in
//      the widget code - it reads this file fresh on every save, no restart
//      needed). After every disbursement or repayment save, the widget will
//      call:
//        POST <web app url>
//        body (JSON): { "token": "<the token from step 2>" }
//
// If you ever need to invalidate the old token (e.g. it leaked), just run
// setupWebhookToken again - it generates a new one and the old one stops
// working immediately.

function setupWebhookToken() {
  var token = Utilities.getUuid();
  PropertiesService.getScriptProperties().setProperty("WEBHOOK_TOKEN", token);
  Logger.log("Webhook token (copy this): " + token);
  SpreadsheetApp.getUi().alert(
    "New webhook token generated. Open View > Logs (or the Executions page) " +
    "to copy it - it will not be shown again here. Token: " + token
  );
}

function handleWebhook_(token) {
  var expected = PropertiesService.getScriptProperties().getProperty("WEBHOOK_TOKEN");
  if (!expected || token !== expected) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: "invalid token" }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  try {
    rebuildLedger();
    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  var token = e && e.parameter ? e.parameter.token : null;
  return handleWebhook_(token);
}

function doPost(e) {
  var token = null;
  if (e && e.parameter && e.parameter.token) {
    token = e.parameter.token;
  } else if (e && e.postData && e.postData.contents) {
    try {
      var body = JSON.parse(e.postData.contents);
      token = body.token;
    } catch (err) {
      token = null;
    }
  }
  return handleWebhook_(token);
}
