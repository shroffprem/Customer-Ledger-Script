// BridgeLine Ledger - consolidated script
//
// What this does, in one run:
//   1. Removes all pre-June-2026 rows from "Accounts" and "M Coll" (archiving
//      them first into "Accounts_Archive_PreJune" / "MColl_Archive_PreJune").
//   2. Rebuilds the "Customer Ledger" tab from scratch as a Tally-style
//      accounting ledger: one section per customer, columns
//      Date | Particulars | Vch Type | Vch No. | Debit | Credit | Balance,
//      with a running balance per section and a bold double-ruled subtotal
//      row closing each customer's section. Reference text (Debit Note /
//      Credit Note / M-Coll note) is shown exactly as entered -- nothing
//      here re-parses or can silently blank it.
//   3. Installs ONE time-based trigger (every 10 minutes) so it keeps itself
//      current automatically. No onChange trigger is used, because onChange
//      fires on the script's own writes too and causes an endless rebuild
//      loop - that loop is what was making the sheet flicker/glitch before.
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
// manual options, and the ledger refreshes on its own every 10 minutes.

var ACCOUNTS_SHEET = "Accounts";
var MCOLL_SHEET = "M Coll";
var LEDGER_SHEET = "Customer Ledger";
var ACC_ARCHIVE_SHEET = "Accounts_Archive_PreJune";
var MC_ARCHIVE_SHEET = "MColl_Archive_PreJune";
var CUTOFF = new Date(2026, 5, 1); // June 1, 2026

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

function getOrCreateArchive_(ss, name, headerRow) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow]);
    sheet.getRange(1, 1, 1, headerRow.length).setFontWeight("bold");
  }
  return sheet;
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

function rebuildLedger() {
  var lock = LockService.getScriptLock();
  var gotLock = lock.tryLock(10000);
  if (!gotLock) return; // another run is already in progress, skip this one

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var accSheet = ss.getSheetByName(ACCOUNTS_SHEET);
    var mcSheet = ss.getSheetByName(MCOLL_SHEET);
    if (!accSheet || !mcSheet) {
      throw new Error("Could not find Accounts or M Coll tabs.");
    }

    var accData = accSheet.getDataRange().getValues();
    var accRows = accData.slice(2).filter(function (r) { return r[0]; });

    var mcData = mcSheet.getDataRange().getValues();
    var mcRows = mcData.slice(1).filter(function (r) { return r[0]; });

    var mcByDisb = {};
    mcRows.forEach(function (r) {
      var disbId = r[0];
      if (!mcByDisb[disbId]) mcByDisb[disbId] = [];
      mcByDisb[disbId].push({ date: r[1], amount: r[2], note: r[3] });
    });

    // One flat list of ledger events (Debit = disbursement, Credit = collection),
    // grouped by customer so they can become per-customer sections below.
    // Reference text (Debit Note / Credit Note / M-Coll note) flows straight
    // through unparsed -- there is no regex here that can silently blank it.
    var eventsByCustomer = {};

    accRows.forEach(function (r) {
      var disbId = r[0], disbDate = r[1];
      var customer = (r[2] || '').toString().trim() || '(No Name)';
      var amount = r[7], charges = r[8], gst = r[9];
      var collDate = r[11], collAmt = r[12];
      var debitNote = r[21], creditNote = r[22];

      if (!eventsByCustomer[customer]) eventsByCustomer[customer] = [];

      eventsByCustomer[customer].push({
        date: disbDate,
        bold: "Disbursed to " + customer,
        note: debitNote ? String(debitNote) : "",
        vchType: "Disbursement",
        vchNo: disbId,
        debit: Number(amount) || 0,
        credit: 0
      });

      // Charges and GST are owed by the customer too (Total Payable =
      // Amount + Charges + GST) -- shown as their own debit rows rather
      // than folded into the disbursement figure, so each is independently
      // traceable in the running balance.
      if (Number(charges) > 0) {
        eventsByCustomer[customer].push({
          date: disbDate,
          bold: "Processing Charges",
          note: "",
          vchType: "Charges",
          vchNo: disbId,
          debit: Number(charges) || 0,
          credit: 0
        });
      }
      if (Number(gst) > 0) {
        eventsByCustomer[customer].push({
          date: disbDate,
          bold: "GST on Charges",
          note: "",
          vchType: "GST",
          vchNo: disbId,
          debit: Number(gst) || 0,
          credit: 0
        });
      }

      var colls = [];
      if (mcByDisb[disbId] && mcByDisb[disbId].length > 0) {
        mcByDisb[disbId].forEach(function (c) {
          colls.push({ date: c.date, amount: c.amount, note: c.note });
        });
      } else if (collAmt) {
        colls.push({ date: collDate, amount: collAmt, note: creditNote });
      }

      colls.forEach(function (c) {
        eventsByCustomer[customer].push({
          date: c.date,
          bold: "Collection received",
          note: c.note ? String(c.note) : "",
          vchType: "Collection",
          vchNo: disbId,
          debit: 0,
          credit: Number(c.amount) || 0
        });
      });
    });

    // Order customer sections chronologically by each customer's earliest
    // transaction date (not alphabetically by name).
    var customerNames = Object.keys(eventsByCustomer).sort(function (a, b) {
      var da = eventsByCustomer[a].reduce(function (min, ev) {
        var d = toDate_(ev.date);
        return (d && (!min || d < min)) ? d : min;
      }, null);
      var db = eventsByCustomer[b].reduce(function (min, ev) {
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
    var sectionTotalRows = [];     // subtotal rows -> bold + top border
    var particularsRichText = [];  // {row, boldLen, hasNote} for bold/italic split

    customerNames.forEach(function (custName) {
      var events = eventsByCustomer[custName].slice();
      events.sort(function (a, b) {
        var da = toDate_(a.date), db = toDate_(b.date);
        if (!da && !db) return 0;
        if (!da) return 1;
        if (!db) return -1;
        return da - db;
      });

      sectionHeaderRows.push(out.length);
      out.push([custName, "", "", "", "", "", ""]);

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
      out.push(["", "Total", "", "", runDebit, runCredit, closingBal]);
      out.push(["", "", "", "", "", "", ""]); // spacer row between customer sections
    });

    var ledgerSheet = ss.getSheetByName(LEDGER_SHEET);
    if (!ledgerSheet) {
      ledgerSheet = ss.insertSheet(LEDGER_SHEET);
    } else {
      ledgerSheet.clear();
      ledgerSheet.clearFormats();
      ledgerSheet.getBandings().forEach(function (b) { b.remove(); });
      ledgerSheet.clearConditionalFormatRules();
    }

    var numRows = out.length;
    var numCols = headers.length;
    ledgerSheet.getRange(1, 1, numRows, numCols).setValues(out);

    ledgerSheet.getRange(1, 1, 1, numCols).setFontWeight("bold");
    ledgerSheet.setFrozenRows(1);

    sectionHeaderRows.forEach(function (r) {
      var rng = ledgerSheet.getRange(r + 1, 1, 1, numCols);
      rng.merge();
      rng.setFontWeight("bold").setFontSize(11).setHorizontalAlignment("left");
    });

    sectionTotalRows.forEach(function (r) {
      ledgerSheet.getRange(r + 1, 1, 1, numCols).setFontWeight("bold");
      ledgerSheet.getRange(r + 1, 1, 1, numCols)
        .setBorder(true, null, null, null, null, null, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    });

    var particularsCol = 2;
    ledgerSheet.getRange(2, particularsCol, numRows - 1, 1)
      .setWrap(true).setHorizontalAlignment("left").setVerticalAlignment("top");

    particularsRichText.forEach(function (info) {
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
    // horizontal scrolling (the old 14-column layout needed it; this doesn't).
    ledgerSheet.setColumnWidth(1, 90);   // Date
    ledgerSheet.setColumnWidth(2, 320);  // Particulars (wide, wraps to 2 lines)
    ledgerSheet.setColumnWidth(3, 100);  // Vch Type
    ledgerSheet.setColumnWidth(4, 120);  // Vch No.
    ledgerSheet.setColumnWidth(5, 100);  // Debit
    ledgerSheet.setColumnWidth(6, 100);  // Credit
    ledgerSheet.setColumnWidth(7, 100);  // Balance
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
}

// Run this once after deleting old sheets/triggers.
function freshSetup() {
  var result = removePreJuneData();
  rebuildLedger();
  setupTriggers_();
  SpreadsheetApp.getUi().alert(
    "Setup complete. Removed " + result.accRemoved + " pre-June disbursement row(s) and " +
    result.mcRemoved + " pre-June M Coll row(s). Customer Ledger rebuilt as a Tally-style " +
    "per-customer ledger. It will auto-refresh every 10 minutes from now on."
  );
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Ledger Tools")
    .addItem("Run Fresh Setup (remove pre-June + rebuild)", "freshSetup")
    .addItem("Rebuild Ledger Now", "rebuildLedger")
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
