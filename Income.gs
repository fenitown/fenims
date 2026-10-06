/*******************************************************
 * TCB-স্টাইল ডিলার ম্যানেজমেন্ট সফটওয়্যার
 * Income.gs — আয় রশিদ (ডিপু ও ডিলার — দুই সাইটেই)
 *
 * MasterRegistry এর একই Apps Script প্রজেক্টে নতুন ফাইল
 * হিসেবে যোগ করুন (নাম দিন: Income)
 *
 * - ডিপু (AGENCY) লগইনে ডাটা Master Spreadsheet এর "IncomeReceipt" ট্যাবে,
 *   ডিলার লগইনে সেই ডিলারের নিজের শীটের "IncomeReceipt" ট্যাবে যায়
 * - এক রশিদে একাধিক সারি (ক্রম, বিবরণ, টাকা); রশিদ নং অটো তৈরি হয়
 * - ট্যাব না থাকলে নিজে থেকে তৈরি হয়, আলাদা সেটআপ লাগে না
 * - যোগ/দেখা: Admin + প্রতিনিধি (ডিপুতে শুধু Admin); এডিট/ডিলিট: শুধু Admin
 *******************************************************/

const INCOME_TAB = "IncomeReceipt";

function checkIncomePermission(token, adminOnly) {
  const perm = checkPermission(token, adminOnly ? ["Admin"] : ["Admin", "প্রতিনিধি"]);
  if (!perm.ok) return perm;
  if (perm.payload.dealerId === AGENCY_ID && perm.payload.role !== "Admin") {
    return { ok: false, message: "এই কাজ করার অনুমতি আপনার নেই" };
  }
  return perm;
}

function getIncomeSheetFor(perm) {
  const ss = perm.payload.dealerId === AGENCY_ID ? getMasterSS() : getDealerSpreadsheet(perm.payload.dealerId);
  let sheet = ss.getSheetByName(INCOME_TAB);
  if (!sheet) {
    const def = {};
    def[INCOME_TAB] = INCOME_HEADERS;
    ensureSheetsWithHeaders(ss, def);
    sheet = ss.getSheetByName(INCOME_TAB);
  }
  return sheet;
}

/* বিদ্যমান সারির সর্বোচ্চ ক্রমিক সংখ্যা (রশিদ নং / EntryID এর শেষের অংশ থেকে) */
function maxSeqInColumn(values, colIdx) {
  let max = 0;
  for (let i = 1; i < values.length; i++) {
    const m = String(values[i][colIdx] || "").match(/(\d+)$/);
    if (m) max = Math.max(max, Number(m[1]));
  }
  return max;
}

function addIncomeReceipt(data) {
  const perm = checkIncomePermission(data.token, false);
  if (!perm.ok) return { success: false, message: perm.message };

  const entries = (data.entries || []).filter(function (e) {
    return String(e["বিবরণ"] || "").trim() !== "";
  });
  if (entries.length === 0) return { success: false, message: "অন্তত একটি এন্ট্রি দিন" };

  const sheet = getIncomeSheetFor(perm);
  const date = data.date ? new Date(data.date) : new Date();

  let grandTotal = 0;
  entries.forEach(function (e) { grandTotal += Number(e["টাকা"]) || 0; });

  // দুজন একসাথে সংরক্ষণ করলেও যেন একই রশিদ নং না হয়
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  let receiptNo;
  try {
    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const values = sheet.getLastRow() > 0 ? sheet.getDataRange().getValues() : [headers];
    const idxReceipt = headers.indexOf("রশিদ নং");
    const idxEntry = headers.indexOf("EntryID");

    receiptNo = "IR" + Utilities.formatString("%04d", maxSeqInColumn(values, idxReceipt) + 1);
    let nextEntry = maxSeqInColumn(values, idxEntry);

    const rows = entries.map(function (e, i) {
      nextEntry++;
      const obj = {
        "EntryID": "INE" + Utilities.formatString("%04d", nextEntry),
        "রশিদ নং": receiptNo,
        "তারিখ": date,
        "ক্রম": i + 1,
        "বিবরণ": String(e["বিবরণ"]).trim(),
        "টাকা": Number(e["টাকা"]) || 0,
        "সর্বমোট": grandTotal
      };
      return headers.map(function (h) { return obj.hasOwnProperty(h) ? obj[h] : ""; });
    });
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
  } finally {
    lock.releaseLock();
  }

  return {
    success: true, receiptNo: receiptNo, total: grandTotal,
    message: "আয় রশিদ সংরক্ষিত হয়েছে (" + entries.length + " টি এন্ট্রি)"
  };
}

function listIncomeReceipts(data) {
  const perm = checkIncomePermission(data.token, false);
  if (!perm.ok) return { success: false, message: perm.message };
  return { success: true, receipts: genericListRows(getIncomeSheetFor(perm)) };
}

/* একটি সারির বিবরণ/টাকা এডিট — পুরো রশিদের সর্বমোট নতুন করে হিসাব হয় */
function updateIncomeReceiptEntry(data) {
  const perm = checkIncomePermission(data.token, true);
  if (!perm.ok) return { success: false, message: perm.message };

  const sheet = getIncomeSheetFor(perm);
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const idxEntry = headers.indexOf("EntryID");
  const idxReceipt = headers.indexOf("রশিদ নং");
  const idxDesc = headers.indexOf("বিবরণ");
  const idxAmount = headers.indexOf("টাকা");
  const idxTotal = headers.indexOf("সর্বমোট");

  let target = -1;
  for (let i = 1; i < values.length; i++) {
    if (values[i][idxEntry] === data.entryId) { target = i; break; }
  }
  if (target === -1) return { success: false, message: "আয় এন্ট্রি পাওয়া যায়নি" };

  const fields = data.fields || {};
  if (fields.hasOwnProperty("বিবরণ")) values[target][idxDesc] = String(fields["বিবরণ"]).trim();
  if (fields.hasOwnProperty("টাকা")) values[target][idxAmount] = Number(fields["টাকা"]) || 0;

  const receiptNo = values[target][idxReceipt];
  let total = 0;
  for (let i = 1; i < values.length; i++) {
    if (values[i][idxReceipt] === receiptNo) total += Number(values[i][idxAmount]) || 0;
  }

  sheet.getRange(target + 1, idxDesc + 1).setValue(values[target][idxDesc]);
  sheet.getRange(target + 1, idxAmount + 1).setValue(values[target][idxAmount]);
  for (let i = 1; i < values.length; i++) {
    if (values[i][idxReceipt] === receiptNo) sheet.getRange(i + 1, idxTotal + 1).setValue(total);
  }
  return { success: true, message: "আপডেট হয়েছে" };
}

/* পুরো রশিদ (তার সব সারি) ডিলিট */
function deleteIncomeReceipt(data) {
  const perm = checkIncomePermission(data.token, true);
  if (!perm.ok) return { success: false, message: perm.message };

  const sheet = getIncomeSheetFor(perm);
  const values = sheet.getDataRange().getValues();
  const idx = values[0].indexOf("রশিদ নং");

  let found = false;
  for (let i = values.length - 1; i >= 1; i--) {
    if (values[i][idx] === data.receiptNo) {
      sheet.deleteRow(i + 1);
      found = true;
    }
  }
  return { success: found, message: found ? "রশিদ ডিলিট হয়েছে" : "আয় রশিদ পাওয়া যায়নি" };
}
