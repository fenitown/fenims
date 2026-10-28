/*******************************************************
 * TCB-স্টাইল ডিলার ম্যানেজমেন্ট সফটওয়্যার
 * Payments.gs — (১) স্টক ভাউচার পুরোটা এডিট
 *               (২) কোম্পানির বকেয়া পরিশোধ (স্টক ভাউচার ভিত্তিক / সকল বকেয়া)
 *               (৩) ডিলারের কাছ থেকে বকেয়া আদায় (ইনভয়েস ভিত্তিক / ডিলারের সকল / সকল ডিলারের)
 *
 * MasterRegistry এর একই Apps Script প্রজেক্টে নতুন ফাইল হিসেবে যোগ করুন (নাম: Payments)
 * সব একশন শুধুমাত্র মূল এজেন্সি (AGENCY) লগইন করতে পারবে
 *
 * কাজের নিয়ম:
 *  - প্রতিটি পরিশোধ/আদায় একটি ইতিহাস-ট্যাবে (StockPayment / DealerCollection) লেখা হয় —
 *    ট্যাব না থাকলে নিজে থেকে তৈরি হয়
 *  - একই সাথে ভাউচার/ইনভয়েসের সব সারির "পরিশোধ" বাড়ে ও "বকেয়া" কমে, তাই ডিপু রিপোর্ট ও
 *    ডিলার রিপোর্টের বকেয়া নিজে থেকেই ঠিক থাকে
 *  - "সকল বকেয়া" দিলে টাকা পুরনো তারিখের ভাউচার/ইনভয়েস থেকে আগে মেটানো হয়
 *  - বকেয়ার চেয়ে বেশি টাকা নেওয়া হয় না
 *******************************************************/

const STOCK_PAY_TAB = "StockPayment";
const STOCK_PAY_HEADERS = ["PaymentID", "তারিখ", "ভাউচার নং", "টাকা", "নোট"];
const DEALER_COL_TAB = "DealerCollection";
const DEALER_COL_HEADERS = ["CollectionID", "তারিখ", "ইনভয়েস নং", "DealerID", "ডিলার নাম", "টাকা", "নোট"];

function payRound(n) { return Math.round((Number(n) || 0) * 100) / 100; }

function getOrCreateTab(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/* কলামের শেষ সংখ্যার সর্বোচ্চ মান ধরে পরের আইডি — সারি ডিলিট হলেও আইডি দ্বিগুণ হয় না */
function makeIdGenerator(sheet, idHeader, prefix, width) {
  let max = 0;
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const idx = headers.indexOf(idHeader);
    if (idx !== -1) {
      sheet.getRange(2, idx + 1, lastRow - 1, 1).getValues().forEach(function (r) {
        const m = String(r[0] || "").match(/(\d+)$/);
        if (m) max = Math.max(max, Number(m[1]));
      });
    }
  }
  return function () { max++; return prefix + Utilities.formatString("%0" + width + "d", max); };
}

/* ভাউচার/ইনভয়েস অনুযায়ী সারির গ্রুপ — প্রতিটি গ্রুপে প্রথম সারির পরিশোধ/বকেয়া/তারিখ */
function payGroupRows(values, headers, keyHeader) {
  const iK = headers.indexOf(keyHeader), iD = headers.indexOf("তারিখ"),
        iPaid = headers.indexOf("পরিশোধ"), iDue = headers.indexOf("বকেয়া"),
        iDealer = headers.indexOf("DealerID"), iDealerName = headers.indexOf("ডিলার নাম");
  const groups = {}, order = [];
  for (let i = 1; i < values.length; i++) {
    const k = String(values[i][iK] || "");
    if (!k) continue;
    if (!groups[k]) {
      groups[k] = {
        key: k, date: values[i][iD], paid: Number(values[i][iPaid]) || 0, due: Number(values[i][iDue]) || 0,
        dealerId: iDealer === -1 ? "" : String(values[i][iDealer] || ""),
        dealerName: iDealerName === -1 ? "" : String(values[i][iDealerName] || ""),
        rows: []
      };
      order.push(k);
    }
    groups[k].rows.push(i);
  }
  return { groups: groups, order: order, iPaid: iPaid, iDue: iDue };
}

function paySortedKeys(G) {
  return G.order.slice().sort(function (a, b) {
    const da = new Date(G.groups[a].date).getTime() || 0, db = new Date(G.groups[b].date).getTime() || 0;
    return (da - db) || (a < b ? -1 : (a > b ? 1 : 0));
  });
}

/* deltas: { key: টাকা } — ধনাত্মক = পরিশোধ/আদায় (পরিশোধ বাড়ে, বকেয়া কমে), ঋণাত্মক = ফেরত */
function payApplyDeltas(sheet, values, G, deltas) {
  const n = values.length - 1;
  if (n < 1 || G.iPaid === -1 || G.iDue === -1) throw new Error("পরিশোধ/বকেয়া কলাম পাওয়া যায়নি");
  const paidCol = [], dueCol = [];
  for (let i = 1; i < values.length; i++) {
    paidCol.push([values[i][G.iPaid]]);
    dueCol.push([values[i][G.iDue]]);
  }
  Object.keys(deltas).forEach(function (k) {
    const g = G.groups[k];
    if (!g) return;
    const d = deltas[k];
    g.rows.forEach(function (i) {
      paidCol[i - 1][0] = payRound((Number(values[i][G.iPaid]) || 0) + d);
      dueCol[i - 1][0] = payRound((Number(values[i][G.iDue]) || 0) - d);
    });
  });
  sheet.getRange(2, G.iPaid + 1, n, 1).setValues(paidCol);
  sheet.getRange(2, G.iDue + 1, n, 1).setValues(dueCol);
}

/* সাধারণ বরাদ্দ: নির্দিষ্ট key হলে শুধু সেটি, না হলে পুরনো থেকে নতুন ক্রমে */
function payAllocate(G, keys, amount) {
  keys = keys.filter(function (k) { return G.groups[k].due > 0.0001; });
  if (keys.length === 0) return { error: "কোনো বকেয়া নেই" };
  let totalDue = 0;
  keys.forEach(function (k) { totalDue += G.groups[k].due; });
  if (amount > totalDue + 0.0001) return { error: "টাকা বকেয়ার (" + payRound(totalDue) + ") চেয়ে বেশি হতে পারবে না" };
  let remaining = amount;
  const allocations = [];
  keys.forEach(function (k) {
    if (remaining <= 0.0001) return;
    const pay = payRound(Math.min(remaining, G.groups[k].due));
    if (pay > 0) { allocations.push({ key: k, pay: pay }); remaining = payRound(remaining - pay); }
  });
  return { allocations: allocations };
}

/*=========================================================
 *  (২) কোম্পানির বকেয়া পরিশোধ
 *  data: { token, date, voucherNo ("" = সকল বকেয়া), amount, note }
 *=======================================================*/
function addStockPayment(data) {
  const perm = checkAgencyPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };
  const amount = payRound(data.amount);
  if (!(amount > 0)) return { success: false, message: "টাকার পরিমাণ দিন" };

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = getMasterSS();
    const vSheet = getSheet(ss, "StockInVoucher");
    const values = vSheet.getDataRange().getValues();
    if (values.length < 2) return { success: false, message: "কোনো স্টক ভাউচার নেই" };
    const G = payGroupRows(values, values[0], "ভাউচার নং");

    let keys;
    if (data.voucherNo) {
      if (!G.groups[data.voucherNo]) return { success: false, message: "ভাউচার পাওয়া যায়নি" };
      keys = [data.voucherNo];
    } else {
      keys = paySortedKeys(G);
    }
    const plan = payAllocate(G, keys, amount);
    if (plan.error) return { success: false, message: plan.error };

    const deltas = {};
    plan.allocations.forEach(function (a) { deltas[a.key] = a.pay; });
    payApplyDeltas(vSheet, values, G, deltas);

    const payTab = getOrCreateTab(ss, STOCK_PAY_TAB, STOCK_PAY_HEADERS);
    const nextId = makeIdGenerator(payTab, "PaymentID", "SP", 4);
    const date = data.date ? new Date(data.date) : new Date();
    const rows = plan.allocations.map(function (a) { return [nextId(), date, a.key, a.pay, String(data.note || "")]; });
    payTab.getRange(payTab.getLastRow() + 1, 1, rows.length, STOCK_PAY_HEADERS.length).setValues(rows);

    return { success: true, applied: plan.allocations.length, message: "পরিশোধ সংরক্ষিত হয়েছে (" + plan.allocations.length + " টি ভাউচারে)" };
  } finally {
    lock.releaseLock();
  }
}

function listStockPayments(data) {
  const perm = checkAgencyPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };
  return { success: true, payments: genericListRows(getOrCreateTab(getMasterSS(), STOCK_PAY_TAB, STOCK_PAY_HEADERS)) };
}

/* ভুল পরিশোধ মুছলে ভাউচারের পরিশোধ কমে ও বকেয়া ফিরে আসে */
function deleteStockPayment(data) {
  const perm = checkAgencyPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = getMasterSS();
    const payTab = getOrCreateTab(ss, STOCK_PAY_TAB, STOCK_PAY_HEADERS);
    const pv = payTab.getDataRange().getValues();
    const h = pv[0];
    let row = -1;
    for (let i = 1; i < pv.length; i++) { if (pv[i][h.indexOf("PaymentID")] === data.paymentId) { row = i; break; } }
    if (row === -1) return { success: false, message: "পরিশোধ রেকর্ড পাওয়া যায়নি" };
    const amount = Number(pv[row][h.indexOf("টাকা")]) || 0;
    const voucherNo = String(pv[row][h.indexOf("ভাউচার নং")] || "");

    const vSheet = getSheet(ss, "StockInVoucher");
    const values = vSheet.getDataRange().getValues();
    if (values.length >= 2) {
      const G = payGroupRows(values, values[0], "ভাউচার নং");
      if (G.groups[voucherNo]) { const d = {}; d[voucherNo] = -amount; payApplyDeltas(vSheet, values, G, d); }
    }
    payTab.deleteRow(row + 1);
    return { success: true, message: "পরিশোধ মুছে ফেলা হয়েছে" };
  } finally {
    lock.releaseLock();
  }
}

/*=========================================================
 *  (৩) ডিলারের কাছ থেকে বকেয়া আদায়
 *  data: { token, date, dealerId (ঐচ্ছিক), invoiceNo (ঐচ্ছিক), amount, note }
 *   - invoiceNo দিলে শুধু সেই ইনভয়েস
 *   - শুধু dealerId দিলে সেই ডিলারের সব বকেয়া ইনভয়েস (পুরনো থেকে)
 *   - কিছুই না দিলে সকল ডিলারের সব বকেয়া ইনভয়েস (পুরনো থেকে)
 *=======================================================*/
function addDealerCollection(data) {
  const perm = checkAgencyPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };
  const amount = payRound(data.amount);
  if (!(amount > 0)) return { success: false, message: "টাকার পরিমাণ দিন" };

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = getMasterSS();
    const sheet = getSheet(ss, "SalesInvoice");
    const values = sheet.getDataRange().getValues();
    if (values.length < 2) return { success: false, message: "কোনো ইনভয়েস নেই" };
    const G = payGroupRows(values, values[0], "ইনভয়েস নং");

    let keys;
    if (data.invoiceNo) {
      if (!G.groups[data.invoiceNo]) return { success: false, message: "ইনভয়েস পাওয়া যায়নি" };
      keys = [data.invoiceNo];
    } else {
      keys = paySortedKeys(G);
      if (data.dealerId) keys = keys.filter(function (k) { return G.groups[k].dealerId === data.dealerId; });
    }
    const plan = payAllocate(G, keys, amount);
    if (plan.error) return { success: false, message: plan.error };

    const deltas = {};
    plan.allocations.forEach(function (a) { deltas[a.key] = a.pay; });
    payApplyDeltas(sheet, values, G, deltas);

    const colTab = getOrCreateTab(ss, DEALER_COL_TAB, DEALER_COL_HEADERS);
    const nextId = makeIdGenerator(colTab, "CollectionID", "DC", 4);
    const date = data.date ? new Date(data.date) : new Date();
    const rows = plan.allocations.map(function (a) {
      const g = G.groups[a.key];
      return [nextId(), date, a.key, g.dealerId, g.dealerName, a.pay, String(data.note || "")];
    });
    colTab.getRange(colTab.getLastRow() + 1, 1, rows.length, DEALER_COL_HEADERS.length).setValues(rows);

    return { success: true, applied: plan.allocations.length, message: "আদায় সংরক্ষিত হয়েছে (" + plan.allocations.length + " টি ইনভয়েসে)" };
  } finally {
    lock.releaseLock();
  }
}

function listDealerCollections(data) {
  const perm = checkAgencyPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };
  return { success: true, collections: genericListRows(getOrCreateTab(getMasterSS(), DEALER_COL_TAB, DEALER_COL_HEADERS)) };
}

function deleteDealerCollection(data) {
  const perm = checkAgencyPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = getMasterSS();
    const colTab = getOrCreateTab(ss, DEALER_COL_TAB, DEALER_COL_HEADERS);
    const cv = colTab.getDataRange().getValues();
    const h = cv[0];
    let row = -1;
    for (let i = 1; i < cv.length; i++) { if (cv[i][h.indexOf("CollectionID")] === data.collectionId) { row = i; break; } }
    if (row === -1) return { success: false, message: "আদায় রেকর্ড পাওয়া যায়নি" };
    const amount = Number(cv[row][h.indexOf("টাকা")]) || 0;
    const invoiceNo = String(cv[row][h.indexOf("ইনভয়েস নং")] || "");

    const sheet = getSheet(ss, "SalesInvoice");
    const values = sheet.getDataRange().getValues();
    if (values.length >= 2) {
      const G = payGroupRows(values, values[0], "ইনভয়েস নং");
      if (G.groups[invoiceNo]) { const d = {}; d[invoiceNo] = -amount; payApplyDeltas(sheet, values, G, d); }
    }
    colTab.deleteRow(row + 1);
    return { success: true, message: "আদায় মুছে ফেলা হয়েছে" };
  } finally {
    lock.releaseLock();
  }
}

/*=========================================================
 *  (১) স্টক ভাউচার পুরোটা এডিট (সব পণ্য-লাইন নতুন করে)
 *  data: { token, voucherNo, items:[{productId, productName, brand, purchasePrice, quantity}], date, paid }
 *  "paid" = ভাউচারের এখনকার মোট পরিশোধ (পরে করা বকেয়া-পরিশোধসহ)
 *=======================================================*/
function updateStockVoucherFull(data) {
  const perm = checkAgencyPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const masterSS = getMasterSS();
    const sheet = getSheet(masterSS, "StockInVoucher");
    ensureSheetColumns(sheet, ["ব্র্যান্ড", "ক্রয় মূল্য"]);

    const values = sheet.getDataRange().getValues();
    const idx = values[0].indexOf("ভাউচার নং");
    let existing = 0;
    for (let i = values.length - 1; i >= 1; i--) { if (values[i][idx] === data.voucherNo) existing++; }
    if (existing === 0) return { success: false, message: "ভাউচার পাওয়া যায়নি" };

    const items = (data.items || []).filter(function (i) { return i.productId && Number(i.quantity) > 0; });
    if (items.length === 0) return { success: false, message: "অন্তত একটি পণ্য ও সংখ্যা দিন" };

    for (let i = values.length - 1; i >= 1; i--) {
      if (values[i][idx] === data.voucherNo) sheet.deleteRow(i + 1);
    }

    const date = data.date ? new Date(data.date) : new Date();
    let grandTotal = 0;
    const lines = items.map(function (item) {
      const price = Number(item.purchasePrice) || 0;
      const total = (Number(item.quantity) || 0) * price;
      grandTotal += total;
      return { item: item, price: price, total: total };
    });
    const paid = Number(data.paid) || 0;
    const due = grandTotal - paid;

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const nextId = makeIdGenerator(sheet, "EntryID", "SVE", 4);
    const rows = lines.map(function (ld) {
      const obj = {
        "EntryID": nextId(), "ভাউচার নং": data.voucherNo, "তারিখ": date,
        "ProductID": ld.item.productId, "পণ্যের নাম": ld.item.productName || "",
        "ব্র্যান্ড": ld.item.brand || "", "ক্রয় মূল্য": ld.price, "সংখ্যা": ld.item.quantity,
        "মোট মূল্য": ld.total, "সর্বমোট": grandTotal, "পরিশোধ": paid, "বকেয়া": due
      };
      return headers.map(function (h) { return obj.hasOwnProperty(h) ? obj[h] : ""; });
    });
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);

    return { success: true, voucherNo: data.voucherNo, total: grandTotal, paid: paid, due: due, message: "স্টক ভাউচার আপডেট হয়েছে" };
  } finally {
    lock.releaseLock();
  }
}
