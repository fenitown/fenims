/*******************************************************
 * TCB-স্টাইল ডিলার ম্যানেজমেন্ট সফটওয়্যার
 * Finance.gs — ডিপু সাইটের "ইনভেস্ট", "ক্যাশ" মেনু ও নতুন ডিপু রিপোর্ট
 *
 * MasterRegistry এর একই Apps Script প্রজেক্টে নতুন ফাইল হিসেবে যোগ করুন
 * (নাম দিন: Finance)। সব একশন শুধুমাত্র মূল এজেন্সি (AGENCY) Admin পারবে।
 * ট্যাব না থাকলে (Invest, CashDeposit, CashWithdraw) নিজে থেকে তৈরি হয় —
 * আলাদা সেটআপ লাগে না।
 *******************************************************/

const INVEST_TAB = "Invest";
const INVEST_HEADERS = ["InvestID", "তারিখ", "নাম", "মোবাইল নং", "পদবি", "ঠিকানা", "টাকা"];
const CASH_DEP_TAB = "CashDeposit";
const CASH_DEP_HEADERS = ["DepositID", "তারিখ", "টাকা", "জমাকারীর নাম", "মোবাইল নং"];
const CASH_WD_TAB = "CashWithdraw";
const CASH_WD_HEADERS = ["WithdrawID", "তারিখ", "টাকা", "উত্তোলনকারীর নাম", "মোবাইল নং"];

function finSheet(tab, headers) { return getOrCreateTab(getMasterSS(), tab, headers); }
function finNum(v) { return Number(v) || 0; }
function finR2(v) { return Math.round((Number(v) || 0) * 100) / 100; }

/* মোবাইল কলাম টেক্সট ফরম্যাটে রেখে সারি যোগ (শুরুর ০ না হারাতে) */
function finAppendRows(sheet, headers, rows) {
  const start = sheet.getLastRow() + 1;
  const mobCol = headers.indexOf("মোবাইল নং");
  if (mobCol !== -1) sheet.getRange(start, mobCol + 1, rows.length, 1).setNumberFormat("@");
  sheet.getRange(start, 1, rows.length, headers.length).setValues(rows);
}

/*******************************************************
 * ইনভেস্ট — data: { token, date, entries: [ {নাম, "মোবাইল নং", পদবি, ঠিকানা, টাকা}, ... ] }
 * একসাথে একাধিক এন্ট্রি; প্রতিটি এন্ট্রি আলাদা সারি ও আলাদা InvestID
 *******************************************************/
function addInvest(data) {
  const perm = checkAgencyPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };

  const entries = (data.entries || []).filter(function (e) {
    return String(e["নাম"] || "").trim() !== "" && finNum(e["টাকা"]) > 0;
  });
  if (entries.length === 0) return { success: false, message: "অন্তত একজনের নাম ও টাকার পরিমাণ দিন" };

  const date = data.date ? new Date(data.date) : new Date();
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sheet = finSheet(INVEST_TAB, INVEST_HEADERS);
    const nextId = makeIdGenerator(sheet, "InvestID", "INV", 4);
    let total = 0;
    const rows = entries.map(function (e) {
      const amt = finNum(e["টাকা"]);
      total += amt;
      return [nextId(), date, String(e["নাম"]).trim(), String(e["মোবাইল নং"] || ""),
              String(e["পদবি"] || ""), String(e["ঠিকানা"] || ""), amt];
    });
    finAppendRows(sheet, INVEST_HEADERS, rows);
    return { success: true, total: total, message: rows.length + " টি ইনভেস্ট এন্ট্রি সংরক্ষিত হয়েছে" };
  } finally {
    lock.releaseLock();
  }
}

function listInvests(data) {
  const perm = checkAgencyPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };
  return { success: true, invests: genericListRows(finSheet(INVEST_TAB, INVEST_HEADERS)) };
}

function updateInvest(data) {
  const perm = checkAgencyPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };
  const ok = genericUpdateRow(finSheet(INVEST_TAB, INVEST_HEADERS), "InvestID", data.investId, data.fields || {});
  return { success: ok, message: ok ? "আপডেট হয়েছে" : "ইনভেস্ট এন্ট্রি পাওয়া যায়নি" };
}

function deleteInvest(data) {
  const perm = checkAgencyPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };
  const ok = genericDeleteRow(finSheet(INVEST_TAB, INVEST_HEADERS), "InvestID", data.investId);
  return { success: ok, message: ok ? "ডিলিট হয়েছে" : "ইনভেস্ট এন্ট্রি পাওয়া যায়নি" };
}

/*******************************************************
 * ক্যাশ — জমা / উত্তোলন
 * data: { token, date, amount, name, mobile }
 *******************************************************/
function finAddCash(data, tab, headers, idPrefix, idHeader, label) {
  const perm = checkAgencyPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };
  const amount = finR2(data.amount);
  if (!(amount > 0)) return { success: false, message: "টাকার পরিমাণ দিন" };
  if (!String(data.name || "").trim()) return { success: false, message: label + "কারীর নাম দিন" };

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sheet = finSheet(tab, headers);
    const id = makeIdGenerator(sheet, idHeader, idPrefix, 4)();
    const date = data.date ? new Date(data.date) : new Date();
    finAppendRows(sheet, headers, [[id, date, amount, String(data.name).trim(), String(data.mobile || "")]]);
    return { success: true, id: id, message: label + " সংরক্ষিত হয়েছে" };
  } finally {
    lock.releaseLock();
  }
}
function addCashDeposit(data) { return finAddCash(data, CASH_DEP_TAB, CASH_DEP_HEADERS, "CD", "DepositID", "জমা"); }
function addCashWithdraw(data) { return finAddCash(data, CASH_WD_TAB, CASH_WD_HEADERS, "CW", "WithdrawID", "উত্তোলন"); }

function finListCash(data, tab, headers, key) {
  const perm = checkAgencyPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };
  const out = { success: true };
  out[key] = genericListRows(finSheet(tab, headers));
  return out;
}
function listCashDeposits(data) { return finListCash(data, CASH_DEP_TAB, CASH_DEP_HEADERS, "deposits"); }
function listCashWithdraws(data) { return finListCash(data, CASH_WD_TAB, CASH_WD_HEADERS, "withdraws"); }

function finUpdateCash(data, tab, headers, idHeader) {
  const perm = checkAgencyPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };
  const ok = genericUpdateRow(finSheet(tab, headers), idHeader, data.id, data.fields || {});
  return { success: ok, message: ok ? "আপডেট হয়েছে" : "রেকর্ড পাওয়া যায়নি" };
}
function finDeleteCash(data, tab, headers, idHeader) {
  const perm = checkAgencyPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };
  const ok = genericDeleteRow(finSheet(tab, headers), idHeader, data.id);
  return { success: ok, message: ok ? "ডিলিট হয়েছে" : "রেকর্ড পাওয়া যায়নি" };
}
function updateCashDeposit(data) { return finUpdateCash(data, CASH_DEP_TAB, CASH_DEP_HEADERS, "DepositID"); }
function deleteCashDeposit(data) { return finDeleteCash(data, CASH_DEP_TAB, CASH_DEP_HEADERS, "DepositID"); }
function updateCashWithdraw(data) { return finUpdateCash(data, CASH_WD_TAB, CASH_WD_HEADERS, "WithdrawID"); }
function deleteCashWithdraw(data) { return finDeleteCash(data, CASH_WD_TAB, CASH_WD_HEADERS, "WithdrawID"); }

/* ক্যাশের সর্বমোট (শুরু থেকে আজ পর্যন্ত): জমা, উত্তোলন, বর্তমান ক্যাশ */
function finCashTotals() {
  let dep = 0, depN = 0, wd = 0, wdN = 0;
  genericListRows(finSheet(CASH_DEP_TAB, CASH_DEP_HEADERS)).forEach(function (r) { dep += finNum(r["টাকা"]); depN++; });
  genericListRows(finSheet(CASH_WD_TAB, CASH_WD_HEADERS)).forEach(function (r) { wd += finNum(r["টাকা"]); wdN++; });
  return { depositCount: depN, depositTotal: finR2(dep), withdrawCount: wdN, withdrawTotal: finR2(wd), cash: finR2(dep - wd) };
}

/* ক্যাশ মেনুর রিপোর্ট — data: { token } */
function cashReport(data) {
  const perm = checkAgencyPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };
  return { success: true, report: finCashTotals() };
}

/*******************************************************
 * ডিপু রিপোর্ট (নতুন) — খাত ভিত্তিক টেবিল + সারাংশ টেবিল
 * data: { token, mode: daily/monthly/yearly/total, date, year, month }
 *
 *  ইনভেস্ট ........... Invest ট্যাব (নির্বাচিত সময়)
 *  পণ্য ক্রয় ........ স্টক ভাউচার, ভাউচার ধরে
 *  পণ্য বিক্রয় ...... অর্ডার ইনভয়েস (অপেক্ষমান বাদ), ইনভয়েস ধরে
 *                      বিক্রয় = Σ পরিশোধযোগ্য মূল্য (কমিশন বাদে) − ডিসকাউন্ট
 *  ব্যয় ............. খরচের ভাউচার, ভাউচার ধরে
 *  অন্যান্য আয় ...... আয় রশিদ, রশিদ ধরে
 *  লাভ/লস ............ (মোট বিক্রি + অন্যান্য আয়) − মোট ব্যয়
 *  বর্তমান ক্যাশ ..... ক্যাশ মেনু (জমা − উত্তোলন), সবসময় আজ পর্যন্ত
 *  স্টক মূল্য ........ বর্তমান স্টকের মোট মূল্য, সবসময় আজ পর্যন্ত
 *******************************************************/
function finGroupBy(rows, keyHeader, build) {
  const seen = {}, out = [];
  rows.forEach(function (r) {
    const k = r[keyHeader];
    if (k === "" || k === null || k === undefined) return;
    if (!seen[k]) { seen[k] = build(r); out.push(seen[k]); }
    else if (seen[k]._add) seen[k]._add(r);
  });
  return out;
}

function depotFinanceReport(data) {
  const perm = checkAgencyPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };

  const masterSS = getMasterSS();
  const mode = data.mode || "total";
  const inPeriod = function (d) { return isInReportPeriod(d, mode, data); };

  // ---- ইনভেস্ট ----
  const invests = genericListRows(finSheet(INVEST_TAB, INVEST_HEADERS))
    .filter(function (r) { return inPeriod(r["তারিখ"]); })
    .map(function (r) { return { date: r["তারিখ"], name: r["নাম"], title: r["পদবি"], amount: finNum(r["টাকা"]) }; });
  let investTotal = 0;
  invests.forEach(function (r) { investTotal += r.amount; });

  // ---- পণ্য ক্রয় (স্টক ভাউচার) ----
  const allVouchers = finGroupBy(genericListRows(getSheet(masterSS, "StockInVoucher")), "ভাউচার নং", function (r) {
    const o = { voucherNo: r["ভাউচার নং"], date: r["তারিখ"], total: finNum(r["মোট মূল্য"]), paid: finNum(r["পরিশোধ"]), due: finNum(r["বকেয়া"]) };
    o._add = function (x) { o.total += finNum(x["মোট মূল্য"]); };
    return o;
  });
  const vouchers = allVouchers.filter(function (v) { return inPeriod(v.date); });
  let purchaseTotal = 0, purchasePaid = 0, purchaseDue = 0;
  vouchers.forEach(function (v) { v.total = finR2(v.total); purchaseTotal += v.total; purchasePaid += v.paid; purchaseDue += v.due; });

  // ---- পণ্য বিক্রয় (অর্ডার ইনভয়েস, অপেক্ষমান বাদ) ----
  const invoiceLines = genericListRows(getSheet(masterSS, "SalesInvoice"))
    .filter(function (e) { return e["স্ট্যাটাস"] !== "অপেক্ষমান"; });
  const allOrders = finGroupBy(invoiceLines, "ইনভয়েস নং", function (r) {
    const o = { invoiceNo: r["ইনভয়েস নং"], date: r["তারিখ"], dealer: r["ডিলার নাম"], status: r["স্ট্যাটাস"] || "পেন্ডিং",
                gross: finNum(r["মোট মূল্য"]), commission: finNum(r["কমিশন মূল্য"]), payable: finNum(r["পরিশোধযোগ্য মূল্য"]),
                discount: finNum(r["ডিসকাউন্ট"]), paid: finNum(r["পরিশোধ"]), due: finNum(r["বকেয়া"]) };
    o._add = function (x) { o.gross += finNum(x["মোট মূল্য"]); o.commission += finNum(x["কমিশন মূল্য"]); o.payable += finNum(x["পরিশোধযোগ্য মূল্য"]); };
    return o;
  });
  const orders = allOrders.filter(function (o) { return inPeriod(o.date); });
  let salesTotal = 0, salesPaid = 0, salesDue = 0;
  orders.forEach(function (o) {
    o.net = finR2(o.payable - o.discount);
    salesTotal += o.net; salesPaid += o.paid; salesDue += o.due;
  });

  // ---- ব্যয় (খরচের ভাউচার) ----
  const expenses = finGroupBy(genericListRows(getSheet(masterSS, "Expense")), "ভাউচার নং", function (r) {
    const o = { voucherNo: r["ভাউচার নং"], date: r["তারিখ"], total: finNum(r["টাকা"]), paid: finNum(r["পরিশোধ"]), due: finNum(r["বকেয়া"]) };
    o._add = function (x) { o.total += finNum(x["টাকা"]); };
    return o;
  }).filter(function (v) { return inPeriod(v.date); });
  let expenseTotal = 0, expensePaid = 0, expenseDue = 0;
  expenses.forEach(function (v) { v.total = finR2(v.total); expenseTotal += v.total; expensePaid += v.paid; expenseDue += v.due; });

  // ---- অন্যান্য আয় (আয় রশিদ) ----
  const incSheet = masterSS.getSheetByName("IncomeReceipt");
  const receipts = finGroupBy(incSheet ? genericListRows(incSheet) : [], "রশিদ নং", function (r) {
    const o = { receiptNo: r["রশিদ নং"], date: r["তারিখ"], total: finNum(r["টাকা"]) };
    o._add = function (x) { o.total += finNum(x["টাকা"]); };
    return o;
  }).filter(function (v) { return inPeriod(v.date); });
  let otherIncomeTotal = 0;
  receipts.forEach(function (v) { v.total = finR2(v.total); otherIncomeTotal += v.total; });

  // ---- লাভ/লস = (মোট বিক্রি + অন্যান্য আয়) − মোট ব্যয় ----
  const sellPlusIncome = salesTotal + otherIncomeTotal;
  const net = sellPlusIncome - expenseTotal;

  // ---- ক্যাশ ও স্টক মূল্য (আজ পর্যন্ত) ----
  const cash = finCashTotals();
  let stockValue = 0;
  computeStockState(masterSS).list.forEach(function (r) {
    const qty = r.qtyIn - r.sold;
    const avg = r.qtyIn > 0 ? r.valueIn / r.qtyIn : 0;
    stockValue += Math.round(qty * avg);
  });

  function clean(list) { list.forEach(function (x) { delete x._add; }); return list; }

  // ---- কোম্পানীর বকেয়া (স্টক ভাউচারের বকেয়া) ও ডিলারদের নিকট বকেয়া (অর্ডার ইনভয়েসের বকেয়া) — আজ পর্যন্ত মোট বকেয়া ----
  const companyDueRows = [];
  allVouchers.forEach(function (v) {
    if (v.due > 0.0001) companyDueRows.push({ voucherNo: v.voucherNo, date: v.date, total: finR2(v.total), paid: v.paid, due: v.due });
  });
  const dealerDueRows = [];
  allOrders.forEach(function (o) {
    if (o.due > 0.0001) dealerDueRows.push({ invoiceNo: o.invoiceNo, date: o.date, dealer: o.dealer, net: finR2(o.payable - o.discount), paid: o.paid, due: o.due });
  });
  let companyDueTotal = 0, companyDueAmt = 0, companyPaidAmt = 0, dealerDueTotal = 0, dealerPaidAmt = 0, dealerDueAmt = 0;
  companyDueRows.forEach(function (x) { companyDueTotal += x.total; companyPaidAmt += x.paid; companyDueAmt += x.due; });
  dealerDueRows.forEach(function (x) { dealerDueTotal += x.net; dealerPaidAmt += x.paid; dealerDueAmt += x.due; });

  return {
    success: true,
    report: {
      invest: { rows: invests, total: finR2(investTotal) },
      purchase: { rows: clean(vouchers), total: finR2(purchaseTotal), paid: finR2(purchasePaid), due: finR2(purchaseDue) },
      sales: { rows: clean(orders), total: finR2(salesTotal), paid: finR2(salesPaid), due: finR2(salesDue) },
      expense: { rows: clean(expenses), total: finR2(expenseTotal), paid: finR2(expensePaid), due: finR2(expenseDue) },
      otherIncome: { rows: clean(receipts), total: finR2(otherIncomeTotal) },
      profitLoss: { sales: finR2(salesTotal), otherIncome: finR2(otherIncomeTotal), income: finR2(sellPlusIncome), expense: finR2(expenseTotal), net: finR2(net) },
      companyDue: { rows: companyDueRows, total: finR2(companyDueTotal), paid: finR2(companyPaidAmt), due: finR2(companyDueAmt) },
      dealerDue: { rows: dealerDueRows, total: finR2(dealerDueTotal), paid: finR2(dealerPaidAmt), due: finR2(dealerDueAmt) },
      cash: cash,
      stockValue: finR2(stockValue)
    }
  };
}
