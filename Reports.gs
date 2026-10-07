/*******************************************************
 * TCB-স্টাইল ডিলার ম্যানেজমেন্ট সফটওয়্যার
 * Reports.gs — দৈনিক, মাসিক, সর্বমোট রিপোর্ট
 *
 * MasterRegistry এর একই Apps Script প্রজেক্টে নতুন ফাইল
 * হিসেবে যোগ করুন (নাম দিন: Reports)
 *******************************************************/

/*******************************************************
 * তারিখ কে yyyy-MM-dd স্ট্রিং এ রূপান্তর (স্ক্রিপ্টের টাইমজোন অনুযায়ী)
 *******************************************************/
function toDateStr(value) {
  const d = (value instanceof Date) ? value : new Date(value);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

/*******************************************************
 * মূল রিপোর্ট বিল্ডার — বিক্রিত সেলস ফিল্টার করে প্যাকেজ-ভিত্তিক গ্রুপ করে
 * (প্যাকেজ তথ্য এখন Master Spreadsheet থেকে আসে, Sales ডিলারের নিজের শীট থেকে)
 *******************************************************/
function buildSalesReport(ss, filterFn) {
  const salesSheet = getSheet(ss, "Sales");
  const masterSS = getMasterSS();
  const pkgSheet = getSheet(masterSS, "Packages");

  const allSales = genericListRows(salesSheet).filter(function (s) {
    return s["স্ট্যাটাস"] === "বিক্রিত";
  });
  const packages = genericListRows(pkgSheet);
  const pkgMap = {};
  packages.forEach(function (p) { pkgMap[p["PackageID"]] = p; });

  const filtered = allSales.filter(filterFn);

  const byPackage = {};
  let grandTotal = 0;

  filtered.forEach(function (s) {
    const pkg = pkgMap[s["PackageID"]];
    // বিক্রির সময় নির্ধারিত মূল্যই আসল হিসাব — এটা প্যাকেজের ডিফল্ট
    // মূল্য থেকে ম্যানুয়ালি বদলানো থাকতে পারে; না থাকলে (পুরনো
    // রেকর্ড) প্যাকেজের বর্তমান মূল্য দিয়ে হিসাব করা হয়
    const price = (s["মূল্য"] !== undefined && s["মূল্য"] !== "" && s["মূল্য"] !== null)
      ? Number(s["মূল্য"]) || 0
      : (pkg ? (Number(pkg["সর্বমোট কম্বো মূল্য (সর্বমোট মূল্য)"]) || 0) : 0);
    grandTotal += price;

    const key = s["PackageID"];
    if (!byPackage[key]) {
      byPackage[key] = {
        packageId: key,
        packageName: pkg ? pkg["নাম"] : "(অজানা প্যাকেজ)",
        count: 0,
        total: 0
      };
    }
    byPackage[key].count += 1;
    byPackage[key].total += price;
  });

  return {
    totalCount: filtered.length,
    grandTotal: grandTotal,
    byPackage: Object.keys(byPackage).map(function (k) { return byPackage[k]; })
  };
}

/*******************************************************
 * দৈনিক রিপোর্ট — data: { token, date: "yyyy-MM-dd" }
 *******************************************************/
function dailyReport(data) {
  const perm = checkPermission(data.token, ["Admin", "প্রতিনিধি"]);
  if (!perm.ok) return { success: false, message: perm.message };

  const ss = getDealerSpreadsheet(perm.payload.dealerId);
  const targetDate = data.date; // "yyyy-MM-dd"

  const report = buildSalesReport(ss, function (s) {
    return toDateStr(s["তারিখ"]) === targetDate;
  });

  return { success: true, date: targetDate, report: report };
}

/*******************************************************
 * মাসিক রিপোর্ট — data: { token, year: 2026, month: 9 }
 *******************************************************/
function monthlyReport(data) {
  const perm = checkPermission(data.token, ["Admin", "প্রতিনিধি"]);
  if (!perm.ok) return { success: false, message: perm.message };

  const ss = getDealerSpreadsheet(perm.payload.dealerId);
  const year = Number(data.year);
  const month = Number(data.month); // ১-১২

  const report = buildSalesReport(ss, function (s) {
    const d = (s["তারিখ"] instanceof Date) ? s["তারিখ"] : new Date(s["তারিখ"]);
    return d.getFullYear() === year && (d.getMonth() + 1) === month;
  });

  return { success: true, year: year, month: month, report: report };
}

/*******************************************************
 * সর্বমোট রিপোর্ট — data: { token }
 *******************************************************/
function totalReport(data) {
  const perm = checkPermission(data.token, ["Admin", "প্রতিনিধি"]);
  if (!perm.ok) return { success: false, message: perm.message };

  const ss = getDealerSpreadsheet(perm.payload.dealerId);
  const report = buildSalesReport(ss, function () { return true; });

  return { success: true, report: report };
}

/*******************************************************
 * ডিলার রিপোর্ট — শুধু নিজের ডিলারের তথ্য (ডিপু/এজেন্সির রিপোর্ট নয়)
 * data: { token, mode: "daily"/"monthly"/"total", date, year, month }
 *
 *  ১) short     — শর্ট রিপোর্ট টেবিল
 *  ২) products  — পণ্য (প্যাকেজ) ভিত্তিক মোট ক্রয়-বিক্রয় তথ্য
 *  ৩) stock     — পণ্য ভিত্তিক মোট স্টক তথ্য (আজ পর্যন্ত)
 *  ৪) expenses  — খাত ভিত্তিক খরচ তথ্য
 *
 * লাভের হিসাব (প্রতিটি বিক্রিত প্যাকেজে):
 *   লাভ = ক্রয় কম্বো মূল্যের ৪% + (বিক্রয় মূল্য − ক্রয় কম্বো মূল্য)
 *   যেমন ক্রয় ৮৫৫, বিক্রি ৮৫৫ → ৩৪.২ ; বিক্রি ৮৭০ → ৪৯.২ ;
 *   ক্রয়ের চেয়ে কমে বিক্রি হলে একই সূত্রে কম/ঋণাত্মক হবে
 * মোট আয় = প্যাকেজ বিক্রি থেকে আয় (উপরের সূত্র) + আয় রশিদ থেকে আয় + কার্ড ফি থেকে আয়
 * বর্তমান লাভ/লস = মোট আয় − মোট খরচ
 * ১,২,৪ নির্বাচিত সময়ের (দৈনিক/মাসিক/সর্বমোট); স্টক ও ক্যাশ সবসময় আজ পর্যন্ত
 *******************************************************/
const DEALER_PROFIT_PERCENT = 4;

function dealerReport(data) {
  const perm = checkPermission(data.token, ["Admin", "প্রতিনিধি"]);
  if (!perm.ok) return { success: false, message: perm.message };
  const dealerId = perm.payload.dealerId;
  if (dealerId === AGENCY_ID) return { success: false, message: "এই রিপোর্ট শুধু ডিলার সাইটের জন্য" };

  const mode = data.mode || "total";
  const ss = getDealerSpreadsheet(dealerId);
  const masterSS = getMasterSS();
  const n = function (v) { return Number(v) || 0; };
  const r2 = function (v) { return Math.round(v * 100) / 100; };

  // ---- প্যাকেজ তথ্য ----
  const pkgMap = {};
  genericListRows(getPackagesSheet(masterSS)).forEach(function (p) {
    pkgMap[p["PackageID"]] = {
      serial: packageNoOf(p),
      name: p["নাম"],
      bazar: n(p["সর্বমোট বাজার মূল্য"]),
      combo: n(p["সর্বমোট কম্বো মূল্য (সর্বমোট মূল্য)"])
    };
  });
  function pkgSerial(id) { return pkgMap[id] ? pkgMap[id].serial : 999999; }
  function pkgName(id) { return pkgMap[id] ? pkgMap[id].name : "(অজানা প্যাকেজ)"; }
  function pkgBazar(id) { return pkgMap[id] ? pkgMap[id].bazar : 0; }

  // ---- অর্ডার (এই ডিলারের) ----
  // confirmedLines: কনফার্ম হওয়া সব অর্ডার ("অপেক্ষমান" বাদ) — ডিপুকে পরিশোধ হিসাবের জন্য
  // orderLines: শুধু "ডেলিভারি সম্পন্ন" অর্ডার — ক্রয়/স্টক এ কেবল ডেলিভারি পাওয়া পণ্যই ধরা হয়
  const confirmedLines = genericListRows(getSheet(masterSS, "SalesInvoice")).filter(function (e) {
    return e["DealerID"] === dealerId && e["স্ট্যাটাস"] !== "অপেক্ষমান";
  });
  const orderLines = confirmedLines.filter(function (e) { return e["স্ট্যাটাস"] === "ডেলিভারি সম্পন্ন"; });

  // প্যাকেজ-ভিত্তিক সর্বমোট ক্রয় (সংখ্যা ও কম্বো মূল্য) — গড় ক্রয় কম্বো মূল্য বের করার জন্য
  const bought = {};            // সব সময়
  const invPaid = {};           // ইনভয়েস ধরে পরিশোধ (প্রতি ইনভয়েসে একবার)
  orderLines.forEach(function (e) {
    const id = e["PackageID"];
    if (!bought[id]) bought[id] = { qty: 0, combo: 0 };
    bought[id].qty += n(e["সংখ্যা"]);
    bought[id].combo += n(e["মোট মূল্য"]);
  });
  confirmedLines.forEach(function (e) {
    const k = e["ইনভয়েস নং"];
    if (invPaid[k] === undefined) invPaid[k] = n(e["পরিশোধ"]);
  });
  function unitCost(id) {
    if (bought[id] && bought[id].qty > 0) return bought[id].combo / bought[id].qty;
    return pkgMap[id] ? pkgMap[id].combo : 0;
  }
  let paidToDepot = 0;
  Object.keys(invPaid).forEach(function (k) { paidToDepot += invPaid[k]; });

  // ---- বিক্রি (গ্রাহকের কাছে) ----
  const allSales = genericListRows(getSheet(ss, "Sales")).filter(function (s) { return s["স্ট্যাটাস"] === "বিক্রিত"; });
  function salePrice(s) {
    if (s["মূল্য"] !== undefined && s["মূল্য"] !== "" && s["মূল্য"] !== null) return n(s["মূল্য"]);
    return pkgMap[s["PackageID"]] ? pkgMap[s["PackageID"]].combo : 0;
  }
  const soldAllQty = {};        // প্যাকেজ → সর্বমোট বিক্রি সংখ্যা (স্টকের জন্য)
  let allSalesTotal = 0;
  allSales.forEach(function (s) {
    const id = s["PackageID"];
    soldAllQty[id] = (soldAllQty[id] || 0) + 1;
    allSalesTotal += salePrice(s);
  });

  // ---- নির্বাচিত সময়ের পণ্যভিত্তিক ক্রয়-বিক্রয় ----
  const prod = {};
  function ensureProd(id) {
    if (!prod[id]) prod[id] = { serial: pkgSerial(id), name: pkgName(id), qty: 0, bazar: 0, combo: 0, soldQty: 0, soldAmount: 0, unitProfit: 0, profit: 0 };
    return prod[id];
  }
  orderLines.filter(function (e) { return isInReportPeriod(e["তারিখ"], mode, data); }).forEach(function (e) {
    const id = e["PackageID"];
    const row = ensureProd(id);
    const qty = n(e["সংখ্যা"]);
    row.qty += qty;
    row.bazar += pkgBazar(id) * qty;
    row.combo += n(e["মোট মূল্য"]);
  });

  let totalProfit = 0;
  allSales.filter(function (s) { return isInReportPeriod(s["তারিখ"], mode, data); }).forEach(function (s) {
    const id = s["PackageID"];
    const price = salePrice(s);
    const cost = unitCost(id);
    const row = ensureProd(id);
    row.soldQty += 1;
    row.soldAmount += price;
    // প্রতি প্যাকেটের আয় = ক্রয় কম্বো মূল্যের ৪% + (বিক্রয় মূল্য − ক্রয় কম্বো মূল্য)
    // (সমান দামে বিক্রি → শুধু ৪%; বেশিতে বিক্রি → বাড়ে; কমে বিক্রি → কমে)
    const saleProfit = cost * DEALER_PROFIT_PERCENT / 100 + (price - cost);
    row.profit += saleProfit;
    totalProfit += saleProfit;
  });
  // প্রতিটি প্যাকেজের "প্রতি প্যাকেটের আয়" — কম্বো মূল্যে বিক্রি হলে যা হতো (ক্রয় কম্বো মূল্যের ৪%)
  Object.keys(prod).forEach(function (id) {
    prod[id].unitProfit = r2(unitCost(id) * DEALER_PROFIT_PERCENT / 100);
    prod[id].profit = r2(prod[id].profit);
  });

  const productRows = Object.keys(prod).map(function (k) { return prod[k]; })
    .sort(function (a, b) { return a.serial - b.serial; });
  const productTotals = { qty: 0, bazar: 0, combo: 0, soldQty: 0, soldAmount: 0, profit: 0 };
  productRows.forEach(function (r) {
    productTotals.qty += r.qty; productTotals.bazar += r.bazar; productTotals.combo += r.combo;
    productTotals.soldQty += r.soldQty; productTotals.soldAmount += r.soldAmount; productTotals.profit += r.profit;
  });
  productTotals.profit = r2(productTotals.profit);
  const purchasedTypes = productRows.filter(function (r) { return r.qty > 0; }).length;

  // ---- স্টক (আজ পর্যন্ত: মোট ক্রয় − মোট বিক্রি) ----
  const stockRows = [];
  const stockTotals = { qty: 0, bazar: 0, combo: 0 };
  Object.keys(bought).forEach(function (id) {
    const qty = Math.max(0, bought[id].qty - (soldAllQty[id] || 0));
    const bazar = pkgBazar(id) * qty;
    const combo = unitCost(id) * qty;
    stockRows.push({ serial: pkgSerial(id), name: pkgName(id), qty: qty, bazar: r2(bazar), combo: r2(combo) });
    stockTotals.qty += qty; stockTotals.bazar += bazar; stockTotals.combo += combo;
  });
  stockRows.sort(function (a, b) { return a.serial - b.serial; });
  stockTotals.bazar = r2(stockTotals.bazar); stockTotals.combo = r2(stockTotals.combo);

  // ---- খরচ (খাত = বিবরণ ভিত্তিক) ----
  const allExpenses = genericListRows(getSheet(ss, "ExpenseVoucher"));
  let allExpenseTotal = 0;
  allExpenses.forEach(function (v) { allExpenseTotal += n(v["পরিমাণ"]); });

  const expBy = {};
  allExpenses.filter(function (v) { return isInReportPeriod(v["তারিখ"], mode, data); }).forEach(function (v) {
    const cat = String(v["বিবরণ"] || "").trim() || "(খাত উল্লেখ নেই)";
    expBy[cat] = (expBy[cat] || 0) + n(v["পরিমাণ"]);
  });
  const expenseRows = Object.keys(expBy).map(function (k) { return { category: k, amount: expBy[k] }; });
  let expenseTotal = 0;
  expenseRows.forEach(function (r) { expenseTotal += r.amount; });

  // ---- আয় রশিদ থেকে আয় (ডিলারের নিজের শীটের IncomeReceipt ট্যাব) ----
  const incSheet = ss.getSheetByName("IncomeReceipt");
  const incomeLines = incSheet ? genericListRows(incSheet) : [];
  let allReceiptIncome = 0, receiptIncome = 0;
  incomeLines.forEach(function (l) {
    const amt = n(l["টাকা"]);
    allReceiptIncome += amt;
    if (isInReportPeriod(l["তারিখ"], mode, data)) receiptIncome += amt;
  });

  // ---- কার্ড ফি থেকে আয় (গ্রাহক ফর্মের "কার্ড ফি"; গ্রাহকের যোগ করার তারিখ ধরে সময় নির্ধারণ) ----
  let allCardFee = 0, cardFeeIncome = 0;
  genericListRows(getSheet(ss, "Customers")).forEach(function (c) {
    const fee = n(c["কার্ড ফি"]);
    if (!fee) return;
    allCardFee += fee;
    if (isInReportPeriod(c["তারিখ"], mode, data)) cardFeeIncome += fee;
  });

  // ---- মোট আয় = প্যাকেজ বিক্রি থেকে আয় + আয় রশিদ থেকে আয় + কার্ড ফি থেকে আয় ----
  const totalIncome = totalProfit + receiptIncome + cardFeeIncome;
  const net = totalIncome - expenseTotal;       // বর্তমান লাভ/লস = মোট আয় − মোট খরচ

  return {
    success: true,
    mode: mode,
    report: {
      short: {
        purchaseTypes: purchasedTypes,
        purchaseQty: productTotals.qty,
        purchaseAmount: r2(productTotals.combo),
        soldQty: productTotals.soldQty,
        stockQty: stockTotals.qty,
        expense: r2(expenseTotal),
        packageProfit: r2(totalProfit),     // প্যাকেজ বিক্রি থেকে আয়
        receiptIncome: r2(receiptIncome),   // আয় রশিদ থেকে আয়
        cardFeeIncome: r2(cardFeeIncome),   // কার্ড ফি থেকে আয়
        totalIncome: r2(totalIncome),       // মোট আয়
        net: r2(net),                       // ধনাত্মক = লাভ, ঋণাত্মক = লস
        // ক্যাশ (আজ পর্যন্ত): বিক্রির টাকা + আয় রশিদের টাকা + কার্ড ফি − ডিপুকে পরিশোধ − খরচ
        cash: r2(allSalesTotal + allReceiptIncome + allCardFee - paidToDepot - allExpenseTotal)
      },
      products: { rows: productRows, totals: productTotals },
      stock: { rows: stockRows, totals: stockTotals },
      expenses: { rows: expenseRows, total: expenseTotal }
    }
  };
}
