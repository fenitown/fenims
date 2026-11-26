/*******************************************************
 * TCB/fenims ডিলার ম্যানেজমেন্ট সফটওয়্যার
 * Stock.gs — স্টক (পণ্য তালিকা, স্টক ইন ভাউচার), বিক্রয় ইনভয়েস
 * (এজেন্সি থেকে ডিলারকে প্যাকেজ বিক্রি), এবং এজেন্সি রিপোর্ট
 *
 * MasterRegistry এর একই Apps Script প্রজেক্টে নতুন ফাইল হিসেবে
 * যোগ করুন (নাম দিন: Stock)
 * সব একশন শুধুমাত্র মূল এজেন্সি (AGENCY) লগইন করতে পারবে
 *******************************************************/

/*=========================================================
 *  পণ্য তালিকা (Products)
 *=======================================================*/
function addProduct(data) {
  const perm = checkAgencyPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };

  const masterSS = getMasterSS();
  const sheet = getSheet(masterSS, "Products");
  ensureSheetColumns(sheet, ["পরিমাণ", "একক", "প্যাকিং ধরন", "প্যাকিং পরিমাণ"]);   // পুরনো শীটে না থাকলে নিজে যোগ হয়
  const productId = generateId(sheet, "PR");

  genericAddRow(sheet, {
    "ProductID": productId,
    "পণ্যের নাম": data.নাম,
    "ব্র্যান্ড": data.brand,
    "পরিমাণ": data.poriman || "",
    "একক": data.unit || "",
    "প্যাকিং ধরন": data.packType || "",
    "প্যাকিং পরিমাণ": Number(data.packSize) || ""
  });

  invalidateAgencyCaches();
  return { success: true, productId: productId, message: "পণ্য যোগ হয়েছে" };
}

function listProducts(data) {
  const perm = checkAgencyPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };

  const cache = CacheService.getScriptCache();
  const cached = cache.get("products_cache_v2");
  if (cached) return { success: true, products: JSON.parse(cached) };

  const masterSS = getMasterSS();
  const sheet = getSheet(masterSS, "Products");
  const products = genericListRows(sheet);

  try { cache.put("products_cache_v2", JSON.stringify(products), 30); } catch (e) { /* বাদ */ }
  return { success: true, products: products };
}

function updateProduct(data) {
  const perm = checkAgencyPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };

  const masterSS = getMasterSS();
  const sheet = getSheet(masterSS, "Products");
  ensureSheetColumns(sheet, ["পরিমাণ", "একক", "প্যাকিং ধরন", "প্যাকিং পরিমাণ"]);
  const ok = genericUpdateRow(sheet, "ProductID", data.productId, data.fields || {});
  invalidateProductCache();
  return { success: ok, message: ok ? "পণ্য আপডেট হয়েছে" : "পণ্য পাওয়া যায়নি" };
}

function deleteProduct(data) {
  const perm = checkAgencyPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };

  const masterSS = getMasterSS();
  const sheet = getSheet(masterSS, "Products");
  const ok = genericDeleteRow(sheet, "ProductID", data.productId);
  invalidateProductCache();
  return { success: ok, message: ok ? "পণ্য ডিলিট হয়েছে" : "পণ্য পাওয়া যায়নি" };
}

function invalidateProductCache() {
  CacheService.getScriptCache().remove("products_cache_v2");
}

/*=========================================================
 *  মূল্য তালিকা (PriceList) — পণ্যের বাজার মূল্য ও কম্বো মূল্য নির্ধারণ
 *  একটি পণ্যের একটি সারি: "পরিমাণ + একক" এর জন্য মূল্য (যেমন চাউল, ১ কেজি, কম্বো ৫০)।
 *  প্যাকেজ ফরম ও অর্ডার ইনভয়েসে পণ্য/পরিমাণ বাছলে দাম = মূল্য × (পরিমাণ ÷ তালিকার পরিমাণ)
 *=======================================================*/
function getPriceListSheet(masterSS) {
  let sheet = masterSS.getSheetByName("PriceList");
  if (!sheet) {
    sheet = masterSS.insertSheet("PriceList");
    const h = MASTER_SHEETS_DEF["PriceList"];
    sheet.getRange(1, 1, 1, h.length).setValues([h]);
    sheet.getRange(1, 1, 1, h.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/* মূল্য তালিকার মূল্য-অনুপাত: "qty unit" কে তালিকার "baseQty baseUnit" দিয়ে ভাগ (কেজি↔গ্রাম, লিটার↔মি.লি রূপান্তরসহ) */
function priceRatio(baseQty, baseUnit, qty, unit) {
  const a = parseSizeText(stockToEn(qty) + " " + (unit || ""));
  const b = parseSizeText(stockToEn(baseQty) + " " + (baseUnit || ""));
  if (a && b && a.u === b.u && b.v > 0) return a.v / b.v;
  const bq = Number(stockToEn(baseQty)) || 0, q = Number(stockToEn(qty)) || 0;
  return bq > 0 ? q / bq : 1;
}

function priceListMap(masterSS) {
  const map = {};
  genericListRows(getPriceListSheet(masterSS)).forEach(function (r) {
    map[r["ProductID"]] = {
      qty: String(r["পরিমাণ"] == null ? "" : r["পরিমাণ"]), unit: String(r["একক"] || ""),
      market: Number(r["বাজার মূল্য"]) || 0, combo: Number(r["কম্বো মূল্য"]) || 0, brand: String(r["ব্র্যান্ড"] || "")
    };
  });
  return map;
}

function listPriceList(data) {
  const perm = checkAgencyPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };
  return { success: true, items: genericListRows(getPriceListSheet(getMasterSS())) };
}

/* data: { token, productId, qty, unit, market, combo } — পণ্য আগে থেকে তালিকায় থাকলে আপডেট, না থাকলে নতুন */
function savePriceItem(data) {
  const perm = checkAgencyPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };
  const masterSS = getMasterSS();
  const prod = genericListRows(getSheet(masterSS, "Products")).filter(function (p) { return p["ProductID"] === data.productId; })[0];
  if (!prod) return { success: false, message: "পণ্য পাওয়া যায়নি" };
  const qty = String(data.qty == null ? "" : data.qty).trim();
  if (!(Number(stockToEn(qty)) > 0)) return { success: false, message: "পরিমাণ দিন (০ এর বেশি)" };
  const combo = Number(data.combo) || 0, market = Number(data.market) || 0;
  if (combo <= 0) return { success: false, message: "কম্বো মূল্য দিন" };

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sheet = getPriceListSheet(masterSS);
    ensureSheetColumns(sheet, ["ব্র্যান্ড"]);   // আগের শীটে না থাকলে নিজে যোগ হয়
    const fields = { "ব্র্যান্ড": String(prod["ব্র্যান্ড"] || ""), "পণ্যের নাম": prod["পণ্যের নাম"], "পরিমাণ": qty, "একক": String(data.unit || ""), "বাজার মূল্য": market, "কম্বো মূল্য": combo };
    const exists = genericListRows(sheet).some(function (r) { return r["ProductID"] === data.productId; });
    if (exists) {
      genericUpdateRow(sheet, "ProductID", data.productId, fields);
    } else {
      fields["PriceID"] = generateId(sheet, "PL");
      fields["ProductID"] = data.productId;
      genericAddRow(sheet, fields);
    }
    // আগের ইনভয়েসের দাম ধরে রেখে (freeze), তারপর প্যাকেজের দাম নতুন মূল্যে হালনাগাদ
    freezeLegacyInvoices(masterSS);
    const n = propagatePriceToPackages(masterSS, data.productId, { qty: qty, unit: String(data.unit || ""), market: market, combo: combo });
    return { success: true, message: "মূল্য তালিকা সংরক্ষিত হয়েছে" + (n ? " — " + n + " টি প্যাকেজের দাম হালনাগাদ হয়েছে" : "") };
  } finally {
    lock.releaseLock();
  }
}

function deletePriceItem(data) {
  const perm = checkAgencyPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };
  const ok = genericDeleteRow(getPriceListSheet(getMasterSS()), "ProductID", data.productId);
  return { success: ok, message: ok ? "মূল্য তালিকা থেকে বাদ দেওয়া হয়েছে" : "পাওয়া যায়নি" };
}

/*=========================================================
 *  স্টক ইন ভাউচার (StockInVoucher) — এক ভাউচারে একাধিক পণ্য
 *=======================================================*/
function addStockVoucher(data) {
  const perm = checkAgencyPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };

  const masterSS = getMasterSS();
  const sheet = getSheet(masterSS, "StockInVoucher");
  ensureSheetColumns(sheet, ["ব্র্যান্ড", "ক্রয় মূল্য"]);   // পুরনো শীটে না থাকলে নিজে যোগ হয়
  const voucherNo = "SV" + Utilities.formatString("%06d", Math.floor(Math.random() * 900000) + 100000);
  const now = data.date ? new Date(data.date) : new Date();

  const items = data.items || [];
  let grandTotal = 0;
  const lineData = items.map(function (item) {
    const purchasePrice = Number(item.purchasePrice) || 0;
    const total = (Number(item.quantity) || 0) * purchasePrice;
    grandTotal += total;
    return { item: item, purchasePrice: purchasePrice, total: total };
  });
  const paid = Number(data.paid) || 0;
  const due = grandTotal - paid;

  const rows = lineData.map(function (ld) {
    return {
      "ভাউচার নং": voucherNo,
      "তারিখ": now,
      "ProductID": ld.item.productId,
      "পণ্যের নাম": ld.item.productName || "",
      "ব্র্যান্ড": ld.item.brand || "",
      "ক্রয় মূল্য": ld.purchasePrice,
      "সংখ্যা": ld.item.quantity,
      "মোট মূল্য": ld.total,
      "সর্বমোট": grandTotal,
      "পরিশোধ": paid,
      "বকেয়া": due
    };
  });

  batchAppendRows(sheet, rows, "SVE", "EntryID");

  return { success: true, voucherNo: voucherNo, total: grandTotal, paid: paid, due: due, message: "স্টক ভাউচার যোগ হয়েছে" };
}

function listStockVouchers(data) {
  const perm = checkAgencyPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };

  const masterSS = getMasterSS();
  const sheet = getSheet(masterSS, "StockInVoucher");
  return { success: true, entries: genericListRows(sheet) };
}

function updateStockVoucherEntry(data) {
  const perm = checkAgencyPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };

  const masterSS = getMasterSS();
  const sheet = getSheet(masterSS, "StockInVoucher");
  ensureSheetColumns(sheet, ["ব্র্যান্ড", "ক্রয় মূল্য"]);
  const ok = genericUpdateRow(sheet, "EntryID", data.entryId, data.fields || {});
  return { success: ok, message: ok ? "আপডেট হয়েছে" : "এন্ট্রি পাওয়া যায়নি" };
}

/*******************************************************
 * পুরো ভাউচার (তার সব পণ্য-লাইন) ডিলিট করা
 *******************************************************/
function deleteStockVoucher(data) {
  const perm = checkAgencyPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };

  const masterSS = getMasterSS();
  const sheet = getSheet(masterSS, "StockInVoucher");
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const idx = headers.indexOf("ভাউচার নং");

  for (let i = values.length - 1; i >= 1; i--) {
    if (values[i][idx] === data.voucherNo) {
      sheet.deleteRow(i + 1);
    }
  }
  return { success: true, message: "ভাউচার ডিলিট হয়েছে" };
}

/*=========================================================
 *  বিক্রয় ইনভয়েস (SalesInvoice) — এজেন্সি থেকে ডিলারকে বিক্রি
 *=======================================================*/
/*******************************************************
 * বিক্রয় ইনভয়েসের সারি তৈরি করার শেয়ার্ড লজিক — addSalesInvoice ও
 * updateSalesInvoiceFull দুটোতেই ব্যবহৃত হয়
 *******************************************************/
function buildSalesInvoiceRows(invoiceNo, dealerInfo, dealerMobile, dealerThikana, date, items, discount, paid, status, seller) {
  let subtotal = 0, totalPayable = 0;
  const lineCalc = items.map(function (item) {
    const total = (Number(item.quantity) || 0) * (Number(item.unitPrice) || 0);
    const commissionPct = Number(item.commissionPercent) || 0;
    const commissionAmt = total * commissionPct / 100;
    subtotal += total;
    totalPayable += (total - commissionAmt);
    return { item: item, total: total, commissionPct: commissionPct, commissionAmt: commissionAmt };
  });
  const netPayable = totalPayable - discount;
  const due = netPayable - paid;

  const rows = lineCalc.map(function (lc) {
    return {
      "ইনভয়েস নং": invoiceNo,
      "তারিখ": date,
      "DealerID": dealerInfo.dealerId,
      "ডিলার নাম": dealerInfo.name,
      "মোবাইল": dealerMobile,
      "ঠিকানা": dealerThikana,
      "PackageID": lc.item.packageId,
      "প্যাকেজ": lc.item.packageName || "",
      "একক মূল্য": lc.item.unitPrice,
      "সংখ্যা": lc.item.quantity,
      "মোট মূল্য": lc.total,
      "কমিশন %": lc.commissionPct,
      "কমিশন মূল্য": lc.commissionAmt,
      "পরিশোধযোগ্য মূল্য": lc.total - lc.commissionAmt,
      "সাবটোটাল": subtotal,
      "ডিসকাউন্ট": discount,
      "পরিশোধ": paid,
      "বকেয়া": due,
      "স্ট্যাটাস": status,
      "বিক্রয়কারী": seller || "",
      "পণ্য পরিবর্তন": lc.item.customJson || "",
      "লাইন সংরক্ষণ": lc.item.snapshotJson || ""
    };
  });

  return { rows: rows, subtotal: subtotal, netPayable: netPayable, due: due };
}

function getDealerContactInfo(dealersSheet, dealerId) {
  const dealersData = dealersSheet.getDataRange().getValues();
  const dHeaders = dealersData[0];
  const idxId = dHeaders.indexOf("DealerID");
  const idxMobile = dHeaders.indexOf("মোবাইল");
  const idxThikana = dHeaders.indexOf("ঠিকানা");
  for (let i = 1; i < dealersData.length; i++) {
    if (dealersData[i][idxId] === dealerId) {
      return { mobile: dealersData[i][idxMobile], thikana: dealersData[i][idxThikana] };
    }
  }
  return { mobile: "", thikana: "" };
}

/*******************************************************
 * অনুরোধকারী ডিপু (এজেন্সি) নাকি নির্দিষ্ট ডিলার তা যাচাই —
 * এজেন্সি হলে null (সব ডিলারের ডাটা দেখতে/এডিট করতে পারবে),
 * ডিলার হলে তার নিজের DealerID ফেরত দেয় (শুধু নিজের ডাটার সীমা)
 *******************************************************/
function resolveRequesterDealerScope(perm) {
  return perm.payload.dealerId === AGENCY_ID ? null : perm.payload.dealerId;
}

function addSalesInvoice(data) {
  const perm = checkPermission(data.token, ["Admin", "প্রতিনিধি"]);
  if (!perm.ok) return { success: false, message: perm.message };
  const scopeDealerId = resolveRequesterDealerScope(perm);

  let targetDealerId;
  if (scopeDealerId === null) {
    // ডিপু থেকে সাবমিট — ফরম থেকে ডিলার সিলেক্ট করা আবশ্যক
    targetDealerId = data.dealerId;
    if (!targetDealerId) return { success: false, message: "ডিলার সিলেক্ট করুন" };
  } else {
    // ডিলার থেকে সাবমিট — সবসময় নিজের একাউন্টের জন্যই তৈরি হবে
    targetDealerId = scopeDealerId;
  }

  const masterSS = getMasterSS();
  const sheet = getSheet(masterSS, "SalesInvoice");
  const dealersSheet = getSheet(masterSS, "Dealers");
  ensureSheetColumns(sheet, ["বিক্রয়কারী", "পণ্য পরিবর্তন", "লাইন সংরক্ষণ"]);   // পুরনো শীটে না থাকলে নিজে যোগ হয়
  const dealerInfo = getDealerRow(dealersSheet, targetDealerId);
  if (!dealerInfo) return { success: false, message: "ডিলার পাওয়া যায়নি" };

  // প্যাকেজের পণ্য বদল (থাকলে) যাচাই করে দাম বসানো
  const customErr = applyInvoiceCustomItems(masterSS, data.items || [], []);
  if (customErr) return { success: false, message: customErr };
  if (scopeDealerId !== null && (data.items || []).some(function (it) { return it.customJson; })) {
    return { success: false, message: "অর্ডারের পণ্য পরিবর্তন শুধু ডিপু করতে পারবে" };
  }

  // স্টকের বেশি অর্ডার করা যাবে না
  const stockErr = checkOrderStock(masterSS, data.items || [], "");
  if (stockErr) return { success: false, message: stockErr };

  const contact = getDealerContactInfo(dealersSheet, targetDealerId);
  const invoiceNo = "SI" + Utilities.formatString("%06d", Math.floor(Math.random() * 900000) + 100000);
  const now = data.date ? new Date(data.date) : new Date();

  const initialStatus = scopeDealerId === null ? "পেন্ডিং" : "অপেক্ষমান";
  const built = buildSalesInvoiceRows(
    invoiceNo, { dealerId: targetDealerId, name: dealerInfo.name }, contact.mobile, contact.thikana,
    now, data.items || [], Number(data.discount) || 0, Number(data.paid) || 0, initialStatus, String(data.seller || "").trim()
  );

  batchAppendRows(sheet, built.rows, "SIE", "EntryID");

  return {
    success: true, invoiceNo: invoiceNo, subtotal: built.subtotal, netPayable: built.netPayable, due: built.due,
    status: initialStatus,
    message: scopeDealerId === null ? "অর্ডার ইনভয়েস যোগ হয়েছে" : "অর্ডার পাঠানো হয়েছে, ডিপুর কনফার্মেশনের অপেক্ষায় আছে"
  };
}

/*******************************************************
 * বিদ্যমান ইনভয়েস সম্পূর্ণ এডিট — পুরনো সব লাইন ডিলিট করে
 * নতুন করে একই ইনভয়েস নং দিয়ে আবার তৈরি করা হয়। ডিপু যেকোনো
 * ইনভয়েস এডিট করতে পারবে, ডিলার শুধু নিজের ইনভয়েস এডিট করতে পারবে
 *******************************************************/
function updateSalesInvoiceFull(data) {
  const perm = checkPermission(data.token, ["Admin", "প্রতিনিধি"]);
  if (!perm.ok) return { success: false, message: perm.message };
  const scopeDealerId = resolveRequesterDealerScope(perm);

  const masterSS = getMasterSS();
  const sheet = getSheet(masterSS, "SalesInvoice");
  const dealersSheet = getSheet(masterSS, "Dealers");

  ensureSheetColumns(sheet, ["বিক্রয়কারী", "পণ্য পরিবর্তন", "লাইন সংরক্ষণ"]);
  // পুরনো স্ট্যাটাস ও মালিকানা (DealerID) যাচাই
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const idxInv = headers.indexOf("ইনভয়েস নং");
  const idxStatus = headers.indexOf("স্ট্যাটাস");
  const idxDealerId = headers.indexOf("DealerID");
  let oldStatus = "পেন্ডিং";
  let existingDealerId = null;
  for (let i = 1; i < values.length; i++) {
    if (values[i][idxInv] === data.invoiceNo) {
      oldStatus = values[i][idxStatus] || "পেন্ডিং";
      existingDealerId = values[i][idxDealerId];
      break;
    }
  }
  if (existingDealerId === null) return { success: false, message: "ইনভয়েস পাওয়া যায়নি" };
  if (scopeDealerId !== null && existingDealerId !== scopeDealerId) {
    return { success: false, message: "এই ইনভয়েস এডিট করার অনুমতি আপনার নেই" };
  }
  const targetDealerId = scopeDealerId === null ? (data.dealerId || existingDealerId) : scopeDealerId;

  // প্যাকেজের পণ্য বদল (থাকলে) যাচাই করে দাম বসানো
  const oldRows = [];
  const idxPkgOld = headers.indexOf("PackageID"), idxCusOld = headers.indexOf("পণ্য পরিবর্তন"), idxSnapOld = headers.indexOf("লাইন সংরক্ষণ");
  for (let i = 1; i < values.length; i++) {
    if (values[i][idxInv] !== data.invoiceNo) continue;
    let sn = null, cu = [];
    try { sn = JSON.parse(values[i][idxSnapOld] || "null"); } catch (e) { sn = null; }
    try { cu = JSON.parse(values[i][idxCusOld] || "[]") || []; } catch (e) { cu = []; }
    oldRows.push({ packageId: values[i][idxPkgOld], snapshot: sn, customArr: cu });
  }
  const customErr = applyInvoiceCustomItems(masterSS, data.items || [], oldRows);
  if (customErr) return { success: false, message: customErr };
  // পণ্য বদল (এডিট মোড) শুধু ডিপু করতে পারবে — ডিলার শুধু আগে থেকে থাকা বদল অপরিবর্তিত রাখতে পারবে
  if (scopeDealerId !== null) {
    const idxCustom = headers.indexOf("পণ্য পরিবর্তন");
    const okKeys = {};
    for (let i = 1; i < values.length; i++) {
      if (values[i][idxInv] === data.invoiceNo) okKeys[values[i][headers.indexOf("PackageID")] + "|" + (idxCustom === -1 ? "" : String(values[i][idxCustom] || ""))] = true;
    }
    const bad = (data.items || []).some(function (it) { return it.customJson && !okKeys[it.packageId + "|" + it.customJson]; });
    if (bad) return { success: false, message: "অর্ডারের পণ্য পরিবর্তন শুধু ডিপু করতে পারবে" };
  }

  // স্টকের বেশি অর্ডার করা যাবে না (এই ইনভয়েসের নিজের পুরনো সংখ্যা বাদ দিয়ে হিসাব)
  const stockErr = checkOrderStock(masterSS, data.items || [], data.invoiceNo);
  if (stockErr) return { success: false, message: stockErr };

  // পুরনো লাইন সব ডিলিট
  for (let i = values.length - 1; i >= 1; i--) {
    if (values[i][idxInv] === data.invoiceNo) sheet.deleteRow(i + 1);
  }

  const dealerInfo = getDealerRow(dealersSheet, targetDealerId);
  if (!dealerInfo) return { success: false, message: "ডিলার পাওয়া যায়নি" };
  const contact = getDealerContactInfo(dealersSheet, targetDealerId);
  const now = data.date ? new Date(data.date) : new Date();

  const built = buildSalesInvoiceRows(
    data.invoiceNo, { dealerId: targetDealerId, name: dealerInfo.name }, contact.mobile, contact.thikana,
    now, data.items || [], Number(data.discount) || 0, Number(data.paid) || 0, oldStatus, String(data.seller || "").trim()
  );

  batchAppendRows(sheet, built.rows, "SIE", "EntryID");

  return {
    success: true, invoiceNo: data.invoiceNo, subtotal: built.subtotal, netPayable: built.netPayable, due: built.due,
    message: "ইনভয়েস আপডেট হয়েছে"
  };
}

/*******************************************************
 * ইনভয়েসের ডেলিভারি স্ট্যাটাস আপডেট (সব লাইনে একসাথে) — শুধু ডিপু
 *******************************************************/
function updateSalesInvoiceStatus(data) {
  const perm = checkAgencyPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };

  const status = data.status || "ডেলিভারি সম্পন্ন";
  const masterSS = getMasterSS();
  const sheet = getSheet(masterSS, "SalesInvoice");

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const values = sheet.getDataRange().getValues();
    const headers = values[0];
    const idxInv = headers.indexOf("ইনভয়েস নং");
    const idxStatus = headers.indexOf("স্ট্যাটাস");
    const idxPaid = headers.indexOf("পরিশোধ");
    const idxDue = headers.indexOf("বকেয়া");
    const idxPayable = headers.indexOf("পরিশোধযোগ্য মূল্য");
    const idxDiscount = headers.indexOf("ডিসকাউন্ট");

    const rowIdx = [];
    for (let i = 1; i < values.length; i++) {
      if (values[i][idxInv] === data.invoiceNo) rowIdx.push(i);
    }
    if (rowIdx.length === 0) return { success: false, message: "ইনভয়েস পাওয়া যায়নি" };

    let collect = 0, newPaid = null, newDue = null;
    if (status === "ডেলিভারি সম্পন্ন") {
      const first = values[rowIdx[0]];
      if (first[idxStatus] === "অপেক্ষমান") {
        return { success: false, message: "অর্ডারটি এখনো কনফার্ম করা হয়নি — আগে কনফার্ম করুন" };
      }

      // বকেয়া যাচাই — বকেয়া থাকলে হয় আদায় করতে হবে, নয়তো বকেয়া রেখে ডেলিভারির অনুমতি (allowDue) লাগবে
      let payable = 0;
      rowIdx.forEach(function (i) { payable += Number(values[i][idxPayable]) || 0; });
      const net = payable - (Number(first[idxDiscount]) || 0);
      const paid = Number(first[idxPaid]) || 0;
      const dueNow = Math.round((net - paid) * 100) / 100;

      collect = Number(data.collectAmount) || 0;
      if (collect < 0 || collect > Math.max(dueNow, 0) + 0.009) {
        return { success: false, message: "আদায়ের পরিমাণ ০ থেকে বকেয়া (" + dueNow + ") এর মধ্যে হতে হবে", due: dueNow };
      }
      newPaid = paid + collect;
      newDue = Math.round((net - newPaid) * 100) / 100;

      if (newDue > 0.009 && data.allowDue !== true) {
        return {
          success: false, needsCollection: true, due: newDue,
          message: "ইনভয়েসে বকেয়া " + newDue + " টাকা আছে — বকেয়া আদায় করুন অথবা বকেয়া রেখে ডেলিভারির অনুমতি দিন"
        };
      }
    }

    rowIdx.forEach(function (i) {
      if (collect > 0) {
        sheet.getRange(i + 1, idxPaid + 1).setValue(newPaid);
        sheet.getRange(i + 1, idxDue + 1).setValue(newDue);
      }
      sheet.getRange(i + 1, idxStatus + 1).setValue(status);
    });
    return {
      success: true, due: newDue, paid: newPaid,
      message: (collect > 0 ? collect + " টাকা আদায় হয়েছে। " : "") + "স্ট্যাটাস আপডেট হয়েছে"
    };
  } finally {
    lock.releaseLock();
  }
}

/*******************************************************
 * ইনভয়েস লিস্ট — ডিপু সব ডিলারের ইনভয়েস দেখবে, ডিলার শুধু
 * নিজের ইনভয়েসগুলো দেখবে (এজেন্সি বা যেকোনো ডিলার থেকে তৈরি হোক না কেন)
 *******************************************************/
function listSalesInvoices(data) {
  const perm = checkPermission(data.token, ["Admin", "প্রতিনিধি"]);
  if (!perm.ok) return { success: false, message: perm.message };
  const scopeDealerId = resolveRequesterDealerScope(perm);

  const masterSS = getMasterSS();
  const sheet = getSheet(masterSS, "SalesInvoice");
  let rows = genericListRows(sheet).filter(function (r) { return r["স্ট্যাটাস"] !== "অপেক্ষমান"; });
  if (scopeDealerId !== null) {
    rows = rows.filter(function (r) { return r["DealerID"] === scopeDealerId; });
  }
  return { success: true, entries: rows };
}

function updateSalesInvoiceEntry(data) {
  const perm = checkAgencyPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };

  const masterSS = getMasterSS();
  const sheet = getSheet(masterSS, "SalesInvoice");
  const ok = genericUpdateRow(sheet, "EntryID", data.entryId, data.fields || {});
  return { success: ok, message: ok ? "আপডেট হয়েছে" : "এন্ট্রি পাওয়া যায়নি" };
}

/*******************************************************
 * ইনভয়েস ডিলিট — ডিপু যেকোনো ইনভয়েস ডিলিট করতে পারবে,
 * ডিলার শুধু নিজের ইনভয়েস ডিলিট করতে পারবে
 *******************************************************/
function deleteSalesInvoice(data) {
  const perm = checkPermission(data.token, ["Admin", "প্রতিনিধি"]);
  if (!perm.ok) return { success: false, message: perm.message };
  const scopeDealerId = resolveRequesterDealerScope(perm);

  const masterSS = getMasterSS();
  const sheet = getSheet(masterSS, "SalesInvoice");
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const idx = headers.indexOf("ইনভয়েস নং");
  const idxDealerId = headers.indexOf("DealerID");

  // মালিকানা যাচাই (ডিলার হলে শুধু নিজের ইনভয়েস)
  if (scopeDealerId !== null) {
    for (let i = 1; i < values.length; i++) {
      if (values[i][idx] === data.invoiceNo && values[i][idxDealerId] !== scopeDealerId) {
        return { success: false, message: "এই ইনভয়েস ডিলিট করার অনুমতি আপনার নেই" };
      }
    }
  }

  let found = false;
  for (let i = values.length - 1; i >= 1; i--) {
    if (values[i][idx] === data.invoiceNo) {
      sheet.deleteRow(i + 1);
      found = true;
    }
  }
  return { success: found, message: found ? "ইনভয়েস ডিলিট হয়েছে" : "ইনভয়েস পাওয়া যায়নি" };
}

/*=========================================================
 *  অর্ডার কনফার্মেশন নোটিফিকেশন — ডিলার থেকে সরাসরি সাবমিট করা
 *  অর্ডার প্রথমে "অপেক্ষমান" থাকে, ডিপু কনফার্ম করলে তবেই সেটি
 *  স্বাভাবিক অর্ডার লিস্টে (উভয় সাইটে) দেখা যায়
 *=======================================================*/
function listPendingOrderConfirmations(data) {
  const perm = checkPermission(data.token, ["Admin", "প্রতিনিধি"]);
  if (!perm.ok) return { success: false, message: perm.message };
  const scopeDealerId = resolveRequesterDealerScope(perm);

  const masterSS = getMasterSS();
  const sheet = getSheet(masterSS, "SalesInvoice");
  let rows = genericListRows(sheet).filter(function (r) { return r["স্ট্যাটাস"] === "অপেক্ষমান"; });
  if (scopeDealerId !== null) {
    rows = rows.filter(function (r) { return r["DealerID"] === scopeDealerId; });
  }
  return { success: true, entries: rows };
}

function confirmPendingOrder(data) {
  const perm = checkAgencyPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };

  const masterSS = getMasterSS();
  const sheet = getSheet(masterSS, "SalesInvoice");
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const idxInv = headers.indexOf("ইনভয়েস নং");
  const idxStatus = headers.indexOf("স্ট্যাটাস");

  let found = false;
  for (let i = 1; i < values.length; i++) {
    if (values[i][idxInv] === data.invoiceNo && values[i][idxStatus] === "অপেক্ষমান") {
      sheet.getRange(i + 1, idxStatus + 1).setValue("পেন্ডিং");
      found = true;
    }
  }
  return { success: found, message: found ? "অর্ডার কনফার্ম হয়েছে" : "অপেক্ষমান অর্ডার পাওয়া যায়নি" };
}

function rejectPendingOrder(data) {
  const perm = checkAgencyPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };

  const masterSS = getMasterSS();
  const sheet = getSheet(masterSS, "SalesInvoice");
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const idxInv = headers.indexOf("ইনভয়েস নং");
  const idxStatus = headers.indexOf("স্ট্যাটাস");

  let found = false;
  for (let i = values.length - 1; i >= 1; i--) {
    if (values[i][idxInv] === data.invoiceNo && values[i][idxStatus] === "অপেক্ষমান") {
      sheet.deleteRow(i + 1);
      found = true;
    }
  }
  return { success: found, message: found ? "অর্ডার বাতিল হয়েছে" : "অপেক্ষমান অর্ডার পাওয়া যায়নি" };
}

/*=========================================================
 *  এজেন্সি রিপোর্ট — বিক্রয় ইনভয়েসের ভিত্তিতে দৈনিক/মাসিক/মোট,
 *  ডিলার-ভিত্তিক অথবা সব ডিলার মিলিয়ে
 *  data: { token, mode: "daily"/"monthly"/"total", date, year, month, dealerId (ঐচ্ছিক) }
 *=======================================================*/
function agencySalesReport(data) {
  const perm = checkAgencyPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };

  const masterSS = getMasterSS();
  const sheet = getSheet(masterSS, "SalesInvoice");
  // কনফার্মেশনের অপেক্ষায় থাকা (অপেক্ষমান) অর্ডার বিক্রি হিসেবে ধরা হয় না
  let entries = genericListRows(sheet).filter(function (e) { return e["স্ট্যাটাস"] !== "অপেক্ষমান"; });

  // একই ইনভয়েসের একাধিক লাইন থেকে ডুপ্লিকেট সাবটোটাল এড়াতে,
  // প্রতিটি ইনভয়েস নং একবারই গোনা হবে সাবটোটাল/ডিসকাউন্ট/পরিশোধ/বকেয়ার জন্য,
  // কিন্তু "মোট মূল্য" প্রতিটি লাইন থেকেই যোগ হবে (পণ্য/প্যাকেজ-ভিত্তিক বিক্রি)
  if (data.dealerId) {
    const wantId = String(data.dealerId).trim();
    entries = entries.filter(function (e) { return String(e["DealerID"] == null ? "" : e["DealerID"]).trim() === wantId; });
  }

  if (data.mode === "daily" && data.date) {
    entries = entries.filter(function (e) {
      return toDateStr(e["তারিখ"]) === data.date;
    });
  } else if (data.mode === "monthly" && data.year && data.month) {
    const year = Number(data.year), month = Number(data.month);
    entries = entries.filter(function (e) {
      const d = (e["তারিখ"] instanceof Date) ? e["তারিখ"] : new Date(e["তারিখ"]);
      return d.getFullYear() === year && (d.getMonth() + 1) === month;
    });
  } else if (data.mode === "yearly" && data.year) {
    const year = Number(data.year);
    entries = entries.filter(function (e) {
      const d = (e["তারিখ"] instanceof Date) ? e["তারিখ"] : new Date(e["তারিখ"]);
      return d.getFullYear() === year;
    });
  }
  // mode === "total" হলে ফিল্টার ছাড়াই সব

  const seenInvoices = {};
  const orderMap = {}, orderList = [];   // ইনভয়েস-ভিত্তিক অর্ডার তালিকা (ডিলার রিপোর্টের নিচের টেবিলের জন্য)
  let totalSale = 0, totalDiscount = 0, totalPaid = 0, totalDue = 0, totalCommission = 0;
  let totalOrders = 0, totalDelivered = 0, totalPending = 0;
  const byDealer = {};

  entries.forEach(function (e) {
    totalSale += Number(e["মোট মূল্য"]) || 0;
    totalCommission += Number(e["কমিশন মূল্য"]) || 0;   // প্রতিটি লাইনের কমিশন যোগ হয়

    const dealerKey = e["DealerID"];
    if (!byDealer[dealerKey]) {
      byDealer[dealerKey] = {
        dealerId: dealerKey, dealerName: e["ডিলার নাম"],
        totalSale: 0, totalDiscount: 0, totalPaid: 0, totalDue: 0, totalCommission: 0,
        totalOrders: 0, totalDelivered: 0, totalPending: 0
      };
    }
    byDealer[dealerKey].totalSale += Number(e["মোট মূল্য"]) || 0;
    byDealer[dealerKey].totalCommission += Number(e["কমিশন মূল্য"]) || 0;

    // ইনভয়েস-ভিত্তিক অর্ডার সারি: প্রতিটি লাইনের মূল্য/কমিশন যোগ, ডিসকাউন্ট/পরিশোধ/বকেয়া একবার
    const ordKey = e["ইনভয়েস নং"];
    if (!orderMap[ordKey]) {
      orderMap[ordKey] = {
        invoiceNo: ordKey, date: e["তারিখ"], dealerId: e["DealerID"], dealerName: e["ডিলার নাম"],
        packages: [], gross: 0, commission: 0, payable: 0,
        discount: Number(e["ডিসকাউন্ট"]) || 0, paid: Number(e["পরিশোধ"]) || 0, due: Number(e["বকেয়া"]) || 0,
        status: e["স্ট্যাটাস"] || "পেন্ডিং"
      };
      orderList.push(orderMap[ordKey]);
    }
    const ord = orderMap[ordKey];
    ord.gross += Number(e["মোট মূল্য"]) || 0;
    ord.commission += Number(e["কমিশন মূল্য"]) || 0;
    ord.payable += Number(e["পরিশোধযোগ্য মূল্য"]) || 0;
    ord.packages.push(String(e["প্যাকেজ"] || "") + " × " + (Number(e["সংখ্যা"]) || 0));

    // ইনভয়েস-লেভেল মান (সাবটোটাল/ডিসকাউন্ট/পরিশোধ/বকেয়া/অর্ডার/স্ট্যাটাস) প্রতিটি ইনভয়েসে একবারই গোনা
    const invKey = e["ইনভয়েস নং"];
    if (!seenInvoices[invKey]) {
      seenInvoices[invKey] = true;
      totalDiscount += Number(e["ডিসকাউন্ট"]) || 0;
      totalPaid += Number(e["পরিশোধ"]) || 0;
      totalDue += Number(e["বকেয়া"]) || 0;
      totalOrders += 1;
      const status = e["স্ট্যাটাস"] || "পেন্ডিং";
      if (status === "ডেলিভারি সম্পন্ন") totalDelivered += 1; else totalPending += 1;

      byDealer[dealerKey].totalDiscount += Number(e["ডিসকাউন্ট"]) || 0;
      byDealer[dealerKey].totalPaid += Number(e["পরিশোধ"]) || 0;
      byDealer[dealerKey].totalDue += Number(e["বকেয়া"]) || 0;
      byDealer[dealerKey].totalOrders += 1;
      if (status === "ডেলিভারি সম্পন্ন") byDealer[dealerKey].totalDelivered += 1; else byDealer[dealerKey].totalPending += 1;
    }
  });

  return {
    success: true,
    report: {
      totalInvoiceCount: Object.keys(seenInvoices).length,
      totalOrders: totalOrders,
      totalDelivered: totalDelivered,
      totalPending: totalPending,
      totalSale: totalSale,
      totalDiscount: totalDiscount,
      totalPaid: totalPaid,
      totalDue: totalDue,
      totalCommission: totalCommission,
      byDealer: Object.keys(byDealer).map(function (k) { return byDealer[k]; }),
      // সংশ্লিষ্ট ডিলারের (বা সব ডিলারের) সকল অর্ডার — তারিখ অনুযায়ী পুরনো থেকে নতুন
      orders: orderList.map(function (o) {
        o.packages = o.packages.join(", ");
        o.net = o.payable - o.discount;
        return o;
      }).sort(function (a, b) { return new Date(a.date) - new Date(b.date); })
    }
  };
}

/*=========================================================
 *  ডিপু রিপোর্ট — সামগ্রিক ব্যবসায়িক সারসংক্ষেপ (স্টক, বিক্রি,
 *  খরচ, লাভ/লস, কোম্পানি পরিশোধ/বকেয়া, ডিলার/অর্ডার সংখ্যা)
 *  data: { token }
 *=======================================================*/
function getPackageProductCountMap(masterSS) {
  const pkgSheet = getSheet(masterSS, "Packages");
  const packages = genericListRows(pkgSheet);
  const map = {};
  packages.forEach(function (p) {
    let count = 0;
    for (let i = 1; i <= MAX_PACKAGE_ITEMS; i++) {
      if (p["পণ্য" + i + " - নাম ও পরিমাণ"]) count += Number(stockToEn(p["পণ্য" + i + " - সংখ্যা"])) || 1;
    }
    map[p["PackageID"]] = count;
  });
  return map;
}

/*******************************************************
 * data.mode অনুযায়ী একটি সারি (তারিখ ফিল্ড সহ) নির্দিষ্ট সময়ের
 * মধ্যে পড়ে কিনা যাচাই — ডিপু ও ডিলার রিপোর্ট উভয়ে ব্যবহারযোগ্য
 *******************************************************/
function isInReportPeriod(dateVal, mode, data) {
  if (!mode || mode === "total") return true;
  const d = (dateVal instanceof Date) ? dateVal : new Date(dateVal);
  if (mode === "daily" && data.date) return toDateStr(d) === data.date;
  if (mode === "monthly" && data.year && data.month) {
    return d.getFullYear() === Number(data.year) && (d.getMonth() + 1) === Number(data.month);
  }
  if (mode === "yearly" && data.year) return d.getFullYear() === Number(data.year);
  return true;
}

function depotReport(data) {
  const perm = checkAgencyPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };

  const masterSS = getMasterSS();
  const mode = data.mode || "total";

  // ---- স্টক (ক্রয়) — সব সময়ের ডাটা (বর্তমান স্টক হিসাবের জন্য) ----
  const allStockLines = genericListRows(getSheet(masterSS, "StockInVoucher"));
  let allStockQty = 0, allStockValue = 0;
  allStockLines.forEach(function (e) {
    allStockQty += Number(e["সংখ্যা"]) || 0;
    allStockValue += Number(e["মোট মূল্য"]) || 0;
  });
  const avgUnitCost = allStockQty > 0 ? (allStockValue / allStockQty) : 0;

  // ---- স্টক — নির্বাচিত সময়ের মধ্যে (period metrics) ----
  const stockLines = allStockLines.filter(function (e) { return isInReportPeriod(e["তারিখ"], mode, data); });
  const seenStockVouchers = {};
  let totalStockQty = 0, totalStockPurchaseValue = 0, totalPaidToCompany = 0, totalDueToCompany = 0;
  stockLines.forEach(function (e) {
    totalStockQty += Number(e["সংখ্যা"]) || 0;
    totalStockPurchaseValue += Number(e["মোট মূল্য"]) || 0;
    const vKey = e["ভাউচার নং"];
    if (!seenStockVouchers[vKey]) {
      seenStockVouchers[vKey] = true;
      totalPaidToCompany += Number(e["পরিশোধ"]) || 0;
      totalDueToCompany += Number(e["বকেয়া"]) || 0;
    }
  });

  // ---- বিক্রি — সব সময়ের ডাটা (বর্তমান স্টক থেকে বিক্রিত বাদ দেওয়ার জন্য) ----
  const pkgProductCountMap = getPackageProductCountMap(masterSS);
  const allSalesLines = genericListRows(getSheet(masterSS, "SalesInvoice"));
  let allProductsSoldQty = 0;
  allSalesLines.forEach(function (e) {
    const productCount = pkgProductCountMap[e["PackageID"]] || 0;
    allProductsSoldQty += productCount * (Number(e["সংখ্যা"]) || 0);
  });
  const currentStockQty = allStockQty - allProductsSoldQty;
  const currentStockValue = currentStockQty * avgUnitCost;

  // ---- বিক্রি — নির্বাচিত সময়ের মধ্যে (period metrics) ----
  const salesLines = allSalesLines.filter(function (e) { return isInReportPeriod(e["তারিখ"], mode, data); });
  const seenSalesInvoices = {};
  let totalProductsSoldQty = 0, totalSalesValue = 0, totalPayableSum = 0;
  let totalOrders = 0, totalDelivered = 0, totalPending = 0, totalDueFromDealers = 0;

  salesLines.forEach(function (e) {
    totalSalesValue += Number(e["মোট মূল্য"]) || 0;
    totalPayableSum += Number(e["পরিশোধযোগ্য মূল্য"]) || 0;
    const productCount = pkgProductCountMap[e["PackageID"]] || 0;
    totalProductsSoldQty += productCount * (Number(e["সংখ্যা"]) || 0);

    const invKey = e["ইনভয়েস নং"];
    if (!seenSalesInvoices[invKey]) {
      seenSalesInvoices[invKey] = true;
      totalOrders += 1;
      totalDueFromDealers += Number(e["বকেয়া"]) || 0;
      const status = e["স্ট্যাটাস"] || "পেন্ডিং";
      if (status === "ডেলিভারি সম্পন্ন") totalDelivered += 1; else totalPending += 1;
    }
  });

  // ---- খরচ (Expense) — নির্বাচিত সময়ের মধ্যে ----
  const allExpenseLines = genericListRows(getSheet(masterSS, "Expense"));
  const expenseLines = allExpenseLines.filter(function (e) { return isInReportPeriod(e["তারিখ"], mode, data); });
  let totalExpense = 0;
  expenseLines.forEach(function (e) { totalExpense += Number(e["টাকা"]) || 0; });

  // ---- লাভ/লস (নির্বাচিত সময়ের বিক্রি/খরচ ভিত্তিতে) ----
  const cogs = totalProductsSoldQty * avgUnitCost;
  const netResult = totalPayableSum - cogs - totalExpense;
  const totalProfit = netResult > 0 ? netResult : 0;
  const totalLoss = netResult < 0 ? -netResult : 0;

  // ---- ডিলার সংখ্যা (সব সময়ে নথিভুক্ত মোট ডিলার) ----
  const totalDealers = genericListRows(getSheet(masterSS, "Dealers")).length;

  return {
    success: true,
    report: {
      totalStockQty: totalStockQty,
      totalStockPurchaseValue: totalStockPurchaseValue,
      totalProductsSoldQty: totalProductsSoldQty,
      totalSalesValue: totalSalesValue,
      currentStockQty: currentStockQty,
      currentStockValue: Math.round(currentStockValue),
      totalExpense: totalExpense,
      totalProfit: Math.round(totalProfit),
      totalLoss: Math.round(totalLoss),
      totalPaidToCompany: totalPaidToCompany,
      totalDueToCompany: totalDueToCompany,
      totalDealers: totalDealers,
      totalOrders: totalOrders,
      totalDelivered: totalDelivered,
      totalPending: totalPending,
      totalSalesAmount: totalSalesValue,
      totalDueFromDealers: totalDueFromDealers
    }
  };
}


/*=========================================================
 *  বর্তমান স্টক — পণ্যভিত্তিক (পণ্যের নাম, ব্র্যান্ড, মোট পরিমাণ, মোট মূল্য)
 *  data: { token }   (ফিল্টার ব্রাউজারে হয়)
 *
 *  - ক্রয়: স্টক ভাউচারের সংখ্যা ও মোট মূল্য, ProductID ধরে
 *  - বিক্রি: প্যাকেজের পণ্য-লাইন ("চিনি - ১ কেজি") এ যে পণ্যের নাম পাওয়া যায়
 *    (একাধিক মিললে সবচেয়ে লম্বা নাম) — প্রতিটি প্যাকেজ-লাইন = ১ ইউনিট, প্যাকেজ বিক্রির
 *    সংখ্যা দিয়ে গুণ (ডিপু রিপোর্টের হিসাবের নিয়মের সাথে একই)
 *  - মোট মূল্য = বর্তমান পরিমাণ × গড় ক্রয় মূল্য
 *=======================================================*/
function stockNormName(v) { return String(v == null ? "" : v).replace(/\s+/g, "").toLowerCase(); }

/* প্যাকেজের পণ্য-লাইনের লেখা (যেমন "চিনি ১ কেজি") থেকে শেষের একক (কেজি/লিটার/পিস) আলাদা করা — পণ্যে "একক" দেওয়া না থাকলে ব্যবহৃত */
function stockUnitFromText(text) {
  const m = String(text == null ? "" : text).trim().match(/[\d০-৯][\d০-৯.,\/]*\s*([^\d০-৯\s.,\/][^\d০-৯]*?)\s*$/);
  const u = m ? m[1].trim() : "";
  return u.length <= 12 ? u : "";
}

/* ইংরেজি অঙ্ক → বাংলা */
function stockToBn(v) { return String(v == null ? "" : v).replace(/[0-9]/g, function (d) { return "০১২৩৪৫৬৭৮৯"[d]; }); }

/* বাংলা অঙ্ক → ইংরেজি */
function stockToEn(v) { return String(v == null ? "" : v).replace(/[০-৯]/g, function (d) { return "০১২৩৪৫৬৭৮৯".indexOf(d); }); }

/* "৫০০ গ্রাম" / "১ কেজি" / "২৫০ মি.লি" লেখা থেকে সাইজ → { v: কেজি বা লিটারে মান, u: "কেজি"|"লিটার" }, না চিনলে null */
function parseSizeText(text) {
  const t = stockToEn(text).replace(/,/g, "");
  const m = t.match(/(\d+(?:\.\d+)?)\s*(কিলোগ্রাম|কেজি|কেজী|kgs?|গ্রাম|গ্রা|gms?|grams?|g|লিটার|ltr|liters?|litres?|l|মি\.?\s*লি|মিলি|ml)/i);
  if (!m) return null;
  const n = Number(m[1]), w = m[2].toLowerCase();
  if (/^(কিলোগ্রাম|কেজি|কেজী|kgs?)$/.test(w)) return { v: n, u: "কেজি" };
  if (/^(গ্রাম|গ্রা|gms?|grams?|g)$/.test(w)) return { v: n / 1000, u: "কেজি" };
  if (/^(লিটার|ltr|liters?|litres?|l)$/.test(w)) return { v: n, u: "লিটার" };
  return { v: n / 1000, u: "লিটার" };
}

/* পণ্যের প্রতি প্যাকেট/পিসের সাইজ: "পরিমাণ" = ৫০০ আর "একক" = গ্রাম (অথবা পরিমাণে "৫০০ গ্রাম" লেখা) */
function productSizeOf(poriman, unit) {
  const raw = String(poriman == null ? "" : poriman).trim(), un = String(unit || "").trim();
  let sz = parseSizeText(raw);
  if (!sz && /^[\d.]+$/.test(stockToEn(raw)) && un) sz = parseSizeText(stockToEn(raw) + " " + un);
  return sz;
}

/* প্যাকেজের একটিতে এই পণ্য কত পিস/প্যাকেট: "১ কেজি", "৫০০ গ্রাম" লিখলে পণ্যের সাইজ দিয়ে ভাগ হয়; শুধু সংখ্যা লিখলে সেটাই পিস সংখ্যা */
function packsPerPackage(raw, size) {
  const s = String(raw == null ? "" : raw).trim();
  if (!s) return 1;
  const a = parseSizeText(s);
  if (a) return (size && size.u === a.u && size.v > 0) ? Math.round(a.v / size.v * 10000) / 10000 : a.v;
  const n = parseFloat(stockToEn(s));   // "২ পিস" / "২ প্যাকেট" / "২" → ২
  return n > 0 ? n : 1;
}

/* সংখ্যা + একক এবং বস্তা/বক্স হিসাব, যেমন "110 কেজি (2 বস্তা 10 কেজি)" */
function formatStockQty(qty, unit, packType, packSize) {
  qty = Math.round((Number(qty) || 0) * 100) / 100;
  let out = qty + (unit ? " " + unit : "");
  packSize = Number(packSize) || 0;
  if (packType && packSize > 0 && qty >= packSize) {
    const n = Math.floor(qty / packSize);
    const rem = Math.round((qty - n * packSize) * 100) / 100;
    out += " (" + n + " " + packType + (rem > 0 ? " " + rem + (unit ? " " + unit : "") : "") + ")";
  }
  return out;
}

/*******************************************************
 * স্টকের শেয়ার্ড হিসাব — currentStockReport ও অর্ডার ইনভয়েসের স্টক যাচাই দুটোতেই ব্যবহৃত
 * excludeInvoiceNo: ইনভয়েস এডিটের সময় ওই ইনভয়েসের নিজের সংখ্যা বিক্রি হিসেবে ধরা হয় না
 *
 * প্যাকেজের পণ্য-লাইনে ProductID ও "সংখ্যা" (প্যাকেজ-প্রতি পরিমাণ) থাকলে সেটাই ব্যবহার হয়;
 * না থাকলে (পুরনো প্যাকেজ) নাম মিলিয়ে ধরা হয় এবং প্যাকেজ-প্রতি ১ একক ধরা হয়
 *******************************************************/
function computeStockState(masterSS, excludeInvoiceNo) {
  const products = genericListRows(getSheet(masterSS, "Products"));
  const byId = {};
  const list = [];
  products.forEach(function (p) {
    const r = { productId: p["ProductID"], name: String(p["পণ্যের নাম"] || ""), brand: String(p["ব্র্যান্ড"] || ""),
                unitRaw: String(p["একক"] || ""), packType: String(p["প্যাকিং ধরন"] || ""), packSize: Number(stockToEn(p["প্যাকিং পরিমাণ"])) || 0,
                size: productSizeOf(p["পরিমাণ"], p["একক"]),
                combo: Number(p["কম্বো মূল্য"]) || 0, poriman: String(p["পরিমাণ"] == null ? "" : p["পরিমাণ"]),
                qtyIn: 0, valueIn: 0, sold: 0, key: stockNormName(p["পণ্যের নাম"]) };
    r.unit = r.size ? "পিস" : (r.unitRaw || "পিস");   // স্টক গোনা হয় পিস/প্যাকেটে; সাইজ না থাকলে পণ্যের এককে
    byId[r.productId] = r;
    list.push(r);
  });

  genericListRows(getSheet(masterSS, "StockInVoucher")).forEach(function (e) {
    const r = byId[e["ProductID"]];
    if (!r) return;
    r.qtyIn += Number(e["সংখ্যা"]) || 0;
    r.valueIn += Number(e["মোট মূল্য"]) || 0;
  });

  // প্যাকেজ → কোন কোন পণ্য (পণ্য-লাইন ধরে)
  const unmatched = {};
  const pkgLines = {};   // PackageID -> [{ r|null, unit, per }, ...] (প্যাকেজের পণ্য-লাইনের ক্রমে)
  genericListRows(getSheet(masterSS, "Packages")).forEach(function (pk) {
    const lines = [];
    for (let i = 1; i <= MAX_PACKAGE_ITEMS; i++) {
      const text = pk["পণ্য" + i + " - নাম ও পরিমাণ"];
      if (!text) continue;
      let best = byId[pk["পণ্য" + i + " - ProductID"]] || null;
      if (!best) {
        const t = stockNormName(text);
        list.forEach(function (r) {
          if (r.key && t.indexOf(r.key) !== -1 && (!best || r.key.length > best.key.length)) best = r;
        });
      }
      if (!best) unmatched[String(text)] = true;
      const perHere = best ? packsPerPackage(pk["পণ্য" + i + " - সংখ্যা"], best.size) : 1;
      lines.push({
        r: best,
        unit: best ? best.unit : stockUnitFromText(text),
        per: perHere
      });
    }
    pkgLines[pk["PackageID"]] = lines;
  });
  // মূল্য তালিকা (PriceList) থেকে দাম
  const plMap = priceListMap(masterSS);
  list.forEach(function (r) {
    const e = plMap[r.productId];
    r.pl = e || null;
  });

  genericListRows(getSheet(masterSS, "SalesInvoice")).forEach(function (e) {
    if (excludeInvoiceNo && e["ইনভয়েস নং"] === excludeInvoiceNo) return;
    const q = Number(e["সংখ্যা"]) || 0;
    effectiveInvoiceLines({ pkgLines: pkgLines, byId: byId }, e["PackageID"], e["পণ্য পরিবর্তন"]).forEach(function (l) { if (l.r) l.r.sold += q * l.per; });
  });

  return { list: list, pkgLines: pkgLines, unmatched: unmatched, byId: byId };
}

/*******************************************************
 * অর্ডার ইনভয়েসে প্যাকেজের পণ্য বদল — কোন লাইনে (i = পণ্য-লাইনের ক্রম, ০ থেকে) কোন পণ্য বসেছে
 * customJson: [{ i, productId, name, text, per, price, orig }] — না থাকলে প্যাকেজের মূল পণ্যই
 * ফেরত: স্টক হিসাবের জন্য পণ্য-লাইনের তালিকা (বদল প্রয়োগ করা)
 *******************************************************/
function effectiveInvoiceLines(st, packageId, customJson) {
  const base = (st.pkgLines[packageId] || []).slice();
  if (!customJson) return base;
  let arr;
  try { arr = JSON.parse(customJson); } catch (err) { return base; }
  if (!Array.isArray(arr)) return base;
  arr.forEach(function (c) {
    const k = Number(c.i);
    if (!(k >= 0 && k < base.length)) return;
    const r = st.byId[c.productId] || null;
    base[k] = { r: r, unit: r ? r.unit : "", per: Number(c.per) || 1 };
  });
  return base;
}

/*******************************************************
 * ইনভয়েসের প্রতিটি প্যাকেজ-লাইনের চূড়ান্ত পণ্য ও দাম হিসাব করে "সংরক্ষণ" (snapshot) তৈরি —
 * পরে প্যাকেজ বা মূল্য তালিকার দাম বদলালেও এই ইনভয়েসের দাম/লেখা বদলাবে না
 *
 * item.custom: [{ i, productId, qty?, unit? }] — i = প্যাকেজের পণ্য-লাইনের ক্রম (০ থেকে)
 * oldRows: এডিটের সময় ইনভয়েসের আগের সারি [{ packageId, snapshot, customArr }] — আগের দাম ধরে রাখতে
 * ফলাফল (প্রতি item এ): unitPrice (প্যাকেজ-প্রতি মোট), customJson (পণ্য পরিবর্তন), snapshotJson (লাইন সংরক্ষণ)
 * নতুন/বদলানো পণ্যের মূল্য = মূল্য তালিকার কম্বো মূল্য × (পরিমাণ ÷ তালিকার পরিমাণ); তালিকায় না থাকলে ০
 * ব্যর্থ হলে বার্তা ফেরত দেয়, ঠিক থাকলে null
 *******************************************************/
function applyInvoiceCustomItems(masterSS, items, oldRows) {
  items = items || [];
  const products = {};
  genericListRows(getSheet(masterSS, "Products")).forEach(function (p) { products[p["ProductID"]] = p; });
  const stState = computeStockState(masterSS, "");   // মূল্য তালিকা (pl) এখানে
  const pkgs = {};
  genericListRows(getSheet(masterSS, "Packages")).forEach(function (p) { pkgs[p["PackageID"]] = p; });
  const olds = (oldRows || []).slice();

  for (let n = 0; n < items.length; n++) {
    const it = items[n];
    const custom = (Array.isArray(it.custom) ? it.custom : []).filter(function (c) { return c && c.productId; });
    it.customJson = ""; it.snapshotJson = "";
    const pk = pkgs[it.packageId];
    if (!pk) { if (custom.length) return "প্যাকেজ পাওয়া যায়নি"; continue; }

    const slots = [];
    for (let i = 1; i <= MAX_PACKAGE_ITEMS; i++) if (pk["পণ্য" + i + " - নাম ও পরিমাণ"]) slots.push(i);

    // এই প্যাকেজের আগের সারি (এডিটে) — আগের সংরক্ষিত দাম/লেখা ধরে রাখতে
    let old = null;
    for (let o = 0; o < olds.length; o++) { if (olds[o].packageId === it.packageId) { old = olds.splice(o, 1)[0]; break; } }
    const oldSnap = old && Array.isArray(old.snapshot) && old.snapshot.length === slots.length ? old.snapshot : null;
    const oldCustom = {};
    ((old && old.customArr) || []).forEach(function (oc) { if (oc && oc.i !== undefined) oldCustom[oc.i] = oc; });

    // শুরুর লাইন: আগের সংরক্ষণ থাকলে সেটাই, না থাকলে প্যাকেজের বর্তমান লাইন
    const lines = slots.map(function (slot, k) {
      if (oldSnap && oldSnap[k]) return { text: String(oldSnap[k].text || ""), price: Number(oldSnap[k].price) || 0 };
      if (oldCustom[k]) return { text: String(oldCustom[k].text || ""), price: Number(oldCustom[k].price) || 0 };
      return { text: String(pk["পণ্য" + slot + " - নাম ও পরিমাণ"] || ""), price: Number(pk["পণ্য" + slot + " - কম্বো মূল্য"]) || 0 };
    });

    const out = [], seen = {};
    for (let c = 0; c < custom.length; c++) {
      const k = Number(custom[c].i);
      if (!(k >= 0 && k < slots.length) || seen[k]) return "পণ্য বদলের তথ্য ঠিক নেই";
      seen[k] = true;
      const p = products[custom[c].productId];
      if (!p) return "বাছাই করা পণ্য পাওয়া যায়নি";
      const cq = custom[c].qty == null ? "" : String(custom[c].qty).trim(), cu = String(custom[c].unit || "").trim();

      // আগের ইনভয়েসে একই বদল থাকলে আগের দামই থাকবে (মূল্য তালিকা বদলালেও)
      const oc = oldCustom[k];
      if (oc && oc.productId === custom[c].productId && String(oc.qty || "") === cq && String(oc.unit || "") === cu) {
        out.push(oc); lines[k] = { text: String(oc.text || ""), price: Number(oc.price) || 0 };
        continue;
      }

      const slot = slots[k];
      let qtyRaw = pk["পণ্য" + slot + " - সংখ্যা"];
      let qtyNum = null, qtyUnit = "";
      if (cq !== "") {
        qtyNum = Number(stockToEn(cq));
        if (!(qtyNum > 0)) return "পরিমাণ ০ এর বেশি হতে হবে";
        qtyUnit = cu;
        qtyRaw = stockToEn(qtyNum) + (qtyUnit ? " " + qtyUnit : "");
      }
      const size = productSizeOf(p["পরিমাণ"], p["একক"]);
      const per = packsPerPackage(qtyRaw, size);
      const sr = stState.byId[p["ProductID"]];
      let price = 0;   // মূল্য তালিকায় দাম নির্ধারণ নেই — অনুমান করে কোনো দাম বসানো হয় না
      if (sr && sr.pl && sr.pl.combo > 0) {
        const qm = stockToEn(qtyRaw).trim().match(/^([\d.]+)\s*(.*)$/);
        price = Math.round(sr.pl.combo * priceRatio(sr.pl.qty, sr.pl.unit, qm ? qm[1] : "1", qm ? qm[2] : "") * 100) / 100;
      }
      const qtyTxt = qtyNum !== null ? stockToBn(qtyNum) + (qtyUnit ? " " + qtyUnit : "") : String(qtyRaw == null ? "" : qtyRaw).trim();
      const entry = { i: k, productId: p["ProductID"], name: String(p["পণ্যের নাম"] || ""), qty: qtyNum !== null ? String(qtyNum) : "", unit: qtyUnit, text: String(p["পণ্যের নাম"] || "") + (qtyTxt ? " " + qtyTxt : ""), per: per, price: price, orig: lines[k].price };
      out.push(entry);
      lines[k] = { text: entry.text, price: price };
    }
    if (slots.length) {
      let unit = 0;
      lines.forEach(function (l) { unit += l.price; });
      it.unitPrice = Math.round(unit * 100) / 100;
    }
    it.customJson = out.length ? JSON.stringify(out) : "";
    it.snapshotJson = JSON.stringify(lines);
  }
  return null;
}

/* আগে তৈরি হওয়া ইনভয়েস (যেগুলোতে "লাইন সংরক্ষণ" নেই) এর বর্তমান দাম/লেখা সংরক্ষণ করে ফেলা —
   এরপর প্যাকেজের দাম বদলালেও পুরনো ইনভয়েস অপরিবর্তিত থাকবে */
function freezeLegacyInvoices(masterSS) {
  const sheet = getSheet(masterSS, "SalesInvoice");
  ensureSheetColumns(sheet, ["পণ্য পরিবর্তন", "লাইন সংরক্ষণ"]);
  const last = sheet.getLastRow();
  if (last < 2) return;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const iPkg = headers.indexOf("PackageID"), iCus = headers.indexOf("পণ্য পরিবর্তন"), iSnap = headers.indexOf("লাইন সংরক্ষণ");
  const n = last - 1;
  const pkgCol = sheet.getRange(2, iPkg + 1, n, 1).getValues();
  const cusCol = sheet.getRange(2, iCus + 1, n, 1).getValues();
  const snapCol = sheet.getRange(2, iSnap + 1, n, 1).getValues();
  const pkgs = {};
  genericListRows(getSheet(masterSS, "Packages")).forEach(function (p) { pkgs[p["PackageID"]] = p; });
  let changed = false;
  for (let r = 0; r < n; r++) {
    if (snapCol[r][0]) continue;
    const pk = pkgs[pkgCol[r][0]];
    if (!pk) continue;
    const lines = [];
    for (let i = 1; i <= MAX_PACKAGE_ITEMS; i++) {
      if (pk["পণ্য" + i + " - নাম ও পরিমাণ"]) lines.push({ text: String(pk["পণ্য" + i + " - নাম ও পরিমাণ"]), price: Number(pk["পণ্য" + i + " - কম্বো মূল্য"]) || 0 });
    }
    try {
      (JSON.parse(cusCol[r][0] || "[]") || []).forEach(function (c) { if (lines[c.i]) lines[c.i] = { text: String(c.text || ""), price: Number(c.price) || 0 }; });
    } catch (e) { /* বাদ */ }
    snapCol[r][0] = JSON.stringify(lines);
    changed = true;
  }
  if (changed) sheet.getRange(2, iSnap + 1, n, 1).setValues(snapCol);
}

/* মূল্য তালিকায় কোনো পণ্যের দাম বসলে/বদলালে সব প্যাকেজের ঐ পণ্যের লাইনের বাজার/কম্বো মূল্য ও সর্বমোট নতুন করে হিসাব */
function propagatePriceToPackages(masterSS, productId, entry) {
  if (!entry || !(entry.combo > 0)) return 0;
  const sheet = getPackagesSheet(masterSS);
  let updated = 0;
  genericListRows(sheet).forEach(function (pk) {
    let hit = false;
    const fields = {};
    let totB = 0, totC = 0;
    for (let i = 1; i <= MAX_PACKAGE_ITEMS; i++) {
      if (!pk["পণ্য" + i + " - নাম ও পরিমাণ"]) continue;
      let b = Number(pk["পণ্য" + i + " - বাজার মূল্য"]) || 0, c = Number(pk["পণ্য" + i + " - কম্বো মূল্য"]) || 0;
      if (pk["পণ্য" + i + " - ProductID"] === productId) {
        const qm = stockToEn(pk["পণ্য" + i + " - সংখ্যা"]).trim().match(/^([\d.]+)\s*(.*)$/);
        if (qm) {
          const ratio = priceRatio(entry.qty, entry.unit, qm[1], qm[2]);
          b = Math.round(entry.market * ratio * 100) / 100;
          c = Math.round(entry.combo * ratio * 100) / 100;
          fields["পণ্য" + i + " - বাজার মূল্য"] = b; fields["পণ্য" + i + " - কম্বো মূল্য"] = c; fields["পণ্য" + i + " - সাশ্রয়"] = Math.round((b - c) * 100) / 100;
          hit = true;
        }
      }
      totB += b; totC += c;
    }
    if (!hit) return;
    fields["সর্বমোট বাজার মূল্য"] = Math.round(totB * 100) / 100;
    fields["সর্বমোট কম্বো মূল্য (সর্বমোট মূল্য)"] = Math.round(totC * 100) / 100;
    fields["সর্বমোট সাশ্রয়"] = Math.round((totB - totC) * 100) / 100;
    genericUpdateRow(sheet, "PackageID", pk["PackageID"], fields);
    updated++;
  });
  if (updated) invalidatePackageCache();
  return updated;
}

function currentStockReport(data) {
  const perm = checkAgencyPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };

  const st = computeStockState(getMasterSS());
  const rows = st.list.map(function (r) {
    const qty = r.qtyIn - r.sold;
    const avg = r.qtyIn > 0 ? r.valueIn / r.qtyIn : 0;
    return { productId: r.productId, name: r.name, brand: r.brand, unit: r.unit, packType: r.packType, packSize: r.packSize,
             sizeV: r.size ? r.size.v : 0, sizeU: r.size ? r.size.u : "",
             qtyIn: r.qtyIn, sold: r.sold, qty: qty, value: Math.round(qty * avg) };
  });
  return { success: true, rows: rows, unmatched: Object.keys(st.unmatched).slice(0, 10) };
}

/*******************************************************
 * অর্ডার ইনভয়েসের স্টক তথ্য — প্রতিটি প্যাকেজের প্রতিটি পণ্য-লাইনে কতটুকু স্টক আছে
 * (ডিপু ও ডিলার দুই সাইট থেকেই ডাকা যায়)
 * data: { token, excludeInvoice? }
 *******************************************************/
function orderStockInfo(data) {
  const perm = checkPermission(data.token, ["Admin", "প্রতিনিধি"]);
  if (!perm.ok) return { success: false, message: perm.message };
  const st = computeStockState(getMasterSS(), data.excludeInvoice || "");
  const packages = {};
  Object.keys(st.pkgLines).forEach(function (pid) {
    packages[pid] = st.pkgLines[pid].map(function (l) {
      return l.r
        ? { productId: l.r.productId, name: l.r.name, avail: l.r.qtyIn - l.r.sold, unit: l.unit,
            per: l.per, packType: l.r.packType, packSize: l.r.packSize,
            sizeV: l.r.size ? l.r.size.v : 0, sizeU: l.r.size ? l.r.size.u : "" }
        : null;   // স্টকের পণ্যের সাথে মেলেনি — সীমা নেই
    });
  });
  const prods = {};
  st.list.forEach(function (r) {
    prods[r.productId] = { productId: r.productId, name: r.name, avail: r.qtyIn - r.sold, unit: r.unit,
                           brand: r.brand, combo: r.combo, poriman: r.poriman, unitRaw: r.unitRaw,
                           plQty: r.pl ? r.pl.qty : "", plUnit: r.pl ? r.pl.unit : "", plCombo: r.pl ? r.pl.combo : 0, plMarket: r.pl ? r.pl.market : 0,
                           packType: r.packType, packSize: r.packSize,
                           sizeV: r.size ? r.size.v : 0, sizeU: r.size ? r.size.u : "" };
  });
  return { success: true, packages: packages, products: prods };
}

/*******************************************************
 * সংরক্ষণের আগে সার্ভার-সাইড স্টক যাচাই — বেশি অর্ডার হলে ব্যর্থ বার্তা ফেরত দেয়, ঠিক থাকলে null
 *******************************************************/
function checkOrderStock(masterSS, items, excludeInvoiceNo) {
  const st = computeStockState(masterSS, excludeInvoiceNo);
  const need = {}, order = [];
  (items || []).forEach(function (item) {
    const q = Number(item.quantity) || 0;
    if (q <= 0) return;
    effectiveInvoiceLines(st, item.packageId, item.customJson).forEach(function (l) {
      if (!l.r) return;
      const k = l.r.productId;
      if (!need[k]) { need[k] = { r: l.r, unit: l.unit, qty: 0 }; order.push(k); }
      need[k].qty += q * l.per;
    });
  });
  const problems = [];
  order.forEach(function (k) {
    const n = need[k], avail = n.r.qtyIn - n.r.sold;
    if (n.qty > avail) {
      problems.push(avail <= 0
        ? n.r.name + " — স্টক নেই"
        : n.r.name + " — স্টক আছে " + formatStockQty(avail, n.unit, n.r.packType, n.r.packSize));
    }
  });
  return problems.length ? problems.join("; ") : null;
}
