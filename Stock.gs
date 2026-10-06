/*******************************************************
 * TCB/MSDS ডিলার ম্যানেজমেন্ট সফটওয়্যার
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
  const productId = generateId(sheet, "PR");
  const marketPrice = Number(data.marketPrice) || 0;
  const comboPrice = Number(data.comboPrice) || 0;

  genericAddRow(sheet, {
    "ProductID": productId,
    "পণ্যের নাম": data.নাম,
    "ব্র্যান্ড": data.brand,
    "বাজার মূল্য": marketPrice,
    "কম্বো মূল্য": comboPrice,
    "সাশ্রয়ী": marketPrice - comboPrice
  });

  invalidateAgencyCaches();
  return { success: true, productId: productId, message: "পণ্য যোগ হয়েছে" };
}

function listProducts(data) {
  const perm = checkAgencyPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };

  const cache = CacheService.getScriptCache();
  const cached = cache.get("products_cache_v1");
  if (cached) return { success: true, products: JSON.parse(cached) };

  const masterSS = getMasterSS();
  const sheet = getSheet(masterSS, "Products");
  const products = genericListRows(sheet);

  try { cache.put("products_cache_v1", JSON.stringify(products), 30); } catch (e) { /* বাদ */ }
  return { success: true, products: products };
}

function updateProduct(data) {
  const perm = checkAgencyPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };

  const masterSS = getMasterSS();
  const sheet = getSheet(masterSS, "Products");
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
  CacheService.getScriptCache().remove("products_cache_v1");
}

/*=========================================================
 *  স্টক ইন ভাউচার (StockInVoucher) — এক ভাউচারে একাধিক পণ্য
 *=======================================================*/
function addStockVoucher(data) {
  const perm = checkAgencyPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };

  const masterSS = getMasterSS();
  const sheet = getSheet(masterSS, "StockInVoucher");
  const voucherNo = "SV" + Utilities.formatString("%06d", Math.floor(Math.random() * 900000) + 100000);
  const now = data.date ? new Date(data.date) : new Date();

  const items = data.items || [];
  let grandTotal = 0;
  const lineData = items.map(function (item) {
    const marketPrice = Number(item.marketPrice) || 0;
    const comboPrice = Number(item.comboPrice) || 0;
    const total = (Number(item.quantity) || 0) * comboPrice;
    grandTotal += total;
    return { item: item, marketPrice: marketPrice, comboPrice: comboPrice, total: total };
  });
  const paid = Number(data.paid) || 0;
  const due = grandTotal - paid;

  const rows = lineData.map(function (ld) {
    return {
      "ভাউচার নং": voucherNo,
      "তারিখ": now,
      "ProductID": ld.item.productId,
      "পণ্যের নাম": ld.item.productName || "",
      "বাজার মূল্য": ld.marketPrice,
      "কম্বো মূল্য": ld.comboPrice,
      "সাশ্রয়ী": ld.marketPrice - ld.comboPrice,
      "সংখ্যা": ld.item.quantity,
      "মোট মূল্য": ld.total,
      "সর্বমোট": grandTotal,
      "পরিশোধ": paid,
      "বকেয়া": due
    };
  });

  batchAppendRows(sheet, rows, "SVE", "EntryID");

  return { success: true, voucherNo: voucherNo, total: grandTotal, paid: paid, due: due, message: "স্টক ইন ভাউচার যোগ হয়েছে" };
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
function buildSalesInvoiceRows(invoiceNo, dealerInfo, dealerMobile, dealerThikana, date, items, discount, paid, status) {
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
      "স্ট্যাটাস": status
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
  const dealerInfo = getDealerRow(dealersSheet, targetDealerId);
  if (!dealerInfo) return { success: false, message: "ডিলার পাওয়া যায়নি" };

  const contact = getDealerContactInfo(dealersSheet, targetDealerId);
  const invoiceNo = "SI" + Utilities.formatString("%06d", Math.floor(Math.random() * 900000) + 100000);
  const now = data.date ? new Date(data.date) : new Date();

  const initialStatus = scopeDealerId === null ? "পেন্ডিং" : "অপেক্ষমান";
  const built = buildSalesInvoiceRows(
    invoiceNo, { dealerId: targetDealerId, name: dealerInfo.name }, contact.mobile, contact.thikana,
    now, data.items || [], Number(data.discount) || 0, Number(data.paid) || 0, initialStatus
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
    now, data.items || [], Number(data.discount) || 0, Number(data.paid) || 0, oldStatus
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

  const masterSS = getMasterSS();
  const sheet = getSheet(masterSS, "SalesInvoice");
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const idxInv = headers.indexOf("ইনভয়েস নং");
  const idxStatus = headers.indexOf("স্ট্যাটাস");

  let found = false;
  for (let i = 1; i < values.length; i++) {
    if (values[i][idxInv] === data.invoiceNo) {
      sheet.getRange(i + 1, idxStatus + 1).setValue(data.status || "ডেলিভারি সম্পন্ন");
      found = true;
    }
  }
  return { success: found, message: found ? "স্ট্যাটাস আপডেট হয়েছে" : "ইনভয়েস পাওয়া যায়নি" };
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
  let entries = genericListRows(sheet);

  // একই ইনভয়েসের একাধিক লাইন থেকে ডুপ্লিকেট সাবটোটাল এড়াতে,
  // প্রতিটি ইনভয়েস নং একবারই গোনা হবে সাবটোটাল/ডিসকাউন্ট/পরিশোধ/বকেয়ার জন্য,
  // কিন্তু "মোট মূল্য" প্রতিটি লাইন থেকেই যোগ হবে (পণ্য/প্যাকেজ-ভিত্তিক বিক্রি)
  if (data.dealerId) {
    entries = entries.filter(function (e) { return e["DealerID"] === data.dealerId; });
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
  let totalSale = 0, totalDiscount = 0, totalPaid = 0, totalDue = 0;
  let totalOrders = 0, totalDelivered = 0, totalPending = 0;
  const byDealer = {};

  entries.forEach(function (e) {
    totalSale += Number(e["মোট মূল্য"]) || 0;

    const dealerKey = e["DealerID"];
    if (!byDealer[dealerKey]) {
      byDealer[dealerKey] = {
        dealerId: dealerKey, dealerName: e["ডিলার নাম"],
        totalSale: 0, totalDiscount: 0, totalPaid: 0, totalDue: 0,
        totalOrders: 0, totalDelivered: 0, totalPending: 0
      };
    }
    byDealer[dealerKey].totalSale += Number(e["মোট মূল্য"]) || 0;

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
      byDealer: Object.keys(byDealer).map(function (k) { return byDealer[k]; })
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
      if (p["পণ্য" + i + " - নাম ও পরিমাণ"]) count++;
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
