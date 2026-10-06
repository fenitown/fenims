/*******************************************************
 * TCB-স্টাইল ডিলার ম্যানেজমেন্ট সফটওয়্যার
 * Crud.gs — গ্রাহক, প্যাকেজ, বিক্রি, অর্ডার এর CRUD
 *
 * এই ফাইলটিও MasterRegistry এর একই Apps Script প্রজেক্টে
 * নতুন ফাইল হিসেবে যোগ করুন (নাম দিন: Crud)
 *
 * নিয়ম: Admin সব করতে পারবে।
 * প্রতিনিধি শুধু গ্রাহক তৈরি, বিক্রি, বিক্রি বাতিল করতে পারবে।
 *******************************************************/

/*******************************************************
 * টোকেন থেকে ডিলারের নিজের Spreadsheet খুলে দেওয়া
 * — এবং প্রয়োজনে হেডার স্বয়ংক্রিয়ভাবে হালনাগাদ করে দেওয়া (নতুন
 * কলাম যোগ হলে পুরনো ডিলারদের শীটেও এটা এমনিতেই বসে যাবে)
 *******************************************************/
function getDealerSpreadsheet(dealerId) {
  const ss = SpreadsheetApp.openById(getDealerSpreadsheetId(dealerId));
  ensureDealerSheetHeadersUpToDate(ss, dealerId);
  return ss;
}

/*******************************************************
 * প্রতিটি ডিলারের নিজস্ব Spreadsheet-এর হেডার বর্তমান
 * DEALER_SHEETS_DEF এর সাথে মিলে কিনা যাচাই করে, না মিললে
 * স্বয়ংক্রিয়ভাবে নতুন কলাম যোগ করে দেয় (পুরনো ডাটা অক্ষত থাকে)।
 * প্রতিটি API কলে বারবার চেক না করে ১ ঘণ্টার জন্য ক্যাশ করা হয়,
 * যাতে পারফরম্যান্সে প্রভাব না পড়ে।
 *******************************************************/
function ensureDealerSheetHeadersUpToDate(ss, dealerId) {
  const cache = CacheService.getScriptCache();
  const cacheKey = "headers_synced_v3_" + dealerId; // v3: "IncomeReceipt" ট্যাব যোগ হওয়ায় নতুন করে সিঙ্ক
  if (cache.get(cacheKey)) return; // সম্প্রতি চেক করা হয়েছে, আবার লাগবে না

  ensureSheetsWithHeaders(ss, DEALER_SHEETS_DEF);

  try { cache.put(cacheKey, "1", 21600); } catch (e) { /* বাদ */ }  // ৬ ঘণ্টা — ভারী হেডার-সিঙ্ক বারবার চলবে না
}

/*******************************************************
 * পারফরম্যান্স: dealerId → SpreadsheetID ম্যাপিং ক্যাশ করা হয়
 * (৬ ঘণ্টা — এটা কখনো বদলায় না) — প্রতিটি ডিলার-সাইড API কলে
 * পুরো Dealers শীট স্ক্যান করার বদলে ক্যাশ থেকে সরাসরি পাওয়া যায়
 *******************************************************/
function getDealerSpreadsheetId(dealerId) {
  const cache = CacheService.getScriptCache();
  const cacheKey = "dealer_ssid_" + dealerId;
  const cached = cache.get(cacheKey);
  if (cached) return cached;

  const masterSS = getMasterSS();
  const dealersSheet = getSheet(masterSS, "Dealers");
  const dealerInfo = getDealerRow(dealersSheet, dealerId);
  if (!dealerInfo) throw new Error("ডিলার পাওয়া যায়নি");

  try { cache.put(cacheKey, dealerInfo.spreadsheetId, 21600); } catch (e) { /* বাদ */ }
  return dealerInfo.spreadsheetId;
}

/*******************************************************
 * জেনেরিক: হেডার-ভিত্তিক রো যোগ (ফরমের ফিল্ড ছাড়া অতিরিক্ত কিছু যোগ হবে না)
 *******************************************************/
function isTextColumnHeader(h) {
  // মোবাইল/NID এর মতো কলামে শুরুর ০ রাখতে হয় — এগুলো সবসময় টেক্সট হিসেবে লেখা হয়
  return /মোবাইল|mobile|NID|জন্মসনদ/i.test(String(h));
}

function genericAddRow(sheet, rowObject) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const newRow = sheet.getLastRow() + 1;
  const textCells = [];
  const row = headers.map(function (h, i) {
    let v = rowObject.hasOwnProperty(h) ? rowObject[h] : "";
    if (isTextColumnHeader(h)) {
      textCells.push(sheet.getRange(newRow, i + 1).getA1Notation());
      if (v !== "" && v !== null && v !== undefined) v = String(v);
    }
    return v;
  });
  // পারফরম্যান্স: মোবাইল/NID এর মতো সব টেক্সট ঘর একটিমাত্র কলে "@" ফরম্যাট করা হয়
  // (আগে প্রতিটির জন্য আলাদা কল হতো), নাহলে শীট 01712.. কে সংখ্যা ধরে ০ কেটে দেয়
  if (textCells.length > 0) sheet.getRangeList(textCells).setNumberFormat("@");
  sheet.getRange(newRow, 1, 1, headers.length).setValues([row]);
  return newRow;
}

/*******************************************************
 * জেনেরিক: সব রো অবজেক্ট আকারে রিড (rowIndex সহ, এডিট/ডিলিটের জন্য দরকার)
 *******************************************************/
function genericListRows(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  const result = [];
  for (let i = 1; i < data.length; i++) {
    const obj = { rowIndex: i + 1 };
    headers.forEach(function (h, idx) {
      obj[h] = data[i][idx];
    });
    result.push(obj);
  }
  return result;
}

/*******************************************************
 * জেনেরিক: idColumn দিয়ে একটি রো খুঁজে বের করা
 *******************************************************/
function genericFindRowIndex(sheet, idColumnName, idValue) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return -1;
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const idx = headers.indexOf(idColumnName);
  if (idx === -1) return -1;
  // পারফরম্যান্স: পুরো শীটের বদলে শুধু ID কলামটি পড়া হয়
  const ids = sheet.getRange(2, idx + 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === idValue) return i + 2; // sheet row number
  }
  return -1;
}

/*******************************************************
 * জেনেরিক: রো আপডেট (শুধু ফরমের ফিল্ড, extra কিছু না)
 *******************************************************/
function genericUpdateRow(sheet, idColumnName, idValue, updatedFields) {
  const rowIndex = genericFindRowIndex(sheet, idColumnName, idValue);
  if (rowIndex === -1) return false;
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  // কোন কোন কলাম বদলাবে তা বের করা — প্রতিটি ঘরে আলাদা কল না করে
  // বদলানো কলামগুলোর সীমা (min..max) একবারে পড়ে একবারেই লেখা হয়
  let minCol = -1, maxCol = -1;
  headers.forEach(function (h, colIdx) {
    if (updatedFields.hasOwnProperty(h)) {
      if (minCol === -1) minCol = colIdx;
      maxCol = colIdx;
    }
  });
  if (minCol === -1) return true; // বদলানোর মতো কিছু নেই

  const width = maxCol - minCol + 1;
  const range = sheet.getRange(rowIndex, minCol + 1, 1, width);
  const values = range.getValues()[0];
  const formulas = range.getFormulas()[0];
  const formats = range.getNumberFormats()[0];
  let formatsChanged = false;

  const out = values.map(function (cur, k) {
    const h = headers[minCol + k];
    if (updatedFields.hasOwnProperty(h)) {
      let v = updatedFields[h];
      if (isTextColumnHeader(h)) {
        formats[k] = "@";
        formatsChanged = true;
        if (v !== "" && v !== null && v !== undefined) v = String(v);
      }
      return v;
    }
    return formulas[k] ? formulas[k] : cur; // অপরিবর্তিত ঘর যেমন ছিল তেমনই থাকে
  });

  if (formatsChanged) range.setNumberFormats([formats]);
  range.setValues([out]);
  return true;
}

/*******************************************************
 * জেনেরিক: রো ডিলিট
 *******************************************************/
function genericDeleteRow(sheet, idColumnName, idValue) {
  const rowIndex = genericFindRowIndex(sheet, idColumnName, idValue);
  if (rowIndex === -1) return false;
  sheet.deleteRow(rowIndex);
  return true;
}

/*=========================================================
 *  গ্রাহক (Customers) — Admin + প্রতিনিধি
 *=======================================================*/
function addCustomer(data) {
  const perm = checkPermission(data.token, ["Admin", "প্রতিনিধি"]);
  if (!perm.ok) return { success: false, message: perm.message };

  const ss = getDealerSpreadsheet(perm.payload.dealerId);
  const sheet = getSheet(ss, "Customers");

  // লক শুধু আইডি তৈরি + সারি লেখার ছোট অংশে — দুজন একসাথে সংরক্ষণ করলেও একই আইডি/সারি হবে না
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  let customerId;
  try {
    customerId = generateId(sheet, "C");
    genericAddRow(sheet, {
      "CustomerID": customerId,
      "তারিখ": new Date(),
      "নাম": data.নাম,
      "পিতার নাম": data.পিতারনাম,
      "মোবাইল নং": data.mobile,
      "NID/জন্মসনদ নং": data.nid,
      "বাড়ির নাম": data.বাড়িরনাম,
      "গ্রাম": data.গ্রাম,
      "ওয়ার্ড নং": data.ward,
      "ইউনিয়ন/পৌরসভা": data.union,
      "প্রাপ্তির স্থান": data.praptirsthan,
      "কার্ড ফি": data.cardFee || 0,
      "রেফারেন্স": data.reference || ""
    });
  } finally {
    lock.releaseLock();
  }
  // (মোবাইল ঘর টেক্সট ফরম্যাট genericAddRow নিজেই করে — আলাদা করে আবার করার দরকার নেই)

  return { success: true, customerId: customerId, message: "গ্রাহক যোগ হয়েছে" };
}

/*******************************************************
 * একসাথে অনেক গ্রাহক সংরক্ষণ (ব্যাকগ্রাউন্ড কিউ থেকে আসে)
 * শীট একবার খোলা, লক একবার, লেখা একবার — তাই ২০ জন গ্রাহকও
 * একজনের সমান সময়েই সংরক্ষিত হয়
 * data: { token, customers: [ {নাম, পিতারনাম, mobile, nid, ...}, ... ] }
 * ফেরত: ids — customers এর ক্রম অনুযায়ী নতুন আইডি
 *******************************************************/
function addCustomers(data) {
  const perm = checkPermission(data.token, ["Admin", "প্রতিনিধি"]);
  if (!perm.ok) return { success: false, message: perm.message };

  const list = (data.customers || []).slice(0, 50);
  if (list.length === 0) return { success: false, message: "সংরক্ষণের মতো কোনো গ্রাহক নেই" };

  const ss = getDealerSpreadsheet(perm.payload.dealerId);
  const sheet = getSheet(ss, "Customers");

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  const ids = [];
  try {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const lastRow = sheet.getLastRow();
    let max = 0;
    if (lastRow >= 2) {
      sheet.getRange(2, 1, lastRow - 1, 1).getValues().forEach(function (r) {
        const m = /^C(\d+)$/.exec(String(r[0]));
        if (m) max = Math.max(max, Number(m[1]));
      });
    }
    const now = new Date();
    const rows = list.map(function (d, i) {
      const id = "C" + Utilities.formatString("%04d", max + 1 + i);
      ids.push(id);
      const obj = {
        "CustomerID": id, "তারিখ": now,
        "নাম": d.নাম, "পিতার নাম": d.পিতারনাম, "মোবাইল নং": d.mobile,
        "NID/জন্মসনদ নং": d.nid, "বাড়ির নাম": d.বাড়িরনাম, "গ্রাম": d.গ্রাম,
        "ওয়ার্ড নং": d.ward, "ইউনিয়ন/পৌরসভা": d.union, "প্রাপ্তির স্থান": d.praptirsthan,
        "কার্ড ফি": d.cardFee || 0, "রেফারেন্স": d.reference || ""
      };
      return headers.map(function (h) {
        let v = obj.hasOwnProperty(h) ? obj[h] : "";
        if (isTextColumnHeader(h) && v !== "" && v !== null && v !== undefined) v = String(v);
        return v;
      });
    });

    const start = lastRow + 1;
    const textCols = [];
    headers.forEach(function (h, i) {
      if (isTextColumnHeader(h)) textCols.push(sheet.getRange(start, i + 1, rows.length, 1).getA1Notation());
    });
    if (textCols.length > 0) sheet.getRangeList(textCols).setNumberFormat("@");
    sheet.getRange(start, 1, rows.length, headers.length).setValues(rows);
  } finally {
    lock.releaseLock();
  }

  return { success: true, ids: ids, message: list.length + " জন গ্রাহক যোগ হয়েছে" };
}

function listCustomers(data) {
  const perm = checkPermission(data.token, ["Admin", "প্রতিনিধি"]);
  if (!perm.ok) return { success: false, message: perm.message };
  const ss = getDealerSpreadsheet(perm.payload.dealerId);
  const sheet = getSheet(ss, "Customers");
  return { success: true, customers: genericListRows(sheet) };
}

/*******************************************************
 * এজেন্সি থেকে যেকোনো নির্দিষ্ট ডিলারের গ্রাহক তালিকা দেখা
 * (গ্রাহক কার্ড তৈরির জন্য ব্যবহৃত) — শুধু মূল এজেন্সি করতে পারবে
 * data: { token, dealerId }
 *******************************************************/
function listCustomersForDealer(data) {
  const perm = checkAgencyPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };

  const ss = getDealerSpreadsheet(data.dealerId);
  const sheet = getSheet(ss, "Customers");
  return { success: true, customers: genericListRows(sheet) };
}

function updateCustomer(data) {
  const perm = checkPermission(data.token, ["Admin"]);
  if (!perm.ok) return { success: false, message: perm.message };
  const ss = getDealerSpreadsheet(perm.payload.dealerId);
  const sheet = getSheet(ss, "Customers");
  const ok = genericUpdateRow(sheet, "CustomerID", data.customerId, data.fields);
  return { success: ok, message: ok ? "আপডেট হয়েছে" : "গ্রাহক পাওয়া যায়নি" };
}

function deleteCustomer(data) {
  const perm = checkPermission(data.token, ["Admin"]);
  if (!perm.ok) return { success: false, message: perm.message };
  const ss = getDealerSpreadsheet(perm.payload.dealerId);
  const sheet = getSheet(ss, "Customers");
  // ডিলিট + আইডি পুনর্বিন্যাস একসাথে (নতুন গ্রাহক যোগের সাথে জট এড়াতে লকের ভেতরে)
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  let ok;
  try {
    ok = genericDeleteRow(sheet, "CustomerID", data.customerId);
    // বাকি গ্রাহকদের আইডি ১ থেকে ক্রমানুসারে নতুন করে সাজানো (বিক্রির রেকর্ডের আইডিও সাথে মিলিয়ে)
    if (ok) renumberCustomerIds(ss, sheet);
  } finally {
    lock.releaseLock();
  }
  if (!ok) return { success: false, message: "গ্রাহক পাওয়া যায়নি" };
  return {
    success: true,
    message: "ডিলিট হয়েছে — গ্রাহক আইডি ক্রমানুসারে সাজানো হয়েছে",
    customers: genericListRows(sheet)      // ব্রাউজার আবার লোড না করে সরাসরি এটাই দেখাবে
  };
}

/*******************************************************
 * গ্রাহক আইডি ধারাবাহিক করা: সারির ক্রম অনুযায়ী C0001, C0002, ... (ডিলিটে ফাঁক পড়লে
 * পরের সবার আইডি এক ধাপ করে কমে যায়)।
 *
 * বিক্রি (Sales) শীটের CustomerID পুরনো → নতুন আইডিতে বদলে দেওয়া হয়, যাতে প্রতিটি
 * বিক্রি আগের গ্রাহকের সাথেই যুক্ত থাকে। যে বিক্রির গ্রাহক আর নেই (ডিলিট হয়েছে) সেই
 * রেকর্ড মুছে ফেলা হয় না (রিপোর্ট ও স্টকের হিসাব ঠিক থাকে) — শুধু আইডি "DELETED" করে
 * দেওয়া হয়, যাতে ওই পুরনো আইডি পরে অন্য কোনো গ্রাহকের সাথে ভুলে জুড়ে না যায়
 *******************************************************/
const DELETED_CUSTOMER_ID = "DELETED";

function renumberCustomerIds(ss, sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return;
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const idx = headers.indexOf("CustomerID");
  if (idx === -1) return;

  const n = Math.max(lastRow - 1, 0);
  const oldIds = n > 0 ? sheet.getRange(2, idx + 1, n, 1).getValues().map(function (r) { return String(r[0]); }) : [];

  const map = {};          // পুরনো আইডি → নতুন আইডি (একই আইডি দুইবার থাকলে প্রথমটির)
  let changed = false;
  const newIds = oldIds.map(function (oldId, i) {
    const nid = "C" + Utilities.formatString("%04d", i + 1);
    if (!map.hasOwnProperty(oldId)) map[oldId] = nid;
    if (oldId !== nid) changed = true;
    return [nid];
  });
  if (changed) sheet.getRange(2, idx + 1, n, 1).setValues(newIds);

  // বিক্রি শীটের CustomerID মেলানো
  const sales = ss.getSheetByName("Sales");
  if (!sales || sales.getLastRow() < 2) return;
  const sHeaders = sales.getRange(1, 1, 1, sales.getLastColumn()).getValues()[0];
  const sIdx = sHeaders.indexOf("CustomerID");
  if (sIdx === -1) return;
  const m = sales.getLastRow() - 1;
  const col = sales.getRange(2, sIdx + 1, m, 1).getValues();
  let sChanged = false;
  const out = col.map(function (r) {
    const v = String(r[0]);
    if (v === "") return [r[0]];
    const nv = map.hasOwnProperty(v) ? map[v] : DELETED_CUSTOMER_ID;
    if (nv !== v) sChanged = true;
    return [nv];
  });
  if (sChanged) sales.getRange(2, sIdx + 1, m, 1).setValues(out);
}

/*=========================================================
 *  প্যাকেজ — এক রো = এক প্যাকেজ, পণ্য কলাম-ভিত্তিক (একবারেই এক লেখা, দ্রুততম)
 *  — যোগ/এডিট/ডিলিট শুধু মূল এজেন্সি (AGENCY) করতে পারবে
 *  — ডিলার (Admin/প্রতিনিধি) শুধু লিস্ট/ভিউ করতে পারবে
 *=======================================================*/
/*******************************************************
 * Packages ট্যাব — আলাদা "প্যাকেজ নং" কলাম আর নেই; প্যাকেজের "নাম"-ই
 * নম্বর বহন করে (যেমন: "প্যাকেজ নং- ০১")
 *******************************************************/
function getPackagesSheet(masterSS) {
  return getSheet(masterSS, "Packages");
}

/*******************************************************
 * প্যাকেজের ক্রমিক নং — নামের ভেতরের প্রথম সংখ্যা (বাংলা বা ইংরেজি অঙ্ক,
 * যেমন "প্যাকেজ নং- ০১" → ১)। নামে সংখ্যা না থাকলে সবার শেষে (এন্ট্রির
 * ক্রম অনুযায়ী) বসবে। এই নং অনুযায়ীই সব জায়গায় সিরিয়াল সাজানো হয়
 *******************************************************/
function packageNoOf(p) {
  const name = String(p["নাম"] || "");
  const bnDigits = "০১২৩৪৫৬৭৮৯";
  const normalized = name.replace(/[০-৯]/g, function (d) { return String(bnDigits.indexOf(d)); });
  const m = normalized.match(/\d+/);
  if (m) return Number(m[0]);
  return 100000 + (p.rowIndex || 2);
}

function addPackage(data) {
  const perm = checkAgencyPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };

  const masterSS = getMasterSS();
  const pkgSheet = getPackagesSheet(masterSS);
  const packageId = generateId(pkgSheet, "P");

  const items = (data.items || []).slice(0, MAX_PACKAGE_ITEMS);
  let totalBazar = 0, totalCombo = 0, totalSashroy = 0;
  const row = { "PackageID": packageId, "নাম": data.নাম, "তারিখ": new Date(), "ধরন": data.ধরন };

  for (let i = 0; i < MAX_PACKAGE_ITEMS; i++) {
    const n = i + 1;
    const item = items[i];
    if (item) {
      const bazar = Number(item.bazarMulyo) || 0;
      const combo = Number(item.comboMulyo) || 0;
      const sashroy = bazar - combo;
      totalBazar += bazar; totalCombo += combo; totalSashroy += sashroy;
      row["পণ্য" + n + " - নাম ও পরিমাণ"] = item.naamPoriman || "";
      row["পণ্য" + n + " - বাজার মূল্য"] = bazar;
      row["পণ্য" + n + " - কম্বো মূল্য"] = combo;
      row["পণ্য" + n + " - সাশ্রয়"] = sashroy;
    } else {
      row["পণ্য" + n + " - নাম ও পরিমাণ"] = "";
      row["পণ্য" + n + " - বাজার মূল্য"] = "";
      row["পণ্য" + n + " - কম্বো মূল্য"] = "";
      row["পণ্য" + n + " - সাশ্রয়"] = "";
    }
  }
  row["সর্বমোট বাজার মূল্য"] = totalBazar;
  row["সর্বমোট কম্বো মূল্য (সর্বমোট মূল্য)"] = totalCombo;
  row["সর্বমোট সাশ্রয়"] = totalSashroy;

  // একটিমাত্র রো, একটিমাত্র লেখার কল — সবচেয়ে দ্রুত পদ্ধতি
  genericAddRow(pkgSheet, row);

  invalidatePackageCache();
  return { success: true, packageId: packageId, total: totalCombo, message: "প্যাকেজ যোগ হয়েছে" };
}

/*******************************************************
 * প্যাকেজ লিস্ট ক্যাশ করা হয় (পরিবর্তনে নিজে থেকে মুছে যায়)
 *******************************************************/
function invalidatePackageCache() {
  CacheService.getScriptCache().remove("packages_cache_v5");
}

/*******************************************************
 * একটি প্যাকেজ-রো কে কলাম-ভিত্তিক ফরম্যাট থেকে items[] আকারে রূপান্তর
 * (ফ্রন্টএন্ডের জন্য সুবিধাজনক গঠন)
 *******************************************************/
function expandPackageRow(pkg) {
  const items = [];
  for (let i = 1; i <= MAX_PACKAGE_ITEMS; i++) {
    const naam = pkg["পণ্য" + i + " - নাম ও পরিমাণ"];
    if (naam) {
      items.push({
        "পণ্যের নাম ও পরিমাণ": naam,
        "বাজার মূল্য": pkg["পণ্য" + i + " - বাজার মূল্য"],
        "কম্বো মূল্য": pkg["পণ্য" + i + " - কম্বো মূল্য"],
        "সাশ্রয়": pkg["পণ্য" + i + " - সাশ্রয়"]
      });
    }
  }
  pkg.items = items;
  pkg.itemCount = items.length;
  pkg["সর্বমোট মূল্য"] = pkg["সর্বমোট কম্বো মূল্য (সর্বমোট মূল্য)"];
  return pkg;
}

function listPackages(data) {
  const perm = checkPermission(data.token, ["Admin", "প্রতিনিধি"]);
  if (!perm.ok) return { success: false, message: perm.message };

  const cache = CacheService.getScriptCache();
  const cached = cache.get("packages_cache_v5");
  if (cached) {
    return { success: true, packages: JSON.parse(cached) };
  }

  const masterSS = getMasterSS();
  const pkgSheet = getPackagesSheet(masterSS);
  // প্যাকেজ নং অনুযায়ী ছোট থেকে বড় — আগে-পরে এন্ট্রি হলেও সিরিয়াল ঠিক থাকবে
  const packages = genericListRows(pkgSheet).map(function (p) {
    p.serial = packageNoOf(p); // শীটের কলাম নয় — নাম থেকে বের করা সিরিয়াল (সাজানোর জন্য)
    return expandPackageRow(p);
  }).sort(function (a, b) {
    return (a.serial - b.serial) || (a.rowIndex - b.rowIndex);
  });

  try {
    cache.put("packages_cache_v5", JSON.stringify(packages), 300); // ৫ মিনিট (যোগ/এডিট/ডিলিটে নিজে থেকে মুছে যায়)
  } catch (e) { /* ক্যাশ সাইজ বেশি হলে চুপচাপ বাদ */ }

  return { success: true, packages: packages };
}

function updatePackage(data) {
  const perm = checkAgencyPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };
  const masterSS = getMasterSS();
  const sheet = getPackagesSheet(masterSS);

  // fields এ items[] থাকলে সম্পূর্ণ রো নতুন করে হিসাব করে বসানো হয়
  let fields = data.fields || {};
  if (data.items) {
    const items = data.items.slice(0, MAX_PACKAGE_ITEMS);
    let totalBazar = 0, totalCombo = 0, totalSashroy = 0;
    for (let i = 0; i < MAX_PACKAGE_ITEMS; i++) {
      const n = i + 1;
      const item = items[i];
      if (item) {
        const bazar = Number(item.bazarMulyo) || 0;
        const combo = Number(item.comboMulyo) || 0;
        const sashroy = bazar - combo;
        totalBazar += bazar; totalCombo += combo; totalSashroy += sashroy;
        fields["পণ্য" + n + " - নাম ও পরিমাণ"] = item.naamPoriman || "";
        fields["পণ্য" + n + " - বাজার মূল্য"] = bazar;
        fields["পণ্য" + n + " - কম্বো মূল্য"] = combo;
        fields["পণ্য" + n + " - সাশ্রয়"] = sashroy;
      } else {
        fields["পণ্য" + n + " - নাম ও পরিমাণ"] = "";
        fields["পণ্য" + n + " - বাজার মূল্য"] = "";
        fields["পণ্য" + n + " - কম্বো মূল্য"] = "";
        fields["পণ্য" + n + " - সাশ্রয়"] = "";
      }
    }
    fields["সর্বমোট বাজার মূল্য"] = totalBazar;
    fields["সর্বমোট কম্বো মূল্য (সর্বমোট মূল্য)"] = totalCombo;
    fields["সর্বমোট সাশ্রয়"] = totalSashroy;
  }

  const ok = genericUpdateRow(sheet, "PackageID", data.packageId, fields);
  invalidatePackageCache();
  return { success: ok, message: ok ? "আপডেট হয়েছে" : "প্যাকেজ পাওয়া যায়নি" };
}

function deletePackage(data) {
  const perm = checkAgencyPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };
  const masterSS = getMasterSS();
  const pkgSheet = getPackagesSheet(masterSS);
  const ok = genericDeleteRow(pkgSheet, "PackageID", data.packageId);
  invalidatePackageCache();
  return { success: ok, message: ok ? "প্যাকেজ ডিলিট হয়েছে" : "প্যাকেজ পাওয়া যায়নি" };
}

/*=========================================================
 *  বিক্রি (Sales) — Admin + প্রতিনিধি
 *=======================================================*/
function addSale(data) {
  const perm = checkPermission(data.token, ["Admin", "প্রতিনিধি"]);
  if (!perm.ok) return { success: false, message: perm.message };

  const ss = getDealerSpreadsheet(perm.payload.dealerId);
  const sheet = getSheet(ss, "Sales");
  const saleId = generateId(sheet, "S");

  genericAddRow(sheet, {
    "SaleID": saleId,
    "CustomerID": data.customerId,
    "PackageID": data.packageId,
    "তারিখ": new Date(),
    "মূল্য": data.price || 0,
    "স্ট্যাটাস": "বিক্রিত"
  });

  return { success: true, saleId: saleId, message: "বিক্রি সম্পন্ন হয়েছে" };
}

/*******************************************************
 * বিক্রির মূল্য পরিবর্তন (প্যাকেজের ডিফল্ট মূল্য থেকে বদলাতে চাইলে)
 *******************************************************/
function updateSalePrice(data) {
  const perm = checkPermission(data.token, ["Admin", "প্রতিনিধি"]);
  if (!perm.ok) return { success: false, message: perm.message };

  const ss = getDealerSpreadsheet(perm.payload.dealerId);
  const sheet = getSheet(ss, "Sales");
  const ok = genericUpdateRow(sheet, "SaleID", data.saleId, { "মূল্য": data.price });

  return { success: ok, message: ok ? "মূল্য আপডেট হয়েছে" : "বিক্রি রেকর্ড পাওয়া যায়নি" };
}

function cancelSale(data) {
  const perm = checkPermission(data.token, ["Admin", "প্রতিনিধি"]);
  if (!perm.ok) return { success: false, message: perm.message };

  const ss = getDealerSpreadsheet(perm.payload.dealerId);
  const sheet = getSheet(ss, "Sales");
  const ok = genericUpdateRow(sheet, "SaleID", data.saleId, { "স্ট্যাটাস": "বাতিল" });

  return { success: ok, message: ok ? "বিক্রি বাতিল হয়েছে" : "বিক্রি রেকর্ড পাওয়া যায়নি" };
}

function listSales(data) {
  const perm = checkPermission(data.token, ["Admin", "প্রতিনিধি"]);
  if (!perm.ok) return { success: false, message: perm.message };
  const ss = getDealerSpreadsheet(perm.payload.dealerId);
  const sheet = getSheet(ss, "Sales");
  return { success: true, sales: genericListRows(sheet) };
}

/*******************************************************
 * পারফরম্যান্স: বিক্রি পেজের জন্য প্যাকেজ+গ্রাহক+বিক্রি — তিনটি আলাদা
 * কল না করে একবারেই সব ডাটা ফেরত দেওয়া (নেটওয়ার্ক রাউন্ড-ট্রিপ কমাতে)
 *******************************************************/
function getSalesPageData(data) {
  const perm = checkPermission(data.token, ["Admin", "প্রতিনিধি"]);
  if (!perm.ok) return { success: false, message: perm.message };

  const ss = getDealerSpreadsheet(perm.payload.dealerId);
  const pkgResult = listPackages(data); // ক্যাশড থাকলে দ্রুত
  const customers = genericListRows(getSheet(ss, "Customers"));
  const sales = genericListRows(getSheet(ss, "Sales"));

  return {
    success: true,
    packages: pkgResult.success ? pkgResult.packages : [],
    customers: customers,
    sales: sales
  };
}

/*=========================================================
 *  ড্যাশবোর্ড সামারি — Admin + প্রতিনিধি
 *=======================================================*/
function getDashboardSummary(data) {
  const perm = checkPermission(data.token, ["Admin", "প্রতিনিধি"]);
  if (!perm.ok) return { success: false, message: perm.message };

  const ss = getDealerSpreadsheet(perm.payload.dealerId);
  const masterSS = getMasterSS();
  const packages = genericListRows(getSheet(masterSS, "Packages"));
  const customers = genericListRows(getSheet(ss, "Customers"));
  const sales = genericListRows(getSheet(ss, "Sales")).filter(function (s) { return s["স্ট্যাটাস"] === "বিক্রিত"; });

  // অর্ডার ইনভয়েস এখন শেয়ার্ড মাস্টার SalesInvoice শীটে থাকে (ইনভয়েস নং
  // অনুযায়ী ডিডুপ করে গোনা হচ্ছে, নিজের DealerID এর সীমার মধ্যে)
  const invoiceLines = genericListRows(getSheet(masterSS, "SalesInvoice"))
    .filter(function (e) { return e["DealerID"] === perm.payload.dealerId; });
  const seenInvoiceNos = {};
  invoiceLines.forEach(function (e) { seenInvoiceNos[e["ইনভয়েস নং"]] = true; });
  const totalOrders = Object.keys(seenInvoiceNos).length;

  const runningPackages = packages.filter(function (p) { return p["ধরন"] === "এক্টিভ"; });

  // মোট কার্ড ফি — সব গ্রাহকের "কার্ড ফি" এর যোগফল
  let totalCardFee = 0;
  customers.forEach(function (c) { totalCardFee += Number(c["কার্ড ফি"]) || 0; });

  return {
    success: true,
    summary: {
      totalPackages: packages.length,
      runningPackages: runningPackages.length,
      totalCustomers: customers.length,
      totalCardFee: totalCardFee,
      totalOrders: totalOrders,
      totalSales: sales.length
    }
  };
}

/*=========================================================
 *  খরচ ভাউচার (ExpenseVoucher) — ডিলারের নিজস্ব খরচ, Admin + প্রতিনিধি
 *=======================================================*/
function addExpenseVoucher(data) {
  const perm = checkPermission(data.token, ["Admin", "প্রতিনিধি"]);
  if (!perm.ok) return { success: false, message: perm.message };

  const ss = getDealerSpreadsheet(perm.payload.dealerId);
  const sheet = getSheet(ss, "ExpenseVoucher");

  const entries = (data.entries || []).map(function (entry) {
    return {
      "তারিখ": entry.তারিখ ? new Date(entry.তারিখ) : new Date(),
      "বিবরণ": entry.বিবরণ,
      "পরিমাণ": entry.পরিমাণ
    };
  });

  const addedIds = batchAppendRows(sheet, entries, "EV", "VoucherID");

  return { success: true, voucherIds: addedIds, message: addedIds.length + " টি খরচ এন্ট্রি যোগ হয়েছে" };
}

function listExpenseVouchers(data) {
  const perm = checkPermission(data.token, ["Admin", "প্রতিনিধি"]);
  if (!perm.ok) return { success: false, message: perm.message };

  const ss = getDealerSpreadsheet(perm.payload.dealerId);
  const sheet = getSheet(ss, "ExpenseVoucher");
  return { success: true, vouchers: genericListRows(sheet) };
}

function updateExpenseVoucher(data) {
  const perm = checkPermission(data.token, ["Admin"]);
  if (!perm.ok) return { success: false, message: perm.message };

  const ss = getDealerSpreadsheet(perm.payload.dealerId);
  const sheet = getSheet(ss, "ExpenseVoucher");
  const ok = genericUpdateRow(sheet, "VoucherID", data.voucherId, data.fields || {});
  return { success: ok, message: ok ? "আপডেট হয়েছে" : "খরচ রেকর্ড পাওয়া যায়নি" };
}

function deleteExpenseVoucher(data) {
  const perm = checkPermission(data.token, ["Admin"]);
  if (!perm.ok) return { success: false, message: perm.message };

  const ss = getDealerSpreadsheet(perm.payload.dealerId);
  const sheet = getSheet(ss, "ExpenseVoucher");
  const ok = genericDeleteRow(sheet, "VoucherID", data.voucherId);
  return { success: ok, message: ok ? "ডিলিট হয়েছে" : "খরচ রেকর্ড পাওয়া যায়নি" };
}

/*=========================================================
 *  এজেন্সি ড্যাশবোর্ড সামারি — মোট ডিলার, ব্যবহারকারী, টাকা, কমিশন, খরচ
 *=======================================================*/
function getAgencyDashboardSummary(data) {
  const perm = checkAgencyPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };

  const masterSS = getMasterSS();
  const dealers = genericListRows(getSheet(masterSS, "Dealers"));
  const users = genericListRows(getSheet(masterSS, "Users"));
  const commissions = genericListRows(getSheet(masterSS, "Commission"));
  const expenses = genericListRows(getSheet(masterSS, "Expense"));

  let totalCommission = 0;
  commissions.forEach(function (c) { totalCommission += Number(c["টাকার পরিমাণ"]) || 0; });

  let totalExpense = 0;
  expenses.forEach(function (e) { totalExpense += Number(e["টাকা"]) || 0; });

  return {
    success: true,
    summary: {
      totalDealers: dealers.length,
      totalUsers: users.length,
      totalCommission: totalCommission,
      totalExpense: totalExpense,
      totalMoney: totalCommission // "মোট টাকা" — এজেন্সি থেকে ডিলারদের দেওয়া মোট কমিশন/টাকার সমষ্টি
    }
  };
}

/*******************************************************
 * QR কোড স্ক্যান করলে যা দেখাবে — গ্রাহক হওয়ার তারিখ, মোট কতবার
 * ক্রয় করেছে, সর্বশেষ ক্রয়ের তারিখ (কোনো টোকেন লাগে না, পাবলিক ভিউ)
 *******************************************************/
function getCustomerCardInfo(dealerId, customerId) {
  try {
    const ss = getDealerSpreadsheet(dealerId);
    const customers = genericListRows(getSheet(ss, "Customers"));
    const customer = customers.find(function (c) { return c["CustomerID"] === customerId; });
    if (!customer) return { success: false, message: "গ্রাহক পাওয়া যায়নি" };

    const sales = genericListRows(getSheet(ss, "Sales")).filter(function (s) {
      return s["CustomerID"] === customerId && s["স্ট্যাটাস"] === "বিক্রিত";
    });

    let lastPurchaseDate = null;
    sales.forEach(function (s) {
      const d = (s["তারিখ"] instanceof Date) ? s["তারিখ"] : new Date(s["তারিখ"]);
      if (!lastPurchaseDate || d > lastPurchaseDate) lastPurchaseDate = d;
    });

    return {
      success: true,
      customerName: customer["নাম"],
      joinDate: customer["তারিখ"],
      purchaseCount: sales.length,
      lastPurchaseDate: lastPurchaseDate
    };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}
