/*******************************************************
 * TCB-স্টাইল ডিলার ম্যানেজমেন্ট সফটওয়্যার
 * Master Spreadsheet এর Apps Script (Code.gs)
 *
 * এই স্ক্রিপ্ট Master Spreadsheet-এ বসাতে হবে।
 * প্রথমবার setupMasterSheet() ফাংশনটি একবার ম্যানুয়ালি রান করলে
 * সকল ট্যাব ও কলাম হেডার নিজে থেকেই তৈরি হয়ে যাবে।
 *******************************************************/

// ==== কনফিগারেশন ====
// এই স্ক্রিপ্ট Master Spreadsheet-এর সাথে বাউন্ড থাকলে নিচের লাইন পরিবর্তনের দরকার নেই।
// যদি স্ট্যান্ডঅ্যালোন স্ক্রিপ্ট হিসেবে রাখেন, তাহলে MASTER_SS_ID বসিয়ে দিন।
const MASTER_SS_ID = ""; // খালি রাখলে বাউন্ড স্প্রেডশিট ব্যবহার হবে
const TEMPLATE_SS_ID = "1NSy3NzQeqYUoo-ov7sKBlNww7-PfOJIXsLo8r5H_c4c"; // DealerTemplate বানানোর পর তার ID এখানে বসান

// Web App এর নিজের Deployment URL — একবার Deploy করার পর এখানে বসিয়ে রাখলে
// বারবার আলাদা করে সংরক্ষণ/শেয়ার করার দরকার নেই, কোডেই থেকে যাবে।
const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbxmseU_Vj8xwWzq9rxK6b8NN5ZPvb3GNlc-LNN-hgwa2kr3kS6lAe11ZWseK7qd9PuB-g/exec";

function getMasterSS() {
  if (MASTER_SS_ID && MASTER_SS_ID.trim() !== "") {
    return SpreadsheetApp.openById(MASTER_SS_ID);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet(ss, tabName) {
  return ss.getSheetByName(tabName);
}

/*******************************************************
 * প্যাকেজ শীটের হেডার তৈরি — এক রো তে একটি প্যাকেজ, প্রতিটি পণ্য
 * নিজের নিজের কলামে বসে (সর্বোচ্চ ১০টি পণ্য স্লট)
 * এটি MASTER_SHEETS_DEF এর উপরে থাকা আবশ্যক — নাহলে
 * "Cannot access before initialization" এরর আসবে
 *******************************************************/
const MAX_PACKAGE_ITEMS = 10;
function buildPackageHeaders() {
  const headers = ["PackageID", "নাম", "তারিখ", "ধরন"];
  for (let i = 1; i <= MAX_PACKAGE_ITEMS; i++) {
    headers.push(
      "পণ্য" + i + " - নাম ও পরিমাণ",
      "পণ্য" + i + " - বাজার মূল্য",
      "পণ্য" + i + " - কম্বো মূল্য",
      "পণ্য" + i + " - সাশ্রয়"
    );
  }
  headers.push("সর্বমোট বাজার মূল্য", "সর্বমোট কম্বো মূল্য (সর্বমোট মূল্য)", "সর্বমোট সাশ্রয়");
  return headers;
}

/*******************************************************
 * আয় রশিদ — ডিপু (Master) ও প্রতিটি ডিলারের শীট, দুই জায়গাতেই একই হেডার।
 * MASTER_SHEETS_DEF / DEALER_SHEETS_DEF এর উপরে থাকা আবশ্যক
 *******************************************************/
const INCOME_HEADERS = ["EntryID", "রশিদ নং", "তারিখ", "ক্রম", "বিবরণ", "টাকা", "সর্বমোট"];

/*******************************************************
 * Master Spreadsheet এর ট্যাব ও হেডার সংজ্ঞা
 *******************************************************/
const MASTER_SHEETS_DEF = {
  "Dealers": [
    "DealerID", "তারিখ", "নাম", "পিতার নাম", "NID/জন্মসনদ",
    "মোবাইল", "Gmail", "ট্রেড লাইসেন্স নং", "ঠিকানা", "ডিলারের ছবি(URL)",
    "SpreadsheetID", "স্ট্যাটাস", "Facebook Link", "এরিয়া",
    "ইউনিয়ন/পৌরসভা", "উপজেলা", "জেলা"
  ],
  "DealerNominee": [
    "NomineeID", "DealerID", "নমিনির নাম", "NID নং", "মোবাইল নং", "সম্পর্ক", "নমিনির ছবি(URL)"
  ],
  "Users": [
    "UserID", "DealerID", "ইউজারনেম", "পাসওয়ার্ড", "রোল", "নাম", "মোবাইল"
  ],
  "Agency": [
    "নাম", "মোবাইল", "লোগো(URL)",
    "ব্যবস্থাপক তালিকা (JSON)",
    "Facebook Link", "Youtube Link"
  ],
  // সংস্থার তথ্য — শুধু এজেন্সি এন্ট্রি করবে, সব ডিলার একই তথ্য দেখবে
  "AboutInfo": [
    "প্রতিষ্ঠাকাল", "উদ্দেশ্য", "প্রধান কার্যালয়ের ঠিকানা",
    "জেলা ডিপুর ঠিকানা", "নিয়মাবলি", "সংক্ষিপ্ত বিবরণ"
  ],
  // প্যাকেজ — এক রো = এক প্যাকেজ, পণ্য সর্বোচ্চ ১০টি পর্যন্ত কলাম-ভিত্তিক
  // (নাম১/বাজার১/কম্বো১/সাশ্রয়১, নাম২/বাজার২/কম্বো২/সাশ্রয়২, ...)
  "Packages": buildPackageHeaders(),
  // কমিশন — এজেন্সি থেকে প্রতিটি ডিলারকে দেওয়া কমিশনের হিসাব
  "Commission": [
    "CommissionID", "DealerID", "ডিলারের নাম", "মোবাইল নং", "ঠিকানা", "টাকার পরিমাণ", "তারিখ"
  ],
  // আয় রশিদ — এজেন্সির নিজস্ব আয়ের হিসাব (রশিদ নং শেয়ার করে একাধিক সারি)
  "IncomeReceipt": INCOME_HEADERS,
  // খরচ — এজেন্সির নিজস্ব খরচের হিসাব (ভাউচার নং এর মতো একাধিক সারি শেয়ার করে)
  "Expense": [
    "EntryID", "ভাউচার নং", "তারিখ", "ক্রম", "বিবরণ", "টাকা", "সর্বমোট", "পরিশোধ", "বকেয়া"
  ],
  // স্টক: পণ্য তালিকা
  "Products": [
    "ProductID", "পণ্যের নাম", "ব্র্যান্ড", "বাজার মূল্য", "কম্বো মূল্য", "সাশ্রয়ী"
  ],
  // স্টক ইন ভাউচার — এক ভাউচারে একাধিক পণ্য (ইনভয়েস নং এর মতো ভাউচার নং শেয়ার করে)
  // কোম্পানি (সাপ্লায়ার) থেকে ক্রয়ের পরিশোধ/বকেয়া ট্র্যাক করার জন্য পরিশোধ/বকেয়া যোগ হয়েছে
  "StockInVoucher": [
    "EntryID", "ভাউচার নং", "তারিখ", "ProductID", "পণ্যের নাম",
    "বাজার মূল্য", "কম্বো মূল্য", "সাশ্রয়ী", "সংখ্যা", "মোট মূল্য",
    "সর্বমোট", "পরিশোধ", "বকেয়া"
  ],
  // বিক্রয় ইনভয়েস — এজেন্সি থেকে ডিলারকে প্যাকেজ বিক্রি (স্ট্যাটাস: পেন্ডিং/ডেলিভারি সম্পন্ন)
  "SalesInvoice": [
    "EntryID", "ইনভয়েস নং", "তারিখ", "DealerID", "ডিলার নাম", "মোবাইল", "ঠিকানা",
    "PackageID", "প্যাকেজ", "একক মূল্য", "সংখ্যা", "মোট মূল্য",
    "কমিশন %", "কমিশন মূল্য", "পরিশোধযোগ্য মূল্য",
    "সাবটোটাল", "ডিসকাউন্ট", "পরিশোধ", "বকেয়া", "স্ট্যাটাস"
  ]
};

/*******************************************************
 * প্রতিটি ডিলার Spreadsheet (টেমপ্লেট) এর ট্যাব ও হেডার সংজ্ঞা
 * (প্যাকেজ এখানে নেই — এজেন্সি থেকে আসে, Master Spreadsheet এ থাকে)
 *******************************************************/
const DEALER_SHEETS_DEF = {
  "Customers": [
    "CustomerID", "তারিখ", "নাম", "পিতার নাম", "মোবাইল নং", "NID/জন্মসনদ নং",
    "বাড়ির নাম", "গ্রাম", "ওয়ার্ড নং", "ইউনিয়ন/পৌরসভা", "প্রাপ্তির স্থান", "কার্ড ফি",
    // সর্বশেষ মেসেজ কার কাছে গেছে (মেসেজ পাঠালে নিজে থেকে আপডেট হয়)
    "WhatsApp", "SMS", "মেসেজ আইডি",
    // রেফারেন্স — শীটের শেষে যোগ হয় (বিদ্যমান কলামের জায়গা বদলায় না); ফরম/তালিকায় কার্ড ফি এর পরে দেখায়
    "রেফারেন্স"
  ],
  "Sales": [
    "SaleID", "CustomerID", "PackageID", "তারিখ", "মূল্য", "স্ট্যাটাস"
  ],
  // Orders (deprecated) — dealer-side "order" tracking has been replaced
  // by the shared master-level SalesInvoice system (same as ডিপু's
  // "অর্ডার ইনভয়েস"), usable by both agency and dealers.
  // খরচ ভাউচার — প্রতিটি ডিলারের নিজস্ব খরচের হিসাব
  "ExpenseVoucher": [
    "VoucherID", "তারিখ", "বিবরণ", "পরিমাণ"
  ],
  // আয় রশিদ — প্রতিটি ডিলারের নিজস্ব আয়ের হিসাব
  "IncomeReceipt": INCOME_HEADERS,
  // পাঠানো মেসেজের তালিকা (গ্রাহক → মেসেজ)
  // স্ট্যাটাস: "সকল গ্রাহক" অথবা প্রাপ্তির স্থানের নাম/ওয়ার্ড নং
  "Messages": [
    "MessageID", "মেসেজ", "তারিখ ও সময়", "ধরন", "স্ট্যাটাস", "মোট প্রাপক"
  ]
};

/*******************************************************
 * ফাংশন: একটি স্প্রেডশিটে ট্যাব ও হেডার তৈরি (না থাকলে)
 *******************************************************/
function ensureSheetsWithHeaders(ss, sheetsDef) {
  const definedNames = Object.keys(sheetsDef);

  definedNames.forEach(function (tabName) {
    let sheet = ss.getSheetByName(tabName);
    if (!sheet) {
      sheet = ss.insertSheet(tabName);
    }
    const headers = sheetsDef[tabName];
    const existingHeaderRange = sheet.getRange(1, 1, 1, headers.length);
    const existingValues = existingHeaderRange.getValues()[0];

    // হেডার না থাকলে বা ভিন্ন হলে বসিয়ে দেওয়া
    let needsHeader = false;
    for (let i = 0; i < headers.length; i++) {
      if (existingValues[i] !== headers[i]) {
        needsHeader = true;
        break;
      }
    }
    if (needsHeader) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
      sheet.setFrozenRows(1);
    }
  });

  // ডিফল্ট "Sheet1" থাকলে এবং খালি হলে মুছে ফেলা
  const defaultSheet = ss.getSheetByName("Sheet1");
  if (defaultSheet && ss.getSheets().length > 1) {
    const isEmpty = defaultSheet.getLastRow() === 0;
    if (isEmpty) {
      ss.deleteSheet(defaultSheet);
    }
  }
}

/*******************************************************
 * এই ফাংশনটি Master Spreadsheet-এ একবার ম্যানুয়ালি রান করুন
 * (Apps Script এডিটরে ফাংশন সিলেক্ট করে ▶ Run চাপুন)
 * এতে Dealers, DealerNominee, Users, Agency ট্যাব ও হেডার
 * স্বয়ংক্রিয়ভাবে তৈরি হয়ে যাবে।
 *******************************************************/
function setupMasterSheet() {
  const ss = getMasterSS();
  ensureSheetsWithHeaders(ss, MASTER_SHEETS_DEF);

  // মোবাইল নং কলামগুলো টেক্সট ফরম্যাট করে রাখা
  formatMobileColumnAsText(ss, "Dealers", "মোবাইল");
  formatMobileColumnAsText(ss, "DealerNominee", "মোবাইল নং");
  formatMobileColumnAsText(ss, "Users", "মোবাইল");
  formatMobileColumnAsText(ss, "Commission", "মোবাইল নং");

  SpreadsheetApp.getUi().alert("Master Spreadsheet সেটআপ সম্পন্ন হয়েছে। সকল ট্যাব ও হেডার তৈরি হয়েছে।");
}

/*******************************************************
 * এই ফাংশনটি একটি খালি Spreadsheet-এ একবার রান করলে সেটি
 * ডিলার-টেমপ্লেট হিসেবে প্রস্তুত হয়ে যাবে (সকল ট্যাব ও হেডার সহ)।
 * এই Spreadsheet-টিকে পরে TEMPLATE_SS_ID হিসেবে ব্যবহার করবেন।
 *******************************************************/
function setupDealerTemplateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheetsWithHeaders(ss, DEALER_SHEETS_DEF);

  formatMobileColumnAsText(ss, "Customers", "মোবাইল নং");

  SpreadsheetApp.getUi().alert("Dealer Template Spreadsheet প্রস্তুত হয়েছে। এই স্প্রেডশিটের ID টি Code.gs এর TEMPLATE_SS_ID তে বসান।");
}

/*******************************************************
 * মোবাইল নং কলামকে টেক্সট ফরম্যাট করার সহায়ক ফাংশন
 * (যাতে শুরুর 0 কখনো হারিয়ে না যায়)
 *******************************************************/
function formatMobileColumnAsText(ss, tabName, headerName) {
  const sheet = ss.getSheetByName(tabName);
  if (!sheet) return;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colIndex = headers.indexOf(headerName);
  if (colIndex === -1) return;
  const colNumber = colIndex + 1;
  // পুরো কলাম (হেডার বাদে ১০০০ রো পর্যন্ত) টেক্সট ফরম্যাট
  sheet.getRange(2, colNumber, 1000, 1).setNumberFormat("@");
}

/*******************************************************
 * আইডি জেনারেটর (prefix + ক্রমিক সংখ্যা, ৪ ডিজিট প্যাডেড)
 *******************************************************/
function generateId(sheet, prefix) {
  // সারির সংখ্যা নয়, বর্তমান সর্বোচ্চ আইডির পরেরটি — মাঝখানের কোনো সারি ডিলিট হলেও
  // নতুন আইডি আগের কোনো আইডির সাথে মিলে (ডুপ্লিকেট হয়ে) যাবে না
  const lastRow = sheet.getLastRow();
  let max = 0;
  if (lastRow >= 2) {
    const re = new RegExp("^" + prefix + "(\\d+)$");
    sheet.getRange(2, 1, lastRow - 1, 1).getValues().forEach(function (r) {
      const m = re.exec(String(r[0]));
      if (m) max = Math.max(max, Number(m[1]));
    });
  }
  return prefix + Utilities.formatString("%04d", max + 1);
}

/*******************************************************
 * পারফরম্যান্স হেল্পার: একাধিক রো একসাথে যোগ করা
 * (প্রতিটি রো এর জন্য আলাদা appendRow কল না করে, একবারে
 * setValues() দিয়ে সব রো লিখলে অনেক দ্রুত হয়)
 * rowObjects: [{header: value, ...}, ...]
 * প্রতিটি অবজেক্টের জন্য একটি করে নতুন ID ফেরত দেয় (idPrefix দিয়ে)
 *******************************************************/
function batchAppendRows(sheet, rowObjects, idPrefix, idHeaderName) {
  if (!rowObjects || rowObjects.length === 0) return [];

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const startRow = sheet.getLastRow() + 1;
  let nextNum = sheet.getLastRow() < 1 ? 1 : sheet.getLastRow();
  const generatedIds = [];

  const rows = rowObjects.map(function (obj) {
    let newId = "";
    if (idPrefix) {
      newId = idPrefix + Utilities.formatString("%04d", nextNum);
      nextNum++;
      generatedIds.push(newId);
    }
    return headers.map(function (h) {
      if (idHeaderName && h === idHeaderName) return newId;
      return obj.hasOwnProperty(h) ? obj[h] : "";
    });
  });

  sheet.getRange(startRow, 1, rows.length, headers.length).setValues(rows);
  return generatedIds;
}

/*******************************************************
 * পাসওয়ার্ড হ্যাশ (SHA-256)
 *******************************************************/
function hashPassword(password) {
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password);
  return raw.map(function (b) {
    return (b < 0 ? b + 256 : b).toString(16).padStart(2, "0");
  }).join("");
}

/*******************************************************
 * base64 ছবি ডেটা Google Drive এ সেভ করে পাবলিক-ভিউ URL রিটার্ন করে
 * imageBase64 ফরম্যাট: "data:image/png;base64,AAAA..." অথবা শুধু base64 অংশ
 *******************************************************/
/*******************************************************
 * ছবি সংরক্ষণ — Google Drive/Web App এর মধ্য দিয়ে সার্ভ করার একাধিক
 * চেষ্টা বারবার অনির্ভরযোগ্য প্রমাণিত হওয়ায় (Drive এর পাবলিক লিংক
 * ব্লক হয়ে যাওয়া, doGet এর মাধ্যমে বাইনারি সার্ভ করা অস্থিতিশীল
 * হওয়া), এখন ছবি সরাসরি একটি কমপ্রেসড base64 data URI হিসেবে
 * শীটের সেলেই সংরক্ষণ করা হয় — এটি কোনো বাহ্যিক লিংক/শেয়ারিং/
 * ডিপ্লয়মেন্ট URL এর উপর নির্ভর করে না, তাই সম্পূর্ণ নির্ভরযোগ্য।
 * ফ্রন্টএন্ড আপলোডের আগেই ছবি ছোট (৪০০পিক্সেল) ও কমপ্রেস করে পাঠায়,
 * তাই Google Sheet এর প্রতি-সেল ৫০,০০০ ক্যারেক্টার সীমার মধ্যেই থাকে।
 *******************************************************/
function processImageForStorage(imageBase64) {
  if (!imageBase64) return "";
  if (imageBase64.indexOf("data:image/") !== 0) return "";
  // Google Sheet এর একটি সেলে সর্বোচ্চ ৫০,০০০ ক্যারেক্টার রাখা যায়;
  // নিরাপদ মার্জিন রেখে ৪৮,০০০ এ সীমা বাঁধা হলো
  if (imageBase64.length > 48000) return null; // খুব বড় — সেভ করা যাবে না
  return imageBase64;
}


/*******************************************************
 * নতুন ডিলার রেজিস্ট্রেশন
 * data প্যারামিটারে যা থাকবে:
 * {
 *   নাম, পিতারনাম, nid, mobile, tradeLicense, thikana,
 *   adminUsername, adminPassword, adminName,
 *   nominee: { নাম, nid, mobile, somporko }   <-- নমিনি তথ্য
 * }
 *******************************************************/
function registerDealer(data) {
  const perm = checkAgencyPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };

  const masterSS = getMasterSS();
  const dealersSheet = getDealersSheet(masterSS);
  const nomineeSheet = getSheet(masterSS, "DealerNominee");
  const usersSheet = getSheet(masterSS, "Users");

  // ০. ডিলারের Gmail আবশ্যক — ডিলারের শীট এই Gmail-এর জন্যই তৈরি হয়,
  //    ডিপুর নিজের Drive-এ থেকে যায় না
  const dealerGmail = String(data.gmail || "").trim();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(dealerGmail)) {
    return { success: false, message: "ডিলারের Gmail দিন — এই Gmail-এই ডিলারের Google শীট তৈরি হবে" };
  }

  // ১. টেমপ্লেট কপি করে নতুন Spreadsheet বানানো (Spreadsheet ID স্বয়ংক্রিয় তৈরি হয়)
  const templateFile = DriveApp.getFileById(TEMPLATE_SS_ID);
  const newFile = templateFile.makeCopy(data.নাম + " - Dealer Sheet");
  const newSpreadsheetId = newFile.getId();

  // ২. শীটের মালিকানা ডিলারের Gmail-এ পাঠানো
  const handover = handOverSheetToDealer(newSpreadsheetId, dealerGmail);

  // ৩. Dealer আইডি তৈরি
  const dealerId = generateId(dealersSheet, "D");

  // ৪. ডিলারের ছবি (base64, ফ্রন্টএন্ডেই কমপ্রেসড) সরাসরি সেভ করা
  let dealerPhotoUrl = "";
  let photoWarning = "";
  if (data.dealerPhotoBase64) {
    const processed = processImageForStorage(data.dealerPhotoBase64);
    if (processed === null) { photoWarning = " (⚠️ ডিলারের ছবি খুব বড়, সেভ হয়নি — ছোট ছবি দিয়ে আবার চেষ্টা করুন)"; }
    else { dealerPhotoUrl = processed; }
  }

  // ৫. Dealers ট্যাবে রো যোগ
  dealersSheet.appendRow([
    dealerId,
    new Date(),
    data.নাম,
    data.পিতারনাম,
    data.nid,
    data.mobile,
    data.gmail || "",
    data.tradeLicense,
    data.thikana,
    dealerPhotoUrl,
    newSpreadsheetId,
    "সক্রিয়",
    data.facebookLink || "",
    data.area || ""
  ]);
  dealersSheet.getRange(dealersSheet.getLastRow(), 6).setNumberFormat("@");

  // ঠিকানার ৩ ঘর (ইউনিয়ন/পৌরসভা, উপজেলা, জেলা) — কলাম না থাকলে যোগ হয়, তারপর হেডার-নাম ধরে বসানো
  ensureDealerAddrCols(dealersSheet);
  const dHeaders = dealersSheet.getRange(1, 1, 1, dealersSheet.getLastColumn()).getValues()[0];
  const addrVals = { "ইউনিয়ন/পৌরসভা": data.union || "", "উপজেলা": data.upozila || "", "জেলা": data.jela || "" };
  Object.keys(addrVals).forEach(function (h) {
    const ci = dHeaders.indexOf(h);
    if (ci !== -1) dealersSheet.getRange(dealersSheet.getLastRow(), ci + 1).setValue(addrVals[h]);
  });

  // ৬. ডিলারের নমিনি তথ্য যোগ (যদি দেওয়া থাকে)
  if (data.nominee) {
    let nomineePhotoUrl = "";
    if (data.nominee.photoBase64) {
      const processedN = processImageForStorage(data.nominee.photoBase64);
      if (processedN === null) { photoWarning += " (⚠️ নমিনির ছবি খুব বড়, সেভ হয়নি)"; }
      else { nomineePhotoUrl = processedN; }
    }
    const nomineeId = generateId(nomineeSheet, "N");
    nomineeSheet.appendRow([
      nomineeId,
      dealerId,
      data.nominee.নাম,
      data.nominee.nid,
      data.nominee.mobile,
      data.nominee.somporko,
      nomineePhotoUrl
    ]);
    nomineeSheet.getRange(nomineeSheet.getLastRow(), 5).setNumberFormat("@");
  }

  // ৫. প্রথম Admin ইউজার তৈরি
  const userId = generateId(usersSheet, "U");
  usersSheet.appendRow([
    userId,
    dealerId,
    data.adminUsername,
    data.adminPassword,
    "Admin",
    data.adminName,
    data.mobile
  ]);
  usersSheet.getRange(usersSheet.getLastRow(), 7).setNumberFormat("@");

  invalidateAgencyCaches();

  return {
    success: true,
    dealerId: dealerId,
    spreadsheetId: newSpreadsheetId,
    message: "ডিলার সফলভাবে রেজিস্ট্রেশন হয়েছে। " + handover.message + photoWarning
  };
}

/*******************************************************
 * ডিলারের শীট ডিলারের Gmail-এর হাতে তুলে দেওয়া
 *
 * Apps Script সরাসরি অন্যের Drive-এ ফাইল বানাতে পারে না, তাই কপি করার পর
 * মালিকানা হস্তান্তর করা হয়:
 *  ১) একই Google Workspace ডোমেইন হলে সাথে সাথে মালিক বদলে যায়
 *  ২) সাধারণ @gmail.com হলে "মালিকানা হস্তান্তরের অনুরোধ" যায় — ডিলার
 *     ইমেইল/Drive থেকে গ্রহণ করলেই শীট তার Drive-এর হয়ে যায়
 *  ৩) দুটোই না হলে অন্তত এডিটর হিসেবে শেয়ার করা হয়
 * ডিপুর একাউন্ট এডিটর হিসেবে থেকে যায়, তাই সফটওয়্যার আগের মতোই চলে
 * ফেরত: { ok, mode: "owner"|"pending"|"editor"|"failed", message }
 *******************************************************/
function handOverSheetToDealer(fileId, gmail) {
  const email = String(gmail || "").trim();
  if (!email) return { ok: false, mode: "failed", message: "⚠️ ডিলারের Gmail নেই — শীট শেয়ার করা যায়নি।" };

  let file;
  try { file = DriveApp.getFileById(fileId); }
  catch (e) { return { ok: false, mode: "failed", message: "⚠️ শীট খুঁজে পাওয়া যায়নি: " + e }; }

  // আগে থেকেই ডিলারের হলে কিছু করার নেই
  try {
    const owner = file.getOwner();
    if (owner && owner.getEmail().toLowerCase() === email.toLowerCase()) {
      return { ok: true, mode: "owner", message: "শীট ডিলারের Gmail-এ আছে।" };
    }
  } catch (e) { /* বাদ */ }

  // ১) সরাসরি মালিক বদল (Workspace)
  try {
    file.setOwner(email);
    const owner2 = file.getOwner();
    if (owner2 && owner2.getEmail().toLowerCase() === email.toLowerCase()) {
      return { ok: true, mode: "owner", message: "ডিলারের Gmail-এ Google শীট তৈরি হয়েছে।" };
    }
  } catch (e) { /* সাধারণ Gmail — নিচের ধাপে যাবে */ }

  // ২) মালিকানা হস্তান্তরের অনুরোধ (সাধারণ Gmail)
  let apiError = "";
  try {
    const resp = UrlFetchApp.fetch(
      "https://www.googleapis.com/drive/v3/files/" + fileId + "/permissions?sendNotificationEmail=true&supportsAllDrives=true",
      {
        method: "post",
        contentType: "application/json",
        headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
        payload: JSON.stringify({ role: "writer", type: "user", emailAddress: email, pendingOwner: true }),
        muteHttpExceptions: true
      }
    );
    if (resp.getResponseCode() >= 200 && resp.getResponseCode() < 300) {
      return {
        ok: true, mode: "pending",
        message: "ডিলারের Gmail-এ শীট পাঠানো হয়েছে — ডিলার ইমেইল/Google Drive থেকে মালিকানা গ্রহণ করলে শীটটি তার Drive-এর হবে।"
      };
    }
    apiError = resp.getContentText().slice(0, 200);
  } catch (e) { apiError = String(e); }

  // ৩) শেষ চেষ্টা — এডিটর হিসেবে শেয়ার
  try {
    file.addEditor(email);
    return {
      ok: false, mode: "editor",
      message: "⚠️ মালিকানা হস্তান্তর করা যায়নি, ডিলারকে শুধু এডিটর হিসেবে শেয়ার করা হয়েছে (শীট এখনো ডিপুর Drive-এ আছে)। কারণ: " + apiError
    };
  } catch (e) {
    return { ok: false, mode: "failed", message: "⚠️ শীট ডিলারের সাথে শেয়ার করা যায়নি — Gmail ঠিক আছে কিনা দেখুন। " + apiError };
  }
}

/*******************************************************
 * একবার ম্যানুয়ালি রান করুন (ফাংশন সিলেক্ট করে ▶ Run) — যেসব ডিলারের
 * শীট আগে ডিপুর Drive-এ তৈরি হয়েছে তাদের Gmail-এ হস্তান্তর শুরু হবে।
 * ফলাফল দেখতে: Execution log
 *******************************************************/
function handOverExistingDealerSheets() {
  const dealers = genericListRows(getDealersSheet(getMasterSS()));
  dealers.forEach(function (d) {
    const gmail = String(d["Gmail"] || "").trim();
    const fileId = String(d["SpreadsheetID"] || "").trim();
    if (!gmail || !fileId) {
      Logger.log((d["DealerID"] || "?") + " " + (d["নাম"] || "") + ": বাদ — Gmail বা শীট নেই");
      return;
    }
    const r = handOverSheetToDealer(fileId, gmail);
    Logger.log((d["DealerID"] || "?") + " " + (d["নাম"] || "") + ": " + r.mode + " — " + r.message);
  });
}

/*******************************************************
 * ওয়েব অ্যাপ এন্ট্রি পয়েন্ট — সকল action এখানে রাউট হয়
 * ফ্রন্টএন্ড থেকে POST বডিতে { action: "...", data: {...} } পাঠাতে হবে
 *******************************************************/
/*******************************************************
 * ডাবল সেভ ঠেকানো (idempotency):
 * সংরক্ষণ/এডিট/ডিলিটের প্রতিটি অনুরোধের সাথে ব্রাউজার একটি ইউনিক requestId
 * পাঠায়। নেটওয়ার্কে সমস্যা হলে ব্রাউজার একই requestId দিয়ে আবার পাঠায় —
 * সার্ভার আগেই সেটা করে ফেলে থাকলে নতুন করে না করে আগের ফলই ফেরত দেয়।
 * LockService দিয়ে একই সময়ে দুটি সংরক্ষণ একই সারিতে লেখা (ওভাররাইট/একই ID)
 * হওয়াও আটকানো হয়।
 *******************************************************/
const IDEMPOTENT_ACTIONS = {
  addCustomer:1, addCustomers:1, updateCustomer:1, deleteCustomer:1,
  addPackage:1, updatePackage:1, deletePackage:1,
  addSale:1, updateSalePrice:1, cancelSale:1,
  addCommission:1, updateCommission:1, deleteCommission:1,
  addExpense:1, updateExpense:1, deleteExpense:1,
  addProduct:1, updateProduct:1, deleteProduct:1,
  addStockVoucher:1, updateStockVoucherEntry:1, deleteStockVoucher:1,
  addSalesInvoice:1, updateSalesInvoiceEntry:1, updateSalesInvoiceFull:1, updateSalesInvoiceStatus:1, deleteSalesInvoice:1,
  confirmPendingOrder:1, rejectPendingOrder:1,
  addExpenseVoucher:1, updateExpenseVoucher:1, deleteExpenseVoucher:1,
  addIncomeReceipt:1, updateIncomeReceiptEntry:1, deleteIncomeReceipt:1
};

/*******************************************************
 * লোড দ্রুত করা — সার্ভার-সাইড রিড ক্যাশ:
 * বড় তালিকা/রিপোর্ট (গ্রাহক, বিক্রি, ড্যাশবোর্ড ইত্যাদি) ৫ মিনিট ক্যাশে (সংকুচিত
 * করে) রাখা হয় — একই তথ্য আবার চাইলে শীট না খুলে সাথে সাথে ফেরত যায়।
 * যেকোনো সংরক্ষণ/এডিট/ডিলিট হলেই "ডাটা ভার্সন" বদলে যায়, ফলে পুরনো ক্যাশ আর
 * ব্যবহার হয় না — সবসময় সর্বশেষ তথ্যই আসে। ক্যাশের কী-তে ডিলার আইডি ও রোল থাকে,
 * তাই একজনের ডাটা অন্যজনের কাছে যাওয়ার সুযোগ নেই (টোকেন আগে যাচাই হয়)।
 * (গুগল শীটে সরাসরি হাতে এডিট করলে অ্যাপে সর্বোচ্চ ৫ মিনিট পরে দেখা যাবে)
 *******************************************************/
const CACHEABLE_READS = {
  listCustomers:1, getSalesPageData:1, dashboardSummary:1, agencyDashboardSummary:1,
  listDealers:1, listSalesInvoices:1, listPendingOrderConfirmations:1,
  listProducts:1, listStockVouchers:1, listExpenses:1, listExpenseVouchers:1,
  listIncomeReceipts:1, listCommissions:1, listCustomersForDealer:1,
  getAgencyInfo:1, getAbout:1, listSales:1,
  dealerReport:1, agencySalesReport:1, depotReport:1, dailyReport:1, monthlyReport:1, totalReport:1
};
const READ_CACHE_TTL = 300;
const READ_CACHE_SALT = "r3";   // তালিকা/রিপোর্টের কাঠামো বদলালে এটা বদলান — পুরনো ক্যাশ আর মিলবে না

/* সার্ভার কোডের সংস্করণ — অ্যাপের নিচে দেখায়; নতুন ভার্সন ডিপ্লয় হয়েছে কিনা বোঝার জন্য */
const BACKEND_VERSION = "2026-10-06.5";
const MUTATING_RE = /^(add|update|delete|cancel|confirm|reject|set|save|send|register)/;

function getDataVersion(cache) {
  let v = cache.get("gver");
  if (!v) {
    v = String(Date.now());
    try { cache.put("gver", v, 21600); } catch (e) { /* বাদ */ }
  }
  return v;
}
function bumpDataVersion() {
  try { CacheService.getScriptCache().put("gver", Date.now() + "_" + Math.floor(Math.random() * 1000000), 21600); }
  catch (e) { /* বাদ */ }
}

function md5Hex(str) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, str, Utilities.Charset.UTF_8)
    .map(function (b) { return ("0" + (b < 0 ? b + 256 : b).toString(16)).slice(-2); }).join("");
}

/* বড় JSON গজিপ + base64 করে ৯০KB টুকরোয় (CacheService এর ১০০KB সীমার ভেতরে) রাখা */
function cachePutBig(cache, key, str, ttl) {
  const b64 = Utilities.base64Encode(Utilities.gzip(Utilities.newBlob(str, "application/json", "d.json")).getBytes());
  const CH = 90000;
  const n = Math.ceil(b64.length / CH);
  if (n > 10) return;                                   // খুব বড় — ক্যাশ করা হবে না
  const obj = {};
  for (let i = 0; i < n; i++) obj[key + "_" + i] = b64.substr(i * CH, CH);
  obj[key + "_m"] = String(n);
  cache.putAll(obj, ttl);
}
function cacheGetBig(cache, key) {
  const meta = cache.get(key + "_m");
  if (!meta) return null;
  const n = Number(meta);
  const keys = [];
  for (let i = 0; i < n; i++) keys.push(key + "_" + i);
  const parts = cache.getAll(keys);
  let b64 = "";
  for (let i = 0; i < n; i++) {
    const p = parts[key + "_" + i];
    if (p == null) return null;
    b64 += p;
  }
  return Utilities.ungzip(Utilities.newBlob(Utilities.base64Decode(b64), "application/x-gzip")).getDataAsString();
}

/* রিড অ্যাকশন: ক্যাশে থাকলে সেখান থেকে, না থাকলে চালিয়ে ক্যাশ করে */
function runActionCached(action, data) {
  const payload = data.token ? verifyToken(data.token) : null;
  if (!payload) return runAction(action, data);          // সঠিক এরর বার্তা সেখান থেকেই আসবে
  const cache = CacheService.getScriptCache();
  const p = {};
  Object.keys(data).forEach(function (k) { if (k !== "token" && k !== "requestId") p[k] = data[k]; });
  const key = "rc_" + md5Hex([READ_CACHE_SALT, action, payload.dealerId, payload.role, getDataVersion(cache), JSON.stringify(p)].join("|"));
  try {
    const hit = cacheGetBig(cache, key);
    if (hit) return hit;
  } catch (e) { /* ক্যাশ নষ্ট — সরাসরি চালান */ }
  const out = runAction(action, data);
  try {
    if (out.slice(0, 40).indexOf('"success":true') !== -1) cachePutBig(cache, key, out, READ_CACHE_TTL);
  } catch (e) { /* বাদ */ }
  return out;
}

/* গ্রাহকের কাজে পুরো অনুরোধ জুড়ে লক ধরে রাখা হয় না — শুধু আইডি তৈরি + সারি লেখার ছোট অংশে
   (addCustomer/deleteCustomer এর ভেতরে) লক নেয়। ফলে একটি ধীর সংরক্ষণ বাকিদের আটকে রাখে না */
const FINE_LOCK_ACTIONS = { addCustomer:1, addCustomers:1, updateCustomer:1, deleteCustomer:1 };

/* উত্তরের JSON এ সার্ভারে কত মিলিসেকেন্ড লেগেছে (_ms) জুড়ে দেওয়া — ধীর হলে কারণ খোঁজার জন্য */
function withTiming(out, t0) {
  if (typeof out !== "string" || out.charAt(out.length - 1) !== "}") return out;
  const ms = Date.now() - t0;
  return out === "{}" ? '{"_ms":' + ms + "}" : out.slice(0, -1) + ',"_ms":' + ms + "}";
}

/* ডুপ্লিকেট-সুরক্ষিত কিন্তু পুরো সময় লক না ধরে চালানো: লক শুধু "এই requestId আগে এসেছে কিনা" চেক ও চিহ্ন বসাতে */
function runIdempotentFine(action, data, reqKey, cache) {
  const lock = LockService.getScriptLock();
  let prev = null;
  try {
    lock.waitLock(15000);
    try {
      prev = cache.get(reqKey);
      if (!prev) cache.put(reqKey, "P", 120);          // "P" = এই অনুরোধ এখন চলছে
    } finally { lock.releaseLock(); }
  } catch (lockErr) {
    return JSON.stringify({ success: false, message: "সার্ভার ব্যস্ত, কিছুক্ষণ পরে আবার চেষ্টা করুন" });
  }

  if (prev && prev !== "P") return prev;               // আগেই হয়ে গেছে — আবার করা হবে না
  if (prev === "P") {                                   // একই অনুরোধ এখনো চলছে — শেষ হওয়া পর্যন্ত অপেক্ষা
    for (let i = 0; i < 40; i++) {
      Utilities.sleep(600);
      const v = cache.get(reqKey);
      if (v && v !== "P") return v;
      if (!v) break;
    }
    return JSON.stringify({ success: false, message: "আগের অনুরোধ এখনো চলছে — কিছুক্ষণ পরে আবার চেষ্টা করুন" });
  }

  const out = runAction(action, data);
  const ok = out.slice(0, 40).indexOf('"success":true') !== -1;
  if (ok) {
    bumpDataVersion();                                  // ক্যাশ করা পুরনো তথ্য বাতিল
    try { cache.put(reqKey, out, 600); }
    catch (c) { try { cache.put(reqKey, JSON.stringify({ success: true, message: "সম্পন্ন হয়েছে" }), 600); } catch (x) { /* বাদ */ } }
  } else {
    try { cache.remove(reqKey); } catch (x) { /* বাদ */ }   // ব্যর্থ হলে আবার চেষ্টা করা যাবে
  }
  return out;
}

/* শপের নাম একরূপ রাখা — শীটের ডাটায় পুরনো নাম থাকলেও সব উত্তরে নতুন নাম যায় */
const ORG_NAME_NEW = "ফেনী মানবিক সহায়তা (স্বস্তির বাজার)";
function normalizeOrgName(str) {
  if (typeof str !== "string") return str;
  return str.replace(/(ফেনী\s*)?মানবিক\s*সহায়তা\s*ডেইল[িী]\s*শপ/g, ORG_NAME_NEW);
}
/* লগইন পেজে দেখানোর জন্য শুধু এজেন্সির লোগো ও নাম — গোপন কিছু নয় */
function getLoginBranding() {
  const cache = CacheService.getScriptCache();
  try { const hit = cacheGetBig(cache, "login_brand"); if (hit) return JSON.parse(hit); } catch (e) { /* বাদ */ }
  const rows = genericListRows(getSheet(getMasterSS(), "Agency"));
  const r = rows.length > 0 ? rows[0] : {};
  const out = { success: true, logo: r["লোগো(URL)"] || "", name: r["নাম"] || "" };
  try { cachePutBig(cache, "login_brand", JSON.stringify(out), 120); } catch (e) { /* বাদ */ }
  return out;
}

function respondJson(out, t0) {
  return ContentService.createTextOutput(normalizeOrgName(withTiming(out, t0))).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const t0 = Date.now();
  const params = JSON.parse(e.postData.contents);
  const action = params.action;
  const data = params.data || {};

  // সংরক্ষণ-জাতীয় অনুরোধ হলে ডুপ্লিকেট চেক সহ চালানো হয়
  if (IDEMPOTENT_ACTIONS[action] && data.requestId && FINE_LOCK_ACTIONS[action]) {
    const reqKeyF = "req_" + String(data.requestId).slice(0, 80);
    return respondJson(runIdempotentFine(action, data, reqKeyF, CacheService.getScriptCache()), t0);
  }
  if (IDEMPOTENT_ACTIONS[action] && data.requestId) {
    const reqKey = "req_" + String(data.requestId).slice(0, 80);
    const cache = CacheService.getScriptCache();
    const lock = LockService.getScriptLock();
    let out;
    try {
      lock.waitLock(25000);
      const prev = cache.get(reqKey);
      if (prev) {
        out = prev;                                   // আগেই হয়ে গেছে — আবার করা হবে না
      } else {
        out = runAction(action, data);
        if (MUTATING_RE.test(action)) bumpDataVersion();     // ক্যাশ করা পুরনো তথ্য বাতিল
        try { cache.put(reqKey, out, 600); } catch (cacheErr) { /* বড় ফল — বাদ */ }
      }
    } catch (lockErr) {
      out = JSON.stringify({ success: false, message: "সার্ভার ব্যস্ত, কিছুক্ষণ পরে আবার চেষ্টা করুন" });
    } finally {
      try { lock.releaseLock(); } catch (x) { /* বাদ */ }
    }
    return respondJson(out, t0);
  }

  let output;
  if (CACHEABLE_READS[action]) {
    output = runActionCached(action, data);
  } else {
    output = runAction(action, data);
    if (MUTATING_RE.test(action)) bumpDataVersion();
  }
  return respondJson(output, t0);
}

/* action চালিয়ে JSON স্ট্রিং ফেরত দেয় */
function runAction(action, data) {
  let result;
  try {
    switch (action) {
      // ---- রেজিস্ট্রেশন ----
      case "registerDealer":
        result = registerDealer(data);
        break;

      // ---- সার্ভার ওয়ার্ম-আপ (কোল্ড স্টার্টের দেরি এড়াতে) ----
      case "ping":
        result = { success: true, version: BACKEND_VERSION };
        break;

      // ---- লগইন পেজের লোগো (লগইনের আগে লাগে, তাই টোকেন ছাড়া) ----
      case "getLoginBranding":
        result = getLoginBranding();
        break;

      // ---- অথেনটিকেশন ----
      case "login":
        result = loginUser(data);
        break;
      case "addUser":
        result = addUser(data);
        break;
      case "recoverPassword":
        result = recoverPassword(data);
        break;

      // ---- গ্রাহক ----
      case "addCustomer":
        result = addCustomer(data);
        break;
      case "addCustomers":
        result = addCustomers(data);
        break;
      case "listCustomers":
        result = listCustomers(data);
        break;
      case "listCustomersForDealer":
        result = listCustomersForDealer(data);
        break;
      case "updateCustomer":
        result = updateCustomer(data);
        break;
      case "deleteCustomer":
        result = deleteCustomer(data);
        break;

      // ---- প্যাকেজ ----
      case "addPackage":
        result = addPackage(data);
        break;
      case "listPackages":
        result = listPackages(data);
        break;
      case "updatePackage":
        result = updatePackage(data);
        break;
      case "deletePackage":
        result = deletePackage(data);
        break;

      // ---- বিক্রি ----
      case "addSale":
        result = addSale(data);
        break;
      case "updateSalePrice":
        result = updateSalePrice(data);
        break;
      case "cancelSale":
        result = cancelSale(data);
        break;
      case "listSales":
        result = listSales(data);
        break;
      case "getSalesPageData":
        result = getSalesPageData(data);
        break;

      // ---- ড্যাশবোর্ড ----
      case "dashboardSummary":
        result = getDashboardSummary(data);
        break;

      // ---- রিপোর্ট ----
      case "dailyReport":
        result = dailyReport(data);
        break;
      case "monthlyReport":
        result = monthlyReport(data);
        break;
      case "totalReport":
        result = totalReport(data);
        break;
      case "dealerReport":
        result = dealerReport(data);
        break;

      // ---- এজেন্সি সেটাপ (শুধু মূল এজেন্সি) ----
      case "getAgencyInfo":
        result = getAgencyInfo(data);
        break;
      case "updateAgencyInfo":
        result = updateAgencyInfo(data);
        break;

      // ---- ডিলার তালিকা/এডিট (শুধু মূল এজেন্সি) ----
      case "listDealers":
        result = listDealers(data);
        break;
      case "updateDealerInfo":
        result = updateDealerInfo(data);
        break;
      case "setDealerStatus":
        result = setDealerStatus(data);
        break;
      case "updateDealerNominee":
        result = updateDealerNominee(data);
        break;
      case "getDealerFullInfo":
        result = getDealerFullInfo(data);
        break;
      case "updateDealerFull":
        result = updateDealerFull(data);
        break;

      // ---- ডিলার সাইট: ডিলার নিজের তথ্য দেখা/আপডেট (শুধু ডিলারের Admin) ----
      case "getMyDealerInfo":
        result = getMyDealerInfo(data);
        break;
      case "updateMyDealerFull":
        result = updateMyDealerFull(data);
        break;

      // ---- প্রতি ডিলারের নিজস্ব "সম্পর্কে" ----
      case "getAbout":
        result = getAbout(data);
        break;
      case "updateAbout":
        result = updateAbout(data);
        break;

      // ---- কমিশন (শুধু মূল এজেন্সি) ----
      case "addCommission":
        result = addCommission(data);
        break;
      case "listCommissions":
        result = listCommissions(data);
        break;
      case "updateCommission":
        result = updateCommission(data);
        break;
      case "deleteCommission":
        result = deleteCommission(data);
        break;

      // ---- খরচ (শুধু মূল এজেন্সি) ----
      case "addExpense":
        result = addExpense(data);
        break;
      case "listExpenses":
        result = listExpenses(data);
        break;
      case "updateExpense":
        result = updateExpense(data);
        break;
      case "deleteExpense":
        result = deleteExpense(data);
        break;

      // ---- স্টক: পণ্য তালিকা (শুধু মূল এজেন্সি) ----
      case "addProduct":
        result = addProduct(data);
        break;
      case "listProducts":
        result = listProducts(data);
        break;
      case "updateProduct":
        result = updateProduct(data);
        break;
      case "deleteProduct":
        result = deleteProduct(data);
        break;

      // ---- স্টক ইন ভাউচার (শুধু মূল এজেন্সি) ----
      case "addStockVoucher":
        result = addStockVoucher(data);
        break;
      case "listStockVouchers":
        result = listStockVouchers(data);
        break;
      case "updateStockVoucherEntry":
        result = updateStockVoucherEntry(data);
        break;
      case "deleteStockVoucher":
        result = deleteStockVoucher(data);
        break;

      // ---- বিক্রয় ইনভয়েস (শুধু মূল এজেন্সি) ----
      case "addSalesInvoice":
        result = addSalesInvoice(data);
        break;
      case "listSalesInvoices":
        result = listSalesInvoices(data);
        break;
      case "updateSalesInvoiceEntry":
        result = updateSalesInvoiceEntry(data);
        break;
      case "updateSalesInvoiceFull":
        result = updateSalesInvoiceFull(data);
        break;
      case "updateSalesInvoiceStatus":
        result = updateSalesInvoiceStatus(data);
        break;
      case "deleteSalesInvoice":
        result = deleteSalesInvoice(data);
        break;
      case "listPendingOrderConfirmations":
        result = listPendingOrderConfirmations(data);
        break;
      case "confirmPendingOrder":
        result = confirmPendingOrder(data);
        break;
      case "rejectPendingOrder":
        result = rejectPendingOrder(data);
        break;

      // ---- এজেন্সি রিপোর্ট ----
      case "agencySalesReport":
        result = agencySalesReport(data);
        break;
      case "depotReport":
        result = depotReport(data);
        break;

      // ---- খরচ ভাউচার (ডিলারের নিজস্ব) ----
      case "addExpenseVoucher":
        result = addExpenseVoucher(data);
        break;
      case "listExpenseVouchers":
        result = listExpenseVouchers(data);
        break;
      case "updateExpenseVoucher":
        result = updateExpenseVoucher(data);
        break;
      case "deleteExpenseVoucher":
        result = deleteExpenseVoucher(data);
        break;

      // ---- আয় রশিদ (ডিপু ও ডিলার — টোকেন অনুযায়ী নিজের শীটে) ----
      case "addIncomeReceipt":
        result = addIncomeReceipt(data);
        break;
      case "listIncomeReceipts":
        result = listIncomeReceipts(data);
        break;
      case "updateIncomeReceiptEntry":
        result = updateIncomeReceiptEntry(data);
        break;
      case "deleteIncomeReceipt":
        result = deleteIncomeReceipt(data);
        break;

      // ---- এজেন্সি ড্যাশবোর্ড ----
      case "agencyDashboardSummary":
        result = getAgencyDashboardSummary(data);
        break;

      // ---- মেসেজিং (SMS/WhatsApp) — শুধু ডিলার সাইট ----
      case "getMessagingStatus":
        result = getMessagingStatus(data);
        break;
      case "getDealerMessagingConfig":
        result = getDealerMessagingConfig(data);
        break;
      case "saveDealerMessagingConfig":
        result = saveDealerMessagingConfig(data);
        break;
      case "sendCustomerMessage":
        result = sendCustomerMessage(data);
        break;
      case "listMessages":
        result = listMessages(data);
        break;
      case "deleteMessage":
        result = deleteMessage(data);
        break;

      default:
        result = { success: false, message: "অজানা action: " + action };
    }
  } catch (err) {
    result = { success: false, message: err.toString() };
  }
  return JSON.stringify(result);
}

function doGet(e) {
  try {
    // গ্রাহক কার্ডের QR কোড স্ক্যান করলে এই পেজ দেখাবে
    if (e && e.parameter && e.parameter.view === "customer" && e.parameter.dealerId && e.parameter.customerId) {
      return renderCustomerCardPage(e.parameter.dealerId, e.parameter.customerId);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: "TCB Dealer System API চালু আছে", version: BACKEND_VERSION, url: ScriptApp.getService().getUrl() }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    // কোনো অপ্রত্যাশিত এরর হলেও যেন Google এর জেনেরিক "unable to open"
    // এরর পেজের বদলে আমাদের নিজস্ব বোধগম্য এরর পেজ দেখায়
    return HtmlService.createHtmlOutput(
      '<div style="font-family:sans-serif;padding:20px;color:#a3352b;">সমস্যা হয়েছে: ' + err.toString() + '</div>'
    ).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

/*******************************************************
 * QR কোড স্ক্যান করলে খোলা হবে এমন একটি সাধারণ, মোবাইল-বান্ধব HTML পেজ
 *******************************************************/
function renderCustomerCardPage(dealerId, customerId) {
  const info = getCustomerCardInfo(dealerId, customerId);

  function fmtDate(d) {
    if (!d) return "—";
    const dt = (d instanceof Date) ? d : new Date(d);
    if (isNaN(dt.getTime())) return "—";
    return Utilities.formatDate(dt, Session.getScriptTimeZone(), "dd/MM/yyyy");
  }

  let bodyHtml;
  if (!info.success) {
    bodyHtml = '<p style="color:#a3352b;">' + info.message + '</p>';
  } else {
    bodyHtml =
      '<h2 style="margin:0 0 14px;color:#153f37;">' + info.customerName + '</h2>' +
      '<table style="width:100%;border-collapse:collapse;">' +
      '<tr><td style="padding:8px;border:1px solid #d8d2c2;font-weight:600;background:#eef0e9;">গ্রাহক হওয়ার তারিখ</td><td style="padding:8px;border:1px solid #d8d2c2;">' + fmtDate(info.joinDate) + '</td></tr>' +
      '<tr><td style="padding:8px;border:1px solid #d8d2c2;font-weight:600;background:#eef0e9;">মোট ক্রয় সংখ্যা</td><td style="padding:8px;border:1px solid #d8d2c2;">' + info.purchaseCount + ' বার</td></tr>' +
      '<tr><td style="padding:8px;border:1px solid #d8d2c2;font-weight:600;background:#eef0e9;">সর্বশেষ ক্রয়ের তারিখ</td><td style="padding:8px;border:1px solid #d8d2c2;">' + fmtDate(info.lastPurchaseDate) + '</td></tr>' +
      '</table>';
  }

  const html =
    '<!DOCTYPE html><html lang="bn"><head><meta charset="utf-8">' +
    '<meta name="viewport" content="width=device-width, initial-scale=1.0">' +
    '<title>গ্রাহক তথ্য</title>' +
    '<style>body{font-family:sans-serif;background:#f6f4ee;padding:20px;color:#1c2b28;}' +
    '.card{background:#fff;border-radius:10px;padding:20px;max-width:400px;margin:0 auto;box-shadow:0 4px 14px rgba(0,0,0,0.08);}</style>' +
    '</head><body><div class="card">' + bodyHtml + '</div></body></html>';

  return HtmlService.createHtmlOutput(normalizeOrgName(html))
    .setTitle("গ্রাহক তথ্য")
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

