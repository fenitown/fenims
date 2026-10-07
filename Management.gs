/*******************************************************
 * TCB-স্টাইল ডিলার ম্যানেজমেন্ট সফটওয়্যার
 * Management.gs — ব্যবস্থাপনা: এজেন্সি সেটাপ, ডিলার সেটাপ/লিস্ট,
 * এবং প্রতি ডিলারের নিজস্ব "সম্পর্কে" তথ্য
 *
 * MasterRegistry এর একই Apps Script প্রজেক্টে নতুন ফাইল
 * হিসেবে যোগ করুন (নাম দিন: Management)
 *
 * ================= গুরুত্বপূর্ণ ধারণা =================
 * এজেন্সি-লেভেল কাজ (এজেন্সি সেটাপ, ডিলার তৈরি/এডিট/নমিনি এডিট) —
 * এগুলো মূল এজেন্সির নিজস্ব "AGENCY" লগইন দিয়ে করতে হবে, কোনো
 * সাধারণ ডিলারের Admin দিয়ে নয়। একবার নিচের
 * createAgencyAdminOnce() ফাংশনটি ম্যানুয়ালি রান করে মূল
 * এজেন্সির জন্য একটি লগইন তৈরি করে নিন।
 *
 * প্রতিটি ডিলারের নিজস্ব "সম্পর্কে" তথ্য (যা তার সাইটের "সম্পর্কে"
 * মেনুতে দেখা যাবে) — এটি সেই ডিলারের নিজের Admin এডিট করতে পারবে,
 * কারণ এটি তার নিজের Spreadsheet এর "About" ট্যাবে থাকে।
 *******************************************************/

const AGENCY_ID = "AGENCY";

/*******************************************************
 * পারফরম্যান্স: এজেন্সি-লেভেল রিড-হেভি ডাটা (ডিলার তালিকা, এজেন্সি
 * তথ্য, সংস্থার তথ্য) ক্যাশ করা — এগুলো কম পরিবর্তন হয় কিন্তু প্রতিটি
 * ডিলারের প্রতিটি পেজ লোডে বারবার পড়া হয়, তাই ক্যাশ করলে বড় গতি লাভ হয়
 *******************************************************/
/*******************************************************
 * Dealers ট্যাব — পুরনো শীটে "এরিয়া" কলাম না থাকলে শেষে নিজে যোগ করে নেয়
 * (বিদ্যমান কলামের জায়গা বদলায় না, তাই আগের ডেটা নিরাপদ)
 *******************************************************/
function getDealersSheet(masterSS) {
  const sheet = getSheet(masterSS, "Dealers");
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  if (headers.indexOf("এরিয়া") === -1) {
    sheet.getRange(1, lastCol + 1).setValue("এরিয়া").setFontWeight("bold");
  }
  return sheet;
}

function invalidateDealerMsgInfo(dealerId) {
  try { CacheService.getScriptCache().remove("dealer_nameaddr2_" + dealerId); } catch (e) { /* বাদ */ }
}

function invalidateAgencyCaches() {
  const cache = CacheService.getScriptCache();
  cache.removeAll(["dealers_cache_v1", "agency_info_cache_v3", "about_info_cache_v1"]);
}

/*******************************************************
 * একবার ম্যানুয়ালি রান করুন — মূল এজেন্সির জন্য প্রথম লগইন তৈরি হবে
 * (ফাংশন ড্রপডাউন থেকে createAgencyAdminOnce সিলেক্ট করে ▶ Run করুন,
 * নিচের username/password চাইলে বদলে নিন)
 *******************************************************/
function createAgencyAdminOnce() {
  const masterSS = getMasterSS();
  const usersSheet = getSheet(masterSS, "Users");

  const username = "Office";          // চাইলে বদলান
  const password = "12345";           // অবশ্যই পরে বদলে নিন

  const userId = generateId(usersSheet, "U");
  usersSheet.appendRow([
    userId,
    AGENCY_ID,
    username,
    password,
    "Admin",
    "মূল ডিপু এডমিন",
    ""
  ]);

  SpreadsheetApp.getUi().alert(
    "ডিপু এডমিন তৈরি হয়েছে।\nইউজারনেম: " + username + "\nপাসওয়ার্ড: " + password +
    "\n\nঅনুগ্রহ করে লগইন করার পর এই পাসওয়ার্ড অবশ্যই পরিবর্তন করুন।"
  );
}

/*******************************************************
 * একবার ম্যানুয়ালি রান করুন (ফাংশন ড্রপডাউন থেকে resetAgencyAdminLogin ▶ Run):
 * মূল ডিপুর বিদ্যমান Admin লগইন বদলে ইউজারনেম "Office" ও পাসওয়ার্ড "12345" করে দেয়।
 * ডিপু Admin না থাকলে নতুন তৈরি করে।
 *******************************************************/
function resetAgencyAdminLogin() {
  const NEW_USER = "Office", NEW_PASS = "12345";
  const usersSheet = getSheet(getMasterSS(), "Users");
  const values = usersSheet.getDataRange().getValues();
  const h = values[0];
  const iDealer = h.indexOf("DealerID"), iRole = h.indexOf("রোল");
  const iUser = h.indexOf("ইউজারনেম"), iPass = h.indexOf("পাসওয়ার্ড");

  let row = -1;
  for (let i = 1; i < values.length; i++) {
    if (values[i][iDealer] === AGENCY_ID && values[i][iRole] === "Admin") { row = i + 1; break; }
  }
  if (row === -1) {
    usersSheet.appendRow([generateId(usersSheet, "U"), AGENCY_ID, NEW_USER, NEW_PASS, "Admin", "মূল ডিপু এডমিন", ""]);
  } else {
    usersSheet.getRange(row, iUser + 1).setValue(NEW_USER);
    usersSheet.getRange(row, iPass + 1).setNumberFormat("@").setValue(NEW_PASS);
  }
  const msg = "ডিপু লগইন হালনাগাদ হয়েছে। ইউজারনেম: " + NEW_USER + ", পাসওয়ার্ড: " + NEW_PASS;
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { /* এডিটর থেকে চালালে লগেই দেখা যাবে */ }
}

/*******************************************************
 * ডিলারের ঠিকানা ৩ ঘরে (ইউনিয়ন/পৌরসভা, উপজেলা, জেলা) — Dealers ট্যাবে কলাম না থাকলে
 * নিজে থেকে যোগ হয় (আলাদা সেটআপ লাগে না)। "ঠিকানা" কলামে কমা দিয়ে জোড়া ঠিকানা থাকে (প্রিন্টের জন্য)
 *******************************************************/
const DEALER_ADDR_COLS = ["ইউনিয়ন/পৌরসভা", "উপজেলা", "জেলা"];
function ensureDealerAddrCols(sheet) {
  const cache = CacheService.getScriptCache();
  if (cache.get("dealer_addr_cols_v1")) return;
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  let col = lastCol;
  DEALER_ADDR_COLS.forEach(function (h) {
    if (headers.indexOf(h) === -1) { col++; sheet.getRange(1, col).setValue(h); }
  });
  try { cache.put("dealer_addr_cols_v1", "1", 21600); } catch (e) { /* বাদ */ }
}

/*******************************************************
 * এজেন্সি-লেভেল পারমিশন চেক (শুধু AGENCY_ID এর Admin)
 *******************************************************/
function checkAgencyPermission(token) {
  const perm = checkPermission(token, ["Admin"]);
  if (!perm.ok) return perm;
  if (perm.payload.dealerId !== AGENCY_ID) {
    return { ok: false, message: "এই কাজ শুধুমাত্র মূল ডিপু করতে পারবে" };
  }
  return perm;
}

/*=========================================================
 *  এজেন্সি সেটাপ (নাম, মোবাইল নং, লোগো)
 *=======================================================*/
function getAgencyInfo(data) {
  // এজেন্সির নাম/মোবাইল/লোগো/ব্যবস্থাপক তালিকা/সোশ্যাল লিংক — স্পর্শকাতর
  // তথ্য না, তাই ডিলার ও এজেন্সি উভয়েই দেখতে পারবে (মেনুবার লোগো,
  // ফেভিকন, ফুটার, "আমাদের সম্পর্কে" ইত্যাদিতে ব্যবহারের জন্য); শুধু
  // এডিট (updateAgencyInfo) এজেন্সি-নির্দিষ্ট থাকবে
  const perm = checkPermission(data.token, ["Admin", "প্রতিনিধি"]);
  if (!perm.ok) return { success: false, message: perm.message };

  const cache = CacheService.getScriptCache();
  const cached = cache.get("agency_info_cache_v3");
  if (cached) return { success: true, agency: JSON.parse(cached) };

  const masterSS = getMasterSS();
  const sheet = getSheet(masterSS, "Agency");
  const rows = genericListRows(sheet);
  const raw = rows.length > 0 ? rows[0] : {
    "নাম": "", "মোবাইল": "", "লোগো(URL)": "",
    "ব্যবস্থাপক তালিকা (JSON)": "[]",
    "Facebook Link": "", "Youtube Link": ""
  };

  let managers = [];
  try { managers = JSON.parse(raw["ব্যবস্থাপক তালিকা (JSON)"] || "[]"); } catch (e) { managers = []; }

  const info = {
    "নাম": raw["নাম"] || "",
    "মোবাইল": raw["মোবাইল"] || "",
    "লোগো(URL)": raw["লোগো(URL)"] || "",
    "managers": managers,
    "Facebook Link": raw["Facebook Link"] || "",
    "Youtube Link": raw["Youtube Link"] || ""
  };

  try { cache.put("agency_info_cache_v3", JSON.stringify(info), 60); } catch (e) { /* বাদ */ }

  return { success: true, agency: info };
}

function updateAgencyInfo(data) {
  const perm = checkAgencyPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };

  const masterSS = getMasterSS();
  const sheet = getSheet(masterSS, "Agency");
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rows = genericListRows(sheet);
  const existing = rows.length > 0 ? rows[0] : {};

  // নতুন লোগো ছবি (base64, ফ্রন্টএন্ডেই কমপ্রেসড) দেওয়া থাকলে সরাসরি সেভ
  let logoUrl = data.logoUrl || "";
  let photoWarning = "";
  if (data.logoBase64) {
    const processed = processImageForStorage(data.logoBase64);
    if (processed === null) { photoWarning = " (⚠️ লোগো ছবি খুব বড়, সেভ হয়নি — ছোট ছবি দিয়ে আবার চেষ্টা করুন)"; }
    else { logoUrl = processed; }
  }

  // ব্যবস্থাপক তালিকা — একাধিক ব্যবস্থাপক (নাম, পদবি, মোবাইল) JSON আকারে সংরক্ষণ
  const managersJson = data.managers ? JSON.stringify(data.managers) : (existing["ব্যবস্থাপক তালিকা (JSON)"] || "[]");

  const fieldMap = {
    "নাম": data.নাম,
    "মোবাইল": data.mobile,
    "লোগো(URL)": logoUrl,
    "ব্যবস্থাপক তালিকা (JSON)": managersJson,
    "Facebook Link": data.facebookLink,
    "Youtube Link": data.youtubeLink
  };

  const rowValues = headers.map(function (h) {
    if (h === "লোগো(URL)") return logoUrl || existing[h] || "";
    if (h === "ব্যবস্থাপক তালিকা (JSON)") return managersJson;
    const v = fieldMap[h];
    return (v !== undefined && v !== null && v !== "") ? v : (existing[h] || "");
  });

  if (rows.length === 0) {
    sheet.appendRow(rowValues);
  } else {
    sheet.getRange(2, 1, 1, headers.length).setValues([rowValues]);
  }
  sheet.getRange(2, headers.indexOf("মোবাইল") + 1).setNumberFormat("@");
  sheet.getRange(2, headers.indexOf("ব্যবস্থাপকের মোবাইল") + 1).setNumberFormat("@");
  invalidateAgencyCaches();

  return { success: true, message: "ডিপু তথ্য সংরক্ষিত হয়েছে" + photoWarning };
}

/*=========================================================
 *  ডিলার তালিকা / এডিট / স্ট্যাটাস (এজেন্সি-লেভেল)
 *=======================================================*/
function listDealers(data) {
  const perm = checkAgencyPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };

  const cache = CacheService.getScriptCache();
  const cached = cache.get("dealers_cache_v1");
  if (cached) return { success: true, dealers: JSON.parse(cached) };

  const masterSS = getMasterSS();
  const dealersSheet = getDealersSheet(masterSS);
  const nomineeSheet = getSheet(masterSS, "DealerNominee");

  const dealers = genericListRows(dealersSheet);
  const nominees = genericListRows(nomineeSheet);

  dealers.forEach(function (d) {
    d.nominee = nominees.find(function (n) { return n["DealerID"] === d["DealerID"]; }) || null;
  });

  try { cache.put("dealers_cache_v1", JSON.stringify(dealers), 30); } catch (e) { /* খুব বড় হলে বাদ */ }

  return { success: true, dealers: dealers };
}

function updateDealerInfo(data) {
  const perm = checkAgencyPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };

  const masterSS = getMasterSS();
  const sheet = getDealersSheet(masterSS);
  const ok = genericUpdateRow(sheet, "DealerID", data.dealerId, data.fields || {});
  invalidateAgencyCaches();

  return { success: ok, message: ok ? "ডিলার তথ্য আপডেট হয়েছে" : "ডিলার পাওয়া যায়নি" };
}

function setDealerStatus(data) {
  const perm = checkAgencyPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };

  const masterSS = getMasterSS();
  const sheet = getDealersSheet(masterSS);
  const ok = genericUpdateRow(sheet, "DealerID", data.dealerId, { "স্ট্যাটাস": data.status });
  invalidateAgencyCaches();

  return { success: ok, message: ok ? "স্ট্যাটাস পরিবর্তন হয়েছে" : "ডিলার পাওয়া যায়নি" };
}

function updateDealerNominee(data) {
  const perm = checkAgencyPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };

  const masterSS = getMasterSS();
  const sheet = getSheet(masterSS, "DealerNominee");
  const existingRowIndex = genericFindRowIndex(sheet, "DealerID", data.dealerId);

  if (existingRowIndex === -1) {
    const nomineeId = generateId(sheet, "N");
    genericAddRow(sheet, {
      "NomineeID": nomineeId,
      "DealerID": data.dealerId,
      "নমিনির নাম": data.fields["নমিনির নাম"] || "",
      "NID নং": data.fields["NID নং"] || "",
      "মোবাইল নং": data.fields["মোবাইল নং"] || "",
      "সম্পর্ক": data.fields["সম্পর্ক"] || ""
    });
  } else {
    genericUpdateRow(sheet, "DealerID", data.dealerId, data.fields || {});
  }

  invalidateAgencyCaches();
  return { success: true, message: "নমিনি তথ্য সংরক্ষিত হয়েছে" };
}

/*=========================================================
 *  সংস্থার তথ্য — শুধু এজেন্সি এন্ট্রি করবে (Master Spreadsheet এ),
 *  সব ডিলার একই তথ্য দেখতে ও প্রিন্ট করতে পারবে (এডিট করতে পারবে না)
 *=======================================================*/
const ABOUT_FIELDS = ["প্রতিষ্ঠাকাল", "উদ্দেশ্য", "প্রধান কার্যালয়ের ঠিকানা", "জেলা ডিপুর ঠিকানা", "নিয়মাবলি", "সংক্ষিপ্ত বিবরণ"];

function getAbout(data) {
  const perm = checkPermission(data.token, ["Admin", "প্রতিনিধি"]);
  if (!perm.ok) return { success: false, message: perm.message };

  const cache = CacheService.getScriptCache();
  const cached = cache.get("about_info_cache_v1");
  if (cached) return { success: true, about: JSON.parse(cached) };

  const masterSS = getMasterSS();
  const sheet = getSheet(masterSS, "AboutInfo");
  const rows = genericListRows(sheet);
  const info = rows.length > 0 ? rows[0] : {};
  ABOUT_FIELDS.forEach(function (f) { if (!(f in info)) info[f] = ""; });

  try { cache.put("about_info_cache_v1", JSON.stringify(info), 60); } catch (e) { /* বাদ */ }

  return { success: true, about: info };
}

function updateAbout(data) {
  const perm = checkAgencyPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };

  const masterSS = getMasterSS();
  const sheet = getSheet(masterSS, "AboutInfo");
  const rows = genericListRows(sheet);

  const fields = {
    "প্রতিষ্ঠাকাল": data.protisthakal || "",
    "উদ্দেশ্য": data.uddessho || "",
    "প্রধান কার্যালয়ের ঠিকানা": data.prodhanKaryaloy || "",
    "জেলা ডিপুর ঠিকানা": data.jelaDipur || "",
    "নিয়মাবলি": data.niyomaboli || "",
    "সংক্ষিপ্ত বিবরণ": data.songkhiptoBiboron || ""
  };

  if (rows.length === 0) {
    genericAddRow(sheet, fields);
  } else {
    sheet.getRange(2, 1, 1, ABOUT_FIELDS.length).setValues([ABOUT_FIELDS.map(function (f) { return fields[f]; })]);
  }

  invalidateAgencyCaches();
  return { success: true, message: "সংস্থার তথ্য সংরক্ষিত হয়েছে" };
}

/*******************************************************
 * একজন ডিলারের সম্পূর্ণ তথ্য (ডিলার + নমিনি + এডমিন লগইন) একসাথে
 * — ডিলার লিস্টের "এডিট" বাটনে ক্লিক করলে এটি কল হয়ে সব তথ্য
 * ফরমে লোড করে দেয়
 *******************************************************/
function getDealerFullInfo(data) {
  const perm = checkAgencyPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };

  const masterSS = getMasterSS();
  const dealersSheet = getDealersSheet(masterSS);
  const nomineeSheet = getSheet(masterSS, "DealerNominee");
  const usersSheet = getSheet(masterSS, "Users");

  const dealers = genericListRows(dealersSheet);
  const dealer = dealers.find(function (d) { return d["DealerID"] === data.dealerId; });
  if (!dealer) return { success: false, message: "ডিলার পাওয়া যায়নি" };

  const nominees = genericListRows(nomineeSheet);
  const nominee = nominees.find(function (n) { return n["DealerID"] === data.dealerId; }) || null;

  const users = genericListRows(usersSheet);
  const adminUser = users.find(function (u) {
    return u["DealerID"] === data.dealerId && u["রোল"] === "Admin";
  }) || null;

  return { success: true, dealer: dealer, nominee: nominee, adminUser: adminUser };
}

/*******************************************************
 * একজন ডিলারের সম্পূর্ণ তথ্য একসাথে আপডেট — ডিলার তথ্য, নমিনি তথ্য,
 * এবং (দেওয়া থাকলে) এডমিন লগইনের ইউজারনেম/পাসওয়ার্ড
 * data: {
 *   token, dealerId,
 *   dealerFields: { নাম, পিতার নাম, মোবাইল, Gmail, NID/জন্মসনদ,
 *                   ট্রেড লাইসেন্স নং, ঠিকানা, "ডিলারের ছবি(URL)"(ঐচ্ছিক) },
 *   nomineeFields: { নমিনির নাম, NID নং, মোবাইল নং, সম্পর্ক, "নমিনির ছবি(URL)"(ঐচ্ছিক) },
 *   adminUsername, adminPassword (ঐচ্ছিক — খালি রাখলে বদলাবে না)
 * }
 *******************************************************/
function updateDealerFull(data) {
  const perm = checkAgencyPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };

  const masterSS = getMasterSS();
  const dealersSheet = getDealersSheet(masterSS);
  ensureDealerAddrCols(dealersSheet);
  const usersSheet = getSheet(masterSS, "Users");

  const dealerFields = data.dealerFields || {};
  const nomineeFields = data.nomineeFields || {};
  let photoWarning = "";

  // নতুন ছবি (base64, ফ্রন্টএন্ডেই কমপ্রেসড) দেওয়া থাকলে সরাসরি সেভ
  if (data.dealerPhotoBase64) {
    const processed = processImageForStorage(data.dealerPhotoBase64);
    if (processed === null) { photoWarning += " (⚠️ ডিলারের নতুন ছবি খুব বড়, সেভ হয়নি)"; }
    else { dealerFields["ডিলারের ছবি(URL)"] = processed; }
  }
  if (data.nomineePhotoBase64) {
    const processedN = processImageForStorage(data.nomineePhotoBase64);
    if (processedN === null) { photoWarning += " (⚠️ নমিনির নতুন ছবি খুব বড়, সেভ হয়নি)"; }
    else { nomineeFields["নমিনির ছবি(URL)"] = processedN; }
  }

  // ১. ডিলারের মূল তথ্য আপডেট
  const dealerOk = genericUpdateRow(dealersSheet, "DealerID", data.dealerId, dealerFields);
  if (!dealerOk) return { success: false, message: "ডিলার পাওয়া যায়নি" };

  // ২. নমিনি তথ্য আপডেট (updateDealerNominee এর একই লজিক পুনঃব্যবহার)
  updateDealerNominee({ token: data.token, dealerId: data.dealerId, fields: nomineeFields });

  // ৩. এডমিন লগইন তথ্য (দেওয়া থাকলে) আপডেট
  if (data.adminUsername || data.adminPassword) {
    const users = usersSheet.getDataRange().getValues();
    const headers = users[0];
    const idxDealerId = headers.indexOf("DealerID");
    const idxRole = headers.indexOf("রোল");
    const idxUsername = headers.indexOf("ইউজারনেম");
    const idxPassword = headers.indexOf("পাসওয়ার্ড");

    for (let i = 1; i < users.length; i++) {
      if (users[i][idxDealerId] === data.dealerId && users[i][idxRole] === "Admin") {
        if (data.adminUsername) usersSheet.getRange(i + 1, idxUsername + 1).setValue(data.adminUsername);
        if (data.adminPassword) usersSheet.getRange(i + 1, idxPassword + 1).setValue(data.adminPassword);
        break;
      }
    }
  }

  invalidateAgencyCaches();
  invalidateDealerMsgInfo(data.dealerId);
  return { success: true, message: "ডিলারের সম্পূর্ণ তথ্য আপডেট হয়েছে" + photoWarning };
}

/*=========================================================
 *  ডিলার সাইট — ডিলার নিজের তথ্য দেখা ও আপডেট করা
 *  (ডিপু সাইটের "ডিলার" মেনুর মতোই কাজ করে, কিন্তু শুধু নিজের তথ্য)
 *
 *  নিরাপত্তা: dealerId সবসময় লগইন টোকেন থেকে নেওয়া হয়, ফ্রন্টএন্ডের
 *  পাঠানো কোনো dealerId গ্রহণ করা হয় না — তাই একজন ডিলার অন্যজনের
 *  তথ্য দেখতে/বদলাতে পারে না। স্ট্যাটাস, SpreadsheetID, DealerID বদলানো যায় না।
 *=======================================================*/
const SELF_DEALER_FIELDS = ["নাম", "পিতার নাম", "মোবাইল", "Gmail", "NID/জন্মসনদ", "ট্রেড লাইসেন্স নং", "এরিয়া", "ঠিকানা", "ইউনিয়ন/পৌরসভা", "উপজেলা", "জেলা", "Facebook Link"];
const SELF_NOMINEE_FIELDS = ["নমিনির নাম", "NID নং", "মোবাইল নং", "সম্পর্ক"];

function checkDealerSelfPermission(token) {
  const perm = checkPermission(token, ["Admin"]);
  if (!perm.ok) return perm;
  if (perm.payload.dealerId === AGENCY_ID) {
    return { ok: false, message: "এই সুবিধা শুধু ডিলার সাইটের জন্য" };
  }
  return perm;
}

/*******************************************************
 * লগইন করা ডিলারের Users শীটের রো খোঁজা (ইউজারনেম দিয়ে; না পেলে
 * ঐ ডিলারের প্রথম Admin) — শীটের রো নম্বর ফেরত দেয়, না পেলে -1
 *******************************************************/
function findOwnUserRow(values, dealerId, username) {
  const headers = values[0];
  const idxDealerId = headers.indexOf("DealerID");
  const idxUsername = headers.indexOf("ইউজারনেম");
  const idxRole = headers.indexOf("রোল");
  let fallback = -1;
  for (let i = 1; i < values.length; i++) {
    if (values[i][idxDealerId] !== dealerId) continue;
    if (values[i][idxUsername] === username) return i + 1;
    if (fallback === -1 && values[i][idxRole] === "Admin") fallback = i + 1;
  }
  return fallback;
}

/*******************************************************
 * data: { token } — নিজের ডিলার তথ্য + নমিনি + লগইন তথ্য
 *******************************************************/
function getMyDealerInfo(data) {
  const perm = checkDealerSelfPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };

  const dealerId = perm.payload.dealerId;
  const masterSS = getMasterSS();

  const dealer = genericListRows(getDealersSheet(masterSS)).find(function (d) {
    return d["DealerID"] === dealerId;
  });
  if (!dealer) return { success: false, message: "ডিলার পাওয়া যায়নি" };
  delete dealer["SpreadsheetID"]; // অভ্যন্তরীণ তথ্য ফ্রন্টএন্ডে পাঠানোর দরকার নেই

  const nominee = genericListRows(getSheet(masterSS, "DealerNominee")).find(function (n) {
    return n["DealerID"] === dealerId;
  }) || null;

  const usersSheet = getSheet(masterSS, "Users");
  const values = usersSheet.getDataRange().getValues();
  const rowNum = findOwnUserRow(values, dealerId, perm.payload.username);
  let adminUser = null;
  if (rowNum !== -1) {
    adminUser = {};
    values[0].forEach(function (h, idx) { adminUser[h] = values[rowNum - 1][idx]; });
  }

  return { success: true, dealer: dealer, nominee: nominee, adminUser: adminUser };
}

/*******************************************************
 * data: {
 *   token,
 *   dealerFields: { নাম, পিতার নাম, মোবাইল, Gmail, NID/জন্মসনদ, ট্রেড লাইসেন্স নং, ঠিকানা, Facebook Link },
 *   nomineeFields: { নমিনির নাম, NID নং, মোবাইল নং, সম্পর্ক },
 *   adminUsername, adminPassword, adminName  (খালি থাকলে বদলাবে না),
 *   dealerPhotoBase64, nomineePhotoBase64     (ঐচ্ছিক)
 * }
 * ইউজারনেম বদলালে নতুন টোকেন ফেরত আসে — ফ্রন্টএন্ডকে সেটা ব্যবহার করতে হবে
 *******************************************************/
function updateMyDealerFull(data) {
  const perm = checkDealerSelfPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };

  const dealerId = perm.payload.dealerId;
  const masterSS = getMasterSS();
  const dealersSheet = getDealersSheet(masterSS);
  ensureDealerAddrCols(dealersSheet);
  const nomineeSheet = getSheet(masterSS, "DealerNominee");
  const usersSheet = getSheet(masterSS, "Users");

  // শুধু অনুমোদিত ফিল্ডগুলো নেওয়া হয়
  const dealerFields = {};
  const nomineeFields = {};
  SELF_DEALER_FIELDS.forEach(function (k) {
    if (data.dealerFields && data.dealerFields.hasOwnProperty(k)) dealerFields[k] = data.dealerFields[k];
  });
  SELF_NOMINEE_FIELDS.forEach(function (k) {
    if (data.nomineeFields && data.nomineeFields.hasOwnProperty(k)) nomineeFields[k] = data.nomineeFields[k];
  });

  if (dealerFields.hasOwnProperty("নাম") && !String(dealerFields["নাম"]).trim()) {
    return { success: false, message: "ডিলারের নাম খালি রাখা যাবে না" };
  }

  // ইউজারনেম ডুপ্লিকেট চেক (কিছু সেভ করার আগেই, যাতে আধা-সেভ না হয়)
  const userValues = usersSheet.getDataRange().getValues();
  const uHeaders = userValues[0];
  const idxUsername = uHeaders.indexOf("ইউজারনেম");
  const idxPassword = uHeaders.indexOf("পাসওয়ার্ড");
  const idxName = uHeaders.indexOf("নাম");
  const ownRow = findOwnUserRow(userValues, dealerId, perm.payload.username);
  const newUsername = String(data.adminUsername || "").trim();

  if (newUsername) {
    for (let i = 1; i < userValues.length; i++) {
      if (userValues[i][idxUsername] === newUsername && (i + 1) !== ownRow) {
        return { success: false, message: "এই ইউজারনেম আগে থেকেই আছে, অন্য একটি দিন" };
      }
    }
  }

  let photoWarning = "";
  let newDealerPhoto = "";
  if (data.dealerPhotoBase64) {
    const processed = processImageForStorage(data.dealerPhotoBase64);
    if (processed === null) { photoWarning += " (⚠️ ডিলারের নতুন ছবি খুব বড়, সেভ হয়নি)"; }
    else { dealerFields["ডিলারের ছবি(URL)"] = processed; newDealerPhoto = processed; }
  }
  if (data.nomineePhotoBase64) {
    const processedN = processImageForStorage(data.nomineePhotoBase64);
    if (processedN === null) { photoWarning += " (⚠️ নমিনির নতুন ছবি খুব বড়, সেভ হয়নি)"; }
    else { nomineeFields["নমিনির ছবি(URL)"] = processedN; }
  }

  // ১. ডিলারের মূল তথ্য
  const dealerOk = genericUpdateRow(dealersSheet, "DealerID", dealerId, dealerFields);
  if (!dealerOk) return { success: false, message: "ডিলার পাওয়া যায়নি" };

  // ২. নমিনি তথ্য (না থাকলে নতুন রো, থাকলে আপডেট)
  if (Object.keys(nomineeFields).length > 0) {
    if (genericFindRowIndex(nomineeSheet, "DealerID", dealerId) === -1) {
      const row = { "NomineeID": generateId(nomineeSheet, "N"), "DealerID": dealerId };
      Object.keys(nomineeFields).forEach(function (k) { row[k] = nomineeFields[k]; });
      genericAddRow(nomineeSheet, row);
    } else {
      genericUpdateRow(nomineeSheet, "DealerID", dealerId, nomineeFields);
    }
  }

  // ৩. লগইন তথ্য (দেওয়া থাকলে)
  let newToken = "";
  if (ownRow !== -1) {
    if (newUsername) usersSheet.getRange(ownRow, idxUsername + 1).setValue(newUsername);
    if (data.adminPassword) usersSheet.getRange(ownRow, idxPassword + 1).setValue(data.adminPassword);
    if (data.adminName && idxName !== -1) usersSheet.getRange(ownRow, idxName + 1).setValue(data.adminName);

    if (newUsername && newUsername !== perm.payload.username) {
      newToken = createToken({
        dealerId: dealerId,
        role: perm.payload.role,
        username: newUsername,
        exp: perm.payload.exp
      });
    }
  }

  invalidateAgencyCaches();
  invalidateDealerMsgInfo(dealerId);
  return {
    success: true,
    message: "আপনার তথ্য আপডেট হয়েছে" + photoWarning,
    token: newToken,
    dealerPhotoUrl: newDealerPhoto
  };
}
