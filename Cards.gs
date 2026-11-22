/*******************************************************
 * TCB-স্টাইল ডিলার ম্যানেজমেন্ট সফটওয়্যার
 * Cards.gs — গ্রাহক কার্ড দেওয়া হয়েছে কিনা (গ্রাহক তালিকার "কার্ড" চেকবক্স)
 *
 * MasterRegistry এর একই Apps Script প্রজেক্টে নতুন ফাইল হিসেবে যোগ করুন
 * (নাম দিন: Cards)। ডিলার সাইটে Admin + প্রতিনিধি ব্যবহার করতে পারবে।
 * Customers শীটে "কার্ড" কলাম না থাকলে নিজে থেকে যোগ হয়
 *
 * data: { token, customerIds: [...], given: true|false }
 *   given=true  → "কার্ড" কলামে "হ্যাঁ" (কার্ড দেওয়া হয়েছে)
 *   given=false → কলাম খালি
 *******************************************************/
function setCustomerCards(data) {
  const perm = checkPermission(data.token, ["Admin", "প্রতিনিধি"]);
  if (!perm.ok) return { success: false, message: perm.message };
  const ids = Array.isArray(data.customerIds) ? data.customerIds : [];
  if (ids.length === 0) return { success: false, message: "কোনো গ্রাহক নির্বাচন করা হয়নি" };

  const ss = getDealerSpreadsheet(perm.payload.dealerId);
  const sheet = getSheet(ss, "Customers");
  ensureSheetColumns(sheet, ["কার্ড"]);

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, message: "গ্রাহক পাওয়া যায়নি" };
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const iId = headers.indexOf("CustomerID"), iCard = headers.indexOf("কার্ড");
    if (iId === -1 || iCard === -1) return { success: false, message: "কার্ড কলাম পাওয়া যায়নি" };

    const n = lastRow - 1;
    const idCol = sheet.getRange(2, iId + 1, n, 1).getValues();
    const cardCol = sheet.getRange(2, iCard + 1, n, 1).getValues();
    const want = {};
    ids.forEach(function (id) { want[String(id)] = true; });
    const value = data.given ? "হ্যাঁ" : "";
    let count = 0;
    for (let r = 0; r < n; r++) {
      if (want[String(idCol[r][0])]) { cardCol[r][0] = value; count++; }
    }
    sheet.getRange(2, iCard + 1, n, 1).setValues(cardCol);
    return { success: true, updated: count, message: "আপডেট হয়েছে" };
  } finally {
    lock.releaseLock();
  }
}

/*******************************************************
 * গ্রাহক তালিকার "WhatsApp আছে" চেকবক্স — যাদের নম্বরে WhatsApp আছে
 * Customers শীটে "WhatsApp আছে" কলাম না থাকলে নিজে থেকে যোগ হয়
 * (আগের "WhatsApp" কলামটি মেসেজ পাঠানোর স্ট্যাটাসের — সেটা আলাদা)
 *
 * data: { token, customerIds: [...], has: true|false }
 *   has=true  → "WhatsApp আছে" কলামে "হ্যাঁ"
 *   has=false → কলাম খালি
 *******************************************************/
function setCustomerWhatsApp(data) {
  const perm = checkPermission(data.token, ["Admin", "প্রতিনিধি"]);
  if (!perm.ok) return { success: false, message: perm.message };
  const ids = Array.isArray(data.customerIds) ? data.customerIds : [];
  if (ids.length === 0) return { success: false, message: "কোনো গ্রাহক নির্বাচন করা হয়নি" };

  const ss = getDealerSpreadsheet(perm.payload.dealerId);
  const sheet = getSheet(ss, "Customers");
  ensureSheetColumns(sheet, ["WhatsApp আছে"]);

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, message: "গ্রাহক পাওয়া যায়নি" };
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const iId = headers.indexOf("CustomerID"), iWa = headers.indexOf("WhatsApp আছে");
    if (iId === -1 || iWa === -1) return { success: false, message: "WhatsApp আছে কলাম পাওয়া যায়নি" };

    const n = lastRow - 1;
    const idCol = sheet.getRange(2, iId + 1, n, 1).getValues();
    const waCol = sheet.getRange(2, iWa + 1, n, 1).getValues();
    const want = {};
    ids.forEach(function (id) { want[String(id)] = true; });
    const value = data.has ? "হ্যাঁ" : "";
    let count = 0;
    for (let r = 0; r < n; r++) {
      if (want[String(idCol[r][0])]) { waCol[r][0] = value; count++; }
    }
    sheet.getRange(2, iWa + 1, n, 1).setValues(waCol);
    return { success: true, updated: count, message: "আপডেট হয়েছে" };
  } finally {
    lock.releaseLock();
  }
}

/*******************************************************
 * গ্রাহকের "স্ট্যাটাস" (এক্টিভ / ডিএক্টিভ) — গ্রাহক তালিকার ড্রপডাউন
 * Customers শীটে "স্ট্যাটাস" কলাম না থাকলে নিজে থেকে যোগ হয় (খালি = এক্টিভ)
 * ডিএক্টিভ গ্রাহক বিক্রি তালিকায় দেখায় না এবং মেসেজ পায় না
 *
 * data: { token, customerIds: [...], status: "এক্টিভ" | "ডিএক্টিভ" }
 *******************************************************/
const CUSTOMER_STATUS_DEACTIVE = "ডিএক্টিভ";
function isCustomerDeactive(c) {
  return String(c["স্ট্যাটাস"] === undefined || c["স্ট্যাটাস"] === null ? "" : c["স্ট্যাটাস"]).trim() === CUSTOMER_STATUS_DEACTIVE;
}

function setCustomerStatus(data) {
  const perm = checkPermission(data.token, ["Admin", "প্রতিনিধি"]);
  if (!perm.ok) return { success: false, message: perm.message };
  const ids = Array.isArray(data.customerIds) ? data.customerIds : [];
  if (ids.length === 0) return { success: false, message: "কোনো গ্রাহক নির্বাচন করা হয়নি" };
  const status = data.status === CUSTOMER_STATUS_DEACTIVE ? CUSTOMER_STATUS_DEACTIVE : "এক্টিভ";

  const ss = getDealerSpreadsheet(perm.payload.dealerId);
  const sheet = getSheet(ss, "Customers");
  ensureSheetColumns(sheet, ["স্ট্যাটাস"]);

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, message: "গ্রাহক পাওয়া যায়নি" };
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const iId = headers.indexOf("CustomerID"), iSt = headers.indexOf("স্ট্যাটাস");
    if (iId === -1 || iSt === -1) return { success: false, message: "স্ট্যাটাস কলাম পাওয়া যায়নি" };

    const n = lastRow - 1;
    const idCol = sheet.getRange(2, iId + 1, n, 1).getValues();
    const stCol = sheet.getRange(2, iSt + 1, n, 1).getValues();
    const want = {};
    ids.forEach(function (id) { want[String(id)] = true; });
    let count = 0;
    for (let r = 0; r < n; r++) {
      if (want[String(idCol[r][0])]) { stCol[r][0] = status; count++; }
    }
    sheet.getRange(2, iSt + 1, n, 1).setValues(stCol);
    return { success: true, updated: count, message: "আপডেট হয়েছে" };
  } finally {
    lock.releaseLock();
  }
}
