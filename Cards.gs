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
