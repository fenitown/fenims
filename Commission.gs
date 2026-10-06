/*******************************************************
 * TCB-স্টাইল ডিলার ম্যানেজমেন্ট সফটওয়্যার
 * Commission.gs — কমিশন (এজেন্সি থেকে ডিলারকে কমিশন প্রদানের হিসাব)
 *
 * MasterRegistry এর একই Apps Script প্রজেক্টে নতুন ফাইল
 * হিসেবে যোগ করুন (নাম দিন: Commission)
 * সব একশন শুধুমাত্র মূল এজেন্সি (AGENCY) লগইন করতে পারবে
 *******************************************************/

function addCommission(data) {
  const perm = checkAgencyPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };

  const masterSS = getMasterSS();
  const commissionSheet = getSheet(masterSS, "Commission");
  const dealersSheet = getSheet(masterSS, "Dealers");

  const dealerInfo = getDealerRow(dealersSheet, data.dealerId);
  if (!dealerInfo) return { success: false, message: "ডিলার পাওয়া যায়নি" };

  // Dealers ট্যাব থেকে মোবাইল ও ঠিকানা অটো নেওয়া
  const dealersData = dealersSheet.getDataRange().getValues();
  const headers = dealersData[0];
  const idxId = headers.indexOf("DealerID");
  const idxMobile = headers.indexOf("মোবাইল");
  const idxThikana = headers.indexOf("ঠিকানা");
  let mobile = "", thikana = "";
  for (let i = 1; i < dealersData.length; i++) {
    if (dealersData[i][idxId] === data.dealerId) {
      mobile = dealersData[i][idxMobile];
      thikana = dealersData[i][idxThikana];
      break;
    }
  }

  const commissionId = generateId(commissionSheet, "CM");
  genericAddRow(commissionSheet, {
    "CommissionID": commissionId,
    "DealerID": data.dealerId,
    "ডিলারের নাম": dealerInfo.name,
    "মোবাইল নং": mobile,
    "ঠিকানা": thikana,
    "টাকার পরিমাণ": data.amount,
    "তারিখ": new Date()
  });
  commissionSheet.getRange(commissionSheet.getLastRow(), 4).setNumberFormat("@");

  return { success: true, commissionId: commissionId, message: "কমিশন যোগ হয়েছে" };
}

function listCommissions(data) {
  const perm = checkAgencyPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };

  const masterSS = getMasterSS();
  const sheet = getSheet(masterSS, "Commission");
  return { success: true, commissions: genericListRows(sheet) };
}

function updateCommission(data) {
  const perm = checkAgencyPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };

  const masterSS = getMasterSS();
  const sheet = getSheet(masterSS, "Commission");
  const ok = genericUpdateRow(sheet, "CommissionID", data.commissionId, data.fields || {});
  return { success: ok, message: ok ? "আপডেট হয়েছে" : "কমিশন রেকর্ড পাওয়া যায়নি" };
}

function deleteCommission(data) {
  const perm = checkAgencyPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };

  const masterSS = getMasterSS();
  const sheet = getSheet(masterSS, "Commission");
  const ok = genericDeleteRow(sheet, "CommissionID", data.commissionId);
  return { success: ok, message: ok ? "ডিলিট হয়েছে" : "কমিশন রেকর্ড পাওয়া যায়নি" };
}
