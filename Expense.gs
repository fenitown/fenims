/*******************************************************
 * TCB-স্টাইল ডিলার ম্যানেজমেন্ট সফটওয়্যার
 * Expense.gs — খরচ (মূল এজেন্সির নিজস্ব খরচের হিসাব, ভাউচার-ভিত্তিক)
 *
 * MasterRegistry এর একই Apps Script প্রজেক্টে নতুন ফাইল
 * হিসেবে যোগ করুন (নাম দিন: Expense)
 * সব একশন শুধুমাত্র মূল এজেন্সি (AGENCY) লগইন করতে পারবে
 *
 * data.entries: [ { বিবরণ, টাকা }, ... ] — একই ভাউচারে একাধিক
 * সারি একসাথে যোগ করা যাবে; ভাউচারের তারিখ ও পরিশোধ একবারই দেওয়া হয়
 *******************************************************/

function addExpense(data) {
  const perm = checkAgencyPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };

  const masterSS = getMasterSS();
  const sheet = getSheet(masterSS, "Expense");
  const voucherNo = "EXV" + Utilities.formatString("%06d", Math.floor(Math.random() * 900000) + 100000);
  const date = data.date ? new Date(data.date) : new Date();

  const entries = data.entries || [];
  let grandTotal = 0;
  entries.forEach(function (e) { grandTotal += Number(e["টাকা"]) || 0; });
  const paid = Number(data.paid) || 0;
  const due = grandTotal - paid;

  const rows = entries.map(function (e, i) {
    return {
      "ভাউচার নং": voucherNo,
      "তারিখ": date,
      "ক্রম": i + 1,
      "বিবরণ": e["বিবরণ"],
      "টাকা": e["টাকা"],
      "সর্বমোট": grandTotal,
      "পরিশোধ": paid,
      "বকেয়া": due
    };
  });

  batchAppendRows(sheet, rows, "EXE", "EntryID");

  return {
    success: true, voucherNo: voucherNo, total: grandTotal, paid: paid, due: due,
    message: rows.length + " টি খরচ এন্ট্রি যোগ হয়েছে"
  };
}

function listExpenses(data) {
  const perm = checkAgencyPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };

  const masterSS = getMasterSS();
  const sheet = getSheet(masterSS, "Expense");
  return { success: true, expenses: genericListRows(sheet) };
}

/*******************************************************
 * একটি নির্দিষ্ট খরচ সারি (লাইন) এডিট করা — বিবরণ/টাকা
 *******************************************************/
function updateExpense(data) {
  const perm = checkAgencyPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };

  const masterSS = getMasterSS();
  const sheet = getSheet(masterSS, "Expense");
  const ok = genericUpdateRow(sheet, "EntryID", data.entryId, data.fields || {});
  return { success: ok, message: ok ? "আপডেট হয়েছে" : "খরচ এন্ট্রি পাওয়া যায়নি" };
}

/*******************************************************
 * পুরো ভাউচার (তার সব সারি) ডিলিট করা
 *******************************************************/
function deleteExpense(data) {
  const perm = checkAgencyPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };

  const masterSS = getMasterSS();
  const sheet = getSheet(masterSS, "Expense");
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const idx = headers.indexOf("ভাউচার নং");

  let found = false;
  for (let i = values.length - 1; i >= 1; i--) {
    if (values[i][idx] === data.voucherNo) {
      sheet.deleteRow(i + 1);
      found = true;
    }
  }
  return { success: found, message: found ? "ভাউচার ডিলিট হয়েছে" : "খরচ ভাউচার পাওয়া যায়নি" };
}
