/*******************************************************
 * TCB-স্টাইল ডিলার ম্যানেজমেন্ট সফটওয়্যার
 * Messaging.gs — ডিলার নিজের গ্রাহকদের কাছে SMS ও WhatsApp এ
 * একসাথে বাল্ক মেসেজ পাঠানো
 *
 * MasterRegistry এর একই Apps Script প্রজেক্টে নতুন ফাইল
 * হিসেবে যোগ করুন (নাম দিন: Messaging)
 *
 * ================= গুরুত্বপূর্ণ =================
 * এই মেসেজিং ফিচার শুধুমাত্র ডিলার সাইটে থাকে — প্রত্যেক ডিলার
 * নিজের গ্রাহকদের মেসেজ পাঠানোর জন্য নিজের নিজের SMS/WhatsApp
 * API নিজে সেট করে নেবে (গ্রাহক → মেসেজিং সেটাপ থেকে, শুধু Admin
 * রোল দেখতে/বদলাতে পারবে)। প্রতিটি ডিলারের কনফিগ আলাদা আলাদাভাবে
 * সংরক্ষিত থাকে, একজনেরটা আরেকজনের সাথে মিশবে না। মূল এজেন্সি
 * (ডিপু) সাইটে কোনো মেসেজিং সুবিধা নেই।
 *
 * Google Apps Script নিজে থেকে SMS/WhatsApp পাঠাতে পারে না —
 * SMS এর জন্য দুইটা অপশন আছে (ডিলার সেটাপ ফরম থেকে যেকোনো একটা
 * বেছে নেবে):
 *
 * === অপশন A: নিজের ফোনের সিম দিয়ে — Traccar SMS Gateway (Play Store) ===
 * Google Play Store থেকে সরাসরি "Traccar SMS Gateway" অ্যাপ ইনস্টল করা
 * যায়, কোনো সাইডলোডের ঝামেলা নেই। অ্যাপ খুলে SMS পারমিশন দিন, সেটিংসে
 * "Cloud" মোড চালু করুন (Local/LAN মোড Apps Script থেকে কাজ করবে
 * না, কারণ Apps Script গুগলের সার্ভারে চলে, আপনার ফোনের একই WiFi
 * নেটওয়ার্কে না) — তখন একটা URL ও Token পাবেন, সেটাপ ফরমে বসান।
 * মেসেজ পাঠালে ঠিক ম্যানুয়ালি ওই ফোন থেকে পাঠানোর মতোই সেই সিমের
 * প্যাকেজ/ব্যালেন্স থেকে কাটবে — সম্পূর্ণ ফ্রি সফটওয়্যার, শুধু
 * ফোনের সিমের স্বাভাবিক SMS খরচ লাগবে।
 *
 * === অপশন B: পেইড SMS গেটওয়ে (BulkSMSBD ইত্যাদি) ===
 *    http://bulksmsbd.net/api/smsapi?api_key={api_key}&type=text&number={to}&senderid={from}&message={message}
 *    URL এর ভেতরে {api_key}, {from}, {to}, {message} — এই চারটি
 *    প্লেসহোল্ডার ঠিক এভাবেই থাকতে হবে, বাকিটা প্রতিটি গেটওয়ের
 *    নিজস্ব ডকুমেন্টেশন অনুযায়ী বসবে।
 * ২) SMS API Key — গেটওয়ে থেকে পাওয়া চাবি
 * ৩) WhatsApp API Endpoint URL — যেমন Meta এর অফিসিয়াল
 *    WhatsApp Business Cloud API:
 *    https://graph.facebook.com/v20.0/<PHONE_NUMBER_ID>/messages
 * ৪) WhatsApp Access Token — Meta/প্রোভাইডার থেকে পাওয়া টোকেন
 *
 * imo মেসেজিং সাপোর্ট করা হয়নি — imo কোনো পাবলিক/বিজনেস API
 * সরবরাহ করে না।
 *
 * SMS ও WhatsApp সম্পূর্ণ স্বাধীন — একটি সেটআপ না থাকলেও অন্যটি
 * কাজ করবে (কনফিগ ফাঁকা থাকলে সেই চ্যানেলটি শুধু স্কিপ হবে)।
 *******************************************************/

const MESSAGING_CONFIG_KEYS = ["smsRate", "smsProviderType", "smsGatewayUrlTemplate", "smsApiKey", "smsSenderId", "smsTraccarUrl", "whatsappApiUrl", "whatsappToken"];
// smsProviderType সম্ভাব্য মান:
//  "traccar"      — নিজের Android ফোনের সিম দিয়ে (Traccar SMS Gateway অ্যাপ — Play Store
//                   থেকে সরাসরি ইনস্টল করা যায়, APK সাইডলোডের ঝামেলা নেই)
//  "url_template" — পেইড SMS গেটওয়ে (BulkSMSBD ইত্যাদি), URL টেমপ্লেট দিয়ে GET রিকোয়েস্ট
// (আগে textbee অপশন ছিল, এখন বাদ — পুরনো "textbee" সেভ করা থাকলে সেটা "traccar"
//  হিসেবে ধরা হয়, এবং Traccar URL/Token না থাকলে "সেটআপ হয়নি" দেখাবে)

/*******************************************************
 * SMS প্রোভাইডারের ধরন ও প্রস্তুত আছে কিনা — একই লজিক সব জায়গায় ব্যবহার হয়
 *******************************************************/
function getSmsProviderType(config) {
  return config.smsProviderType === "url_template" ? "url_template" : "traccar";
}

function isSmsReady(config) {
  return getSmsProviderType(config) === "traccar"
    ? !!(config.smsTraccarUrl && config.smsApiKey)
    : !!config.smsGatewayUrlTemplate;
}

function isWhatsAppReady(config) {
  return !!(config.whatsappApiUrl && config.whatsappToken);
}

/*******************************************************
 * ScriptProperties থেকে একটি নির্দিষ্ট ডিলারের মেসেজিং কনফিগ পড়া
 * (প্রতিটি ডিলারের জন্য আলাদা কী দিয়ে সংরক্ষিত)
 *******************************************************/
function getDealerMessagingConfigRaw(dealerId) {
  const props = PropertiesService.getScriptProperties();
  const config = {};
  MESSAGING_CONFIG_KEYS.forEach(function (k) {
    config[k] = props.getProperty("MSG_" + dealerId + "_" + k) || "";
  });
  return config;
}

/*******************************************************
 * ডিলার — নিজের মেসেজিং কনফিগ দেখা (শুধু Admin)
 * data: { token }
 *******************************************************/
function getDealerMessagingConfig(data) {
  const perm = checkPermission(data.token, ["Admin"]);
  if (!perm.ok) return { success: false, message: perm.message };
  return { success: true, config: getDealerMessagingConfigRaw(perm.payload.dealerId) };
}

/*******************************************************
 * ডিলার — নিজের মেসেজিং কনফিগ সংরক্ষণ (শুধু Admin)
 * data: { token, smsProviderType, smsGatewayUrlTemplate, smsApiKey, smsSenderId,
 *         smsTraccarUrl, whatsappApiUrl, whatsappToken }
 *******************************************************/
function saveDealerMessagingConfig(data) {
  const perm = checkPermission(data.token, ["Admin"]);
  if (!perm.ok) return { success: false, message: perm.message };

  const props = PropertiesService.getScriptProperties();
  const dealerId = perm.payload.dealerId;
  MESSAGING_CONFIG_KEYS.forEach(function (k) {
    props.setProperty("MSG_" + dealerId + "_" + k, data[k] || "");
  });

  return { success: true, message: "মেসেজিং সেটিংস সংরক্ষিত হয়েছে" };
}

/*******************************************************
 * "মেসেজ" ফরম খোলার সময় ব্যবহারের জন্য হালকা স্ট্যাটাস —
 * Admin ও প্রতিনিধি দুজনেই পড়তে পারবে (আসল API Key/Token
 * শেয়ার না করে শুধু দরকারি তথ্যটুকু দেওয়া হয়, যাতে মেসেজ ফরমে
 * ঠিকভাবে বোঝা যায় গেটওয়ে সেটআপ করা আছে কিনা)
 *******************************************************/
function getMessagingStatus(data) {
  const perm = checkPermission(data.token, ["Admin", "প্রতিনিধি"]);
  if (!perm.ok) return { success: false, message: perm.message };

  const config = getDealerMessagingConfigRaw(perm.payload.dealerId);

  return {
    success: true,
    smsProviderType: getSmsProviderType(config),
    smsSenderId: config.smsSenderId || "",
    smsReady: isSmsReady(config),
    waReady: isWhatsAppReady(config),
    smsRate: Number(config.smsRate) > 0 ? Number(config.smsRate) : 0.35   // প্রতি SMS (টাকা)
  };
}

/*******************************************************
 * বাংলাদেশি লোকাল নম্বর (01XXXXXXXXX) কে আন্তর্জাতিক ফরম্যাটে
 * (৮৮01XXXXXXXXX) রূপান্তর — অধিকাংশ BD SMS/WhatsApp গেটওয়ে এই
 * ফরম্যাটই চায়। ইতিমধ্যে + বা ৮৮ দিয়ে শুরু থাকলে অপরিবর্তিত রাখা হয়
 *******************************************************/
function normalizeMobileBD(mobile) {
  let m = String(mobile === undefined || mobile === null ? "" : mobile);
  m = m.replace(/[০-৯]/g, function (d) { return "০১২৩৪৫৬৭৮৯".indexOf(d); }); // বাংলা অঙ্ক → ইংরেজি
  m = m.replace(/[^0-9+]/g, "");
  if (m.indexOf("+") === 0) m = m.substring(1);
  if (m.indexOf("880") === 0) return m;
  if (m.indexOf("0") === 0) return "88" + m;
  // শীটে শুরুর ০ কেটে গিয়ে ১০ ডিজিট (1712345678) থাকলে ০ বসিয়ে নেওয়া
  if (m.length === 10 && m.charAt(0) === "1") return "880" + m;
  return m;
}

function isValidBdMobile(normalized) {
  return /^8801[0-9]{9}$/.test(normalized);
}

/*******************************************************
 * একটি নম্বরে SMS গেটওয়ে দিয়ে মেসেজ পাঠানো
 * config.smsGatewayUrlTemplate এ {api_key},{from},{to},{message}
 * প্লেসহোল্ডারগুলো আসল মান দিয়ে বদলে দিয়ে GET রিকোয়েস্ট পাঠানো হয়
 * {from} এর জায়গায় Sender ID (BulkSMSBD এর ক্ষেত্রে এটা মোবাইল
 * নম্বর না, তাদের দেওয়া Approved Sender ID কোড) বসে
 *******************************************************/
function sendSmsViaGateway(config, fromMobile, toMobile, message) {
  if (!config.smsGatewayUrlTemplate) return { ok: false, skipped: true };
  try {
    const url = config.smsGatewayUrlTemplate
      .replace(/\{api_key\}/g, encodeURIComponent(config.smsApiKey || ""))
      .replace(/\{from\}/g, encodeURIComponent(fromMobile || ""))
      .replace(/\{to\}/g, encodeURIComponent(normalizeMobileBD(toMobile)))
      .replace(/\{message\}/g, encodeURIComponent(message));
    const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const code = resp.getResponseCode();
    return { ok: code >= 200 && code < 300, response: resp.getContentText() };
  } catch (e) {
    return { ok: false, reason: e.toString() };
  }
}

/*******************************************************
 * একটি নম্বরে WhatsApp Business Cloud API (Meta) ফরম্যাটে
 * মেসেজ পাঠানো — অন্য প্রোভাইডার ব্যবহার করলে payload বদলাতে হবে
 *******************************************************/
function sendWhatsAppMessage(config, toMobile, message) {
  if (!config.whatsappApiUrl || !config.whatsappToken) return { ok: false, skipped: true };
  try {
    const payload = {
      messaging_product: "whatsapp",
      to: normalizeMobileBD(toMobile),
      type: "text",
      text: { body: message }
    };
    const resp = UrlFetchApp.fetch(config.whatsappApiUrl, {
      method: "post",
      contentType: "application/json",
      headers: { Authorization: "Bearer " + config.whatsappToken },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    const code = resp.getResponseCode();
    return { ok: code >= 200 && code < 300, response: resp.getContentText() };
  } catch (e) {
    return { ok: false, reason: e.toString() };
  }
}

/*******************************************************
 * Traccar SMS Gateway (Play Store অ্যাপ, cloud মোডে) — নিজের
 * Android ফোনকে SMS গেটওয়ে বানিয়ে সেই ফোনের সিম দিয়ে মেসেজ
 * পাঠানো। প্রতি নম্বরে আলাদা POST রিকোয়েস্ট লাগে (bulk সাপোর্ট নেই)
 *******************************************************/
function sendSmsViaTraccar(config, toMobile, message) {
  if (!config.smsTraccarUrl || !config.smsApiKey) return { ok: false, skipped: true };
  try {
    const payload = { to: "+" + normalizeMobileBD(toMobile), message: message };
    const resp = UrlFetchApp.fetch(config.smsTraccarUrl, {
      method: "post",
      contentType: "application/json",
      headers: { Authorization: config.smsApiKey },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    const code = resp.getResponseCode();
    return { ok: code >= 200 && code < 300, response: resp.getContentText() };
  } catch (e) {
    return { ok: false, reason: e.toString() };
  }
}

/*******************************************************
 * ডিলারের নাম ও ঠিকানা (Master "Dealers" ট্যাব থেকে) — মেসেজের নিচে অটো বসে
 *******************************************************/
function getDealerNameAndAddress(dealerId) {
  const cache = CacheService.getScriptCache();
  const key = "dealer_nameaddr2_" + dealerId;
  try {
    const hit = cache.get(key);
    if (hit) return JSON.parse(hit);
  } catch (e) { /* ক্যাশ না থাকলে সরাসরি পড়া হবে */ }

  const sheet = getDealersSheet(getMasterSS());
  const rows = genericListRows(sheet);
  let info = { name: "", area: "", address: "" };
  for (let i = 0; i < rows.length; i++) {
    if (rows[i]["DealerID"] === dealerId) {
      info = {
        name: String(rows[i]["নাম"] || "").trim(),
        area: String(rows[i]["এরিয়া"] || "").trim(),
        address: String(rows[i]["ঠিকানা"] || "").trim()
      };
      break;
    }
  }
  try { cache.put(key, JSON.stringify(info), 300); } catch (e) { /* বাদ */ }
  return info;
}

/*******************************************************
 * চূড়ান্ত মেসেজ তৈরি — সব মেসেজে শুধু গ্রাহকের নাম অটো বসে (প্রথম লাইনে):
 *
 *   <গ্রাহকের নাম>
 *   <ডিলারের লেখা পুরো মেসেজ>
 *
 * ডিলারের নাম/এরিয়া আর অটো বসে না — ডিলার নিজে মেসেজের ভেতরেই লিখবেন।
 * (dealerInfo ও prefixName প্যারামিটার আগের কলগুলোর সাথে মিল রাখতে থাকলো, এখন ব্যবহার হয় না)
 *******************************************************/
function composeCustomerMessage(customerName, mainText, dealerInfo, prefixName) {
  const nameLine = String(customerName || "").trim() ? String(customerName).trim() + "\n" : "";
  return nameLine + String(mainText).trim();
}

/*******************************************************
 * ঠিকানা থেকে "সোনাগাজী", "ফেনী" অংশ বাদ দেওয়া (কমা দিয়ে আলাদা অংশ ধরে)
 * যেমন "৮নং আমিরাবাদ ইউনিয়ন, সোনাগাজী, ফেনী" → "৮নং আমিরাবাদ ইউনিয়ন"
 *******************************************************/
function cleanDealerAddress(address) {
  return String(address || "")
    .split(/[,،]/)
    .map(function (p) { return p.trim(); })
    .filter(function (p) {
      return p && !/^(সোনাগাজী|সোনাগাজি|ফেনী|ফেনি)(\s*(উপজেলা|জেলা))?$/.test(p);
    })
    .join(", ");
}

/* মেসেজের নিচের স্বাক্ষর: ডিলারের নাম + এরিয়া (দুই লাইন)।
   এরিয়া ফাঁকা থাকলে (পুরনো ডিলার) আগের মতো ঠিকানা থেকে সোনাগাজী/ফেনী বাদ দিয়ে বসে */
function buildSignature(info) {
  const place = String(info.area || "").trim() || cleanDealerAddress(info.address);
  return [String(info.name || "").trim(), place]
    .filter(function (s) { return s; })
    .join("\n");
}

/*******************************************************
 * একাধিক প্রাপককে (ডুপ্লিকেট/খালি নম্বর বাদ দিয়ে) SMS + WhatsApp —
 * যেটা কনফিগার করা আছে সেটাই পাঠানো হবে, অন্যটা স্কিপ হবে।
 * প্রতিটি গ্রাহকের নাম আলাদা বলে প্রতি নম্বরে আলাদা মেসেজ তৈরি হয়।
 * recipients: [ { mobile, name }, ... ]
 *******************************************************/
function sendBulkToRecipients(config, fromMobile, mainText, recipients, dealerInfo, prefixName, channels) {
  channels = channels || { sms: true, wa: true };
  const providerType = getSmsProviderType(config);
  const smsEnabled = !!channels.sms && isSmsReady(config);
  const waEnabled = !!channels.wa && isWhatsAppReady(config);

  // ডুপ্লিকেট/খালি/ভুল নম্বর বাদ — একই নম্বরে একাধিক গ্রাহক থাকলে সবার আইডি মনে রাখা হয়
  const unique = [];
  const byNorm = {};
  const invalidIds = [];
  let invalid = 0;
  (recipients || []).forEach(function (r) {
    const clean = normalizeMobileBD(r && r.mobile);
    if (!clean || !isValidBdMobile(clean)) {
      if (clean) invalid++;
      if (r && r.customerId) invalidIds.push(r.customerId);
      return;
    }
    if (!byNorm[clean]) {
      byNorm[clean] = { norm: clean, name: r.name, ids: [] };
      unique.push(byNorm[clean]);
    }
    if (r.customerId) byNorm[clean].ids.push(r.customerId);
  });

  // সব রিকোয়েস্ট আগে তৈরি করে একসাথে (প্যারালাল) পাঠানো হয় — একটা একটা করে পাঠালে অনেক ধীর
  const smsReqs = [], waReqs = [];
  unique.forEach(function (rcp) {
    const text = composeCustomerMessage(rcp.name, mainText, dealerInfo, prefixName);
    if (smsEnabled) {
      if (providerType === "traccar") {
        smsReqs.push({
          url: config.smsTraccarUrl, method: "post", contentType: "application/json",
          headers: { Authorization: config.smsApiKey },
          payload: JSON.stringify({ to: "+" + rcp.norm, message: text }),
          muteHttpExceptions: true
        });
      } else {
        smsReqs.push({
          url: config.smsGatewayUrlTemplate
            .replace(/\{api_key\}/g, encodeURIComponent(config.smsApiKey || ""))
            .replace(/\{from\}/g, encodeURIComponent(fromMobile || ""))
            .replace(/\{to\}/g, encodeURIComponent(rcp.norm))
            .replace(/\{message\}/g, encodeURIComponent(text)),
          method: "get", muteHttpExceptions: true
        });
      }
    }
    if (waEnabled) {
      waReqs.push({
        url: config.whatsappApiUrl, method: "post", contentType: "application/json",
        headers: { Authorization: "Bearer " + config.whatsappToken },
        payload: JSON.stringify({ messaging_product: "whatsapp", to: rcp.norm, type: "text", text: { body: text } }),
        muteHttpExceptions: true
      });
    }
  });

  const smsRes = smsEnabled ? fetchAllSafe(smsReqs) : [];
  const waRes = waEnabled ? fetchAllSafe(waReqs) : [];

  let smsSent = 0, smsFailed = 0, waSent = 0, waFailed = 0;
  let smsLastError = "", waLastError = "";
  const smsByCustomer = {}, waByCustomer = {};   // customerId → true/false (গেটওয়ে গ্রহণ করেছে কিনা)
  smsRes.forEach(function (r, i) {
    if (r.ok) smsSent++; else { smsFailed++; smsLastError = String(r.response || r.reason || "").substring(0, 300); }
    unique[i].ids.forEach(function (id) { smsByCustomer[id] = !!r.ok; });
  });
  waRes.forEach(function (r, i) {
    if (r.ok) waSent++; else { waFailed++; waLastError = String(r.response || r.reason || "").substring(0, 300); }
    unique[i].ids.forEach(function (id) { waByCustomer[id] = !!r.ok; });
  });

  return {
    total: unique.length, invalid: invalid,
    firstNumber: unique.length ? unique[0].norm : "",
    smsEnabled: smsEnabled, smsSent: smsSent, smsFailed: smsFailed, smsLastError: smsLastError,
    waEnabled: waEnabled, waSent: waSent, waFailed: waFailed, waLastError: waLastError,
    smsByCustomer: smsByCustomer, waByCustomer: waByCustomer, invalidIds: invalidIds
  };
}

/*******************************************************
 * UrlFetchApp.fetchAll দিয়ে প্যারালালে পাঠানো (২০টি করে ব্যাচ)।
 * ফেরত: প্রতিটি রিকোয়েস্টের জন্য { ok, response, reason }
 *******************************************************/
function fetchAllSafe(requests) {
  const out = [];
  const CHUNK = 20;
  for (let i = 0; i < requests.length; i += CHUNK) {
    const chunk = requests.slice(i, i + CHUNK);
    try {
      const resps = UrlFetchApp.fetchAll(chunk);
      resps.forEach(function (resp) {
        const code = resp.getResponseCode();
        out.push({ ok: code >= 200 && code < 300, response: resp.getContentText() });
      });
    } catch (e) {
      // ব্যাচে কোনো একটা রিকোয়েস্ট ভুল হলে পুরো ব্যাচ ব্যর্থ হয় — একটা একটা করে আবার চেষ্টা
      chunk.forEach(function (req) {
        try {
          const resp = UrlFetchApp.fetch(req.url, req);
          const code = resp.getResponseCode();
          out.push({ ok: code >= 200 && code < 300, response: resp.getContentText() });
        } catch (e2) {
          out.push({ ok: false, reason: e2.toString() });
        }
      });
    }
  }
  return out;
}

/*******************************************************
 * ডিলার সাইট — নিজের গ্রাহকদের মেসেজ পাঠানো (Admin + প্রতিনিধি)
 * ডিলার নিজের সেট করা কনফিগ ব্যবহার করেই পাঠানো হয়
 * data: { token, fromMobile, message (শুধু মূল লেখা), type: "all"|"location"|"ward"|"single" (single এ customerId লাগে), channel: "whatsapp"|"sms_all"|"sms_nowa" (একটাই), baseMessageId (তালিকা থেকে বাছা আগের মেসেজের আইডি, ঐচ্ছিক), locations: [...] (প্রাপ্তির স্থানের নাম অথবা ওয়ার্ড নং) }
 * গ্রাহকের নাম এবং ডিলারের নাম/ঠিকানা সার্ভার নিজেই বসিয়ে দেয়
 *******************************************************/
function sendCustomerMessage(data) {
  const perm = checkPermission(data.token, ["Admin", "প্রতিনিধি"]);
  if (!perm.ok) return { success: false, message: perm.message };

  if (!data.message || !String(data.message).trim()) return { success: false, message: "মেসেজ লিখুন" };

  const dealerId = perm.payload.dealerId;

  // একই অনুরোধ দুইবার এলে (ডাবল ক্লিক/ধীর নেটওয়ার্কে আবার চাপা) দ্বিতীয়টি আটকে দেওয়া হয়
  if (data.requestId) {
    const lock = LockService.getScriptLock();
    lock.waitLock(5000);
    try {
      const cache = CacheService.getScriptCache();
      const key = "msgreq_" + dealerId + "_" + data.requestId;
      if (cache.get(key)) return { success: false, message: "এই মেসেজটি ইতিমধ্যে পাঠানো হয়েছে বা পাঠানো চলছে" };
      cache.put(key, "1", 600);
    } finally {
      lock.releaseLock();
    }
  }
  const config = getDealerMessagingConfigRaw(dealerId);
  const providerType = getSmsProviderType(config);
  if (!isSmsReady(config) && !isWhatsAppReady(config)) {
    return { success: false, message: "এখনো কোনো SMS/WhatsApp সেটআপ করা হয়নি। গ্রাহক → মেসেজিং সেটাপ থেকে আগে সেটআপ করুন।" };
  }
  // শুধু url_template (পেইড গেটওয়ে) মোডে Sender ID আবশ্যক — traccar এ
  // ফোনের নিজের নম্বর স্বয়ংক্রিয়ভাবে ব্যবহৃত হয়, আলাদা করে লাগে না
  if (providerType === "url_template" && !data.fromMobile) {
    return { success: false, message: "Sender ID / মোবাইল নং দিন" };
  }

  // মাধ্যম: শুধু একটাই — whatsapp | sms_all | sms_nowa (শুধু যাদের WhatsApp যায়নি তাদের SMS)।
  // দুটোতে একসাথে কখনো যায় না। মাধ্যম না এলে WhatsApp (সেটআপ থাকলে), নইলে SMS
  let channel = data.channel;
  if (["whatsapp", "sms_all", "sms_nowa"].indexOf(channel) === -1) {
    channel = isWhatsAppReady(config) ? "whatsapp" : "sms_all";
  }
  const wantWA = channel === "whatsapp";
  const wantSMS = !wantWA;
  if (channel === "whatsapp" && !isWhatsAppReady(config)) {
    return { success: false, message: "WhatsApp এখনো সেটআপ করা হয়নি। গ্রাহক → মেসেজ → মেসেজ সেটাপ থেকে সেটআপ করুন।" };
  }
  if ((channel === "sms_all" || channel === "sms_nowa") && !isSmsReady(config)) {
    return { success: false, message: "SMS এখনো সেটআপ করা হয়নি। গ্রাহক → মেসেজ → মেসেজ সেটাপ থেকে সেটআপ করুন।" };
  }

  const ss = getDealerSpreadsheet(dealerId);
  const custSheet = getSheet(ss, "Customers");
  ensureSheetColumns(custSheet, ["WhatsApp", "SMS", "মেসেজ আইডি"]);
  const customers = genericListRows(custSheet);

  // তিন ধরনের মেসেজ: সকল গ্রাহক / প্রাপ্তির স্থান ভিত্তিক / ওয়ার্ড ভিত্তিক (+ নতুন গ্রাহকের একক মেসেজ)
  let targeted = customers;
  let statusText = "সকল গ্রাহক";
  const picked = (data.locations || []).map(function (x) { return String(x).trim(); });
  if (data.type === "location" && picked.length) {
    targeted = customers.filter(function (c) {
      return picked.indexOf(String(c["প্রাপ্তির স্থান"] || "").trim()) !== -1;
    });
    statusText = picked.join(", ");
  } else if (data.type === "single" && data.customerId) {
    // নতুন গ্রাহক সংরক্ষণের পর শুধু সেই একজনকে মেসেজ
    targeted = customers.filter(function (c) { return c["CustomerID"] === data.customerId; });
    statusText = "নতুন গ্রাহক";
  } else if (data.type === "ward" && picked.length) {
    targeted = customers.filter(function (c) {
      return picked.indexOf(String(c["ওয়ার্ড নং"] === undefined || c["ওয়ার্ড নং"] === null ? "" : c["ওয়ার্ড নং"]).trim()) !== -1;
    });
    statusText = "ওয়ার্ড: " + picked.join(", ");
  }

  const mainText = String(data.message).trim();
  const msgSheet = getMessagesSheet(ss);
  const rec = resolveMessageRecord(msgSheet, mainText, data.type, statusText, data.baseMessageId);

  // "যাদের নিকট WhatsApp মেসেজ যায়নি" — এই একই মেসেজ WhatsApp এ যাদের কাছে গেছে তাদের বাদ
  if (channel === "sms_nowa") {
    targeted = targeted.filter(function (c) {
      return !(c["মেসেজ আইডি"] === rec.id && c["WhatsApp"] === MSG_STATUS_OK);
    });
    if (targeted.length === 0) {
      return {
        success: true, total: 0, invalid: 0, smsEnabled: true, smsSent: 0, smsFailed: 0, waEnabled: false, waSent: 0, waFailed: 0,
        message: "সবার কাছেই WhatsApp এ মেসেজ গেছে — SMS পাঠানোর কেউ নেই"
      };
    }
  }

  const recipients = targeted.map(function (c) {
    return { mobile: c["মোবাইল নং"], name: c["নাম"], customerId: c["CustomerID"] };
  });
  const dealerInfo = getDealerNameAndAddress(dealerId);
  const result = sendBulkToRecipients(
    config, data.fromMobile, mainText, recipients, dealerInfo,
    data.type === "single" && !!data.includeName,
    { sms: wantSMS, wa: wantWA }
  );

  // গ্রাহক তালিকায় WhatsApp/SMS কার কাছে গেছে তা আপডেট (সর্বশেষ মেসেজ অনুযায়ী)
  try {
    updateCustomerMessageStatus(custSheet, rec.id, targeted, result, result.waEnabled, result.smsEnabled);
  } catch (e) { /* স্ট্যাটাস আপডেট ব্যর্থ হলে পাঠানোর ফলাফল আটকাবে না */ }

  // মেসেজ তালিকায় সংরক্ষণ — একই মেসেজ একবারই (হাজার গ্রাহকের কাছে গেলেও একটি এন্ট্রি)
  try {
    if (rec.rowNum) {
      // আগের মেসেজ আবার (অন্য এরিয়ায়/মাধ্যমে) পাঠালে নতুন সারি হয় না — একই এন্ট্রি আপডেট হয়
      if (data.type === "single" || rec.viaBase) {
        const idxTime = rec.headers.indexOf("তারিখ ও সময়");
        if (idxTime !== -1) msgSheet.getRange(rec.rowNum, idxTime + 1).setValue(new Date());
      }
      if (rec.viaBase) {
        // স্ট্যাটাসে নতুন প্রাপক-ধরন যোগ (যেমন "ওয়ার্ড: ৩ | ওয়ার্ড: ৪")
        const idxStatus = rec.headers.indexOf("স্ট্যাটাস");
        if (idxStatus !== -1) {
          const cell = msgSheet.getRange(rec.rowNum, idxStatus + 1);
          const parts = String(cell.getValue()).split(" | ").filter(function (p) { return p; });
          if (parts.indexOf(statusText) === -1) { parts.push(statusText); cell.setValue(parts.join(" | ")); }
        }
      }
    } else {
      genericAddRow(msgSheet, {
        "MessageID": rec.id,
        "মেসেজ": mainText,
        "তারিখ ও সময়": new Date(),
        "ধরন": data.type === "location" ? "প্রাপ্তির স্থান" : (data.type === "ward" ? "ওয়ার্ড" : (data.type === "single" ? MSG_TYPE_NEW_CUSTOMER : "সকল গ্রাহক")),
        "স্ট্যাটাস": statusText,
        "মোট প্রাপক": result.total
      });
    }
  } catch (e) { /* লগ সংরক্ষণ ব্যর্থ হলে বাদ */ }

  // ফ্রন্টএন্ডে বড় ম্যাপ পাঠানোর দরকার নেই
  delete result.smsByCustomer; delete result.waByCustomer; delete result.invalidIds;
  return Object.assign({ success: true, message: "মেসেজ পাঠানো সম্পন্ন হয়েছে", messageId: rec.id }, result);
}

const MSG_STATUS_OK = "গেছে";
const MSG_STATUS_FAIL = "যায়নি";
const MSG_TYPE_NEW_CUSTOMER = "নতুন গ্রাহক";
const MSG_REUSE_WINDOW_MS = 6 * 60 * 60 * 1000;   // একই মেসেজের ফলো-আপ (WhatsApp→SMS) ৬ ঘণ্টার মধ্যে ধরা হয়

/*******************************************************
 * শীটে কলাম না থাকলে শেষে যোগ করে (বিদ্যমান ডিলারদের শীটের জন্য)
 *******************************************************/
function ensureSheetColumns(sheet, names) {
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  let col = headers.filter(function (h) { return h !== ""; }).length;
  names.forEach(function (n) {
    if (headers.indexOf(n) === -1) {
      col++;
      sheet.getRange(1, col).setValue(n).setFontWeight("bold");
    }
  });
}

/*******************************************************
 * এই মেসেজটি মেসেজ তালিকায় আগে থেকে আছে কিনা — থাকলে সেই আইডিই ব্যবহার
 *  • নতুন-গ্রাহকের মেসেজ: একই লেখা থাকলে (সময় যাই হোক) একই এন্ট্রি
 *  • অন্য মেসেজ: সর্বশেষ এন্ট্রির লেখা ও প্রাপক-ধরন একই এবং ৬ ঘণ্টার মধ্যে হলে একই এন্ট্রি
 *    (যেমন WhatsApp পাঠিয়ে পরে একই মেসেজ SMS এ পাঠানো)
 *******************************************************/
function resolveMessageRecord(msgSheet, text, type, statusText, baseId) {
  const values = msgSheet.getDataRange().getValues();
  const headers = values[0] || [];
  const iId = headers.indexOf("MessageID"), iText = headers.indexOf("মেসেজ"), iTime = headers.indexOf("তারিখ ও সময়");
  const iType = headers.indexOf("ধরন"), iStatus = headers.indexOf("স্ট্যাটাস");
  const isWelcome = function (t) { return t === MSG_TYPE_NEW_CUSTOMER || t === "একক গ্রাহক"; };
  const wantWelcome = (type === "single");

  // ডিলার তালিকা থেকে আগের মেসেজ বেছে (লেখা না বদলিয়ে) পাঠালে সেই এন্ট্রিই ব্যবহার
  if (!wantWelcome && baseId) {
    for (let i = values.length - 1; i >= 1; i--) {
      if (values[i][iId] === baseId && !isWelcome(values[i][iType]) && String(values[i][iText]).trim() === text) {
        return { id: baseId, rowNum: i + 1, headers: headers, viaBase: true };
      }
    }
  }

  for (let i = values.length - 1; i >= 1; i--) {
    const rowWelcome = isWelcome(values[i][iType]);
    if (wantWelcome) {
      if (rowWelcome && String(values[i][iText]).trim() === text) {
        return { id: values[i][iId], rowNum: i + 1, headers: headers };
      }
      continue;
    }
    if (rowWelcome) continue;
    // সর্বশেষ নন-স্বাগত এন্ট্রিটাই শুধু পরীক্ষা
    const t = new Date(values[i][iTime]).getTime();
    const statusParts = String(values[i][iStatus]).split(" | ");
    if (String(values[i][iText]).trim() === text && statusParts.indexOf(statusText) !== -1 &&
        !isNaN(t) && (Date.now() - t) < MSG_REUSE_WINDOW_MS) {
      return { id: values[i][iId], rowNum: i + 1, headers: headers };
    }
    break;
  }
  const newId = "MS" + Utilities.formatDate(new Date(), "UTC", "yyMMddHHmmss") + Math.floor(100 + Math.random() * 900);
  return { id: newId, rowNum: 0, headers: headers };
}

/*******************************************************
 * Customers শীটে WhatsApp / SMS / মেসেজ আইডি কলাম আপডেট।
 * গ্রাহকের শীটে সংরক্ষিত আইডি বর্তমান মেসেজের আইডির সাথে না মিললে
 * (অর্থাৎ নতুন মেসেজ) আগের স্ট্যাটাস মুছে নতুন করে শুরু — তাই সবসময়
 * সর্বশেষ মেসেজের অবস্থাই দেখায়।
 *******************************************************/
function updateCustomerMessageStatus(custSheet, msgId, targeted, result, waAttempted, smsAttempted) {
  const lastRow = custSheet.getLastRow();
  if (lastRow < 2) return;
  const lastCol = custSheet.getLastColumn();
  const headers = custSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const iId = headers.indexOf("CustomerID");
  const iWa = headers.indexOf("WhatsApp"), iSms = headers.indexOf("SMS"), iMid = headers.indexOf("মেসেজ আইডি");
  if (iId === -1 || iWa === -1 || iSms === -1 || iMid === -1) return;

  const n = lastRow - 1;
  const idCol = custSheet.getRange(2, iId + 1, n, 1).getValues();
  const waCol = custSheet.getRange(2, iWa + 1, n, 1).getValues();
  const smsCol = custSheet.getRange(2, iSms + 1, n, 1).getValues();
  const midCol = custSheet.getRange(2, iMid + 1, n, 1).getValues();

  const target = {};
  targeted.forEach(function (c) { target[c["CustomerID"]] = true; });
  const bad = {};
  (result.invalidIds || []).forEach(function (id) { bad[id] = true; });

  for (let r = 0; r < n; r++) {
    const cid = idCol[r][0];
    if (!target[cid]) continue;
    if (midCol[r][0] !== msgId) { waCol[r][0] = ""; smsCol[r][0] = ""; midCol[r][0] = msgId; }
    if (waAttempted) waCol[r][0] = (!bad[cid] && result.waByCustomer[cid] === true) ? MSG_STATUS_OK : MSG_STATUS_FAIL;
    if (smsAttempted) smsCol[r][0] = (!bad[cid] && result.smsByCustomer[cid] === true) ? MSG_STATUS_OK : MSG_STATUS_FAIL;
  }

  custSheet.getRange(2, iWa + 1, n, 1).setValues(waCol);
  custSheet.getRange(2, iSms + 1, n, 1).setValues(smsCol);
  custSheet.getRange(2, iMid + 1, n, 1).setValues(midCol);
}

/*******************************************************
 * "Messages" ট্যাব — আগের ডিলারদের শীটে না থাকলে নিজে তৈরি হবে
 *******************************************************/
function getMessagesSheet(ss) {
  let sheet = ss.getSheetByName("Messages");
  if (!sheet) {
    ensureSheetsWithHeaders(ss, { "Messages": DEALER_SHEETS_DEF["Messages"] });
    sheet = ss.getSheetByName("Messages");
  }
  return sheet;
}

/*******************************************************
 * পাঠানো মেসেজের তালিকা (পুরনো → নতুন ক্রমে) — Admin + প্রতিনিধি
 * data: { token }
 *******************************************************/
function listMessages(data) {
  const perm = checkPermission(data.token, ["Admin", "প্রতিনিধি"]);
  if (!perm.ok) return { success: false, message: perm.message };

  const ss = getDealerSpreadsheet(perm.payload.dealerId);
  const sheet = getMessagesSheet(ss);
  return { success: true, messages: genericListRows(sheet) };
}

/*******************************************************
 * মেসেজ তালিকা থেকে একটি মেসেজ ডিলিট (শুধু তালিকার রেকর্ড মোছে,
 * আগে পাঠানো মেসেজ ফেরত যায় না) — Admin + প্রতিনিধি
 * data: { token, messageId }
 *******************************************************/
function deleteMessage(data) {
  const perm = checkPermission(data.token, ["Admin", "প্রতিনিধি"]);
  if (!perm.ok) return { success: false, message: perm.message };

  const ss = getDealerSpreadsheet(perm.payload.dealerId);
  const sheet = getMessagesSheet(ss);
  const ok = genericDeleteRow(sheet, "MessageID", data.messageId);
  return { success: ok, message: ok ? "মেসেজ ডিলিট হয়েছে" : "মেসেজ পাওয়া যায়নি" };
}

/*******************************************************
 * একবারের অনুমতি (UrlFetchApp / external_request) নেওয়ার জন্য —
 * এডিটর থেকে একবার Run করলে "Connect to an external service"
 * পারমিশন চাইবে। এরপর চাইলে এটা রেখে দিতে পারেন
 *******************************************************/
function authorizeExternalRequest() {
  UrlFetchApp.fetch("https://www.google.com", { muteHttpExceptions: true });
  Logger.log("অনুমতি সম্পন্ন");
}
