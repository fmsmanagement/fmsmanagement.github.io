/*******************************************************
 * TCB-স্টাইল ডিলার ম্যানেজমেন্ট সফটওয়্যার
 * Auth.gs — ডিপু লগইন / অথেনটিকেশন / সেশন টোকেন
 *
 * এই ফাইলটি MasterRegistry এর একই Apps Script প্রজেক্টে
 * Code.gs এর পাশে নতুন ফাইল হিসেবে যোগ করুন
 * (Apps Script এডিটরে + বাটনে ক্লিক করে "Script" সিলেক্ট করে
 * নাম দিন Auth, তারপর এই কোড পেস্ট করুন)
 *******************************************************/

// টোকেন সাইন করার জন্য গোপন চাবি — একবার তৈরি হয়ে ScriptProperties এ সেভ থাকবে
function getSecretKey() {
  const props = PropertiesService.getScriptProperties();
  let key = props.getProperty("SECRET_KEY");
  if (!key) {
    key = Utilities.getUuid() + Utilities.getUuid();
    props.setProperty("SECRET_KEY", key);
  }
  return key;
}


/*******************************************************
 * লগইন — শুধু ডিপু (AGENCY) ইউজার; ডিলার লগইন এই সাইটে চলবে না
 * data: { username, password }
 *******************************************************/
function loginUser(data) {
  const usersSheet = getSheet(getMasterSS(), "Users");
  const usersData = usersSheet.getDataRange().getValues();
  const headers = usersData[0];
  const idxUsername = headers.indexOf("ইউজারনেম");
  const idxPassword = headers.indexOf("পাসওয়ার্ড");
  const idxDealerId = headers.indexOf("DealerID");
  const idxRole = headers.indexOf("রোল");
  const idxName = headers.indexOf("নাম");

  let foundUser = null;
  for (let i = 1; i < usersData.length; i++) {
    if (usersData[i][idxUsername] === data.username && String(usersData[i][idxPassword]) === String(data.password)) {
      foundUser = usersData[i];
      break;
    }
  }
  if (!foundUser) return { success: false, message: "ভুল ইউজারনেম অথবা পাসওয়ার্ড" };
  if (foundUser[idxDealerId] !== AGENCY_ID) {
    return { success: false, message: "এই সাইটে শুধু ডিপু লগইন করতে পারবে" };
  }

  const token = createToken({
    dealerId: AGENCY_ID,
    role: foundUser[idxRole],
    username: data.username,
    exp: Date.now() + (12 * 60 * 60 * 1000) // ১২ ঘণ্টা মেয়াদ
  });
  return {
    success: true, token: token, role: foundUser[idxRole], name: foundUser[idxName],
    dealerId: AGENCY_ID, dealerName: "মূল ডিপু"
  };
}

/*******************************************************
 * Dealers ট্যাব থেকে DealerID দিয়ে তথ্য খোঁজা
 *******************************************************/
function getDealerRow(dealersSheet, dealerId) {
  const lastRow = dealersSheet.getLastRow();
  const lastCol = dealersSheet.getLastColumn();
  if (lastRow < 2) return null;

  const headers = dealersSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const idxDealerId = headers.indexOf("DealerID");
  if (idxDealerId === -1) return null;

  // পারফরম্যান্স: পুরো শীট (ডিলারদের base64 ছবিসহ) না পড়ে আগে শুধু ID কলাম
  // পড়ে সারি খুঁজে নেওয়া হয়, তারপর কেবল সেই একটি সারি পড়া হয়
  const ids = dealersSheet.getRange(2, idxDealerId + 1, lastRow - 1, 1).getValues();
  let rowNum = -1;
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === dealerId) { rowNum = i + 2; break; }
  }
  if (rowNum === -1) return null;

  const row = dealersSheet.getRange(rowNum, 1, 1, lastCol).getValues()[0];
  const idxName = headers.indexOf("নাম");
  const idxSpreadsheetId = headers.indexOf("SpreadsheetID");
  const idxStatus = headers.indexOf("স্ট্যাটাস");
  const idxPhoto = headers.indexOf("ডিলারের ছবি(URL)");
  const idxMobile = headers.indexOf("মোবাইল");
  const idxFacebook = headers.indexOf("Facebook Link");
  const idxUnion = headers.indexOf("ইউনিয়ন/পৌরসভা");
  const idxArea = headers.indexOf("এরিয়া");
  const idxSign = headers.indexOf("ডিলারের স্বাক্ষর(URL)");

  return {
    name: row[idxName],
    spreadsheetId: row[idxSpreadsheetId],
    status: row[idxStatus],
    photoUrl: idxPhoto !== -1 ? row[idxPhoto] : "",
    mobile: idxMobile !== -1 ? row[idxMobile] : "",
    facebookLink: idxFacebook !== -1 ? row[idxFacebook] : "",
    union: idxUnion !== -1 ? row[idxUnion] : "",
    area: idxArea !== -1 ? row[idxArea] : "",
    signatureUrl: idxSign !== -1 ? row[idxSign] : ""
  };
}

/*******************************************************
 * টোকেন তৈরি (HMAC সাইন করা, স্টেটলেস)
 *******************************************************/
function createToken(payload) {
  const payloadStr = JSON.stringify(payload);
  const payloadB64 = Utilities.base64EncodeWebSafe(payloadStr);
  const signature = Utilities.computeHmacSha256Signature(payloadB64, getSecretKey());
  const signatureB64 = Utilities.base64EncodeWebSafe(signature);
  return payloadB64 + "." + signatureB64;
}

/*******************************************************
 * টোকেন যাচাই — সঠিক হলে payload অবজেক্ট ফেরত দেয়, ভুল হলে null
 *******************************************************/
function verifyToken(token) {
  try {
    const parts = token.split(".");
    if (parts.length !== 2) return null;
    const payloadB64 = parts[0];
    const signatureB64 = parts[1];

    const expectedSignature = Utilities.computeHmacSha256Signature(payloadB64, getSecretKey());
    const expectedSignatureB64 = Utilities.base64EncodeWebSafe(expectedSignature);

    if (signatureB64 !== expectedSignatureB64) return null;

    const payloadStr = Utilities.newBlob(Utilities.base64DecodeWebSafe(payloadB64)).getDataAsString();
    const payload = JSON.parse(payloadStr);

    if (payload.exp && Date.now() > payload.exp) return null; // মেয়াদ শেষ

    return payload;
  } catch (e) {
    return null;
  }
}

/*******************************************************
 * পারমিশন চেক
 * allowedRoles: ["Admin"] অথবা ["Admin", "প্রতিনিধি"]
 *******************************************************/
function checkPermission(token, allowedRoles) {
  const payload = verifyToken(token);
  if (!payload) {
    return { ok: false, message: "সেশন মেয়াদোত্তীর্ণ, আবার লগইন করুন" };
  }
  if (allowedRoles.indexOf(payload.role) === -1) {
    return { ok: false, message: "এই কাজ করার অনুমতি আপনার নেই" };
  }
  return { ok: true, payload: payload };
}

/*******************************************************
 * নতুন ইউজার (প্রতিনিধি/এডমিন) যোগ করা — শুধু Admin করতে পারবে
 * data: { token, username, password, name, mobile, role }
 *******************************************************/
function addUser(data) {
  const perm = checkAgencyPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };

  const masterSS = getMasterSS();
  const usersSheet = getSheet(masterSS, "Users");

  const userId = generateId(usersSheet, "U");
  usersSheet.appendRow([
    userId,
    perm.payload.dealerId,
    data.username,
    data.password,
    data.role, // "Admin" অথবা "প্রতিনিধি"
    data.name,
    data.mobile
  ]);
  usersSheet.getRange(usersSheet.getLastRow(), 7).setNumberFormat("@");

  return { success: true, userId: userId, message: "ইউজার যোগ হয়েছে" };
}


