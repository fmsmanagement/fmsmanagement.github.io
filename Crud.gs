/*******************************************************
 * TCB-স্টাইল ডিলার ম্যানেজমেন্ট সফটওয়্যার
 * Crud.gs — জেনেরিক শীট হেল্পার, প্যাকেজ ও এজেন্সি ড্যাশবোর্ড (শুধু ডিপু)
 *
 * এই ফাইলটিও MasterRegistry এর একই Apps Script প্রজেক্টে
 * নতুন ফাইল হিসেবে যোগ করুন (নাম দিন: Crud)
 *******************************************************/




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









/*******************************************************
 * গ্রাহক ডিলিট হলে অন্য গ্রাহকদের আইডি আর বদলায় না (আগের মতো ১ থেকে নতুন করে
 * সাজানো হয় না)।
 *
 * ডিলিট হওয়া গ্রাহকের বিক্রির রেকর্ড মুছে ফেলা হয় না (রিপোর্ট ও স্টকের হিসাব ঠিক
 * থাকে) — শুধু ওই রেকর্ডের CustomerID "DELETED" করে দেওয়া হয়। এতে ডিলিট হওয়া
 * আইডি (যেমন সর্বশেষ গ্রাহকের আইডি) পরে নতুন কোনো গ্রাহকের সাথে ভুলে জুড়ে যায় না
 *******************************************************/


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

/* প্যাকেজ শীটের নতুন কলাম — প্রতিটি পণ্য-লাইনের ProductID ও প্যাকেজ-প্রতি পরিমাণ (শীটের শেষে যোগ হয়) */
function packageExtraHeaders() {
  const h = [];
  for (let i = 1; i <= MAX_PACKAGE_ITEMS; i++) h.push("পণ্য" + i + " - ProductID", "পণ্য" + i + " - সংখ্যা");
  return h;
}

function addPackage(data) {
  const perm = checkAgencyPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };

  const masterSS = getMasterSS();
  const pkgSheet = getPackagesSheet(masterSS);
  ensureSheetColumns(pkgSheet, packageExtraHeaders());   // পুরনো শীটে না থাকলে নিজে যোগ হয়
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
      row["পণ্য" + n + " - ProductID"] = item.productId || "";
      row["পণ্য" + n + " - সংখ্যা"] = String(item.qty == null ? "" : item.qty).trim();
    } else {
      row["পণ্য" + n + " - নাম ও পরিমাণ"] = "";
      row["পণ্য" + n + " - বাজার মূল্য"] = "";
      row["পণ্য" + n + " - কম্বো মূল্য"] = "";
      row["পণ্য" + n + " - সাশ্রয়"] = "";
      row["পণ্য" + n + " - ProductID"] = "";
      row["পণ্য" + n + " - সংখ্যা"] = "";
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
  CacheService.getScriptCache().remove("packages_cache_v6");
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
        "সাশ্রয়": pkg["পণ্য" + i + " - সাশ্রয়"],
        "ProductID": pkg["পণ্য" + i + " - ProductID"] || "",
        "পণ্যের সংখ্যা": pkg["পণ্য" + i + " - সংখ্যা"] === undefined || pkg["পণ্য" + i + " - সংখ্যা"] === null ? "" : pkg["পণ্য" + i + " - সংখ্যা"]   // প্যাকেজ-প্রতি এই পণ্যের পরিমাণ (যেমন "১ কেজি", "৫০০ গ্রাম")
      });
    }
  }
  pkg.items = items;
  pkg.itemCount = items.length;
  pkg["সর্বমোট মূল্য"] = pkg["সর্বমোট কম্বো মূল্য (সর্বমোট মূল্য)"];
  return pkg;
}

function listPackages(data) {
  const perm = checkAgencyPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };

  const cache = CacheService.getScriptCache();
  const cached = cache.get("packages_cache_v6");
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
    cache.put("packages_cache_v6", JSON.stringify(packages), 300); // ৫ মিনিট (যোগ/এডিট/ডিলিটে নিজে থেকে মুছে যায়)
  } catch (e) { /* ক্যাশ সাইজ বেশি হলে চুপচাপ বাদ */ }

  return { success: true, packages: packages };
}

function updatePackage(data) {
  const perm = checkAgencyPermission(data.token);
  if (!perm.ok) return { success: false, message: perm.message };
  const masterSS = getMasterSS();
  const sheet = getPackagesSheet(masterSS);
  ensureSheetColumns(sheet, packageExtraHeaders());

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
        fields["পণ্য" + n + " - ProductID"] = item.productId || "";
        fields["পণ্য" + n + " - সংখ্যা"] = String(item.qty == null ? "" : item.qty).trim();
      } else {
        fields["পণ্য" + n + " - নাম ও পরিমাণ"] = "";
        fields["পণ্য" + n + " - বাজার মূল্য"] = "";
        fields["পণ্য" + n + " - কম্বো মূল্য"] = "";
        fields["পণ্য" + n + " - সাশ্রয়"] = "";
        fields["পণ্য" + n + " - ProductID"] = "";
        fields["পণ্য" + n + " - সংখ্যা"] = "";
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

