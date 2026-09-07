function getUser_(userId) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.USERS);
  if (!sh) throw new Error("Users シートが見つかりません。");
  const v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    if (String(v[i][0]).trim() === userId) {
      return {
        row: i + 1,
        userId: String(v[i][0]).trim(),
        name: String(v[i][1] || "").trim(),
        balance: Number(v[i][2] || 0),
        isActive: toBool_(v[i][3]),
      };
    }
  }
  return null;
}

function setUserBalance_(userId, balance) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.USERS);
  const u = getUser_(userId);
  if (!u) throw new Error("ユーザーが見つかりません。");
  sh.getRange(u.row, 3).setValue(balance);
}

function getShop_(shopId) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.SHOPS);
  if (!sh) throw new Error("Shops シートが見つかりません。");
  const v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    if (String(v[i][0]).trim() === shopId) {
      return {
        row: i + 1,
        shopId: String(v[i][0]).trim(),
        shopName: String(v[i][1] || "").trim(),
        balance: Number(v[i][2] || 0),
        isActive: toBool_(v[i][3]),
      };
    }
  }
  return null;
}

function setShopBalance_(shopId, balance) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.SHOPS);
  const s = getShop_(shopId);
  if (!s) throw new Error("店が見つかりません。");
  sh.getRange(s.row, 3).setValue(balance);
}

function listUsers_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.USERS);
  const v = sh.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < v.length; i++) {
    if (!v[i][0]) continue;
    out.push({
      userId: String(v[i][0]).trim(),
      name: String(v[i][1] || "").trim(),
      balance: Number(v[i][2] || 0),
      isActive: toBool_(v[i][3]),
    });
  }
  return out;
}

function listShops_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.SHOPS);
  const v = sh.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < v.length; i++) {
    if (!v[i][0]) continue;
    out.push({
      shopId: String(v[i][0]).trim(),
      shopName: String(v[i][1] || "").trim(),
      balance: Number(v[i][2] || 0),
      isActive: toBool_(v[i][3]),
    });
  }
  return out;
}

function appendTx_(obj) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.TX);
  if (!sh) throw new Error("Tx シートが見つかりません。");

  sh.appendRow([
    obj.txId,
    obj.at,
    obj.type,
    obj.userId,
    obj.userName,
    obj.shopId,
    obj.shopName,
    obj.amount,
    obj.status,
    obj.note,
    obj.meta,
  ]);
}

function listTx_(limit) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.TX);
  const v = sh.getDataRange().getValues();
  const rows = v.slice(1).filter(r => r[0]);
  const tail = rows.slice(Math.max(0, rows.length - limit)).reverse();
  return tail.map(r => ({
    txId: r[0], at: r[1], type: r[2],
    userId: r[3], userName: r[4],
    shopId: r[5], shopName: r[6],
    amount: r[7], status: r[8],
    note: r[9], meta: r[10],
  }));
}

function toBool_(x){
  const s = String(x).toUpperCase();
  return (x === true || x === 1 || s === "TRUE" || s === "1" || s === "YES");
}
