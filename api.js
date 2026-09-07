function api_appInfo() {
  return {
    appName: getConfig_("APP_NAME", "ClassPay"),
    allowGoogleChartQr: boolConfig_("ALLOW_GOOGLE_CHART_QR", true),
    nowJst: Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy-MM-dd HH:mm:ss"),
  };
}

/* ===== 共通ユーティリティ ===== */
function _headerMap_(headerRow) {
  const h = headerRow.map(x => String(x || "").trim().toLowerCase());
  const idx = (name) => h.indexOf(String(name).toLowerCase());
  return { h, idx };
}
/**
 * Txシートから userId の履歴を新しい順で返す
 * header: txId at type userId userName shopId shopName amount status note meta
 */
/**
 * Txシートから userId の履歴を新しい順で返す（Date/文字列混在対応）
 * header: txId at type userId userName shopId shopName amount status note meta
 */
function _getTxByUserId_(userId, limit) {
  limit = limit || 30;
  userId = String(userId || "").trim().toUpperCase();

  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.TX);
  if (!sh) return [];

  const values = sh.getDataRange().getValues();
  if (values.length <= 1) return [];

  const { idx } = _headerMap_(values[0]);

  const iTxId     = idx("txId");
  const iAt       = idx("at");
  const iType     = idx("type");
  const iUserId   = idx("userId");
  const iUserName = idx("userName");
  const iShopId   = idx("shopId");
  const iShopName = idx("shopName");
  const iAmount   = idx("amount");
  const iStatus   = idx("status");
  const iNote     = idx("note");
  const iMeta     = idx("meta"); // 無い場合もある

  if (iUserId < 0 || iAt < 0) return [];

  // ★ at を “確実に数値時刻” にする
  function _toTime_(v){
    if (!v && v !== 0) return 0;

    // Date型
    if (Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v)) {
      return v.getTime();
    }

    const s = String(v).trim();
    if (!s) return 0;

    // "yyyy-MM-dd HH:mm:ss" を最優先でパース（JST想定）
    const m = s.match(/^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2})(?::(\d{2}))?/);
    if (m){
      const Y = Number(m[1]), Mo = Number(m[2]) - 1, D = Number(m[3]);
      const H = Number(m[4]), Mi = Number(m[5]), Se = Number(m[6] || 0);
      // Apps ScriptはスクリプトTZ依存になりがちなので、ここは Date(Y,Mo,...) のローカル生成でOK
      return new Date(Y, Mo, D, H, Mi, Se).getTime();
    }

    // それ以外は Date() に委ねる（最後の保険）
    const d = new Date(s);
    if (!isNaN(d)) return d.getTime();

    return 0;
  }

  // ★ at を “表示用文字列” に正規化（Date型でも必ず yyyy-MM-dd HH:mm:ss に）
  function _toAtString_(v){
    if (Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v)) {
      return _fmtJst_(v);
    }
    const s = String(v || "").trim();
    // すでに yyyy-MM-dd HH:mm:ss っぽければそのまま
    if (/^\d{4}-\d{2}-\d{2}[ T]\d{2}:\d{2}/.test(s)) return s.replace("T"," ");
    // それ以外はパース→JST整形（可能な時）
    const t = _toTime_(v);
    if (t > 0) return _fmtJst_(new Date(t));
    return s;
  }

  const rows = values.slice(1)
    .filter(r => String(r[iUserId] || "").trim().toUpperCase() === userId)
    // ★ 数値時刻で降順にする（混在でも崩れない）
    .sort((a, b) => _toTime_(b[iAt]) - _toTime_(a[iAt]))
    .slice(0, limit)
    .map(r => ({
      txId:     String(r[iTxId] || ""),
      at:       _toAtString_(r[iAt]),
      type:     String(r[iType] || ""),
      userId:   String(r[iUserId] || ""),
      userName: String(r[iUserName] || ""),
      shopId:   String(r[iShopId] || ""),
      shopName: String(r[iShopName] || ""),
      amount:   Number(r[iAmount] || 0),
      status:   String(r[iStatus] || ""),
      note:     String(r[iNote] || ""),
      meta:     (iMeta >= 0) ? String(r[iMeta] || "") : "",
    }));

  return rows;
}


function _isTrue_(v) {
  return v === true || v === 1 || String(v).trim().toUpperCase() === "TRUE" || String(v).trim().toUpperCase() === "YES";
}

function _fmtJst_(d) {
  return Utilities.formatDate(d, "Asia/Tokyo", "yyyy-MM-dd HH:mm:ss");
}

function _getUsersSheet_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.USERS);
  if (!sh) throw new Error("Usersシートが見つかりません");
  return sh;
}

function _getShopsSheet_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.SHOPS);
  if (!sh) throw new Error("Shopsシートが見つかりません");
  return sh;
}

function _getTxSheet_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.TX);
  if (!sh) throw new Error("Txシートが見つかりません");
  return sh;
}

function _findUser_(userId) {
  const sh = _getUsersSheet_();
  const values = sh.getDataRange().getValues();
  if (values.length < 2) throw new Error("Usersにデータがありません");

  const m = _headerMap_(values[0]);
  const cUserId = m.idx("userid");
  const cName = m.idx("name");
  const cBal = m.idx("balance");
  const cActive = m.idx("isactive");
  const cPin = m.idx("pin");

  if (cUserId === -1 || cName === -1 || cBal === -1) throw new Error("Usersヘッダに userId/name/balance が必要です");
  if (cPin === -1) throw new Error("Usersヘッダに pin 列が必要です");

  for (let r = 1; r < values.length; r++) {
    if (String(values[r][cUserId]).trim().toUpperCase() === String(userId).trim().toUpperCase()) {
      const active = (cActive === -1) ? true : _isTrue_(values[r][cActive]);
      return {
        sh, row: r + 1,
        cBal: cBal + 1,
        userId: String(values[r][cUserId]).trim().toUpperCase(),
        name: String(values[r][cName] || "").trim(),
        balance: Number(values[r][cBal] || 0),
        active,
        pin: String(values[r][cPin] || "").trim(),
      };
    }
  }
  return null;
}

function _findShop_(shopId) {
  const sh = _getShopsSheet_();
  const values = sh.getDataRange().getValues();
  if (values.length < 2) throw new Error("Shopsにデータがありません");

  const m = _headerMap_(values[0]);
  const cShopId = m.idx("shopid");
  const cName = m.idx("shopname");
  const cBal = m.idx("balance");
  const cActive = m.idx("isactive");
  const cPass = m.idx("companypass");

  if (cShopId === -1 || cName === -1 || cBal === -1) {
    throw new Error("Shopsヘッダに shopId/shopName/balance が必要です");
  }
  if (cPass === -1) {
    throw new Error("Shopsヘッダに companyPass 列が必要です");
  }

  for (let r = 1; r < values.length; r++) {
    if (String(values[r][cShopId]).trim().toUpperCase() === String(shopId).trim().toUpperCase()) {
      const active = (cActive === -1) ? true : _isTrue_(values[r][cActive]);
      return {
        sh, row: r + 1,
        cBal: cBal + 1,
        shopId: String(values[r][cShopId]).trim().toUpperCase(),
        shopName: String(values[r][cName] || "").trim(),
        balance: Number(values[r][cBal] || 0),
        active,
        companyPass: String(values[r][cPass] || "").trim(),
      };
    }
  }
  return null;
}

function _normalizePin_(x){
  // 数値0や空も含めて「4桁文字列」に正規化
  const s = String(x ?? "").trim();
  if (s === "") return "";
  // すでに数字以外が混ざるなら弾く
  if (!/^\d+$/.test(s)) return s; // ここは後でエラーにしてもOK
  return s.padStart(4, "0");
}

function _assertPin_(userId, pin) {
  const u = _findUser_(userId);
  if (!u) throw new Error("ユーザーが見つかりません");
  if (!u.active) throw new Error("このユーザーは無効です");

  const saved = _normalizePin_(u.pin);
  const input = _normalizePin_(pin);

  if (!saved) throw new Error("PINが未設定です（先生に確認してください）");
  if (saved !== input) throw new Error("PINが違います");

  return u;
}


function _appendTx_(tx) {
  const sh = _getTxSheet_();
  sh.appendRow([
    tx.txId,
    tx.at,
    tx.type,
    tx.userId,
    tx.userName,
    tx.shopId,
    tx.shopName,
    tx.amount,
    tx.status,
    tx.note,
    tx.meta || ""
  ]);
}

/* ===== 支払い（PIN方式） ===== */
function api_pay(userId, pin, shopId, amount, note) {
  return lockRun_(() => {
    userId = String(userId || "").trim().toUpperCase();
    pin = String(pin || "").trim();
    shopId = String(shopId || "").trim().toUpperCase();
    amount = Number(amount);
    note = String(note || "").trim();

    if (!userId) throw new Error("支払い者IDが空です");
    if (!pin) throw new Error("PINを入力してください");
    if (!shopId) throw new Error("店IDが空です");
    if (!Number.isFinite(amount) || amount <= 0) throw new Error("金額が不正です");

    const payer = _assertPin_(userId, pin);
    const isGovernment = shopId === "GOV";
    const shop = isGovernment ? null : _findShop_(shopId);
    if (!isGovernment && !shop) throw new Error("店が見つかりません");
    if (shop && !shop.active) throw new Error("この店は無効です");
    if (isGovernment) {
      _getGovernmentAccount_();
      _sheetByNameOrThrow_(SHEETS.GOVERNMENT_LEDGER || "GovernmentLedger");
    }

    if (payer.balance < amount) throw new Error("残高が足りません");

    const newUserBal = payer.balance - amount;
    const newShopBal = shop ? shop.balance + amount : 0;

    payer.sh.getRange(payer.row, payer.cBal).setValue(newUserBal);
    if (shop) shop.sh.getRange(shop.row, shop.cBal).setValue(newShopBal);

    const txId = uuid_();
    const at = _fmtJst_(new Date());

    _appendTx_({
      txId, at,
      type: isGovernment ? "GOV_PAY" : "PAY",
      userId: payer.userId,
      userName: payer.name,
      shopId: isGovernment ? "GOV" : shop.shopId,
      shopName: isGovernment ? "政府" : shop.shopName,
      amount,
      status: "OK",
      note: isGovernment ? (note || "権利料") : note
    });

    const recipientBalance = isGovernment
      ? _phase2GovernmentMove_(amount, "RIGHTS_FEE_IN", note || "権利料", txId)
      : newShopBal;

    return {
      txId,
      at,
      user: { userId: payer.userId, name: payer.name, balance: newUserBal },
      shop: {
        shopId: isGovernment ? "GOV" : shop.shopId,
        shopName: isGovernment ? "政府" : shop.shopName,
        balance: recipientBalance,
        isGovernment
      }
    };
  });
}

/* ===== 管理：残高調整 ===== */
function api_adminAdjust(adminPass, userId, amount, reason) {
  _assertAdminPassValue_(adminPass);
  return lockRun_(() => {
    userId = String(userId || "").trim().toUpperCase();
    amount = Number(amount);
    reason = String(reason || "").trim();

    if (!userId) throw new Error("ユーザーIDが空です");
    if (!Number.isFinite(amount) || amount === 0) throw new Error("調整額が不正です（0は不可）");

    const u = _findUser_(userId);
    if (!u) throw new Error("ユーザーが見つかりません");
    if (!u.active) throw new Error("このユーザーは無効です");

    const newBal = u.balance + amount;
    u.sh.getRange(u.row, u.cBal).setValue(newBal);

    const txId = uuid_();
    const at = _fmtJst_(new Date());

    _appendTx_({
      txId, at,
      type: "ADJUST",
      userId: u.userId,
      userName: u.name,
      shopId: "",
      shopName: "",
      amount,
      status: "OK",
      note: reason
    });

    return { txId, at, user: { userId: u.userId, name: u.name, balance: newBal } };
  });
}

function api_adminListTx(limit) {
  limit = Math.min(Math.max(Number(limit || 50), 1), 300);

  const sh = _getTxSheet_();
  const values = sh.getDataRange().getValues();
  if (values.length <= 1) return [];

  const rows = values.slice(1).filter(r => r[0]);
  const tail = rows.slice(Math.max(0, rows.length - limit)).reverse();

  return tail.map(r => ({
    txId: String(r[0] || ""),
    at: String(r[1] || ""),
    type: String(r[2] || ""),
    userId: String(r[3] || ""),
    userName: String(r[4] || ""),
    shopId: String(r[5] || ""),
    shopName: String(r[6] || ""),
    amount: Number(r[7] || 0),
    status: String(r[8] || ""),
    note: String(r[9] || "")
  }));
}

function api_listUsers() {
  // 管理画面用（PINは返さない）
  const sh = _getUsersSheet_();
  const v = sh.getDataRange().getValues();
  if (v.length < 2) return [];

  const m = _headerMap_(v[0]);
  const cUserId = m.idx("userid");
  const cName = m.idx("name");
  const cBal = m.idx("balance");
  const cActive = m.idx("isactive");

  const out = [];
  for (let i = 1; i < v.length; i++) {
    const id = v[i][cUserId];
    if (!id) continue;
    const active = (cActive === -1) ? true : _isTrue_(v[i][cActive]);
    if (!active) continue;
    out.push({
      userId: String(id).trim().toUpperCase(),
      name: String(v[i][cName] || "").trim(),
      balance: Number(v[i][cBal] || 0)
    });
  }
  out.sort((a,b)=>a.userId.localeCompare(b.userId));
  return out;
}

function api_getShopName(shopId) {
  shopId = String(shopId || "").trim().toUpperCase();
  if (!shopId) return { ok: false, shopName: "" };
  const s = _findShop_(shopId);
  if (!s) return { ok: false, shopName: "" };
  return { ok: true, shopName: s.shopName };
}
function api_shopInfo(shopId){
  shopId = String(shopId || "").trim().toUpperCase();
  if (!shopId) throw new Error("店IDが空です");

  const s = _findShop_(shopId);
  if (!s) throw new Error("店が見つかりません");
  if (!s.active) throw new Error("この店は無効です");

  return {
    shopId: s.shopId,
    shopName: s.shopName,
    balance: s.balance
  };
}
function api_balance(userId, pin, limit){
  return lockRun_(() => {
    userId = String(userId || "").trim().toUpperCase();
    pin    = String(pin || "").trim();

    // ★追加：limit
    limit  = Math.min(Math.max(Number(limit || 30), 1), 300); // 1〜300

    if (!userId) throw new Error("ユーザーIDが空です");
    if (!pin) throw new Error("PINを入力してください");

    const u = _assertPin_(userId, pin);
    const membershipHistory = _membershipHistoryRows_(u.userId);

    return {
      userId: u.userId,
      name: u.name,
      balance: u.balance,
      companies: membershipHistory.filter(x=>x.isActive).map(x=>({shopId:x.shopId,shopName:x.shopName,role:x.role})),
      membershipHistory,
      tx: _getTxByUserId_(userId, limit), // ★ここをlimitに
    };
  });
}

function api_adminDistributeAll(adminPass, amount, reason){
  _assertAdminPassValue_(adminPass);
  // 全児童（Users.isActive=TRUE）に一括配付：残高加算 + Txに記録
  return lockRun_(() => {
    amount = Number(amount);
    reason = String(reason || "").trim();

    if (!Number.isFinite(amount) || amount <= 0) throw new Error("配付額は 1以上の数値にしてください");
    if (!reason) reason = "一括配付";

    const usersSh = _getUsersSheet_();
    const txSh = _getTxSheet_();

    const v = usersSh.getDataRange().getValues();
    if (v.length < 2) throw new Error("Usersにデータがありません");

    const m = _headerMap_(v[0]);
    const cUserId = m.idx("userid");
    const cName   = m.idx("name");
    const cBal    = m.idx("balance");
    const cActive = m.idx("isactive");

    if (cUserId === -1 || cName === -1 || cBal === -1) {
      throw new Error("Usersヘッダに userId / name / balance が必要です");
    }

    // 対象行（isActive が無い場合は全員）
    const targets = [];
    for (let i = 1; i < v.length; i++){
      const userId = String(v[i][cUserId] || "").trim();
      if (!userId) continue;

      const active = (cActive === -1) ? true : _isTrue_(v[i][cActive]);
      if (!active) continue;

      const name = String(v[i][cName] || "").trim();
      const bal  = Number(v[i][cBal] || 0);

      targets.push({
        row: i + 1,
        userId: userId.toUpperCase(),
        name,
        oldBal: bal,
        newBal: bal + amount,
      });
    }

    if (targets.length === 0) throw new Error("配付対象（isActive=TRUE）が見つかりません");

    // 残高を一括更新（balance列だけ）
    const newBalValues = targets.map(t => [t.newBal]);
    usersSh.getRange(2, cBal + 1, v.length - 1, 1).setValues(
      // 行数が合わないので「全行」を作って上書きする方式にする
      (function(){
        const out = [];
        let ptr = 0;
        for (let i=1;i<v.length;i++){
          const userId = String(v[i][cUserId] || "").trim();
          if (!userId) { out.push([v[i][cBal] || 0]); continue; }

          const active = (cActive === -1) ? true : _isTrue_(v[i][cActive]);
          if (!active) { out.push([v[i][cBal] || 0]); continue; }

          out.push([targets[ptr].newBal]);
          ptr++;
        }
        return out;
      })()
    );

    // Txを一括でappend（1人1行）
    const at = _fmtJst_(new Date());
    const rows = targets.map(t => ([
      uuid_(),
      at,
      "DISTRIB",           // 種別（見分けやすいように）
      t.userId,
      t.name,
      "",                  // shopId
      "",                  // shopName
      amount,
      "OK",
      reason
    ]));
    txSh.getRange(txSh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);

    return {
      count: targets.length,
      amountEach: amount,
      total: targets.length * amount,
      at,
      reason
    };
  });
}
function api_adminCollectTaxAll(adminPass, taxAmount, reason, mode){
  _assertAdminPassValue_(adminPass);
  // mode: "CAP"（足りない子はある分だけ回収） / "SKIP"（足りない子は回収しない）
  return lockRun_(() => {
    taxAmount = Number(taxAmount);
    reason = String(reason || "").trim();
    mode = String(mode || "CAP").trim().toUpperCase(); // default CAP

    if (!Number.isFinite(taxAmount) || taxAmount <= 0) throw new Error("税額は 1以上の数値にしてください");
    if (!reason) reason = "税金回収";
    if (!["CAP","SKIP"].includes(mode)) throw new Error("mode は CAP / SKIP のどちらかです");

    const usersSh = _getUsersSheet_();
    const txSh = _getTxSheet_();

    const v = usersSh.getDataRange().getValues();
    if (v.length < 2) throw new Error("Usersにデータがありません");

    const m = _headerMap_(v[0]);
    const cUserId = m.idx("userid");
    const cName   = m.idx("name");
    const cBal    = m.idx("balance");
    const cActive = m.idx("isactive");

    if (cUserId === -1 || cName === -1 || cBal === -1) {
      throw new Error("Usersヘッダに userId / name / balance が必要です");
    }

    const updates = [];   // 残高更新対象
    const rowsTx = [];    // Tx行

    let countApplied = 0;
    let countSkipped = 0;
    let totalCollected = 0;

    const at = _fmtJst_(new Date());

    // まず全行分の新balance配列を作る（setValues一括）
    const newBalCol = [];
    for (let i=1;i<v.length;i++){
      const userIdRaw = String(v[i][cUserId] || "").trim();
      const oldBal = Number(v[i][cBal] || 0);

      if (!userIdRaw) { newBalCol.push([oldBal]); continue; }

      const active = (cActive === -1) ? true : _isTrue_(v[i][cActive]);
      if (!active) { newBalCol.push([oldBal]); continue; }

      const userId = userIdRaw.toUpperCase();
      const name = String(v[i][cName] || "").trim();

      let take = taxAmount;

      if (mode === "SKIP" && oldBal < taxAmount){
        // 不足は回収しない
        countSkipped++;
        newBalCol.push([oldBal]);
        continue;
      }

      // CAP：不足ならある分だけ回収（0まで）
      if (oldBal < taxAmount) take = Math.max(0, oldBal);

      const newBal = oldBal - take;
      newBalCol.push([newBal]);

      // take が0なら実質回収なし（残高0の子など）
      if (take > 0){
        countApplied++;
        totalCollected += take;

        rowsTx.push([
          uuid_(),
          at,
          "TAX",
          userId,
          name,
          "",      // shopId
          "",      // shopName
          -take,   // ★マイナスで記録（回収）
          "OK",
          `${reason}${mode==="CAP" && oldBal < taxAmount ? "（不足分は0まで回収）" : ""}`
        ]);
      } else {
        // take=0 の場合はTxを書かない（必要なら書く仕様にもできる）
      }
    }

    if (countApplied === 0 && totalCollected === 0){
      // SKIPで全員不足/0などの時
      return { countApplied: 0, countSkipped, taxAmount, totalCollected: 0, at, reason, mode };
    }

    // 残高一括更新
    usersSh.getRange(2, cBal+1, newBalCol.length, 1).setValues(newBalCol);

    // Tx一括追記
    if (rowsTx.length > 0){
      txSh.getRange(txSh.getLastRow()+1, 1, rowsTx.length, rowsTx[0].length).setValues(rowsTx);
    }

    return {
      countApplied,
      countSkipped,
      taxAmount,
      totalCollected,
      at,
      reason,
      mode
    };
  });
}
function api_checkAdminPass(pass){
  pass = String(pass || "").trim();
  if (!pass) throw new Error("パスワードを入力してください");

  const correct = String(getConfig_("ADMIN_PASS", "")).trim();
  if (!correct) throw new Error("Configに ADMIN_PASS が設定されていません");

  if (pass !== correct) throw new Error("パスワードが違います");
  return { ok: true };
}

function api_listShops() {
  const sh = _getShopsSheet_();
  const v = sh.getDataRange().getValues();
  if (v.length < 2) return [];

  const m = _headerMap_(v[0]);
  const cShopId = m.idx("shopid");
  const cName   = m.idx("shopname");
  const cBal    = m.idx("balance");
  const cActive = m.idx("isactive");

  if (cShopId === -1 || cName === -1 || cBal === -1) {
    throw new Error("Shopsヘッダに shopId / shopName / balance が必要です");
  }

  const out = [];
  for (let i = 1; i < v.length; i++) {
    const id = v[i][cShopId];
    if (!id) continue;

    const active = (cActive === -1) ? true : _isTrue_(v[i][cActive]);
    if (!active) continue;

    out.push({
      shopId: String(id).trim().toUpperCase(),
      shopName: String(v[i][cName] || "").trim(),
      balance: Number(v[i][cBal] || 0),
    });
  }

  out.sort((a,b)=>a.shopId.localeCompare(b.shopId));
  return out;
}
function api_adminAdjustShop(adminPass, shopId, amount, reason) {
  _assertAdminPassValue_(adminPass);
  return lockRun_(() => {
    shopId = String(shopId || "").trim().toUpperCase();
    amount = Number(amount);
    reason = String(reason || "").trim();

    if (!shopId) throw new Error("会社IDが空です");
    if (!Number.isFinite(amount) || amount === 0) throw new Error("調整額が不正です（0は不可）");

    const s = _findShop_(shopId);
    if (!s) throw new Error("会社が見つかりません");
    if (!s.active) throw new Error("この会社は無効です");

    const newBal = s.balance + amount;
    if (newBal < 0) throw new Error("会社残高がマイナスになります（引き出し過ぎ）");

    s.sh.getRange(s.row, s.cBal).setValue(newBal);

    const txId = uuid_();
    const at = _fmtJst_(new Date());

    _appendTx_({
      txId, at,
      type: "SHOP_ADJ",
      userId: "",
      userName: "",
      shopId: s.shopId,
      shopName: s.shopName,
      amount,
      status: "OK",
      note: reason || "会社調整"
    });

    return { txId, at, shop: { shopId: s.shopId, shopName: s.shopName, balance: newBal } };
  });
}
function api_adminDistributeAllShops(adminPass, amount, reason){
  _assertAdminPassValue_(adminPass);
  return lockRun_(() => {
    amount = Number(amount);
    reason = String(reason || "").trim();
    if (!Number.isFinite(amount) || amount <= 0) throw new Error("配付額は 1以上の数値にしてください");
    if (!reason) reason = "会社一括配付";

    const shopsSh = _getShopsSheet_();
    const txSh = _getTxSheet_();

    const v = shopsSh.getDataRange().getValues();
    if (v.length < 2) throw new Error("Shopsにデータがありません");

    const m = _headerMap_(v[0]);
    const cShopId = m.idx("shopid");
    const cName   = m.idx("shopname");
    const cBal    = m.idx("balance");
    const cActive = m.idx("isactive");

    if (cShopId === -1 || cName === -1 || cBal === -1) {
      throw new Error("Shopsヘッダに shopId / shopName / balance が必要です");
    }

    const targets = [];
    for (let i=1;i<v.length;i++){
      const shopIdRaw = String(v[i][cShopId] || "").trim();
      if (!shopIdRaw) continue;

      const active = (cActive === -1) ? true : _isTrue_(v[i][cActive]);
      if (!active) continue;

      const shopId = shopIdRaw.toUpperCase();
      const name = String(v[i][cName] || "").trim();
      const bal  = Number(v[i][cBal] || 0);
      targets.push({ idx:i, shopId, name, oldBal:bal, newBal: bal + amount });
    }
    if (targets.length === 0) throw new Error("配付対象（isActive=TRUE）が見つかりません");

    // balance列を全行一括更新
    const newBalCol = [];
    let ptr=0;
    for (let i=1;i<v.length;i++){
      const shopIdRaw = String(v[i][cShopId] || "").trim();
      const oldBal = Number(v[i][cBal] || 0);

      if (!shopIdRaw) { newBalCol.push([oldBal]); continue; }
      const active = (cActive === -1) ? true : _isTrue_(v[i][cActive]);
      if (!active) { newBalCol.push([oldBal]); continue; }

      newBalCol.push([targets[ptr].newBal]);
      ptr++;
    }
    shopsSh.getRange(2, cBal+1, newBalCol.length, 1).setValues(newBalCol);

    // Tx追記
    const at = _fmtJst_(new Date());
    const rows = targets.map(t => ([
      uuid_(),
      at,
      "SHOP_DIST",
      "", "",              // user
      t.shopId, t.name,    // shop
      amount,
      "OK",
      reason
    ]));
    txSh.getRange(txSh.getLastRow()+1, 1, rows.length, rows[0].length).setValues(rows);

    return { count: targets.length, amountEach: amount, total: targets.length * amount, at, reason };
  });
}
function api_adminWithdrawAllShops(adminPass, withdrawAmount, reason, mode){
  _assertAdminPassValue_(adminPass);
  return lockRun_(() => {
    withdrawAmount = Number(withdrawAmount);
    reason = String(reason || "").trim();
    mode = String(mode || "CAP").trim().toUpperCase();

    if (!Number.isFinite(withdrawAmount) || withdrawAmount <= 0) throw new Error("引き出し額は 1以上の数値にしてください");
    if (!reason) reason = "会社一括引き出し";
    if (!["CAP","SKIP"].includes(mode)) throw new Error("mode は CAP / SKIP");

    const shopsSh = _getShopsSheet_();
    const txSh = _getTxSheet_();

    const v = shopsSh.getDataRange().getValues();
    if (v.length < 2) throw new Error("Shopsにデータがありません");

    const m = _headerMap_(v[0]);
    const cShopId = m.idx("shopid");
    const cName   = m.idx("shopname");
    const cBal    = m.idx("balance");
    const cActive = m.idx("isactive");

    if (cShopId === -1 || cName === -1 || cBal === -1) {
      throw new Error("Shopsヘッダに shopId / shopName / balance が必要です");
    }

    const at = _fmtJst_(new Date());

    let countApplied = 0;
    let countSkipped = 0;
    let totalWithdrawn = 0;

    const newBalCol = [];
    const rowsTx = [];

    for (let i=1;i<v.length;i++){
      const shopIdRaw = String(v[i][cShopId] || "").trim();
      const oldBal = Number(v[i][cBal] || 0);

      if (!shopIdRaw) { newBalCol.push([oldBal]); continue; }

      const active = (cActive === -1) ? true : _isTrue_(v[i][cActive]);
      if (!active) { newBalCol.push([oldBal]); continue; }

      const shopId = shopIdRaw.toUpperCase();
      const name = String(v[i][cName] || "").trim();

      if (mode === "SKIP" && oldBal < withdrawAmount){
        countSkipped++;
        newBalCol.push([oldBal]);
        continue;
      }

      let take = withdrawAmount;
      if (oldBal < withdrawAmount) take = Math.max(0, oldBal); // CAP

      const newBal = oldBal - take;
      newBalCol.push([newBal]);

      if (take > 0){
        countApplied++;
        totalWithdrawn += take;
        rowsTx.push([
          uuid_(),
          at,
          "SHOP_WD",
          "", "",
          shopId, name,
          -take, // ★マイナスで記録
          "OK",
          `${reason}${mode==="CAP" && oldBal < withdrawAmount ? "（不足分は0まで）" : ""}`
        ]);
      }
    }

    shopsSh.getRange(2, cBal+1, newBalCol.length, 1).setValues(newBalCol);

    if (rowsTx.length > 0){
      txSh.getRange(txSh.getLastRow()+1, 1, rowsTx.length, rowsTx[0].length).setValues(rowsTx);
    }

    return { countApplied, countSkipped, withdrawAmount, totalWithdrawn, at, reason, mode };
  });
}

/**
 * 毎週月曜日の時間主導型トリガー用（利率5%）。
 * 関数名を維持しているため、既存トリガーの作り直しは不要です。
 */
function api_adminInterestOnBalance_Trigger() {
  _ensureCoreSheets_();
  return api_phase2UserInterest(
    String(getConfig_("ADMIN_PASS", "")),
    5,
    "週次利息（月曜日）"
  );
}

function api_adminInterestOnBalance(adminPass, rate, reason, settleGovernment) {
  _assertAdminPassValue_(adminPass);
  // balance に対して利子を付けて balance に加算（=複利）
  return lockRun_(() => {
    // 引数が渡されない（トリガー実行）場合は 0.05 をデフォルトに
    rate = Number(rate || 0.05);
    reason = String(reason || "利子（残高）").trim();

    if (!Number.isFinite(rate) || rate < 0) throw new Error("利率が不正です");

    const usersSh = _getUsersSheet_();
    const txSh = _getTxSheet_();

    const v = usersSh.getDataRange().getValues();
    if (v.length < 2) throw new Error("Usersにデータがありません");

    const m = _headerMap_(v[0]);
    const cUserId = m.idx("userid");
    const cName   = m.idx("name");
    const cBal    = m.idx("balance");
    const cActive = m.idx("isactive");

    if (cUserId === -1 || cName === -1 || cBal === -1) {
      throw new Error("Usersヘッダに userId / name / balance が必要です");
    }

    // 端数処理設定
    const roundMode = String(getConfig_("INTEREST_ROUND", "FLOOR")).trim().toUpperCase();
    const roundFn = (x) => {
      if (roundMode === "ROUND") return Math.round(x);
      if (roundMode === "CEIL")  return Math.ceil(x);
      return Math.floor(x);
    };

    const at = _fmtJst_(new Date());
    const newBalCol = [];
    const rowsTx = [];
    let countApplied = 0;
    let totalInterest = 0;

    for (let i = 1; i < v.length; i++) {
      const userIdRaw = String(v[i][cUserId] || "").trim();
      const oldBal = Number(v[i][cBal] || 0);

      // IDがない、または非アクティブなユーザーはスキップ
      const active = (cActive === -1) ? true : _isTrue_(v[i][cActive]);
      if (!userIdRaw || !active) {
        newBalCol.push([oldBal]);
        continue;
      }

      const base = Math.max(0, oldBal);
      const interest = roundFn(base * rate);
      const newBal = oldBal + (Number.isFinite(interest) ? interest : 0);
      newBalCol.push([newBal]);

      if (interest > 0) {
        countApplied++;
        totalInterest += interest;

        rowsTx.push([
          uuid_(),
          at,
          "INTEREST_B",
          userIdRaw.toUpperCase(),
          String(v[i][cName] || "").trim(),
          "", "",
          interest,
          "OK",
          `${reason} ${(rate * 100).toFixed(2)}%（balance:${oldBal}→${newBal}）`
        ]);
      }
    }

    // 一括更新
    usersSh.getRange(2, cBal + 1, newBalCol.length, 1).setValues(newBalCol);
    if (rowsTx.length > 0) {
      txSh.getRange(txSh.getLastRow() + 1, 1, rowsTx.length, rowsTx[0].length).setValues(rowsTx);
    }

    let governmentBalance = null;
    if (settleGovernment && totalInterest > 0) {
      governmentBalance = _phase2GovernmentMove_(
        -totalInterest,
        "USER_INTEREST_OUT",
        reason,
        ""
      );
    }

    return { at, rate, roundMode, countApplied, totalInterest, governmentBalance };
  });
}
function api_adminInterestOnShopBalance(adminPass, rate, reason){
  _assertAdminPassValue_(adminPass);
  // 会社（Shops）の balance に対して利子を付けて balance に加算（=複利）
  // rate: 0.05 = 5%
  return lockRun_(() => {
    rate = Number(rate);
    reason = String(reason || "").trim() || "利子（会社残高）";

    if (!Number.isFinite(rate) || rate < 0) throw new Error("利率が不正です（例：0.05）");

    const shopsSh = _getShopsSheet_();
    const txSh = _getTxSheet_();

    const v = shopsSh.getDataRange().getValues();
    if (v.length < 2) throw new Error("Shopsにデータがありません");

    const m = _headerMap_(v[0]);
    const cShopId = m.idx("shopid");
    const cName   = m.idx("shopname");
    const cBal    = m.idx("balance");
    const cActive = m.idx("isactive");

    if (cShopId === -1 || cName === -1 || cBal === -1) {
      throw new Error("Shopsヘッダに shopId / shopName / balance が必要です");
    }

    // 端数処理（Users利子と同じConfigを流用）
    // FLOOR=切り捨て / ROUND=四捨五入 / CEIL=切り上げ
    const roundMode = String(getConfig_("INTEREST_ROUND", "FLOOR")).trim().toUpperCase();
    const roundFn = (x) => {
      if (roundMode === "ROUND") return Math.round(x);
      if (roundMode === "CEIL")  return Math.ceil(x);
      return Math.floor(x); // default
    };

    const at = _fmtJst_(new Date());

    const newBalCol = [];
    const rowsTx = [];

    let countApplied = 0;
    let totalInterest = 0;

    for (let i=1;i<v.length;i++){
      const shopIdRaw = String(v[i][cShopId] || "").trim();
      const oldBal = Number(v[i][cBal] || 0);

      if (!shopIdRaw) { newBalCol.push([oldBal]); continue; }

      const active = (cActive === -1) ? true : _isTrue_(v[i][cActive]);
      if (!active) { newBalCol.push([oldBal]); continue; }

      const base = Math.max(0, oldBal);              // マイナス残高は0扱い
      const interest = roundFn(base * rate);         // 利子

      const newBal = oldBal + (Number.isFinite(interest) ? interest : 0);
      newBalCol.push([newBal]);

      if (interest > 0){
        countApplied++;
        totalInterest += interest;

        const shopId = shopIdRaw.toUpperCase();
        const name = String(v[i][cName] || "").trim();

        rowsTx.push([
          uuid_(),
          at,
          "INTEREST_S",       // ★会社残高利子
          "",                 // userId
          "",                 // userName
          shopId,
          name,
          interest,           // プラス
          "OK",
          `${reason} ${(rate*100).toFixed(2)}%（balance:${oldBal}→${newBal}）`
        ]);
      }
    }

    // balance 一括更新
    shopsSh.getRange(2, cBal+1, newBalCol.length, 1).setValues(newBalCol);

    // Tx 一括追記
    if (rowsTx.length > 0){
      txSh.getRange(txSh.getLastRow()+1, 1, rowsTx.length, rowsTx[0].length).setValues(rowsTx);
    }

    return { at, rate, roundMode, countApplied, totalInterest };
  });
}

function api_companyMembers(shopId){
  shopId = String(shopId || "").trim().toUpperCase();
  if (!shopId) throw new Error("会社IDが空です");

  const shop = _findShop_(shopId);
  if (!shop) throw new Error("会社が見つかりません");
  if (!shop.active) throw new Error("この会社は無効です");

  const sh = _getCompanyMembersSheet_();
  const v = sh.getDataRange().getValues();
  if (v.length < 2) return [];

  const m = _headerMap_(v[0]);
  const cShop = m.idx("shopid");
  const cUser = m.idx("userid");
  const cRole = m.idx("role");
  const cActive = m.idx("isactive");

  return v.slice(1)
    .filter(r => String(r[cShop] || "").trim().toUpperCase() === shopId)
    .filter(r => cActive < 0 || _isTrue_(r[cActive]))
    .map(r => {
      const userId = String(r[cUser] || "").trim().toUpperCase();
      const u = _findUser_(userId);
      return u ? {
        userId: u.userId,
        name: u.name,
        role: String(r[cRole] || "MEMBER").trim().toUpperCase()
      } : null;
    })
    .filter(Boolean)
    .sort((a,b) => {
      if (a.role === "PRESIDENT" && b.role !== "PRESIDENT") return -1;
      if (b.role === "PRESIDENT" && a.role !== "PRESIDENT") return 1;
      return a.userId.localeCompare(b.userId);
    });
}

// 3) 給与一括送金（最大5人・別額・会社PASS一致時のみ）
function api_salaryPayBatch(payload){
  return lockRun_(() => {
    payload = payload || {};
    const shopId = String(payload.shopId || "").trim().toUpperCase();
    const companyPass = String(payload.companyPass || "").trim();
    const items = Array.isArray(payload.items) ? payload.items : [];

    if (!shopId) throw new Error("会社IDが空です");
    if (!companyPass) throw new Error("会社PASSを入力してください");
    if (items.length === 0) throw new Error("送金する行がありません");
    if (items.length > 5) throw new Error("一度に送金できるのは最大5人までです");

    const shop = _findShop_(shopId);
    if (!shop) throw new Error("会社が見つかりません");
    if (!shop.active) throw new Error("この会社は無効です");

    if (!shop.companyPass) throw new Error("この会社はPASS未設定です（先生に確認してください）");
    if (companyPass !== shop.companyPass) throw new Error("会社PASSが違います");

    // 正規化 + 重複チェック
    const norm = items.map((x, i) => {
      const to = String(x.to || "").trim().toUpperCase();
      const amount = Number(x.amount);
      if (!to) throw new Error(`支払い先が不正です（${i+1}件目）`);
      if (!Number.isFinite(amount) || amount <= 0) throw new Error(`金額が不正です（${i+1}件目）`);
      if (!Number.isInteger(amount)) throw new Error(`金額は整数で入力してください（${i+1}件目）`);
      return { to, amount };
    });

    const seen = new Set();
    for (const it of norm){
      if (seen.has(it.to)) throw new Error(`同じ支払い先が重複しています：${it.to}`);
      seen.add(it.to);
    }

    const total = norm.reduce((s,x)=>s + x.amount, 0);
    if (shop.balance < total) throw new Error("会社残高が足りません");

    // 受取側ユーザー取得
    const receivers = norm.map(it => {
      const u = _findUser_(it.to);
      if (!u) throw new Error(`ユーザーが見つかりません：${it.to}`);
      if (!u.active) throw new Error(`このユーザーは無効です：${it.to}`);
      return u;
    });

    // 会社残高更新
    const newShopBal = shop.balance - total;
    if (newShopBal < 0) throw new Error("会社残高がマイナスになります");
    shop.sh.getRange(shop.row, shop.cBal).setValue(newShopBal);

    // ユーザー残高更新 + Tx一括追記
    const txSh = _getTxSheet_();
    const at = _fmtJst_(new Date());

    const txRows = [];
    const results = [];

    for (let i=0;i<norm.length;i++){
      const it = norm[i];
      const u = receivers[i];

      const newUserBal = u.balance + it.amount;
      u.sh.getRange(u.row, u.cBal).setValue(newUserBal);

      txRows.push([
        uuid_(),
        at,
        "SALARY",
        u.userId,
        u.name,
        shop.shopId,
        shop.shopName,
        it.amount,
        "OK",
        "給与"
      ]);

      results.push({ to: u.userId, amount: it.amount, ok: true });
    }

    if (txRows.length > 0){
      txSh.getRange(txSh.getLastRow() + 1, 1, txRows.length, txRows[0].length).setValues(txRows);
    }

    return {
      ok: true,
      at,
      shop: { shopId: shop.shopId, shopName: shop.shopName, balance: newShopBal },
      results
    };
  });
}

/*
【シート要件】
- Shops シートに companyPass 列（必須）
  例: shopId | shopName | balance | isActive | companyPass
- Users は紐づけ不要（shopId列は不要）
*/

function api_adminChangeShopPass(adminPass, shopId, newPass){
  _assertAdminPassValue_(adminPass);
  if (String(newPass).length < 4){
    throw new Error("PASSは4文字以上にしてください");
  }

  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.SHOPS);
  const rows = sh.getDataRange().getValues();
  const header = rows.shift();

  const iId   = header.indexOf("shopId");
  const iPass = header.indexOf("companyPass"); // ★ここが重要

  if (iPass === -1) throw new Error("companyPass 列が見つかりません");

  for (let i = 0; i < rows.length; i++){
    if (String(rows[i][iId]).toUpperCase() === shopId){
      sh.getRange(i + 2, iPass + 1).setValue(newPass);
      return { ok:true };
    }
  }
  throw new Error("会社が見つかりません");
}
function api_adminChangeUserPin(adminPass, userId, newPin){
  _assertAdminPassValue_(adminPass);


  if (!/^\d{4}$/.test(newPin)){
    throw new Error("PINは4桁の数字にしてください");
  }

  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.USERS);
  const rows = sh.getDataRange().getValues();
  const header = rows.shift();

  const iId  = header.indexOf("userId");
  const iPin = header.indexOf("pin");

  for (let r of rows){
    if (String(r[iId]).toUpperCase() === userId){
      sh.getRange(rows.indexOf(r)+2, iPin+1).setValue(newPin);
      return { ok:true };
    }
  }
  throw new Error("児童が見つかりません");
}

// ==============================
// 株API 完全版（ClassPay）
// 依存：SHEETS / _getShopsSheet_ / _findShop_ / _findUser_ / _assertPin_
//      _isTrue_ / _fmtJst_ / uuid_ / lockRun_ / _appendTx_ / api_pay
// ==============================

// ====== Holdings（保有） ======
function _getHoldingsSheet_(){
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.HOLDINGS);
  if (!sh) throw new Error("Holdingsシートが見つかりません");
  return sh;
}

function _ensureHoldingsHeader_(){
  const sh = _getHoldingsSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow === 0){
    sh.appendRow(["userId","shopId","shares","updatedAt"]);
    return;
  }
  // 1行目がヘッダじゃないケースの軽い保険（必要なら厳密化）
  const header = sh.getRange(1,1,1,Math.max(1, sh.getLastColumn())).getValues()[0];
  const m = _headerMap_(header);
  if (m.idx("userid") < 0 || m.idx("shopid") < 0 || m.idx("shares") < 0){
    // 既存データがある場合は勝手に触らない方が安全なのでエラー
    throw new Error("Holdingsヘッダが不正です（userId, shopId, shares が必要）");
  }
}

function _getHolding_(userId, shopId){
  userId = String(userId||"").trim().toUpperCase();
  shopId = String(shopId||"").trim().toUpperCase();

  _ensureHoldingsHeader_();

  const sh = _getHoldingsSheet_();
  const v = sh.getDataRange().getValues();
  if (v.length < 2) return { row: null, shares: 0 };

  const m = _headerMap_(v[0]);
  const cUserId = m.idx("userid");
  const cShopId = m.idx("shopid");
  const cShares = m.idx("shares");

  for (let i=1;i<v.length;i++){
    if (String(v[i][cUserId]||"").trim().toUpperCase() === userId &&
        String(v[i][cShopId]||"").trim().toUpperCase() === shopId){
      return { row: i+1, shares: Number(v[i][cShares]||0) };
    }
  }
  return { row: null, shares: 0 };
}

function _setHolding_(userId, shopId, shares){
  userId = String(userId||"").trim().toUpperCase();
  shopId = String(shopId||"").trim().toUpperCase();
  shares = Math.max(0, Math.floor(Number(shares||0)));

  _ensureHoldingsHeader_();

  const sh = _getHoldingsSheet_();
  const v = sh.getDataRange().getValues();
  const m = _headerMap_(v[0]);
  const cShares = m.idx("shares");
  const cAt     = m.idx("updatedat");

  const cur = _getHolding_(userId, shopId);
  const now = _fmtJst_(new Date());

  if (cur.row){
    sh.getRange(cur.row, cShares+1).setValue(shares);
    if (cAt >= 0) sh.getRange(cur.row, cAt+1).setValue(now);
  } else {
    sh.appendRow([userId, shopId, shares, now]);
  }
}


// ====== 価格計算（スプレッド + impact込み） ======
function _clamp_(x, mn, mx){
  return Math.max(mn, Math.min(mx, x));
}

function _applySpread_(base, spread, side){
  spread = Number(spread||0);
  const b = Number(base||0);
  if (String(side).toUpperCase()==="BUY")  return Math.ceil(b * (1 + spread));
  return Math.floor(b * (1 - spread)); // SELL
}

// 通常の基準価格（表示用）
function _calcBasePrice_(assets, totalShares, basePrice, k, priceMin, priceMax){
  const raw = Number(basePrice||10) + Number(k||0.1) * (Number(assets||0) / Math.max(1, Number(totalShares||1)));
  const p = Math.round(raw);
  return _clamp_(p, Number(priceMin||1), Number(priceMax||9999));
}

// 「注文量による影響」を織り込んだ基準価格（スプレッド前）
// side: BUY / SELL
function _calcImpactBasePrice_(assets, totalShares, basePrice, k, priceMin, priceMax, side, shares){
  assets = Number(assets||0);
  totalShares = Math.max(1, Number(totalShares||1));
  basePrice = Number(basePrice||10);
  k = Number(k||0.1);
  shares = Math.max(1, Math.floor(Number(shares||1)));

  const a = basePrice + k * (assets / totalShares);
  const ratio = (k * shares) / totalShares;

  let denom = (String(side).toUpperCase()==="BUY") ? (1 - ratio) : (1 + ratio);
  if (denom <= 0.05) denom = 0.05; // 異常値ガード

  const p0 = a / denom;
  const p = Math.round(p0);
  return _clamp_(p, Number(priceMin||1), Number(priceMax||9999));
}

// Shopsから株パラメータ（任意列はデフォルト）
function _getStockParamsFromShops_(shopId){
  shopId = String(shopId||"").trim().toUpperCase();

  const sh = _getShopsSheet_();
  const values = sh.getDataRange().getValues();
  if (values.length < 2) throw new Error("Shopsにデータがありません");

  const { idx } = _headerMap_(values[0]);

  const cShopId = idx("shopid");
  const cBal    = idx("balance");
  const cActive = idx("isactive");

  const cTotal  = idx("totalshares");
  const cBase   = idx("baseprice");
  const cK      = idx("k");
  const cMin    = idx("pricemin");
  const cMax    = idx("pricemax");
  const cStockA = idx("stockactive");
  const cSpr    = idx("spread");

  for (let r=1; r<values.length; r++){
    if (String(values[r][cShopId]||"").trim().toUpperCase() !== shopId) continue;

    const assets      = Number(values[r][cBal] || 0);
    const totalShares = (cTotal >= 0) ? Number(values[r][cTotal] || 1000) : 1000;
    const basePrice   = (cBase  >= 0) ? Number(values[r][cBase]  || 10)   : 10;
    const k           = (cK     >= 0) ? Number(values[r][cK]     || 0.1)  : 0.1;
    const priceMin    = (cMin   >= 0) ? Number(values[r][cMin]   || 1)    : 1;
    const priceMax    = (cMax   >= 0) ? Number(values[r][cMax]   || 9999) : 9999;

    const spread      = (cSpr >= 0 && String(values[r][cSpr]||"").trim()!=="")
      ? Number(values[r][cSpr]||0)
      : Number(getConfig_("STOCK_SPREAD", 0.10)); // デフォ10%

    const isActive = (cStockA >= 0 && String(values[r][cStockA] || "").trim() !== "")
      ? _isTrue_(values[r][cStockA])
      : ((cActive === -1) ? true : _isTrue_(values[r][cActive]));

    return { shopId, assets, totalShares, basePrice, k, priceMin, priceMax, spread, isActive };
  }

  throw new Error("会社が見つかりません");
}

// 表示用の「現在株価（基準＋買値/売値）」を返す
function _getQuote_(shopId){
  shopId = String(shopId||"").trim().toUpperCase();

  const p = _getStockParamsFromShops_(shopId);
  const s = _findShop_(shopId);

  const base = _calcBasePrice_(p.assets, p.totalShares, p.basePrice, p.k, p.priceMin, p.priceMax);

  return {
    shopId,
    shopName: s ? s.shopName : "",
    isActive: p.isActive,
    assets: p.assets,
    totalShares: p.totalShares,
    basePrice: base,
    buyPrice:  _applySpread_(base, p.spread, "BUY"),
    sellPrice: _applySpread_(base, p.spread, "SELL"),
    spread: p.spread
  };
}


// ====== 会社→児童 の払い戻し（売却用） ======
function _shopToUserPay_(shopId, userId, amount, note){
  shopId = String(shopId||"").trim().toUpperCase();
  userId = String(userId||"").trim().toUpperCase();
  amount = Number(amount);
  note = String(note||"").trim();

  if (!Number.isFinite(amount) || amount <= 0) throw new Error("金額が不正です");

  const u = _findUser_(userId);
  if (!u) throw new Error("ユーザーが見つかりません");
  if (!u.active) throw new Error("このユーザーは無効です");

  const s = _findShop_(shopId);
  if (!s) throw new Error("会社が見つかりません");
  if (!s.active) throw new Error("この会社は無効です");

  if (s.balance < amount) throw new Error("会社残高が足りません（買い戻しできません）");

  const newShopBal = s.balance - amount;
  const newUserBal = u.balance + amount;

  s.sh.getRange(s.row, s.cBal).setValue(newShopBal);
  u.sh.getRange(u.row, u.cBal).setValue(newUserBal);

  const txId = uuid_();
  const at = _fmtJst_(new Date());

  _appendTx_({
    txId, at,
    type: "STOCK_SELL",
    userId: u.userId,
    userName: u.name,
    shopId: s.shopId,
    shopName: s.shopName,
    amount: amount,
    status: "OK",
    note: note
  });

  return {
    txId, at,
    user: { userId: u.userId, name: u.name, balance: newUserBal },
    shop: { shopId: s.shopId, shopName: s.shopName, balance: newShopBal }
  };
}


// ==============================
// 公開API
// ==============================

// 株価（表示用）
function api_stockQuote(shopId){
  const q = _getQuote_(shopId);
  return { ok:true, quote:q };
}

// 保有確認（PIN必要）
function api_stockHolding(userId, pin, shopId){
  return lockRun_(() => {
    userId = String(userId||"").trim().toUpperCase();
    pin    = String(pin||"").trim();
    shopId = String(shopId||"").trim().toUpperCase();

    _assertPin_(userId, pin);

    const q = _getQuote_(shopId);
    const h = _getHolding_(userId, shopId);

    return { ok:true, quote:q, holding:{ userId, shopId, shares:h.shares } };
  });
}

// 株を買う（提示の買値で約定：大量でも単価固定）
function api_stockBuy(userId, pin, shopId, shares, expectedBuyPrice){
  return lockRun_(() => {
    userId = String(userId||"").trim().toUpperCase();
    pin    = String(pin||"").trim();
    shopId = String(shopId||"").trim().toUpperCase();
    shares = Math.floor(Number(shares||0));

    if (!userId) throw new Error("ユーザーIDが空です");
    if (!pin) throw new Error("PINを入力してください");
    if (!shopId) throw new Error("会社IDが空です");
    if (!(shares > 0)) throw new Error("株数が不正です（1以上）");

    // 本人確認
    _assertPin_(userId, pin);

    // 株が有効か
    const p = _getStockParamsFromShops_(shopId);
    if (!p.isActive) throw new Error("この会社の株は停止中です");

    // ★提示価格（買値）を取得して、その単価で約定
    const q = _getQuote_(shopId);
    const buyPrice = Number(q.buyPrice ?? q.basePrice);
    if (!Number.isFinite(buyPrice) || buyPrice <= 0) throw new Error("買値が不正です");

    // ★画面に出ている価格と一致しているか（渡ってきた場合のみチェック）
    _assertExpectedPrice_(expectedBuyPrice, buyPrice, "買値");

    const cost = shares * buyPrice;

    // お金移動：ユーザー→会社。同じロック内で処理し、二重ロックを避ける。
    const buyer = _findUser_(userId);
    const shop = _findShop_(shopId);
    if (!buyer || !buyer.active) throw new Error("ユーザーが見つからない、または無効です");
    if (!shop || !shop.active) throw new Error("会社が見つからない、または無効です");
    if (buyer.balance < cost) throw new Error("残高が足りません");
    const newUserBalance = buyer.balance - cost;
    const newShopBalance = shop.balance + cost;
    buyer.sh.getRange(buyer.row, buyer.cBal).setValue(newUserBalance);
    shop.sh.getRange(shop.row, shop.cBal).setValue(newShopBalance);
    _appendTx_({
      txId:uuid_(), at:_fmtJst_(new Date()), type:"STOCK_BUY",
      userId:buyer.userId, userName:buyer.name,
      shopId:shop.shopId, shopName:shop.shopName,
      amount:-cost, status:"OK", note:`株購入 ${shares}株 @${buyPrice}`,
      meta:JSON.stringify({shares:shares,unitPrice:buyPrice})
    });

    // 保有更新
    const h = _getHolding_(userId, shopId);
    const newShares = h.shares + shares;
    _setHolding_(userId, shopId, newShares);

    return {
      ok:true,
      quote: _getQuote_(shopId), // 約定後の新しい提示価格
      holding:{ userId, shopId, shares: newShares },
      paid: cost,
      buyPrice,
      user: {userId:buyer.userId, name:buyer.name, balance:newUserBalance}
    };
  });
}


// 株を売る（提示の売値で約定：大量でも単価固定）
function api_stockSell(userId, pin, shopId, shares, expectedSellPrice){
  return lockRun_(() => {
    userId = String(userId||"").trim().toUpperCase();
    pin    = String(pin||"").trim();
    shopId = String(shopId||"").trim().toUpperCase();
    shares = Math.floor(Number(shares||0));

    if (!userId) throw new Error("ユーザーIDが空です");
    if (!pin) throw new Error("PINを入力してください");
    if (!shopId) throw new Error("会社IDが空です");
    if (!(shares > 0)) throw new Error("株数が不正です（1以上）");

    // 本人確認
    _assertPin_(userId, pin);

    // 保有チェック
    const h = _getHolding_(userId, shopId);
    if (h.shares < shares) throw new Error("保有株数が足りません");

    // 株が有効か
    const p = _getStockParamsFromShops_(shopId);
    if (!p.isActive) throw new Error("この会社の株は停止中です");

    // ★提示価格（売値）を取得して、その単価で約定
    const q = _getQuote_(shopId);
    const sellPrice = Number(q.sellPrice ?? q.basePrice);
    if (!Number.isFinite(sellPrice) || sellPrice <= 0) throw new Error("売値が不正です");

    // ★画面に出ている価格と一致しているか（渡ってきた場合のみチェック）
    _assertExpectedPrice_(expectedSellPrice, sellPrice, "売値");

    const proceeds = shares * sellPrice;

    // 会社→ユーザー（買い戻し）
    const payRes = _shopToUserPay_(shopId, userId, proceeds, `株売却 ${shares}株 @${sellPrice}`);

    // 保有更新
    const newShares = h.shares - shares;
    _setHolding_(userId, shopId, newShares);

    return {
      ok:true,
      quote: _getQuote_(shopId), // 約定後の新しい提示価格
      holding:{ userId, shopId, shares: newShares },
      received: proceeds,
      sellPrice,
      user: payRes.user
    };
  });
}


function _assertExpectedPrice_(expected, actual, label){
  // expected が未指定ならチェックしない（後方互換）
  if (expected === null || expected === undefined || expected === "") return;

  const e = Number(expected);
  const a = Number(actual);

  if (!Number.isFinite(e)) throw new Error(`${label}（提示価格）が不正です`);
  if (e !== a) {
    throw new Error(`価格が変わりました（提示:${e} / 現在:${a}）。更新してからもう一度おねがいします。`);
  }
}
function api_adminCreateUser(adminPass, payload){
  _assertAdminPassValue_(adminPass);
  return lockRun_(() => {
    const sh = _getUsersSheet_();

    const userId = String(payload.userId || "").trim().toUpperCase();
    const name   = String(payload.name || "").trim();
    const balance = Number(payload.balance || 0);
    const pin    = String(payload.pin || "").trim();
    const isActive = payload.isActive ? true : false;

    if (!userId) throw new Error("児童IDを入力してください");
    if (!name) throw new Error("名前を入力してください");
    if (!/^\d{4}$/.test(pin)) throw new Error("PINは4桁の数字にしてください");

    const values = sh.getDataRange().getValues();
    const m = _headerMap_(values[0]);

    const cUserId = m.idx("userid");

    for (let i=1;i<values.length;i++){
      if (String(values[i][cUserId]).trim().toUpperCase() === userId){
        throw new Error("同じ児童IDがすでに存在します");
      }
    }

    sh.appendRow([
      userId,
      name,
      balance,
      isActive,
      pin,
    ]);

    return { ok:true };
  });
}function api_adminCreateShop(adminPass, payload){
  _assertAdminPassValue_(adminPass);
  return lockRun_(() => {
    const sh = _getShopsSheet_();

    const shopId = String(payload.shopId || "").trim().toUpperCase();
    const shopName = String(payload.shopName || "").trim();
    const balance = Number(payload.balance || 0);
    const pass = String(payload.pass || "").trim();
    const isActive = payload.isActive ? true : false;

    if (!shopId) throw new Error("会社IDを入力してください");
    if (!shopName) throw new Error("会社名を入力してください");
    if (!pass) throw new Error("PASSを入力してください");

    const values = sh.getDataRange().getValues();
    const m = _headerMap_(values[0]);

    const cShopId = m.idx("shopid");

    for (let i=1;i<values.length;i++){
      if (String(values[i][cShopId]).trim().toUpperCase() === shopId){
        throw new Error("同じ会社IDがすでに存在します");
      }
    }

    sh.appendRow([
      shopId,
      shopName,
      balance,
      isActive,
      pass   // ★companyPass列に入る想定
    ]);

    return { ok:true };
  });
}
function api_adminChangeAdminPass(oldPass, newPass){
  return lockRun_(() => {

    oldPass = String(oldPass || "").trim();
    newPass = String(newPass || "").trim();

    if (!oldPass || !newPass){
      throw new Error("パスワードを入力してください");
    }

    if (newPass.length < 4){
      throw new Error("新しいパスワードは4文字以上にしてください");
    }

    const current = String(getConfig_("ADMIN_PASS","")).trim();

    if (oldPass !== current){
      throw new Error("現在のパスワードが違います");
    }

    setConfig_("ADMIN_PASS", newPass);

    return { ok:true };
  });
}

function api_shopList(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Shops");
  if (!sh) throw new Error("Shopsシートが見つかりません");

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  // A〜D列取得
  const values = sh.getRange(2, 1, lastRow - 1, 4).getValues();

  return values
    .filter(r => {
      const shopId = r[0];
      const shopName = r[1];
      const isActive = r[3];

      // TRUEのみ表示（文字TRUEにも対応）
      return shopId && shopName && (isActive === true || isActive === "TRUE");
    })
    .map(r => ({
      shopId: String(r[0]).trim(),
      shopName: String(r[1]).trim()
    }));
}

function api_getActivePayOptions(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const usersSh = ss.getSheetByName("Users");
  const shopsSh = ss.getSheetByName("Shops");

  const users = usersSh.getDataRange().getValues();
  const shops = shopsSh.getDataRange().getValues();

  const userHeader = users[0];
  const shopHeader = shops[0];

  const userIdCol = userHeader.indexOf("userId");
  const nameCol = userHeader.indexOf("name");
  const userActiveCol = userHeader.indexOf("isActive");

  const shopIdCol = shopHeader.indexOf("shopId");
  const shopNameCol = shopHeader.indexOf("shopName");
  const shopActiveCol = shopHeader.indexOf("isActive");

  const activeUsers = users.slice(1)
    .filter(r => r[userActiveCol] === true || String(r[userActiveCol]).toUpperCase() === "TRUE")
    .map(r => ({
      userId: String(r[userIdCol]).trim(),
      name: String(r[nameCol]).trim()
    }))
    .filter(u => u.userId);

  const activeShops = shops.slice(1)
    .filter(r => r[shopActiveCol] === true || String(r[shopActiveCol]).toUpperCase() === "TRUE")
    .map(r => ({
      shopId: String(r[shopIdCol]).trim(),
      shopName: String(r[shopNameCol]).trim()
    }))
    .filter(s => s.shopId);

  return {
    users: activeUsers,
    shops: activeShops,
    recipients: activeShops.concat([{shopId:"GOV",shopName:"政府（権利料など）",isGovernment:true}])
  };
}

function api_getActiveUsers(){
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
  const data = sh.getDataRange().getValues();

  const header = data[0];
  const idCol = header.indexOf("userId");
  const nameCol = header.indexOf("name");
  const activeCol = header.indexOf("isActive");

  return data.slice(1)
    .filter(r => r[activeCol] === true || String(r[activeCol]).toUpperCase() === "TRUE")
    .map(r => ({
      userId: r[idCol],
      name: r[nameCol]
    }));
}







/* =========================================================
 * ClassPay Phase 1: 会社申請 / 所属管理 / 政府口座
 * ========================================================= */

function _sheetByNameOrThrow_(name){
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error(name + "シートが見つかりません。setupClassPayPhase1() を一度実行してください。");
  return sh;
}

function _getCompanyMembersSheet_(){
  return _sheetByNameOrThrow_(SHEETS.COMPANY_MEMBERS || "CompanyMembers");
}
function _getCompanyApplicationsSheet_(){
  return _sheetByNameOrThrow_(SHEETS.COMPANY_APPLICATIONS || "CompanyApplications");
}
function _getGovernmentSheet_(){
  return _sheetByNameOrThrow_(SHEETS.GOVERNMENT || "Government");
}
function _getCompanySnapshotsSheet_(){
  return _sheetByNameOrThrow_(SHEETS.COMPANY_SNAPSHOTS || "CompanySnapshots");
}

function _ensureSheetWithHeader_(name, header){
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0){
    sh.getRange(1,1,1,header.length).setValues([header]);
    sh.setFrozenRows(1);
  }
  return sh;
}

/** 初回だけ手動実行するセットアップ */
function setupClassPayPhase1(){
  const members = _ensureSheetWithHeader_(SHEETS.COMPANY_MEMBERS || "CompanyMembers", [
    "shopId","userId","role","isActive","joinedAt","leftAt","leaveReason","leftBy"
  ]);
  if (typeof _ensureColumns_ === "function") _ensureColumns_(members,["shopId","userId","role","isActive","joinedAt","leftAt","leaveReason","leftBy"]);
  const apps = _ensureSheetWithHeader_(SHEETS.COMPANY_APPLICATIONS || "CompanyApplications", [
    "applicationId","at","companyName","presidentUserId","memberUserIds","companyPass","activity","status","reviewedAt","reviewNote","shopId"
  ]);
  const gov = _ensureSheetWithHeader_(SHEETS.GOVERNMENT || "Government", [
    "accountId","accountName","balance","updatedAt"
  ]);
  _ensureSheetWithHeader_(SHEETS.COMPANY_SNAPSHOTS || "CompanySnapshots", [
    "snapshotAt","shopId","shopName","balance"
  ]);

  if (gov.getLastRow() < 2){
    gov.appendRow(["GOV","政府",0,_fmtJst_(new Date())]);
  }

  return {
    ok:true,
    sheets:[members.getName(), apps.getName(), gov.getName(), SHEETS.COMPANY_SNAPSHOTS || "CompanySnapshots"]
  };
}

function _assertAdminPassValue_(pass){
  pass = String(pass || "").trim();
  const correct = String(getConfig_("ADMIN_PASS", "")).trim();
  if (!correct) throw new Error("Configに ADMIN_PASS が設定されていません");
  if (!pass || pass !== correct) throw new Error("管理パスワードが違います");
  return true;
}

/** 管理画面専用。PINを含むため管理PASS必須 */
function api_adminListUsersWithPin(adminPass){
  _assertAdminPassValue_(adminPass);
  const sh = _getUsersSheet_();
  const v = sh.getDataRange().getValues();
  if (v.length < 2) return [];
  const m = _headerMap_(v[0]);
  const cUserId = m.idx("userid"), cName=m.idx("name"), cBal=m.idx("balance"), cActive=m.idx("isactive"), cPin=m.idx("pin");
  if (cUserId<0 || cName<0 || cBal<0 || cPin<0) throw new Error("Usersヘッダを確認してください");

  // 所属会社は児童ごとにシートを再読込せず、最初に一括で索引化する。
  const shopMap = {};
  api_listShops().forEach(s => { shopMap[s.shopId] = s; });
  const companyMap = {};
  const membersSh = _getCompanyMembersSheet_();
  const members = membersSh.getDataRange().getValues();
  if (members.length > 1) {
    const mm = _headerMap_(members[0]);
    const cMShop=mm.idx("shopid"), cMUser=mm.idx("userid"), cMRole=mm.idx("role"), cMActive=mm.idx("isactive");
    members.slice(1).forEach(r => {
      const userId=String(r[cMUser]||"").trim().toUpperCase();
      const shopId=String(r[cMShop]||"").trim().toUpperCase();
      if (!userId || !shopMap[shopId] || (cMActive>=0 && !_isTrue_(r[cMActive]))) return;
      if (!companyMap[userId]) companyMap[userId]=[];
      companyMap[userId].push({shopId,shopName:shopMap[shopId].shopName,role:String(r[cMRole]||"MEMBER").trim().toUpperCase()});
    });
  }

  return v.slice(1).map(r => {
    const userId = String(r[cUserId] || "").trim().toUpperCase();
    if (!userId) return null;
    return {
      userId,
      name:String(r[cName] || "").trim(),
      balance:Number(r[cBal] || 0),
      isActive:cActive<0 ? true : _isTrue_(r[cActive]),
      pin:_normalizePin_(r[cPin]),
      companies:(companyMap[userId]||[]).sort((a,b)=>{
        if (a.role==="PRESIDENT" && b.role!=="PRESIDENT") return -1;
        if (b.role==="PRESIDENT" && a.role!=="PRESIDENT") return 1;
        return a.shopName.localeCompare(b.shopName,"ja");
      })
    };
  }).filter(Boolean).sort((a,b)=>a.userId.localeCompare(b.userId));
}

function _getCompaniesByUserId_(userId){
  userId = String(userId || "").trim().toUpperCase();
  if (!userId) return [];
  const sh = _getCompanyMembersSheet_();
  const v = sh.getDataRange().getValues();
  if (v.length < 2) return [];
  const m = _headerMap_(v[0]);
  const cShop=m.idx("shopid"), cUser=m.idx("userid"), cRole=m.idx("role"), cActive=m.idx("isactive");
  const out=[];
  v.slice(1).forEach(r=>{
    if (String(r[cUser]||"").trim().toUpperCase() !== userId) return;
    if (cActive>=0 && !_isTrue_(r[cActive])) return;
    const shopId=String(r[cShop]||"").trim().toUpperCase();
    const s=_findShop_(shopId);
    if (!s || !s.active) return;
    out.push({shopId:s.shopId, shopName:s.shopName, role:String(r[cRole]||"MEMBER").trim().toUpperCase()});
  });
  return out.sort((a,b)=>{
    if (a.role==="PRESIDENT" && b.role!=="PRESIDENT") return -1;
    if (b.role==="PRESIDENT" && a.role!=="PRESIDENT") return 1;
    return a.shopName.localeCompare(b.shopName,"ja");
  });
}

function api_myCompanies(userId, pin){
  const u = _assertPin_(String(userId||"").trim().toUpperCase(), pin);
  return _getCompaniesByUserId_(u.userId);
}

function _membershipHistoryRows_(onlyUserId){
  onlyUserId=String(onlyUserId||"").trim().toUpperCase();
  const sh=_getCompanyMembersSheet_(),v=sh.getDataRange().getValues();
  if(v.length<2)return [];
  const m=_headerMap_(v[0]),ix=n=>m.idx(n),shopNames={},userNames={};
  const sv=_getShopsSheet_().getDataRange().getValues();
  if(sv.length>1){const sm=_headerMap_(sv[0]),cId=sm.idx("shopid"),cName=sm.idx("shopname"),cActive=sm.idx("isactive");sv.slice(1).forEach(r=>{const id=String(r[cId]||"").trim().toUpperCase();if(id)shopNames[id]={name:String(r[cName]||""),active:cActive<0?true:_isTrue_(r[cActive])};});}
  const uv=_getUsersSheet_().getDataRange().getValues();
  if(uv.length>1){const um=_headerMap_(uv[0]),cId=um.idx("userid"),cName=um.idx("name");uv.slice(1).forEach(r=>{const id=String(r[cId]||"").trim().toUpperCase();if(id)userNames[id]=String(r[cName]||"");});}
  return v.slice(1).map((r,i)=>{
    const userId=String(r[ix("userid")]||"").trim().toUpperCase();
    if(!userId||(onlyUserId&&userId!==onlyUserId))return null;
    const shopId=String(r[ix("shopid")]||"").trim().toUpperCase(),shop=shopNames[shopId]||{name:shopId,active:false};
    return {rowNumber:i+2,shopId,shopName:shop.name,userId,userName:userNames[userId]||userId,role:String(r[ix("role")]||"MEMBER").trim().toUpperCase(),isActive:(ix("isactive")<0||_isTrue_(r[ix("isactive")]))&&shop.active,joinedAt:String(r[ix("joinedat")]||""),leftAt:ix("leftat")<0?"":String(r[ix("leftat")]||""),leaveReason:ix("leavereason")<0?"":String(r[ix("leavereason")]||""),leftBy:ix("leftby")<0?"":String(r[ix("leftby")]||"")};
  }).filter(Boolean).sort((a,b)=>String(b.joinedAt).localeCompare(String(a.joinedAt)));
}

function api_myMembershipHistory(userId,pin){
  const u=_assertPin_(String(userId||"").trim().toUpperCase(),pin);
  return _membershipHistoryRows_(u.userId);
}

function api_leaveCompany(userId,pin,shopId,reason){
  return lockRun_(()=>{
    const u=_assertPin_(String(userId||"").trim().toUpperCase(),pin);
    shopId=String(shopId||"").trim().toUpperCase();
    const count=_endCompanyMemberRows_(shopId,u.userId,String(reason||"本人による退社").trim()||"本人による退社",u.userId,false);
    if(!count)throw new Error("有効な所属が見つかりません");
    return {ok:true,userId:u.userId,shopId};
  });
}

function api_adminMembershipHistory(adminPass,userId){
  _assertAdminPassValue_(adminPass);
  return _membershipHistoryRows_(userId);
}

function _hasPresidentCompany_(userId){
  return _getCompaniesByUserId_(userId).some(x => x.role === "PRESIDENT");
}

function _hasPendingPresidentApplication_(userId){
  const sh=_getCompanyApplicationsSheet_();
  const v=sh.getDataRange().getValues();
  if (v.length<2) return false;
  const m=_headerMap_(v[0]);
  const cPresident=m.idx("presidentuserid"), cStatus=m.idx("status");
  return v.slice(1).some(r =>
    String(r[cPresident]||"").trim().toUpperCase()===String(userId||"").trim().toUpperCase() &&
    String(r[cStatus]||"").trim().toUpperCase()==="PENDING"
  );
}

/** 会社設立申請フォームの児童候補 */
function api_companyApplicationUsers(){
  return api_listUsers().map(u => ({userId:u.userId, name:u.name}));
}

/** オンライン会社設立申請 */
function api_submitCompanyApplication(payload){
  return lockRun_(() => {
    payload = payload || {};
    const companyName=String(payload.companyName||"").trim();
    const presidentUserId=String(payload.presidentUserId||"").trim().toUpperCase();
    let memberUserIds=Array.isArray(payload.memberUserIds) ? payload.memberUserIds : [];
    memberUserIds=memberUserIds.map(x=>String(x||"").trim().toUpperCase()).filter(Boolean);
    const companyPass=String(payload.companyPass||"").trim();
    const activity=String(payload.activity||"").trim();

    if (!companyName) throw new Error("会社名を入力してください");
    if (!presidentUserId) throw new Error("社長を選択してください");
    if (!companyPass) throw new Error("会社パスワードを入力してください");
    if (!activity) throw new Error("会社の活動内容を入力してください");
    if (!_findUser_(presidentUserId)) throw new Error("社長の児童が見つかりません");
    if (_hasPresidentCompany_(presidentUserId)) throw new Error("この児童はすでに別の会社の社長です");
    if (_hasPendingPresidentApplication_(presidentUserId)) throw new Error("この児童を社長とする申請がすでに審査中です");

    memberUserIds = Array.from(new Set(memberUserIds.filter(id => id !== presidentUserId)));
    memberUserIds.forEach(id => { if (!_findUser_(id)) throw new Error("社員の児童が見つかりません："+id); });

    const shops = api_listShops();
    if (shops.some(s => String(s.shopName||"").trim() === companyName)) throw new Error("同じ会社名がすでに使われています");

    const sh=_getCompanyApplicationsSheet_();
    const applicationId="APP-" + Utilities.formatDate(new Date(),"Asia/Tokyo","yyyyMMdd-HHmmss") + "-" + uuid_().slice(0,6).toUpperCase();
    const at=_fmtJst_(new Date());
    sh.appendRow([applicationId,at,companyName,presidentUserId,JSON.stringify(memberUserIds),companyPass,activity,"PENDING","","",""]);
    return {ok:true, applicationId, at, status:"PENDING"};
  });
}

function api_adminListCompanyApplications(adminPass, status){
  _assertAdminPassValue_(adminPass);
  status=String(status||"").trim().toUpperCase();
  const sh=_getCompanyApplicationsSheet_();
  const v=sh.getDataRange().getValues();
  if (v.length<2) return [];
  const m=_headerMap_(v[0]);
  const ix=n=>m.idx(n);
  const userMap={};
  api_listUsers().forEach(u=>{userMap[u.userId]=u;});
  return v.slice(1).map(r=>{
    const presidentUserId=String(r[ix("presidentuserid")]||"").trim().toUpperCase();
    const p=userMap[presidentUserId];
    let memberIds=[];
    try{ memberIds=JSON.parse(String(r[ix("memberuserids")]||"[]")); }catch(e){}
    const members=memberIds.map(id=>{ const u=userMap[String(id||"").trim().toUpperCase()]; return u?{userId:u.userId,name:u.name}:null; }).filter(Boolean);
    return {
      applicationId:String(r[ix("applicationid")]||""),
      at:String(r[ix("at")]||""),
      companyName:String(r[ix("companyname")]||""),
      president:p?{userId:p.userId,name:p.name}:{userId:presidentUserId,name:""},
      members,
      companyPass:String(r[ix("companypass")]||""),
      activity:String(r[ix("activity")]||""),
      status:String(r[ix("status")]||"").toUpperCase(),
      reviewedAt:String(r[ix("reviewedat")]||""),
      reviewNote:String(r[ix("reviewnote")]||""),
      shopId:String(r[ix("shopid")]||"")
    };
  }).filter(x=>!status || x.status===status).reverse();
}

function _nextShopId_(){
  const sh=_getShopsSheet_();
  const v=sh.getDataRange().getValues();
  const m=_headerMap_(v[0]);
  const c=m.idx("shopid");
  let max=0;
  v.slice(1).forEach(r=>{
    const mm=String(r[c]||"").trim().toUpperCase().match(/^S(\d+)$/);
    if(mm) max=Math.max(max,Number(mm[1]));
  });
  return "S" + String(max+1).padStart(3,"0");
}

function _findApplicationRow_(applicationId){
  const sh=_getCompanyApplicationsSheet_();
  const v=sh.getDataRange().getValues();
  if(v.length<2) return null;
  const m=_headerMap_(v[0]);
  const c=m.idx("applicationid");
  for(let i=1;i<v.length;i++){
    if(String(v[i][c]||"").trim()===String(applicationId||"").trim()) return {sh,row:i+1,values:v[i],map:m};
  }
  return null;
}

/** 承認すると Shops と CompanyMembers を自動作成 */
function api_adminApproveCompanyApplication(adminPass, applicationId){
  _assertAdminPassValue_(adminPass);
  return lockRun_(() => {
    const a=_findApplicationRow_(applicationId);
    if(!a) throw new Error("申請が見つかりません");
    const idx=n=>a.map.idx(n);
    const status=String(a.values[idx("status")]||"").trim().toUpperCase();
    if(status!=="PENDING") throw new Error("この申請はすでに処理済みです");

    const companyName=String(a.values[idx("companyname")]||"").trim();
    const presidentUserId=String(a.values[idx("presidentuserid")]||"").trim().toUpperCase();
    const companyPass=String(a.values[idx("companypass")]||"").trim();
    let memberIds=[]; try{memberIds=JSON.parse(String(a.values[idx("memberuserids")]||"[]"));}catch(e){}

    if(_hasPresidentCompany_(presidentUserId)) throw new Error("この児童はすでに別会社の社長です");
    const shopId=_nextShopId_();
    _appendShopPhase2_(shopId,companyName,companyPass);

    const joinedAt=_fmtJst_(new Date());
    const memSh=_getCompanyMembersSheet_();
    const rows=[[shopId,presidentUserId,"PRESIDENT"]];
    Array.from(new Set(memberIds.map(x=>String(x||"").trim().toUpperCase()).filter(Boolean)))
      .filter(id=>id!==presidentUserId)
      .forEach(id=>rows.push([shopId,id,"MEMBER"]));
    rows.forEach(r=>_appendCompanyMember_(r[0],r[1],r[2],joinedAt));

    const now=_fmtJst_(new Date());
    a.sh.getRange(a.row, idx("status")+1).setValue("APPROVED");
    a.sh.getRange(a.row, idx("reviewedat")+1).setValue(now);
    a.sh.getRange(a.row, idx("shopid")+1).setValue(shopId);
    return {ok:true,shopId,companyName,status:"APPROVED"};
  });
}

function api_adminRejectCompanyApplication(adminPass, applicationId, reviewNote){
  _assertAdminPassValue_(adminPass);
  return lockRun_(() => {
    const a=_findApplicationRow_(applicationId);
    if(!a) throw new Error("申請が見つかりません");
    const idx=n=>a.map.idx(n);
    if(String(a.values[idx("status")]||"").trim().toUpperCase()!=="PENDING") throw new Error("この申請はすでに処理済みです");
    a.sh.getRange(a.row,idx("status")+1).setValue("REJECTED");
    a.sh.getRange(a.row,idx("reviewedat")+1).setValue(_fmtJst_(new Date()));
    a.sh.getRange(a.row,idx("reviewnote")+1).setValue(String(reviewNote||"").trim());
    return {ok:true,status:"REJECTED"};
  });
}

function _getGovernmentAccount_(){
  const sh=_getGovernmentSheet_();
  const v=sh.getDataRange().getValues();
  if(v.length<2) throw new Error("Government口座がありません。setupClassPayPhase1() を実行してください");
  const m=_headerMap_(v[0]);
  const cId=m.idx("accountid"), cName=m.idx("accountname"), cBal=m.idx("balance"), cUpdated=m.idx("updatedat");
  for(let i=1;i<v.length;i++){
    if(String(v[i][cId]||"").trim().toUpperCase()==="GOV"){
      return {sh,row:i+1,cBal:cBal+1,cUpdated:cUpdated+1,accountId:"GOV",accountName:String(v[i][cName]||"政府"),balance:Number(v[i][cBal]||0)};
    }
  }
  throw new Error("Governmentシートに GOV 行がありません");
}

function api_adminGovernmentInfo(adminPass){
  _assertAdminPassValue_(adminPass);
  const g=_getGovernmentAccount_();
  return {accountId:g.accountId,accountName:g.accountName,balance:g.balance};
}

/** 政府残高はマイナス可。第2段階で税・利息・支援金をこの関数へ接続する */
function _changeGovernmentBalance_(delta){
  delta=Number(delta);
  if(!Number.isFinite(delta)) throw new Error("政府口座の増減額が不正です");
  const g=_getGovernmentAccount_();
  const next=g.balance+delta;
  g.sh.getRange(g.row,g.cBal).setValue(next);
  if(g.cUpdated>0) g.sh.getRange(g.row,g.cUpdated).setValue(_fmtJst_(new Date()));
  return next;
}

function api_adminAdjustGovernment(adminPass, amount, reason){
  _assertAdminPassValue_(adminPass);
  return lockRun_(() => {
    amount=Number(amount); reason=String(reason||"").trim();
    if(!Number.isFinite(amount)||amount===0) throw new Error("調整額は0以外の数値にしてください");
    const balance=_changeGovernmentBalance_(amount);
    return {ok:true,balance,amount,reason};
  });
}
