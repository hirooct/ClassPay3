function updater_getConfigValue_(spreadsheetId, key, fallback) {
  var ss = SpreadsheetApp.openById(spreadsheetId), sh = ss.getSheetByName("Config");
  if (!sh || sh.getLastRow() < 2) return fallback;
  var values = sh.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) if (String(values[i][0]).trim() === key) return values[i][1];
  return fallback;
}

function updater_setConfigValue_(ss, key, value, onlyIfMissing) {
  var sh = ss.getSheetByName("Config") || ss.insertSheet("Config");
  if (sh.getLastColumn() < 2 || sh.getLastRow() < 1) sh.getRange(1,1,1,2).setValues([["key","value"]]);
  var values = sh.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim() === key) { if (!onlyIfMissing) sh.getRange(i+1,2).setValue(value); return; }
  }
  sh.appendRow([key,value]);
}

function updater_ensureSheet_(ss, name, headers) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  updater_ensureColumns_(sh, headers || []);
  return sh;
}

function updater_ensureColumns_(sh, headers) {
  if (!headers.length) return;
  if (sh.getLastRow() < 1 || sh.getLastColumn() < 1 || !sh.getRange(1,1).getValue()) {
    sh.getRange(1,1,1,headers.length).setValues([headers]); sh.setFrozenRows(1); return;
  }
  var width = sh.getLastColumn(), existing = sh.getRange(1,1,1,width).getValues()[0].map(function(x){ return String(x).trim().toLowerCase(); });
  headers.forEach(function(header) { if (existing.indexOf(String(header).toLowerCase()) < 0) { sh.getRange(1, sh.getLastColumn()+1).setValue(header); existing.push(String(header).toLowerCase()); } });
  sh.setFrozenRows(1);
}

function updater_fetchMigration_(spec) {
  var response = UrlFetchApp.fetch(spec.url, { muteHttpExceptions: true, followRedirects: true });
  if (response.getResponseCode() !== 200) throw new Error("Migrationを取得できません: " + spec.from + " → " + spec.to);
  var text = response.getContentText();
  if (updater_sha256_(text).toLowerCase() !== String(spec.sha256).toLowerCase()) throw new Error("MigrationのSHA-256が一致しません");
  return JSON.parse(text);
}

function updater_buildMigrationPath_(manifest, current, latest) {
  var path = [], cursor = current, specs = manifest.migrations || [], guard = 0;
  while (updater_compareSemver_(cursor, latest) < 0 && guard++ < 50) {
    var next = specs.filter(function(x){ return String(x.from) === String(cursor); }).sort(function(a,b){ return updater_compareSemver_(a.to,b.to); })[0];
    if (!next) throw new Error(cursor + "からのMigrationがありません");
    path.push(next); cursor = next.to;
  }
  if (String(cursor) !== String(latest)) throw new Error("Migration経路が最新版までつながっていません");
  return path;
}

function updater_runMigrations_(spreadsheetId, manifest, currentVersion) {
  var ss = SpreadsheetApp.openById(spreadsheetId), path = updater_buildMigrationPath_(manifest, currentVersion, manifest.latestVersion), applied = [];
  path.forEach(function(spec) {
    var migration = updater_fetchMigration_(spec);
    (migration.actions || []).forEach(function(action) {
      if (action.type === "ensureSheet") updater_ensureSheet_(ss, action.name, action.headers || []);
      else if (action.type === "ensureColumns") updater_ensureColumns_(updater_ensureSheet_(ss, action.sheet, []), action.headers || []);
      else if (action.type === "setConfigDefault") updater_setConfigValue_(ss, action.key, action.value, true);
      else if (action.type === "setConfig") updater_setConfigValue_(ss, action.key, action.value, false);
      else throw new Error("未対応のMigration action: " + action.type);
    });
    updater_setConfigValue_(ss, "CLASS_PAY_VERSION", spec.to, false);
    applied.push(spec.from + "→" + spec.to);
  });
  SpreadsheetApp.flush();
  return applied;
}

function updater_appendHistory_(spreadsheetId, row) {
  var ss = SpreadsheetApp.openById(spreadsheetId), headers = ["updateId","startedAt","finishedAt","fromVersion","toVersion","status","backupFolderId","gasVersion","deploymentId","message"];
  var sh = updater_ensureSheet_(ss, "UpdateHistory", headers), values = headers.map(function(key){ return row[key] === undefined ? "" : row[key]; });
  sh.appendRow(values);
}

function updater_protectedDataFingerprint_(spreadsheetId) {
  var ss = SpreadsheetApp.openById(spreadsheetId), result = {};
  ["Users","Shops","Tx","Holdings"].forEach(function(name) {
    var sh = ss.getSheetByName(name);
    if (!sh) { result[name] = { rows: 0, balance: 0 }; return; }
    var values = sh.getDataRange().getValues(), balance = 0, balanceIndex = values.length ? values[0].map(function(x){return String(x).toLowerCase();}).indexOf("balance") : -1;
    if (balanceIndex >= 0) for (var i = 1; i < values.length; i++) balance += Number(values[i][balanceIndex] || 0);
    result[name] = { rows: Math.max(0, values.length - 1), balance: balance };
  });
  return JSON.stringify(result);
}
