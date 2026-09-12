/* =========================================================
 * ClassPay 3.3: 外部Updater連携（本体は自分自身を書き換えない）
 * ========================================================= */

const CLASS_PAY_VERSION = "3.3.0";
const DEFAULT_RELEASE_MANIFEST_URL = "https://raw.githubusercontent.com/hirooct/ClassPay3/main/release/manifest.json";

function _semverParts_(value) {
  var match = String(value || "0.0.0").trim().replace(/^v/i, "").match(/^(\d+)\.(\d+)\.(\d+)/);
  return match ? [Number(match[1]), Number(match[2]), Number(match[3])] : [0, 0, 0];
}

function _compareSemver_(left, right) {
  var a = _semverParts_(left), b = _semverParts_(right);
  for (var i = 0; i < 3; i++) if (a[i] !== b[i]) return a[i] < b[i] ? -1 : 1;
  return 0;
}

function _releaseManifestUrl_() {
  return String(getConfig_("RELEASE_MANIFEST_URL", DEFAULT_RELEASE_MANIFEST_URL)).trim() || DEFAULT_RELEASE_MANIFEST_URL;
}

function _fetchReleaseManifest_() {
  var response = UrlFetchApp.fetch(_releaseManifestUrl_(), {
    muteHttpExceptions: true,
    followRedirects: true,
    headers: { Accept: "application/json" }
  });
  if (response.getResponseCode() !== 200) throw new Error("更新情報を取得できません（HTTP " + response.getResponseCode() + "）");
  var manifest = JSON.parse(response.getContentText());
  if (!manifest.latestVersion || !manifest.releaseDate || !manifest.updateType) throw new Error("更新情報の形式が正しくありません");
  return manifest;
}

function api_adminUpdateStatus(adminPass) {
  _assertAdminPassValue_(adminPass);
  var current = String(getConfig_("CLASS_PAY_VERSION", CLASS_PAY_VERSION));
  var manifest = _fetchReleaseManifest_();
  var minimum = String(manifest.minimumVersion || "0.0.0");
  return {
    currentVersion: current,
    latestVersion: String(manifest.latestVersion),
    releaseDate: String(manifest.releaseDate),
    updateType: String(manifest.updateType).toUpperCase(),
    releaseNotes: manifest.releaseNotes || [],
    updateAvailable: _compareSemver_(current, manifest.latestVersion) < 0,
    compatible: _compareSemver_(current, minimum) >= 0,
    minimumVersion: minimum,
    migrationRequired: !!manifest.migrationRequired,
    updaterConfigured: !!String(getConfig_("UPDATER_URL", "")).trim()
  };
}

function api_adminUpdaterLaunchInfo(adminPass) {
  _assertAdminPassValue_(adminPass);
  var updaterUrl = String(getConfig_("UPDATER_URL", "")).trim();
  if (!updaterUrl) throw new Error("初期設定でUPDATER_URLを登録してください");
  return {
    updaterUrl: updaterUrl,
    scriptId: ScriptApp.getScriptId(),
    spreadsheetId: SpreadsheetApp.getActive().getId(),
    deploymentId: String(getConfig_("DEPLOYMENT_ID", "")).trim(),
    webAppUrl: String(getConfig_("BASE_URL", "")).trim(),
    manifestUrl: _releaseManifestUrl_(),
    currentVersion: String(getConfig_("CLASS_PAY_VERSION", CLASS_PAY_VERSION))
  };
}

function api_adminSaveUpdaterSettings(adminPass, payload) {
  _assertAdminPassValue_(adminPass);
  payload = payload || {};
  _setConfigValue_("UPDATER_URL", String(payload.updaterUrl || "").trim());
  _setConfigValue_("DEPLOYMENT_ID", String(payload.deploymentId || "").trim());
  _setConfigValue_("RELEASE_MANIFEST_URL", String(payload.manifestUrl || DEFAULT_RELEASE_MANIFEST_URL).trim() || DEFAULT_RELEASE_MANIFEST_URL);
  return { ok: true };
}

function api_adminUpdateHistory(adminPass, limit) {
  _assertAdminPassValue_(adminPass);
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.UPDATE_HISTORY || "UpdateHistory");
  if (!sh || sh.getLastRow() < 2) return [];
  var values = sh.getDataRange().getValues(), headers = values[0].map(function(x){ return String(x); });
  return values.slice(1).reverse().slice(0, Math.max(1, Math.min(50, Number(limit || 10)))).map(function(row){
    var result = {};
    headers.forEach(function(key, i){ result[key] = row[i]; });
    return result;
  });
}
