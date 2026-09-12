const UPDATER_VERSION = "1.0.0";
const UPDATER_PROP_TARGET = "CLASSPAY_TARGET";
const UPDATER_BACKUP_FOLDER = "ClassPay Backups";

function doGet(e) {
  var t = HtmlService.createTemplateFromFile("ui_index");
  t.params = (e && e.parameter) ? e.parameter : {};
  return t.evaluate().setTitle("ClassPay Updater").setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function api_updaterBootstrap(params) {
  var saved = updater_loadTarget_();
  params = params || {};
  ["scriptId", "spreadsheetId", "deploymentId", "webAppUrl", "manifestUrl", "currentVersion"].forEach(function(key) {
    if (params[key]) saved[key] = String(params[key]).trim();
  });
  if (saved.scriptId || saved.spreadsheetId) updater_saveTarget_(saved);
  return { updaterVersion: UPDATER_VERSION, email: Session.getActiveUser().getEmail() || "", target: saved };
}

function api_updaterSaveTarget(target) {
  target = updater_normalizeTarget_(target);
  if (!target.scriptId || !target.spreadsheetId) throw new Error("Script IDとSpreadsheet IDが必要です");
  updater_saveTarget_(target);
  return { ok: true, target: target };
}

function api_updaterInspect(target) {
  target = updater_targetOrSaved_(target);
  var manifest = updater_fetchJson_(target.manifestUrl);
  updater_validateManifest_(manifest);
  var current = updater_getConfigValue_(target.spreadsheetId, "CLASS_PAY_VERSION", target.currentVersion || "3.2.0");
  var deployment = target.deploymentId ? updater_scriptApi_("get", "/projects/" + encodeURIComponent(target.scriptId) + "/deployments/" + encodeURIComponent(target.deploymentId)) : null;
  return {
    target: target,
    currentVersion: current,
    latestVersion: manifest.latestVersion,
    minimumVersion: manifest.minimumVersion,
    updateType: manifest.updateType,
    releaseDate: manifest.releaseDate,
    releaseNotes: manifest.releaseNotes || [],
    migrationRequired: !!manifest.migrationRequired,
    compatible: updater_compareSemver_(current, manifest.minimumVersion || "0.0.0") >= 0,
    updateAvailable: updater_compareSemver_(current, manifest.latestVersion) < 0,
    deploymentVersion: deployment && deployment.deploymentConfig ? deployment.deploymentConfig.versionNumber : null
  };
}

function api_updaterRun(target, options) {
  return updater_runUpdate_(updater_targetOrSaved_(target), options || {});
}

function api_updaterBackups(target) {
  target = updater_targetOrSaved_(target);
  return updater_listBackups_(target.scriptId, 20);
}

function api_updaterRollback(target, backupFolderId) {
  return updater_runRollback_(updater_targetOrSaved_(target), String(backupFolderId || ""));
}

function api_installerCreate(payload) {
  return updater_installClassPay_(payload || {});
}

function updater_normalizeTarget_(target) {
  target = target || {};
  return {
    scriptId: String(target.scriptId || "").trim(),
    spreadsheetId: String(target.spreadsheetId || "").trim(),
    deploymentId: String(target.deploymentId || "").trim(),
    webAppUrl: String(target.webAppUrl || "").trim(),
    manifestUrl: String(target.manifestUrl || "https://raw.githubusercontent.com/hirooct/ClassPay3/main/release/manifest.json").trim(),
    currentVersion: String(target.currentVersion || "").trim()
  };
}

function updater_targetOrSaved_(target) {
  var merged = updater_loadTarget_(), incoming = updater_normalizeTarget_(target);
  Object.keys(incoming).forEach(function(key) { if (incoming[key]) merged[key] = incoming[key]; });
  merged = updater_normalizeTarget_(merged);
  if (!merged.scriptId || !merged.spreadsheetId || !merged.manifestUrl) throw new Error("更新対象の設定が不足しています");
  updater_saveTarget_(merged);
  return merged;
}

function updater_saveTarget_(target) {
  PropertiesService.getUserProperties().setProperty(UPDATER_PROP_TARGET, JSON.stringify(updater_normalizeTarget_(target)));
}

function updater_loadTarget_() {
  try { return updater_normalizeTarget_(JSON.parse(PropertiesService.getUserProperties().getProperty(UPDATER_PROP_TARGET) || "{}")); }
  catch (e) { return updater_normalizeTarget_({}); }
}
