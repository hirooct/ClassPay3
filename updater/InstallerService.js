function updater_installClassPay_(payload) {
  var appName = String(payload.appName || "ClassPay").trim() || "ClassPay";
  var adminPass = String(payload.adminPass || "").trim();
  var manifestUrl = String(payload.manifestUrl || "https://raw.githubusercontent.com/hirooct/ClassPay3/main/release/manifest.json").trim();
  if (adminPass.length < 4 || adminPass !== String(payload.confirmPass || "").trim()) throw new Error("管理パスワードを4文字以上で正しく確認入力してください");
  var manifest = updater_fetchJson_(manifestUrl); updater_validateManifest_(manifest);
  var spreadsheet = SpreadsheetApp.create(appName + " データ"), ssId = spreadsheet.getId();
  try {
    updater_ensureSheet_(spreadsheet,"Users",["userId","name","balance","isActive","pin"]);
    updater_ensureSheet_(spreadsheet,"Shops",["shopId","shopName","balance","isActive","companyPass","URL"]);
    updater_ensureSheet_(spreadsheet,"Tx",["txId","at","type","userId","userName","shopId","shopName","amount","status","note","meta"]);
    updater_ensureSheet_(spreadsheet,"Config",["key","value"]);
    updater_setConfigValue_(spreadsheet,"APP_NAME",appName,false);
    updater_setConfigValue_(spreadsheet,"ADMIN_PASS",adminPass,false);
    updater_setConfigValue_(spreadsheet,"CLASS_PAY_VERSION","0.0.0",false);
    updater_setConfigValue_(spreadsheet,"RELEASE_MANIFEST_URL",manifestUrl,false);
    updater_setConfigValue_(spreadsheet,"RELEASE_CHANNEL","stable",false);
    var project = updater_scriptApi_("post","/projects",{title:appName,parentId:ssId});
    updater_putProjectContent_(project.scriptId,updater_loadReleaseFiles_(manifest));
    updater_runMigrations_(ssId,manifest,"0.0.0");
    var version = updater_createVersion_(project.scriptId,"ClassPay v"+manifest.latestVersion);
    var deployment = updater_createDeployment_(project.scriptId,version.versionNumber,"ClassPay v"+manifest.latestVersion);
    var deploymentId = deployment.deploymentId || "", webAppUrl = "";
    (deployment.entryPoints || []).forEach(function(entry){ if(entry.webApp && entry.webApp.url) webAppUrl = entry.webApp.url; });
    updater_setConfigValue_(spreadsheet,"BASE_URL",webAppUrl,false);
    updater_setConfigValue_(spreadsheet,"DEPLOYMENT_ID",deploymentId,false);
    updater_setConfigValue_(spreadsheet,"UPDATER_URL",ScriptApp.getService().getUrl() || "",false);
    updater_saveTarget_({scriptId:project.scriptId,spreadsheetId:ssId,deploymentId:deploymentId,webAppUrl:webAppUrl,manifestUrl:manifestUrl,currentVersion:manifest.latestVersion});
    return {ok:true,scriptId:project.scriptId,spreadsheetId:ssId,deploymentId:deploymentId,webAppUrl:webAppUrl,spreadsheetUrl:spreadsheet.getUrl(),version:manifest.latestVersion};
  } catch (error) {
    throw new Error("作成途中で停止しました。作成済みSpreadsheet ID: " + ssId + " / " + error.message);
  }
}
