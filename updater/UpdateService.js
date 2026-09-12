function updater_runUpdate_(target, options) {
  var lock = LockService.getUserLock(); lock.waitLock(30000);
  var startedAt = new Date().toISOString(), backup = null, oldContent = null, oldDeployment = null, contentChanged = false;
  var current = updater_getConfigValue_(target.spreadsheetId, "CLASS_PAY_VERSION", target.currentVersion || "3.2.0"), manifest;
  try {
    manifest = updater_fetchJson_(target.manifestUrl); updater_validateManifest_(manifest);
    if (updater_compareSemver_(current, manifest.minimumVersion || "0.0.0") < 0) throw new Error("現在版が古いため、中間版から更新してください");
    if (updater_compareSemver_(current, manifest.latestVersion) >= 0 && !options.force) return { ok:true, skipped:true, message:"すでに最新版です", currentVersion:current };
    oldContent = updater_getProjectContent_(target.scriptId);
    if (target.deploymentId) oldDeployment = updater_scriptApi_("get", "/projects/" + encodeURIComponent(target.scriptId) + "/deployments/" + encodeURIComponent(target.deploymentId));
    backup = updater_createBackup_(target, current, oldContent, oldDeployment);
    var protectedBefore = updater_protectedDataFingerprint_(target.spreadsheetId);
    var releaseFiles = updater_loadReleaseFiles_(manifest);
    updater_putProjectContent_(target.scriptId, releaseFiles); contentChanged = true;
    updater_verifyProjectFiles_(target.scriptId, releaseFiles);
    var migrations = updater_runMigrations_(target.spreadsheetId, manifest, current);
    if (protectedBefore !== updater_protectedDataFingerprint_(target.spreadsheetId)) throw new Error("保護対象データの件数または残高が変化したため停止しました");
    var version = updater_createVersion_(target.scriptId, "ClassPay v" + manifest.latestVersion);
    var deployment = target.deploymentId ? updater_updateDeployment_(target.scriptId, target.deploymentId, version.versionNumber, "ClassPay v" + manifest.latestVersion) : null;
    if (target.deploymentId) {
      var deployed = updater_scriptApi_("get", "/projects/" + encodeURIComponent(target.scriptId) + "/deployments/" + encodeURIComponent(target.deploymentId));
      if (!deployed.deploymentConfig || Number(deployed.deploymentConfig.versionNumber) !== Number(version.versionNumber)) throw new Error("Deploymentの更新確認に失敗しました");
    }
    var finishedAt = new Date().toISOString();
    updater_appendHistory_(target.spreadsheetId,{updateId:Utilities.getUuid(),startedAt:startedAt,finishedAt:finishedAt,fromVersion:current,toVersion:manifest.latestVersion,status:"SUCCESS",backupFolderId:backup.folderId,gasVersion:version.versionNumber,deploymentId:target.deploymentId,message:"更新完了 / Migration: "+(migrations.join(", ")||"なし")});
    return {ok:true,fromVersion:current,toVersion:manifest.latestVersion,backupFolderId:backup.folderId,gasVersion:version.versionNumber,deployment:deployment,migrations:migrations};
  } catch (error) {
    var rollbackMessage = "";
    if (contentChanged && oldContent) {
      try { updater_putProjectContent_(target.scriptId, oldContent.files); rollbackMessage += "コード復元済み。"; } catch (restoreError) { rollbackMessage += "コード復元失敗: " + restoreError.message + "。"; }
    }
    if (oldDeployment && oldDeployment.deploymentConfig && target.deploymentId) {
      try { updater_updateDeployment_(target.scriptId,target.deploymentId,oldDeployment.deploymentConfig.versionNumber,oldDeployment.deploymentConfig.description||"ClassPay rollback"); rollbackMessage += "Deployment復元済み。"; } catch (deploymentError) { rollbackMessage += "Deployment復元失敗: " + deploymentError.message + "。"; }
    }
    try { updater_appendHistory_(target.spreadsheetId,{updateId:Utilities.getUuid(),startedAt:startedAt,finishedAt:new Date().toISOString(),fromVersion:current,toVersion:manifest?manifest.latestVersion:"",status:"FAILED",backupFolderId:backup?backup.folderId:"",gasVersion:"",deploymentId:target.deploymentId,message:error.message+" / "+rollbackMessage}); } catch (ignore) {}
    throw new Error(error.message + (rollbackMessage ? " / " + rollbackMessage : ""));
  } finally { lock.releaseLock(); }
}

function updater_runRollback_(target, backupFolderId) {
  if (!backupFolderId) throw new Error("戻すバックアップを選択してください");
  var lock = LockService.getUserLock(); lock.waitLock(30000);
  try {
    var backup = updater_readBackup_(backupFolderId), meta = backup.metadata;
    if (String(meta.target.scriptId) !== target.scriptId) throw new Error("別のClassPayのバックアップです");
    var beforeVersion = updater_getConfigValue_(target.spreadsheetId,"CLASS_PAY_VERSION","");
    updater_putProjectContent_(target.scriptId, backup.source.files);
    var version = updater_createVersion_(target.scriptId, "ClassPay rollback to " + meta.currentVersion);
    if (target.deploymentId) updater_updateDeployment_(target.scriptId,target.deploymentId,version.versionNumber,"ClassPay rollback to "+meta.currentVersion);
    updater_setConfigValue_(SpreadsheetApp.openById(target.spreadsheetId),"CLASS_PAY_VERSION",meta.currentVersion,false);
    updater_appendHistory_(target.spreadsheetId,{updateId:Utilities.getUuid(),startedAt:new Date().toISOString(),finishedAt:new Date().toISOString(),fromVersion:beforeVersion,toVersion:meta.currentVersion,status:"ROLLBACK",backupFolderId:backupFolderId,gasVersion:version.versionNumber,deploymentId:target.deploymentId,message:"コードとDeploymentを復元。データは自動復元していません"});
    return {ok:true,version:meta.currentVersion,gasVersion:version.versionNumber,spreadsheetBackupId:meta.spreadsheetBackupId};
  } finally { lock.releaseLock(); }
}
