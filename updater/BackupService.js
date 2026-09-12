function updater_backupRoot_() {
  var found = DriveApp.getFoldersByName(UPDATER_BACKUP_FOLDER);
  return found.hasNext() ? found.next() : DriveApp.createFolder(UPDATER_BACKUP_FOLDER);
}

function updater_createBackup_(target, currentVersion, content, deployment) {
  var stamp = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyyMMdd-HHmmss");
  var folder = updater_backupRoot_().createFolder("ClassPay-" + String(currentVersion).replace(/[^0-9A-Za-z._-]/g, "_") + "-" + stamp);
  var metadata = {
    schemaVersion: 1,
    createdAt: new Date().toISOString(),
    updaterVersion: UPDATER_VERSION,
    target: target,
    currentVersion: currentVersion,
    deployment: deployment || null,
    spreadsheetBackupId: ""
  };
  folder.createFile("source.json", JSON.stringify(content, null, 2), MimeType.PLAIN_TEXT);
  var sheetFile = DriveApp.getFileById(target.spreadsheetId);
  var sheetCopy = sheetFile.makeCopy(sheetFile.getName() + " backup " + stamp, folder);
  metadata.spreadsheetBackupId = sheetCopy.getId();
  folder.createFile("metadata.json", JSON.stringify(metadata, null, 2), MimeType.PLAIN_TEXT);
  return { folderId: folder.getId(), metadata: metadata };
}

function updater_readBackup_(folderId) {
  var folder = DriveApp.getFolderById(folderId), sourceFiles = folder.getFilesByName("source.json"), metaFiles = folder.getFilesByName("metadata.json");
  if (!sourceFiles.hasNext() || !metaFiles.hasNext()) throw new Error("バックアップファイルが不足しています");
  return { folder: folder, source: JSON.parse(sourceFiles.next().getBlob().getDataAsString()), metadata: JSON.parse(metaFiles.next().getBlob().getDataAsString()) };
}

function updater_listBackups_(scriptId, limit) {
  var folders = updater_backupRoot_().getFolders(), rows = [];
  while (folders.hasNext()) {
    var folder = folders.next(), meta = folder.getFilesByName("metadata.json");
    if (!meta.hasNext()) continue;
    try {
      var value = JSON.parse(meta.next().getBlob().getDataAsString());
      if (String(value.target && value.target.scriptId) !== String(scriptId)) continue;
      rows.push({ folderId: folder.getId(), name: folder.getName(), createdAt: value.createdAt, version: value.currentVersion, spreadsheetBackupId: value.spreadsheetBackupId });
    } catch (ignore) {}
  }
  return rows.sort(function(a,b){ return String(b.createdAt).localeCompare(String(a.createdAt)); }).slice(0, Number(limit || 20));
}
