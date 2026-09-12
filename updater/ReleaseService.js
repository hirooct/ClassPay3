function updater_fetchJson_(url) {
  var response = UrlFetchApp.fetch(String(url || ""), { muteHttpExceptions: true, followRedirects: true, headers: { Accept: "application/json" } });
  if (response.getResponseCode() !== 200) throw new Error("Release情報を取得できません（HTTP " + response.getResponseCode() + "）");
  return JSON.parse(response.getContentText());
}

function updater_validateManifest_(manifest) {
  if (!manifest || manifest.schemaVersion !== 1) throw new Error("未対応のmanifestです");
  ["latestVersion", "minimumVersion", "releaseDate", "updateType", "files"].forEach(function(key){ if (!manifest[key]) throw new Error("manifestに" + key + "がありません"); });
  if (!["OPTIONAL", "RECOMMENDED", "REQUIRED"].includes(String(manifest.updateType).toUpperCase())) throw new Error("updateTypeが不正です");
  if (!Array.isArray(manifest.files) || !manifest.files.length) throw new Error("更新ファイルがありません");
  var names = {};
  manifest.files.forEach(function(file) {
    if (!file.name || !file.type || !file.url || !/^[a-f0-9]{64}$/i.test(String(file.sha256 || ""))) throw new Error("ファイル定義が不正です: " + (file.name || "unknown"));
    if (names[file.name]) throw new Error("ファイル名が重複しています: " + file.name);
    names[file.name] = true;
  });
  if (!names.appsscript) throw new Error("appsscript manifestがありません");
}

function updater_loadReleaseFiles_(manifest) {
  updater_validateManifest_(manifest);
  return manifest.files.map(function(file) {
    var response = UrlFetchApp.fetch(file.url, { muteHttpExceptions: true, followRedirects: true });
    if (response.getResponseCode() !== 200) throw new Error(file.name + "を取得できません（HTTP " + response.getResponseCode() + "）");
    var source = response.getContentText();
    var actual = updater_sha256_(source);
    if (actual.toLowerCase() !== String(file.sha256).toLowerCase()) throw new Error(file.name + "のSHA-256が一致しません");
    return { name: String(file.name), type: String(file.type), source: source };
  });
}

function updater_verifyProjectFiles_(scriptId, expectedFiles) {
  var actual = updater_getProjectContent_(scriptId).files || [], byName = {};
  actual.forEach(function(file){ byName[file.name] = file; });
  expectedFiles.forEach(function(file) {
    var found = byName[file.name];
    if (!found || found.type !== file.type || updater_sha256_(found.source) !== updater_sha256_(file.source)) throw new Error("更新後コード検証に失敗しました: " + file.name);
  });
  if (actual.length !== expectedFiles.length) throw new Error("更新後のGASファイル数が一致しません");
  return true;
}

function updater_sha256_(text) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(text), Utilities.Charset.UTF_8).map(function(byte) {
    return (byte + 256).toString(16).slice(-2);
  }).join("");
}

function updater_compareSemver_(left, right) {
  function parts(v) { var m = String(v || "0.0.0").replace(/^v/i, "").match(/^(\d+)\.(\d+)\.(\d+)/); return m ? [Number(m[1]),Number(m[2]),Number(m[3])] : [0,0,0]; }
  var a = parts(left), b = parts(right);
  for (var i = 0; i < 3; i++) if (a[i] !== b[i]) return a[i] < b[i] ? -1 : 1;
  return 0;
}
