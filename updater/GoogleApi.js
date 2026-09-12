function updater_googleApi_(method, url, body) {
  var options = {
    method: String(method || "get").toLowerCase(),
    muteHttpExceptions: true,
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() }
  };
  if (body !== undefined) {
    options.contentType = "application/json";
    options.payload = JSON.stringify(body);
  }
  var response = UrlFetchApp.fetch(url, options), code = response.getResponseCode(), text = response.getContentText();
  if (code < 200 || code >= 300) {
    var detail = text;
    try { detail = JSON.parse(text).error.message || text; } catch (ignore) {}
    throw new Error("Google API " + code + ": " + detail);
  }
  return text ? JSON.parse(text) : {};
}

function updater_scriptApi_(method, path, body) {
  return updater_googleApi_(method, "https://script.googleapis.com/v1" + path, body);
}

function updater_getProjectContent_(scriptId, versionNumber) {
  var path = "/projects/" + encodeURIComponent(scriptId) + "/content";
  if (versionNumber) path += "?versionNumber=" + encodeURIComponent(versionNumber);
  return updater_scriptApi_("get", path);
}

function updater_putProjectContent_(scriptId, files) {
  if (!files || !files.length || !files.some(function(f){ return f.name === "appsscript" && f.type === "JSON"; })) {
    throw new Error("更新ファイルにappsscript manifestがありません");
  }
  return updater_scriptApi_("put", "/projects/" + encodeURIComponent(scriptId) + "/content", { files: files });
}

function updater_createVersion_(scriptId, description) {
  return updater_scriptApi_("post", "/projects/" + encodeURIComponent(scriptId) + "/versions", { description: String(description || "ClassPay update") });
}

function updater_updateDeployment_(scriptId, deploymentId, versionNumber, description) {
  return updater_scriptApi_("put", "/projects/" + encodeURIComponent(scriptId) + "/deployments/" + encodeURIComponent(deploymentId), {
    deploymentConfig: { scriptId: scriptId, versionNumber: Number(versionNumber), manifestFileName: "appsscript", description: String(description || "ClassPay") }
  });
}

function updater_createDeployment_(scriptId, versionNumber, description) {
  return updater_scriptApi_("post", "/projects/" + encodeURIComponent(scriptId) + "/deployments", {
    versionNumber: Number(versionNumber), manifestFileName: "appsscript", description: String(description || "ClassPay web app")
  });
}
