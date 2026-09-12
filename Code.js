const SHEETS = {
  USERS: "Users",
  SHOPS: "Shops",
  TX: "Tx",
  CONFIG: "Config",
  COMPANY_MEMBERS: "CompanyMembers",
  COMPANY_APPLICATIONS: "CompanyApplications",
  GOVERNMENT: "Government",
  COMPANY_SNAPSHOTS: "CompanySnapshots",
  GOVERNMENT_LEDGER: "GovernmentLedger",
  HOLDINGS: "Holdings",
  RETIREMENT_APPLICATIONS: "RetirementApplications",
  WEEKLY_REPORTS: "WeeklyReports",
  RULE_PROPOSALS: "RuleProposals",
  RULE_VOTES: "RuleVotes",
  RECRUITMENT_POSTINGS: "RecruitmentPostings",
  EMPLOYMENT_APPLICATIONS: "EmploymentApplications",
  COMPANY_ANNOUNCEMENTS: "CompanyAnnouncements",
  PRODUCT_CATALOG: "ProductCatalog",
  PRODUCT_ORDERS: "ProductOrders",
  COMPANY_CONTRACTS: "CompanyContracts",
  WEEKLY_SETTLEMENTS: "WeeklySettlements",
  UPDATE_HISTORY: "UpdateHistory",
};

function doGet(e) {
  if (e && e.parameter && String(e.parameter.health || "") === "1") {
    return ContentService.createTextOutput(JSON.stringify({
      ok: true,
      app: "ClassPay",
      version: String(getConfig_("CLASS_PAY_VERSION", "3.3.0")),
      spreadsheetReady: !!SpreadsheetApp.getActive().getSheetByName(SHEETS.CONFIG),
      checkedAt: new Date().toISOString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
  const p = (e && e.parameter && e.parameter.p) ? String(e.parameter.p) : "home";
  const allow = { home: 1, shop: 1, pay: 1, admin: 1, balance: 1 };
  // ?p=ui_home はホーム画面への正式な別名として扱う
  const page = (p === "ui_home") ? "home" : (allow[p] ? p : "home");

  const t = HtmlService.createTemplateFromFile("index");
  t.page = page;
  t.params = (e && e.parameter) ? e.parameter : {};

  return t.evaluate()
    .setTitle(getConfig_("APP_NAME", "ClassPay"))
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include_(name) {
  const candidates = [name, name + ".html", name.replace(/\.html$/, "")];
  let lastErr = null;

  for (const file of candidates) {
    try {
      const content = HtmlService.createHtmlOutputFromFile(file).getContent();
      if (content && content.trim()) return content;
    } catch (e) {
      lastErr = e;
    }
  }
  throw new Error("include_ failed: " + candidates.join(", ") + (lastErr ? (" / " + lastErr.message) : ""));
}

function uuid_() {
  return Utilities.getUuid();
}

function lockRun_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(8000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function getConfig_(key, defaultValue) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.CONFIG);
  if (!sh) return defaultValue;

  const v = typeof _cpSheetValues_ === "function"
    ? _cpSheetValues_(sh)
    : sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    if (String(v[i][0]).trim() === key) return v[i][1];
  }
  return defaultValue;
}

function boolConfig_(key, defaultBool) {
  const v = String(getConfig_(key, defaultBool ? "TRUE" : "FALSE")).trim().toUpperCase();
  return (v === "TRUE" || v === "1" || v === "YES");
}

/**
 * HTML側から呼ぶ用
 * Configシートの BASE_URL を返す
 */
function getBaseUrlForClient() {
  let base = String(getConfig_("BASE_URL", "")).trim();
  if (!base) {
    throw new Error("Configシートに BASE_URL がありません。");
  }

  // 末尾を整える
  base = base.replace(/\?p=.*$/, "");
  base = base.replace(/\?$/, "");

  return base;
}
