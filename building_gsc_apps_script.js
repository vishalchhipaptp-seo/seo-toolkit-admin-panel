// ============================================
// SEO Toolkit Pro - GSC Crawl Tracker (BUILDING edition)
// One self-contained Google Apps Script per BUILDING. Bind it to that building's
// GSC Google Sheet (Extensions > Apps Script), or deploy standalone with SHEET_ID.
// It manages that building's GSC accounts, OAuth tokens, domain sync and crawl.
// Login/roles live in the CENTRAL gateway, NOT here.
//
// Google Sheet tabs used: GSC_Config, GSC_Accounts, Auth_Status, GSC_Settings,
//   CrawlHistory, GSC_Help, Log_YYYY_MM
//
// Script Properties (Project Settings > Script Properties):
//   SHEET_ID            = this building's GSC sheet ID
//   CLIENT_ID           = GSC OAuth client id  (Google Cloud project)
//   CLIENT_SECRET       = GSC OAuth client secret
//   FALLBACK_ALERT_EMAIL (optional) = where alerts go if "Alert Email TO" in
//                          GSC_Settings is left blank. Rarely needed - if
//                          left unset too, alerts automatically go to this
//                          building's own admin_emails (set centrally by the
//                          super admin, in the Buildings tab), looked up
//                          automatically. Only set this if you want alerts
//                          to go somewhere OTHER than the registered admins.
//   (ADMIN_KEY is no longer used - every admin action checks the calling
//   admin's own SEO Toolkit Pro login instead. Don't add it.)
//
// Deploy: Deploy > New deployment > Web app > Execute as: Me, Access: Anyone.
// Put the /exec URL into the CENTRAL sheet Buildings tab (gsc_script_url) for this building.
//
// OAuth2 Library ID (Libraries > add): 1B7FSrk5Zi6L1rSxxTDgDEUsPzlukDsi4KGuTMorsTQHhGBzBkMun4iDF
// ============================================

function getProp(name) { return PropertiesService.getScriptProperties().getProperty(name); }
function getSheetId()      { return getProp("SHEET_ID"); }
function getClientId()     { return getProp("CLIENT_ID"); }
function getClientSecret() { return getProp("CLIENT_SECRET"); }
function getAdminKey()     { return getProp("ADMIN_KEY"); }

// The one central gateway every building reports to - same for all
// buildings, so it's a constant here rather than something each building
// admin has to enter. Used only as a last-resort alert-routing fallback:
// this script identifies itself by its OWN deployed URL (which only a
// script actually deployed as a real, already-registered building would
// know), so no separate secret is needed for this specific lookup.
var CENTRAL_GATEWAY_URL = "https://script.google.com/macros/s/AKfycbyYhBdjOinGEFgN_unzHXlSuhQWpqdGoipN4dB1iVTRxrqoq_5c_grAi1LAjCG80VTwmw/exec";

function _fallbackAdminEmails() {
  try {
    var selfUrl = ScriptApp.getService().getUrl();
    if (!selfUrl) return "";
    var resp = UrlFetchApp.fetch(
      CENTRAL_GATEWAY_URL + "?action=building_admins_get&caller_url=" + encodeURIComponent(selfUrl),
      { muteHttpExceptions: true }
    );
    var data = JSON.parse(resp.getContentText());
    return data.success ? (data.admin_emails || "") : "";
  } catch (e) {
    return "";
  }
}

const CACHE = CacheService.getScriptCache();

// ============================================
// Web App entry points - GSC / Crawl actions only
// ============================================
function doGet(e) {
  try {
    var params = (e && e.parameter) ? e.parameter : {};
    var action = (params.action || "").toString().trim();
    if (action === "get_config")        return jsonOut(getConfigForExtension());
    if (action === "get_settings")      return handleGetSettings();
    if (action === "get_crawl_history") return handleGetCrawlHistory(params);
    if (action === "ping")              return jsonOut({ status: "GSC Backend Active", timestamp: new Date().toISOString() });
    return jsonOut({ status: "GSC Backend Active" });
  } catch (err) { return jsonOut({ success: false, error: err.message }); }
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var action = String(data.action || "").trim();
    switch (action) {
      case "send_otp":         return sendOtp(data);
      case "send_alert":       return jsonOut(handleSendAlert(data));
      case "save_settings":    return jsonOut(handleSaveSettings(data.settings));
      case "validate_batch":   return jsonOut(handleValidateBatch(data));
      case "inspect_single":   return jsonOut(handleInspectSingle(data));
      case "save_crawl_batch": return jsonOut(handleSaveCrawlBatch(data));
      case "get_keys":         return getKeys();
      default: return jsonOut({ success: false, error: "Unknown action: " + action });
    }
  } catch (err) { return jsonOut({ success: false, error: err.message }); }
}

// ============================================
// Shared helpers used by the GSC code below
// ============================================

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function _getAuthSheet(name) {
  var ss = SpreadsheetApp.openById(getSheetId());
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (name === 'users') {
      sheet.appendRow(['Email', 'Password', 'Device_ID', 'Approved', 'Name', 'Registered', 'Last Login', 'Notes', 'is_admin', 'formats', 'tools']);
    } else if (name === 'config') {
      sheet.appendRow(['min_version', '3.2']);
    }
  }
  return sheet;
}


// Everything below is the original GSC script
// (unchanged except doGet/doPost moved above)
// ============================================

// Batch progress tracking
function getBatchProgress(batchId) {
  const raw = CACHE.get('batch_' + batchId);
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (e) { return null; }
}

function setBatchProgress(batchId, data) {
  CACHE.put('batch_' + batchId, JSON.stringify(data), 21600);
}

function incrementBatchProcessed(batchId) {
  const progress = getBatchProgress(batchId);
  if (!progress) return null;
  progress.processed = (progress.processed || 0) + 1;
  setBatchProgress(batchId, progress);
  return progress;
}

// Settings tab helpers
function getSettings() {
  const ss = SpreadsheetApp.openById(getSheetId());
  let settingsSheet = ss.getSheetByName('GSC_Settings');
  if (!settingsSheet) settingsSheet = createSettingsTab();
  const data = settingsSheet.getDataRange().getValues();
  const settings = {};
  for (let i = 1; i < data.length; i++) {
    const key   = data[i][0] ? data[i][0].toString().trim() : '';
    const value = data[i][1] !== undefined && data[i][1] !== null ? data[i][1].toString().trim() : '';
    if (key) settings[key] = value;
  }
  return settings;
}

function getNotificationEmails() {
  const settings = getSettings();
  const raw = settings['Notification Emails'] || '';
  return raw.split(',').map(e => e.trim()).filter(e => e.length > 0 && e.indexOf('@') !== -1);
}

function isWeeklySummaryEnabled() {
  const settings = getSettings();
  const val = (settings['Send Weekly Summary'] || 'Yes').toLowerCase();
  return val === 'yes' || val === 'true' || val === '1';
}

function createSettingsTab() {
  const ss = SpreadsheetApp.openById(getSheetId());
  let settingsSheet = ss.getSheetByName('GSC_Settings');
  if (settingsSheet) return settingsSheet;
  settingsSheet = ss.insertSheet('GSC_Settings');
  const rows = [
    ['Setting', 'Value', 'Description'],
    ['Notification Emails', '', 'Comma-separated emails who receive the weekly sync summary.'],
    ['Send Weekly Summary', 'Yes', 'Yes or No. If No, weekly sync still runs but no email is sent.'],
  ];
  settingsSheet.getRange(1, 1, rows.length, 3).setValues(rows);
  settingsSheet.getRange('A1:C1').setFontWeight('bold').setBackground('#1e1e2e').setFontColor('#00ff88');
  settingsSheet.setColumnWidth(1, 200); settingsSheet.setColumnWidth(2, 400); settingsSheet.setColumnWidth(3, 500);
  settingsSheet.setFrozenRows(1);
  return settingsSheet;
}

function handleGetSettings() {
  var settings = getSettings();
  return jsonOut({
    success: true,
    settings: {
      notificationEmails: settings['Notification Emails'] || '',
      sendWeeklySummary: settings['Send Weekly Summary'] || 'Yes',
      alertEmailTo: settings['Alert Email TO'] || '',
      alertEmailCc: settings['Alert Email CC'] || ''
    }
  });
}

function handleSaveSettings(settings) {
  if (!settings) return { success: false, error: 'No settings provided' };
  var ss = SpreadsheetApp.openById(getSheetId());
  var sheet = ss.getSheetByName('GSC_Settings');
  if (!sheet) sheet = createSettingsTab();
  var data = sheet.getDataRange().getValues();
  var keyMap = {
    'Notification Emails': 'notificationEmails',
    'Send Weekly Summary': 'sendWeeklySummary',
    'Alert Email TO': 'alertEmailTo',
    'Alert Email CC': 'alertEmailCc'
  };
  for (var i = 1; i < data.length; i++) {
    var settingName = String(data[i][0]).trim();
    var jsKey = keyMap[settingName];
    if (jsKey && settings[jsKey] !== undefined) sheet.getRange(i + 1, 2).setValue(settings[jsKey]);
  }
  return { success: true };
}

// OAuth2 service per account
// hasAccess() reads PropertiesService and can involve a live token check -
// handleValidateBatch was calling this fresh for EVERY url in a batch, so a
// 5-url batch against the same domain/account did the same OAuth access check
// 5 times. Cache the boolean result per accountKey (short TTL, and a shorter
// TTL on a "not authorised" result so a freshly-fixed auth issue is picked up
// quickly) - this is server-side and per-building (this script's own
// CacheService), never exposed to the client beyond the per-URL ready/error
// result it already returned.
function hasAccessCached(accountKey) {
  const cacheKey = 'oauth_access_' + accountKey;
  const cached = CACHE.get(cacheKey);
  if (cached !== null) return cached === '1';
  let ok = getOAuthService(accountKey).hasAccess();
  // A single flaky/transient hasAccess() check (network blip, momentary quota
  // hiccup) would otherwise get cached as "not authorised" for 30s and then
  // poison EVERY other URL for the same account checked in that window - for
  // a validate_batch call processing many URLs under the same account key
  // back-to-back, that turns one bad check into the whole batch reading
  // "Account not authorised" even though the account genuinely has access
  // (confirmed real case: costumesdelivered.com.au showed Full (domain)
  // access with today's sync in GSC_Config, yet every URL in that batch was
  // skipped). Retry once before accepting a negative result.
  if (!ok) {
    Utilities.sleep(500);
    ok = getOAuthService(accountKey).hasAccess();
  }
  CACHE.put(cacheKey, ok ? '1' : '0', ok ? 300 : 30);
  return ok;
}

function getOAuthService(accountKey) {
  let clientId = getClientId();
  let clientSecret = getClientSecret();
  try {
    const acct = getAccountsMap()[accountKey];
    if (acct && acct.project) {
      const creds = getProjectCreds(acct.project);
      if (creds && creds.clientId && creds.clientSecret) {
        clientId = creds.clientId; clientSecret = creds.clientSecret;
      }
    }
  } catch (e) {
    Logger.log('Project lookup failed for ' + accountKey + ': ' + e.message);
  }
  return OAuth2.createService(accountKey)
    .setAuthorizationBaseUrl('https://accounts.google.com/o/oauth2/auth')
    .setTokenUrl('https://oauth2.googleapis.com/token')
    .setClientId(clientId).setClientSecret(clientSecret)
    .setCallbackFunction('authCallback')
    .setPropertyStore(PropertiesService.getScriptProperties())
    .setScope(['https://www.googleapis.com/auth/webmasters.readonly'])
    .setParam('access_type', 'offline').setParam('prompt', 'consent')
    .setParam('login_hint', getEmailForKey(accountKey));
}

// Lookups
function getEmailForKey(accountKey) {
  try { const acct = getAccountsMap()[accountKey]; if (acct && acct.email) return acct.email; } catch (e) {}
  const ss = SpreadsheetApp.openById(getSheetId());
  const config = ss.getSheetByName('GSC_Config');
  const data = config.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][2] && data[i][2].toString().trim() === accountKey) return data[i][1].toString().trim();
  }
  return null;
}

// Was reading + scanning the whole GSC_Config sheet fresh on EVERY call.
// handleValidateBatch calls this once per URL in the batch (up to 100 per
// request), so a 100-URL batch meant up to 100 fresh SpreadsheetApp reads --
// the slowest kind of Apps Script operation, and it only gets worse as this
// building's master domain list grows. Cache the whole domain->accountKey
// map instead, same pattern as getAccountsMap(), so a batch does ONE sheet
// read total.
function getDomainAccountMap() {
  const cached = CACHE.get('domain_account_map_v1');
  if (cached) { try { return JSON.parse(cached); } catch (e) {} }
  const ss = SpreadsheetApp.openById(getSheetId());
  const config = ss.getSheetByName('GSC_Config');
  const data = config.getDataRange().getValues();
  // domain -> [accountKey, ...] - a domain can be reachable through MORE THAN
  // ONE connected GSC account (dedup'd, in sheet order), not just whichever
  // row happened to be seen last.
  const map = {};
  for (let i = 1; i < data.length; i++) {
    const domain = data[i][0] ? data[i][0].toString().trim().toLowerCase() : '';
    const key = data[i][2] ? data[i][2].toString().trim() : '';
    const access = data[i][3] ? data[i][3].toString().trim() : '';
    if (!domain || !key) continue;
    if (access.indexOf('Removed') === 0) continue;  // no live access - not a usable candidate
    if (!map[domain]) map[domain] = [];
    if (map[domain].indexOf(key) === -1) map[domain].push(key);
  }
  try { CACHE.put('domain_account_map_v1', JSON.stringify(map), 300); } catch (e) {
    // Master list too large for one cache entry (100KB CacheService limit) --
    // fall back silently to per-call reads rather than breaking the lookup.
    Logger.log('domain_account_map cache put failed: ' + e.message);
  }
  return map;
}

function getAccountKeysForDomain(domain) {
  const target = domain.toString().trim().toLowerCase();
  const map = getDomainAccountMap();
  return Object.prototype.hasOwnProperty.call(map, target) ? map[target] : [];
}

// Tries every account with access to this domain, one after another, and
// returns the first one that actually has LIVE access right now - previously
// only a single (last-seen-in-sheet) account key was ever tried, so if THAT
// one account happened to be quota-exhausted or unauthorised, the domain
// read as failed even when another connected account with access to the
// same domain was working fine the whole time.
function getAccountKeyForDomain(domain) {
  const candidates = getAccountKeysForDomain(domain);
  for (let i = 0; i < candidates.length; i++) {
    if (hasAccessCached(candidates[i])) return candidates[i];
  }
  // None currently have live access - fall back to the first candidate so
  // the caller still gets a real, specific reason (e.g. "Account not
  // authorised") instead of a generic "domain not in master list" for a
  // domain that DOES exist, just with no currently-working account.
  return candidates.length > 0 ? candidates[0] : null;
}

// Call after editing GSC_Config (add/remove/change a domain mapping) so the
// next batch picks up the change immediately instead of waiting out the
// 5-minute cache TTL.
function clearDomainAccountMapCache() {
  CACHE.remove('domain_account_map_v1');
}

function extractDomainFromUrl(url) {
  try {
    const match = url.match(/^https?:\/\/([^\/]+)/i);
    if (!match) return null;
    return match[1].toLowerCase().replace(/^www\./, '');
  } catch (e) { return null; }
}

// Fetches + filters the account's GSC property list, cached per accountKey for
// 30 min. resolveSiteProperty() used to call sites.list() fresh on EVERY
// domain row - for an account managing many client domains under one GSC
// account, that meant redundant, identical sites.list() calls per sweep
// instead of one - very likely why a large sheet's version-check sweep only
// gets through a handful of domains per 3-minute time-slice.
function _getAccountSiteEntries(accountKey, accessToken) {
  var cacheKey = accountKey ? ('site_entries_' + accountKey) : null;
  if (cacheKey) {
    var cached = CACHE.get(cacheKey);
    if (cached) { try { return JSON.parse(cached); } catch (e) {} }
  }
  const response = UrlFetchApp.fetch('https://www.googleapis.com/webmasters/v3/sites', {
    method: 'GET', headers: { 'Authorization': 'Bearer ' + accessToken }, muteHttpExceptions: true
  });
  if (response.getResponseCode() !== 200) return null;
  const allEntries = (JSON.parse(response.getContentText()).siteEntry) || [];
  // Exclude properties this account can see but doesn't actually own/verify -
  // otherwise a pending "sc-domain:" entry can shadow a real, owned URL-prefix property.
  const entries = allEntries.filter((e) => e.permissionLevel !== 'siteUnverifiedUser');
  if (cacheKey) {
    try { CACHE.put(cacheKey, JSON.stringify(entries), 1800); } catch (e) {
      // Account has enough properties to exceed the 100KB CacheService limit -
      // fall back silently to per-call fetches rather than breaking resolution.
    }
  }
  return entries;
}

// GSC property resolution. accountKey is optional - when given, the account's
// property list is cached (see _getAccountSiteEntries above); omit it only for
// one-off/manual calls where caching doesn't matter.
function resolveSiteProperty(domain, accessToken, sampleUrl, accountKey) {
  const entries = _getAccountSiteEntries(accountKey, accessToken);
  if (entries === null) return { error: 'Could not list GSC properties' };
  const bareHost = (s) => s.toLowerCase().replace(/^sc-domain:/, '').replace(/^https?:\/\//, '').replace(/^www\./, '').replace(/\/.*$/, '');
  const target = bareHost(domain);
  // A sc-domain: property only actually COVERS the site if the live host still
  // matches the configured domain name. Previously this matched on the
  // property's NAME alone - if the site had been redirected off-domain
  // entirely (e.g. perfume-samples.co.uk -> refachemical.com), a
  // sc-domain:perfume-samples.co.uk property still "matched" by name and got
  // reported as Correct/covers-it, which is false: a domain property never
  // covers a genuinely different domain.
  const runningHost = sampleUrl ? bareHost(sampleUrl) : null;
  for (let i = 0; i < entries.length; i++) {
    if (entries[i].siteUrl.indexOf('sc-domain:') === 0 && bareHost(entries[i].siteUrl) === target) {
      if (runningHost && runningHost !== target) {
        return { error: 'Live site now runs on a different domain (' + sampleUrl + ') than the ' +
                         'configured GSC domain property (' + domain + ') - that domain property ' +
                         'does not cover it.', movedOffDomain: true };
      }
      return { siteUrl: entries[i].siteUrl, type: 'domain' };
    }
  }
  const hostPrefixes = [];
  const sample = sampleUrl ? sampleUrl.toLowerCase() : '';
  for (let i = 0; i < entries.length; i++) {
    const su = entries[i].siteUrl;
    if (su.indexOf('sc-domain:') !== 0 && bareHost(su) === target) {
      hostPrefixes.push(su);
      if (sample && sample.indexOf(su.toLowerCase()) === 0) return { siteUrl: su, type: 'prefix' };
    }
  }
  if (hostPrefixes.length > 0) return { mismatch: true, available: hostPrefixes };
  return { notFound: true };
}

function getCachedSiteProperty(domain, accountKey, accessToken, sampleUrl) {
  const origin = (sampleUrl.match(/^https?:\/\/[^\/]+/i) || [''])[0].toLowerCase();
  const cacheKey = 'prop_' + accountKey + '_' + origin;
  const cached = CACHE.get(cacheKey);
  if (cached) { try { return JSON.parse(cached); } catch (e) {} }
  const prop = resolveSiteProperty(domain, accessToken, sampleUrl, accountKey);
  if (prop.siteUrl) CACHE.put(cacheKey, JSON.stringify(prop), 1800);
  return prop;
}

// ONE-OFF: clears the cached (possibly stale/wrong) resolved property for a
// single accountKey + URL origin. Run manually from the Apps Script editor.
function clearPropCacheForTest() {
  const accountKey = 'searchconsoleanalyticsaccess8';
  const sampleUrl = 'https://aerokitchen.com/kitchens/';
  const origin = (sampleUrl.match(/^https?:\/\/[^\/]+/i) || [''])[0].toLowerCase();
  const cacheKey = 'prop_' + accountKey + '_' + origin;
  CACHE.remove(cacheKey);
  Logger.log('Cleared cache key: ' + cacheKey);
}

function inspectUrl(inspectionUrl, siteUrl, accessToken) {
  const response = UrlFetchApp.fetch(
    'https://searchconsole.googleapis.com/v1/urlInspection/index:inspect',
    {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + accessToken, 'Content-Type': 'application/json' },
      payload: JSON.stringify({ inspectionUrl: inspectionUrl, siteUrl: siteUrl }),
      muteHttpExceptions: true
    }
  );
  let body = {};
  try { body = JSON.parse(response.getContentText() || '{}'); } catch (e) {}
  return { code: response.getResponseCode(), body: body };
}

// Monthly audit log
function getOrCreateLogSheet() {
  const ss = SpreadsheetApp.openById(getSheetId());
  const monthTag = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy_MM');
  const tabName = 'Log_' + monthTag;
  let logSheet = ss.getSheetByName(tabName);
  if (!logSheet) {
    logSheet = ss.insertSheet(tabName);
    const headers = ['Timestamp (IST)', 'Domain(s)', 'Account Key', 'URL Count', 'Status', 'Processed', 'Error / Notes'];
    logSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    logSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#1e1e2e').setFontColor('#00ff88');
    logSheet.setFrozenRows(1);
  }
  return logSheet;
}

function logRequest(entry) {
  try {
    const logSheet = getOrCreateLogSheet();
    const timestamp = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'dd MMM yyyy, HH:mm:ss');
    logSheet.appendRow([timestamp, entry.domain || '', entry.accountKey || '', entry.urlCount || 0, entry.status || '', entry.processed || 0, entry.notes || '']);
    return logSheet.getLastRow();
  } catch (err) { Logger.log('Log write failed: ' + err.message); return null; }
}

function updateLogProcessed(rowNumber, processedCount) {
  if (!rowNumber) return;
  try { getOrCreateLogSheet().getRange(rowNumber, 6).setValue(processedCount); } catch (err) {}
}

// OAuth callback
function authCallback(request) {
  const accountKey = request.parameter.accountKey;
  const service = getOAuthService(accountKey);
  const authorized = service.handleCallback(request);
  return HtmlService.createHtmlOutput(
    authorized ? 'Account authorised: ' + accountKey + '. You can close this tab.'
               : 'Authorisation failed for: ' + accountKey
  );
}

// Bulk auth helpers
function generateAuthSheet() {
  const ss = SpreadsheetApp.openById(getSheetId());
  const config = ss.getSheetByName('GSC_Config');
  const data = config.getDataRange().getValues();
  const seen = {}; const accounts = [];
  for (let i = 1; i < data.length; i++) {
    const email = data[i][1] ? data[i][1].toString().trim() : '';
    const key   = data[i][2] ? data[i][2].toString().trim() : '';
    if (!email || !key || seen[key]) continue;
    seen[key] = true; accounts.push({ email: email, key: key });
  }
  let authSheet = ss.getSheetByName('Auth_Status');
  if (authSheet) { authSheet.clear(); } else { authSheet = ss.insertSheet('Auth_Status'); }
  authSheet.getRange(1, 1, 1, 4).setValues([['Email', 'Account Key', 'Auth URL', 'Status']]);
  authSheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#1e1e2e').setFontColor('#00ff88');
  const rows = [];
  accounts.forEach(a => {
    const service = getOAuthService(a.key);
    if (service.hasAccess()) { rows.push([a.email, a.key, '', 'Authorised']); }
    else { rows.push([a.email, a.key, service.getAuthorizationUrl({ accountKey: a.key }), 'Pending']); }
  });
  authSheet.getRange(2, 1, rows.length, 4).setValues(rows);
  authSheet.setFrozenRows(1);
}

// Sync GSC_Config
function dedupeGscConfig() {
  const ss = SpreadsheetApp.openById(getSheetId());
  const config = ss.getSheetByName('GSC_Config');
  if (!config) return 0;
  const data = config.getDataRange().getValues();
  if (data.length < 2) return 0;
  const parseDate = (v) => { if (v instanceof Date) return v.getTime(); const t = Date.parse(v); return isNaN(t) ? 0 : t; };
  const winner = {};       // dedup key -> {idx, dateVal} of the row to KEEP
  const passthrough = [];  // rows with no domain/key - never deduped, always kept
  for (let i = 1; i < data.length; i++) {
    const domain = data[i][0] ? data[i][0].toString().trim().toLowerCase() : '';
    const key = data[i][2] ? data[i][2].toString().trim() : '';
    if (!domain || !key || domain === '_setup') { passthrough.push(i); continue; }
    const prop = data[i][5] ? data[i][5].toString().trim().toLowerCase() : '';
    const k = (prop || domain) + '||' + key;
    const dateVal = parseDate(data[i][4]);
    if (!winner[k] || dateVal >= winner[k].dateVal) winner[k] = { idx: i, dateVal: dateVal };
  }
  const keepIdx = passthrough.concat(Object.keys(winner).map((k) => winner[k].idx)).sort((a, b) => a - b);
  const removedCount = (data.length - 1) - keepIdx.length;
  if (removedCount <= 0) return 0;

  // Rewrite with ONE bulk write + ONE trim instead of one deleteRow() call
  // per duplicate - deleteRow() reshuffles every subsequent row on every
  // call, and doing that once per duplicate (found via real Execution log
  // data: syncDomainsFromGSC was timing out at ~360s on essentially every
  // run for the past week) was almost certainly the dominant cost, running
  // unconditionally before the account-sync loop even starts.
  const kept = keepIdx.map((i) => data[i]);
  if (kept.length > 0) config.getRange(2, 1, kept.length, data[0].length).setValues(kept);
  const newLastRow = 1 + kept.length;
  if (config.getLastRow() > newLastRow) config.deleteRows(newLastRow + 1, config.getLastRow() - newLastRow);
  return removedCount;
}

function syncDomainsFromGSC() {
  const ss = SpreadsheetApp.openById(getSheetId());
  try { syncAccountsRegistry(); } catch (e) { Logger.log('Registry sync skipped: ' + e.message); }
  try { dedupeGscConfig(); } catch (e) { Logger.log('Dedupe skipped: ' + e.message); }
  const runDateShort = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'dd MMM yyyy');
  const runDateLong  = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'dd MMM yyyy, HH:mm');
  const authSheet = ss.getSheetByName('Auth_Status');
  if (!authSheet) { Logger.log('Auth_Status tab missing.'); return; }
  // getRange() throws "number of rows must be at least 1" if the sheet has
  // only a header row (getLastRow() === 1) - a real crash risk whenever
  // Auth_Status is briefly empty (freshly created, or between a clear and a
  // re-populate), and one of the few uncaught exceptions in this function
  // that would show up as a failed trigger execution.
  if (authSheet.getLastRow() < 2) { Logger.log('Auth_Status has no account rows yet.'); return; }
  const authData = authSheet.getRange(2, 1, authSheet.getLastRow() - 1, 4).getValues();
  const accounts = [];
  authData.forEach(row => {
    const email = row[0] ? row[0].toString().trim() : '';
    const key = row[1] ? row[1].toString().trim() : '';
    const status = row[3] ? row[3].toString().trim() : '';
    if (email && key && status === 'Authorised') accounts.push({ email, key });
  });
  if (accounts.length === 0) { Logger.log('No authorised accounts.'); return; }

  const config = ss.getSheetByName('GSC_Config');
  if (!config) { Logger.log('GSC_Config missing!'); return; }
  config.getRange(1, 1, 1, 6).setValues([['Domain', 'GSC Email', 'Account Key', 'Access Level', 'Last Synced', 'GSC Property']]);

  const existingData = config.getDataRange().getValues();
  const byProp = {}; const byDomain = {};
  const fmtDate_ = (v) => v ? (v instanceof Date ? Utilities.formatDate(v, 'Asia/Kolkata', 'dd MMM yyyy') : v.toString().trim()) : '';
  for (let i = 1; i < existingData.length; i++) {
    const domain = existingData[i][0] ? existingData[i][0].toString().trim().toLowerCase() : '';
    const key = existingData[i][2] ? existingData[i][2].toString().trim() : '';
    const prop = existingData[i][5] ? existingData[i][5].toString().trim() : '';
    if (!key) continue;
    const rec = { row: i + 1, previousLevel: existingData[i][3] ? existingData[i][3].toString().trim() : '', previousDate: fmtDate_(existingData[i][4]), previousProperty: prop, adopted: false };
    if (prop) byProp[prop.toLowerCase() + '||' + key] = rec;
    else if (domain) { (byDomain[domain + '||' + key] = byDomain[domain + '||' + key] || []).push(rec); }
  }

  const bareHost = (s) => s.toLowerCase().replace(/^sc-domain:/, '').replace(/^https?:\/\//, '').replace(/^www\./, '').replace(/\/.*$/, '');
  const labelFor = (pl, isDom) => { const nice = ({siteOwner:'Owner',siteFullUser:'Full',siteRestrictedUser:'Restricted',siteUnverifiedUser:'Unverified'})[pl] || pl; return nice + (isDom ? ' (domain)' : ' (URL prefix)'); };

  const newRows = []; const updates = []; const seenRows = {};
  const accountsChecked = {}; const accountsFailed = [];
  const changes = { newDomains: [], accessChanged: [], removed: [], accountsWithErrors: [] };
  let totalProperties = 0;

  for (let a = 0; a < accounts.length; a++) {
    const account = accounts[a];
    const service = getOAuthService(account.key);
    if (!service.hasAccess()) { accountsFailed.push(account.key + ' (no token)'); continue; }
    let response;
    try { response = UrlFetchApp.fetch('https://www.googleapis.com/webmasters/v3/sites', { headers: { 'Authorization': 'Bearer ' + service.getAccessToken() }, muteHttpExceptions: true }); }
    catch (e) { accountsFailed.push(account.key + ' (fetch error)'); continue; }
    if (response.getResponseCode() !== 200) { accountsFailed.push(account.key + ' (HTTP ' + response.getResponseCode() + ')'); continue; }
    const entries = (JSON.parse(response.getContentText()).siteEntry) || [];
    accountsChecked[account.key] = true;

    entries.forEach(entry => {
      const siteUrl = entry.siteUrl; const domain = bareHost(siteUrl);
      const isDomainProperty = siteUrl.indexOf('sc-domain:') === 0;
      if (entry.permissionLevel === 'siteUnverifiedUser') return;
      const accessLevel = labelFor(entry.permissionLevel, isDomainProperty);
      totalProperties++;
      let rec = byProp[siteUrl.toLowerCase() + '||' + account.key];
      if (!rec) { const cands = byDomain[domain + '||' + account.key]; if (cands) { for (let j = 0; j < cands.length; j++) { if (!cands[j].adopted) { rec = cands[j]; rec.adopted = true; break; } } } }
      if (rec) {
        seenRows[rec.row] = true;
        updates.push({ row: rec.row, level: accessLevel, property: siteUrl });
        if (rec.previousLevel && rec.previousLevel !== accessLevel && rec.previousLevel.indexOf('Removed') === -1)
          changes.accessChanged.push({ domain, key: account.key, before: rec.previousLevel, beforeDate: rec.previousDate, beforeProperty: rec.previousProperty, after: accessLevel, afterProperty: siteUrl, afterDate: runDateShort });
      } else {
        newRows.push([domain, account.email, account.key, accessLevel, runDateShort, siteUrl]);
        changes.newDomains.push({ domain, email: account.email, key: account.key, level: accessLevel, property: siteUrl });
      }
    });
    Utilities.sleep(100);
  }

  // Both the per-row update loop AND the "mark removed" loop below used to
  // issue one Range write per row (the update loop: one 3-column setValues()
  // per row; the removed loop: TWO single-cell setValue() calls per row) -
  // each is a real API round-trip, and with hundreds of accounts/domains this
  // was the exact same "N individual writes" anti-pattern that was already
  // found and fixed in dedupeGscConfig() above (see its comment - that one
  // alone was confirmed via Execution log data to cause ~360s timeouts on
  // almost every run). This loop is the sibling of that bug, just not yet
  // fixed - both are now folded into ONE bulk write: mutate the already-
  // in-memory existingData array for every update/removal, then write the
  // whole range back in a single setValues() call.
  updates.forEach(u => { existingData[u.row - 1][3] = u.level; existingData[u.row - 1][4] = runDateShort; existingData[u.row - 1][5] = u.property; });

  let removedCount = 0;
  for (let i = 1; i < existingData.length; i++) {
    const domain = existingData[i][0] ? existingData[i][0].toString().trim().toLowerCase() : '';
    const key = existingData[i][2] ? existingData[i][2].toString().trim() : '';
    if (!domain || !key || domain === '_setup') continue;
    if (!accountsChecked[key]) continue;
    if (!seenRows[i + 1]) {
      const currentValue = existingData[i][3] ? existingData[i][3].toString() : '';
      if (currentValue.indexOf('Removed') === -1) { existingData[i][3] = 'Removed (' + runDateShort + ')'; existingData[i][4] = runDateShort; removedCount++; changes.removed.push({ domain, key, lastLevel: currentValue, lastProperty: existingData[i][5] ? existingData[i][5].toString().trim() : '' }); }
    }
  }
  if (existingData.length > 1) config.getRange(2, 1, existingData.length - 1, 6).setValues(existingData.slice(1));
  if (newRows.length > 0) config.getRange(config.getLastRow() + 1, 1, newRows.length, 6).setValues(newRows);
  changes.accountsWithErrors = accountsFailed;
  if (config.getLastRow() > 1) config.getRange(2, 1, config.getLastRow() - 1, 6).sort([{ column: 3, ascending: true }, { column: 1, ascending: true }]);

  if (isWeeklySummaryEnabled()) sendRoutedSyncEmails({ runTimestamp: runDateLong, runDateShort, changes });

  // The automatic Version_Check_Status re-check on every new/changed domain
  // was DISABLED here (not deleted - checkVersionForChangedDomains/
  // _runVersionCheckSweep still exist and work, just aren't auto-triggered
  // any more) because it draws from the SAME account-wide daily UrlFetchApp
  // quota as everything else on this script, including live/user-triggered
  // Crawl Tracker checks - confirmed real case: Crawl Tracker started
  // failing with "Service invoked too many times for one day: urlfetch"
  // while this was running in the background. Live tools should always get
  // quota priority. Re-enable by uncommenting below, or call
  // checkVersionForChangedDomains(...)/_runVersionCheckSweep() manually
  // (e.g. from a menu item) when quota headroom isn't a concern.
  //
  // var toVersionCheck = changes.newDomains.map(function (d) { return { domain: d.domain, key: d.key }; })
  //   .concat(changes.accessChanged.map(function (x) { return { domain: x.domain, key: x.key }; }));
  // if (toVersionCheck.length > 0) {
  //   try { checkVersionForChangedDomains(toVersionCheck); } catch (e) { Logger.log('checkVersionForChangedDomains failed: ' + e.message); }
  // }
}

// Get config for extension
function getConfigForExtension() {
  try {
    const ss = SpreadsheetApp.openById(getSheetId());
    const config = ss.getSheetByName('GSC_Config');
    if (!config) return { success: false, error: 'GSC_Config tab not found' };
    const data = config.getDataRange().getValues();
    if (data.length < 2) return { success: true, mapping: {}, domains: [] };
    const accessRank = { 'Full (domain)': 100, 'Full (URL prefix)': 90, 'Owner (domain)': 80, 'Owner (URL prefix)': 70, 'Restricted (domain)': 60, 'Restricted (URL prefix)': 50 };
    const best = {}; const allAccess = {};
    for (let i = 1; i < data.length; i++) {
      const domain = data[i][0] ? data[i][0].toString().trim().toLowerCase() : '';
      const email = data[i][1] ? data[i][1].toString().trim().toLowerCase() : '';
      const accountKey = data[i][2] ? data[i][2].toString().trim() : '';
      const accessLevel = data[i][3] ? data[i][3].toString().trim() : '';
      if (!domain || !email || domain === '_setup') continue;
      if (accessLevel.indexOf('Removed') !== -1 || accessLevel.indexOf('Unverified') !== -1) continue;
      const rank = accessRank[accessLevel] || 0;
      if (!best[domain] || rank > best[domain].rank) best[domain] = { email, accountKey, accessLevel, rank };
      if (!allAccess[domain]) allAccess[domain] = [];
      allAccess[domain].push({ email, accountKey, accessLevel, rank });
    }
    const mapping = {};
    Object.keys(best).forEach(d => {
      const b = best[d]; const others = (allAccess[d] || []).filter(a => a.accountKey !== b.accountKey);
      const isRestricted = b.accessLevel.indexOf('Restricted') !== -1;
      const betterOptions = others.filter(a => a.rank > b.rank);
      mapping[d] = { email: b.email, accountKey: b.accountKey, accessLevel: b.accessLevel, isRestrictedOnly: isRestricted && betterOptions.length === 0,
        betterAccess: betterOptions.length > 0 ? betterOptions.map(a => ({ email: a.email, accountKey: a.accountKey, accessLevel: a.accessLevel })) : null };
    });
    return { success: true, mapping, domains: Object.keys(mapping).sort(), total: Object.keys(mapping).length, generated: new Date().toISOString() };
  } catch (err) { return { success: false, error: err.message }; }
}

// Keys
function getKeys() {
  const sheet = SpreadsheetApp.openById(getSheetId()).getSheetByName('Keys');
  if (!sheet) return jsonOut({ success: false, error: 'Sheet "Keys" not found.' });
  const values = sheet.getDataRange().getValues();
  const keys = {};
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] && values[i][1]) keys[String(values[i][0]).trim()] = String(values[i][1]).trim();
  }
  return jsonOut({ success: true, keys: keys });
}

// Validate batch
function handleValidateBatch(data) {
  var logEntry = { domain: '', accountKey: '', urlCount: 0, status: '', processed: 0, notes: '' };
  const urls = Array.isArray(data.urls) ? data.urls : [];
  logEntry.urlCount = urls.length;
  if (urls.length === 0) { logEntry.status = 'Bad Request'; logRequest(logEntry); return { success: false, error: 'No URLs provided.' }; }
  if (urls.length > 100) { logEntry.status = 'Bad Request'; logRequest(logEntry); return { success: false, error: 'Maximum 100 URLs allowed.' }; }
  const items = urls.map(function(url) {
    const cleanUrl = url.toString().trim();
    if (!cleanUrl) return { url: cleanUrl, ready: false, error: 'Empty URL' };
    const domain = extractDomainFromUrl(cleanUrl);
    if (!domain) return { url: cleanUrl, ready: false, error: 'Invalid URL format' };
    const accountKey = getAccountKeyForDomain(domain);
    if (!accountKey) return { url: cleanUrl, domain: domain, ready: false, error: 'Domain not in master list.' };
    if (!hasAccessCached(accountKey)) return { url: cleanUrl, domain: domain, accountKey: accountKey, ready: false, error: 'Account not authorised.' };
    return { url: cleanUrl, domain: domain, accountKey: accountKey, ready: true };
  });
  const readyCount = items.filter(function(i) { return i.ready; }).length;
  const uniqueDomains = [...new Set(items.filter(function(i) { return i.domain; }).map(function(i) { return i.domain; }))];
  const uniqueAccountKeys = [...new Set(items.filter(function(i) { return i.accountKey; }).map(function(i) { return i.accountKey; }))];
  logEntry.domain = uniqueDomains.join(', ').substring(0, 250);
  logEntry.accountKey = uniqueAccountKeys.join(', ').substring(0, 250);
  logEntry.status = 'Batch Validated';
  logEntry.notes = readyCount + ' ready, ' + (items.length - readyCount) + ' will fail';
  const logRow = logRequest(logEntry);
  const batchId = Utilities.getUuid();
  setBatchProgress(batchId, { logRow: logRow, readyCount: readyCount, processed: 0 });
  return { success: true, batchId: batchId, items: items, totalUrls: items.length, readyCount: readyCount };
}

// Inspect single URL
function handleInspectSingle(data) {
  const url = (data.url || '').toString().trim();
  const domain = (data.domain || '').toString().trim();
  const accountKey = (data.accountKey || '').toString().trim();
  const batchId = (data.batchId || '').toString().trim();
  if (!url || !domain || !accountKey) return { success: false, error: 'Missing url, domain, or accountKey' };
  if (!hasAccessCached(accountKey)) return { success: false, url: url, error: 'Account not authorised.' };
  const accessToken = getOAuthService(accountKey).getAccessToken();
  const acctEmail = getEmailForKey(accountKey) || accountKey;
  const prop = getCachedSiteProperty(domain, accountKey, accessToken, url);
  if (prop.error) return { success: false, url: url, error: prop.error };
  if (prop.notFound) return { success: false, url: url, error: 'No GSC property for ' + domain + ' on ' + acctEmail + '.' };
  if (prop.mismatch) return { success: false, url: url, error: 'URL version not covered. Available: ' + prop.available.join(', ') };
  const res = inspectUrl(url, prop.siteUrl, accessToken);
  if (res.code !== 200) { const msg = (res.body.error && res.body.error.message) || ('HTTP ' + res.code); return { success: false, url: url, error: 'GSC API ' + res.code + ': ' + msg }; }
  const r = res.body.inspectionResult && res.body.inspectionResult.indexStatusResult;
  const crawlTime = r && r.lastCrawlTime; const coverage = r && r.coverageState; const verdict = r && r.verdict;
  if (!crawlTime) {
    const reason = coverage || (verdict ? ('verdict ' + verdict) : 'no crawl data');
    if (batchId) { const p = incrementBatchProcessed(batchId); if (p && p.logRow) updateLogProcessed(p.logRow, p.processed); }
    return { success: true, url: url, domain: domain, accountKey: accountKey, lastCrawlDate: 'Never Crawled', indexStatus: reason, property: prop.siteUrl };
  }
  const formatted = Utilities.formatDate(new Date(crawlTime), 'Asia/Kolkata', 'dd MMM yyyy');
  if (batchId) { const p = incrementBatchProcessed(batchId); if (p && p.logRow) updateLogProcessed(p.logRow, p.processed); }
  return { success: true, url: url, domain: domain, accountKey: accountKey, lastCrawlDate: formatted, indexStatus: coverage || 'Unknown', property: prop.siteUrl };
}

// Save crawl batch
function handleSaveCrawlBatch(data) {
  var ss = SpreadsheetApp.openById(getSheetId());
  var sheet = ss.getSheetByName('CrawlHistory');
  if (!sheet) {
    sheet = ss.insertSheet('CrawlHistory');
    sheet.appendRow(['BatchId', 'UserId', 'UserName', 'AccountKey', 'UrlCount', 'Results', 'CreatedAt']);
    sheet.getRange(1, 1, 1, 7).setFontWeight('bold');
  } else {
    var existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(h) { return h.toString().trim(); });
    if (existingHeaders.indexOf('AccountKey') === -1) {
      sheet.insertColumnAfter(3);
      sheet.getRange(1, 4).setValue('AccountKey').setFontWeight('bold');
    }
  }
  var batchId = Utilities.getUuid();
  var results = Array.isArray(data.results) ? data.results : [];
  var uniqueAccountKeys = [...new Set(results.filter(function(r) { return r && r.accountKey; }).map(function(r) { return r.accountKey; }))];
  sheet.appendRow([batchId, data.userId || '', data.userName || '', uniqueAccountKeys.join(', '), data.urlCount || 0, JSON.stringify(results), new Date().toISOString()]);
  return { success: true, batchId: batchId };
}

// Get crawl history
function handleGetCrawlHistory(params) {
  var ss = SpreadsheetApp.openById(getSheetId());
  var sheet = ss.getSheetByName('CrawlHistory');
  if (!sheet) return ContentService.createTextOutput(JSON.stringify({ success: true, batches: [] })).setMimeType(ContentService.MimeType.JSON);
  var data = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h) { return h.toString().trim(); });
  var accountKeyCol = headers.indexOf('AccountKey');
  var urlCountCol = headers.indexOf('UrlCount');
  var resultsCol = headers.indexOf('Results');
  var createdAtCol = headers.indexOf('CreatedAt');
  var userId = params.userId || ''; var role = params.role || '';
  var all = [];
  for (var i = data.length - 1; i >= 1; i--) {
    var row = {
      id: data[i][0], userId: data[i][1], userName: data[i][2],
      accountKey: accountKeyCol >= 0 ? String(data[i][accountKeyCol] || '') : '',
      urlCount: Number(data[i][urlCountCol]) || 0,
      results: String(data[i][resultsCol]),
      createdAt: String(data[i][createdAtCol])
    };
    if (role === 'admin' || row.userId === userId) all.push(row);
    if (all.length >= 5) break;
  }
  return ContentService.createTextOutput(JSON.stringify({ success: true, batches: all })).setMimeType(ContentService.MimeType.JSON);
}

// Cleanup old crawl history.
function cleanupCrawlHistory() {
  var ss = SpreadsheetApp.openById(getSheetId());
  var sheet = ss.getSheetByName('CrawlHistory');
  if (!sheet) return;
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return;
  var headers = data[0].map(function(h) { return h.toString().trim(); });
  var createdAtCol = headers.indexOf('CreatedAt');
  if (createdAtCol === -1) createdAtCol = data[0].length - 1;
  var now = new Date().getTime(); var sixHours = 6 * 60 * 60 * 1000;
  // One deleteRow() call per stale row used to run here - same "N individual
  // writes" anti-pattern as the syncDomainsFromGSC/dedupeGscConfig timeout
  // bug (each deleteRow() is its own API round-trip AND reshuffles every row
  // below it). Replaced with the same keep-and-bulk-rewrite pattern: keep the
  // fresh rows in memory, write them back in one call, then trim the tail.
  var kept = [data[0]];
  for (var i = 1; i < data.length; i++) { if (now - new Date(data[i][createdAtCol]).getTime() <= sixHours) kept.push(data[i]); }
  if (kept.length === data.length) return;  // nothing stale - avoid a no-op rewrite
  if (kept.length > 1) sheet.getRange(1, 1, kept.length, data[0].length).setValues(kept);
  var newLastRow = kept.length;
  if (sheet.getLastRow() > newLastRow) sheet.deleteRows(newLastRow + 1, sheet.getLastRow() - newLastRow);
}

// Alert email handler
function handleSendAlert(payload) {
  const domain = payload.domain || 'Unknown domain'; const issues = payload.issues || [];
  const timestamp = payload.timestamp || new Date().toISOString();
  const accountEmail = payload.accountEmail || ''; const accessLevel = payload.accessLevel || '';
  const betterAccess = payload.betterAccess || null; const isRestricted = accessLevel.toLowerCase().indexOf('restricted') !== -1;
  const extraEmail = payload.extraEmail || '';
  if (issues.length === 0) return { success: true, sent: false, reason: 'No issues to report' };
  const settings = getSettings();
  const toRaw = settings['Alert Email TO'] || getProp('FALLBACK_ALERT_EMAIL') || _fallbackAdminEmails(); const ccRaw = settings['Alert Email CC'] || '';
  const toList = toRaw.split(',').map(e => e.trim()).filter(e => e.includes('@'));
  const ccList = ccRaw.split(',').map(e => e.trim()).filter(e => e.includes('@'));
  if (extraEmail && extraEmail.includes('@') && !toList.includes(extraEmail) && !ccList.includes(extraEmail)) ccList.push(extraEmail);
  if (toList.length === 0) return { success: false, error: 'No valid TO recipients' };
  const subject = '⚠️ GSC Audit Alert: Issues detected on ' + domain;
  let body = 'GSC Audit Alert for ' + domain + '\n' + issues.length + ' issue(s) found.\n';
  let auditTime;
  try { auditTime = Utilities.formatDate(new Date(timestamp), 'Asia/Kolkata', 'dd MMM yyyy, HH:mm') + ' IST'; }
  catch (e) { auditTime = timestamp; }
  const row = function (label, value) {
    return value ? '<tr><td style="padding:4px 12px 4px 0;color:#666;font-size:13px;white-space:nowrap">' + label + '</td>' +
                    '<td style="padding:4px 0;font-size:13px"><strong>' + value + '</strong></td></tr>' : '';
  };
  let html = '<div style="font-family:Arial,sans-serif;max-width:620px;margin:0 auto">' +
    '<div style="background:#0D3D4A;padding:22px 28px;border-radius:6px 6px 0 0">' +
    '<h2 style="color:#C9A84C;margin:0">⚠️ GSC Audit Alert</h2>' +
    '<p style="color:#cfd8dc;margin:6px 0 0;font-size:13px">Issues detected during automated webmaster audit</p></div>' +
    '<div style="background:#fff;padding:24px 28px;border:1px solid #dee2e6">' +
    '<table style="border-collapse:collapse;margin-bottom:18px">' +
    row('Domain', domain) + row('Audit Time', auditTime) + row('GSC Account', accountEmail) +
    row('Access Level', accessLevel) + row('Issues Found', issues.length + ' issue(s) require attention') +
    '</table>' +
    '<h3 style="font-size:15px;margin:0 0 10px;color:#0D3D4A">Issue Details</h3>';
  issues.forEach(function (issue, i) {
    html += '<div style="background:#fff8f8;border-left:4px solid #e74c3c;padding:14px 16px;margin-bottom:14px">' +
      '<strong>' + (i + 1) + '. ' + issue.type + '</strong>' +
      '<div style="margin-top:8px;color:#555;font-size:13px;white-space:pre-line">' + (issue.detail || '') + '</div></div>';
  });
  html += '<a href="https://search.google.com/search-console" ' +
    'style="display:inline-block;margin-top:8px;padding:10px 20px;background:#0D3D4A;color:#fff;' +
    'text-decoration:none;border-radius:4px;font-size:13px">Open Search Console</a>' +
    '</div></div>';
  try {
    const options = { htmlBody: html }; if (ccList.length > 0) options.cc = ccList.join(',');
    GmailApp.sendEmail(toList.join(','), subject, body, options);
    return { success: true, sent: true, to: toList, cc: ccList, issueCount: issues.length };
  } catch (err) { return { success: false, error: 'GmailApp error: ' + err.message }; }
}

// ============================================
// GSC PROPERTY VERSION CHECK
// ============================================
// For each domain in GSC_Config, checks whether the account we have GSC
// access for actually covers its live URL version:
//   - a domain property (sc-domain:) covers every version -> always correct
//   - a URL-prefix property matching the GSC-verified canonical -> correct
//   - only OTHER url-prefix version(s) accessible          -> WRONG VERSION
// Writes the result into GSC_Config (Version_Check_Status / _At / _Detail
// columns, auto-added if missing). Runs automatically, targeted at just the
// domains that changed, from checkVersionForChangedDomains() below (called
// by syncDomainsFromGSC's own trigger) - no separate daily sweep needed.

function _ensureVersionCheckColumns(sheet, headers) {
  var need = ['Version_Check_Status', 'Version_Check_At', 'Version_Check_Detail'];
  var added = false;
  need.forEach(function (name) {
    if (headers.indexOf(name) === -1) {
      sheet.getRange(1, headers.length + 1).setValue(name);
      headers.push(name);
      added = true;
    }
  });
  return added;
}

// First phase of the version check: resolves whether this domain's status is
// already FINAL (domain property / not found / no auth / error - no live
// inspection needed) or still NEEDS a live inspectUrl() call. Splitting this
// out lets the batch sweep fire many inspectUrl() calls together via
// UrlFetchApp.fetchAll() instead of one at a time. Returns either
// { final: <result-or-null> } or { propertyUrl, accessToken }.
function _prepareDomainForVersionCheck(domain, accountKey) {
  try {
    var service = getOAuthService(accountKey);
    if (!service.hasAccess()) {
      return { final: { status: 'No Auth', detail: 'Account not authorised.' } };
    }
    var accessToken = service.getAccessToken();

    // Domain (sc-domain:) properties inherently cover every protocol/www/
    // subdomain variant of their domain, so "does this property match the
    // exact live URL" isn't a meaningful question for them the way it is for
    // URL-prefix properties (which ARE strict about exact origin match).
    // Only run the live-version comparison for URL-prefix properties - a
    // domain property is Correct on its own, no inspection call needed.
    var domainProp = resolveSiteProperty(domain, accessToken, null, accountKey);
    if (domainProp.error) {
      return { final: { status: 'Error', detail: domainProp.error } };
    }
    if (domainProp.type === 'domain') {
      return { final: { status: 'Correct', detail: 'Domain property (sc-domain:) - covers all URL variants of ' + domain + '.' } };
    }
    if (domainProp.notFound) {
      return { final: { status: 'Wrong Version', detail: 'No GSC property found for ' + domain + ' on this account at all.' } };
    }
    if (!domainProp.mismatch || !domainProp.available || domainProp.available.length === 0) {
      return { final: null };   // unexpected shape - inconclusive, leave blank, retried next sweep
    }
    return { propertyUrl: domainProp.available[0], accessToken: accessToken };
  } catch (e) {
    return { final: { status: 'Error', detail: e.message } };
  }
}

// Second phase: turns ONE inspectUrl() response into the Correct/Wrong
// Version result shape. Shared by both the single-domain path and the
// batched path so there's exactly one place that defines "is this domain's
// GSC access the right URL version".
function _resultFromInspection(propertyUrl, code, body) {
  if (code !== 200) {
    return null;   // API error/timeout - inconclusive, retried next sweep
  }
  var idx = body.inspectionResult && body.inspectionResult.indexStatusResult;
  var canonical = idx && (idx.googleCanonical || idx.userCanonical);
  if (!canonical) {
    return null;   // not yet crawled / no canonical info yet - inconclusive
  }
  var bareHost = function (u) { return (u.match(/^https?:\/\/([^\/]+)/i) || ['', ''])[1].toLowerCase().replace(/^www\./, ''); };
  if (bareHost(canonical) === bareHost(propertyUrl)) {
    return { status: 'Correct', detail: 'GSC-verified canonical matches registered property: ' + canonical };
  }
  return { status: 'Wrong Version',
           detail: 'GSC-verified canonical is ' + canonical + ' but GSC access is only for: ' + propertyUrl };
}

// Runs the two-phase version check for ONE domain without touching the
// sheet - used by checkVersionForChangedDomains() below, the only caller
// now that this is check-on-change rather than a daily full-sheet sweep.
function _checkOneDomainVersion(domain, accountKey) {
  var prep = _prepareDomainForVersionCheck(domain, accountKey);
  if ('final' in prep) {
    return prep.final;
  }
  var res = inspectUrl(prep.propertyUrl, prep.propertyUrl, prep.accessToken);
  if (res.code !== 200) {
    return null;
  }
  return _resultFromInspection(prep.propertyUrl, res.code, res.body);
}

// Targeted version check for JUST the domains whose GSC access is new or
// changed in this sync run -- called from syncDomainsFromGSC() (the existing
// 6-hourly access-sync job) instead of waiting for the once-a-day full sweep
// of the whole (3000+ row) GSC_Config. Runs AFTER syncDomainsFromGSC's sort,
// so it re-locates each domain's row fresh rather than trusting stale row
// numbers from before the sort.
var VERSION_CHECK_BATCH_SIZE = 15;
// Leaves real headroom under Apps Script's 6-min hard execution ceiling,
// since this can run right after syncDomainsFromGSC's own work in the SAME
// invocation and still needs time left over to schedule its own
// continuation trigger if it runs out.
var VERSION_CHECK_BUDGET_MS = 4 * 60 * 1000;

// Clears Version_Check_Status/_At/_Detail for the given {domain, key} pairs
// (marking them as needing a fresh check), then kicks off the resumable
// sweep below. The sweep re-scans the SHEET for blank-status rows rather
// than being handed an explicit list to carry across executions - a list of
// a few thousand {domain,key} pairs doesn't fit in PropertiesService's ~9KB
// per-value cap (confirmed real case: clearing and re-syncing all 3000+
// GSC_Config rows at once left only 240 checked before Apps Script's 6-min
// execution ceiling killed the run, with no way to resume once the old
// dailyVersionCheck's row-by-row resume logic was removed along with the
// daily sweep). The sheet's own blank cells ARE the resumable queue - no
// size limit, and it self-heals even if the app crashes/redeploys mid-sweep.
function checkVersionForChangedDomains(domainKeyPairs) {
  if (!domainKeyPairs || domainKeyPairs.length === 0) return;
  var ss = SpreadsheetApp.openById(getSheetId());
  var config = ss.getSheetByName('GSC_Config');
  if (!config) return;
  var data = config.getDataRange().getValues();
  var headers = data[0].map(function (h) { return h.toString().trim(); });
  if (_ensureVersionCheckColumns(config, headers)) {
    data = config.getDataRange().getValues();
    headers = data[0].map(function (h) { return h.toString().trim(); });
  }
  var h = getHeaderMap(headers);
  var domainCol = h['Domain'] !== undefined ? h['Domain'] : 0;
  var acctCol = h['Account Key'] !== undefined ? h['Account Key'] : 2;
  var statusCol = h['Version_Check_Status'];

  var rowByDomainKey = {};
  for (var i = 1; i < data.length; i++) {
    var d = (data[i][domainCol] || '').toString().trim().toLowerCase();
    var k = (data[i][acctCol] || '').toString().trim();
    if (d && k) rowByDomainKey[d + '||' + k] = i + 1;
  }
  domainKeyPairs.forEach(function (pair) {
    var domain = (pair.domain || '').toString().trim().toLowerCase();
    var accountKey = (pair.key || pair.accountKey || '').toString().trim();
    var row = rowByDomainKey[domain + '||' + accountKey];
    if (row) config.getRange(row, statusCol + 1).setValue('');
  });

  _runVersionCheckSweep();
}

// Resumable: processes every GSC_Config row with an active (non-"Removed")
// Access Level and a blank Version_Check_Status, in fetchAll()-batched
// rounds, until either none are left or the time budget runs out - in which
// case it schedules a ONE-SHOT continuation trigger (not a recurring daily
// one) to pick up where it left off, then stops entirely once caught up.
// Guarded against overlapping runs the same way the old dailyVersionCheck
// was (a fresh call always supersedes an older one still mid-sweep).
function _runVersionCheckSweep() {
  var props = PropertiesService.getScriptProperties();
  var myToken = Utilities.getUuid();
  props.setProperty('version_check_active_token', myToken);

  var ss = SpreadsheetApp.openById(getSheetId());
  var config = ss.getSheetByName('GSC_Config');
  if (!config) return;
  var data = config.getDataRange().getValues();
  var headers = data[0].map(function (h) { return h.toString().trim(); });
  if (_ensureVersionCheckColumns(config, headers)) {
    data = config.getDataRange().getValues();
    headers = data[0].map(function (h) { return h.toString().trim(); });
  }
  var h = getHeaderMap(headers);
  var domainCol = h['Domain'] !== undefined ? h['Domain'] : 0;
  var acctCol = h['Account Key'] !== undefined ? h['Account Key'] : 2;
  var accessCol = h['Access Level'] !== undefined ? h['Access Level'] : 3;
  var statusCol = h['Version_Check_Status'];
  var atCol = h['Version_Check_At'];
  var detailCol = h['Version_Check_Detail'];

  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var domain = (data[i][domainCol] || '').toString().trim();
    var accountKey = (data[i][acctCol] || '').toString().trim();
    var access = (data[i][accessCol] || '').toString().trim();
    var status = (data[i][statusCol] || '').toString().trim();
    if (!domain || !accountKey || status) continue;
    if (access.indexOf('Removed') === 0) continue;  // no live access - nothing to check
    rows.push({ rowIndex: i + 1, domain: domain, accountKey: accountKey });
  }
  if (rows.length === 0) return;

  var rowInfoByIndex = {};
  rows.forEach(function (r) { rowInfoByIndex[r.rowIndex] = r; });

  var startTime = Date.now();
  var now = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'dd MMM yyyy HH:mm');
  var wrong = [], errors = [], correctCount = 0, checkedCount = 0;

  for (var idx = 0; idx < rows.length; idx += VERSION_CHECK_BATCH_SIZE) {
    if (props.getProperty('version_check_active_token') !== myToken) {
      Logger.log('_runVersionCheckSweep: superseded by a newer run, stopping.');
      return;
    }
    if (Date.now() - startTime > VERSION_CHECK_BUDGET_MS) {
      Logger.log('_runVersionCheckSweep: time budget reached (' + checkedCount + '/' + rows.length +
        ' checked this round), continuing in ~2 min.');
      _scheduleVersionCheckSweepContinuation();
      if (wrong.length > 0 || errors.length > 0) {
        _sendVersionCheckSummaryEmail(wrong, errors, correctCount, checkedCount);
      }
      return;
    }

    var batch = rows.slice(idx, idx + VERSION_CHECK_BATCH_SIZE);
    var pending = [];         // {rowIndex, propertyUrl, accessToken}
    var rowResults = {};      // rowIndex -> result

    batch.forEach(function (r) {
      try {
        var prep = _prepareDomainForVersionCheck(r.domain, r.accountKey);
        if ('final' in prep) {
          rowResults[r.rowIndex] = prep.final;
        } else if (prep.propertyUrl) {
          pending.push({ rowIndex: r.rowIndex, propertyUrl: prep.propertyUrl, accessToken: prep.accessToken });
        }
      } catch (e) {
        rowResults[r.rowIndex] = { status: 'Error', detail: e.message };
      }
    });

    if (pending.length > 0) {
      var inspectResults = _batchInspectUrls(pending);
      pending.forEach(function (p, pi) { rowResults[p.rowIndex] = inspectResults[pi]; });
    }

    // Daily UrlFetchApp quota is shared account-wide with everything else
    // running on this script (Crawl Tracker's live checks, syncDomainsFromGSC,
    // etc.) - once it's exhausted, every remaining call in this sweep would
    // just fail with the same "Service invoked too many times for one day"
    // error, and retrying every ~2 minutes for the rest of the day only adds
    // more failed calls competing for whatever trickle of quota live tools
    // need. Detect it and stop the sweep entirely for today instead - it
    // resumes on its own next time syncDomainsFromGSC's regular trigger runs
    // (quota resets daily), so live/user-triggered usage always gets
    // priority over this background catch-up work.
    var quotaExhausted = false;
    Object.keys(rowResults).forEach(function (rowIndexStr) {
      var result = rowResults[rowIndexStr];
      if (!result) return;
      if (result.status === 'Error' && /too many times for one day/i.test(result.detail || '')) {
        quotaExhausted = true;
      }
      var rn = Number(rowIndexStr);
      var info = rowInfoByIndex[rn] || { domain: '' };
      config.getRange(rn, statusCol + 1).setValue(result.status);
      config.getRange(rn, atCol + 1).setValue(now);
      config.getRange(rn, detailCol + 1).setValue(result.detail);
      checkedCount++;
      if (result.status === 'Correct') correctCount++;
      else if (result.status === 'Error') errors.push({ domain: info.domain, error: result.detail });
      else wrong.push({ domain: info.domain, detail: result.detail });
    });

    if (quotaExhausted) {
      Logger.log('_runVersionCheckSweep: daily UrlFetchApp quota exhausted (' + checkedCount + '/' +
        rows.length + ' checked) - stopping for today instead of retrying every 2 min. Will resume ' +
        'automatically once quota resets and syncDomainsFromGSC runs again.');
      _clearVersionCheckSweepContinuationTriggers();
      if (wrong.length > 0 || errors.length > 0) {
        _sendVersionCheckSummaryEmail(wrong, errors, correctCount, checkedCount);
      }
      return;
    }
  }

  _clearVersionCheckSweepContinuationTriggers();
  if (wrong.length > 0 || errors.length > 0) {
    _sendVersionCheckSummaryEmail(wrong, errors, correctCount, checkedCount);
  }
}

function _scheduleVersionCheckSweepContinuation() {
  _clearVersionCheckSweepContinuationTriggers();
  try {
    var trigger = ScriptApp.newTrigger('_runVersionCheckSweep').timeBased().after(2 * 60 * 1000).create();
    PropertiesService.getScriptProperties().setProperty('version_check_sweep_trigger_id', trigger.getUniqueId());
  } catch (e) {
    Logger.log('Could not schedule version-check sweep continuation: ' + e.message);
  }
}

function _clearVersionCheckSweepContinuationTriggers() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('version_check_sweep_trigger_id');
  if (!id) return;
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getUniqueId() === id) { ScriptApp.deleteTrigger(triggers[i]); break; }
  }
  props.deleteProperty('version_check_sweep_trigger_id');
}

// Fires inspectUrl() for several domains at once via UrlFetchApp.fetchAll() -
// Apps Script executes the whole batch concurrently instead of one HTTP
// round-trip at a time, each of which takes ~1-3s regardless. `pending` is
// an array of {propertyUrl, accessToken}; returns an array of results in the
// SAME order (fetchAll preserves request order in its response array).
function _batchInspectUrls(pending) {
  var requests = pending.map(function (p) {
    return {
      url: 'https://searchconsole.googleapis.com/v1/urlInspection/index:inspect',
      method: 'post',
      headers: { 'Authorization': 'Bearer ' + p.accessToken, 'Content-Type': 'application/json' },
      payload: JSON.stringify({ inspectionUrl: p.propertyUrl, siteUrl: p.propertyUrl }),
      muteHttpExceptions: true
    };
  });
  var responses;
  try {
    responses = UrlFetchApp.fetchAll(requests);
  } catch (e) {
    // A single malformed request in the batch can make fetchAll() throw for
    // the WHOLE batch rather than just that one - fall back to running this
    // batch one at a time rather than losing every domain in it.
    Logger.log('fetchAll batch failed (' + e.message + '), falling back to sequential for this batch.');
    return pending.map(function (p) {
      var res = inspectUrl(p.propertyUrl, p.propertyUrl, p.accessToken);
      return _resultFromInspection(p.propertyUrl, res.code, res.body);
    });
  }
  return responses.map(function (resp, i) {
    var body = {};
    try { body = JSON.parse(resp.getContentText() || '{}'); } catch (e) {}
    return _resultFromInspection(pending[i].propertyUrl, resp.getResponseCode(), body);
  });
}

function _sendVersionCheckSummaryEmail(wrong, errors, correctCount, checkedCount) {
  var settings = getSettings();
  var toRaw = settings['Alert Email TO'] || getProp('FALLBACK_ALERT_EMAIL') || _fallbackAdminEmails();
  var ccRaw = settings['Alert Email CC'] || '';
  var toList = toRaw.split(',').map(function (e) { return e.trim(); }).filter(function (e) { return e.indexOf('@') !== -1; });
  var ccList = ccRaw.split(',').map(function (e) { return e.trim(); }).filter(function (e) { return e.indexOf('@') !== -1; });
  if (toList.length === 0) return;

  var subject = String.fromCodePoint(0x26A0) + ' GSC Property Version Check: ' + wrong.length + ' domain(s) need attention';
  var summaryLine = (checkedCount || 0) + ' domain(s) checked -- ' + (correctCount || 0) + ' correct, '
    + wrong.length + ' wrong version, ' + errors.length + ' could not be checked.';
  var html = '<div style="font-family:Arial,sans-serif;max-width:680px;margin:0 auto">'
    + '<div style="background:#0D3D4A;padding:22px 28px;border-radius:6px 6px 0 0">'
    + '<h2 style="color:#C9A84C;margin:0">Daily GSC Property Version Check</h2></div>'
    + '<div style="background:#fff;padding:24px 28px;border:1px solid #dee2e6">'
    + '<p style="color:#333;font-size:14px;margin-top:0">' + summaryLine + '</p>';
  if (wrong.length > 0) {
    html += '<h3 style="color:#e74c3c">Wrong URL version (' + wrong.length + ')</h3>';
    wrong.forEach(function (w) {
      html += '<div style="background:#fff8f8;border-left:4px solid #e74c3c;padding:14px 16px;margin-bottom:12px">'
        + '<strong>' + w.domain + '</strong><br>'
        + '<span style="color:#555;font-size:13px">' + w.detail + '</span></div>';
    });
  }
  if (errors.length > 0) {
    html += '<h3 style="color:#e67e22">Could not check (' + errors.length + ')</h3>';
    errors.forEach(function (e) {
      html += '<div style="background:#fffaf0;border-left:4px solid #e67e22;padding:14px 16px;margin-bottom:12px">'
        + '<strong>' + e.domain + '</strong><br>'
        + '<span style="color:#555;font-size:13px">' + e.error + '</span></div>';
    });
  }
  html += '<p style="color:#888;font-size:12px">Full detail + status for every domain is in the GSC_Config sheet (Version_Check_Status / Version_Check_At / Version_Check_Detail columns).</p>';
  html += '</div></div>';

  var options = { htmlBody: html };
  if (ccList.length > 0) options.cc = ccList.join(',');
  try {
    GmailApp.sendEmail(toList.join(','), subject, subject, options);
  } catch (e) {
    Logger.log('Version check summary email failed: ' + e.message);
  }
}

// OTP
function sendOtp(data) {
  var to = data.email; var code = String(data.code || ''); var mins = data.expiresMinutes || 10; var name = data.name || '';
  var subject = 'Your James SEO Tools verification code: ' + code;
  var html = '<div style="font-family:Arial;max-width:480px;margin:auto"><h2 style="color:#E63812">Password change verification</h2><p>Hi ' + (name || 'there') + ',</p><p>Use this code:</p><div style="font-size:34px;font-weight:bold;letter-spacing:8px;background:#FBF3EA;padding:16px;text-align:center;border-radius:10px">' + code + '</div><p style="color:#6B4A3A;font-size:13px">Expires in ' + mins + ' minutes.</p></div>';
  MailApp.sendEmail({ to: to, subject: subject, htmlBody: html });
  return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
}

// Multi-project support
function getProjects() {
  const cached = CACHE.get('projects_v1');
  if (cached) { try { return JSON.parse(cached); } catch (e) {} }
  const props = PropertiesService.getScriptProperties(); const caps = {};
  (props.getProperty('PROJECT_CAPS') || '').split(',').forEach(function (pair) { const b = pair.split(':'); if (b.length === 2 && b[0].trim()) caps[b[0].trim()] = Number(b[1]) || 100; });
  const projects = [];
  if (props.getProperty('CLIENT_ID')) projects.push({ name: 'project-1', clientId: props.getProperty('CLIENT_ID'), clientSecret: props.getProperty('CLIENT_SECRET'), maxUsers: caps['project-1'] || 100 });
  for (let n = 2; n <= 12; n++) { const id = props.getProperty('CLIENT_ID_' + n); if (!id) continue; projects.push({ name: 'project-' + n, clientId: id, clientSecret: props.getProperty('CLIENT_SECRET_' + n), maxUsers: caps['project-' + n] || 100 }); }
  CACHE.put('projects_v1', JSON.stringify(projects), 300);
  return projects;
}

function getProjectCreds(name) { const p = getProjects(); for (let i = 0; i < p.length; i++) if (p[i].name === name) return p[i]; return null; }

function getAccountsTab() {
  const ss = SpreadsheetApp.openById(getSheetId());
  let sh = ss.getSheetByName('GSC_Accounts');
  if (!sh) { sh = ss.insertSheet('GSC_Accounts'); sh.getRange(1, 1, 1, 4).setValues([['Account Key', 'GSC Email', 'Project', 'Report Recipient']]).setFontWeight('bold').setBackground('#1e1e2e').setFontColor('#00ff88'); sh.setFrozenRows(1); }
  return sh;
}

function getAccountsMap() {
  const cached = CACHE.get('accounts_map_v1');
  if (cached) { try { return JSON.parse(cached); } catch (e) {} }
  const sh = getAccountsTab(); const data = sh.getDataRange().getValues(); const map = {};
  for (let i = 1; i < data.length; i++) {
    const key = data[i][0] ? data[i][0].toString().trim() : '';
    if (!key) continue;
    map[key] = { row: i + 1, key: key, email: data[i][1] ? data[i][1].toString().trim() : '', project: data[i][2] ? data[i][2].toString().trim() : '', recipient: data[i][3] ? data[i][3].toString().trim() : '' };
  }
  CACHE.put('accounts_map_v1', JSON.stringify(map), 300);
  return map;
}

function syncAccountsRegistry() {
  const ss = SpreadsheetApp.openById(getSheetId());
  const config = ss.getSheetByName('GSC_Config');
  if (!config) throw new Error('GSC_Config tab missing.');
  const cfg = config.getDataRange().getValues();
  const distinct = {};
  for (let i = 1; i < cfg.length; i++) { const email = cfg[i][1] ? cfg[i][1].toString().trim() : ''; const key = cfg[i][2] ? cfg[i][2].toString().trim() : ''; if (!key) continue; if (!distinct[key]) distinct[key] = email; }
  const sh = getAccountsTab();
  try { applyProjectMapping(); } catch (e) {}
  CACHE.remove('accounts_map_v1'); CACHE.remove('projects_v1'); CACHE.remove('domain_account_map_v1');
  const existing = getAccountsMap(); const projects = getProjects();
  if (projects.length === 0) throw new Error('No Cloud projects found.');
  const store = PropertiesService.getScriptProperties();
  const hasToken = function (key) { return !!store.getProperty('oauth2.' + key); };
  const counts = {}; projects.forEach(function (p) { counts[p.name] = 0; });
  Object.keys(existing).forEach(function (k) { const pr = existing[k].project; if (pr && counts[pr] !== undefined) counts[pr]++; });
  const pickProject = function () { for (let i = 0; i < projects.length; i++) { const p = projects[i]; if ((counts[p.name] || 0) < p.maxUsers) { counts[p.name]++; return p.name; } } return 'UNASSIGNED'; };
  const newRows = [];
  Object.keys(distinct).forEach(function (key) {
    if (existing[key]) {
      // One setValues() call for both columns instead of two separate
      // setValue() calls - same batching principle as the fixes above,
      // bounded by distinct account count here so a smaller win, but free.
      const needEmail = !existing[key].email && distinct[key];
      const needProject = !existing[key].project;
      if (needEmail || needProject) {
        const email = needEmail ? distinct[key] : existing[key].email;
        const project = needProject ? (hasToken(key) ? 'project-1' : pickProject()) : existing[key].project;
        sh.getRange(existing[key].row, 2, 1, 2).setValues([[email, project]]);
      }
      return;
    }
    let proj; if (hasToken(key)) { proj = 'project-1'; counts['project-1'] = (counts['project-1'] || 0) + 1; } else { proj = pickProject(); }
    newRows.push([key, distinct[key] || '', proj, '']);
  });
  if (newRows.length > 0) sh.getRange(sh.getLastRow() + 1, 1, newRows.length, 4).setValues(newRows);
  CACHE.remove('accounts_map_v1'); CACHE.remove('projects_v1'); CACHE.remove('domain_account_map_v1');
  return newRows.length;
}

// Project mapping
const PROJECT_NAME_MAP = { 'gsc for claude': 'project-1', 'gsc-tracker-2': 'project-1' };
const PROJECT_MAPPING_TAB = 'All accounts with Pass';

function getProjectMappingByEmail() {
  const ss = SpreadsheetApp.openById(getSheetId()); const sh = ss.getSheetByName(PROJECT_MAPPING_TAB);
  if (!sh) throw new Error('Mapping tab not found: ' + PROJECT_MAPPING_TAB);
  const data = sh.getDataRange().getValues(); const headers = data[0].map(h => h.toString().trim().toLowerCase());
  const emailCol = headers.indexOf('gsc email'); const projCol = headers.indexOf('test users cloud project');
  if (emailCol === -1 || projCol === -1) throw new Error('Mapping tab needs headers "GSC Email" and "Test Users Cloud Project".');
  const map = {};
  for (let i = 1; i < data.length; i++) { const email = data[i][emailCol] ? data[i][emailCol].toString().trim().toLowerCase() : ''; const name = data[i][projCol] ? data[i][projCol].toString().trim().toLowerCase() : ''; if (!email || !name) continue; const code = PROJECT_NAME_MAP[name]; if (code) map[email] = code; }
  return map;
}

function applyProjectMapping() {
  const mapping = getProjectMappingByEmail(); const sh = getAccountsTab(); const data = sh.getDataRange().getValues();
  let changed = 0;
  for (let i = 1; i < data.length; i++) { const email = data[i][1] ? data[i][1].toString().trim().toLowerCase() : ''; if (!email) continue; const code = mapping[email]; if (!code) continue; if (data[i][2] !== code) { data[i][2] = code; changed++; } }
  if (changed > 0) sh.getRange(2, 1, data.length - 1, data[0].length).setValues(data.slice(1));
  CACHE.remove('accounts_map_v1'); CACHE.remove('domain_account_map_v1');
  return { changed: changed };
}

// CSV email helper
function csvCell(v) { v = (v === undefined || v === null) ? '' : v.toString(); return /[",\n\r]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v; }

function withProp(level, prop) { return (level && prop) ? (level + ' [' + prop + ']') : (level || ''); }

function sendRoutedSyncEmails(info) {
  const accounts = getAccountsMap(); const defaultRecipients = getNotificationEmails(); const c = info.changes;
  const recipientOf = function (key) { const a = accounts[key]; return (a && a.recipient && a.recipient.indexOf('@') !== -1) ? a.recipient : null; };
  const emailOf = function (key) { return (accounts[key] && accounts[key].email) ? accounts[key].email : ''; };
  const buckets = {}; const bucketFor = function (rcpt) { const k = rcpt || '__DEFAULT__'; if (!buckets[k]) buckets[k] = { rows: [], n: 0, c: 0, r: 0, e: 0 }; return buckets[k]; };
  c.newDomains.forEach(function (d) { const b = bucketFor(recipientOf(d.key)); b.rows.push(['New', d.domain, d.email, d.key, '', withProp(d.level, d.property), info.runDateShort]); b.n++; });
  c.accessChanged.forEach(function (x) { const b = bucketFor(recipientOf(x.key)); b.rows.push(['Access Changed', x.domain, emailOf(x.key), x.key, withProp(x.before, x.beforeProperty), withProp(x.after, x.afterProperty), info.runDateShort]); b.c++; });
  c.removed.forEach(function (x) { const b = bucketFor(recipientOf(x.key)); b.rows.push(['Removed', x.domain, emailOf(x.key), x.key, withProp(x.lastLevel, x.lastProperty), 'Removed', info.runDateShort]); b.r++; });
  c.accountsWithErrors.forEach(function (s) { const key = s.split(' ')[0]; const b = bucketFor(recipientOf(key)); b.rows.push(['Account Error', '', emailOf(key), key, '', s, info.runDateShort]); b.e++; });
  Object.keys(buckets).forEach(function (k) {
    const bucket = buckets[k]; if (bucket.rows.length === 0) return;
    let toList; if (k === '__DEFAULT__') { if (defaultRecipients.length === 0) return; toList = defaultRecipients.join(','); } else { toList = k; }
    sendOneRoutedEmail(toList, bucket, info);
  });
}

function sendOneRoutedEmail(toList, bucket, info) {
  const header = ['Change Type', 'Domain', 'GSC Email', 'Account Key', 'Before Access', 'After Access', 'Date'];
  let csv = 'ï»¿' + header.map(csvCell).join(',') + '\r\n';
  bucket.rows.forEach(function (row) { csv += row.map(csvCell).join(',') + '\r\n'; });
  const fileName = 'GSC_Config_changes_' + info.runDateShort.replace(/\s+/g, '_') + '.csv';
  const blob = Utilities.newBlob(csv, 'text/csv', fileName);
  const subject = 'GSC_Config sync: ' + bucket.n + ' new, ' + bucket.c + ' changed, ' + bucket.r + ' removed';
  let html = '<div style="font-family:Arial"><h2>GSC_Config changes</h2><p>Run: ' + info.runTimestamp + '</p><ul><li>' + bucket.n + ' new</li><li>' + bucket.c + ' changed</li><li>' + bucket.r + ' removed</li></ul><p>See attached CSV.</p></div>';
  try { MailApp.sendEmail({ to: toList, subject: subject, htmlBody: html, attachments: [blob] }); } catch (err) { Logger.log('Email failed: ' + err.message); }
}

// Auth status check
function recheckAuthStatus() {
  const ss = SpreadsheetApp.openById(getSheetId()); const config = ss.getSheetByName('GSC_Config');
  const data = config.getDataRange().getValues(); const seen = {}; const keys = [];
  for (let i = 1; i < data.length; i++) { const key = data[i][2] ? data[i][2].toString().trim() : ''; if (!key || seen[key]) continue; seen[key] = true; keys.push(key); }
  const result = {}; let done = 0, pending = 0; const pendingKeys = [];
  keys.forEach(function (key) { const ok = getOAuthService(key).hasAccess(); result[key] = ok; if (ok) done++; else { pending++; pendingKeys.push(key); } });
  try {
    const authSheet = ss.getSheetByName('Auth_Status');
    if (authSheet && authSheet.getLastRow() > 1) { const range = authSheet.getRange(2, 2, authSheet.getLastRow() - 1, 3); const vals = range.getValues(); for (let i = 0; i < vals.length; i++) { const key = vals[i][0] ? vals[i][0].toString().trim() : ''; if (key && result[key] === true) { vals[i][1] = ''; vals[i][2] = 'Authorised'; } } range.setValues(vals); }
  } catch (e) {}
  return { done, pending, pendingKeys };
}

// Sheet menu
function onOpen() {
  SpreadsheetApp.getUi().createMenu('📊 GSC Tools')
    .addItem('🔄 Sync Domains', 'menuSyncDomains')
    .addItem('✅ Check Authorisation', 'menuCheckAuthStatus')
    .addItem('🔐 Generate Auth URLs', 'menuGenerateAuthSheet')
    .addItem('📒 Sync Account Registry', 'menuSyncAccounts')
    .addItem('🏷️ Apply Project Labels', 'menuApplyProjectMapping')
    .addItem('♻️ Reset One Account', 'menuResetOneAccount')
    .addItem('🚨 Clear ALL Tokens - DO NOT USE until confirmed', 'menuClearAllTokens')
    .addSeparator()
    .addItem('⚙️ Settings', 'menuOpenSettings')
    .addItem('📧 Send Test Email', 'menuSendTestEmail')
    .addItem('ℹ️ Info', 'menuShowInfo')
    .addToUi();
}

// Menu handlers
function menuSyncDomains() { const ui = SpreadsheetApp.getUi(); if (ui.alert('Sync domains from GSC', 'This pulls all properties from every authorised account.\n\nContinue?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return; try { syncDomainsFromGSC(); ui.alert('Done', 'Sync complete.', ui.ButtonSet.OK); } catch (e) { ui.alert('Error', e.message, ui.ButtonSet.OK); } }
function menuCheckAuthStatus() { const ui = SpreadsheetApp.getUi(); try { const r = recheckAuthStatus(); ui.alert('Auth Status', 'Authorised: ' + r.done + '\nPending: ' + r.pending + (r.pending > 0 ? '\n\n' + r.pendingKeys.join('\n') : ''), ui.ButtonSet.OK); } catch (e) { ui.alert('Error', e.message, ui.ButtonSet.OK); } }
function menuGenerateAuthSheet() { const ui = SpreadsheetApp.getUi(); if (ui.alert('Generate auth URLs', 'Refresh Auth_Status tab?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return; try { generateAuthSheet(); ui.alert('Done', 'Auth_Status refreshed.', ui.ButtonSet.OK); } catch (e) { ui.alert('Error', e.message, ui.ButtonSet.OK); } }
function menuResetOneAccount() { const ui = SpreadsheetApp.getUi(); const r = ui.prompt('Reset ONE account', 'Enter account key:', ui.ButtonSet.OK_CANCEL); if (r.getSelectedButton() !== ui.Button.OK) return; const key = r.getResponseText().trim(); if (!key) return; try { getOAuthService(key).reset(); ui.alert('Done', 'Token reset for: ' + key, ui.ButtonSet.OK); } catch (e) { ui.alert('Error', e.message, ui.ButtonSet.OK); } }
function menuClearAllTokens() { const ui = SpreadsheetApp.getUi(); if (ui.alert('Clear ALL tokens', 'This wipes ALL tokens. Continue?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return; try { const store = PropertiesService.getScriptProperties(); const keys = store.getKeys().filter(k => k.indexOf('oauth2.') === 0); keys.forEach(k => store.deleteProperty(k)); ui.alert('Done', keys.length + ' tokens deleted.', ui.ButtonSet.OK); } catch (e) { ui.alert('Error', e.message, ui.ButtonSet.OK); } }
function menuOpenSettings() { const ss = SpreadsheetApp.openById(getSheetId()); let sh = ss.getSheetByName('GSC_Settings'); if (!sh) sh = createSettingsTab(); sh.activate(); }
function menuSendTestEmail() { const ui = SpreadsheetApp.getUi(); try { const r = getNotificationEmails(); if (r.length === 0) { ui.alert('No recipients', 'Set emails in GSC_Settings.', ui.ButtonSet.OK); return; } MailApp.sendEmail({ to: r.join(','), subject: 'SEO Toolkit Pro â€” Test email', htmlBody: '<p>Test email working. Recipients: ' + r.join(', ') + '</p>' }); ui.alert('Sent', 'Test sent to: ' + r.join(', '), ui.ButtonSet.OK); } catch (e) { ui.alert('Error', e.message, ui.ButtonSet.OK); } }
function menuShowInfo() { const ui = SpreadsheetApp.getUi(); try { ui.alert('Script Info', 'Owner: ' + Session.getEffectiveUser().getEmail() + '\nSheet: ' + getSheetId() + '\nRedirect URI: https://script.google.com/macros/d/' + ScriptApp.getScriptId() + '/usercallback', ui.ButtonSet.OK); } catch (e) { ui.alert('Error', e.message, ui.ButtonSet.OK); } }
function menuSyncAccounts() { const ui = SpreadsheetApp.getUi(); try { const added = syncAccountsRegistry(); ui.alert('Done', added + ' new account(s) added.', ui.ButtonSet.OK); } catch (e) { ui.alert('Error', e.message, ui.ButtonSet.OK); } }
function menuApplyProjectMapping() { const ui = SpreadsheetApp.getUi(); try { const r = applyProjectMapping(); ui.alert('Done', 'Updated: ' + r.changed, ui.ButtonSet.OK); } catch (e) { ui.alert('Error', e.message, ui.ButtonSet.OK); } }

// Debug helpers
function printExactRedirectUri() { Logger.log('https://script.google.com/macros/d/' + ScriptApp.getScriptId() + '/usercallback'); }
function verifyCredentials() { Logger.log('CLIENT_ID: ' + (getClientId() ? 'OK' : 'MISSING')); Logger.log('CLIENT_SECRET: ' + (getClientSecret() ? 'OK' : 'MISSING')); Logger.log('SHEET_ID: ' + (getSheetId() ? 'OK' : 'MISSING')); Logger.log('ADMIN_KEY: ' + (getAdminKey() ? 'OK' : 'MISSING')); }

// ============================================
// ONE-OFF DIAGNOSTIC â€” run manually from the Apps Script editor.
// Confirms which real Google identity the access token belongs to,
// which clientId was used, what property got resolved, and the raw
// inspect response â€” for a single accountKey + URL.
//
// Usage: edit the three constants below, then Run > debugInspectOne.
// Read the log (View > Logs / Executions). Delete or ignore afterwards.
// ============================================
function debugInspectOne() {
  var ACCOUNT_KEY = 'searchconsoleanalyticsaccess8';
  var DOMAIN = 'aerokitchen.com';
  var URL_TO_INSPECT = 'https://aerokitchen.com/kitchens/';

  var acctRow = getAccountsMap()[ACCOUNT_KEY];
  Logger.log('GSC_Accounts row for key: ' + JSON.stringify(acctRow));

  var clientId = getClientId();
  if (acctRow && acctRow.project) {
    var creds = getProjectCreds(acctRow.project);
    if (creds && creds.clientId) clientId = creds.clientId;
  }
  Logger.log('clientId that will be used (masked): ...' + String(clientId).slice(-12));

  var service = getOAuthService(ACCOUNT_KEY);
  Logger.log('hasAccess(): ' + service.hasAccess());
  if (!service.hasAccess()) { Logger.log('lastError: ' + service.getLastError()); return; }

  var accessToken = service.getAccessToken();

  // Ask Google directly whose token this is and what scopes it carries.
  var tokenInfo = UrlFetchApp.fetch(
    'https://oauth2.googleapis.com/tokeninfo?access_token=' + encodeURIComponent(accessToken),
    { muteHttpExceptions: true }
  );
  Logger.log('tokeninfo (' + tokenInfo.getResponseCode() + '): ' + tokenInfo.getContentText());
  // Compare the "email" field above to expected account email below:
  Logger.log('Expected email for this key: ' + getEmailForKey(ACCOUNT_KEY));

  var prop = resolveSiteProperty(DOMAIN, accessToken, URL_TO_INSPECT);
  Logger.log('resolveSiteProperty result: ' + JSON.stringify(prop));

  if (prop.siteUrl) {
    var res = inspectUrl(URL_TO_INSPECT, prop.siteUrl, accessToken);
    Logger.log('inspectUrl code: ' + res.code + ' body: ' + JSON.stringify(res.body));
  }
}



function setAdminKey() {
  PropertiesService.getScriptProperties().setProperty('ADMIN_KEY', '44bXCd9yaa#xMR!E');
  Logger.log('44bXCd9yaa#xMR!E.');
}

// ============================================
// ONE-TIME: run this manually from the Apps Script editor (Run button) to
// create the first super admin. Change SUPER_ADMIN_PASSWORD below first,
// then change the password again from the Admin Panel after your first
// login (or just re-run this function with a new password).
// ============================================
function bootstrapSuperAdmin() {
  var SUPER_ADMIN_EMAIL = 'vishal.chhipa.ptp@digigyan.org';
  var SUPER_ADMIN_PASSWORD = 'CHANGE_ME_BEFORE_RUNNING'; // <-- set a real password first

  if (SUPER_ADMIN_PASSWORD === 'CHANGE_ME_BEFORE_RUNNING') {
    throw new Error('Set SUPER_ADMIN_PASSWORD to a real password before running bootstrapSuperAdmin().');
  }

  var email = SUPER_ADMIN_EMAIL.trim().toLowerCase();
  var sheet = _getAuthSheet('users');
  var data = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h) { return h.toString().trim(); });
  var h = getHeaderMap(headers);

  var rowIndex = -1;
  for (var i = 1; i < data.length; i++) {
    if ((data[i][h['Email']] || '').toString().trim().toLowerCase() === email) { rowIndex = i + 1; break; }
  }

  var existing = rowIndex !== -1 ? data[rowIndex - 1] : [];
  var row = [];
  row[h['Email']] = email;
  row[h['Password']] = SUPER_ADMIN_PASSWORD;
  row[h['Device_ID']] = existing[h['Device_ID']] || '';
  row[h['Approved']] = true;
  row[h['Name']] = existing[h['Name']] || 'Super Admin';
  row[h['Registered']] = existing[h['Registered']] || new Date();
  row[h['Last Login']] = existing[h['Last Login']] || '';
  row[h['Notes']] = 'Super admin (bootstrap)';
  row[h['is_admin']] = true;
  row[h['formats']] = 'All';
  row[h['tools']] = 'All';
  var finalRow = headers.map(function(_, i) { return row[i] !== undefined ? row[i] : ''; });

  if (rowIndex === -1) sheet.appendRow(finalRow);
  else sheet.getRange(rowIndex, 1, 1, finalRow.length).setValues([finalRow]);

  Logger.log('Super admin ready: ' + email);
}

function _syncUserDevice(email, mac) {
  email = (email || '').toString().trim().toLowerCase();
  mac = (mac || '').toString().trim().toUpperCase();
  if (!email) return;

  var sheet = _getAuthSheet('users');
  var data = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h) { return h.toString().trim(); });
  var h = getHeaderMap(headers);
  if (h['Email'] === undefined) return;

  for (var i = 1; i < data.length; i++) {
    if ((data[i][h['Email']] || '').toString().trim().toLowerCase() !== email) continue;
    if (h['Device_ID'] !== undefined) sheet.getRange(i + 1, h['Device_ID'] + 1).setValue(mac);
    if (h['Last Login'] !== undefined) sheet.getRange(i + 1, h['Last Login'] + 1).setValue(new Date());
  }
}

function _authAdminUpsert(p) {
  if (!_requireAdmin(p)) {
    return jsonOut({ error: 'Unauthorized' });
  }

  const email = (p.email || '').toString().trim().toLowerCase();
  if (!email) return jsonOut({ error: 'Missing email' });

  const sheet = _getAuthSheet('users');
  let data = sheet.getDataRange().getValues();

  let headers = data[0].map(h => h.toString().trim());

  // Ensure the 'tools' column exists on an existing sheet, then re-read.
  if (headers.indexOf('tools') < 0) {
    sheet.getRange(1, headers.length + 1).setValue('tools');
    data = sheet.getDataRange().getValues();
    headers = data[0].map(h => h.toString().trim());
  }

  const h = getHeaderMap(headers);

  let rowIndex = -1;

  // find user by Email column (NOT column 0 assumption)
  for (let i = 1; i < data.length; i++) {
    const rowEmail = (data[i][h['Email']] || '').toString().trim().toLowerCase();
    if (rowEmail === email) {
      rowIndex = i + 1;
      break;
    }
  }

  const existing = rowIndex !== -1 ? data[rowIndex - 1] : [];

  const row = [];
  row[h['Email']] = email;
  row[h['Password']] = p.password ? p.password : (existing[h['Password']] || '');
  row[h['Name']] = p.name || '';
  row[h['Approved']] = (p.approved === true || p.approved === 'true');
  row[h['is_admin']] = (p.is_admin === true || p.is_admin === 'true');
  row[h['Device_ID']] = p.clear_mac ? '' : (p.device !== undefined ? p.device : (existing[h['Device_ID']] || ''));
  row[h['formats']] = p.formats || '';
  row[h['tools']] = p.tools || '';
  row[h['Last Login']] = existing[h['Last Login']] || '';
  row[h['Notes']] = p.notes || '';

  // normalize missing columns (avoid undefined holes)
  const finalRow = headers.map((_, i) => row[i] !== undefined ? row[i] : '');

  if (rowIndex === -1) {
    sheet.appendRow(finalRow);
    if (p.clear_mac || p.device !== undefined) {
      _syncUserDevice(email, p.clear_mac ? '' : (p.device !== undefined ? p.device : (existing[h['Device_ID']] || '')));
    }
    return jsonOut({ success: true, mode: 'created' });
  }

  sheet.getRange(rowIndex, 1, 1, finalRow.length).setValues([finalRow]);
  if (p.clear_mac || p.device !== undefined) {
    _syncUserDevice(email, p.clear_mac ? '' : (p.device !== undefined ? p.device : (existing[h['Device_ID']] || '')));
  }
  return jsonOut({ success: true, mode: 'updated' });
}

function getHeaderMap(headers) {
  const map = {};
  headers.forEach((h, i) => {
    map[h.toString().trim()] = i;
  });
  return map;
}



