/**
 * QuickBooks Online integration (read-only, SANDBOX company only for now).
 * =========================================================================
 * Uses the community "OAuth2 for Apps Script" library
 * (https://github.com/googleworkspace/apps-script-oauth2), referenced in
 * appsscript.json as the "OAuth2" library.
 *
 * Script Properties required (set once via Apps Script editor or clasp run,
 * never committed to source):
 *   QBO_CLIENT_ID, QBO_CLIENT_SECRET  — from the Intuit Developer app's
 *   "Keys & OAuth" page (use the DEVELOPMENT/sandbox keys, not production).
 *
 * QBO_REALM_ID (the sandbox company id) is written automatically to Script
 * Properties the first time someone completes the Connect flow.
 */

var QBO_SANDBOX_BASE_URL = 'https://sandbox-quickbooks.api.intuit.com';

/** Owner-only gate (same convention as the push-notifications admin card): admin role AND owner email. */
function authorizeQuickBooksOwner_(payload) {
  var auth = authorizeQuickBooksOwner_(payload);
  if (!auth.ok) return auth;
  if (!isOwnerEmail(auth.email)) {
    return { ok: false, code: 'FORBIDDEN', error: 'You do not have permission to do this.' };
  }
  return auth;
}

function getQuickBooksService_() {
  var props = PropertiesService.getScriptProperties();
  var clientId = props.getProperty('QBO_CLIENT_ID');
  var clientSecret = props.getProperty('QBO_CLIENT_SECRET');
  return OAuth2.createService('quickbooks')
    .setAuthorizationBaseUrl('https://appcenter.intuit.com/connect/oauth2')
    .setTokenUrl('https://oauth.platform.intuit.com/oauth2/v1/tokens/bearer')
    .setClientId(clientId)
    .setClientSecret(clientSecret)
    .setCallbackFunction('quickbooksAuthCallback_')
    .setPropertyStore(props)
    .setScope('com.intuit.quickbooks.accounting')
    // Intuit's token endpoint requires client credentials via HTTP Basic auth,
    // not in the request body (the library's default) - without this the
    // token exchange fails with invalid_client.
    .setTokenHeaders({
      'Authorization': 'Basic ' + Utilities.base64Encode(clientId + ':' + clientSecret),
      'Content-Type': 'application/x-www-form-urlencoded'
    });
}

/** Called from the "Connect QuickBooks" admin action; client opens the returned URL in a new tab. */
function getQuickBooksAuthorizationUrl(payload) {
  var auth = authorizeQuickBooksOwner_(payload);
  if (!auth.ok) return { error: auth.error, code: auth.code };
  return { success: true, url: getQuickBooksService_().getAuthorizationUrl() };
}

/** Invoked directly by doGet() when Intuit redirects back with ?qboCallback=1. */
function quickbooksAuthCallback_(e) {
  var service = getQuickBooksService_();
  var isAuthorized = service.handleCallback(e);
  if (isAuthorized && e.parameter.realmId) {
    PropertiesService.getScriptProperties().setProperty('QBO_REALM_ID', e.parameter.realmId);
  }
  var message = isAuthorized
    ? 'QuickBooks connected. You can close this tab and return to the app.'
    : 'QuickBooks authorization failed or was denied. You can close this tab and try again.';
  return HtmlService.createHtmlOutput('<p style="font-family:sans-serif;padding:24px">' + message + '</p>');
}

function getQuickBooksStatus(payload) {
  var auth = authorizeQuickBooksOwner_(payload);
  if (!auth.ok) return { error: auth.error, code: auth.code };
  var connected = getQuickBooksService_().hasAccess();
  var realmId = PropertiesService.getScriptProperties().getProperty('QBO_REALM_ID') || '';
  return { success: true, connected: connected, realmId: realmId };
}

function disconnectQuickBooks(payload) {
  var auth = authorizeQuickBooksOwner_(payload);
  if (!auth.ok) return { error: auth.error, code: auth.code };
  getQuickBooksService_().reset();
  PropertiesService.getScriptProperties().deleteProperty('QBO_REALM_ID');
  return { success: true };
}

function quickbooksApiGet_(path) {
  var service = getQuickBooksService_();
  if (!service.hasAccess()) {
    return { success: false, error: 'QuickBooks is not connected yet.' };
  }
  var realmId = PropertiesService.getScriptProperties().getProperty('QBO_REALM_ID');
  if (!realmId) {
    return { success: false, error: 'Missing QuickBooks company (realm) id - reconnect QuickBooks.' };
  }
  var url = QBO_SANDBOX_BASE_URL + '/v3/company/' + realmId + path;
  var response = UrlFetchApp.fetch(url, {
    headers: {
      Authorization: 'Bearer ' + service.getAccessToken(),
      Accept: 'application/json'
    },
    muteHttpExceptions: true
  });
  var code = response.getResponseCode();
  var body = response.getContentText();
  if (code !== 200) {
    return { success: false, error: 'QuickBooks API error ' + code + ': ' + body };
  }
  return { success: true, data: JSON.parse(body) };
}

/** Read-only connectivity test: company info proves the OAuth round-trip works end to end. */
function testQuickBooksConnection(payload) {
  var auth = authorizeQuickBooksOwner_(payload);
  if (!auth.ok) return { error: auth.error, code: auth.code };
  var realmId = PropertiesService.getScriptProperties().getProperty('QBO_REALM_ID');
  if (!realmId) return { success: false, error: 'Not connected yet.' };
  return quickbooksApiGet_('/companyinfo/' + realmId);
}

/** Read-only: pulls a handful of vendors to sanity-check data access beyond company info. */
function testQuickBooksVendors(payload) {
  var auth = authorizeQuickBooksOwner_(payload);
  if (!auth.ok) return { error: auth.error, code: auth.code };
  return quickbooksApiGet_('/query?query=' + encodeURIComponent('select Id, DisplayName from Vendor maxresults 20'));
}
