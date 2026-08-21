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
var QBO_PRODUCTION_BASE_URL = 'https://quickbooks.api.intuit.com';

// ─── Warehouse/stock lane ────────────────────────────────────────────────────
// Stock POs use this fixed placeholder as their Job Ref (createPO requires a
// jobRef, so there's no such thing as "no job" in the app's own schema) --
// purely an app-side label, never looked up as or sent to QBO as a real
// Customer/Job. Every guardrail that distinguishes "warehouse stock" from
// "direct-to-job" keys off this same check.
var WAREHOUSE_JOB_REF = 'WAREHOUSE';
function isWarehouseJob_(jobRef) {
  return (jobRef || '').toString().trim().toUpperCase() === WAREHOUSE_JOB_REF;
}

/** Reads the QBO_ENVIRONMENT script property ('production' | anything else = sandbox, the safe default). */
function getQuickBooksBaseUrl_() {
  var env = (PropertiesService.getScriptProperties().getProperty('QBO_ENVIRONMENT') || '').toLowerCase();
  return env === 'production' ? QBO_PRODUCTION_BASE_URL : QBO_SANDBOX_BASE_URL;
}

/** Owner-only gate (same convention as the push-notifications admin card): admin role AND owner email. */
function authorizeQuickBooksOwner_(payload) {
  var auth = authorizeCaller(payload, ['admin']);
  if (!auth.ok) return auth;
  if (!isOwnerEmail(auth.email)) {
    return { ok: false, code: 'FORBIDDEN', error: 'You do not have permission to do this.' };
  }
  return auth;
}

/** Invoice-review pipeline gate: Admin or Office role (no owner-email restriction) -- upload/view/approve/create-bill. */
function authorizeInvoiceReviewer_(payload) {
  return authorizeCaller(payload, ['admin', 'office']);
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
    })
    // The library still puts client_id/client_secret in the request BODY by
    // default, in addition to the Basic-Auth header above -- the OAuth2 spec
    // treats Basic auth and body-based client auth as alternatives, not to
    // be combined. This worked against Sandbox regardless, but Intuit's
    // Production token endpoint has been reported (Aug 2026) to reject the
    // duplicate with invalid_client where Sandbox tolerates it. Stripping
    // them from the body leaves Basic Auth as the sole credential channel.
    .setTokenPayloadHandler(function(payload) {
      delete payload.client_id;
      delete payload.client_secret;
      return payload;
    });
}

/** Called from the "Connect QuickBooks" admin action; client opens the returned URL in a new tab. */
function getQuickBooksAuthorizationUrl(payload) {
  var auth = authorizeQuickBooksOwner_(payload);
  if (!auth.ok) return { error: auth.error, code: auth.code };
  return { success: true, url: getQuickBooksService_().getAuthorizationUrl() };
}

function qboEscHtml_(s) {
  return (s || '').toString()
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

/**
 * Invoked directly by the Apps Script platform's built-in OAuth2 callback
 * endpoint (https://script.google.com/macros/d/{scriptId}/usercallback) -
 * this bypasses doGet() entirely, which is the OAuth2 library's normal,
 * documented usage. Do not try to route this through doGet().
 *
 * handleCallback() throws (rather than returning false) when the token
 * exchange itself fails (e.g. invalid_client) -- previously uncaught, which
 * crashed straight to Apps Script's generic error page (no real message,
 * and no way to reset the half-completed OAuth2 library state short of a
 * successful Disconnect, which never shows because status never reached
 * "connected"). Catching it surfaces the actual error Intuit returned and
 * resets the service so the next Connect attempt starts clean.
 */
function quickbooksAuthCallback_(e) {
  var service = getQuickBooksService_();
  try {
    var isAuthorized = service.handleCallback(e);
    if (isAuthorized && e.parameter.realmId) {
      PropertiesService.getScriptProperties().setProperty('QBO_REALM_ID', e.parameter.realmId);
    }
    var message = isAuthorized
      ? 'QuickBooks connected. You can close this tab and return to the app.'
      : 'QuickBooks authorization failed or was denied. You can close this tab and try again.';
    return HtmlService.createHtmlOutput('<p style="font-family:sans-serif;padding:24px">' + message + '</p>');
  } catch (err) {
    service.reset();
    return HtmlService.createHtmlOutput(
      '<p style="font-family:sans-serif;padding:24px">QuickBooks connection failed:<br>' +
      '<code style="white-space:pre-wrap">' + qboEscHtml_(err && err.message) + '</code>' +
      '<br><br>You can close this tab and try again.</p>'
    );
  }
}

function getQuickBooksStatus(payload) {
  var auth = authorizeQuickBooksOwner_(payload);
  if (!auth.ok) return { error: auth.error, code: auth.code };
  var connected = getQuickBooksService_().hasAccess();
  var realmId = PropertiesService.getScriptProperties().getProperty('QBO_REALM_ID') || '';
  var env = (PropertiesService.getScriptProperties().getProperty('QBO_ENVIRONMENT') || '').toLowerCase() === 'production' ? 'production' : 'sandbox';
  return { success: true, connected: connected, realmId: realmId, environment: env };
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
  var url = getQuickBooksBaseUrl_() + '/v3/company/' + realmId + path;
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

/** Parses QBO's error body shape ({Fault:{Error:[{Message, Detail}]}}) into a single readable string, falling back to the raw body. */
function qboParseErrorBody_(body) {
  try {
    var parsed = JSON.parse(body);
    var errs = parsed.Fault && parsed.Fault.Error;
    if (errs && errs.length) {
      return errs.map(function(e) { return e.Message + (e.Detail ? ' -- ' + e.Detail : ''); }).join('; ');
    }
  } catch (e) { /* not JSON, or not QBO's Fault shape -- fall through */ }
  return body;
}

/** Same auth/realm handling as quickbooksApiGet_, but POST with a JSON body -- the write counterpart. */
function quickbooksApiPost_(path, bodyObj) {
  var service = getQuickBooksService_();
  if (!service.hasAccess()) {
    return { success: false, error: 'QuickBooks is not connected yet.' };
  }
  var realmId = PropertiesService.getScriptProperties().getProperty('QBO_REALM_ID');
  if (!realmId) {
    return { success: false, error: 'Missing QuickBooks company (realm) id - reconnect QuickBooks.' };
  }
  var url = getQuickBooksBaseUrl_() + '/v3/company/' + realmId + path;
  var response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(bodyObj),
    headers: {
      Authorization: 'Bearer ' + service.getAccessToken(),
      Accept: 'application/json'
    },
    muteHttpExceptions: true
  });
  var code = response.getResponseCode();
  var body = response.getContentText();
  if (code < 200 || code >= 300) {
    return { success: false, error: 'QuickBooks API error ' + code + ': ' + qboParseErrorBody_(body) };
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

/**
 * Read-only diagnostic: pulls Customers/Projects (Job=true rows are
 * sub-customers, which is what Projects are built on) with their actual
 * Accounting-API Id, DisplayName, and ParentRef. A purely-numeric filter is
 * treated as an exact Id lookup (Id is always queryable); anything else is
 * a DisplayName-prefix search (QBO renders sub-customers as "Parent:Child")
 * -- QBO's query engine rejects ParentRef in a WHERE clause entirely
 * (error 4001, "not queryable"), so these two are the only reliable ways
 * in. A QBO Projects-tab URL id (e.g. the "id=" in /app/projects/...) is
 * NOT the same identifier as CustomerRef.value needs -- confirmed via QBO's
 * own "New Bill" form URL, which separately exposes customerId=<real
 * CustomerRef value> and projectRef=<the Projects-tab id>.
 */
function testQuickBooksCustomers(payload) {
  var auth = authorizeQuickBooksOwner_(payload);
  if (!auth.ok) return { error: auth.error, code: auth.code };
  var filter = (payload.nameFilter || '').toString().trim();
  var isNumeric = /^[0-9]+$/.test(filter);
  var escaped = filter.replace(/'/g, "\\'");
  var query = !filter
    ? 'select Id, DisplayName, Job, Active, ParentRef, SubCustomer, BillWithParent from Customer maxresults 100'
    : isNumeric
      ? "select Id, DisplayName, Job, Active, ParentRef, SubCustomer, BillWithParent from Customer where Id = '" + escaped + "'"
      : "select Id, DisplayName, Job, Active, ParentRef, SubCustomer, BillWithParent from Customer where DisplayName like '" + escaped + "%' maxresults 100";
  return quickbooksApiGet_('/query?query=' + encodeURIComponent(query));
}

/**
 * Read-only diagnostic: every sub-customer/job (Job=true) in one shot, with
 * its real Id and ParentRef -- the bulk alternative to looking up jobs one
 * at a time via testQuickBooksCustomers. Projects created through QBO's
 * Projects tab do create an underlying Job=true Customer record (confirmed
 * empirically), just not always named predictably as "Parent:Child", so a
 * full export to cross-reference by name/parent against the Projects sheet
 * is more reliable than guessing a search prefix per job. Single page (up
 * to QBO's 1000-per-query cap) -- not paginated, since this is a manual
 * one-off lookup, not a cached/repeated call.
 */
function testQuickBooksAllJobs(payload) {
  var auth = authorizeQuickBooksOwner_(payload);
  if (!auth.ok) return { error: auth.error, code: auth.code };
  return quickbooksApiGet_('/query?query=' + encodeURIComponent('select Id, DisplayName, ParentRef from Customer where Job = true maxresults 1000'));
}

// ─── QBO Item catalog (Products/Services) + deterministic matching ──────────
var QBO_ITEM_CATALOG_CACHE_KEY = 'qbo_item_catalog_v1';
var QBO_ITEM_CATALOG_CACHE_TTL_SEC = 21600; // 6h, CacheService's max -- refreshQuickBooksItemCatalog forces a refetch on demand

/**
 * Fetches the full active Item catalog (Categories + the normalized
 * Products/Services nested under them, per the locked design) via the
 * existing read-only quickbooksApiGet_, paginating past QBO's 1000-row
 * MAXRESULTS cap if needed. Cached -- the review screen loads this on every
 * open, and re-fetching per load would be wasteful for data that only
 * changes when someone edits the QBO catalog.
 */
function getQuickBooksItemCatalog_() {
  var cache = CacheService.getScriptCache();
  try {
    var cached = cache.get(QBO_ITEM_CATALOG_CACHE_KEY);
    if (cached) return JSON.parse(cached);
  } catch (e) { /* fall through and refetch */ }

  var items = [];
  var startPosition = 1;
  var pageSize = 1000;
  while (true) {
    var query = "select Id, Name, Type, ParentRef, FullyQualifiedName, Active from Item where Active = true" +
      " startposition " + startPosition + " maxresults " + pageSize;
    var res = quickbooksApiGet_('/query?query=' + encodeURIComponent(query));
    if (!res.success) return res; // surface the error as-is; nothing to cache

    var page = (res.data && res.data.QueryResponse && res.data.QueryResponse.Item) || [];
    items = items.concat(page.map(function(it) {
      return {
        id: it.Id,
        name: it.Name,
        fullyQualifiedName: it.FullyQualifiedName,
        type: it.Type,
        parentId: it.ParentRef ? it.ParentRef.value : null
      };
    }));
    if (page.length < pageSize) break;
    startPosition += pageSize;
  }

  var result = { success: true, items: items };
  try { cache.put(QBO_ITEM_CATALOG_CACHE_KEY, JSON.stringify(result), QBO_ITEM_CATALOG_CACHE_TTL_SEC); } catch (e) { /* e.g. over 100KB -- fine, just skip caching this round */ }
  return result;
}

/** Thin wrapper exposing a manual "refresh catalog" action to the reviewer (invalidate + refetch). */
function refreshQuickBooksItemCatalog(payload) {
  var auth = authorizeInvoiceReviewer_(payload);
  if (!auth.ok) return { error: auth.error, code: auth.code };
  try { CacheService.getScriptCache().remove(QBO_ITEM_CATALOG_CACHE_KEY); } catch (e) {}
  return getQuickBooksItemCatalog_();
}

/**
 * Client-facing: cached catalog fetch for populating the review screen's
 * Item dropdown. Also resolves the optional QBO_TAX_ITEM_NAME/
 * QBO_FREIGHT_ITEM_NAME Script Properties (exact item name, e.g. "Sales
 * Tax") to their catalog Item Id -- irLoadCatalogAndMatch uses these to
 * pre-fill tax/freight lines the same way material lines get fuzzy-matched,
 * so the reviewer isn't re-picking the same tax/freight Item on every
 * invoice. '' (unset or not found in the catalog) means "leave unmatched,
 * same as before this existed" -- purely additive, no behavior change if
 * these properties are never set.
 */
function getQuickBooksItemCatalogForReview(payload) {
  var auth = authorizeInvoiceReviewer_(payload);
  if (!auth.ok) return { error: auth.error, code: auth.code };
  var catalog = getQuickBooksItemCatalog_();
  if (!catalog.success) return catalog;
  var props = PropertiesService.getScriptProperties();
  return {
    success: true,
    items: catalog.items,
    taxItemId: findQBOItemIdByName_(catalog.items, props.getProperty('QBO_TAX_ITEM_NAME')),
    freightItemId: findQBOItemIdByName_(catalog.items, props.getProperty('QBO_FREIGHT_ITEM_NAME'))
  };
}

/**
 * Client-facing: active Income + Expense/COGS + Inventory Asset accounts for
 * the "Add New QuickBooks Item" form's account pickers. IncomeAccountRef is a
 * required field on every Service/NonInventory/Inventory Item per QBO's
 * schema, even for an item that will only ever appear on Bills --
 * ExpenseAccountRef is optional but lets the new item categorize correctly on
 * the Bill it's about to be used on instead of falling back to QBO's default
 * expense account. assetAccounts (AccountSubType 'Inventory') is only used
 * when creating a stocked Inventory-type item -- see createQuickBooksItem.
 */
function getQuickBooksAccountsForNewItem(payload) {
  var auth = authorizeInvoiceReviewer_(payload);
  if (!auth.ok) return { error: auth.error, code: auth.code };
  var query = 'select Id, Name, AccountType, AccountSubType from Account where Active = true maxresults 1000';
  var res = quickbooksApiGet_('/query?query=' + encodeURIComponent(query));
  if (!res.success) return res;
  var accounts = ((res.data && res.data.QueryResponse && res.data.QueryResponse.Account) || []).map(function(a) {
    return { id: a.Id, name: a.Name, accountType: a.AccountType, accountSubType: a.AccountSubType };
  });
  var incomeTypes = { 'Income': true, 'Other Income': true };
  var expenseTypes = { 'Expense': true, 'Cost of Goods Sold': true, 'Other Expense': true };
  return {
    success: true,
    incomeAccounts: accounts.filter(function(a) { return incomeTypes[a.accountType]; }),
    expenseAccounts: accounts.filter(function(a) { return expenseTypes[a.accountType]; }),
    assetAccounts: accounts.filter(function(a) { return a.accountSubType === 'Inventory'; })
  };
}

/**
 * Creates a new Product/Service/Inventory Item directly in QuickBooks so an
 * invoice line that doesn't match anything in the catalog can be paired
 * without leaving the review screen. Service/NonInventory stay gated to the
 * regular admin+office reviewer group; Type: 'Inventory' is admin-only --
 * creating a stocked item means picking real GL accounts (Inventory Asset,
 * COGS), a bigger financial decision than matching an existing item, per the
 * locked "dummy proof" requirement. Invalidates the cached catalog on
 * success so the next full catalog load picks the new item up too; also
 * returns the item directly so the caller can splice it into an
 * already-loaded catalog without waiting on a refetch.
 */
function createQuickBooksItem(payload) {
  var type = payload.type === 'NonInventory' ? 'NonInventory' : (payload.type === 'Inventory' ? 'Inventory' : 'Service');
  var auth = type === 'Inventory' ? authorizeCaller(payload, ['admin']) : authorizeInvoiceReviewer_(payload);
  if (!auth.ok) return { success: false, error: auth.error, code: auth.code };

  var name = (payload.name || '').toString().trim();
  if (!name) return { success: false, error: 'Item name is required.' };
  var incomeAccountId = (payload.incomeAccountId || '').toString().trim();
  if (!incomeAccountId) return { success: false, error: 'Income account is required.' };

  var body = {
    Name: name,
    Type: type,
    IncomeAccountRef: { value: incomeAccountId }
  };
  if (payload.expenseAccountId) body.ExpenseAccountRef = { value: payload.expenseAccountId.toString() };
  if (payload.parentId) {
    body.SubItem = true;
    body.ParentRef = { value: payload.parentId.toString() };
  }

  if (type === 'Inventory') {
    var assetAccountId = (payload.assetAccountId || '').toString().trim();
    if (!assetAccountId) return { success: false, error: 'Inventory Asset account is required.' };
    if (!payload.expenseAccountId) return { success: false, error: 'COGS/Expense account is required for an Inventory item.' };
    // Bootstrapped at 0 on-hand, per the locked design -- the receiving Bill
    // that follows right after is what establishes the real FIFO layer at
    // the vendor's actual cost, so starting at the invoice quantity here
    // would double-count it.
    body.TrackQtyOnHand = true;
    body.QtyOnHand = 0;
    body.InvStartDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    body.AssetAccountRef = { value: assetAccountId };
  }

  var postRes = quickbooksApiPost_('/item', body);
  if (!postRes.success) return { success: false, error: postRes.error };

  var item = postRes.data && postRes.data.Item;
  if (!item) return { success: false, error: 'Unexpected response creating item.' };

  try { CacheService.getScriptCache().remove(QBO_ITEM_CATALOG_CACHE_KEY); } catch (e) { /* best-effort invalidation */ }

  return {
    success: true,
    item: {
      id: item.Id,
      name: item.Name,
      fullyQualifiedName: item.FullyQualifiedName,
      type: item.Type,
      parentId: item.ParentRef ? item.ParentRef.value : null
    }
  };
}

/** Resolves a Script-Property-configured item name (case/punctuation-insensitive exact match, via the same normalizer as the fuzzy matcher below) to its QBO Item Id, or '' if unset or not found. */
function findQBOItemIdByName_(items, name) {
  if (!name) return '';
  var wantNorm = qboNormalizeItemText_(name);
  var match = (items || []).find(function(it) { return it.type !== 'Category' && qboNormalizeItemText_(it.name) === wantNorm; });
  return match ? match.id : '';
}

/** Lowercases, strips punctuation, collapses whitespace, and normalizes common size notation so "8-inch"/"8 in"/"8\"" compare equal. */
function qboNormalizeItemText_(s) {
  return (s || '').toString().toLowerCase()
    .replace(/(\d+)\s*(?:"|in\b|inch(?:es)?\b)/g, '$1in')
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
}

/**
 * Deterministic string matcher -- no AI/ML, per the locked "never guess"
 * requirement; this needs to stay inspectable. Scores each catalog item by
 * token-overlap ratio against the extracted description, with a bonus for
 * one string containing the other. Returns the best match with its score,
 * or null if nothing scores above MATCH_THRESHOLD -- callers treat that as
 * "unmatched, flag for manual pairing."
 */
var QBO_MATCH_THRESHOLD = 0.55;

function matchLineItemToQBOItem_(description, catalogItems) {
  var wantNorm = qboNormalizeItemText_(description);
  if (!wantNorm) return null;
  var wantTokens = wantNorm.split(' ').filter(Boolean);

  var best = null;
  (catalogItems || []).forEach(function(item) {
    // Only leaf Items are billable lines on a Bill -- QBO Categories
    // (Type: 'Category') are parents for rollup reporting, not selectable
    // as an ItemRef themselves.
    if (item.type === 'Category') return;

    var itemNorm = qboNormalizeItemText_(item.name);
    var itemTokens = itemNorm.split(' ').filter(Boolean);
    if (!itemTokens.length) return;

    var overlap = wantTokens.filter(function(t) { return itemTokens.indexOf(t) !== -1; }).length;
    var score = overlap / Math.max(wantTokens.length, itemTokens.length);
    if (itemNorm && (wantNorm.indexOf(itemNorm) !== -1 || itemNorm.indexOf(wantNorm) !== -1)) {
      score = Math.max(score, 0.75); // substring-containment bonus, e.g. "lp 8in lap siding" contains "8in lap siding"
    }

    if (!best || score > best.score) best = { item: item, score: score };
  });

  if (!best || best.score < QBO_MATCH_THRESHOLD) return null;
  return { qboItemId: best.item.id, qboItemName: best.item.name, matchConfidence: Math.round(best.score * 100) / 100 };
}

/**
 * Loads the full learned item-mapping sheet (QB_ITEM_MAP_SHEET, written by
 * saveQBItemMapping_ in PO_Manager_Code.gs on Approve) in one read: keyed by
 * normalized description + "|" + lane ('stock' | 'direct') -> {qboItemId,
 * qboItemName}. Lane-keyed rather than description-only so a material that's
 * sometimes bought direct-for-a-job and sometimes into stock can hold both
 * mappings without one overwriting the other. Returns {} (never throws) if
 * the sheet is empty or doesn't exist yet -- matchInvoiceLineItems then
 * falls through to the fuzzy matcher exactly as it did before this existed.
 */
function getQBItemMap_() {
  var map = {};
  try {
    var sheet = ensureSheetWithHeaders_(QB_ITEM_MAP_SHEET, QB_ITEM_MAP_HEADERS);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return map;
    var data = sheet.getRange(2, 1, lastRow - 1, QB_ITEM_MAP_HEADERS.length).getValues();
    data.forEach(function(row) {
      var key = (row[0] || '').toString().trim();
      var qboItemId = (row[2] || '').toString().trim();
      var lane = (row[4] || 'direct').toString().trim() || 'direct';
      if (!key || !qboItemId) return;
      map[key + '|' + lane] = { qboItemId: qboItemId, qboItemName: (row[3] || '').toString().trim() };
    });
  } catch (e) { /* best-effort, same fail-open convention as getQuickBooksVendorId_ */ }
  return map;
}

/**
 * Client-facing: matches a batch of line-item descriptions against the
 * learned item-mapping sheet first, falling back to the cached QBO catalog
 * fuzzy matcher on a miss (used when the review screen opens a staging
 * row). payload: {descriptions: string[], jobRef?: string, learnedOnly?:
 * bool}. Returns one match (or null) per input description, same order.
 *
 * jobRef determines the lane ('stock' when it's the WAREHOUSE placeholder,
 * 'direct' otherwise) -- both the learned-map lookup and the fuzzy-match
 * fallback are lane-scoped. Inventory-type QBO Items are never offered
 * outside the 'stock' lane: an Inventory item accidentally attached to a
 * real job's Bill would silently misbook that cost into Inventory Asset
 * instead of the job's COGS (QBO doesn't move inventory cost to a job just
 * from a CustomerRef tag), so this is a hard exclusion, not a preference --
 * see the matching hard-stop in createQuickBooksBill for the
 * can't-be-bypassed-by-the-client backstop.
 *
 * learnedOnly (used for tax/freight lines) skips the fuzzy-matcher fallback
 * entirely -- scoring "Sales Tax" by token overlap against a materials
 * catalog is just noise, per the locked "never guess" requirement. The
 * learned-map lookup itself is a safe exact normalized-description match
 * either way, so tax/freight lines still benefit from it: pick "Sales Tax"
 * once, approve, and it's remembered for next time instead of needing to be
 * repicked (or relying solely on the QBO_TAX_ITEM_NAME/QBO_FREIGHT_ITEM_NAME
 * Script Properties, which this now sits alongside as a fallback for).
 */
function matchInvoiceLineItems(payload) {
  var auth = authorizeInvoiceReviewer_(payload);
  if (!auth.ok) return { error: auth.error, code: auth.code };

  var catalogRes = getQuickBooksItemCatalog_();
  if (!catalogRes.success) return { error: catalogRes.error || 'Could not load QuickBooks Item catalog' };

  var lane = isWarehouseJob_(payload.jobRef) ? 'stock' : 'direct';
  var matchableItems = lane === 'stock' ? catalogRes.items : catalogRes.items.filter(function(it) { return it.type !== 'Inventory'; });

  var learnedMap = getQBItemMap_();
  var descriptions = payload.descriptions || [];
  var learnedOnly = !!payload.learnedOnly;
  var matches = descriptions.map(function(d) {
    var key = qboNormalizeItemText_(d);
    var learned = key && learnedMap[key + '|' + lane];
    if (learned) return { qboItemId: learned.qboItemId, qboItemName: learned.qboItemName, matchConfidence: 1 };
    return learnedOnly ? null : matchLineItemToQBOItem_(d, matchableItems);
  });
  return { success: true, matches: matches };
}

/**
 * Uploads a file as a QuickBooks Attachable linked to a specific entity
 * (e.g. a just-created Bill) via AttachableRef -- this is what makes the
 * invoice PDF show up as an attachment on the Bill inside QuickBooks
 * itself, not just linked from our own Drive copy. QBO's Attachable API is
 * a dedicated multipart endpoint (/upload), separate from the JSON-only
 * quickbooksApiPost_ above, so it needs its own request shape: a JSON
 * metadata part (AttachableRef) plus the raw file bytes.
 */
function quickbooksUploadAttachment_(fileBlob, entityType, entityId) {
  var service = getQuickBooksService_();
  if (!service.hasAccess()) {
    return { success: false, error: 'QuickBooks is not connected yet.' };
  }
  var realmId = PropertiesService.getScriptProperties().getProperty('QBO_REALM_ID');
  if (!realmId) {
    return { success: false, error: 'Missing QuickBooks company (realm) id - reconnect QuickBooks.' };
  }

  var metadata = {
    AttachableRef: [{
      EntityRef: { type: entityType, value: entityId.toString() },
      IncludeOnSend: false
    }],
    FileName: fileBlob.getName(),
    ContentType: fileBlob.getContentType()
  };
  var metadataBlob = Utilities.newBlob(JSON.stringify(metadata), 'application/json', 'metadata.json');

  var url = getQuickBooksBaseUrl_() + '/v3/company/' + realmId + '/upload';
  var response = UrlFetchApp.fetch(url, {
    method: 'post',
    headers: {
      Authorization: 'Bearer ' + service.getAccessToken(),
      Accept: 'application/json'
    },
    // UrlFetchApp builds a multipart/form-data body automatically when
    // payload is an object of Blobs -- one part per key, named after the
    // key, per Intuit's documented two-part (metadata + content) contract.
    payload: {
      file_metadata_01: metadataBlob,
      file_content_01: fileBlob
    },
    muteHttpExceptions: true
  });

  var code = response.getResponseCode();
  var body = response.getContentText();
  if (code < 200 || code >= 300) {
    return { success: false, error: 'QuickBooks attachment upload error ' + code + ': ' + qboParseErrorBody_(body) };
  }
  var parsed;
  try { parsed = JSON.parse(body); } catch (e) { return { success: false, error: 'Unexpected attachment response: ' + body }; }
  var entry = parsed.AttachableResponse && parsed.AttachableResponse[0];
  if (!entry || entry.Fault) {
    return { success: false, error: qboParseErrorBody_(body) };
  }
  return { success: true, attachableId: entry.Attachable && entry.Attachable.Id };
}

/**
 * For a WAREHOUSE-lane (stock) invoice, folds freight/delivery-fee AND tax
 * line amounts into the material lines' Amount pro-rata by dollar share,
 * instead of posting them as their own Bill lines -- standard freight-in
 * capitalization: a QBO inventory line's FIFO cost is just Amount/Qty, and
 * there's no separate "freight" field on an inventory line, so getting an
 * accurate landed cost per unit means folding it in before posting rather
 * than expensing it separately. Sales-tax treatment on stock receipts is
 * still an open question with whoever does the books (freight is the
 * confirmed case) -- tax is folded the same way as freight here as the
 * simplest consistent Phase-1 default, not because it's confirmed correct;
 * revisit if the bookkeeper wants tax kept separate.
 *
 * Returns a NEW array of material lines only (shallow-copied, originals
 * untouched) -- freight/tax lines never make it into the returned list, so
 * callers building Bill lines from this never see them. Direct-to-job
 * invoices never call this; their freight/tax lines keep posting as their
 * own Bill line exactly as before.
 */
function applyWarehouseNonMaterialAllocation_(lineItems) {
  var materialLines = (lineItems || [])
    .filter(function(li) { return !li.skip && li.lineType !== 'freight' && li.lineType !== 'tax'; })
    .map(function(li) { return Object.assign({}, li); });

  var extraTotal = (lineItems || [])
    .filter(function(li) { return !li.skip && (li.lineType === 'freight' || li.lineType === 'tax'); })
    .reduce(function(s, li) { return s + (parseFloat(li.amount) || 0); }, 0);
  if (!extraTotal || !materialLines.length) return materialLines;

  var materialTotal = materialLines.reduce(function(s, li) { return s + (parseFloat(li.amount) || 0); }, 0);
  if (!materialTotal) return materialLines; // nothing sensible to allocate a $ share against

  materialLines.forEach(function(li) {
    var share = (parseFloat(li.amount) || 0) / materialTotal;
    var added = Math.round(extraTotal * share * 100) / 100;
    li.amount = Math.round(((parseFloat(li.amount) || 0) + added) * 100) / 100;
    var qty = parseFloat(li.qty);
    if (qty) li.rate = Math.round((li.amount / qty) * 10000) / 10000;
  });
  return materialLines;
}

// ─── QuickBooks Bill creation (write path) ───────────────────────────────────
/**
 * Creates a Bill in QuickBooks from an Approved staging row. Owner-gated,
 * same as the rest of QuickBooks access -- Bill creation was explicitly
 * kept owner-only rather than opened to admin/office, per the locked
 * decision. Reloads the staging row server-side rather than trusting
 * whatever line items the client submits, since this posts real financial
 * data. Idempotent: a staging row that already has a QB Bill Id returns
 * that Bill instead of posting again, covering retries/double-clicks. Also
 * attaches the uploaded invoice PDF to the new Bill (see
 * quickbooksUploadAttachment_) -- best-effort, reported via
 * attachmentWarning on the response rather than failing the whole call.
 */
function createQuickBooksBill(payload) {
  var auth = authorizeInvoiceReviewer_(payload);
  if (!auth.ok) return { success: false, error: auth.error, code: auth.code };

  var stagingId = payload.stagingId;
  if (!stagingId) return { success: false, error: 'Missing stagingId' };

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { success: false, error: 'Server is busy - try again in a moment.' };

  try {
    var sheet = ensureSheetWithHeaders_(QB_STAGING_SHEET, QB_STAGING_HEADERS);
    var rowIdx = findStagingRowIndex_(sheet, stagingId);
    if (rowIdx === -1) return { success: false, error: 'Staging row not found' };

    var rowValues = sheet.getRange(rowIdx, 1, 1, QB_STAGING_HEADERS.length).getValues()[0];
    var staging = stagingRowToObject_(rowValues);

    // Idempotency: already posted, hand back the existing Bill rather than
    // creating a duplicate. Covers a retried click or a repeated call.
    if (staging.status === 'Posted' && staging.qbBillId) {
      return { success: true, qbBillId: staging.qbBillId, alreadyPosted: true };
    }
    if (staging.status !== 'Approved') {
      return { success: false, error: 'This invoice must be Approved before a Bill can be created (currently: ' + staging.status + ').' };
    }
    if (!staging.qbVendorId) {
      return { success: false, error: 'No QuickBooks Vendor Id linked for "' + staging.vendor + '" -- add it in the QB Vendor Map before creating this Bill.' };
    }
    // Stock (WAREHOUSE) POs receive with no job attached -- the whole point
    // is that the material isn't attributed to a job yet -- so the usual
    // Customer/Job requirement is skipped for them. Any other job still
    // requires one exactly as before.
    var isWarehouse = isWarehouseJob_(staging.jobRef);
    if (!isWarehouse && !staging.qbCustomerId) {
      return { success: false, error: 'No QuickBooks Customer/Job Id linked for "' + staging.builder + ' / ' + staging.jobRef + '" -- add it to the Projects sheet before creating this Bill.' };
    }

    // Server-side re-validation of every line's QBO Item Type, right before
    // posting -- never trusts whatever the client submitted, since this
    // posts real financial data. This is the hard backstop for the
    // "Inventory items can never attach to a non-warehouse job" guardrail:
    // QBO doesn't move inventory cost to a job's COGS just from a
    // CustomerRef tag, so an Inventory item slipping onto a real job's Bill
    // would silently misbook that cost into Inventory Asset instead.
    var catalogRes = getQuickBooksItemCatalog_();
    if (!catalogRes.success) return { success: false, error: catalogRes.error || 'Could not load QuickBooks Item catalog to validate this Bill.' };
    var itemTypeById = {};
    catalogRes.items.forEach(function(it) { itemTypeById[it.id] = it.type; });

    var billableLines = isWarehouse ? applyWarehouseNonMaterialAllocation_(staging.lineItems) : (staging.lineItems || []).filter(function(li) { return !li.skip; });

    var badLine = null;
    billableLines.forEach(function(li) {
      if (badLine || !li.qboItemId) return;
      var isInventoryItem = itemTypeById[li.qboItemId] === 'Inventory';
      if (isWarehouse && !isInventoryItem) badLine = li;
      if (!isWarehouse && isInventoryItem) badLine = li;
    });
    if (badLine) {
      return {
        success: false,
        error: isWarehouse
          ? 'Line "' + (badLine.description || '(no description)') + '" is not a stocked Inventory item -- a stock PO can only bill Inventory-type items.'
          : 'Line "' + (badLine.description || '(no description)') + '" is a stocked Inventory item, which can\'t be billed to a real job -- fix the item pairing before creating this Bill.'
      };
    }

    var billLines = [];
    var missingItem = null;
    billableLines.forEach(function(li) {
      if (!li.qboItemId) { missingItem = li; return; }
      var lineDetail = {
        ItemRef: { value: li.qboItemId },
        Qty: li.qty !== '' && li.qty != null ? parseFloat(li.qty) : undefined,
        UnitPrice: li.rate !== '' && li.rate != null ? parseFloat(li.rate) : undefined
      };
      if (!isWarehouse) {
        // CustomerRef alone attributes the cost to the job for job-cost
        // reporting without exposing it as a pass-through charge to be
        // re-invoiced to the customer later -- NotBillable set explicitly
        // (not omitted) so this stays deliberate, not QBO's ambient default.
        lineDetail.CustomerRef = { value: staging.qbCustomerId };
        lineDetail.BillableStatus = 'NotBillable';
      }
      billLines.push({
        DetailType: 'ItemBasedExpenseLineDetail',
        Amount: parseFloat(li.amount) || 0,
        Description: li.description || '',
        ItemBasedExpenseLineDetail: lineDetail
      });
    });
    if (missingItem) {
      return { success: false, error: 'Line "' + (missingItem.description || '(no description)') + '" has no matched QuickBooks Item -- pair or skip it before approving/creating the Bill.' };
    }
    if (!billLines.length) {
      return { success: false, error: 'No line items to bill (everything is skipped).' };
    }

    var billPayload = {
      VendorRef: { value: staging.qbVendorId },
      // A blank vendorInvoice used to leave this key entirely absent,
      // which fails with "value must not be null : DocNumber" on a QBO
      // company that has custom transaction numbers enabled (confirmed).
      // Falls back to a timestamp-based number rather than leaving it unset.
      DocNumber: staging.vendorInvoice || ('BILL-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyMMdd-HHmmss')),
      Line: billLines
    };

    var postRes = quickbooksApiPost_('/bill', billPayload);
    if (!postRes.success) {
      // Leave Status at 'Approved' (not reverted) so this can be fixed and retried without redoing the whole review.
      // sentPayload included for debugging "Invalid Reference Id" faults --
      // shows the exact JSON QuickBooks received, since a rejected reference
      // can look valid everywhere except in the literal request bytes.
      return { success: false, error: postRes.error, sentPayload: billPayload };
    }

    var qbBillId = postRes.data && postRes.data.Bill && postRes.data.Bill.Id;
    var postedAt = new Date();
    sheet.getRange(rowIdx, QB_STAGING_COL['Status'] + 1).setValue('Posted');
    sheet.getRange(rowIdx, QB_STAGING_COL['QB Bill Id'] + 1).setValue(qbBillId || '');
    sheet.getRange(rowIdx, QB_STAGING_COL['Posted At'] + 1).setValue(postedAt);

    // Best-effort: attach the same invoice PDF the reviewer uploaded (already
    // sitting in Drive at staging.invoiceFileUrl) to the Bill we just created,
    // so the file is visible directly on the Bill in QuickBooks. Non-fatal --
    // the Bill itself is already posted and Approved->Posted at this point,
    // so a Drive/attachment hiccup surfaces as a warning, not a failure.
    var attachmentWarning;
    if (qbBillId && staging.invoiceFileUrl) {
      try {
        var invoiceFileId = extractDriveFileId_(staging.invoiceFileUrl);
        if (!invoiceFileId) {
          attachmentWarning = 'Bill ' + qbBillId + ' created, but the invoice file link could not be read to attach it.';
        } else {
          var attachRes = quickbooksUploadAttachment_(DriveApp.getFileById(invoiceFileId).getBlob(), 'Bill', qbBillId);
          if (!attachRes.success) {
            attachmentWarning = 'Bill ' + qbBillId + ' created, but attaching the invoice PDF failed: ' + attachRes.error;
          }
        }
      } catch (attachErr) {
        attachmentWarning = 'Bill ' + qbBillId + ' created, but attaching the invoice PDF failed: ' + attachErr.toString();
      }
    }

    // Best-effort: flatten this Bill's line items into the Purchase Line Item
    // Log for later analytics (PPV, per-vendor price trends, per-material job
    // cost) -- see logPurchaseLineItems_ in PO_Manager_Code.gs. Non-fatal,
    // same reasoning as the attachment step above: the Bill has already
    // posted, so a logging hiccup surfaces as a warning, not a failure.
    var logWarning = logPurchaseLineItems_(staging, qbBillId, auth.email, postedAt);

    return { success: true, qbBillId: qbBillId, attachmentWarning: attachmentWarning, logWarning: logWarning || undefined };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}
