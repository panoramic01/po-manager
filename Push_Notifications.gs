// === Push Notifications (Web Push via FCM HTTP v1) ===========================
// Sends real push notifications (arrive even when the app/tab is closed) via
// Firebase Cloud Messaging's HTTP v1 send API. Authenticated with this
// script's own ScriptApp.getOAuthToken() (requires the
// https://www.googleapis.com/auth/firebase.messaging scope in appsscript.json)
// rather than a downloaded service-account key -- Firebase blocks key
// creation by default on new projects, and the script's own OAuth identity
// (Aidan, who also owns the Firebase project) already has send permission.

// Temporary kill switch for the whole notification system (push + the
// automatic emails in Code.gs / Form_Response.gs / Received_Form_Response.gs)
// while it's paused pending a redesign. Flip back to true to re-enable --
// nothing else needs to change. Checked in sendPushNotification() below and
// at each MailApp/GmailApp.sendEmail call site guarded by this same flag.
var NOTIFICATIONS_ENABLED = false;

var FCM_PROJECT_ID = 'panoramic-ops-push';
var PUSH_SHEET = 'Push Subscriptions';
var PUSH_SHEET_HEADERS = ['Email', 'FCM Token', 'Device Info', 'Created At', 'Last Seen At'];

function getOrCreatePushSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(PUSH_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(PUSH_SHEET);
    sheet.appendRow(PUSH_SHEET_HEADERS);
  }
  return sheet;
}

/**
 * Registers (or refreshes) an FCM token for the calling user's device.
 * Any signed-in user may manage their own push subscription -- no role
 * restriction here, since who can even see the "enable push" toggle is
 * already gated client-side.
 */
function registerPushToken(payload) {
  var auth = requireVerifiedEmail_(payload);
  if (auth.error) return { success: false, error: auth.error, code: auth.code };

  var token = ((payload && payload.fcmToken) || '').toString().trim();
  if (!token) return { success: false, error: 'Missing push token.' };
  var deviceInfo = ((payload && payload.deviceInfo) || '').toString().slice(0, 300);

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { success: false, error: 'Server is busy - try again in a moment.' };
  try {
    var sheet = getOrCreatePushSheet_();
    var data = sheet.getDataRange().getValues();
    var now = new Date();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === auth.email && data[i][1] === token) {
        sheet.getRange(i + 1, 5).setValue(now); // Last Seen At
        return { success: true };
      }
    }
    sheet.appendRow([auth.email, token, deviceInfo, now, now]);
    return { success: true };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Removes an FCM token for the calling user's device (push disabled on that
 * device). Any signed-in user may remove their own token.
 */
function unregisterPushToken(payload) {
  var auth = requireVerifiedEmail_(payload);
  if (auth.error) return { success: false, error: auth.error, code: auth.code };

  var token = ((payload && payload.fcmToken) || '').toString().trim();
  if (!token) return { success: false, error: 'Missing push token.' };

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { success: false, error: 'Server is busy - try again in a moment.' };
  try {
    var sheet = getOrCreatePushSheet_();
    var data = sheet.getDataRange().getValues();
    for (var i = data.length - 1; i >= 1; i--) {
      if (data[i][0] === auth.email && data[i][1] === token) sheet.deleteRow(i + 1);
    }
    return { success: true };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Deletes a push-subscription row by its FCM token, regardless of owner --
 * used to self-clean tokens FCM reports as dead/unregistered so the sheet
 * doesn't accumulate stale rows for uninstalled/reinstalled devices.
 */
function deletePushTokenRow_(token) {
  try {
    var sheet = getOrCreatePushSheet_();
    var data = sheet.getDataRange().getValues();
    for (var i = data.length - 1; i >= 1; i--) {
      if (data[i][1] === token) sheet.deleteRow(i + 1);
    }
  } catch (e) {
    // best-effort cleanup only
  }
}

/**
 * Sends a push notification to one or more users (by email), to every
 * device/token they currently have registered. Uses a data-only FCM
 * payload (no top-level "notification" key) so the service worker's own
 * `push` event handler always fires and renders the notification itself,
 * instead of the browser auto-rendering FCM's default notification.
 *
 * Silently no-ops for any email with no registered token -- expected right
 * now since only the 'aidan' role can enable push. Callers can call this
 * inline the same way MailApp.sendEmail(...) is called elsewhere; it never
 * throws.
 *
 * Returns { sent, failed }.
 */
function sendPushNotification(targetEmails, title, body, url) {
  if (!NOTIFICATIONS_ENABLED) return { sent: 0, failed: 0 };
  var emails = Array.isArray(targetEmails) ? targetEmails : [targetEmails];
  emails = emails.map(function(e) { return (e || '').toString().toLowerCase().trim(); }).filter(Boolean);
  if (!emails.length) return { sent: 0, failed: 0 };

  var sheet = getOrCreatePushSheet_();
  var data = sheet.getDataRange().getValues();
  var tokensByEmail = {};
  for (var i = 1; i < data.length; i++) {
    var rowEmail = (data[i][0] || '').toString().toLowerCase().trim();
    if (emails.indexOf(rowEmail) === -1) continue;
    if (!tokensByEmail[rowEmail]) tokensByEmail[rowEmail] = [];
    tokensByEmail[rowEmail].push(data[i][1]);
  }

  var accessToken;
  try {
    accessToken = ScriptApp.getOAuthToken();
  } catch (e) {
    logError_('sendPushNotification', 'Could not get OAuth token: ' + e.toString(), { targetEmails: emails });
    return { sent: 0, failed: emails.length };
  }

  // One entry per (email, token) pair, built up front so the FCM calls
  // below can go out as a single parallel batch instead of one-by-one --
  // sequential UrlFetchApp.fetch() calls here used to sit directly in the
  // critical path of createPO/createSubPO (every recipient x every device),
  // which could push a PO submission's total response time past the
  // client's fetch timeout even though the PO row itself had already been
  // written. fetchAll() dispatches every request concurrently, so total
  // latency is roughly that of the single slowest recipient instead of the
  // sum of all of them.
  var jobs = [];
  emails.forEach(function(email) {
    (tokensByEmail[email] || []).forEach(function(token) {
      jobs.push({ email: email, token: token });
    });
  });
  if (!jobs.length) return { sent: 0, failed: 0 };

  var requests = jobs.map(function(job) {
    return {
      url: 'https://fcm.googleapis.com/v1/projects/' + FCM_PROJECT_ID + '/messages:send',
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + accessToken },
      payload: JSON.stringify({
        message: {
          token: job.token,
          data: {
            title: title || 'Panoramic Ops',
            body: body || '',
            url: url || '/'
          }
        }
      }),
      muteHttpExceptions: true
    };
  });

  var sent = 0, failed = 0;
  var responses;
  try {
    responses = UrlFetchApp.fetchAll(requests);
  } catch (e) {
    logError_('sendPushNotification', 'fetchAll failed: ' + e.toString(), { targetEmails: emails });
    return { sent: 0, failed: jobs.length };
  }

  responses.forEach(function(resp, i) {
    var job = jobs[i];
    var code = resp.getResponseCode();
    if (code === 200) {
      sent++;
    } else {
      failed++;
      var errText = resp.getContentText();
      if (code === 404 || errText.indexOf('UNREGISTERED') !== -1 || errText.indexOf('INVALID_ARGUMENT') !== -1) {
        deletePushTokenRow_(job.token);
      }
      logError_('sendPushNotification', 'FCM send failed (' + code + '): ' + errText, { email: job.email });
    }
  });

  return { sent: sent, failed: failed };
}
