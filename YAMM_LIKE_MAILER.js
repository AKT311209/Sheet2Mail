function doGet() {
  return HtmlService.createHtmlOutputFromFile('YAMM_Like_Mailer_UI')
    .setWidth(700)
    .setHeight(800);
}

/**
 * Returns the current user's email for the sender field.
 */
function getCurrentUserEmail() {
  return Session.getActiveUser().getEmail();
}

/**
 * Returns the user's remaining daily Gmail sending quota.
 */
function getEmailQuota() {
  return MailApp.getRemainingDailyQuota();
}

/**
 * Given a Google Sheet ID or URL, return its sheets and headers.
 */
function getSheetsAndHeaders(sheetIdOrUrl) {
  var ss;
  if (sheetIdOrUrl.startsWith('http')) {
    ss = SpreadsheetApp.openByUrl(sheetIdOrUrl);
  } else {
    ss = SpreadsheetApp.openById(sheetIdOrUrl);
  }
  var sheets = ss.getSheets();
  var result = [];
  sheets.forEach(function (sheet) {
    var headers = sheet.getDataRange().getValues()[0];
    result.push({
      name: sheet.getName(),
      headers: headers
    });
  });
  return result;
}

/**
 * Send emails based on config, using the Sheet ID or URL provided.
 * Supports masking the sender with a display name if provided.
 * Set config.test=true to send a test mail to user using testValue for each mapping.
 * If config.scheduleDateTime is set (ISO string), schedule the send using Apps Script triggers.
 * @param {Object} config
 * {
 *   sheetIdOrUrl: string,
 *   sheetName: string,
 *   startRow: number, // NEW: 1-based, inclusive
 *   endRow: number,   // NEW: 1-based, inclusive
 *   emailCol: string,
 *   subject: string,
 *   body: string,
 *   mappings: [{placeholder: string, column: string, testValue: string}],
 *   fromName: string,
 *   fromEmail: string,
 *   test: boolean,
 *   scheduleDateTime: string (optional, ISO 8601)
 * }
 */
function sendMailMerge(config) {
  if (config.scheduleDateTime && !config.test) {
    // Schedule the sendMailMerge to run at the specified time
    var dt = new Date(config.scheduleDateTime);
    if (isNaN(dt.getTime())) throw new Error('Invalid schedule date/time.');
    // Remove scheduleDateTime to avoid recursion
    var configCopy = JSON.parse(JSON.stringify(config));
    delete configCopy.scheduleDateTime;

    // Basic safeguard: avoid creating too many triggers for this project
    try {
      var existingTriggers = ScriptApp.getProjectTriggers();
      // Typical Apps Script limit is ~20 time-based triggers per project; use 18 as a soft threshold
      if (existingTriggers && existingTriggers.length >= 18) {
        throw new Error('Too many triggers exist for this script. Please remove old scheduled tasks before creating new ones.');
      }
    } catch (err) {
      // If ScriptApp.getProjectTriggers fails for any reason, continue and let trigger creation attempt
    }

    var trigger = null;
    try {
      trigger = ScriptApp.newTrigger('sendMailMergeScheduled')
        .timeBased()
        .at(dt)
        .create();
    } catch (e) {
      throw new Error('Failed to create scheduled trigger: ' + (e && e.message ? e.message : e));
    }

    // Try to store the config keyed by the trigger's unique id for reliable lookup
    try {
      var triggerUid = (trigger && typeof trigger.getUniqueId === 'function') ? trigger.getUniqueId() : null;
      var props = PropertiesService.getScriptProperties();
      if (triggerUid) {
        var key = 'SCHEDULED_MAILMERGE_' + triggerUid;
        var payload = {
          config: configCopy,
          scheduledFor: dt.getTime(),
          createdBy: (Session && Session.getActiveUser && Session.getActiveUser().getEmail) ? Session.getActiveUser().getEmail() : null
        };
        props.setProperty(key, JSON.stringify(payload));
      } else {
        // Fallback: store by timestamp (legacy behaviour) but include scheduledFor
        var keyTs = 'SCHEDULED_MAILMERGE_' + dt.getTime();
        var payload = {
          config: configCopy,
          scheduledFor: dt.getTime(),
          createdBy: (Session && Session.getActiveUser && Session.getActiveUser().getEmail) ? Session.getActiveUser().getEmail() : null
        };
        PropertiesService.getScriptProperties().setProperty(keyTs, JSON.stringify(payload));
      }
    } catch (e) {
      // If saving config fails, delete the trigger to avoid orphan triggers and surface error
      try { if (trigger) ScriptApp.deleteTrigger(trigger); } catch (ex) { }
      throw new Error('Failed to store scheduled configuration: ' + (e && e.message ? e.message : e));
    }

    // Format date as dd/mm/yy for user-friendly message
    var day = dt.getDate().toString().padStart(2, '0');
    var month = (dt.getMonth() + 1).toString().padStart(2, '0');
    var year = dt.getFullYear().toString().slice(-2);
    var formatted = day + '/' + month + '/' + year;
    var time = dt.toTimeString().slice(0, 8); // HH:MM:SS
    return 'Scheduled email send for ' + formatted + ' ' + time + '.';
  }

  if (config.test) {
    // Send test email using test values
    var subject = config.subject;
    var body = config.body;
    var map = {};
    (config.mappings || []).forEach(function (m) {
      if (m.placeholder) {
        map[m.placeholder] = m.testValue || '';
      }
    });
    for (var p in map) {
      subject = subject.replace(new RegExp('{{' + p + '}}', 'g'), map[p]);
      body = body.replace(new RegExp('{{' + p + '}}', 'g'), map[p]);
    }
    var mailOptions = {
      to: config.fromEmail || Session.getActiveUser().getEmail(),
      subject: subject,
      htmlBody: body
    };
    if (config.fromName) {
      mailOptions.name = config.fromName;
    }
    MailApp.sendEmail(mailOptions);
    return 'Test email sent to ' + mailOptions.to + '!';
  }

  // ...existing mail merge logic...
  var ss;
  if (config.sheetIdOrUrl.startsWith('http')) {
    ss = SpreadsheetApp.openByUrl(config.sheetIdOrUrl);
  } else {
    ss = SpreadsheetApp.openById(config.sheetIdOrUrl);
  }
  var sheet = ss.getSheetByName(config.sheetName);
  if (!sheet) {
    throw new Error('Sheet "' + config.sheetName + '" not found. Please check your sheet selection.');
  }
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  if (!headers) {
    throw new Error('No header row found in the selected sheet.');
  }
  var colMap = {};
  headers.forEach(function (h, idx) {
    colMap[h] = idx;
  });

  if (!(config.emailCol in colMap)) {
    throw new Error('Email column "' + config.emailCol + '" not found in the selected sheet.');
  }

  var map = {};
  (config.mappings || []).forEach(function (m) {
    if (m.placeholder && m.column && (m.column in colMap)) {
      map[m.placeholder] = colMap[m.column];
    }
  });

  // Use startRow and endRow (1-based, including header row as row 1)
  var startRow = Math.max(2, config.startRow || 2); // default to 2 (first data row)
  var endRow = Math.min(data.length, config.endRow || data.length);
  var sentCount = 0;
  var firstRow = null;
  var lastRow = null;
  var errorMsg = null;
  for (var i = startRow - 1; i < endRow; i++) {
    var row = data[i];
    var email = row[colMap[config.emailCol]];
    if (!email) continue;
    if (firstRow === null) firstRow = i + 1; // +1 for 1-based row number
    lastRow = i + 1;
    var subject = config.subject;
    var body = config.body;
    for (var p in map) {
      var value = row[map[p]];
      subject = subject.replace(new RegExp('{{' + p + '}}', 'g'), value);
      body = body.replace(new RegExp('{{' + p + '}}', 'g'), value);
    }
    var mailOptions = {
      to: email,
      subject: subject,
      htmlBody: body
    };
    if (config.fromName) {
      mailOptions.name = config.fromName;
    }
    try {
      MailApp.sendEmail(mailOptions);
      sentCount++;
    } catch (e) {
      errorMsg = e && e.message ? e.message : 'Unknown error';
      break;
    }
  }
  if (sentCount === 0) {
    return 'No emails sent (no valid recipients in the selected rows or quota exceeded).';
  }
  if (errorMsg) {
    var nextRow = (lastRow || startRow) + 1;
    return 'Stopped after sending ' + sentCount + ' emails, from row ' + firstRow + ' to row ' + lastRow + '. Error: ' + errorMsg + '. You can continue from row ' + nextRow + '.';
  }
  if (sentCount < (endRow - startRow + 1)) {
    var nextRow = (lastRow || startRow) + 1;
    return 'Quota exceeded! Sent ' + sentCount + ' emails, from row ' + firstRow + ' to row ' + lastRow + '. You can continue from row ' + nextRow + '.';
  }
  return 'Sending ' + sentCount + ' emails, from row ' + firstRow + ' to row ' + lastRow + '.';
}

/**
 * Trigger handler for scheduled mail merge.
 * Looks up config in PropertiesService and calls sendMailMerge.
 */
function sendMailMergeScheduled(e) {
  // Attempt to delete the trigger that called this function (cleanup)
  try {
    if (e && e.triggerUid) {
      var allTriggers = ScriptApp.getProjectTriggers();
      for (var i = 0; i < allTriggers.length; i++) {
        if (allTriggers[i].getUniqueId && allTriggers[i].getUniqueId() === e.triggerUid) {
          try { ScriptApp.deleteTrigger(allTriggers[i]); } catch (ex) { /* ignore deletion errors */ }
          break;
        }
      }
    }
  } catch (err) {
    // ignore errors when enumerating/deleting triggers
  }

  var props = PropertiesService.getScriptProperties();
  var foundKey = null;
  var foundConfig = null;

  // Preferred: lookup by trigger UID (reliable mapping)
  if (e && e.triggerUid) {
    var key = 'SCHEDULED_MAILMERGE_' + e.triggerUid;
    var raw = props.getProperty(key);
    if (raw) {
      try {
        var parsed = JSON.parse(raw);
        foundConfig = parsed.config || parsed; // support either {config:..., ...} or direct config
        foundKey = key;
      } catch (ex) {
        // parsing failed; continue to fallback
      }
    }
  }

  // Fallback: keep the old time-window search for legacy timestamp keys
  if (!foundConfig) {
    var now = Date.now();
    var all = props.getProperties();
    for (var k in all) {
      if (k.startsWith('SCHEDULED_MAILMERGE_')) {
        try {
          var entry = all[k];
          var parsed = JSON.parse(entry);
          var t = parsed && parsed.scheduledFor ? parseInt(parsed.scheduledFor, 10) : parseInt(k.replace('SCHEDULED_MAILMERGE_', ''), 10);
          if (!isNaN(t) && Math.abs(now - t) < 2 * 60 * 1000) { // within 2 minutes
            foundKey = k;
            foundConfig = parsed.config || parsed;
            break;
          }
        } catch (err) {
          // ignore parse errors and continue
        }
      }
    }
  }

  if (foundConfig) {
    try {
      sendMailMerge(foundConfig);
    } catch (err) {
      // If sendMailMerge throws, surface a log but don't rethrow to avoid stopping cleanup
      Logger.log('sendMailMergeScheduled: sendMailMerge failed: ' + (err && err.message ? err.message : err));
    }
    if (foundKey) {
      try { props.deleteProperty(foundKey); } catch (ex) { /* ignore delete failures */ }
    }
  }
}