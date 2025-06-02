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
  sheets.forEach(function(sheet) {
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
 * @param {Object} config
 * {
 *   sheetIdOrUrl: string,
 *   sheetName: string,
 *   rangeA1: string,
 *   emailCol: string,
 *   subject: string,
 *   body: string,
 *   mappings: [{placeholder: string, column: string, testValue: string}],
 *   fromName: string,
 *   fromEmail: string,
 *   test: boolean
 * }
 */
function sendMailMerge(config) {
  if (config.test) {
    // Send test email using test values
    var subject = config.subject;
    var body = config.body;
    var map = {};
    (config.mappings || []).forEach(function(m) {
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
  var range = sheet.getRange(config.rangeA1);
  var data = range.getValues();
  var headers = data[0];
  if (!headers) {
    throw new Error('No header row found in the specified range.');
  }
  var colMap = {};
  headers.forEach(function(h, idx) {
    colMap[h] = idx;
  });

  if (!(config.emailCol in colMap)) {
    throw new Error('Email column "' + config.emailCol + '" not found in the selected sheet.');
  }

  var map = {};
  config.mappings.forEach(function(m) {
    if (m.placeholder && m.column && (m.column in colMap)) {
      map[m.placeholder] = colMap[m.column];
    }
  });

  var sentCount = 0;
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var email = row[colMap[config.emailCol]];
    if (!email) continue;

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
    MailApp.sendEmail(mailOptions);
    sentCount++;
  }
  return 'Sent ' + sentCount + ' emails!';
}