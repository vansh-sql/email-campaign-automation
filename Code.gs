/**
 * Vns.WorkSpace — Email Bulk Tool
 * --------------------------------
 * BACKEND / SERVICE LAYER
 *
 * Architecture:
 *   Web App (Index.html) -> Apps Script service functions -> Google Sheet -> Gmail
 *
 * The existing Tool sheet remains the data source so the current mail workflow
 * stays compatible. UI logic is intentionally kept out of this file.
 */

var APP_NAME = 'Vns.WorkSpace';
var SHEET_NAME = 'Tool';
var START_ROW = 2;

// Phase 4 stores campaign metadata separately so the existing Tool sheet contract remains unchanged.
var CAMPAIGN_SHEET_NAME = 'Email_Campaigns';
var CAMPAIGN_LOG_SHEET_NAME = 'Email_Campaign_Logs';

// Tool sheet contract: keep these columns stable for backward compatibility.
var COL_NAME = 1;        // A: Name
var COL_EMAIL = 2;       // B: Email Id
var COL_TEMPLATE = 3;    // C: Google Doc template link
var COL_ATTACHMENT = 4;  // D: Drive attachment link
var COL_SUBJECT = 5;     // E: Subject
var COL_BODY = 6;        // F: Plain-text fallback body
var COL_STATUS = 7;      // G: Sending status

// Sender display-name mapping used by the existing mail engine.
var SENDER_NAME_MAP = {
  'singhv@gmail.com': 'Vansh Pratap Singh',
  'rajputboyyvansh.com': 'kunal Singh',
  'vanshpiparpur@gmail.com': 'Vinod'
};

/**
 * Web App router. Authentication is handled before the application shell is served.
 * Login.html is public; Index.html is the employee workspace; Admin.html is the
 * isolated Super Admin control tower. The existing mail/data services remain intact.
 */
function doGet(e) {
  e = e || {parameter:{}};
  var params = e.parameter || {};
  if (params.logout === '1' && params.session) {
    logoutSession(params.session);
  }

  var token = String(params.session || '').trim();
  if (!token) {
    return HtmlService.createTemplateFromFile('Login')
      .evaluate()
      .setTitle(APP_NAME + ' — Secure Sign In');
  }

  var ctx = getWorkspaceAccessContext(token);
  if (!ctx.allowed) {
    return HtmlService.createTemplateFromFile('Login')
      .evaluate()
      .setTitle(APP_NAME + ' — Secure Sign In');
  }

  if (ctx.role === 'Admin') {
    var admin = HtmlService.createTemplateFromFile('Admin');
    admin.sessionToken = token;
    admin.userContext = JSON.stringify(ctx).replace(/<\//g, '<\\/');
    return admin.evaluate().setTitle(APP_NAME + ' — Admin Control Center');
  }

  var employee = HtmlService.createTemplateFromFile('Index');
  employee.sessionToken = token;
  employee.userContext = JSON.stringify(ctx).replace(/<\//g, '<\\/');
  return employee.evaluate().setTitle(APP_NAME + ' — Workspace');
}

/** Optional include helper for future modular UI expansion. */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}

/**
 * Spreadsheet menu for admins who still work directly in the Tool sheet.
 * The existing manual workflow is preserved.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📧 Template Mail')
    .addItem('Send Mail (Selected Row)', 'templateSendSelectedRow')
    .addItem('Send All Pending Mails', 'templateSendAllPending')
    .addToUi();
}

/** Server-side authorization guard for every web-app service. UI visibility is not security. */
function requirePermission_(permission, sessionToken) {
  var ctx = getWorkspaceAccessContext(sessionToken);
  if (!ctx.allowed) throw new Error('Access denied. Your Vns.WorkSpace session is invalid or inactive.');
  if (!ctx.permissions || ctx.permissions[permission] !== true) {
    throw new Error('Access denied for the ' + permission + ' module.');
  }
  return ctx;
}

/** Prevents accidental self-lockout of the last/current administrator. */
function getCurrentIdentity_() {
  var active = '';
  try { active = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase(); } catch (e) {}
  var effective = '';
  try { effective = String(Session.getEffectiveUser().getEmail() || '').trim().toLowerCase(); } catch (e) {}
  return active || effective || '';
}

/**
 * Dashboard API: returns only the data required by the web UI.
 * No sheet formulas or internal implementation details are exposed.
 */
function getDashboardData(sessionToken) {
  requirePermission_('dashboard', sessionToken);
  var sheet = getToolSheet_();
  var lastRow = sheet.getLastRow();
  var rows = [];

  if (lastRow >= START_ROW) {
    var values = sheet.getRange(START_ROW, 1, lastRow - START_ROW + 1, COL_STATUS).getValues();
    values.forEach(function (r, index) {
      var email = String(r[COL_EMAIL - 1] || '').trim();
      if (!email) return;

      rows.push({
        row: START_ROW + index,
        name: String(r[COL_NAME - 1] || ''),
        email: email,
        subject: String(r[COL_SUBJECT - 1] || ''),
        status: String(r[COL_STATUS - 1] || ''),
        body: String(r[COL_BODY - 1] || ''),
        hasTemplate: !!String(r[COL_TEMPLATE - 1] || '').trim(),
        hasAttachment: !!String(r[COL_ATTACHMENT - 1] || '').trim()
      });
    });
  }

  var sent = rows.filter(function (r) { return r.status.indexOf('Sent Successfully') !== -1; }).length;
  var failed = rows.filter(function (r) { return r.status.indexOf('Failed:') === 0; }).length;
  var pending = rows.length - sent - failed;
  var quota = 0;
  try { quota = MailApp.getRemainingDailyQuota(); } catch (e) { quota = null; }

  return {
    appName: APP_NAME,
    total: rows.length,
    sent: sent,
    failed: failed,
    pending: pending,
    quota: quota,
    updatedAt: formatNow_(),
    sender: getActiveSender_(),
    recent: rows.slice(Math.max(0, rows.length - 8)).reverse(),
    pendingRows: rows.filter(function (r) { return r.status.indexOf('Sent Successfully') === -1; }).slice(0, 100)
  };
}


/**
 * Recipient management service.
 * Phase 2 only appends new recipient rows; existing Tool rows are never edited.
 * Supported import fields: Name, Email, Template, Attachment, Subject, Body.
 */
function getRecipientsData(sessionToken) {
  requirePermission_('recipients', sessionToken);
  var sheet = getToolSheet_();
  var lastRow = sheet.getLastRow();
  var rows = [];

  if (lastRow >= START_ROW) {
    var values = sheet.getRange(START_ROW, 1, lastRow - START_ROW + 1, COL_STATUS).getValues();
    values.forEach(function (r, index) {
      var email = String(r[COL_EMAIL - 1] || '').trim();
      if (!email) return;
      var status = String(r[COL_STATUS - 1] || '').trim();
      rows.push({
        row: START_ROW + index,
        name: String(r[COL_NAME - 1] || ''),
        email: email,
        subject: String(r[COL_SUBJECT - 1] || ''),
        status: status,
        hasTemplate: !!String(r[COL_TEMPLATE - 1] || '').trim(),
        hasAttachment: !!String(r[COL_ATTACHMENT - 1] || '').trim()
      });
    });
  }

  return {
    total: rows.length,
    sent: rows.filter(function (r) { return r.status.indexOf('Sent Successfully') !== -1; }).length,
    pending: rows.filter(function (r) { return r.status.indexOf('Sent Successfully') === -1 && r.status.indexOf('Failed:') !== 0; }).length,
    failed: rows.filter(function (r) { return r.status.indexOf('Failed:') === 0; }).length,
    rows: rows
  };
}

/**
 * Safely appends imported recipients to the existing Tool sheet.
 * Existing rows are read-only in this operation; duplicate emails are skipped.
 */
function importRecipients(rows, sessionToken) {
  requirePermission_('recipients', sessionToken);
  if (!Array.isArray(rows)) throw new Error('Invalid recipient import data.');
  if (!rows.length) throw new Error('No recipients found in the import file.');
  if (rows.length > 1000) throw new Error('Maximum 1000 recipients per import is allowed.');

  var sheet = getToolSheet_();
  var lastRow = sheet.getLastRow();
  var existing = {};
  if (lastRow >= START_ROW) {
    var emails = sheet.getRange(START_ROW, COL_EMAIL, lastRow - START_ROW + 1, 1).getValues();
    emails.forEach(function (r) {
      var email = String(r[0] || '').trim().toLowerCase();
      if (email) existing[email] = true;
    });
  }

  var appendRows = [];
  var invalid = 0;
  var duplicate = 0;

  rows.forEach(function (item) {
    var name = String(item.name || '').trim();
    var email = String(item.email || '').trim().toLowerCase();
    var template = String(item.template || '').trim();
    var attachment = String(item.attachment || '').trim();
    var subject = String(item.subject || '').trim();
    var body = String(item.body || '');

    if (!isValidEmail_(email)) {
      invalid++;
      return;
    }
    if (existing[email]) {
      duplicate++;
      return;
    }

    existing[email] = true;
    appendRows.push([name, email, template, attachment, subject, body, '']);
  });

  if (appendRows.length) {
    var start = Math.max(sheet.getLastRow() + 1, START_ROW);
    sheet.getRange(start, 1, appendRows.length, COL_STATUS).setValues(appendRows);
  }

  return {
    success: true,
    imported: appendRows.length,
    duplicate: duplicate,
    invalid: invalid,
    message: appendRows.length + ' recipient(s) imported. Existing rows were not changed.'
  };
}

/** Sends one row from the web UI. */
function sendWebRow(row, campaignName, senderDisplayName, sessionToken, selectedTemplateRow) {
  var accessCtx = requirePermission_('send', sessionToken);
  var requestedSenderName = String(senderDisplayName || '').trim();
  var allowedSenderNames = Object.keys(SENDER_NAME_MAP).map(function(email) {
    return SENDER_NAME_MAP[email];
  });
  allowedSenderNames.push(APP_NAME);
  allowedSenderNames = allowedSenderNames.filter(function(value, index, all) {
    return value && all.indexOf(value) === index;
  });
  if (!requestedSenderName || allowedSenderNames.indexOf(requestedSenderName) === -1) {
    throw new Error('Please choose a valid sender display name.');
  }
  var lock = LockService.getScriptLock();
  lock.tryLock(5000);
  try {
  var sheet = getToolSheet_();
  var rowNumber = Number(row);
  if (!rowNumber || rowNumber < START_ROW || rowNumber > sheet.getLastRow()) {
    throw new Error('Invalid recipient row.');
  }

  var quota = MailApp.getRemainingDailyQuota();
  if (quota <= 0) {
    throw new Error('Daily email quota is currently exhausted.');
  }

  var name = String(campaignName || '').trim() || 'Web Campaign';
  var campaign = createCampaign_(name, 'Single');
  var templateOverride = null;
  var templateRowNumber = Number(selectedTemplateRow || 0);
  if (templateRowNumber) {
    var templateSheet = getTemplateSheet_();
    if (templateRowNumber < 2 || templateRowNumber > templateSheet.getLastRow()) {
      throw new Error('Please choose a valid saved template.');
    }
    var templateValues = templateSheet.getRange(templateRowNumber, 1, 1, TEMPLATE_HEADERS.length).getValues()[0];
    templateOverride = {
      googleDocUrl: String(templateValues[4] || '').trim(),
      attachmentUrl: String(templateValues[5] || '').trim(),
      subject: String(templateValues[2] || '').trim(),
      body: String(templateValues[3] || '')
    };
    if (!templateOverride.subject) throw new Error('The selected template has no subject.');
    if (!templateOverride.googleDocUrl && !templateOverride.body.trim()) throw new Error('The selected template has no body or Google Doc.');
  }
  var result = sendTemplateMailForRow(sheet, rowNumber, requestedSenderName, templateOverride);
  appendCampaignLog_(campaign.id, rowNumber, result.success ? 'Sent' : 'Failed', result.message);
  recordEmailActivity_(campaign.id, campaign.name, rowNumber, result.success ? 'Sent' : 'Failed', result.message, 'Web Campaign', accessCtx);
  finalizeCampaign_(campaign.id, result.success ? 1 : 0, result.success ? 0 : 1);
  return result;
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/** Sends all pending rows using the existing quota-safe engine. */
function sendAllPendingWeb(campaignName, trackHistory, sessionToken) {
  var accessCtx = requirePermission_('send', sessionToken);
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) throw new Error('Another sending operation is already running. Please wait and try again.');
  try {
  var sheet = getToolSheet_();
  var lastRow = sheet.getLastRow();
  var sentCount = 0;
  var skippedCount = 0;
  var failedCount = 0;
  var quotaStoppedAt = null;
  var campaign = trackHistory ? createCampaign_(String(campaignName || '').trim() || 'Pending Queue', 'Bulk') : null;

  for (var row = START_ROW; row <= lastRow; row++) {
    var email = String(sheet.getRange(row, COL_EMAIL).getValue() || '').trim();
    if (!email) continue;

    var statusVal = String(sheet.getRange(row, COL_STATUS).getValue() || '');
    if (statusVal.indexOf('Sent Successfully') !== -1) {
      skippedCount++;
      continue;
    }

    if (MailApp.getRemainingDailyQuota() <= 0) {
      quotaStoppedAt = row;
      break;
    }

    var result = sendTemplateMailForRow(sheet, row);
    if (result.success) {
      sentCount++;
      if (campaign) appendCampaignLog_(campaign.id, row, 'Sent', result.message);
      recordEmailActivity_(campaign ? campaign.id : '', campaign ? campaign.name : 'Pending Queue', row, 'Sent', result.message, 'Bulk Send', accessCtx);
    } else {
      failedCount++;
      if (campaign) appendCampaignLog_(campaign.id, row, 'Failed', result.message);
      recordEmailActivity_(campaign ? campaign.id : '', campaign ? campaign.name : 'Pending Queue', row, 'Failed', result.message, 'Bulk Send', accessCtx);
    }
  }

  if (campaign) finalizeCampaign_(campaign.id, sentCount, failedCount);

  return {
    success: true,
    campaignId: campaign ? campaign.id : null,
    sent: sentCount,
    failed: failedCount,
    skipped: skippedCount,
    quotaStoppedAt: quotaStoppedAt,
    message: quotaStoppedAt
      ? 'Daily quota reached. Remaining rows are still pending.'
      : 'Pending mail processing completed.'
  };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}


/** Existing selected-row spreadsheet workflow. */
function templateSendSelectedRow() {
  var sheet = getToolSheet_();
  var ui = SpreadsheetApp.getUi();
  if (MailApp.getRemainingDailyQuota() <= 0) {
    ui.alert('Aaj ki daily mail quota poori ho chuki hai.');
    return;
  }

  var row = sheet.getActiveCell().getRow();
  if (row < START_ROW) {
    ui.alert('Valid data row select karein.');
    return;
  }

  var result = sendTemplateMailForRow(sheet, row);
  recordEmailActivity_('', 'Manual Selected Row', row, result.success ? 'Sent' : 'Failed', result.message, 'Spreadsheet Menu');
  ui.alert(result.message);
}

/** Existing bulk spreadsheet workflow, retained for backward compatibility. */
function templateSendAllPending() {
  var result = sendAllPendingWeb();
  var msg = 'Total ' + result.sent + ' mail bheji gayi. ' + result.skipped + ' already sent rows skip hui.';
  if (result.failed) msg += '\nFailed: ' + result.failed;
  if (result.quotaStoppedAt) msg += '\n\nDaily quota complete. Remaining rows pending hain.';
  SpreadsheetApp.getUi().alert(msg);
}

/**
 * Core mail engine.
 * This is the business/service layer used by both the sheet menu and web app.
 */
function sendTemplateMailForRow(sheet, row, senderDisplayName, templateOverride) {
  var name = sheet.getRange(row, COL_NAME).getValue();
  var email = String(sheet.getRange(row, COL_EMAIL).getValue() || '').trim();
  var templateLink = templateOverride ? String(templateOverride.googleDocUrl || '').trim() : String(sheet.getRange(row, COL_TEMPLATE).getValue() || '').trim();
  var attachLink = templateOverride ? String(templateOverride.attachmentUrl || '').trim() : String(sheet.getRange(row, COL_ATTACHMENT).getValue() || '').trim();
  var subject = templateOverride ? String(templateOverride.subject || '').trim() : String(sheet.getRange(row, COL_SUBJECT).getValue() || '').trim();
  var bodyText = templateOverride ? String(templateOverride.body || '') : String(sheet.getRange(row, COL_BODY).getValue() || '');
  var statusCell = sheet.getRange(row, COL_STATUS);

  if (!email) {
    statusCell.setValue('Email ID Missing');
    return { success: false, message: 'Row ' + row + ': Email ID missing hai.' };
  }
  if (!isValidEmail_(email)) {
    statusCell.setValue('Invalid Email');
    return { success: false, message: 'Row ' + row + ': Email format invalid hai.' };
  }
  if (!subject) {
    statusCell.setValue('Subject Missing');
    return { success: false, message: 'Row ' + row + ': Subject missing hai.' };
  }

  var htmlBody = null;
  var inlineImages = {};

  // Prefer the formatted Google Doc template; fall back to column F body.
  if (templateLink) {
    try {
      var templateFileId = extractDriveFileId(templateLink);
      if (templateFileId) {
        var converted = convertDocToHtmlWithImages(templateFileId);
        htmlBody = converted.html;
        inlineImages = converted.inlineImages;
      }
    } catch (templateErr) {
      Logger.log('Template fetch error row ' + row + ': ' + templateErr.message);
    }
  }

  if (!htmlBody) {
    if (!bodyText.trim()) {
      statusCell.setValue('Body Missing');
      return { success: false, message: 'Row ' + row + ': Template aur Body dono missing hain.' };
    }
    htmlBody = '<div style="font-family:Arial,sans-serif;font-size:14px;color:#222;line-height:1.6;">' + escapeHtml(bodyText).replace(/\n/g, '<br>') + '</div>';
  }

  var mailOptions = { htmlBody: htmlBody };
  var currentUserEmail = getActiveSender_().toLowerCase();
  // Web campaigns may choose a display name; sheet-menu sends retain the existing mapping.
  if (senderDisplayName) {
    mailOptions.name = String(senderDisplayName).trim();
  } else if (SENDER_NAME_MAP[currentUserEmail]) {
    mailOptions.name = SENDER_NAME_MAP[currentUserEmail];
  }

  // Attach the selected Drive file. If it is an image, also embed it inline
  // so the recipient can see it in the message body as well as open the attachment.
  if (attachLink) {
    var fileId = extractDriveFileId(attachLink);
    if (!fileId) {
      statusCell.setValue('Failed: Invalid attachment link');
      return { success: false, message: 'Row ' + row + ': Attachment link invalid hai. Drive file ka valid share link select karein.' };
    }
    try {
      var driveFile = DriveApp.getFileById(fileId);
      var fileBlob = driveFile.getBlob().setName(driveFile.getName());
      var mimeType = String(fileBlob.getContentType() || '').toLowerCase();
      mailOptions.attachments = [fileBlob];
      if (mimeType.indexOf('image/') === 0) {
        var attachCid = 'campaignImage';
        inlineImages[attachCid] = fileBlob;
        mailOptions.inlineImages = inlineImages;
        var imageHtml = '<div style="margin:18px 0;text-align:center;"><img src="cid:' + attachCid + '" alt="Email image" style="display:block;max-width:100%;height:auto;margin:0 auto;"></div>';
        if (/<\/body\s*>/i.test(mailOptions.htmlBody)) {
          mailOptions.htmlBody = mailOptions.htmlBody.replace(/<\/body\s*>/i, imageHtml + '</body>');
        } else {
          mailOptions.htmlBody = mailOptions.htmlBody + imageHtml;
        }
      }
    } catch (attachErr) {
      Logger.log('Attachment error for row ' + row + ': ' + attachErr.message);
      statusCell.setValue('Failed: Attachment could not be loaded');
      return { success: false, message: 'Row ' + row + ': Attachment load nahi hui. Drive permission/link check karein. Details: ' + attachErr.message };
    }
  }

  if (Object.keys(inlineImages).length > 0) mailOptions.inlineImages = inlineImages;

  try {
    GmailApp.sendEmail(email, subject, bodyText || 'Please view this email in an HTML-compatible mail client.', mailOptions);
    var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MMM-yyyy HH:mm');
    statusCell.setValue('Sent Successfully - ' + timestamp);
    return { success: true, message: 'Row ' + row + ' (' + name + '): Mail successfully bhej di gayi ' + email + ' par.' };
  } catch (err) {
    statusCell.setValue('Failed: ' + err.message);
    return { success: false, message: 'Row ' + row + ': Mail bhejne me error - ' + err.message };
  }
}


/** Returns a safe preview payload for the campaign review screen. */
function getWebRowPreview(row, sessionToken, selectedTemplateRow) {
  requirePermission_('send', sessionToken);
  var sheet = getToolSheet_();
  var rowNumber = Number(row);
  if (!rowNumber || rowNumber < START_ROW || rowNumber > sheet.getLastRow()) throw new Error('Invalid recipient row.');

  var templateLink = String(sheet.getRange(rowNumber, COL_TEMPLATE).getValue() || '').trim();
  var bodyText = String(sheet.getRange(rowNumber, COL_BODY).getValue() || '');
  var selectedRow = Number(selectedTemplateRow || 0);
  if (selectedRow) {
    var templateSheet = getTemplateSheet_();
    if (selectedRow < 2 || selectedRow > templateSheet.getLastRow()) throw new Error('Invalid saved template.');
    var selectedValues = templateSheet.getRange(selectedRow, 1, 1, TEMPLATE_HEADERS.length).getValues()[0];
    templateLink = String(selectedValues[4] || '').trim();
    bodyText = String(selectedValues[3] || '');
  }

  if (templateLink) {
    try {
      var templateFileId = extractDriveFileId(templateLink);
      if (templateFileId) {
        var converted = convertDocToHtmlWithImages(templateFileId);
        if (converted.html) return { html: converted.html, text: bodyText };
      }
    } catch (e) {
      Logger.log('Preview error row ' + rowNumber + ': ' + e.message);
    }
  }

  return { html: '', text: bodyText };
}

/** Sends the selected row content to a test address without changing the Tool row. */
function sendWebTest(row, testEmail, sessionToken) {
  var accessCtx = requirePermission_('send', sessionToken);
  var sheet = getToolSheet_();
  var rowNumber = Number(row);
  var email = String(testEmail || '').trim();
  if (!rowNumber || rowNumber < START_ROW || rowNumber > sheet.getLastRow()) throw new Error('Invalid recipient row.');
  if (!isValidEmail_(email)) throw new Error('Test email address is invalid.');
  if (MailApp.getRemainingDailyQuota() <= 0) throw new Error('Daily email quota is currently exhausted.');

  var subject = String(sheet.getRange(rowNumber, COL_SUBJECT).getValue() || '').trim();
  var templateLink = String(sheet.getRange(rowNumber, COL_TEMPLATE).getValue() || '').trim();
  var bodyText = String(sheet.getRange(rowNumber, COL_BODY).getValue() || '');
  if (!subject) throw new Error('Subject is missing in the selected row.');

  var htmlBody = null;
  var inlineImages = {};
  if (templateLink) {
    try {
      var templateFileId = extractDriveFileId(templateLink);
      if (templateFileId) {
        var converted = convertDocToHtmlWithImages(templateFileId);
        htmlBody = converted.html;
        inlineImages = converted.inlineImages;
      }
    } catch (e) {
      Logger.log('Test template error row ' + rowNumber + ': ' + e.message);
    }
  }
  if (!htmlBody) {
    if (!bodyText.trim()) throw new Error('Template and plain-text body are both missing.');
    htmlBody = '<div style="font-family:Arial,sans-serif;font-size:14px;color:#222;line-height:1.6;">' + escapeHtml(bodyText).replace(/\n/g, '<br>') + '</div>';
  }

  var options = { htmlBody: htmlBody };
  var sender = getActiveSender_().toLowerCase();
  if (SENDER_NAME_MAP[sender]) options.name = SENDER_NAME_MAP[sender];
  if (Object.keys(inlineImages).length) options.inlineImages = inlineImages;

  GmailApp.sendEmail(email, '[TEST] ' + subject, bodyText || 'Test email preview.', options);
  recordEmailActivity_('', 'Test Email', rowNumber, 'Sent', 'Test email sent to ' + email + '.', 'Test', accessCtx);
  return { success: true, message: 'Test email sent to ' + email + '.' };
}

/**
 * Phase 4 campaign history service.
 * Campaign data lives outside the Tool sheet so the existing mail data contract stays stable.
 */
function getCampaignsData(sessionToken) {
  requirePermission_('campaigns', sessionToken);
  var campaignSheet = getCampaignSheet_();
  var logSheet = getCampaignLogSheet_();
  var campaigns = [];
  var logs = [];

  if (campaignSheet.getLastRow() >= 2) {
    var values = campaignSheet.getRange(2, 1, campaignSheet.getLastRow() - 1, 9).getValues();
    values.forEach(function (r, index) {
      if (!r[0]) return;
      campaigns.push({
        row: index + 2,
        id: String(r[0]),
        name: String(r[1] || ''),
        mode: String(r[2] || ''),
        createdAt: String(r[3] || ''),
        status: String(r[4] || ''),
        total: Number(r[5] || 0),
        sent: Number(r[6] || 0),
        failed: Number(r[7] || 0),
        skipped: Number(r[8] || 0)
      });
    });
  }

  if (logSheet.getLastRow() >= 2) {
    var logValues = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 7).getValues();
    logValues.forEach(function (r, index) {
      if (!r[0]) return;
      logs.push({
        row: index + 2,
        campaignId: String(r[0]),
        recipientRow: Number(r[1] || 0),
        email: String(r[2] || ''),
        name: String(r[3] || ''),
        status: String(r[4] || ''),
        message: String(r[5] || ''),
        timestamp: String(r[6] || '')
      });
    });
  }

  return { campaigns: campaigns.reverse(), logs: logs.reverse() };
}

function createCampaign_(name, mode) {
  var sheet = getCampaignSheet_();
  var id = Utilities.getUuid();
  var now = formatNow_();
  sheet.appendRow([id, name, mode, now, 'Processing', 0, 0, 0, 0]);
  return { id: id, name: name, mode: mode };
}

function appendCampaignLog_(campaignId, recipientRow, status, message) {
  var toolSheet = getToolSheet_();
  var row = Number(recipientRow);
  var name = row >= START_ROW ? String(toolSheet.getRange(row, COL_NAME).getValue() || '') : '';
  var email = row >= START_ROW ? String(toolSheet.getRange(row, COL_EMAIL).getValue() || '') : '';
  getCampaignLogSheet_().appendRow([campaignId, row, email, name, status, message, formatNow_()]);
}

function finalizeCampaign_(campaignId, sent, failed) {
  var sheet = getCampaignSheet_();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(campaignId)) {
      var row = i + 1;
      var total = Number(sent || 0) + Number(failed || 0);
      var status = failed > 0 ? (sent > 0 ? 'Completed with Errors' : 'Failed') : 'Completed';
      sheet.getRange(row, 5, 1, 5).setValues([[status, total, Number(sent || 0), Number(failed || 0), 0]]);
      return;
    }
  }
}

function getCampaignSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CAMPAIGN_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CAMPAIGN_SHEET_NAME);
    sheet.getRange(1, 1, 1, 9).setValues([['Campaign ID', 'Campaign Name', 'Mode', 'Created At', 'Status', 'Total', 'Sent', 'Failed', 'Skipped']]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getCampaignLogSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CAMPAIGN_LOG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CAMPAIGN_LOG_SHEET_NAME);
    sheet.getRange(1, 1, 1, 7).setValues([['Campaign ID', 'Recipient Row', 'Email', 'Name', 'Status', 'Message', 'Timestamp']]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getToolSheet_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Sheet "' + SHEET_NAME + '" nahi mili.');
  return sheet;
}

function getActiveSender_() {
  try { return Session.getActiveUser().getEmail() || ''; } catch (e) { return ''; }
}

function formatNow_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy, hh:mm a');
}

function isValidEmail_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function extractDriveFileId(url) {
  var match = String(url || '').match(/[-\w]{25,}/);
  return match ? match[0] : null;
}

/** Converts supported Google Doc paragraph/list/table content to email HTML. */
function convertDocToHtmlWithImages(fileId) {
  var doc = DocumentApp.openById(fileId);
  var body = doc.getBody();
  var htmlParts = [];
  var inlineImages = {};
  var imageCounter = 0;

  for (var i = 0; i < body.getNumChildren(); i++) {
    var element = body.getChild(i);
    var type = element.getType();

    if (type === DocumentApp.ElementType.PARAGRAPH || type === DocumentApp.ElementType.LIST_ITEM) {
      var result = elementToHtml(element, inlineImages, imageCounter);
      imageCounter = result.counter;
      if (result.html.trim() || result.hasImageOnly) htmlParts.push('<p style="margin:0 0 12px 0;">' + result.html + '</p>');
    } else if (type === DocumentApp.ElementType.TABLE) {
      htmlParts.push('<div>' + escapeHtml(element.asTable().getText()) + '</div>');
    }
  }

  return {
    html: '<div style="font-family:Arial,sans-serif;font-size:14px;color:#222;line-height:1.6;">' + htmlParts.join('') + '</div>',
    inlineImages: inlineImages
  };
}

function elementToHtml(element, inlineImages, imageCounter) {
  var html = '';
  var hasImage = false;

  for (var j = 0; j < element.getNumChildren(); j++) {
    var child = element.getChild(j);
    var childType = child.getType();

    if (childType === DocumentApp.ElementType.TEXT) {
      var text = child.asText();
      var content = text.getText();
      if (!content) continue;

      var escaped = escapeHtml(content);
      var linkUrl = null;
      var isBold = false;
      try { linkUrl = text.getLinkUrl(); } catch (e1) {}
      try { isBold = text.isBold(); } catch (e2) {}

      if (isBold) escaped = '<b>' + escaped + '</b>';
      if (linkUrl) escaped = '<a href="' + escapeHtml(linkUrl) + '" style="color:#1a73e8;">' + escaped + '</a>';
      html += escaped;
    } else if (childType === DocumentApp.ElementType.INLINE_IMAGE) {
      var img = child.asInlineImage();
      var blob = img.getBlob();
      var cid = 'templateImage' + imageCounter++;
      inlineImages[cid] = blob;
      html += '<br><img src="cid:' + cid + '" style="max-width:100%;height:auto;"><br>';
      hasImage = true;
    }
  }

  return { html: html, counter: imageCounter, hasImageOnly: hasImage };
}

function escapeHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

/* --------------------------------------------------------------------------
 * PHASE 3 — TEMPLATE LIBRARY
 *
 * This module is intentionally isolated from the existing Tool-sheet contract.
 * It creates/uses a separate Email_Templates sheet and does not rewrite any
 * existing Tool rows. The existing mail engine continues to read columns A:G.
 * -------------------------------------------------------------------------- */

var TEMPLATE_SHEET_NAME = 'Email_Templates';
var TEMPLATE_HEADERS = ['Template ID', 'Template Name', 'Subject', 'Body', 'Google Doc URL', 'Attachment URL', 'Created At', 'Updated At'];

function getTemplatesData(sessionToken) {
  requirePermission_('templates', sessionToken);
  var sheet = getTemplateSheet_();
  var lastRow = sheet.getLastRow();
  var rows = [];

  if (lastRow >= 2) {
    var values = sheet.getRange(2, 1, lastRow - 1, TEMPLATE_HEADERS.length).getValues();
    values.forEach(function (r, index) {
      var name = String(r[1] || '').trim();
      if (!name) return;
      rows.push({
        row: index + 2,
        id: String(r[0] || ''),
        name: name,
        subject: String(r[2] || ''),
        body: String(r[3] || ''),
        googleDocUrl: String(r[4] || ''),
        attachmentUrl: String(r[5] || ''),
        createdAt: String(r[6] || ''),
        updatedAt: String(r[7] || '')
      });
    });
  }

  return {
    total: rows.length,
    rows: rows.reverse()
  };
}

function saveEmailTemplate(template, sessionToken) {
  requirePermission_('templates', sessionToken);
  if (!template || typeof template !== 'object') throw new Error('Invalid template data.');

  var name = String(template.name || '').trim();
  var subject = String(template.subject || '').trim();
  var body = String(template.body || '');
  var googleDocUrl = String(template.googleDocUrl || '').trim();
  var attachmentUrl = String(template.attachmentUrl || '').trim();

  if (!name) throw new Error('Template name is required.');
  if (name.length > 120) throw new Error('Template name is too long.');
  if (!subject) throw new Error('Subject is required.');
  if (!body && !googleDocUrl) throw new Error('Add a body or Google Doc template URL.');
  if (googleDocUrl && !extractDriveFileId(googleDocUrl)) throw new Error('Google Doc URL is not valid.');
  if (attachmentUrl && !extractDriveFileId(attachmentUrl)) throw new Error('Attachment URL is not valid.');

  var sheet = getTemplateSheet_();
  var now = formatNow_();
  var rowNumber = Number(template.row || 0);

  if (rowNumber >= 2 && rowNumber <= sheet.getLastRow()) {
    var existingName = String(sheet.getRange(rowNumber, 2).getValue() || '').trim();
    if (!existingName) throw new Error('Template record was not found. Refresh and try again.');

    sheet.getRange(rowNumber, 2, 1, TEMPLATE_HEADERS.length - 1).setValues([[
      name, subject, body, googleDocUrl, attachmentUrl,
      sheet.getRange(rowNumber, 7).getValue() || now, now
    ]]);

    return { success: true, mode: 'updated', row: rowNumber, message: 'Template updated successfully.' };
  }

  var id = 'TPL-' + Utilities.getUuid().slice(0, 8).toUpperCase();
  sheet.appendRow([id, name, subject, body, googleDocUrl, attachmentUrl, now, now]);

  return { success: true, mode: 'created', row: sheet.getLastRow(), id: id, message: 'Template created successfully.' };
}

function deleteEmailTemplate(row, sessionToken) {
  requirePermission_('templates', sessionToken);
  var sheet = getTemplateSheet_();
  var rowNumber = Number(row);
  if (!rowNumber || rowNumber < 2 || rowNumber > sheet.getLastRow()) throw new Error('Invalid template row.');
  var name = String(sheet.getRange(rowNumber, 2).getValue() || '').trim();
  if (!name) throw new Error('Template record was not found.');
  sheet.deleteRow(rowNumber);
  return { success: true, message: 'Template deleted: ' + name };
}

function getTemplatePreview(row, sessionToken) {
  requirePermission_('templates', sessionToken);
  var sheet = getTemplateSheet_();
  var rowNumber = Number(row);
  if (!rowNumber || rowNumber < 2 || rowNumber > sheet.getLastRow()) throw new Error('Invalid template row.');

  var values = sheet.getRange(rowNumber, 1, 1, TEMPLATE_HEADERS.length).getValues()[0];
  var body = String(values[3] || '');
  var docUrl = String(values[4] || '').trim();

  if (docUrl) {
    try {
      var fileId = extractDriveFileId(docUrl);
      if (fileId) {
        var converted = convertDocToHtmlWithImages(fileId);
        if (converted.html) return { html: converted.html, text: body };
      }
    } catch (e) {
      Logger.log('Template preview error row ' + rowNumber + ': ' + e.message);
    }
  }

  return {
    html: body ? '<div style="font-family:Arial,sans-serif;font-size:14px;color:#222;line-height:1.6;white-space:pre-wrap;">' + escapeHtml(body) + '</div>' : '',
    text: body
  };
}

function getTemplateSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(TEMPLATE_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(TEMPLATE_SHEET_NAME);
    sheet.getRange(1, 1, 1, TEMPLATE_HEADERS.length).setValues([TEMPLATE_HEADERS]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, TEMPLATE_HEADERS.length).setFontWeight('bold');
  } else if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, TEMPLATE_HEADERS.length).setValues([TEMPLATE_HEADERS]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}


/* --------------------------------------------------------------------------
 * PHASE 6 — REPORTS & ANALYTICS
 *
 * Reporting is additive. A separate activity sheet captures sender/employee,
 * campaign, recipient and result details without changing Tool A:G or the
 * existing campaign/template/settings sheet contracts.
 * -------------------------------------------------------------------------- */

var ACTIVITY_SHEET_NAME = 'Email_Activity_Log';
var ACTIVITY_HEADERS = ['Activity ID', 'Timestamp', 'Employee Email', 'Employee Name', 'Campaign ID', 'Campaign Name', 'Recipient Row', 'Recipient Name', 'Recipient Email', 'Status', 'Message', 'Source', 'Is Test'];

function recordEmailActivity_(campaignId, campaignName, recipientRow, status, message, source, actor) {
  try {
    var sheet = getActivitySheet_();
    var toolSheet = getToolSheet_();
    var row = Number(recipientRow || 0);
    var employeeEmail = actor && actor.email ? String(actor.email).trim().toLowerCase() : getActiveSender_();
    var employeeName = actor && actor.name ? String(actor.name).trim() : (SENDER_NAME_MAP[String(employeeEmail || '').toLowerCase()] || employeeEmail || 'Unknown Employee');
    var recipientName = row >= START_ROW ? String(toolSheet.getRange(row, COL_NAME).getValue() || '') : '';
    var recipientEmail = row >= START_ROW ? String(toolSheet.getRange(row, COL_EMAIL).getValue() || '') : '';
    var isTest = String(source || '').toLowerCase() === 'test';
    sheet.appendRow([
      Utilities.getUuid(), new Date(), employeeEmail, employeeName,
      String(campaignId || ''), String(campaignName || ''), row,
      recipientName, recipientEmail, String(status || ''), String(message || ''),
      String(source || ''), isTest ? 'Yes' : 'No'
    ]);
  } catch (e) {
    // Reporting must never block the existing mail-sending workflow.
    Logger.log('Activity log error: ' + e.message);
  }
}

function getReportsData(filters, sessionToken) {
  requirePermission_('reports', sessionToken);
  filters = filters || {};
  var start = parseReportDate_(filters.startDate, false);
  var end = parseReportDate_(filters.endDate, true);
  var employeeFilter = String(filters.employee || '').trim().toLowerCase();
  var statusFilter = String(filters.status || 'all').trim().toLowerCase();
  var includeTests = String(filters.includeTests || 'false') === 'true';
  var rows = [];
  var sheet = getActivitySheet_();

  if (sheet.getLastRow() >= 2) {
    var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, ACTIVITY_HEADERS.length).getValues();
    values.forEach(function (r, index) {
      if (!r[0]) return;
      var timestamp = r[1] instanceof Date ? r[1] : new Date(r[1]);
      if (isNaN(timestamp.getTime())) return;
      var employeeEmail = String(r[2] || '').trim();
      var employeeName = String(r[3] || '').trim() || employeeEmail || 'Unknown Employee';
      var status = String(r[9] || '').trim();
      var isTest = String(r[12] || '').toLowerCase() === 'yes';
      if (!includeTests && isTest) return;
      if (start && timestamp < start) return;
      if (end && timestamp > end) return;
      if (employeeFilter && employeeFilter !== '__legacy__' && (employeeEmail + ' ' + employeeName).toLowerCase().indexOf(employeeFilter) === -1) return;
      if (employeeFilter === '__legacy__') return;
      if (statusFilter !== 'all' && status.toLowerCase() !== statusFilter) return;
      rows.push({
        activityId: String(r[0]), timestamp: timestamp.getTime(),
        timestampLabel: Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'dd MMM yyyy, hh:mm a'),
        employeeEmail: employeeEmail, employeeName: employeeName,
        campaignId: String(r[4] || ''), campaignName: String(r[5] || '') || 'Unassigned / Manual',
        recipientRow: Number(r[6] || 0), recipientName: String(r[7] || ''),
        recipientEmail: String(r[8] || ''), status: status,
        message: String(r[10] || ''), source: String(r[11] || ''), isTest: isTest
      });
    });
  }

  /* Phase 6 compatibility: include historical Tool-sheet outcomes that were created
   * before Email_Activity_Log existed. They are intentionally marked unattributed
   * instead of assigning them to an employee without evidence. */
  var loggedKeys = {};
  rows.forEach(function(r){
    if (r.recipientRow && r.status) loggedKeys[String(r.recipientRow) + '|' + r.status.toLowerCase()] = true;
  });
  var toolSheetForReport = getToolSheet_();
  if (toolSheetForReport.getLastRow() >= START_ROW) {
    var toolValues = toolSheetForReport.getRange(START_ROW, 1, toolSheetForReport.getLastRow() - START_ROW + 1, COL_STATUS).getValues();
    toolValues.forEach(function(r, index){
      var email = String(r[COL_EMAIL - 1] || '').trim();
      var statusRaw = String(r[COL_STATUS - 1] || '').trim();
      var statusLower = statusRaw.toLowerCase();
      if (!email || (statusLower.indexOf('sent successfully') !== 0 && statusLower.indexOf('failed:') !== 0)) return;
      var status = statusLower.indexOf('sent successfully') === 0 ? 'Sent' : 'Failed';
      var recipientRow = START_ROW + index;
      if (loggedKeys[String(recipientRow) + '|' + status.toLowerCase()]) return;
      var parsed = parseStatusTimestamp_(statusRaw);
      var timestamp = parsed || new Date(0);
      if (start && timestamp.getTime() && timestamp < start) return;
      if (end && timestamp.getTime() && timestamp > end) return;
      var legacyEmployee = 'Unattributed / Existing Tool Data';
      if (employeeFilter && employeeFilter !== '__legacy__') return;
      if (employeeFilter === '__legacy__' || !employeeFilter) {
        rows.push({
          activityId: 'legacy-' + recipientRow,
          timestamp: timestamp.getTime(),
          timestampLabel: parsed ? Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'dd MMM yyyy, hh:mm a') : 'Existing record',
          employeeEmail: '', employeeName: legacyEmployee,
          campaignId: '', campaignName: 'Existing Tool Data',
          recipientRow: recipientRow, recipientName: String(r[COL_NAME - 1] || ''),
          recipientEmail: email, status: status,
          message: statusRaw, source: 'Legacy Tool Sheet', isTest: false, legacy: true
        });
      }
    });
  }

  rows.sort(function(a,b){ return b.timestamp - a.timestamp; });

  var sent = rows.filter(function(r){return r.status.toLowerCase() === 'sent';}).length;
  var failed = rows.filter(function(r){return r.status.toLowerCase() === 'failed';}).length;
  var total = rows.length;
  var successRate = total ? Math.round((sent / total) * 1000) / 10 : 0;
  var employees = {};
  var campaigns = {};
  var days = {};
  rows.forEach(function(r){
    var ek = r.employeeEmail || r.employeeName || 'unknown';
    if(!employees[ek]) employees[ek] = {employeeEmail:r.employeeEmail, employeeName:r.employeeName, total:0, sent:0, failed:0};
    employees[ek].total++;
    if(r.status.toLowerCase()==='sent') employees[ek].sent++; else if(r.status.toLowerCase()==='failed') employees[ek].failed++;
    var ck = r.campaignId || ('manual:' + r.campaignName);
    if(!campaigns[ck]) campaigns[ck] = {campaignId:r.campaignId, campaignName:r.campaignName, total:0, sent:0, failed:0};
    campaigns[ck].total++;
    if(r.status.toLowerCase()==='sent') campaigns[ck].sent++; else if(r.status.toLowerCase()==='failed') campaigns[ck].failed++;
    var day = Utilities.formatDate(new Date(r.timestamp), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if(!days[day]) days[day] = {date:day,total:0,sent:0,failed:0};
    days[day].total++;
    if(r.status.toLowerCase()==='sent') days[day].sent++; else if(r.status.toLowerCase()==='failed') days[day].failed++;
  });

  var employeeRows = Object.keys(employees).map(function(k){
    var r=employees[k]; r.successRate=r.total?Math.round((r.sent/r.total)*1000)/10:0; return r;
  }).sort(function(a,b){return b.sent-a.sent || b.total-a.total;});
  var campaignRows = Object.keys(campaigns).map(function(k){
    var r=campaigns[k]; r.successRate=r.total?Math.round((r.sent/r.total)*1000)/10:0; return r;
  }).sort(function(a,b){return b.total-a.total;});
  var trendRows = Object.keys(days).map(function(k){return days[k];}).sort(function(a,b){return a.date.localeCompare(b.date);});

  return {
    generatedAt: formatNow_(), filters: {startDate: filters.startDate || '', endDate: filters.endDate || '', employee: filters.employee || '', status: statusFilter, includeTests: includeTests},
    summary: {total:total,sent:sent,failed:failed,successRate:successRate,employees:employeeRows.length,campaigns:campaignRows.length,uniqueRecipients:uniqueCount_(rows)},
    employees: employeeRows, campaigns: campaignRows, trend: trendRows,
    recent: rows.slice(0,100), hasActivityLog: sheet.getLastRow() >= 2
  };
}

function uniqueCount_(rows) {
  var set = {};
  rows.forEach(function(r){ if(r.recipientEmail) set[r.recipientEmail.toLowerCase()] = true; });
  return Object.keys(set).length;
}

function getReportEmployees(sessionToken) {
  var sheet = getActivitySheet_();
  var map = {};
  if(sheet.getLastRow() >= 2){
    sheet.getRange(2,3,sheet.getLastRow()-1,2).getValues().forEach(function(r){
      var email=String(r[0]||'').trim(), name=String(r[1]||'').trim();
      if(email || name) map[email || name] = {email:email,name:name||email};
    });
  }
  var toolSheet = getToolSheet_();
  var hasLegacy = false;
  if (toolSheet.getLastRow() >= START_ROW) {
    toolSheet.getRange(START_ROW, COL_EMAIL, toolSheet.getLastRow()-START_ROW+1, COL_STATUS-COL_EMAIL+1).getValues().forEach(function(r){
      var email=String(r[0]||'').trim(), status=String(r[COL_STATUS-COL_EMAIL]||'').toLowerCase();
      if(email && (status.indexOf('sent successfully')===0 || status.indexOf('failed:')===0)) hasLegacy=true;
    });
  }
  var result = Object.keys(map).map(function(k){return map[k];}).sort(function(a,b){return a.name.localeCompare(b.name);});
  if (hasLegacy) result.unshift({email:'__legacy__',name:'Unattributed / Existing Tool Data'});
  return result;
}

function exportReports(filters, format, sessionToken) {
  requirePermission_('reports', sessionToken);
  var data = getReportsData(filters || {}, sessionToken);
  var fmt = String(format || 'xlsx').toLowerCase();
  if (fmt !== 'xlsx' && fmt !== 'pdf') throw new Error('Unsupported report format.');

  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmm');
  var title = 'Vns.WorkSpace_Report_' + stamp;
  var temp = SpreadsheetApp.create(title + '_TEMP');
  try {
    var summary = temp.getSheets()[0];
    summary.setName('Executive Summary');
    var summaryRows = [
      ['Vns.WorkSpace — Email Analytics Report'],
      ['Generated', data.generatedAt],
      ['Date Range', (data.filters.startDate || 'All') + ' to ' + (data.filters.endDate || 'All')],
      ['', ''],
      ['Metric','Value'],
      ['Total Emails',data.summary.total],['Sent',data.summary.sent],['Failed',data.summary.failed],
      ['Success Rate',data.summary.successRate + '%'],['Unique Recipients',data.summary.uniqueRecipients],
      ['Employees',data.summary.employees],['Campaigns',data.summary.campaigns]
    ];
    summary.getRange(1,1,summaryRows.length,2).setValues(summaryRows);
    styleReportSheet_(summary, 2);

    var emp = temp.insertSheet('Employee Performance');
    emp.getRange(1,1,1,6).setValues([['Employee','Email','Total','Sent','Failed','Success Rate']]);
    if(data.employees.length) emp.getRange(2,1,data.employees.length,6).setValues(data.employees.map(function(r){return [r.employeeName,r.employeeEmail,r.total,r.sent,r.failed,r.successRate+'%'];}));
    styleReportSheet_(emp,6); emp.getRange(1,1,Math.max(1,data.employees.length+1),6).createFilter();

    var camp = temp.insertSheet('Campaign Performance');
    camp.getRange(1,1,1,6).setValues([['Campaign','Total','Sent','Failed','Success Rate','Campaign ID']]);
    if(data.campaigns.length) camp.getRange(2,1,data.campaigns.length,6).setValues(data.campaigns.map(function(r){return [r.campaignName,r.total,r.sent,r.failed,r.successRate+'%',r.campaignId];}));
    styleReportSheet_(camp,6); camp.getRange(1,1,Math.max(1,data.campaigns.length+1),6).createFilter();

    var trend = temp.insertSheet('Daily Trend');
    trend.getRange(1,1,1,4).setValues([['Date','Total','Sent','Failed']]);
    if(data.trend.length) trend.getRange(2,1,data.trend.length,4).setValues(data.trend.map(function(r){return [r.date,r.total,r.sent,r.failed];}));
    styleReportSheet_(trend,4); trend.getRange(1,1,Math.max(1,data.trend.length+1),4).createFilter();

    var recent = temp.insertSheet('Activity Log');
    recent.getRange(1,1,1,8).setValues([['Timestamp','Employee','Email','Campaign','Recipient','Recipient Email','Status','Source']]);
    if(data.recent.length) recent.getRange(2,1,data.recent.length,8).setValues(data.recent.map(function(r){return [new Date(r.timestamp),r.employeeName,r.employeeEmail,r.campaignName,r.recipientName,r.recipientEmail,r.status,r.source];}));
    styleReportSheet_(recent,8); recent.getRange(1,1,Math.max(1,data.recent.length+1),8).createFilter();

    SpreadsheetApp.flush();
    var file;
    if(fmt === 'xlsx') {
      var xurl = 'https://docs.google.com/spreadsheets/d/' + temp.getId() + '/export?format=xlsx';
      var response = UrlFetchApp.fetch(xurl, {headers:{Authorization:'Bearer ' + ScriptApp.getOAuthToken()}, muteHttpExceptions:true});
      if(response.getResponseCode() < 200 || response.getResponseCode() >= 300) throw new Error('Excel export failed (HTTP ' + response.getResponseCode() + ').');
      file = DriveApp.createFile(response.getBlob().setName(title + '.xlsx'));
    } else {
      var purl = 'https://docs.google.com/spreadsheets/d/' + temp.getId() + '/export?format=pdf&size=A4&portrait=false&fitw=true&sheetnames=true&printtitle=true&pagenumbers=true&gridlines=false&fzr=true';
      var presponse = UrlFetchApp.fetch(purl, {headers:{Authorization:'Bearer ' + ScriptApp.getOAuthToken()}, muteHttpExceptions:true});
      if(presponse.getResponseCode() < 200 || presponse.getResponseCode() >= 300) throw new Error('PDF export failed (HTTP ' + presponse.getResponseCode() + ').');
      file = DriveApp.createFile(presponse.getBlob().setName(title + '.pdf'));
    }
    return {success:true,name:file.getName(),url:file.getUrl(),id:file.getId(),format:fmt};
  } finally {
    try { DriveApp.getFileById(temp.getId()).setTrashed(true); } catch(e) { Logger.log('Temp report cleanup: ' + e.message); }
  }
}

function styleReportSheet_(sheet, columns) {
  sheet.setFrozenRows(1);
  sheet.getRange(1,1,1,columns).setFontWeight('bold').setBackground('#1e3a8a').setFontColor('#ffffff');
  sheet.autoResizeColumns(1, columns);
  if(sheet.getLastRow() > 1) sheet.getRange(1,1,sheet.getLastRow(),columns).setVerticalAlignment('middle');
}

function getActivitySheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(ACTIVITY_SHEET_NAME);
  if(!sheet){
    sheet = ss.insertSheet(ACTIVITY_SHEET_NAME);
    sheet.getRange(1,1,1,ACTIVITY_HEADERS.length).setValues([ACTIVITY_HEADERS]);
    sheet.setFrozenRows(1);
    sheet.getRange(1,1,1,ACTIVITY_HEADERS.length).setFontWeight('bold');
  } else if(sheet.getLastRow() === 0){
    sheet.getRange(1,1,1,ACTIVITY_HEADERS.length).setValues([ACTIVITY_HEADERS]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function parseStatusTimestamp_(status) {
  var s = String(status || '');
  var match = s.match(/-\s*(\d{1,2}[-\/]\d{1,2}[-\/]\d{2,4}(?:\s+\d{1,2}:\d{2}(?::\d{2})?\s*(?:AM|PM)?)?)/i);
  if (!match) return null;
  var raw = match[1].trim().replace(/\//g, '-');
  var d = new Date(raw);
  if (isNaN(d.getTime())) {
    var parts = raw.split(/[-\s:]+/);
    if (parts.length >= 3) {
      var day = Number(parts[0]), month = Number(parts[1]) - 1, year = Number(parts[2]);
      if (year < 100) year += 2000;
      d = new Date(year, month, day);
    }
  }
  return isNaN(d.getTime()) ? null : d;
}

function parseReportDate_(value, endOfDay) {
  var s = String(value || '').trim();
  if(!s) return null;
  var parts = s.split('-');
  if(parts.length !== 3) return null;
  var d = new Date(Number(parts[0]), Number(parts[1])-1, Number(parts[2]), endOfDay ? 23 : 0, endOfDay ? 59 : 0, endOfDay ? 59 : 0);
  return isNaN(d.getTime()) ? null : d;
}

/* --------------------------------------------------------------------------
 * PHASE 5 — WORKSPACE SETTINGS
 *
 * Settings are stored in a separate Email_Settings sheet. This keeps the
 * existing Tool sheet, campaign sheets, template sheet, and mail engine
 * contracts unchanged. Current sender identity remains read-only here.
 * -------------------------------------------------------------------------- */

var SETTINGS_SHEET_NAME = 'Email_Settings';
var SETTINGS_HEADERS = ['Setting Key', 'Setting Value', 'Updated At'];

function getWorkspaceSettings(sessionToken) {
  requirePermission_('settings', sessionToken);
  var sheet = getSettingsSheet_();
  var values = sheet.getLastRow() >= 2
    ? sheet.getRange(2, 1, sheet.getLastRow() - 1, SETTINGS_HEADERS.length).getValues()
    : [];
  var map = {};

  values.forEach(function (r) {
    var key = String(r[0] || '').trim();
    if (key) map[key] = String(r[1] || '');
  });

  var defaults = getDefaultWorkspaceSettings_();
  Object.keys(defaults).forEach(function (key) {
    if (map[key] === undefined) map[key] = defaults[key];
  });

  return {
    profile: {
      companyName: map.companyName,
      productName: map.productName,
      supportEmail: map.supportEmail,
      website: map.website
    },
    campaign: {
      defaultCampaignName: map.defaultCampaignName
    },
    system: {
      sender: getActiveSender_(),
      timezone: Session.getScriptTimeZone(),
      toolSheet: SHEET_NAME,
      templateSheet: TEMPLATE_SHEET_NAME,
      campaignSheet: CAMPAIGN_SHEET_NAME,
      campaignLogSheet: CAMPAIGN_LOG_SHEET_NAME,
      settingsSheet: SETTINGS_SHEET_NAME
    }
  };
}

function saveWorkspaceSettings(settings, sessionToken) {
  requirePermission_('settings', sessionToken);
  if (!settings || typeof settings !== 'object') {
    throw new Error('Invalid settings data.');
  }

  var profile = settings.profile || {};
  var campaign = settings.campaign || {};

  var companyName = String(profile.companyName || '').trim();
  var productName = String(profile.productName || '').trim();
  var supportEmail = String(profile.supportEmail || '').trim();
  var website = String(profile.website || '').trim();
  var defaultCampaignName = String(campaign.defaultCampaignName || '').trim();

  if (!companyName) throw new Error('Company name is required.');
  if (!productName) throw new Error('Product name is required.');
  if (supportEmail && !isValidEmail_(supportEmail)) {
    throw new Error('Support email is not valid.');
  }
  if (companyName.length > 120 || productName.length > 120) {
    throw new Error('Company or product name is too long.');
  }
  if (defaultCampaignName.length > 80) {
    throw new Error('Default campaign name is too long.');
  }

  var sheet = getSettingsSheet_();
  var updates = {
    companyName: companyName,
    productName: productName,
    supportEmail: supportEmail,
    website: website,
    defaultCampaignName: defaultCampaignName
  };

  var now = formatNow_();
  var existingRows = {};
  if (sheet.getLastRow() >= 2) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().forEach(function (r, i) {
      var key = String(r[0] || '').trim();
      if (key) existingRows[key] = i + 2;
    });
  }

  Object.keys(updates).forEach(function (key) {
    var value = updates[key];
    if (existingRows[key]) {
      sheet.getRange(existingRows[key], 2, 1, 2).setValues([[value, now]]);
    } else {
      sheet.appendRow([key, value, now]);
    }
  });

  return {
    success: true,
    message: 'Workspace settings saved successfully.',
    settings: getWorkspaceSettings(sessionToken)
  };
}

function resetWorkspaceSettings(sessionToken) {
  requirePermission_('settings', sessionToken);
  var sheet = getSettingsSheet_();
  var defaults = getDefaultWorkspaceSettings_();
  var now = formatNow_();

  var existingRows = {};
  if (sheet.getLastRow() >= 2) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().forEach(function (r, i) {
      var key = String(r[0] || '').trim();
      if (key) existingRows[key] = i + 2;
    });
  }

  Object.keys(defaults).forEach(function (key) {
    if (existingRows[key]) {
      sheet.getRange(existingRows[key], 2, 1, 2).setValues([[defaults[key], now]]);
    } else {
      sheet.appendRow([key, defaults[key], now]);
    }
  });

  return {
    success: true,
    message: 'Workspace settings restored to defaults.',
    settings: getWorkspaceSettings(sessionToken)
  };
}

function getDefaultWorkspaceSettings_() {
  return {
    companyName: APP_NAME,
    productName: 'Email Bulk Tool',
    supportEmail: '',
    website: '',
    defaultCampaignName: 'Vns.WorkSpace Campaign'
  };
}

function getSettingsSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SETTINGS_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SETTINGS_SHEET_NAME);
    sheet.getRange(1, 1, 1, SETTINGS_HEADERS.length).setValues([SETTINGS_HEADERS]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, SETTINGS_HEADERS.length).setFontWeight('bold');
    return sheet;
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, SETTINGS_HEADERS.length).setValues([SETTINGS_HEADERS]);
    sheet.setFrozenRows(1);
  }

  return sheet;
}



/* --------------------------------------------------------------------------
 * IDENTITY / ACCESS DIRECTORY
 *
 * Product model:
 *   - Super Admin uses one fixed first-time/trial credential.
 *   - All employees are stored in one central Employee_Master sheet.
 *   - Employee passwords are never generated automatically.
 *   - Employee ID can be generated automatically only when the Admin enables
 *     the auto-ID option; otherwise the Admin supplied ID is used.
 *   - Employee passwords are stored exactly as entered by the Super Admin so the
 *     assigned login password remains visible in Employee_Master.
 *   - Legacy identity tabs and old Employee_<ID> sheets are migrated once,
 *     then removed from the active product structure.
 * -------------------------------------------------------------------------- */

var SESSION_PREFIX = 'VNS_WS_SESSION_';
var SESSION_TTL_SECONDS = 21600;

// FIRST-TIME / TRIAL CREDENTIAL. Change this before client production use.
var SUPER_ADMIN_EMPLOYEE_ID = 'ADMIN001';
var SUPER_ADMIN_PASSWORD = 'Vns@Admin#7314';
var SUPER_ADMIN_USER_ID = 'SUPER_ADMIN';
var SUPER_ADMIN_NAME = 'Vns Super Admin';
var SUPER_ADMIN_EMAIL = 'admin@vnsworkspace.local';

var EMPLOYEE_MASTER_SHEET = 'Employee_Master';
var EMPLOYEE_HEADERS = [
  'Employee ID',
  'User ID',
  'Name',
  'Email',
  'Role',
  'Status',
  'Department',
  'Manager ID',
  'Password',
  'Failed Attempts',
  'Locked Until',
  'Last Login',
  'Created At',
  'Updated At'
];

var LEGACY_IDENTITY_SHEETS = [
  'Workspace_Auth',
  'Workspace_Users',
  'Workspace_User_Profile'
];

function authenticateUser(employeeId, password) {
  ensureIdentityDirectory_();

  var id = String(employeeId || '').trim().toUpperCase();
  var pass = String(password || '');
  if (!id || !pass) throw new Error('Employee ID and password are required.');
  if (id.length > 40 || pass.length > 128) throw new Error('Invalid login details.');

  var now = new Date();

  if (id === SUPER_ADMIN_EMPLOYEE_ID) {
    if (pass !== SUPER_ADMIN_PASSWORD) {
      throw new Error('Employee ID or password is incorrect.');
    }

    return {
      success: true,
      token: createSession_(
        SUPER_ADMIN_USER_ID,
        SUPER_ADMIN_EMPLOYEE_ID,
        SUPER_ADMIN_EMAIL,
        SUPER_ADMIN_NAME,
        'Admin'
      ),
      role: 'Admin',
      name: SUPER_ADMIN_NAME,
      employeeId: SUPER_ADMIN_EMPLOYEE_ID,
      redirect: 'admin',
      appUrl: ScriptApp.getService().getUrl() || ''
    };
  }

  var master = getEmployeeMasterSheet_(false);
  var rowNumber = findEmployeeRow_(id);
  if (!master || !rowNumber) {
    throw new Error('Employee ID or password is incorrect.');
  }

  var record = readEmployeeRecord_(master, rowNumber);
  if (!record || record.employeeId !== id) {
    throw new Error('Employee ID or password is incorrect.');
  }
  if (record.status !== 'Active') {
    throw new Error('This employee account is inactive. Please contact the Super Admin.');
  }

  var lockedUntil = record.lockedUntil instanceof Date
    ? record.lockedUntil
    : (record.lockedUntil ? new Date(record.lockedUntil) : null);

  if (lockedUntil && !isNaN(lockedUntil.getTime()) && lockedUntil.getTime() > now.getTime()) {
    var minutes = Math.max(1, Math.ceil((lockedUntil.getTime() - now.getTime()) / 60000));
    throw new Error('Too many failed attempts. Try again in ' + minutes + ' minute(s).');
  }

  var storedPassword = String(record.password || '');
  var passwordMatches = storedPassword === pass;

  // Backward compatibility: older records may still contain a SHA-256 hash.
  if (!passwordMatches && /^[a-f0-9]{64}$/i.test(storedPassword)) {
    passwordMatches = hashPassword_(pass) === storedPassword;
  }

  if (!passwordMatches) {
    var attempts = Number(record.failedAttempts || 0) + 1;
    var lock = '';

    if (attempts >= 5) {
      lock = new Date(now.getTime() + 15 * 60000);
      attempts = 0;
    }

    updateEmployeeSecurity_(master, rowNumber, attempts, lock, record.lastLogin, record.createdAt);
    throw new Error(
      lock
        ? 'Too many failed attempts. Account locked for 15 minutes.'
        : 'Employee ID or password is incorrect.'
    );
  }

  var token = createSession_(
    record.userId,
    record.employeeId,
    record.email,
    record.name,
    record.role
  );

  updateEmployeeSecurity_(master, rowNumber, 0, '', now, record.createdAt);

  return {
    success: true,
    token: token,
    role: record.role,
    name: record.name,
    employeeId: record.employeeId,
    redirect: record.role === 'Admin' ? 'admin' : 'employee',
    appUrl: ScriptApp.getService().getUrl() || ''
  };
}

function createSession_(userId, employeeId, email, name, role) {
  var token = Utilities.getUuid().replace(/-/g, '') +
    Utilities.getUuid().replace(/-/g, '').slice(0, 16);

  var session = {
    userId: String(userId || ''),
    employeeId: String(employeeId || ''),
    email: String(email || ''),
    name: String(name || ''),
    role: String(role || 'Employee'),
    issuedAt: Date.now()
  };

  CacheService.getScriptCache().put(
    SESSION_PREFIX + token,
    JSON.stringify(session),
    SESSION_TTL_SECONDS
  );

  return token;
}

function logoutSession(sessionToken) {
  var token = String(sessionToken || '').trim();
  if (token) CacheService.getScriptCache().remove(SESSION_PREFIX + token);
  return {success: true};
}

function getSession_(sessionToken) {
  var token = String(sessionToken || '').trim();
  if (!token || token.length < 24 || token.length > 100) return null;

  var raw = CacheService.getScriptCache().get(SESSION_PREFIX + token);
  if (!raw) return null;

  try {
    return JSON.parse(raw);
  } catch (e) {
    return null;
  }
}

function hashPassword_(password) {
  var bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(password),
    Utilities.Charset.UTF_8
  );

  return bytes.map(function(b) {
    var value = b < 0 ? b + 256 : b;
    return ('0' + value.toString(16)).slice(-2);
  }).join('');
}

function ensureIdentityDirectory_() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('VNS_IDENTITY_DIRECTORY_V3') === '1') {
    ensureEmployeeMasterSheet_();
    return;
  }

  migrateLegacyIdentityDirectory_();
  props.setProperty('VNS_IDENTITY_DIRECTORY_V3', '1');
}

function ensureEmployeeMasterSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(EMPLOYEE_MASTER_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(EMPLOYEE_MASTER_SHEET);
    sheet.setFrozenRows(1);
  }

  // Preserve the existing Employee_Master layout. Only add headers to a truly blank sheet;
  // never rewrite an existing header row or move employee data between columns.
  if (sheet.getLastRow() === 0 || sheet.getRange(1, 1).isBlank()) {
    sheet.getRange(1, 1, 1, EMPLOYEE_HEADERS.length).setValues([EMPLOYEE_HEADERS]);
    sheet.getRange(1, 1, 1, EMPLOYEE_HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  sheet.showColumns(9);

  return sheet;
}

function migrateLegacyIdentityDirectory_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var master = ensureEmployeeMasterSheet_();
  var existingIds = getEmployeeMasterIds_(master);

  // Migrate the previous Employee_<ID> structure first.
  ss.getSheets().slice().forEach(function(sheet) {
    var name = sheet.getName();
    if (name.indexOf('Employee_') !== 0 || name === EMPLOYEE_MASTER_SHEET) return;
    if (sheet.getLastRow() < 2) return;

    var row = sheet.getRange(2, 1, 1, Math.min(EMPLOYEE_HEADERS.length, sheet.getLastColumn())).getValues()[0];
    var employeeId = String(row[0] || '').trim().toUpperCase();
    if (!employeeId || existingIds[employeeId]) return;

    master.appendRow(normalizeEmployeeRow_(row));
    existingIds[employeeId] = true;
  });

  // Migrate the earlier three-sheet identity model when those tabs still exist.
  var usersSheet = ss.getSheetByName('Workspace_Users');
  var profileSheet = ss.getSheetByName('Workspace_User_Profile');
  var authSheet = ss.getSheetByName('Workspace_Auth');
  var profiles = {};
  var credentials = {};

  if (profileSheet && profileSheet.getLastRow() >= 2) {
    profileSheet.getRange(2, 1, profileSheet.getLastRow() - 1, 5).getValues().forEach(function(row) {
      var userId = String(row[0] || '').trim();
      if (!userId) return;
      profiles[userId] = {
        employeeId: String(row[1] || '').trim().toUpperCase(),
        department: String(row[2] || 'General').trim() || 'General',
        managerId: String(row[3] || '').trim()
      };
    });
  }

  if (authSheet && authSheet.getLastRow() >= 2) {
    authSheet.getRange(2, 1, authSheet.getLastRow() - 1, 9).getValues().forEach(function(row) {
      var userId = String(row[1] || '').trim();
      if (!userId) return;
      credentials[userId] = {
        employeeId: String(row[0] || '').trim().toUpperCase(),
        password: String(row[2] || ''),
        status: String(row[3] || 'Active'),
        failedAttempts: Number(row[4] || 0),
        lockedUntil: row[5] || '',
        lastLogin: row[6] || '',
        createdAt: row[7] || '',
        updatedAt: row[8] || ''
      };
    });
  }

  if (usersSheet && usersSheet.getLastRow() >= 2) {
    usersSheet.getRange(2, 1, usersSheet.getLastRow() - 1, 7).getValues().forEach(function(row) {
      var userId = String(row[0] || '').trim();
      var role = String(row[3] || 'Employee').trim();
      if (!userId || role === 'Admin') return;

      var profile = profiles[userId] || {};
      var credential = credentials[userId] || {};
      var employeeId = String(profile.employeeId || credential.employeeId || '').trim().toUpperCase();
      if (!employeeId || existingIds[employeeId]) return;

      master.appendRow([
        employeeId,
        userId,
        String(row[1] || '').trim(),
        String(row[2] || '').trim().toLowerCase(),
        role,
        String(row[4] || credential.status || 'Active'),
        String(profile.department || 'General'),
        String(profile.managerId || ''),
        String(credential.password || ''),
        Number(credential.failedAttempts || 0),
        credential.lockedUntil || '',
        credential.lastLogin || '',
        credential.createdAt || row[5] || new Date(),
        credential.updatedAt || row[6] || new Date()
      ]);
      existingIds[employeeId] = true;
    });
  }

  // Remove obsolete identity tabs and old per-employee tabs.
  ss.getSheets().slice().forEach(function(sheet) {
    var name = sheet.getName();
    if (name === EMPLOYEE_MASTER_SHEET) return;

    var obsoleteIdentity = LEGACY_IDENTITY_SHEETS.indexOf(name) !== -1;
    var oldEmployeeSheet = name.indexOf('Employee_') === 0;

    if ((obsoleteIdentity || oldEmployeeSheet) && ss.getSheets().length > 1) {
      ss.deleteSheet(sheet);
    }
  });
}

function normalizeEmployeeRow_(row) {
  var values = row.slice(0, EMPLOYEE_HEADERS.length);
  while (values.length < EMPLOYEE_HEADERS.length) values.push('');
  return values;
}

function getEmployeeMasterSheet_(create) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(EMPLOYEE_MASTER_SHEET);
  if (!sheet && create) sheet = ensureEmployeeMasterSheet_();
  return sheet;
}

function getEmployeeMasterIds_(sheet) {
  var ids = {};
  if (!sheet || sheet.getLastRow() < 2) return ids;

  sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().forEach(function(row) {
    var id = String(row[0] || '').trim().toUpperCase();
    if (id) ids[id] = true;
  });

  return ids;
}

function findEmployeeRow_(employeeId) {
  var id = String(employeeId || '').trim().toUpperCase();
  var sheet = getEmployeeMasterSheet_(false);
  if (!sheet || sheet.getLastRow() < 2) return 0;

  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0] || '').trim().toUpperCase() === id) return i + 2;
  }
  return 0;
}

function createOrUpdateEmployeeSheet(data) {
  var employeeId = String(data.employeeId || '').trim().toUpperCase();
  if (!employeeId) throw new Error('Employee ID is required.');

  var sheet = ensureEmployeeMasterSheet_();
  var rowNumber = findEmployeeRow_(employeeId);
  var values = [
    employeeId,
    String(data.userId || Utilities.getUuid()),
    String(data.name || ''),
    String(data.email || '').toLowerCase(),
    String(data.role || 'Employee'),
    String(data.status || 'Active'),
    String(data.department || 'General'),
    String(data.managerId || ''),
    String(data.password || ''),
    Number(data.failedAttempts || 0),
    data.lockedUntil || '',
    data.lastLogin || '',
    data.createdAt || new Date(),
    data.updatedAt || new Date()
  ];

  if (rowNumber) {
    sheet.getRange(rowNumber, 1, 1, EMPLOYEE_HEADERS.length).setValues([values]);
  } else {
    sheet.appendRow(values);
  }

  return sheet;
}

function readEmployeeRecord_(sheet, rowNumber) {
  if (!sheet || !rowNumber) return null;
  var row = sheet.getRange(rowNumber, 1, 1, EMPLOYEE_HEADERS.length).getValues()[0];

  return {
    employeeId: String(row[0] || '').trim().toUpperCase(),
    userId: String(row[1] || '').trim(),
    name: String(row[2] || '').trim(),
    email: String(row[3] || '').trim().toLowerCase(),
    role: String(row[4] || 'Employee').trim(),
    status: String(row[5] || 'Active').trim(),
    department: String(row[6] || 'General').trim(),
    managerId: String(row[7] || '').trim(),
    password: String(row[8] || ''),
    failedAttempts: Number(row[9] || 0),
    lockedUntil: row[10] || '',
    lastLogin: row[11] || '',
    createdAt: row[12] || '',
    updatedAt: row[13] || ''
  };
}

function updateEmployeeSecurity_(sheet, rowNumber, failedAttempts, lockedUntil, lastLogin, createdAt) {
  sheet.getRange(rowNumber, 10, 1, 5).setValues([[
    Number(failedAttempts || 0),
    lockedUntil || '',
    lastLogin || '',
    createdAt || new Date(),
    new Date()
  ]]);
}

function getAdminDashboardData(sessionToken) {
  var ctx = requirePermission_('users', sessionToken);
  if (ctx.role !== 'Admin') {
    throw new Error('Only Super Admin can open the Admin Control Center.');
  }

  var report = getReportsData({includeTests:'false'}, sessionToken);
  var users = getWorkspaceUsers_().rows;

  var employeeMap = {};
  users.forEach(function(u) {
    employeeMap[u.email.toLowerCase()] = {
      id: u.id,
      name: u.name,
      email: u.email,
      role: u.role,
      status: u.status,
      employeeId: u.employeeId || '',
      department: u.department || 'General',
      total: 0,
      sent: 0,
      failed: 0
    };
  });

  (report.employees || []).forEach(function(a) {
    if (!a.employeeEmail) return;
    var key = String(a.employeeEmail).toLowerCase();
    if (!employeeMap[key]) {
      employeeMap[key] = {
        id: '',
        name: a.employeeName || 'Unknown Employee',
        email: a.employeeEmail,
        role: 'Unknown',
        status: 'Unknown',
        employeeId: '',
        department: 'Unassigned',
        total: 0,
        sent: 0,
        failed: 0
      };
    }
    employeeMap[key].total = Number(a.total || 0);
    employeeMap[key].sent = Number(a.sent || 0);
    employeeMap[key].failed = Number(a.failed || 0);
  });

  var employeeRows = Object.keys(employeeMap).map(function(k) {
    var r = employeeMap[k];
    r.successRate = r.total ? Math.round(r.sent / r.total * 1000) / 10 : 0;
    return r;
  }).sort(function(a, b) {
    return b.sent - a.sent || a.name.localeCompare(b.name);
  });

  var recent = (report.recent || []).slice(0, 20).map(function(a) {
    return {
      timestamp: a.timestampLabel || a.timestamp,
      employeeEmail: a.employeeEmail || '',
      employeeName: a.employeeName || 'Unattributed / Existing Tool Data',
      campaignName: a.campaignName || 'Existing Tool Data',
      recipientName: a.recipientName || '',
      recipientEmail: a.recipientEmail || '',
      status: a.status || '',
      source: a.source || ''
    };
  });

  var campaignRows = (report.campaigns || []).slice(0, 12).map(function(c) {
    return {
      name: c.campaignName || 'Unassigned',
      total: c.total || 0,
      sent: c.sent || 0,
      failed: c.failed || 0,
      successRate: c.successRate || 0
    };
  });

  return {
    user: ctx,
    summary: {
      users: users.length + 1,
      activeUsers: users.filter(function(u) { return u.status === 'Active'; }).length + 1,
      inactiveUsers: users.filter(function(u) { return u.status !== 'Active'; }).length,
      totalEmails: report.summary.total,
      sent: report.summary.sent,
      failed: report.summary.failed,
      successRate: report.summary.successRate
    },
    employees: employeeRows,
    campaigns: campaignRows,
    recent: recent,
    legacyEmails: (report.employees || []).filter(function(a) {
      return !a.employeeEmail && a.employeeName === 'Unattributed / Existing Tool Data';
    }).reduce(function(n, a) {
      return n + Number(a.total || 0);
    }, 0),
    updatedAt: formatNow_()
  };
}

function saveEmployeeCredentials(data, sessionToken) {
  var ctx = requirePermission_('users', sessionToken);
  if (ctx.role !== 'Admin') {
    throw new Error('Only Super Admin can manage login credentials.');
  }

  data = data || {};
  var employeeId = String(data.employeeId || '').trim().toUpperCase();
  var password = String(data.password || '');

  if (!/^[A-Z0-9_-]{3,40}$/.test(employeeId)) {
    throw new Error('Employee ID must be 3-40 characters: A-Z, 0-9, _ or -.');
  }
  if (password.length < 8 || password.length > 128) {
    throw new Error('Password must be 8-128 characters.');
  }

  var sheet = getEmployeeMasterSheet_(false);
  var rowNumber = findEmployeeRow_(employeeId);
  if (!sheet || !rowNumber) throw new Error('Employee account not found.');

  sheet.getRange(rowNumber, 9).setValue(password);
  sheet.getRange(rowNumber, 10, 1, 5).setValues([[0, '', '', sheet.getRange(rowNumber, 13).getValue() || new Date(), new Date()]]);

  var record = readEmployeeRecord_(sheet, rowNumber);
  return {
    success: true,
    message: 'Login credentials saved for ' + record.name + '.'
  };
}

function getLoginDirectory(sessionToken) {
  var ctx = requirePermission_('users', sessionToken);
  if (ctx.role !== 'Admin') {
    throw new Error('Only Super Admin can view login directory.');
  }

  return getWorkspaceUsers_().rows;
}

/* --------------------------------------------------------------------------
 * WORKSPACE USERS / ROLES
 * The directory is now sheet-per-employee. No central password table exists.
 * -------------------------------------------------------------------------- */
var USER_ROLES = ['Admin', 'Manager', 'Employee', 'Viewer'];
var USER_STATUSES = ['Active', 'Inactive'];

function getUserProfiles_() {
  var profiles = {};
  getWorkspaceUsers_().rows.forEach(function(user) {
    profiles[user.id] = {
      userId: user.id,
      employeeId: user.employeeId || '',
      department: user.department || 'General',
      managerId: user.managerId || '',
      updatedAt: user.updatedAt || ''
    };
  });
  return profiles;
}

function saveUserProfile_(userId, employeeId, department, managerId) {
  if (!userId || !employeeId) return;

  var sheet = getEmployeeMasterSheet_(false);
  var rowNumber = findEmployeeRow_(employeeId);
  if (!sheet || !rowNumber) return;

  sheet.getRange(rowNumber, 7, 1, 2).setValues([[
    String(department || 'General'),
    String(managerId || '')
  ]]);
  sheet.getRange(rowNumber, 14).setValue(new Date());
}

function getWorkspaceUsers(sessionToken) {
  requirePermission_('users', sessionToken);
  return getWorkspaceUsers_();
}

function getWorkspaceUsers_() {
  ensureIdentityDirectory_();

  var rows = [{
    id: SUPER_ADMIN_USER_ID,
    name: SUPER_ADMIN_NAME,
    email: SUPER_ADMIN_EMAIL,
    role: 'Admin',
    status: 'Active',
    employeeId: SUPER_ADMIN_EMPLOYEE_ID,
    department: 'Administration',
    managerId: '',
    createdAt: '',
    updatedAt: ''
  }];

  var sheet = getEmployeeMasterSheet_(false);
  if (sheet && sheet.getLastRow() >= 2) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, EMPLOYEE_HEADERS.length).getValues().forEach(function(row) {
      var record = {
        id: String(row[1] || ''),
        name: String(row[2] || ''),
        email: String(row[3] || '').toLowerCase(),
        role: String(row[4] || 'Employee'),
        status: String(row[5] || 'Active'),
        employeeId: String(row[0] || '').toUpperCase(),
        department: String(row[6] || 'General'),
        managerId: String(row[7] || ''),
        createdAt: row[12] || '',
        updatedAt: row[13] || ''
      };
      if (record.id && record.employeeId) rows.push(record);
    });
  }

  rows.sort(function(a, b) {
    if (a.role === 'Admin') return -1;
    if (b.role === 'Admin') return 1;
    return a.name.localeCompare(b.name);
  });

  return {
    rows: rows,
    total: rows.length,
    active: rows.filter(function(r) { return r.status === 'Active'; }).length,
    inactive: rows.filter(function(r) { return r.status !== 'Active'; }).length,
    roles: USER_ROLES
  };
}

function requireAdminUser_(sessionToken) {
  var ctx = getWorkspaceAccessContext(sessionToken);
  if (!ctx.allowed || ctx.role !== 'Admin') {
    throw new Error('Only Workspace Admin can manage users.');
  }
  return ctx;
}

function saveWorkspaceUser(user, sessionToken) {
  var adminCtx = requireAdminUser_(sessionToken);
  user = user || {};

  var name = String(user.name || '').trim();
  var email = String(user.email || '').trim().toLowerCase();
  var role = String(user.role || 'Employee').trim();
  var status = String(user.status || 'Active').trim();
  var employeeId = String(user.employeeId || '').trim().toUpperCase();
  var password = String(user.password || '');
  var department = String(user.department || 'General').trim() || 'General';
  var managerId = String(user.managerId || '').trim();
  var autoGenerateEmployeeId = user.autoGenerateEmployeeId === true || String(user.autoGenerateEmployeeId).toLowerCase() === 'true';

  if (!name) throw new Error('User name is required.');
  if (name.length > 120) throw new Error('User name is too long.');
  if (!isValidEmail_(email)) throw new Error('Enter a valid user email.');
  if (USER_ROLES.indexOf(role) === -1 || role === 'Admin') throw new Error('New employee accounts must use a non-admin role.');
  if (USER_STATUSES.indexOf(status) === -1) throw new Error('Invalid user status.');
  if (password.length < 8 || password.length > 128) throw new Error('Set an employee password of 8-128 characters.');

  // Auto ID is opt-in. If it is off, the Admin supplied ID is mandatory.
  if (autoGenerateEmployeeId) {
    employeeId = generateNextEmployeeId_();
  }

  if (!employeeId) throw new Error('Employee ID is required, or enable Auto Generate ID.');
  if (!/^[A-Z0-9_-]{3,40}$/.test(employeeId)) {
    throw new Error('Employee ID must be 3-40 characters: A-Z, 0-9, _ or -.');
  }

  var sheet = ensureEmployeeMasterSheet_();
  var existingRow = findEmployeeRow_(employeeId);
  var existingUserId = String(user.id || '').trim();
  var now = new Date();

  if (existingRow) {
    var existingRecord = readEmployeeRecord_(sheet, existingRow);
    if (!existingUserId || existingUserId !== existingRecord.userId) {
      throw new Error('This Employee ID is already assigned. Use another ID.');
    }

    if (existingRecord.email !== email) {
      throw new Error('Email change is not allowed through the employee creation flow.');
    }

    sheet.getRange(existingRow, 3, 1, 6).setValues([[
      name,
      email,
      role,
      status,
      department,
      managerId
    ]]);
    sheet.getRange(existingRow, 9).setValue(hashPassword_(password));
    sheet.getRange(existingRow, 10, 1, 3).setValues([[0, '', '']]);
    sheet.getRange(existingRow, 14).setValue(now);

    return {
      success: true,
      id: existingRecord.userId,
      employeeId: employeeId,
      message: 'Employee updated successfully.',
      users: getWorkspaceUsers_()
    };
  }

  var userId = existingUserId || Utilities.getUuid();
  createOrUpdateEmployeeSheet({
    employeeId: employeeId,
    userId: userId,
    name: name,
    email: email,
    role: role,
    status: status,
    department: department,
    managerId: managerId,
    password: password,
    failedAttempts: 0,
    lockedUntil: '',
    lastLogin: '',
    createdAt: now,
    updatedAt: now
  });

  return {
    success: true,
    id: userId,
    employeeId: employeeId,
    message: 'Employee added successfully.',
    users: getWorkspaceUsers_()
  };
}

function generateNextEmployeeId_() {
  var maxNumber = 0;
  var sheet = getEmployeeMasterSheet_(false);
  if (!sheet || sheet.getLastRow() < 2) return 'EMP001';

  sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().forEach(function(row) {
    var id = String(row[0] || '').trim().toUpperCase();
    var match = id.match(/^EMP(\d+)$/);
    if (match) maxNumber = Math.max(maxNumber, Number(match[1]));
  });

  return 'EMP' + String(maxNumber + 1).padStart(3, '0');
}

function setWorkspaceUserStatus(userId, status, sessionToken) {
  var adminCtx = requireAdminUser_(sessionToken);
  status = String(status || '').trim();
  if (USER_STATUSES.indexOf(status) === -1) throw new Error('Invalid user status.');
  if (String(userId) === SUPER_ADMIN_USER_ID) {
    if (status !== 'Active') throw new Error('The Super Admin account cannot be disabled.');
    return {success:true, message:'Super Admin status is Active.', users:getWorkspaceUsers_()};
  }

  var users = getWorkspaceUsers_().rows;
  var user = users.filter(function(r) { return String(r.id) === String(userId); })[0];
  if (!user) throw new Error('User not found.');

  var sheet = getEmployeeMasterSheet_(false);
  var rowNumber = findEmployeeRow_(user.employeeId);
  if (!sheet || !rowNumber) throw new Error('Employee record not found.');

  sheet.getRange(rowNumber, 6).setValue(status);
  sheet.getRange(rowNumber, 14).setValue(new Date());

  return {success:true, message:'User status updated.', users:getWorkspaceUsers_()};
}

function getWorkspaceAccessContext(sessionToken) {
  var token = String(sessionToken || '').trim();
  if (token) {
    var session = getSession_(token);
    if (!session) {
      return {configured:true, allowed:false, email:'', name:'Unauthorized User', role:'None', permissions:{}};
    }

    if (session.role === 'Admin' && session.userId === SUPER_ADMIN_USER_ID) {
      return {
        configured: true,
        allowed: true,
        email: SUPER_ADMIN_EMAIL,
        name: SUPER_ADMIN_NAME,
        employeeId: SUPER_ADMIN_EMPLOYEE_ID,
        userId: SUPER_ADMIN_USER_ID,
        role: 'Admin',
        permissions: getRolePermissions_('Admin')
      };
    }

    var users = getWorkspaceUsers_().rows;
    var user = users.filter(function(r) {
      return r.status === 'Active' && String(r.id) === String(session.userId);
    })[0];

    if (!user) {
      return {configured:true, allowed:false, email:session.email || '', name:'Inactive User', role:'None', permissions:{}};
    }

    return {
      configured:true,
      allowed:true,
      email:user.email,
      name:user.name,
      employeeId:session.employeeId || user.employeeId || '',
      userId:user.id,
      role:user.role,
      permissions:getRolePermissions_(user.role)
    };
  }

  var identity = getCurrentIdentity_();
  if (!identity) {
    return {
      configured:true,
      allowed:true,
      email:SUPER_ADMIN_EMAIL,
      name:SUPER_ADMIN_NAME,
      employeeId:SUPER_ADMIN_EMPLOYEE_ID,
      userId:SUPER_ADMIN_USER_ID,
      role:'Admin',
      permissions:getRolePermissions_('Admin')
    };
  }

  var usersLegacy = getWorkspaceUsers_().rows;
  var userLegacy = usersLegacy.filter(function(r) {
    return r.status === 'Active' && r.email.toLowerCase() === identity;
  })[0];

  if (!userLegacy) {
    return {configured:true, allowed:false, email:identity, name:'Unauthorized User', role:'None', permissions:{}};
  }

  return {
    configured:true,
    allowed:true,
    email:userLegacy.email,
    name:userLegacy.name,
    employeeId:userLegacy.employeeId || '',
    userId:userLegacy.id,
    role:userLegacy.role,
    permissions:getRolePermissions_(userLegacy.role)
  };
}

function getRolePermissions_(role) {
  var all = ['dashboard','send','recipients','templates','campaigns','reports','settings','users'];
  if (role === 'Admin') return all.reduce(function(o,k){o[k]=true;return o;},{});
  if (role === 'Manager') return {dashboard:true,send:true,recipients:true,templates:true,campaigns:true,reports:true,settings:false,users:false};
  if (role === 'Employee') return {dashboard:true,send:true,recipients:true,templates:true,campaigns:true,reports:false,settings:false,users:false};
  if (role === 'Viewer') return {dashboard:true,send:false,recipients:false,templates:false,campaigns:true,reports:true,settings:false,users:false};
  return {};
}



/* --------------------------------------------------------------------------
 * Google Drive Picker + local-file upload endpoints.
 * Uses existing session/permission guards and does not change the Tool columns.
 * Script Property required: VNS_DRIVE_PICKER_API_KEY
 * -------------------------------------------------------------------------- */
function getDrivePickerConfig(purpose, sessionToken) {
  var ctx = getWorkspaceAccessContext(sessionToken);
  if (!ctx || !ctx.allowed) throw new Error('Your session has expired. Please sign in again.');
  if (purpose !== 'templateDoc' && purpose !== 'templateAttachment' && purpose !== 'campaignAttachment') {
    throw new Error('Invalid Drive Picker purpose.');
  }
  var apiKey = PropertiesService.getScriptProperties().getProperty('VNS_DRIVE_PICKER_API_KEY');
  if (!apiKey) {
    return { apiKey: '', oauthToken: '' };
  }
  // Google Picker uses the active Apps Script OAuth token. Ensure Drive scopes are
  // present in appsscript.json and the deployment has been re-authorized.
  return { apiKey: apiKey, oauthToken: ScriptApp.getOAuthToken() };
}

function uploadWorkspaceAttachment(fileData, purpose, sessionToken) {
  var ctx = getWorkspaceAccessContext(sessionToken);
  if (!ctx || !ctx.allowed) throw new Error('Your session has expired. Please sign in again.');
  if (purpose !== 'templateAttachment' && purpose !== 'campaignAttachment' && purpose !== 'templateDoc') {
    throw new Error('Invalid upload purpose.');
  }
  if (!fileData || !fileData.name || !fileData.base64) throw new Error('Please select a valid file.');
  var name = String(fileData.name).replace(/[\\\/:*?"<>|#%{}~&]/g, '_').trim();
  if (!name) name = 'Workspace attachment';
  var mimeType = String(fileData.mimeType || 'application/octet-stream');
  var base64 = String(fileData.base64 || '');
  // Browser UI limits uploads to 10 MB; server validates again.
  if (base64.length > 15 * 1024 * 1024) throw new Error('File is larger than the 10 MB upload limit.');
  var bytes;
  try { bytes = Utilities.base64Decode(base64); } catch (e) { throw new Error('Uploaded file data is invalid.'); }
  if (bytes.length > 10 * 1024 * 1024) throw new Error('File is larger than the 10 MB upload limit.');
  var blob = Utilities.newBlob(bytes, mimeType, name);
  var file = DriveApp.createFile(blob);
  return { success: true, id: file.getId(), name: file.getName(), url: 'https://drive.google.com/open?id=' + file.getId(), mimeType: file.getMimeType() };
}
