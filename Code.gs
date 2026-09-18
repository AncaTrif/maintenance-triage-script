/**
 * Maintenance Triage Automation
 * Bound Apps Script for the shared Maintenance spreadsheet (Vendors / Triage / Tracker).
 *
 * Secrets (Claude API key, Slack webhook URL, Dana's Slack member ID) are never stored in
 * code or in the spreadsheet — they live in Script Properties. Set them via the
 * "Triage Tools" menu after opening the Sheet, or directly under
 * Project Settings > Script Properties in the Apps Script editor.
 */

// ===== CONFIGURATION =====

const SHEET_TRIAGE = 'Triage';
const SHEET_TRACKER = 'Tracker';
const SHEET_VENDORS = 'Vendors';
const SHEET_CONFIG = 'Config';

const TRIAGE_HEADERS = [
  'Timestamp', 'Unit #', 'Tenant Name', 'Tenant Text', 'Source (Portal/Phone)',
  'Suggested Category', 'Suggested Priority', 'Suggested Vendor',
  'Confidence (Low/Medium/High)', 'Confidence Reason', 'Status'
];
const STATUS_OPTIONS = ['', 'Approved', 'Waiting on Tenant Callback', 'Skip'];
const VALIDATION_ROWS = 2000; // rows to pre-apply the Status dropdown to at setup time

const DEFAULT_GMAIL_LABEL = 'Maintenance-Portal';
const GMAIL_PROCESSED_LABEL = 'Triage-Processed';

const CLAUDE_MODEL = 'claude-sonnet-5';
const CLAUDE_API_URL = 'https://api.anthropic.com/v1/messages';

const PROP_CLAUDE_API_KEY = 'CLAUDE_API_KEY';
const PROP_SLACK_WEBHOOK_URL = 'SLACK_WEBHOOK_URL';
const PROP_DANA_SLACK_ID = 'DANA_SLACK_ID';
const PROP_GMAIL_LABEL = 'GMAIL_LABEL';
const PROP_TRIAGE_GUIDE_DOC_ID = 'TRIAGE_GUIDE_DOC_ID';

// ===== MENU =====

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Triage Tools')
    .addItem('Run setup (create Triage/Config tabs)', 'setupProject')
    .addItem('Install 15-minute trigger', 'installTriggers')
    .addSeparator()
    .addItem('Set API keys / webhook', 'promptForSecrets')
    .addSeparator()
    .addItem('Scan Gmail now', 'scanGmailForNewTickets')
    .addItem('Process Triage rows now', 'processTriageRows')
    .addToUi();
}

// ===== SETUP =====

function setupProject() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const triage = getOrCreateSheet(ss, SHEET_TRIAGE);
  setupTriageSheet(triage);

  const config = getOrCreateSheet(ss, SHEET_CONFIG);
  setupConfigSheet(config);

  SpreadsheetApp.getUi().alert(
    'Setup complete. Next: open "Triage Tools > Set API keys / webhook", fill in the ' +
    'Config tab values, then run "Triage Tools > Install 15-minute trigger".'
  );
}

function getOrCreateSheet(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function setupTriageSheet(sheet) {
  sheet.getRange(1, 1, 1, TRIAGE_HEADERS.length).setValues([TRIAGE_HEADERS]);
  sheet.getRange(1, 1, 1, TRIAGE_HEADERS.length).setFontWeight('bold');
  sheet.setFrozenRows(1);

  const statusCol = TRIAGE_HEADERS.indexOf('Status') + 1;
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(STATUS_OPTIONS, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, statusCol, VALIDATION_ROWS, 1).setDataValidation(rule);
}

function setupConfigSheet(sheet) {
  if (sheet.getRange('A1').getValue() === '') {
    sheet.getRange('A1:B1').setValues([['Setting', 'Value']]);
    sheet.getRange('A1:B1').setFontWeight('bold');
    sheet.getRange('A2:B2').setValues([['Gmail label to watch', DEFAULT_GMAIL_LABEL]]);
    sheet.getRange('A3:B3').setValues([['Triage guide Google Doc ID', '']]);
    sheet.setColumnWidth(1, 220);
    sheet.setColumnWidth(2, 320);
  }
}

function installTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'scanAndTriage') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('scanAndTriage').timeBased().everyMinutes(15).create();
  SpreadsheetApp.getUi().alert('15-minute trigger installed.');
}

function promptForSecrets() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();

  promptForOneSecret(ui, props, PROP_CLAUDE_API_KEY, 'Claude API key');
  promptForOneSecret(ui, props, PROP_SLACK_WEBHOOK_URL, 'Slack incoming webhook URL');
  promptForOneSecret(ui, props, PROP_DANA_SLACK_ID, "Dana's Slack member ID (e.g. U01ABC2DEF, optional)");
}

function promptForOneSecret(ui, props, key, label) {
  const existing = props.getProperty(key);
  const hint = existing ? ' (already set — leave blank to keep it)' : '';
  const response = ui.prompt('Set ' + label + hint, ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  const value = response.getResponseText().trim();
  if (value !== '') props.setProperty(key, value);
}

// ===== TRIGGER ENTRY POINT =====

function scanAndTriage() {
  scanGmailForNewTickets();
  processTriageRows();
}

// ===== STEP 1: GMAIL SCAN =====

function scanGmailForNewTickets() {
  const label = getConfigValue('Gmail label to watch') || DEFAULT_GMAIL_LABEL;
  const processedLabel = GmailApp.getUserLabelByName(GMAIL_PROCESSED_LABEL)
    || GmailApp.createLabel(GMAIL_PROCESSED_LABEL);

  const threads = GmailApp.search('label:"' + label + '" -label:"' + GMAIL_PROCESSED_LABEL + '"', 0, 50);
  if (threads.length === 0) return;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_TRIAGE);
  if (!sheet) return;

  threads.forEach(function (thread) {
    thread.getMessages().forEach(function (message) {
      const ticket = parseTicketEmail(message.getSubject(), message.getPlainBody());
      sheet.appendRow([
        new Date(), ticket.unit, ticket.tenant, ticket.tenantText, 'Portal',
        '', '', '', '', '', ''
      ]);
    });
    thread.addLabel(processedLabel);
  });
}

/**
 * Subject: "New Maintenance Request – Unit [X] – [Property]"
 * Body lines: Property / Unit / Tenant / Category selected / Description / Submitted.
 * Category selected and Description may each be blank in the source email.
 */
function parseTicketEmail(subject, body) {
  const unit = firstMatch(body, /Unit\s*:\s*(.*)/i) || firstMatch(subject, /Unit\s*#?\s*([^–—-]+?)\s*[–—-]/i) || '';
  const tenant = firstMatch(body, /Tenant\s*:\s*(.*)/i) || '';
  const category = firstMatch(body, /Category selected\s*:\s*(.*)/i) || '';
  const description = firstMatch(body, /Description\s*:\s*([\s\S]*?)(?:\n\s*Submitted|$)/i) || '';

  let tenantText = description.trim();
  if (tenantText === '') {
    tenantText = category.trim() !== ''
      ? '(No description provided. Category selected in portal: ' + category.trim() + ')'
      : '(No description or category provided.)';
  }

  return { unit: unit.trim(), tenant: tenant.trim(), tenantText: tenantText };
}

function firstMatch(text, regex) {
  const m = text.match(regex);
  return m ? m[1] : null;
}

// ===== STEP 2: CLAUDE TRIAGE =====

function processTriageRows() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_TRIAGE);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const col = triageColumnIndexes();
  const apiKey = PropertiesService.getScriptProperties().getProperty(PROP_CLAUDE_API_KEY);
  if (!apiKey) {
    console.error('Claude API key not set. Run Triage Tools > Set API keys / webhook.');
    return;
  }

  const guideText = getTriageGuideText();
  const vendorsText = buildVendorsReference();

  const data = sheet.getRange(2, 1, lastRow - 1, TRIAGE_HEADERS.length).getValues();

  data.forEach(function (row, i) {
    const sheetRow = i + 2;
    const category = row[col.suggestedCategory];
    const tenantText = String(row[col.tenantText] || '').trim();
    if (category !== '') return; // already processed

    if (tenantText === '') {
      writeSuggestion(sheet, sheetRow, {
        category: 'Other / Unclear', priority: 'Low', vendor: 'Needs Manual Assignment',
        confidence: 'Low', reason: 'No tenant text was provided to classify.'
      });
      return;
    }

    const result = classifyWithClaude(apiKey, tenantText, guideText, vendorsText);
    if (!result) return; // leave blank, retried on next 15-minute run

    writeSuggestion(sheet, sheetRow, result);

    const unit = row[col.unit];
    if (result.priority === 'Emergency' || result.confidence === 'Low' || result.confidence === 'Medium') {
      postSlackAlert(unit, result.category, result.priority, result.reason);
    }
  });
}

function triageColumnIndexes() {
  return {
    unit: TRIAGE_HEADERS.indexOf('Unit #'),
    tenantText: TRIAGE_HEADERS.indexOf('Tenant Text'),
    suggestedCategory: TRIAGE_HEADERS.indexOf('Suggested Category')
  };
}

function writeSuggestion(sheet, row, result) {
  const startCol = TRIAGE_HEADERS.indexOf('Suggested Category') + 1;
  sheet.getRange(row, startCol, 1, 5).setValues([[
    result.category, result.priority, result.vendor, result.confidence, result.reason
  ]]);
}

function getTriageGuideText() {
  const docId = getConfigValue('Triage guide Google Doc ID');
  if (!docId) return '(No triage guide configured yet — using category judgment only.)';
  try {
    return DocumentApp.openById(docId).getBody().getText();
  } catch (e) {
    console.error('Could not open triage guide doc: ' + e);
    return '(Triage guide doc could not be loaded.)';
  }
}

function buildVendorsReference() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_VENDORS);
  if (!sheet || sheet.getLastRow() < 1) return '(No vendors configured.)';
  const values = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
  return values.map(function (r) { return r.join(' | '); }).join('\n');
}

/**
 * Sends ONLY the tenant's free text to Claude — never tenant name or unit number.
 * Guide + Vendors sheet are sent as non-personal reference context.
 */
function classifyWithClaude(apiKey, tenantText, guideText, vendorsText) {
  const systemPrompt = [
    'You triage residential maintenance requests for a property management company.',
    'Use the triage guide below to choose a category and priority, and the vendor list to suggest a vendor.',
    'Respond with ONLY a JSON object, no other text, in exactly this shape:',
    '{"category": "...", "priority": "Emergency|High|Medium|Low", "vendor": "...", "confidence": "Low|Medium|High", "reason": "one sentence"}',
    '',
    '=== TRIAGE GUIDE ===',
    guideText,
    '',
    '=== AVAILABLE VENDORS ===',
    vendorsText
  ].join('\n');

  const payload = {
    model: CLAUDE_MODEL,
    max_tokens: 300,
    system: systemPrompt,
    messages: [{ role: 'user', content: tenantText }]
  };

  const response = UrlFetchApp.fetch(CLAUDE_API_URL, {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    console.error('Claude API error ' + response.getResponseCode() + ': ' + response.getContentText());
    return null;
  }

  try {
    const body = JSON.parse(response.getContentText());
    const parsed = JSON.parse(body.content[0].text);
    return {
      category: parsed.category, priority: parsed.priority, vendor: parsed.vendor,
      confidence: parsed.confidence, reason: parsed.reason
    };
  } catch (e) {
    console.error('Could not parse Claude response: ' + e);
    return null;
  }
}

// ===== STEP 3: SLACK ALERTS =====

function postSlackAlert(unit, category, priority, reason) {
  const webhook = PropertiesService.getScriptProperties().getProperty(PROP_SLACK_WEBHOOK_URL);
  if (!webhook) {
    console.error('Slack webhook not set. Run Triage Tools > Set API keys / webhook.');
    return;
  }
  const danaId = PropertiesService.getScriptProperties().getProperty(PROP_DANA_SLACK_ID);
  const mention = danaId ? '<@' + danaId + '>' : '@Dana';

  const text = ':rotating_light: ' + mention + ' new maintenance ticket needs review\n' +
    '*Unit:* ' + unit + '\n' +
    '*Category:* ' + category + '\n' +
    '*Priority:* ' + priority + '\n' +
    '*Confidence reason:* ' + reason;

  UrlFetchApp.fetch(webhook, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ text: text }),
    muteHttpExceptions: true
  });
}

// ===== STEP 4: APPROVAL -> TRACKER =====

/**
 * Simple onEdit trigger. Fires once per actual edit, so an "Approved" row logs to
 * Tracker exactly once. Matches Tracker columns by header name so this never
 * assumes or changes Tracker's existing column order.
 */
function onEdit(e) {
  const sheet = e.range.getSheet();
  if (sheet.getName() !== SHEET_TRIAGE) return;

  const statusCol = TRIAGE_HEADERS.indexOf('Status') + 1;
  if (e.range.getColumn() !== statusCol || e.range.getRow() === 1) return;
  if (e.value !== 'Approved') return;

  const row = e.range.getRow();
  const rowValues = sheet.getRange(row, 1, 1, TRIAGE_HEADERS.length).getValues()[0];
  const col = triageColumnIndexes();

  appendToTracker({
    'Date Logged': new Date(),
    'Unit #': rowValues[col.unit],
    'Tenant Name': rowValues[TRIAGE_HEADERS.indexOf('Tenant Name')],
    'Source': rowValues[TRIAGE_HEADERS.indexOf('Source (Portal/Phone)')],
    'Category': rowValues[TRIAGE_HEADERS.indexOf('Suggested Category')],
    'Priority': rowValues[TRIAGE_HEADERS.indexOf('Suggested Priority')],
    'Assigned Vendor': rowValues[TRIAGE_HEADERS.indexOf('Suggested Vendor')],
    'Status': 'Sent to Vendor'
  });
}

function appendToTracker(valuesByHeader) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_TRACKER);
  if (!sheet) {
    console.error('Tracker sheet not found.');
    return;
  }
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const newRow = new Array(headers.length).fill('');

  headers.forEach(function (header, i) {
    const key = String(header).trim();
    if (Object.prototype.hasOwnProperty.call(valuesByHeader, key)) {
      newRow[i] = valuesByHeader[key];
    }
  });

  sheet.appendRow(newRow);
}

// ===== CONFIG HELPERS =====

function getConfigValue(settingName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CONFIG);
  if (!sheet) return '';
  const values = sheet.getRange(1, 1, sheet.getLastRow(), 2).getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim() === settingName) return String(values[i][1]).trim();
  }
  return '';
}
