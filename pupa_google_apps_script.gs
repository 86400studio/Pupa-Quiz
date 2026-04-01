
const LOG_SHEET_NAME = 'Quiz_Log';
const SUMMARY_SHEET_NAME = 'Summary';

const FIXED_HEADERS = [
  'server_timestamp',
  'timestamp_iso_client',
  'event_type',
  'session_id',
  'flow_stage',
  'quiz_id',
  'quiz_title',
  'quiz_number',
  'result_key',
  'result_label',
  'tied_result_keys',
  'journey_quiz_order',
  'prompt_target_quiz_id',
  'recommend_to_friends',
  'email',
  'first_name',
  'consent',
  'answer_1_key', 'answer_1_text',
  'answer_2_key', 'answer_2_text',
  'answer_3_key', 'answer_3_text',
  'answer_4_key', 'answer_4_text',
  'answer_5_key', 'answer_5_text',
  'answer_6_key', 'answer_6_text',
  'answer_7_key', 'answer_7_text',
  'answer_8_key', 'answer_8_text',
  'answer_9_key', 'answer_9_text',
  'answer_10_key', 'answer_10_text',
  'answers_json',
  'page_url',
  'referrer',
  'user_agent',
  'note'
];

function doGet(e) {
  const payload = normalizePayload_(e);
  return json_({ ok: true, mode: 'GET', received: payload });
}

function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getOrCreateSheet_(ss, LOG_SHEET_NAME);
    ensureSummarySheet_(ss);

    const payload = normalizePayload_(e);
    payload.server_timestamp = new Date().toISOString();

    const headers = ensureHeaders_(sheet, payload);
    const row = headers.map(function(header) {
      return Object.prototype.hasOwnProperty.call(payload, header) ? payload[header] : '';
    });

    sheet.appendRow(row);
    return json_({ ok: true });
  } catch (error) {
    return json_({ ok: false, error: String(error) });
  }
}

function normalizePayload_(e) {
  const payload = {};

  if (e && e.parameter) {
    Object.keys(e.parameter).forEach(function(key) {
      payload[key] = e.parameter[key];
    });
  }

  if (e && e.postData && e.postData.contents) {
    const contentType = String(e.postData.type || '');
    if (contentType.indexOf('application/json') !== -1) {
      try {
        const json = JSON.parse(e.postData.contents);
        Object.keys(json).forEach(function(key) {
          payload[key] = json[key];
        });
      } catch (err) {}
    }
  }

  return payload;
}

function getOrCreateSheet_(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  return sheet;
}

function ensureHeaders_(sheet, payload) {
  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  const existingHeaders = sheet.getLastRow() > 0
    ? sheet.getRange(1, 1, 1, lastColumn).getValues()[0].filter(String)
    : [];

  let headers = existingHeaders.slice();

  if (!headers.length) {
    headers = FIXED_HEADERS.slice();
  }

  Object.keys(payload).forEach(function(key) {
    if (headers.indexOf(key) === -1) {
      headers.push(key);
    }
  });

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  return headers;
}

function ensureSummarySheet_(ss) {
  let sheet = ss.getSheetByName(SUMMARY_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SUMMARY_SHEET_NAME);
    sheet.getRange('A1:B8').setValues([
      ['Metric', 'Value'],
      ['Total quiz starts', '=COUNTIF(Quiz_Log!C:C,"quiz_started")'],
      ['Total quiz completions', '=COUNTIF(Quiz_Log!C:C,"quiz_completed")'],
      ['Unique participants', '=COUNTA(UNIQUE(FILTER(Quiz_Log!D:D,Quiz_Log!D:D<>"")))'],
      ['Unique emails captured', '=COUNTA(UNIQUE(FILTER(Quiz_Log!O:O,Quiz_Log!O:O<>"")))'],
      ['Referral yes', '=COUNTIFS(Quiz_Log!C:C,"exit_feedback",Quiz_Log!N:N,"Yes")'],
      ['Website clicks', '=COUNTIF(Quiz_Log!C:C,"website_click")'],
      ['Last event timestamp', '=MAX(Quiz_Log!A:A)']
    ]);
    sheet.getRange('A1:B1').setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, 2);
  }
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
