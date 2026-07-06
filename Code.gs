const SHEET_ID = 'YOUR_SHEET_ID';
const SHEET_NAME = 'RSVP';

function doGet(e) {
  try {
    // Health check
    if (e && e.parameter && e.parameter.health === '1') {
      return ContentService
        .createTextOutput(JSON.stringify({ ok: true, time: new Date().toISOString() }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    const data = sheet.getDataRange().getValues();

    if (data.length < 2) {
      return ContentService
        .createTextOutput(JSON.stringify({ ok: true, data: [] }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const headers = data[0];
    const result = data.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    });

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, data: result }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    const body = JSON.parse(e.postData.contents);

    sheet.appendRow([
      body.id,
      body.firstName,
      body.lastName,
      body.email,
      body.status,
      body.adults,
      body.children,
      body.message || '',
      body.timestamp
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
