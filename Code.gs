/**
 * Google Apps Script — RSVP backend za vjenčanje
 * 
 * 1. Otvori https://script.google.com, novi projekt
 * 2. Zalijepi ovaj kod
 * 3. Zamijeni 'YOUR_SHEET_ID' sa ID-jem svoje Google tablice
 * 4. Spremi → Deploy → New deployment → Web app
 *    Execute as: Me, Access: Anyone
 * 5. Kopiraj URL i zalijepi u admin panel na stranici
 */

const SHEET_ID = 'YOUR_SHEET_ID';
const SHEET_NAME = 'RSVP';

function doGet() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  
  if (data.length < 2) {
    return ContentService
      .createTextOutput(JSON.stringify({ data: [] }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  const headers = data[0];
  const result = data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
  
  return ContentService
    .createTextOutput(JSON.stringify({ data: result }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
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
}
