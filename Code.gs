const SHEET_ID = "1-_1y4CGDyToIBLgNwXSgTRnUpvCWuV1DPNEQUOkuEt8";
const SHEET_NAME = "RSVP";
function doPost(e) {
  const data = JSON.parse(e.postData.contents || "{}");
  const name = String(data.name || "").trim();
  const answer = String(data.answer || "").trim();
  const token = String(data.token || "").trim();
  if (!name || !answer || !token) return ContentService.createTextOutput("bad request");
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) { sheet = ss.insertSheet(SHEET_NAME); sheet.appendRow(["Имя", "Ответ", "Token"]); }
  const rows = sheet.getDataRange().getValues();
  let existingRow = -1;
  for (let i = 1; i < rows.length; i++) if (String(rows[i][2]) === token) { existingRow = i + 1; break; }
  if (existingRow > 0) sheet.getRange(existingRow, 1, 1, 3).setValues([[name, answer, token]]);
  else sheet.appendRow([name, answer, token]);
  return ContentService.createTextOutput("ok");
}
function doGet() { return ContentService.createTextOutput("RSVP TEST OK"); }
