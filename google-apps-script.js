// 1. Go to https://sheets.google.com and create a new Sheet.
// 2. Go to Extensions > Apps Script.
// 3. Delete any code in the editor and paste this entire script.
// 4. Click "Deploy" > "New deployment".
// 5. Select type: "Web app".
// 6. Description: "Meow Meow Waitlist".
// 7. Execute as: "Me" (your email).
// 8. Who has access: "Anyone" (IMPORTANT!).
// 9. Click "Deploy".
// 10. Copy the "Web App URL" and paste it in the chat.

const SHEET_NAME = "Meow Meow";

function doGet(e) {
    return ContentService.createTextOutput("Meow! The script is working and accessible! 🐱");
}

function doPost(e) {
    const lock = LockService.getScriptLock();
    lock.tryLock(10000);

    try {
        const doc = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = doc.getSheetByName(SHEET_NAME);

        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        const nextRow = sheet.getLastRow() + 1;

        const newRow = headers.map(function (header) {
            return header === 'timestamp' ? new Date() : e.parameter[header];
        });

        sheet.getRange(nextRow, 1, 1, newRow.length).setValues([newRow]);

        return ContentService
            .createTextOutput(JSON.stringify({ 'result': 'success', 'row': nextRow }))
            .setMimeType(ContentService.MimeType.JSON);
    }

    catch (e) {
        return ContentService
            .createTextOutput(JSON.stringify({ 'result': 'error', 'error': e }))
            .setMimeType(ContentService.MimeType.JSON);
    }

    finally {
        lock.releaseLock();
    }
}

function setup() {
    const doc = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = doc.getSheetByName(SHEET_NAME);
    sheet.getRange(1, 1).setValue("timestamp");
    sheet.getRange(1, 2).setValue("email");
}
