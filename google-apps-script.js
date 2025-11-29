// 1. Go to https://sheets.google.com and create a new Sheet.
// 2. Go to Extensions > Apps Script.
// 3. Delete any code in the editor and paste this entire script.
// 4. Click "Deploy" > "New deployment".
// 5. Select type: "Web app".
// 6. Description: "Meow Meow Waitlist v2".
// 7. Execute as: "Me" (your email).
// 8. Who has access: "Anyone" (IMPORTANT!).
// 9. Click "Deploy".
// 10. Copy the "Web App URL" and paste it in the chat.

const SHEET_NAME = "Meow Meow";
const PIXEL_ID = "25424083083870050";
const ACCESS_TOKEN = "EAAapIUvypdkBQHUrL1izBPesc5Ulfxp9rUCjiXTQa3L4EZBLahU00Os0Vb0bceiKlXW8Dx3WMlS1zotaHzSbHDjZBceY4cZC52Q8xspq3XYqsc59JtAXiJkoPupWeXoSiSR1cDEJjZA3SKwglxmJ34T0rPXh8KfahTQp2myBSqZCAbbqVQRwYFbVP8MMsPgZDZD"; // Replace with your actual access token

function doGet(e) {
    return ContentService.createTextOutput("Meow! The script is working and accessible! 🐱");
}

function doPost(e) {
    const lock = LockService.getScriptLock();
    lock.tryLock(10000);

    try {
        const doc = SpreadsheetApp.getActiveSpreadsheet();
        let sheet = doc.getSheetByName(SHEET_NAME);

        // Fallback: If sheet doesn't exist, use the first sheet or create one
        if (!sheet) {
            sheet = doc.getSheets()[0];
            // Optional: Rename it to avoid confusion later? 
            // sheet.setName(SHEET_NAME); 
        }

        // Ensure headers exist
        if (sheet.getLastRow() === 0) {
            const initialHeaders = ["timestamp", "email"];
            sheet.getRange(1, 1, 1, initialHeaders.length).setValues([initialHeaders]);
        }

        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        const nextRow = sheet.getLastRow() + 1;

        // Map incoming parameters to headers (Case Insensitive)
        const newRow = headers.map(function (header) {
            if (header.toLowerCase() === 'timestamp') return new Date();

            // Find matching key in e.parameter (case insensitive)
            const paramKey = Object.keys(e.parameter).find(key => key.toLowerCase() === header.toLowerCase());
            return paramKey ? e.parameter[paramKey] : "";
        });

        sheet.getRange(nextRow, 1, 1, newRow.length).setValues([newRow]);

        // Send event to Facebook CAPI
        if (e.parameter.email) {
            sendFacebookEvent(e.parameter.email);
        }

        return ContentService
            .createTextOutput(JSON.stringify({ 'result': 'success', 'row': nextRow }))
            .setMimeType(ContentService.MimeType.JSON);
    }

    catch (e) {
        console.error("Error in doPost:", e); // Log to Apps Script dashboard
        return ContentService
            .createTextOutput(JSON.stringify({ 'result': 'error', 'error': e.toString() }))
            .setMimeType(ContentService.MimeType.JSON);
    }

    finally {
        lock.releaseLock();
    }
}

function sendFacebookEvent(email) {
    if (!ACCESS_TOKEN || ACCESS_TOKEN === "YOUR_ACCESS_TOKEN_HERE") {
        console.log("Skipping Facebook Event: Access Token not set.");
        return;
    }

    const hashedEmail = hashEmail(email);
    const url = `https://graph.facebook.com/v18.0/${PIXEL_ID}/events?access_token=${ACCESS_TOKEN}`;

    const payload = {
        "data": [
            {
                "event_name": "Lead",
                "event_time": Math.floor(new Date().getTime() / 1000),
                "action_source": "website",
                "user_data": {
                    "em": [hashedEmail]
                },
                "custom_data": {
                    "currency": "USD",
                    "value": 0
                }
            }
        ]
    };

    const options = {
        "method": "post",
        "contentType": "application/json",
        "payload": JSON.stringify(payload)
    };

    try {
        UrlFetchApp.fetch(url, options);
    } catch (e) {
        console.error("Error sending to Facebook CAPI: " + e);
    }
}

function hashEmail(email) {
    const cleanEmail = email.trim().toLowerCase();
    const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, cleanEmail);
    let txtHash = "";
    for (let i = 0; i < digest.length; i++) {
        let hashVal = digest[i];
        if (hashVal < 0) {
            hashVal += 256;
        }
        if (hashVal.toString(16).length == 1) {
            txtHash += "0";
        }
        txtHash += hashVal.toString(16);
    }
    return txtHash;
}

function setup() {
    const doc = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = doc.getSheetByName(SHEET_NAME);
    if (!sheet) {
        sheet = doc.insertSheet(SHEET_NAME);
    }
    sheet.getRange(1, 1).setValue("timestamp");
    sheet.getRange(1, 2).setValue("email");
}
