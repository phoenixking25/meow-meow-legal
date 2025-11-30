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
            // Send welcome email
            sendWelcomeEmail(e.parameter.email);
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

function sendWelcomeEmail(email) {
    try {
        const subject = "🐱 Your Meow Meow Early Access is Here!";

        const htmlBody = `
<!DOCTYPE html>
<html>
<head>
    <style>
        body { font-family: 'Arial', sans-serif; line-height: 1.6; color: #333; max-width: 600px; margin: 0 auto; }
        .container { padding: 20px; background-color: #f9f9f9; }
        .header { text-align: center; padding: 30px 20px; background: linear-gradient(135deg, #fff0f5 0%, #e6e6fa 100%); border-radius: 10px; }
        .content { background: white; padding: 30px; border-radius: 10px; margin-top: 20px; }
        .button { display: inline-block; padding: 15px 30px; background: #6366f1; color: white; text-decoration: none; border-radius: 10px; font-weight: bold; margin: 20px 0; }
        .footer { text-align: center; padding: 20px; color: #888; font-size: 14px; }
        h1 { margin: 0; font-size: 28px; }
        h2 { color: #6366f1; font-size: 20px; }
        ul { padding-left: 20px; }
        li { margin: 10px 0; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🐱 Welcome to Meow Meow!</h1>
            <p style="font-size: 18px; margin: 10px 0 0 0;">You're in! 🎉</p>
        </div>
        
        <div class="content">
            <p>Hi there, Cat Lover! 🐾</p>
            
            <p><strong>Thank you for joining our waitlist</strong> — we're so excited to have you as one of our first Meow Meow users.</p>
            
            <p>Your home screen is about to get a whole lot cuter. 😻</p>
            
            <p style="background: #fff3cd; padding: 15px; border-radius: 8px; border-left: 4px solid #ffc107;">
                <strong>🤖 Android Only:</strong> Meow Meow is exclusively available for Android devices.
            </p>
            
            <h2>📲 What's Next?</h2>
            
            <p>You're now on our early access waitlist! We'll send you a <strong>separate email with your Play Store closed testing group invite link</strong> very soon.</p>
            
            <p>Please <strong>keep an eye on your inbox</strong> (and check your spam folder just in case) for the invitation email. Once you receive it, you'll be able to download and install Meow Meow before it's publicly available! 🎉</p>
            
            <h2>🎨 Getting Started (Once You Get the Invite!)</h2>
            
            <ol>
                <li><strong>Wait for our email</strong> with the Play Store testing invite link</li>
                <li><strong>Accept the invite</strong> and join the closed testing group</li>
                <li><strong>Download</strong> the app from Google Play Store</li>
                <li><strong>Add the widget</strong> to your home screen (long press → Widgets → Meow Meow)</li>
                <li><strong>Choose your favorite cat</strong> or customize with time/date overlay</li>
                <li><strong>Smile</strong> every time you check your phone!</li>
            </ol>
            
            <p>That's it — you're done! 🎉</p>
            
            <h2>🐱 What Makes Meow Meow Special?</h2>
            
            <ul>
                <li>✨ <strong>Real adorable cats</strong> — carefully curated for maximum cuteness</li>
                <li>⏰ <strong>Optional time display</strong> — stay on schedule in style</li>
                <li>📅 <strong>Optional date display</strong> — never miss an important day</li>
                <li>💜 <strong>Simple & beautiful</strong> — just one tap to set up</li>
                <li>🤖 <strong>Android exclusive</strong> — designed for your Android device</li>
            </ul>
            
            <h2>💌 We'd Love Your Feedback</h2>
            
            <p>As an early access member, your opinion matters! Reply to this email with:</p>
            <ul>
                <li>What you love about Meow Meow</li>
                <li>Features you'd like to see next</li>
                <li>Any bugs or issues (we'll fix them fast!)</li>
            </ul>
            
            <p>Thanks for being part of our journey. We hope Meow Meow brings a smile to your face every day. 🐱💕</p>
            
            <p><strong>Meow you later!</strong> 😸<br>
            The Meow Meow Team</p>
            
            <p style="font-size: 14px; color: #888; margin-top: 30px;">
                P.S. Follow us on Instagram <a href="https://www.instagram.com/kawai_neko_neko_">@kawai_neko_neko_</a> for daily cat cuteness and app updates!
            </p>
        </div>
        
        <div class="footer">
            <p>Having trouble? Reply to this email and we'll help you out!</p>
            <p style="margin-top: 20px; font-size: 12px;">
                Note: Meow Meow is currently available for Android devices only.
            </p>
        </div>
    </div>
</body>
</html>
        `;

        const plainBody = `
Hi there, Cat Lover! 🐾

You're in! Thank you for joining our waitlist — we're so excited to have you as one of our first Meow Meow users.

Your home screen is about to get a whole lot cuter. 😻

🤖 ANDROID ONLY: Meow Meow is exclusively available for Android devices.

📲 WHAT'S NEXT?

You're now on our early access waitlist! We'll send you a SEPARATE EMAIL with your Play Store closed testing group invite link very soon.

Please keep an eye on your inbox (and check your spam folder just in case) for the invitation email. Once you receive it, you'll be able to download and install Meow Meow before it's publicly available! 🎉

GETTING STARTED (ONCE YOU GET THE INVITE!)
1. Wait for our email with the Play Store testing invite link
2. Accept the invite and join the closed testing group
2. Download the app from Google Play Store
3. Add the widget to your home screen (long press → Widgets → Meow Meow)
4. Choose your favorite cat or customize with time/date overlay
5. Smile every time you check your phone!

WHAT MAKES MEOW MEOW SPECIAL?
✨ Real adorable cats — carefully curated for maximum cuteness
⏰ Optional time display — stay on schedule in style
📅 Optional date display — never miss an important day
💜 Simple & beautiful — just one tap to set up
🤖 Android exclusive — designed for your Android device

WE'D LOVE YOUR FEEDBACK
Reply to this email with what you love, what features you'd like to see, or any bugs!

Thanks for being part of our journey. We hope Meow Meow brings a smile to your face every day. 🐱💕

Meow you later! 😸
The Meow Meow Team

P.S. Follow us on Instagram @kawai_neko_neko_ for daily cat cuteness and app updates!

---
Note: Meow Meow is currently available for Android devices only.
        `;

        MailApp.sendEmail({
            to: email,
            subject: subject,
            htmlBody: htmlBody,
            body: plainBody,
            name: "Meow Meow Team"
        });

        console.log("Welcome email sent to: " + email);
    } catch (e) {
        console.error("Error sending welcome email: " + e);
    }
}

// Test function to authorize email sending
// Run this manually in Apps Script editor to grant permissions
function testEmail() {
    // Replace with the email you want to test with
    sendWelcomeEmail("test@example.com");
}
