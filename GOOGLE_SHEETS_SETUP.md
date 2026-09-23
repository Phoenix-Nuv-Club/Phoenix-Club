# Google Sheets Integration Guide for Phoenix Club Forms

This guide walks you through connecting your **Contact Us Form** and **Event Registration Form** directly to a **Google Sheet** so all inquiries and submissions are automatically appended as rows in real time.

---

## ⚡ Quick 3-Step Setup

### Step 1: Create a Google Sheet & Open Apps Script
1. Open [Google Sheets](https://sheets.new) in your browser.
2. Name your spreadsheet (e.g., **"Phoenix Club — Form Submissions & Inquiries"**).
3. In the top menu, click **Extensions** &rarr; **Apps Script**.

---

### Step 2: Paste the Automation Script
1. Delete any code currently in the `Code.gs` editor.
2. Copy and paste the script below into `Code.gs`:

```javascript
/**
 * Phoenix Club — Google Sheets Automation Script
 * Receives Contact Form inquiries & Registration submissions and appends rows automatically.
 */

// If you created this script directly inside your Google Sheet (Extensions > Apps Script),
// leave SPREADSHEET_ID empty. If created as a standalone script, paste your Sheet ID here:
var SPREADSHEET_ID = "";

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({
    status: 'online',
    message: 'Phoenix Club Google Sheets Webhook is active and healthy.'
  })).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // 10s timeout to prevent concurrency collisions

    var data = {};
    if (e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    } else if (e.parameter) {
      data = e.parameter;
    }

    var doc = SpreadsheetApp.getActiveSpreadsheet();
    if (!doc && SPREADSHEET_ID) {
      doc = SpreadsheetApp.openById(SPREADSHEET_ID);
    }

    if (!doc) {
      throw new Error("Spreadsheet not found. Please bind this script to your Google Sheet via Extensions > Apps Script, or set SPREADSHEET_ID.");
    }

    var formType = data.form_type || 'contact';

    if (formType === 'contact') {
      // 1. Contact Form Submissions Sheet
      var sheet = doc.getSheetByName('Contact Inquiries');
      if (!sheet) {
        sheet = doc.insertSheet('Contact Inquiries');
      }
      
      // Add formatted headers on first run
      if (sheet.getLastRow() === 0) {
        sheet.appendRow(['Timestamp', 'Full Name', 'Email Address', 'Inquiry Purpose', 'Message']);
        var headerRange = sheet.getRange(1, 1, 1, 5);
        headerRange.setFontWeight('bold');
        headerRange.setBackground('#f97316');
        headerRange.setFontColor('#ffffff');
        sheet.setFrozenRows(1);
      }

      sheet.appendRow([
        data.timestamp || new Date().toLocaleString(),
        data.full_name || '',
        data.email || '',
        data.purpose || 'General Student Inquiry',
        data.message || ''
      ]);

    } else if (formType === 'registration') {
      // 2. Student Registrations Sheet
      var sheet = doc.getSheetByName('Registrations');
      if (!sheet) {
        sheet = doc.insertSheet('Registrations');
      }

      // Add formatted headers on first run
      if (sheet.getLastRow() === 0) {
        sheet.appendRow([
          'Timestamp',
          'Pass Code',
          'Full Name',
          'Enrollment ID',
          'School / Department',
          'Academic Year',
          'Email Address',
          'Phone Number',
          'Event Name',
          'Event Date / Schedule',
          'Additional Notes'
        ]);
        var headerRange = sheet.getRange(1, 1, 1, 11);
        headerRange.setFontWeight('bold');
        headerRange.setBackground('#7c3aed');
        headerRange.setFontColor('#ffffff');
        sheet.setFrozenRows(1);
      }

      sheet.appendRow([
        data.timestamp || new Date().toLocaleString(),
        data.reg_code || '',
        data.full_name || '',
        data.enrollment_id || '',
        data.department || '',
        data.academic_year || '',
        data.email || '',
        data.phone || '',
        data.event_name || '',
        data.event_date || '',
        data.additional_info || ''
      ]);
    }

    return ContentService.createTextOutput(JSON.stringify({
      result: 'success',
      message: 'Form submission appended to Google Sheets successfully'
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      result: 'error',
      error: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    try {
      lock.releaseLock();
    } catch (e) {}
  }
}
```

---

### Step 3: Deploy as a Web App (Critical Settings)

> [!IMPORTANT]
> **To prevent the `"You do not have permission to access the requested document"` error:**
> You must configure **Execute as: Me**! If it is set to "User accessing the web app", anonymous form submissions will be blocked by Google.

1. In the Apps Script editor, click the blue **Deploy** button (top-right) &rarr; Select **Manage deployments**.
2. Click the pencil (`✏️`) **Edit** icon next to your active deployment.
3. Configure these exact settings:
   - **Execute as**: **`Me (your Google email)`** &larr; *(CRITICAL)*
   - **Who has access**: **`Anyone`** &larr; *(CRITICAL: allows web forms to submit entries without requiring students to log into Google)*
   - **Version**: Select **`New version`** so your code updates take effect!
4. Click **Deploy**.
5. If Google prompts *"Authorization required"*:
   - Click **Review permissions**.
   - Select your Google account.
   - Click **Advanced** &rarr; Click **Go to Phoenix Webhook (unsafe)** &rarr; Click **Allow**.
6. Copy the generated **Web App URL** (ends in `/exec`).

---

## 🔗 Connect Your URL to the Website

### Option A: If running locally with Python Flask Server
Open your `.env` file and paste the Web App URL:
```env
GOOGLE_SHEET_WEBHOOK_URL=https://script.google.com/macros/s/AKfycb.../exec
```

### Option B: If deploying statically on Vercel / GitHub Pages
Open [contact.html](file:///c:/Users/Nair%20Prashobh%20Manoj/OneDrive/Desktop/Web/Ravichandran/contact.html) and [registration.html](file:///c:/Users/Nair%20Prashobh%20Manoj/OneDrive/Desktop/Web/Ravichandran/registration.html) and set the URL near the bottom:
```javascript
window.GOOGLE_SHEETS_WEBHOOK_URL = "https://script.google.com/macros/s/AKfycb.../exec";
```

---

## ✅ What Happens When a Student Submits?
1. **Contact Inquiries**: A new row is automatically added to the **"Contact Inquiries"** tab with Timestamp, Full Name, Email, Inquiry Purpose, and Message.
2. **Registrations**: A new row is added to the **"Registrations"** tab with the student's Registration Pass Code, Name, Roll No., Department, Email, WhatsApp Phone, and Event info.
3. **Local Database & Activity Logs**: All submissions are also recorded in your local database `phoenix_club.db` for the Admin Dashboard and Excel/CSV exports!
