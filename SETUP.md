# Vajra Farms Staff Portal — Setup Guide

## Step 1: Set up Google Apps Script

1. Open your Google Sheet: https://docs.google.com/spreadsheets/d/15muhKJlmGZXL7cN5zhHvosonlU6Dx_nVbNEHD3RkCMc/edit
2. Go to **Extensions > Apps Script**
3. Delete any existing code in the editor
4. Copy the entire contents of `google-apps-script.js` and paste it in
5. Click **Save** (Ctrl+S)
6. Click **Deploy > New deployment**
7. Click the gear icon next to "Select type" and choose **Web app**
8. Set **Description**: "Vajra Farms API"
9. Set **Execute as**: "Me"
10. Set **Who has access**: "Anyone"
11. Click **Deploy**
12. Authorize when prompted (click through "Advanced" > "Go to Vajra Farms" if you see a warning)
13. **Copy the Web App URL** — it looks like: `https://script.google.com/macros/s/XXXX/exec`

## Step 2: Connect the Portal

1. Open `portal.html`
2. Find this line near the top of the `<script>` section:
   ```
   const SCRIPT_URL = "YOUR_GOOGLE_APPS_SCRIPT_URL_HERE";
   ```
3. Replace `YOUR_GOOGLE_APPS_SCRIPT_URL_HERE` with the URL you copied

## Step 3: Default Logins

The system auto-creates these accounts on first use:

| Username | Password   | Role  |
|----------|-----------|-------|
| admin    | vajra2024 | Admin |
| staff    | staff123  | Staff |

**Change these passwords** after first login by editing the Users sheet directly, or remove/add users from the Admin panel.

## Features

- **Staff** can submit new milk orders
- **Admin** can submit orders + view all orders + manage user accounts
- All data is stored in your Google Sheet (Orders tab + Users tab)
- The sheets are auto-created on first use
