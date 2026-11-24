# Arkan Group — Tour Expenses Sheet

A lightweight, static HTML/CSS/JS Tour Expenses Sheet replicating the exact **Arkan Group — Tour Expenses (Out of Khulna)** Excel format, with:

- Auto-calculating daily **Transport Total**, **Meal Total**, **Total Day Expense**, and bottom **Grand Total**
- **Monthly Sheet Generator** (auto working days excluding Fridays, clone previous month routes, or blank manual mode)
- Dynamic multi-route rows per day with searchable **From / To** dropdowns
- **Drag-and-drop** (`⋮⋮`) and **`▲ Up` / `▼ Down`** reordering that keeps dates fixed while moving routes & expenses
- **Copy / Paste** day routes & multi-date paste
- **Signature Upload & Draw Pad** for `Prepared by`
- **A4 Landscape Print / PDF** and native offline **Excel (`.xlsx`)** export with live formulas
- **Free Google Sheets Cloud Auto-Save** (via Google Apps Script with Secret PIN protection)

## Security Setup (Recommended for Public GitHub Pages)
1. Open your Google Apps Script (`Extensions → Apps Script` in your Google Sheet).
2. Set a secret PIN on line 18:
   ```javascript
   const SYNC_KEY = "YOUR_SECRET_PIN";
   ```
3. Click **Deploy → Manage deployments → Edit (pencil) → Version: New version → Deploy**.
4. Open the live GitHub Pages site, click **Google Sheet**, paste your Web App URL (`.../exec`) and Secret PIN, and click **Test & Connect**.
5. Your URL and PIN are stored only in your browser's local storage (use **Clear This Device** when logging out of a shared computer).
