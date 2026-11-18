# Google Sheets & Google Drive Webhook Setup Guide

This guide walks you through connecting your **Vercel-deployed Fitts' Law Experiment** to your personal Google account. 

Once configured:
1. **Consent Form PDFs** automatically save directly into a `Fitts Law - Consent Forms` folder in your Google Drive.
2. **Participant Progress** is saved automatically after every trial in the participant's browser (resumes immediately across breaks and page reloads).
3. **Completed Results (All 320 Trials)** automatically stream directly into your master Google Sheet (`Fitts_Law_Master_Results`) so you can download the consolidated Excel file anytime!

---

## Step-by-Step Instructions (Takes ~2 Minutes)

### Step 1: Open Google Apps Script
1. Go to [https://script.google.com](https://script.google.com) in your browser.
2. Sign in with your Google account.
3. Click the **"New project"** button (top left).
4. Name the project (top left) e.g., `Fitts Law Backend`.

---

### Step 2: Paste the Serverless Backend Code
1. In the Apps Script editor, delete any default code inside `Code.gs`.
2. Open the file [google_apps_script.js](file:///c:/Users/nirma/Downloads/582/Fit's%20Law/google_apps_script.js) from your workspace.
3. Copy the entire contents of `google_apps_script.js` and paste them into `Code.gs`.
4. Click the **Save** icon (disk icon) or press `Ctrl + S`.

---

### Step 3: Deploy as a Public Web App
1. At the top right, click the blue **"Deploy"** button and select **"New deployment"**.
2. Click the gear icon next to "Select type" and choose **"Web app"**.
3. Fill in the deployment settings:
   - **Description:** `Fitts Law API v1`
   - **Execute as:** `Me (your_email@gmail.com)` *(Important!)*
   - **Who has access:** `Anyone` *(Important: Allows participants' browsers to upload results)*
4. Click **"Deploy"**.

---

### Step 4: Authorize Google Permissions
1. Google will show an **"Authorize access"** prompt. Click **"Authorize access"** and select your Google account.
2. If you see an *"Advanced / Google hasn't verified this app"* warning (normal for your own personal scripts):
   - Click **"Advanced"** (bottom left).
   - Click **"Go to Fitts Law Backend (unsafe)"**.
   - Click **"Allow"**.
3. You will now see a **Web app URL** that looks like:
   ```
   https://script.google.com/macros/s/AKfycbx.../exec
   ```
4. Copy this **Web app URL**.

---

### Step 5: Paste Your URL into `index.html`
1. Open `index.html` (around line 725) and locate:
   ```javascript
   const GOOGLE_SCRIPT_WEBHOOK_URL = "";
   ```
2. Paste your copied URL inside the quotes:
   ```javascript
   const GOOGLE_SCRIPT_WEBHOOK_URL = "https://script.google.com/macros/s/AKfycbx.../exec";
   ```
3. Commit and push to your GitHub repository:
   ```bash
   git add .
   git commit -m "Connect Google Sheets & Drive backend"
   git push origin main
   ```
4. Vercel will automatically redeploy within seconds!

---

## How to View and Download Your Results

### 1. View Consent Form PDFs
- Open [Google Drive](https://drive.google.com).
- You will see a new folder named **`Fitts Law - Consent Forms`**.
- Inside, you will find `Consent_Form_P01.pdf`, `Consent_Form_P02.pdf`, etc., with timestamps and electronic signature records.

### 2. Download the Consolidated Master Excel File
- In [Google Drive](https://drive.google.com), open the spreadsheet named **`Fitts_Law_Master_Results`**.
- It contains two tabs:
  - **`Raw_Trials`**: Every single trial (320 rows per participant) with amplitudes, widths, directions, Shannon IDs, movement times, and error logs.
  - **`Participant_Summary`**: Cohort-level summary metrics for each participant (mean MT, error rate, throughput, regression parameters).
- To get the complete master Excel file, simply click:
  **`File` > `Download` > `Microsoft Excel (.xlsx)`**.
