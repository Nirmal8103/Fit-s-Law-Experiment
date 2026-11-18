/**
 * FITTS' LAW EXPERIMENT — GOOGLE APPS SCRIPT BACKEND
 * 
 * This script runs in Google Drive / Google Sheets and serves as the serverless
 * backend for your Vercel deployment:
 * 1. Automatically saves signed Consent PDFs into a "Fitts Law - Consent Forms" Google Drive folder.
 * 2. Automatically logs all 320 raw trials per participant into the "Raw_Trials" tab.
 * 3. Automatically logs cohort metrics into the "Participant_Summary" tab.
 * 
 * Instructions to deploy:
 * 1. Go to https://script.google.com and click "New project".
 * 2. Paste this entire script into Code.gs (replacing any existing code).
 * 3. Click "Deploy" > "New deployment".
 * 4. Select type: "Web app".
 * 5. Set:
 *    - Description: "Fitts Law Backend"
 *    - Execute as: "Me" (your Google account)
 *    - Who has access: "Anyone" (allows participants' browsers to submit)
 * 6. Click "Deploy", authorize permissions, and copy the Web App URL.
 * 7. Paste that Web App URL into GOOGLE_SCRIPT_WEBHOOK_URL in index.html!
 */

const SPREADSHEET_NAME = "Fitts_Law_Master_Results";
const CONSENT_FOLDER_NAME = "Fitts Law - Consent Forms";

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return responseJSON({ status: "error", message: "No post data received" });
    }

    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;

    // ACTION 1: Save Signed Consent Form PDF to Google Drive
    if (action === "save_consent") {
      const folder = getOrCreateFolder(CONSENT_FOLDER_NAME);
      const fileName = payload.fileName || ("Consent_Form_" + payload.participantId + ".pdf");
      
      // Decode base64 PDF
      const decodedPdf = Utilities.base64Decode(payload.pdfBase64);
      const blob = Utilities.newBlob(decodedPdf, "application/pdf", fileName);
      
      // Check if file with same name already exists in folder; if so, trash it first to avoid duplicates
      const existingFiles = folder.getFilesByName(fileName);
      while (existingFiles.hasNext()) {
        existingFiles.next().setTrashed(true);
      }
      
      const file = folder.createFile(blob);
      return responseJSON({
        status: "success",
        action: "save_consent",
        participantId: payload.participantId,
        fileName: fileName,
        fileUrl: file.getUrl()
      });
    }

    // ACTION 2: Save All 320 Raw Trials & Participant Summary to Master Google Sheet
    if (action === "save_results") {
      const ss = getOrCreateSpreadsheet(SPREADSHEET_NAME);
      
      // 1. Process Raw Trials Sheet
      const rawSheet = getOrCreateSheet(ss, "Raw_Trials");
      ensureRawHeaders(rawSheet);
      
      const trials = payload.trials || [];
      if (trials.length > 0) {
        const rowsToInsert = trials.map(t => [
          t.participantId || "",
          t.age || "",
          t.device || "",
          t.handedness || "",
          t.block || "",
          t.trialInBlock || "",
          t.overallTrial || "",
          t.amplitude || "",
          t.width || "",
          t.direction || "",
          t.idShannon || "",
          t.idFitts || "",
          t.movementTime || "",
          t.actualDistance || "",
          t.optimalDistance || "",
          t.trajectoryEfficiency || "",
          t.errors || 0,
          t.success || (t.errors === 0 ? 1 : 0),
          payload.submittedAt || new Date().toISOString()
        ]);
        
        // Batch write rows for high-speed insertion
        const startRow = rawSheet.getLastRow() + 1;
        const numRows = rowsToInsert.length;
        const numCols = rowsToInsert[0].length;
        rawSheet.getRange(startRow, 1, numRows, numCols).setValues(rowsToInsert);
      }
      
      // 2. Process Participant Summary Sheet
      const summarySheet = getOrCreateSheet(ss, "Participant_Summary");
      ensureSummaryHeaders(summarySheet);
      
      const summary = payload.summary || {};
      const summaryRow = [
        payload.participantId || "",
        summary.age || "",
        summary.device || "",
        summary.handedness || "",
        trials.length,
        summary.meanMt || "",
        summary.errorRate || "",
        summary.meanDistance || "",
        summary.trajectoryEfficiency || "",
        summary.throughput || "",
        summary.regressionA || "",
        summary.regressionB || "",
        summary.rSquared || "",
        payload.submittedAt || new Date().toISOString()
      ];
      summarySheet.appendRow(summaryRow);

      return responseJSON({
        status: "success",
        action: "save_results",
        participantId: payload.participantId,
        trialsRecorded: trials.length,
        spreadsheetUrl: ss.getUrl()
      });
    }

    // Unknown action
    return responseJSON({ status: "error", message: "Unknown action: " + action });

  } catch (error) {
    return responseJSON({
      status: "error",
      message: error.toString()
    });
  }
}

// Return JSON response with CORS headers
function responseJSON(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Helper: Get or Create Google Drive Folder
function getOrCreateFolder(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(folderName);
}

// Helper: Get or Create Google Spreadsheet
function getOrCreateSpreadsheet(name) {
  const files = DriveApp.getFilesByName(name);
  if (files.hasNext()) {
    const file = files.next();
    return SpreadsheetApp.open(file);
  }
  const ss = SpreadsheetApp.create(name);
  return ss;
}

// Helper: Get or Create Worksheet in Spreadsheet
function getOrCreateSheet(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    // If default "Sheet1" is empty, remove it
    const defaultSheet = ss.getSheetByName("Sheet1");
    if (defaultSheet && defaultSheet.getLastRow() === 0 && ss.getSheets().length > 1) {
      try { ss.deleteSheet(defaultSheet); } catch (e) {}
    }
  }
  return sheet;
}

// Ensure Headers for Raw_Trials tab
function ensureRawHeaders(sheet) {
  if (sheet.getLastRow() === 0) {
    const headers = [
      "Participant_ID", "Age", "Device", "Handedness", "Block", "Trial_In_Block", "Overall_Trial",
      "Amplitude_px", "Width_px", "Direction", "ID_Shannon_bits", "ID_Fitts_bits",
      "Movement_Time_ms", "Actual_Distance_px", "Optimal_Distance_px", "Trajectory_Efficiency",
      "Errors", "Success", "Submitted_At"
    ];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#1E3A8A").setFontColor("#FFFFFF");
    sheet.setFrozenRows(1);
  }
}

// Ensure Headers for Participant_Summary tab
function ensureSummaryHeaders(sheet) {
  if (sheet.getLastRow() === 0) {
    const headers = [
      "Participant_ID", "Age", "Device", "Handedness", "Trials_Completed",
      "Mean_MT_ms", "Error_Rate_Pct", "Mean_Distance_px", "Trajectory_Efficiency",
      "Throughput_bits_s", "Intercept_a_ms", "Slope_b_ms_bit", "R_Squared", "Submitted_At"
    ];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#1E3A8A").setFontColor("#FFFFFF");
    sheet.setFrozenRows(1);
  }
}
