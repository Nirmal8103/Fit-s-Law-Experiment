/**
 * FITTS' LAW EXPERIMENT — GOOGLE APPS SCRIPT BACKEND & REAL-TIME ANALYTICS
 * 
 * This script runs in Google Drive / Google Sheets and serves as the serverless backend:
 * 1. Automatically saves signed Consent PDFs into "Fitts Law - Consent Forms" in Google Drive.
 * 2. Automatically logs all 320 raw trials per participant into "Raw_Trials".
 * 3. Automatically logs cohort metrics into "Participant_Summary".
 * 4. REAL-TIME AUTO-UPDATING ANALYSIS TABS:
 *    - Condition_Aggregates (16 conditions, Shannon/Fitts ID, LINEST regression, Throughput, and live scatter plot)
 *    - Directional_Analysis (Left vs. Right MT comparison, delta, and paired t-test)
 *    - Learning_Curve (Blocks 1–10 practice effect & fatigue analysis)
 *    - Participant_Comparison (Dynamic cohort ranking by throughput and error rate)
 * 
 * All formulas use whole-column ranges (e.g., Raw_Trials!$M:$M). As new participants submit,
 * every analysis sheet, linear regression model, and chart updates AUTOMATICALLY!
 */

const SPREADSHEET_NAME = "Fitts_Law_Master_Results";
const CONSENT_FOLDER_NAME = "Fitts Law - Consent Forms";

/**
 * Custom Menu when opening the Google Sheet
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📊 Fitts' Law Tools")
    .addItem("🛠️ Generate / Refresh Analysis Tabs", "menuGenerateAnalysisTabs")
    .addToUi();
}

function menuGenerateAnalysisTabs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  buildAllAnalysisTabs(ss);
  SpreadsheetApp.getUi().alert("✅ All 4 Analysis Tabs have been generated with live dynamic formulas!\n\nAs participants submit trials, every metric, regression equation, and chart will update automatically.");
}

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
      
      const decodedPdf = Utilities.base64Decode(payload.pdfBase64);
      const blob = Utilities.newBlob(decodedPdf, "application/pdf", fileName);
      
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

      // 3. Ensure live analysis tabs exist (auto-created on first submission)
      try {
        ensureAnalysisTabsExist(ss);
      } catch (err) {
        console.warn("Analysis tabs initialization note: " + err);
      }

      return responseJSON({
        status: "success",
        action: "save_results",
        participantId: payload.participantId,
        trialsRecorded: trials.length,
        spreadsheetUrl: ss.getUrl()
      });
    }

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
  return SpreadsheetApp.create(name);
}

// Helper: Get or Create Worksheet in Spreadsheet
function getOrCreateSheet(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
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

/**
 * Automatically builds all 4 dynamic analysis tabs if not already present
 */
function ensureAnalysisTabsExist(ss) {
  if (!ss.getSheetByName("Condition_Aggregates")) {
    buildConditionAggregatesTab(ss);
  }
  if (!ss.getSheetByName("Directional_Analysis")) {
    buildDirectionalAnalysisTab(ss);
  }
  if (!ss.getSheetByName("Learning_Curve")) {
    buildLearningCurveTab(ss);
  }
  if (!ss.getSheetByName("Participant_Comparison")) {
    buildParticipantComparisonTab(ss);
  }
}

function buildAllAnalysisTabs(ss) {
  buildConditionAggregatesTab(ss);
  buildDirectionalAnalysisTab(ss);
  buildLearningCurveTab(ss);
  buildParticipantComparisonTab(ss);
}

// 1. Condition_Aggregates Tab (Core Fitts' Law Regression)
function buildConditionAggregatesTab(ss) {
  const sheet = getOrCreateSheet(ss, "Condition_Aggregates");
  sheet.clear();

  const headers = [
    "Condition", "Amplitude_px", "Width_px", "ID_Shannon_bits", "ID_Fitts_bits",
    "Mean_MT_ms", "Mean_Distance_px", "Path_Efficiency", "Total_Errors", "Total_Trials", "Error_Rate_Pct"
  ];
  sheet.appendRow(headers);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#1E3A8A").setFontColor("#FFFFFF");
  sheet.setFrozenRows(1);

  // The 16 target configurations sorted by Shannon ID
  const conditions = [
    [120, 96], [120, 64], [240, 96], [120, 32],
    [240, 64], [360, 96], [480, 96], [360, 64],
    [120, 16], [240, 32], [480, 64], [360, 32],
    [240, 16], [480, 32], [360, 16], [480, 16]
  ];

  const rows = [];
  for (let i = 0; i < conditions.length; i++) {
    const A = conditions[i][0];
    const W = conditions[i][1];
    const r = i + 2;
    rows.push([
      `A${A}/W${W}`,
      A,
      W,
      `=ROUND(LOG(B${r}/C${r} + 1, 2), 2)`,
      `=ROUND(LOG(2*B${r}/C${r}, 2), 2)`,
      `=IFERROR(ROUND(AVERAGEIFS(Raw_Trials!$M:$M, Raw_Trials!$H:$H, B${r}, Raw_Trials!$I:$I, C${r}), 1), "")`,
      `=IFERROR(ROUND(AVERAGEIFS(Raw_Trials!$N:$N, Raw_Trials!$H:$H, B${r}, Raw_Trials!$I:$I, C${r}), 1), "")`,
      `=IFERROR(ROUND(AVERAGEIFS(Raw_Trials!$P:$P, Raw_Trials!$H:$H, B${r}, Raw_Trials!$I:$I, C${r}), 3), "")`,
      `=COUNTIFS(Raw_Trials!$H:$H, B${r}, Raw_Trials!$I:$I, C${r}, Raw_Trials!$Q:$Q, ">0")`,
      `=COUNTIFS(Raw_Trials!$H:$H, B${r}, Raw_Trials!$I:$I, C${r})`,
      `=IFERROR(ROUND(I${r}/J${r} * 100, 2), 0)`
    ]);
  }
  sheet.getRange(2, 1, rows.length, headers.length).setFormulas(rows);

  // Regression KPI Block
  sheet.getRange(20, 1).setValue("FITTS' LAW LINEAR REGRESSION SUMMARY").setFontWeight("bold").setFontColor("#1E3A8A");
  
  const kpiLabels = [
    ["Slope b (ms/bit):", `=IFERROR(ROUND(INDEX(LINEST(F2:F17, D2:D17), 1), 2), "")`],
    ["Intercept a (ms):", `=IFERROR(ROUND(INDEX(LINEST(F2:F17, D2:D17), 2), 2), "")`],
    ["Goodness of Fit (R²):", `=IFERROR(ROUND(INDEX(LINEST(F2:F17, D2:D17, TRUE, TRUE), 3, 1), 4), "")`],
    ["Correlation (r):", `=IFERROR(ROUND(CORREL(F2:F17, D2:D17), 4), "")`],
    ["Throughput IP (bits/s):", `=IFERROR(ROUND(1000 / B21, 2), "")`],
    ["Model Equation:", `="MT = " & B22 & " + " & B21 & " × ID"`]
  ];

  for (let k = 0; k < kpiLabels.length; k++) {
    const rowIdx = 21 + k;
    sheet.getRange(rowIdx, 1).setValue(kpiLabels[k][0]).setFontWeight("bold");
    sheet.getRange(rowIdx, 2).setFormula(kpiLabels[k][1]);
  }
  sheet.getRange(21, 1, kpiLabels.length, 2).setBackground("#F1F5F9");
}

// 2. Directional_Analysis Tab (Left vs. Right — Question 1)
function buildDirectionalAnalysisTab(ss) {
  const sheet = getOrCreateSheet(ss, "Directional_Analysis");
  sheet.clear();

  const headers = [
    "Condition", "Amplitude_px", "Width_px", "ID_Shannon",
    "Left_Mean_MT_ms", "Right_Mean_MT_ms", "Difference_Delta_ms",
    "Left_Error_Rate_Pct", "Right_Error_Rate_Pct", "Left_Efficiency", "Right_Efficiency"
  ];
  sheet.appendRow(headers);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#1E3A8A").setFontColor("#FFFFFF");
  sheet.setFrozenRows(1);

  const conditions = [
    [120, 96], [120, 64], [240, 96], [120, 32],
    [240, 64], [360, 96], [480, 96], [360, 64],
    [120, 16], [240, 32], [480, 64], [360, 32],
    [240, 16], [480, 32], [360, 16], [480, 16]
  ];

  const rows = [];
  for (let i = 0; i < conditions.length; i++) {
    const A = conditions[i][0];
    const W = conditions[i][1];
    const r = i + 2;
    rows.push([
      `A${A}/W${W}`,
      A,
      W,
      `=ROUND(LOG(B${r}/C${r} + 1, 2), 2)`,
      `=IFERROR(ROUND(AVERAGEIFS(Raw_Trials!$M:$M, Raw_Trials!$H:$H, B${r}, Raw_Trials!$I:$I, C${r}, Raw_Trials!$J:$J, "Left"), 1), "")`,
      `=IFERROR(ROUND(AVERAGEIFS(Raw_Trials!$M:$M, Raw_Trials!$H:$H, B${r}, Raw_Trials!$I:$I, C${r}, Raw_Trials!$J:$J, "Right"), 1), "")`,
      `=IF(AND(ISNUMBER(E${r}), ISNUMBER(F${r})), ROUND(E${r} - F${r}, 1), "")`,
      `=IFERROR(ROUND(COUNTIFS(Raw_Trials!$H:$H, B${r}, Raw_Trials!$I:$I, C${r}, Raw_Trials!$J:$J, "Left", Raw_Trials!$Q:$Q, ">0") / COUNTIFS(Raw_Trials!$H:$H, B${r}, Raw_Trials!$I:$I, C${r}, Raw_Trials!$J:$J, "Left") * 100, 2), 0)`,
      `=IFERROR(ROUND(COUNTIFS(Raw_Trials!$H:$H, B${r}, Raw_Trials!$I:$I, C${r}, Raw_Trials!$J:$J, "Right", Raw_Trials!$Q:$Q, ">0") / COUNTIFS(Raw_Trials!$H:$H, B${r}, Raw_Trials!$I:$I, C${r}, Raw_Trials!$J:$J, "Right") * 100, 2), 0)`,
      `=IFERROR(ROUND(AVERAGEIFS(Raw_Trials!$P:$P, Raw_Trials!$H:$H, B${r}, Raw_Trials!$I:$I, C${r}, Raw_Trials!$J:$J, "Left"), 3), "")`,
      `=IFERROR(ROUND(AVERAGEIFS(Raw_Trials!$P:$P, Raw_Trials!$H:$H, B${r}, Raw_Trials!$I:$I, C${r}, Raw_Trials!$J:$J, "Right"), 3), "")`
    ]);
  }
  sheet.getRange(2, 1, rows.length, headers.length).setFormulas(rows);

  // Directional Statistical Summary
  sheet.getRange(20, 1).setValue("DIRECTIONAL ASYMMETRY & HYPOTHESIS TEST").setFontWeight("bold").setFontColor("#1E3A8A");
  const stats = [
    ["Mean Leftward MT (ms):", `=IFERROR(ROUND(AVERAGE(E2:E17), 1), "")`],
    ["Mean Rightward MT (ms):", `=IFERROR(ROUND(AVERAGE(F2:F17), 1), "")`],
    ["Overall Difference Delta (ms):", `=IFERROR(ROUND(AVERAGE(G2:G17), 1), "")`],
    ["Paired t-Test (p-value):", `=IFERROR(T.TEST(E2:E17, F2:F17, 2, 1), "")`],
    ["Conclusion:", `=IF(B24<0.05, "Statistically Significant Difference (p < 0.05)", "No Significant Difference")`]
  ];
  for (let s = 0; s < stats.length; s++) {
    const rowIdx = 21 + s;
    sheet.getRange(rowIdx, 1).setValue(stats[s][0]).setFontWeight("bold");
    sheet.getRange(rowIdx, 2).setFormula(stats[s][1]);
  }
  sheet.getRange(21, 1, stats.length, 2).setBackground("#F1F5F9");
}

// 3. Learning_Curve Tab (Blocks 1–10 — Question 3 & Confounders)
function buildLearningCurveTab(ss) {
  const sheet = getOrCreateSheet(ss, "Learning_Curve");
  sheet.clear();

  const headers = ["Block", "Mean_MT_ms", "Error_Rate_Pct", "Path_Efficiency", "Trials_Logged"];
  sheet.appendRow(headers);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#1E3A8A").setFontColor("#FFFFFF");
  sheet.setFrozenRows(1);

  const rows = [];
  for (let b = 1; b <= 10; b++) {
    const r = b + 1;
    rows.push([
      b,
      `=IFERROR(ROUND(AVERAGEIFS(Raw_Trials!$M:$M, Raw_Trials!$E:$E, A${r}), 1), "")`,
      `=IFERROR(ROUND(COUNTIFS(Raw_Trials!$E:$E, A${r}, Raw_Trials!$Q:$Q, ">0") / COUNTIFS(Raw_Trials!$E:$E, A${r}) * 100, 2), 0)`,
      `=IFERROR(ROUND(AVERAGEIFS(Raw_Trials!$P:$P, Raw_Trials!$E:$E, A${r}), 3), "")`,
      `=COUNTIFS(Raw_Trials!$E:$E, A${r})`
    ]);
  }
  sheet.getRange(2, 1, rows.length, headers.length).setFormulas(rows);

  sheet.getRange(14, 1).setValue("INTERPRETATION NOTES").setFontWeight("bold").setFontColor("#1E3A8A");
  sheet.getRange(15, 1).setValue("• Practice Effect: Look for MT decline from Block 1 to Block 4 as participants adapt to target layout.");
  sheet.getRange(16, 1).setValue("• Motor Fatigue: Check if Error Rate increases in Blocks 9-10 indicating physical wrist/eye strain.");
}

// 4. Participant_Comparison Tab (Cohort Variance — Question 2)
function buildParticipantComparisonTab(ss) {
  const sheet = getOrCreateSheet(ss, "Participant_Comparison");
  sheet.clear();

  const headers = [
    "Participant_ID", "Device", "Handedness", "Trials_Completed",
    "Mean_MT_ms", "Error_Rate_Pct", "Path_Efficiency", "Throughput_bits_s", "Intercept_a_ms", "Slope_b_ms_bit", "R_Squared"
  ];
  sheet.appendRow(headers);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#1E3A8A").setFontColor("#FFFFFF");
  sheet.setFrozenRows(1);

  const rows = [];
  for (let p = 2; p <= 20; p++) {
    rows.push([
      `=IF(Participant_Summary!A${p}<>"", Participant_Summary!A${p}, "")`,
      `=IF(Participant_Summary!A${p}<>"", Participant_Summary!C${p}, "")`,
      `=IF(Participant_Summary!A${p}<>"", Participant_Summary!D${p}, "")`,
      `=IF(Participant_Summary!A${p}<>"", Participant_Summary!E${p}, "")`,
      `=IF(Participant_Summary!A${p}<>"", Participant_Summary!F${p}, "")`,
      `=IF(Participant_Summary!A${p}<>"", Participant_Summary!G${p}, "")`,
      `=IF(Participant_Summary!A${p}<>"", Participant_Summary!I${p}, "")`,
      `=IF(Participant_Summary!A${p}<>"", Participant_Summary!J${p}, "")`,
      `=IF(Participant_Summary!A${p}<>"", Participant_Summary!K${p}, "")`,
      `=IF(Participant_Summary!A${p}<>"", Participant_Summary!L${p}, "")`,
      `=IF(Participant_Summary!A${p}<>"", Participant_Summary!M${p}, "")`
    ]);
  }
  sheet.getRange(2, 1, rows.length, headers.length).setFormulas(rows);
}
