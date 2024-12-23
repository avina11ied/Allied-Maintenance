function generateAndEmailReport() {
  const sheetName = "Form Responses 1"; // Update with your sheet name
  const reportSheetName = "Weekly Maintenance Report";
  const emailRecipient = "avin@allied.com.sg"; // Replace with the recipient's email
  const emailSubject = "Weekly Maintenance Report";
  const emailBody = "Dear Team,\n\nPlease find attached the weekly maintenance report.\n\nBest regards,\nMaintenance Team";

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Step 1: Verify the original sheet exists
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error(`Sheet "${sheetName}" does not exist.`);

    // Step 2: Create a new report sheet
    let reportSheet = ss.getSheetByName(reportSheetName);
    if (reportSheet) {
      Logger.log(`Deleting existing report sheet: ${reportSheetName}`);
      ss.deleteSheet(reportSheet);
    }
    reportSheet = ss.insertSheet(reportSheetName);

    // Step 3: Copy data
    Logger.log("Copying data to report sheet...");
    const dataRange = sheet.getDataRange();
    reportSheet.getRange(1, 1, dataRange.getNumRows(), dataRange.getNumColumns()).setValues(dataRange.getValues());

    // Step 4: Auto-resize columns based on content
    Logger.log("Auto-resizing columns...");
    const numColumns = dataRange.getNumColumns();
    for (let i = 1; i <= numColumns; i++) {
      reportSheet.autoResizeColumn(i);
    }

    // Step 5: Add conditional formatting (e.g., highlight overdue tasks)
    Logger.log("Adding conditional formatting...");
    const range = reportSheet.getDataRange();
    const rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$E2="Overdue"') // Adjust column and condition as needed
      .setBackground("#FFCCCC")
      .setRanges([range])
      .build();
    reportSheet.setConditionalFormatRules([rule]);

    // Step 6: Export the report sheet as an Excel file (.xlsx)
    Logger.log("Exporting sheet to Excel...");
    const spreadsheetId = ss.getId();
    const xlsxBlob = exportSheetAsExcel(spreadsheetId, reportSheet.getSheetId(), "Weekly_Maintenance_Report.xlsx");

    // Step 7: Send the Excel file via email
    Logger.log("Sending email with the report...");
    GmailApp.sendEmail(emailRecipient, emailSubject, emailBody, {
      attachments: [xlsxBlob]
    });

    // Step 8: Log completion
    Logger.log("Report sent successfully to " + emailRecipient);
  } catch (error) {
    Logger.log(`Error: ${error.message}`);
    throw error;
  }
}

/**
 * Export a specific sheet as an Excel file (.xlsx).
 * @param {string} spreadsheetId - The ID of the spreadsheet.
 * @param {number} sheetId - The ID of the sheet to export.
 * @param {string} xlsxName - The name of the Excel file.
 * @returns {Blob} - The Excel file blob.
 */
function exportSheetAsExcel(spreadsheetId, sheetId, xlsxName) {
  const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?format=xlsx&gid=${sheetId}`;
  const params = {
    headers: { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` }
  };
  const response = UrlFetchApp.fetch(url, params);
  const xlsxBlob = response.getBlob().setName(xlsxName);
  return xlsxBlob;
}
