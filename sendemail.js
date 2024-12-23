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

    // Step 3: Copy and transpose data
    Logger.log("Copying and transposing data to report sheet...");
    const dataRange = sheet.getDataRange();
    const data = dataRange.getValues();
    const transposedData = transposeArray(data);
    reportSheet.getRange(1, 1, transposedData.length, transposedData[0].length).setValues(transposedData);

    // Step 4: Auto-resize columns based on content
    Logger.log("Auto-resizing columns...");
    const numColumns = transposedData[0].length;
    for (let i = 1; i <= numColumns; i++) {
      reportSheet.autoResizeColumn(i);
    }

    // Step 5: Add conditional formatting (e.g., highlight overdue tasks)
    Logger.log("Adding conditional formatting...");
    const range = reportSheet.getDataRange();
    const rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$B2="Overdue"') // Adjust column and condition as needed
      .setBackground("#FFCCCC")
      .setRanges([range])
      .build();
    reportSheet.setConditionalFormatRules([rule]);

    // Step 6: Export the report sheet as a PDF
    Logger.log("Exporting sheet to PDF...");
    const spreadsheetId = ss.getId();
    const pdfBlob = exportSheetAsPDF(spreadsheetId, reportSheet.getSheetId(), "Weekly_Maintenance_Report.pdf");

    // Step 7: Send the PDF via email
    Logger.log("Sending email with the report...");
    GmailApp.sendEmail(emailRecipient, emailSubject, emailBody, {
      attachments: [pdfBlob]
    });

    // Step 8: Log completion
    Logger.log("Report sent successfully to " + emailRecipient);
  } catch (error) {
    Logger.log(`Error: ${error.message}`);
    throw error;
  }
}

/**
 * Transpose a 2D array (rows become columns and vice versa).
 * @param {Array<Array<any>>} array - The array to transpose.
 * @returns {Array<Array<any>>} - The transposed array.
 */
function transposeArray(array) {
  return array[0].map((_, colIndex) => array.map(row => row[colIndex]));
}

/**
 * Export a specific sheet as a PDF.
 * @param {string} spreadsheetId - The ID of the spreadsheet.
 * @param {number} sheetId - The ID of the sheet to export.
 * @param {string} pdfName - The name of the PDF file.
 * @returns {Blob} - The PDF file blob.
 */
function exportSheetAsPDF(spreadsheetId, sheetId, pdfName) {
  const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?format=pdf&gid=${sheetId}`;
  const params = {
    headers: { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` }
  };
  const response = UrlFetchApp.fetch(url, params);
  const pdfBlob = response.getBlob().setName(pdfName);
  return pdfBlob;
}
