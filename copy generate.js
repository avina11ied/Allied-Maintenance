function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Menu')
    .addItem('Create Document from Row', 'createDocumentFromRow')
    .addToUi();
}

function createDocumentFromRow() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = sheet.getActiveRange();
  const row = range.getRow();
  const lastColumn = 20; // Column T (20th column)

  // Ensure we don't write to the header row
  if (row === 1) {
    SpreadsheetApp.getUi().alert('Please select a data row, not the header row.');
    return;
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Get machine number and timestamp
  const { machineNumber, timestamp } = getMachineAndTimestamp(headers, rowData);

  // Create the document
  const templateDoc = DocumentApp.create(`Machine ${machineNumber} - Generated Report`);
  const docBody = templateDoc.getBody();

  // Add title (Machine Number as the first header)
  docBody.appendParagraph(`Machine Number: ${machineNumber}`).setHeading(DocumentApp.ParagraphHeading.HEADING1);

  // Add timestamp as second header
  docBody.appendParagraph(`Timestamp: ${timestamp}`).setHeading(DocumentApp.ParagraphHeading.HEADING2);

  docBody.appendHorizontalRule();

  // Add non-empty and relevant data to the document, excluding Machine Number and Timestamp
  for (let i = 0; i < rowData.length; i++) {
    if (rowData[i] && headers[i].toLowerCase() !== 'report' && headers[i].toLowerCase() !== 'machine' && headers[i].toLowerCase() !== 'timestamp') {
      docBody.appendParagraph(`${headers[i]}: ${rowData[i]}`);
      docBody.appendParagraph(''); // Add a blank line after each entry
    }
  }

  // Add generation timestamp at the bottom of the document
  docBody.appendHorizontalRule();
  docBody.appendParagraph(`Generated on: ${new Date()}`).setHeading(DocumentApp.ParagraphHeading.NORMAL);

  // Get the document URL
  const docUrl = templateDoc.getUrl();

  // Write the clickable link to column T for the current row
  const linkFormula = `=HYPERLINK("${docUrl}", "Open Document")`;
  sheet.getRange(row, lastColumn).setFormula(linkFormula);

  // Notify the user
  SpreadsheetApp.getUi().alert('Document created successfully! A link has been added to column T.');
}

// Helper function to extract machine number and timestamp
function getMachineAndTimestamp(headers, rowData) {
  const machineIndex = headers.findIndex(header => header.toLowerCase().includes('machine'));
  const timestampIndex = headers.findIndex(header => header.toLowerCase().includes('timestamp'));

  const machineNumber = machineIndex !== -1 ? rowData[machineIndex] : 'Unknown';
  const timestamp = timestampIndex !== -1 ? rowData[timestampIndex] : 'Not Available';

  return { machineNumber, timestamp };
}

function copyDocumentData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = sheet.getActiveRange();
  const row = range.getRow();

  // Ensure the user has selected a row and not the header
  if (row === 1) {
    SpreadsheetApp.getUi().alert('Please select a data row, not the header row.');
    return;
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Collect relevant data
  let documentData = '';
  const { machineNumber, timestamp } = getMachineAndTimestamp(headers, rowData);

  // Add Machine Number and Timestamp
  documentData += `Machine Number: ${machineNumber}\n\n`;
  documentData += `Timestamp: ${timestamp}\n\n`;

  // Append other data excluding duplicates and unwanted fields
  for (let i = 0; i < rowData.length; i++) {
    if (
      rowData[i] &&
      headers[i].toLowerCase() !== 'timestamp' &&
      headers[i].toLowerCase() !== 'machine' &&
      headers[i].toLowerCase() !== 'report'
    ) {
      documentData += `${headers[i]}:\n${rowData[i]}\n\n`;
    }
  }

  // Create the HTML for the sidebar
  const htmlContent = `
    <div>
      <h3>Copy Data</h3>
      <textarea id="dataToCopy" rows="15" style="width: 100%;">${documentData}</textarea>
      <br>
      <button onclick="copyToClipboard()">Copy to Clipboard</button>
      <button onclick="openWhatsApp()">Open WhatsApp Web</button>
      <script>
        function copyToClipboard() {
          const textarea = document.getElementById('dataToCopy');
          textarea.select();
          document.execCommand('copy');
          alert('Data copied to clipboard!');
        }
        function openWhatsApp() {
          window.open('https://web.whatsapp.com/', '_blank');
        }
      </script>
    </div>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setTitle('Copy and Share Data')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}
