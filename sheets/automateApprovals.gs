/**
 * Sends an approval request email with a link
 * to the active Google Spreadsheet.
 */
function sendApprovalRequest() {
  const recipientEmail = 'jawherbouhouch7@gmail.com';
  const spreadsheet = SpreadsheetApp.getActive();
  const spreadsheetUrl = spreadsheet.getUrl();
  const spreadsheetName = spreadsheet.getName();

  const emailSubject = 'Approval Request Required';
  const emailBody = buildApprovalEmailBody(spreadsheetName, spreadsheetUrl);

  MailApp.sendEmail({
    to: recipientEmail,
    subject: emailSubject,
    htmlBody: emailBody
  });

  SpreadsheetApp.getActive().toast('Approval request email sent successfully.');
}

function buildApprovalEmailBody(sheetName, sheetUrl) {
  return `
    <p>Hello,</p>

    <p>An approval is required for the following spreadsheet:</p>

    <ul>
      <li><strong>Name:</strong> ${sheetName}</li>
      <li><strong>Link:</strong>
        <a href="${sheetUrl}" target="_blank">Open Spreadsheet</a>
      </li>
    </ul>

    <p>Please review and approve at your convenience.</p>

    <p>Thank you.</p>
  `;
}
