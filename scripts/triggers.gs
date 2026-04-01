function onFormSubmit(e) {
  // Get the active sheet where the form submitted data
  const sheet = e.source.getSheetByName("Form Responses 1");
  const row = e.range.getRow();
  
  // Retrieve data from the newly added row
  const timestamp = sheet.getRange(row, 1).getValue(); // Form Timestamp
  const facultyName = sheet.getRange(row, 2).getValue();
  const facultyEmail = sheet.getRange(row,3).getValue();
  const department = sheet.getRange(row, 4).getValue();
  const expenseType = sheet.getRange(row, 5).getValue();
  const description = sheet.getRange(row, 6).getValue();
  const amount = sheet.getRange(row, 7).getValue();

  // Email notification to Admin/HOD
  MailApp.sendEmail({
    to: "parjanya.gr@gmail.com", // Change to your admin mail
    subject: "New Expense Submission from " + facultyName,
    htmlBody: `
      <p><b>${facultyName}</b> from <b>${department}</b> has submitted a new expense request.</p>
      <p><b>Type:</b> ${expenseType}<br>
      <b>Amount:</b> ₹${amount}<br>
      <b>Description:</b> ${description}</p>
      <p>Submitted on: ${timestamp}</p>
      <p>View it in the <a href="${SpreadsheetApp.getActiveSpreadsheet().getUrl()}">Budget Tracker Sheet</a>.</p>
    `
  });
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Budget Tools")
    .addItem("Approve Selected row", "manualApproval")
    .addItem("Reject Selected Row", "rejectEntry")
    .addToUi();
}

function manualApproval() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const row = sheet.getActiveRange().getRow();
  generateApprovalPDF(row);
}
