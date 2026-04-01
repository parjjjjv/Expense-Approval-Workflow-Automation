function rejectEntry() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  const row = sheet.getActiveRange().getRow();

  const response = ui.prompt(
    "Reject Expense",
    "Enter reason for rejection:",
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    ui.alert("Rejection cancelled.");
    return;
  }

  const reason = response.getResponseText().trim();
  if (!reason) {
    ui.alert("Reason cannot be empty!");
    return;
  }

  // Update columns
  sheet.getRange(row, 9).setValue("Rejected");   // I = Status
  sheet.getRange(row, 10).setValue("");          // J = Clear PDF Link
  sheet.getRange(row, 11).setValue(reason);      // K = Rejection Reason

  notifyRejection(row, reason);

  SpreadsheetApp.getActiveSpreadsheet().toast("Entry Rejected & Mail Sent");
}
