function approveEntry() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const row = sheet.getActiveRange().getRow();

  // 1️⃣ CHECK BUDGET FIRST
  if (isBudgetExceeded(row)) {
    const reason = "Budget exceeded. Cannot approve. Auto-rejected.";
    sheet.getRange(row, 9).setValue("Rejected");   // I
    sheet.getRange(row, 10).setValue("");          // J
    sheet.getRange(row, 11).setValue(reason);      // K
    notifyRejection(row, reason);
    SpreadsheetApp.getActiveSpreadsheet().toast("Auto-Rejected: Budget Exceeded");
    return;  // ❗ STOP approval
  }

  // 2️⃣ Deduct Budget
  deductBudget(row);

  // 3️⃣ Approve
  sheet.getRange(row, 9).setValue("Approved");
  sheet.getRange(row, 11).setValue("");

  generateApprovalPDF(row);
}
