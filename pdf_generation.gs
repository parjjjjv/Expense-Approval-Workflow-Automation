function generateApprovalPDF(row) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
  const data = sheet.getRange(row, 1, 1, 7).getValues()[0];

  const docId = "1Zz_x1niBjX96HdbelHMMEi-WeJ_Y_cFE9ywC8Nz4HTc";
  const folder = DriveApp.getFolderById("1MAqLjRJaqiMs_5WdCo5m1T4089o02J3U");

  const copy = DriveApp.getFileById(docId)
    .makeCopy(`Approval_${data[1]}_${new Date().toDateString()}`, folder);
  const doc = DocumentApp.openById(copy.getId());
  const body = doc.getBody();

  const replacements = {
    "{{FacultyName}}": data[1],
    "{{FacultyEmail}}": data[2],
    "{{Department}}": data[3],
    "{{ExpenseType}}": data[4],
    "{{Description}}": data[5],
    "{{Amount}}": data[6],
    "{{Status}}": "Approved",
    "{{ApprovedBy}}": "VC",
    "{{Date}}": new Date().toLocaleDateString()
  };

  for (const [k, v] of Object.entries(replacements)) body.replaceText(k, v);

  const tables = body.getTables();
  tables.forEach(t => {
    for (let r = 0; r < t.getNumRows(); r++) {
      for (let c = 0; c < t.getRow(r).getNumCells(); c++) {
        const cell = t.getRow(r).getCell(c);
        for (const [k, v] of Object.entries(replacements)) cell.replaceText(k, v);
      }
    }
  });

  const header = doc.getHeader();
  if (header) for (const [k, v] of Object.entries(replacements)) header.replaceText(k, v);

  const footer = doc.getFooter();
  if (footer) for (const [k, v] of Object.entries(replacements)) footer.replaceText(k, v);

  doc.saveAndClose();

  // PDF generation
  const pdf = copy.getAs("application/pdf");
  const pdfFile = folder.createFile(pdf)
    .setName(`Approval_${data[1]}_${data[4]}_${new Date().toLocaleDateString()}.pdf`);

  // Auto-share
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // Update J (PDF Link), I (Status)
  sheet.getRange(row, 10).setValue(pdfFile.getUrl());    // J
  sheet.getRange(row, 9).setValue("Approved");           // I

  // Email faculty
  notifyApproval(row, pdfFile.getUrl());
}
