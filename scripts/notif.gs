function notifyRejection(row, reason) {
  const sheet = SpreadsheetApp.getActiveSheet();
  const data = sheet.getRange(row, 1, 1, 7).getValues()[0];

  const name = data[1];
  const email = data[2];
  const type = data[4];
  const amount = data[6];
  const dept = data[3];

  MailApp.sendEmail({
    to: email,
    subject: `Your Expense Request for ${type} Has Been Rejected`,
    htmlBody: `
      <p>Dear <b>${name}</b>,</p>
      <p>Your expense request for <b>${type}</b> (₹${amount}) has been <b>rejected</b>.</p>
      <p><b>Reason:</b> ${reason}</p>
      <p>Regards,<br>VC Office</p>
    `
  });
}
function notifyApproval(row, pdfUrl) {
  const sheet = SpreadsheetApp.getActiveSheet();
  const data = sheet.getRange(row, 1, 1, 7).getValues()[0];

  const name = data[1];
  const email = data[2];
  const dept = data[3];
  const type = data[4];
  const amount = data[6];

  MailApp.sendEmail({
    to: email,
    subject: `Your Expense Request for ${type} Has Been Approved`,
    htmlBody: `
      <p>Dear <b>${name}</b>,</p>
      <p>Your request for <b>${type}</b> amounting to ₹${amount} has been approved.</p>
      <p><a href="${pdfUrl}">Click here to view your approval PDF</a></p>
      <p>Regards,<br>VC OFfice</p>
    `
  });
}
