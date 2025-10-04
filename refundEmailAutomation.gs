/**
 * Sends a personalized refund email automatically
 * when a new response is submitted to the "Refund Responses" sheet.
 */
function onFormSubmit(e) {
  if (!e || !e.range) {
    Logger.log("⚠️ Must be triggered by a form submission.");
    return;
  }

  const sheet = e.range.getSheet();
  if (sheet.getName() !== "Refund Responses") return;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // match columns (case-insensitive)
  const nameCol  = headers.findIndex(h => /full\s*name/i.test(h));   // <-- Full Name
  const emailCol = headers.findIndex((h, idx) => /email/i.test(h) && idx > 0);  
  // above: picks the SECOND "Email Address" (the one filled in form, not Timestamp one)

  let sentCol = headers.findIndex(h => /email\s*sent/i.test(h));

  if (nameCol < 0 || emailCol < 0) {
    Logger.log("❌ Couldn’t find Full Name or Email Address headers.");
    return;
  }

  // if missing, add "Email Sent" column
  if (sentCol < 0) {
    sentCol = headers.length;
    sheet.getRange(1, sentCol + 1).setValue("Email Sent");
  }

  const row = e.range.getRow();
  if (row === 1) return;

  const studentName  = (sheet.getRange(row, nameCol + 1).getValue()  || "").toString().trim();
  const studentEmail = (sheet.getRange(row, emailCol + 1).getValue() || "").toString().trim();

  // skip if already marked
  const sentStatus = sheet.getRange(row, sentCol + 1).getValue();
  if (["SENT", "ERROR", "INVALID"].includes(sentStatus)) return;

  if (!studentName || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(studentEmail)) {
    Logger.log(`⚠️ Invalid entry row ${row}: name="${studentName}", email="${studentEmail}"`);
    sheet.getRange(row, sentCol + 1).setValue("INVALID");
    return;
  }

  const subject = "Refund Request Received";
  const body =
    `Dear ${studentName},\n\n` +
    `We’ve received your refund request, and the amount will be processed and refunded within 15–20 working days.\n\n` +
    `We appreciate your patience during this time. If you have any questions or need further assistance, feel free to reply to this email. We’re here to help.\n\n` +
    `Best regards,\n` +
    `Team atomcamp`;

  try {
    MailApp.sendEmail(studentEmail, subject, body);
    sheet.getRange(row, sentCol + 1).setValue("SENT");
    Logger.log(`✅ Email sent to ${studentEmail} (row ${row})`);
  } catch (err) {
    sheet.getRange(row, sentCol + 1).setValue("ERROR");
    Logger.log(`❌ Error sending to ${studentEmail}: ${err}`);
  }
}
