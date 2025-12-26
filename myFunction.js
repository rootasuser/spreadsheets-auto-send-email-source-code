function myFunction() {

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");

  if (!sheet) {
    Logger.log("❌ Sheet NOT FOUND. Available sheets:");
    checkSheetNames();
    return;
  }

  var data = sheet.getDataRange().getValues();
  var positionTitle = "Application for Administrative Aide VI (ADA6)";

  for (var i = 1; i < data.length; i++) {

    var emailAddress = data[i][1];
    var existingCode = data[i][25];

    Logger.log("Row " + (i+1) + " Email: " + emailAddress + " | Existing Code: " + existingCode);

    if (!emailAddress) continue;

    if (!existingCode) {

      var uniqueCode = generateADA6Code(positionTitle, emailAddress);

      sheet.getRange(i + 1, 26).setValue(uniqueCode);

      var subject = "Application Confirmation – " + positionTitle;
      var body =
`Dear Applicant,

Thank you for applying for the ${positionTitle} position.

Your unique application tracking code is:

👉 ${uniqueCode}

Please keep this code for tracking your application status.

HRM Office`;

      try {
        MailApp.sendEmail({
          to: emailAddress,
          subject: subject,
          body: body
        });

        Logger.log("📨 Email sent to: " + emailAddress);

      } catch (e) {
        Logger.log("❌ Email FAILED for " + emailAddress + " | " + e);
      }
    }
  }
}

function generateADA6Code(position, email) {
  var initials = "ADA6";
  var randomNum = Math.floor(100000 + Math.random() * 900000);
  var emailTag = email.split('@')[0].slice(-2).toUpperCase();
  return initials + "-" + emailTag + "-" + randomNum;
}
