function myFunction() {

  // Exact sheet name
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");

  if (!sheet) {
    Logger.log("❌ Sheet 'Form Responses 1' not found! Check the sheet name.");
    return;
  }

  var data = sheet.getDataRange().getValues();
  var positionTitle = "Contract of Service – Information and Communications Technology (COS-ICT)";

  for (var i = 1; i < data.length; i++) { // Skip header row

    var emailAddress = data[i][1];  // Column 2 = Email Address
    var existingCode = data[i][25]; // Column 26 = Code (index 25 because index starts at 0)

    if (!emailAddress) continue;

    if (!existingCode) {

      var uniqueCode = generateCOSICTCode(positionTitle, emailAddress);

      // Write generated code to Column 26
      sheet.getRange(i + 1, 26).setValue(uniqueCode);

      // Email message
      var subject = "Application Confirmation – " + positionTitle;
      var body =
`Dear Applicant,

Thank you for submitting your application for the ${positionTitle} position.

Your unique application tracking code is:

👉 ${uniqueCode}

Please keep this code for future reference as it will be used to track and verify your application status.

If you have questions or require assistance, feel free to contact the HR Office.

Thank you and we wish you success in the selection process.

Sincerely,
Human Resource Management Office
`;

      try {
        MailApp.sendEmail({
          to: emailAddress,
          subject: subject,
          body: body
        });

        Logger.log("📨 Email sent to: " + emailAddress + " | Code: " + uniqueCode);

      } catch (e) {
        Logger.log("❌ Failed to send to: " + emailAddress + " | Error: " + e);
      }
    }
  }
}

function generateCOSICTCode(position, email) {
  var letters = position.match(/\b\w/g) || [];
  var initials = letters.join('').toUpperCase(); // COSICT
  var randomNum = Math.floor(100000 + Math.random() * 900000);
  var emailTag = email.split('@')[0].slice(-2).toUpperCase();

  return initials + "-" + emailTag + "-" + randomNum;
}
