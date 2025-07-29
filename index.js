function sendMailOnCondition() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();

  // Collect all email addresses from column O (15th column)
  var emailRange = sheet.getRange(1, 15, lastRow, 1); // Column O
  var emailAddresses = emailRange.getValues().flat().filter(String).join(',');

  // Always fetch client details from D2
  var clientDetailsLink = sheet.getRange(2, 4).getValue() || 'No Client Details Provided';
  if (clientDetailsLink !== 'No Client Details Provided') {
    clientDetailsLink = '<a href="' + clientDetailsLink + '" target="_blank">Click here</a>';
  }

  for (var row = 1; row <= lastRow; row++) {
    var sendStatus = sheet.getRange(row, 13).getValue(); // Column M
    var emailSentStatus = sheet.getRange(row, 14).getValue(); // Column N

    if (sendStatus.toLowerCase() === 'send' && emailSentStatus.toLowerCase() !== 'sent') {
      var subject = sheet.getRange(row, 11).getValue() + " Task"; // Column K
      var client = sheet.getName();
      var subtopic = sheet.getRange(row, 7).getValue() || 'Nothing Specific'; // Column G
      var idea = sheet.getRange(row, 8).getValue() || 'Design on your own'; // Column H
      var reference = sheet.getRange(row, 10).getValue() || 'No Reference Available'; // Column J
      if (reference !== 'No Reference Available') {
        reference = '<a href="' + reference + '" target="_blank">' + reference + '</a>';
      }
      var deadline = sheet.getRange(row, 12).getValue() || 'Submit by Day End'; // Column L

      var randomChoice = Math.random() < 0.5 ? 'Rishi Agrawal' : 'Abhishek Tiwari';
      var number = randomChoice === 'Rishi Agrawal' ? '9045655504' : '7505990012';

      var message = '<b>Client:</b> ' + client + '<br>';
      message += '<b>Client Details:</b> ' + clientDetailsLink + '<br>'; // Always from D2
      message += '<b>Subtopic:</b> ' + subtopic + '<br>';
      message += '<b>Idea:</b> ' + idea + '<br>';
      message += '<b>Reference:</b> ' + reference + '<br>';
      message += '<b>Submission Deadline:</b> ' + deadline + '<br>';
      message += '<b>Submit Your design to your mentor over WhatsApp before deadline.</b><br><br>';
      message += '<b>Best Regards,</b><br>';
      message += '<b>' + randomChoice + '</b><br>';
      message += '<b>' + number + '</b><br><br>';
      message += '<span style="color: green;">If you encounter any issues, please reach out to your mentor.</span><br>';
      message += '<span style="color: red;">This is an automated response. Please do not reply to this email.</span>';

      var senderEmail = 'moodale2020@gmail.com';
      var ccEmail = 'mridulagrawal60@gmail.com';

      try {
        MailApp.sendEmail({
          to: emailAddresses,
          subject: subject,
          htmlBody: message,
          cc: ccEmail,
          name: "Moodale Content Department"
        });

        sheet.getRange(row, 14).setValue('Sent'); // Mark as Sent
        Logger.log('Email sent to: ' + emailAddresses + ' for row: ' + row);
      } catch (e) {
        sheet.getRange(row, 14).setValue('Not Sent');
        Logger.log('Failed to send email for row: ' + row + ' due to error: ' + e.message);
      }
    }
  }
}
