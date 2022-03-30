
// This constant is written in column G for rows for which an email
// has been sent successfully.
let EMAIL_SENT = 'Sent';

/**
 * Sends non-duplicate emails with data from the current spreadsheet.
 */
function sendNonDuplicateEmails() {
  try{
    // Get the active sheet in spreadsheet
    const sheet = SpreadsheetApp.getActiveSheet();
    let startRow = 3; // First row of data to process
    let numRows = 17; // Number of rows to process
    // Fetch the range of cells C2:E18
    const dataRange = sheet.getRange(startRow, 3, numRows, 3);
    // Fetch values for each row in the Range.
    const data = dataRange.getValues();
    for (let i = 0; i < data.length; ++i) {
      const row = data[i];
      const emailAddress = row[1]; // Fourth Column in our selection
      console.log(row)
      var email_body = 
      `Greetings,${row[0]} <br/><br/>
      &nbsp;&nbsp;&nbsp;&nbsp;Hope this email finds you well.
      Kindly find your credentials to join ITI- وثبات - on pluralsight as follows
      
      
      <br/>
      <ul>
        <li><strong>Name:</strong> ${row[0]}</li>
        <li><strong>Email:</strong>  ${row[2]}</li>
        <li><strong>Registeration link:</strong>  ${row[4]}</li>
      </ul>
      &nbsp;&nbsp;&nbsp;&nbsp; 
      <h3 style="color:red"> Please note that the link is valid for pnly one click, and You should compelete the registeration when you click it. </h3>
      <br/> 
      <br/>
    Best of luck ^^ <br/>
    Yours,<br/>
     Noha Shehab <br/>`
      const emailSent = row[7]; // Fifth Column in our selection
      if (emailSent !== EMAIL_SENT) { // Prevents sending duplicates
        let subject = 'ITI- Devops وثبات';
        // Send emails to emailAddresses which are presents in Fourth column
        MailApp.sendEmail( {to:emailAddress, name:'Noha Shehab', subject:subject, body:email_body,htmlBody:email_body});
        sheet.getRange(startRow + i, 8).setValue(EMAIL_SENT);
        // Make sure the cell is updated right away in case the script is interrupted
        SpreadsheetApp.flush();
      }
    }
  }
  catch(err){
    Logger.log(err)
  }
}
