function extractDetails(html) {
  const regexPattern =
    /(\d{2}\/\d{2}\/\d{4} - \d{2}\/\d{2}\/\d{4})<br>([\w\d]+ - .+? - .+?)<br>Listing ID: (\d+)/g;

  let matches = [];
  let match;

  while ((match = regexPattern.exec(html)) !== null) {
    matches.push({
      dateRange: match[1],
      details: match[2],
      listingId: match[3],
    });
  }

  return matches;
}

function generateIncrementingID(dateStr, currentIncrement) {
  const dateParts = dateStr.split("-");
  const month = new Date(dateParts[1] + " 1, " + dateParts[2]).getMonth() + 1;
  const twoDigitMonth = String(month).padStart(2, "0");
  const twoDigitYear = dateParts[2].slice(-2);
  const monthYear = twoDigitYear + twoDigitMonth;

  const formattedCount = String(currentIncrement).padStart(2, "0");
  return monthYear + formattedCount;
}

function createReceiptsAndDocs() {
  try {
    const parentFolderId = "136g8n4fia__gnHdzFlEESMpiHEDln4FK";
    const parentFolder = DriveApp.getFolderById(parentFolderId);
    const templateId = "1yOWzl8ZjI2aadUny_5MvQxiR334f_wmO1GfHc4z5yW8";

    const currentDate = new Date();
    const firstDayOfMonth = new Date(
      currentDate.getFullYear(),
      currentDate.getMonth(),
      1
    );
    const monthYearFolderName =
      Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "yyMM") +
      " Payout & Receipts";
    const searchQuery = `from:express@airbnb.com after:${Utilities.formatDate(
      firstDayOfMonth,
      Session.getScriptTimeZone(),
      "yyyy/MM/dd"
    )} "Payout of ฿"`;
    const threads = GmailApp.search(searchQuery);
    const reversedThreads = threads.reverse(); // Reverse the threads to start from the beginning of the month

    let monthYearFolders = parentFolder.getFoldersByName(monthYearFolderName);
    let monthYearFolder;
    if (monthYearFolders.hasNext()) {
      monthYearFolder = monthYearFolders.next();
    } else {
      monthYearFolder = parentFolder.createFolder(monthYearFolderName);
    }

    let receiptCount = 0;
    let totalAmount = 0;
    let currentIncrement = 1; // Initialize the incrementing number

    reversedThreads.forEach((thread) => {
      const messages = thread.getMessages();
      messages.forEach((message) => {
        const htmlBody = message.getBody();
        const subject = message.getSubject();
        const dateReceived = message.getDate();
        const from = message.getFrom();

        const amountMatch = htmlBody.match(/Payout of ฿([\d,\.]+)/);
        const dateMatch = htmlBody.match(
          /arrive in your account by ([\w\s,]+)/
        );

        if (amountMatch && dateMatch) {
          const amount = parseFloat(amountMatch[1].replace(",", ""));
          totalAmount += amount;
          const formattedAmount = amount.toLocaleString();
          const dateStr = dateMatch[1].split(",")[0];
          const year = dateMatch[1].split(",")[1].trim();
          const formattedDate = `${dateStr}-${year}`.replace(" ", "-");

          const incrementingID = generateIncrementingID(
            formattedDate,
            currentIncrement
          ); // Pass the current increment
          currentIncrement++; // Increment for the next set

          const templateFile = DriveApp.getFileById(templateId);
          const newFile = templateFile.makeCopy(
            `${incrementingID} Receipt`,
            monthYearFolder
          );
          const newSheet = SpreadsheetApp.openById(newFile.getId());
          const sheet = newSheet.getSheets()[0];
          sheet.getRange("L25").setValue(formattedAmount);
          sheet.getRange("L8").setValue(formattedDate);
          sheet.getRange("L9").setValue(incrementingID);

          const extractedDetails = extractDetails(htmlBody);
          let row = 20;
          extractedDetails.forEach((detail) => {
            sheet.getRange(`A${row}`).setValue(detail.details);
            row++;
          });

          const docName = `${incrementingID} Payout`;
          const docFile = DocumentApp.create(docName);
          const doc = DocumentApp.openById(docFile.getId());
          let header = doc.getHeader();
          if (!header) {
            header = doc.addHeader();
          }
          header.setText(`ID: ${incrementingID}`);

          const docText =
            extractedDetails
              .map(
                (detail) =>
                  `Date Range: ${detail.dateRange}\nDetails: ${detail.details}\nListing ID: ${detail.listingId}`
              )
              .join("\n\n") + `\n\nPayout Amount: ${formattedAmount}`;
          const docBody = doc.getBody();
          docBody.setText(docText);

          const attachments = message.getAttachments();
          attachments.forEach((attachment) => {
            const blob = attachment.copyBlob();
            monthYearFolder.createFile(blob);
          });

          const docFileInDrive = DriveApp.getFileById(docFile.getId());
          docFileInDrive.moveTo(monthYearFolder);

          receiptCount++;
        }
      });
    });

    const formattedTotalAmount = "฿" + totalAmount.toLocaleString();
    GmailApp.sendEmail(
      "mangotreevillaphuket@gmail.com",
      "Receipts and Docs Created",
      `Receipts and Docs have been created for ${monthYearFolderName}. Total Receipts: ${receiptCount}. Total Amount: ${formattedTotalAmount}`
    );
  } catch (error) {
    console.error("Error creating receipts and docs:", error);
  }
}

function createTimeDrivenTrigger() {
  const currentDate = new Date();
  ScriptApp.newTrigger("createReceiptsAndDocs")
    .timeBased()
    .atDate(currentDate.getFullYear(), currentDate.getMonth() + 1, 0) // 0 represents the last day of the current month
    .atHour(23)
    .nearMinute(59)
    .create();
}
