function createReceiptsAndDocs() {
  try {
    const parentFolderId = "1EuXjX9fYAuZ2pkz1_o9eZEzZIxnJ1H9K"; // Replace with your actual folder ID
    const parentFolder = DriveApp.getFolderById(parentFolderId);
    const templateId = "1Tb9VPZeVIiwQkCA0lshZwzMRmEELJxtxFytuJXNE3ZI"; // Replace with your actual template ID

    const currentDate = new Date();
    const firstDayOfMonth = new Date(
      currentDate.getFullYear(),
      currentDate.getMonth(),
      1
    );
    const monthYearFolderName =
      Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "MMyyyy") +
      "payout";
    const searchQuery = `from:express@airbnb.com after:${Utilities.formatDate(
      firstDayOfMonth,
      Session.getScriptTimeZone(),
      "yyyy/MM/dd"
    )} "Payout of ฿"`;
    const threads = GmailApp.search(searchQuery);

    let monthYearFolders = parentFolder.getFoldersByName(monthYearFolderName);
    let monthYearFolder;
    if (monthYearFolders.hasNext()) {
      monthYearFolder = monthYearFolders.next();
    } else {
      monthYearFolder = parentFolder.createFolder(monthYearFolderName);
    }

    let receiptCount = 0;
    let totalAmount = 0;
    threads.forEach((thread) => {
      const messages = thread.getMessages();
      messages.forEach((message) => {
        const htmlBody = message.getBody();
        const extractedDetails = extractDetails(htmlBody);

        if (extractedDetails.length > 0) {
          const templateFile = DriveApp.getFileById(templateId);
          const newFile = templateFile.makeCopy(
            `Payment ${extractedDetails[0].dateRange}`,
            monthYearFolder
          );
          const newSheet = SpreadsheetApp.openById(newFile.getId());
          const sheet = newSheet.getSheets()[0];

          extractedDetails.forEach((detail, index) => {
            sheet.getRange(`D${20 + index}`).setValue(detail.dateRange);
            const reservationCode = detail.details.split(" - ")[0];
            const guestName = detail.details.split(" - ")[1];
            sheet.getRange(`B${20 + index}`).setValue(reservationCode);
            sheet.getRange(`C${20 + index}`).setValue(guestName);
            // Assuming the amount is the last part of the details string
            const amount = detail.details.split(" - ").pop();
            totalAmount += parseFloat(amount.replace("฿", "").replace(",", ""));
            sheet.getRange(`L${20 + index}`).setValue(amount);
          });

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

function createTimeDrivenTrigger() {
  const currentDate = new Date();
  ScriptApp.newTrigger("createReceiptsAndDocs")
    .timeBased()
    .atDate(currentDate.getFullYear(), currentDate.getMonth() + 1, 0) // 0 represents the last day of the current month
    .atHour(23)
    .nearMinute(59)
    .create();
}
