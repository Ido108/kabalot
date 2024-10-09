async function downloadGmailAttachments(auth, startDate, endDate, folderPath) {
  const gmail = google.gmail({ version: 'v1', auth });

  // Define keywords to exclude based on the subject line
  const excludedSubjectKeywords = ['חשבון עסקה'];

  // Adjust the end date to include the entire day
  endDate.setHours(23, 59, 59, 999);
  const queryEndDate = new Date(endDate.getTime() + 24 * 60 * 60 * 1000); // Add one day to include end of day

  // Prepare date queries
  const startDateQuery = formatDateForGmail(startDate);
  const endDateQuery = formatDateForGmail(queryEndDate);

  const query = `after:${startDateQuery} before:${endDateQuery}`;
  console.log('Gmail query:', query);

  // Senders to exclude and keywords to look for
  const excludedSenders = [
    'חברת חשמל לישראל',
    'עיריית תל אביב-יפו',
    'ארנונה - עיריית תל-אביב-יפו',
  ];
  const keywords = ['קבלה', 'חשבונית', 'חשבונית מס', 'הקבלה'];

  let nextPageToken = null;
  const allMessageIds = [];

  // Fetch all message IDs matching the query
  do {
    const res = await gmail.users.messages.list({
      userId: 'me',
      q: query,
      pageToken: nextPageToken,
      maxResults: 500,
    });
    const messages = res.data.messages || [];
    allMessageIds.push(...messages);
    nextPageToken = res.data.nextPageToken;
  } while (nextPageToken);

  // Process each message
  for (const messageData of allMessageIds) {
    const msg = await gmail.users.messages.get({
      userId: 'me',
      id: messageData.id,
      format: 'full',
    });

    const headers = msg.data.payload.headers;
    const fromHeader = headers.find((h) => h.name === 'From');
    const subjectHeader = headers.find((h) => h.name === 'Subject');
    const dateHeader = headers.find((h) => h.name === 'Date');

    const sender = fromHeader ? fromHeader.value : '';
    const subject = subjectHeader ? subjectHeader.value : '';
    const messageDateStr = dateHeader ? dateHeader.value : '';
    const messageDate = new Date(messageDateStr);

    // Check if message date is within range
    if (messageDate < startDate || messageDate > endDate) {
      continue;
    }

    // Exclude messages with certain keywords in the subject
    let subjectContainsExcludedKeyword = false;
    for (const excludedKeyword of excludedSubjectKeywords) {
      if (subject.includes(excludedKeyword)) {
        subjectContainsExcludedKeyword = true;
        console.log(
          `Skipping message with subject containing excluded keyword: ${excludedKeyword}`
        );
        break;
      }
    }

    if (subjectContainsExcludedKeyword) {
      // Skip this message
      continue;
    }

    // Exclusion logic based on sender and keywords
    let excludeThread = false;
    let keywordFound = false;

    for (const excludedSender of excludedSenders) {
      if (sender.includes(excludedSender)) {
        excludeThread = true;
        for (const keyword of keywords) {
          if (subject.includes(keyword)) {
            keywordFound = true;
            break;
          }
        }
        break;
      }
    }

    // Decide whether to skip the thread
    if (excludeThread && !keywordFound) {
      console.log('Skipping message from excluded sender:', sender);
      continue;
    }

    // Initialize flag to check for receipts in the message
    let receiptFoundInThread = false;

    // Ensure msg.data.payload.parts exists
    if (msg.data.payload && msg.data.payload.parts) {
      // First pass: check if there's a receipt PDF in the message
      for (const part of msg.data.payload.parts) {
        if (part.filename && part.filename.length > 0) {
          const normalizedFileName = part.filename.toLowerCase();
          if (
            normalizedFileName.startsWith('receipt') &&
            part.mimeType === 'application/pdf'
          ) {
            receiptFoundInThread = true;
            break; // Found a receipt in this message
          }
        }
      }

      // Second pass: process the attachments based on whether receipt was found
      for (const part of msg.data.payload.parts) {
        if (part.filename && part.filename.length > 0) {
          const attachmentId = part.body.attachmentId;
          if (!attachmentId) continue;

          const attachment = await gmail.users.messages.attachments.get({
            userId: 'me',
            messageId: messageData.id,
            id: attachmentId,
          });

          const data = attachment.data.data;
          const buffer = Buffer.from(data, 'base64');

          const contentType = part.mimeType;
          const fileName = part.filename;
          const isPDF = isPdfFile(contentType, fileName);

          if (isPDF) {
            const normalizedFileName = fileName.toLowerCase();
            if (receiptFoundInThread) {
              // If receipt is found in the thread, collect only PDFs starting with "receipt"
              if (normalizedFileName.startsWith('receipt')) {
                const filePath = path.join(folderPath, sanitize(fileName));
                fs.writeFileSync(filePath, buffer);
                console.log(`Saved attachment: ${filePath}`);
              } else {
                // Skip other PDFs in this thread
                console.log('Skipping non-receipt PDF in receipt thread:', fileName);
              }
            } else {
              // If no receipt is found, collect the PDF as usual
              const filePath = path.join(folderPath, sanitize(fileName));
              fs.writeFileSync(filePath, buffer);
              console.log(`Saved attachment: ${filePath}`);
            }
          }
        }
      }
    }
  }

  return folderPath;
}