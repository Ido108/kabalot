// server.js

require('dotenv').config(); // Load environment variables
const { v4: uuidv4 } = require('uuid');
const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs-extra');
const axios = require('axios');
const FormData = require('form-data');
const sanitize = require('sanitize-filename');
const { PDFDocument } = require('pdf-lib');
const { google } = require('googleapis');
const cors = require('cors');
const Excel = require('exceljs');
const { parse, format } = require('date-fns');
const session = require('express-session');
const archiver = require('archiver');
const nodemailer = require('nodemailer');
const EventEmitter = require('events');

const app = express();

// Middleware
app.use(cors());
app.use(express.static('public'));
app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.set('view engine', 'ejs');

// Configure session
app.use(
  session({
    secret: process.env.SESSION_SECRET || 'your_session_secret',
    resave: false,
    saveUninitialized: true,
    cookie: {
      maxAge: 3600000, // 1 hour
    },
  })
);

const PDFCO_API_KEY = process.env.PDFCO_API_KEY;
const PASSWORD_PROTECTED_PDF_PASSWORD =
  process.env.PASSWORD_PROTECTED_PDF_PASSWORD || 'your-default-password';

const DOCUMENT_AI_CONFIG = {
  projectId: process.env.DOCUMENT_AI_PROJECT_ID || 'your-project-id',
  location: process.env.DOCUMENT_AI_LOCATION || 'us',
  processorId: process.env.DOCUMENT_AI_PROCESSOR_ID || 'your-processor-id',
};

const SERVICE_ACCOUNT_BASE64 = process.env.SERVICE_ACCOUNT_BASE64;
const SCOPES = ['https://www.googleapis.com/auth/cloud-platform'];

function authenticateServiceAccount() {
  if (!SERVICE_ACCOUNT_BASE64) {
    throw new Error('SERVICE_ACCOUNT_BASE64 not set');
  }
  const serviceAccountJson = Buffer.from(SERVICE_ACCOUNT_BASE64, 'base64').toString('utf8');
  const serviceAccount = JSON.parse(serviceAccountJson);
  const jwtClient = new google.auth.JWT(
    serviceAccount.client_email,
    null,
    serviceAccount.private_key,
    SCOPES,
    null
  );
  return jwtClient;
}

function formatDate(date) {
  return format(date, 'yyyy-MM-dd');
}

function cleanAndParseAmount(amountStr) {
  if (!amountStr) return 0;
  amountStr = amountStr.replace(/[^0-9.\-]+/g, '');
  const parsedAmount = parseFloat(amountStr);
  return isNaN(parsedAmount) ? 0 : parsedAmount;
}

async function getUsdToIlsExchangeRate(date) {
  // Implement actual logic or use a fixed rate
  return 3.5;
}

async function unlockPdf(filePath, password) {
  if (!PDFCO_API_KEY) {
    throw new Error('PDFCO_API_KEY is not set');
  }

  try {
    const formData = new FormData();
    formData.append('file', fs.createReadStream(filePath));
    const uploadResponse = await axios.post(
      'https://api.pdf.co/v1/file/upload',
      formData,
      {
        headers: {
          'x-api-key': PDFCO_API_KEY,
          ...formData.getHeaders(),
        },
      }
    );

    if (!uploadResponse.data || !uploadResponse.data.url) {
      console.error('Error uploading PDF:', uploadResponse.data);
      return null;
    }

    const uploadedFileUrl = uploadResponse.data.url;
    const unlockResponse = await axios.post(
      'https://api.pdf.co/v1/pdf/security/remove',
      {
        url: uploadedFileUrl,
        password: password,
        name:
          path.basename(filePath, path.extname(filePath)) +
          '_unlocked' +
          path.extname(filePath),
      },
      {
        headers: {
          'x-api-key': PDFCO_API_KEY,
          'Content-Type': 'application/json',
        },
      }
    );

    if (unlockResponse.data && unlockResponse.data.url) {
      const unlockedPdfResponse = await axios.get(unlockResponse.data.url, {
        responseType: 'arraybuffer',
      });
      const unlockedPdfPath = filePath.replace(
        path.extname(filePath),
        '_unlocked' + path.extname(filePath)
      );
      fs.writeFileSync(unlockedPdfPath, unlockedPdfResponse.data);
      console.log('PDF unlocked successfully:', unlockedPdfPath);
      return unlockedPdfPath;
    } else {
      console.error('Error unlocking PDF:', unlockResponse.data);
      return null;
    }
  } catch (error) {
    console.error('Error in unlockPdf:', error.message);
    return null;
  }
}

async function isPdfEncrypted(filePath) {
  try {
    const pdfBuffer = await fs.promises.readFile(filePath);
    const pdfDoc = await PDFDocument.load(pdfBuffer, { ignoreEncryption: true });
    try {
      pdfDoc.getPage(0);
      return false;
    } catch {
      return true;
    }
  } catch (error) {
    console.error('Error checking PDF encryption:', error.message);
    return true;
  }
}

function detectCurrency(value) {
  const usdRegex = /\$|USD/;
  const ilsRegex = /₪|ILS|ש"ח/;
  if (usdRegex.test(value)) {
    return 'USD';
  } else if (ilsRegex.test(value)) {
    return 'ILS';
  } else {
    return 'ILS';
  }
}

const exchangeRateCache = {};

async function parseReceiptWithDocumentAI(filePath, serviceAccountAuth) {
  const { projectId, location, processorId } = DOCUMENT_AI_CONFIG;
  const url = `https://${location}-documentai.googleapis.com/v1/projects/${projectId}/locations/${location}/processors/${processorId}:process`;

  try {
    const fileContent = fs.readFileSync(filePath);
    const encodedFile = fileContent.toString('base64');

    const ext = path.extname(filePath).toLowerCase();
    let mimeType = 'application/pdf';
    if (ext === '.jpg' || ext === '.jpeg') mimeType = 'image/jpeg';
    if (ext === '.png') mimeType = 'image/png';
    if (ext === '.tiff' || ext === '.tif') mimeType = 'image/tiff';

    if (!serviceAccountAuth.credentials || !serviceAccountAuth.credentials.access_token) {
      await serviceAccountAuth.authorize();
    }

    const accessToken = serviceAccountAuth.credentials.access_token;

    const response = await axios.post(url, {
      rawDocument: {
        content: encodedFile,
        mimeType: mimeType,
      }
    }, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': 'application/json',
      },
    });

    if (response.status !== 200) {
      console.error('Error processing document:', response.data);
      return {};
    }

    const document = response.data.document;
    const entities = document.entities || [];
    const result = {
      FileName: path.basename(filePath),
      BusinessName: '',
      BusinessNumber: '',
      Date: '',
      InvoiceNumber: '',
      OriginalTotalUSD: '',
      PriceWithoutVat: '',
      VAT: '',
      TotalPrice: '',
      Currency: '',
    };

    let hasUSD = false;
    let invoiceDate = '';

    for (const entity of entities) {
      let value = '';
      let currencyCode = '';
      if (entity.normalizedValue && entity.normalizedValue.moneyValue) {
        value = entity.normalizedValue.moneyValue.amount;
        currencyCode = entity.normalizedValue.moneyValue.currencyCode || '';
      } else {
        value = entity.mentionText || '';
        currencyCode = detectCurrency(value);
      }

      if (currencyCode === 'USD') {
        hasUSD = true;
      }

      if (entity.type === 'Date' && !invoiceDate) {
        invoiceDate = value;
        result['Date'] = value;
      }
    }

    if (!invoiceDate) {
      invoiceDate = formatDate(new Date());
      result['Date'] = invoiceDate;
    } else {
      console.log('Attempting to parse invoice date:', invoiceDate);
      const dateFormats = [
        'yyyy-MM-dd',
        'MM/dd/yyyy',
        'dd/MM/yyyy',
        'dd-MM-yyyy',
        'MM-dd-yyyy',
        'yyyy/MM/dd',
        'dd.MM.yyyy',
        'yyyy.MM.dd',
        'MMMM dd, yyyy',
        'MMM dd, yyyy',
      ];
      let parsedDate;
      for (const df of dateFormats) {
        const pd = parse(invoiceDate, df, new Date());
        if (!isNaN(pd)) {
          console.log(`Date parsed with "${df}":`, pd);
          parsedDate = pd;
          break;
        }
      }
      if (!parsedDate) {
        console.error('Failed to parse invoice date:', invoiceDate);
        invoiceDate = formatDate(new Date());
        result['Date'] = invoiceDate;
      } else {
        invoiceDate = formatDate(parsedDate);
        result['Date'] = invoiceDate;
      }
    }

    let exchangeRate;
    if (hasUSD) {
      exchangeRate = exchangeRateCache[invoiceDate];
      if (!exchangeRate) {
        exchangeRate = await getUsdToIlsExchangeRate(invoiceDate);
        exchangeRateCache[invoiceDate] = exchangeRate;
      }
      console.log(`Using exchange rate: ${exchangeRate} for date: ${invoiceDate}`);
    }

    let originalUSD = 0;
    let convertedTotalPrice = 0;

    for (const entity of entities) {
      let value = '';
      if (entity.normalizedValue && entity.normalizedValue.moneyValue) {
        value = entity.normalizedValue.moneyValue.amount;
      } else {
        value = entity.mentionText || '';
      }

      switch (entity.type) {
        case 'Business-Name':
          result['BusinessName'] = value;
          break;
        case 'Business-Number':
          result['BusinessNumber'] = value;
          break;
        case 'Invoice-Number':
          result['InvoiceNumber'] = value;
          break;
        case 'Price-Without-Vat':
          const pwv = cleanAndParseAmount(value);
          if (hasUSD) {
            originalUSD += pwv;
            result['PriceWithoutVat'] = pwv * exchangeRate;
          } else {
            result['PriceWithoutVat'] = pwv;
          }
          break;
        case 'VAT':
          const vat = cleanAndParseAmount(value);
          if (hasUSD) {
            result['VAT'] = vat * exchangeRate;
          } else {
            result['VAT'] = vat;
          }
          break;
        case 'Total-Price':
          const tp = cleanAndParseAmount(value);
          if (hasUSD) {
            convertedTotalPrice += tp * exchangeRate;
          } else {
            result['TotalPrice'] = tp;
          }
          break;
      }
    }

    if (hasUSD) {
      result['OriginalTotalUSD'] = originalUSD;
      result['Currency'] = 'ILS';
      result['TotalPrice'] = convertedTotalPrice;
    } else {
      result['Currency'] = 'ILS';
    }

    console.log(`Total amount for ${result['FileName']}: ${result['TotalPrice']} (ILS)`);
    return result;
  } catch (error) {
    console.error('Error in parseReceiptWithDocumentAI:', error.message);
    return {};
  }
}

async function createExpenseExcel(expenses, folderPath, filePrefix, startDate, endDate, name) {
  const validStart = parse(startDate, 'yyyy-MM-dd', new Date());
  const validEnd = parse(endDate, 'yyyy-MM-dd', new Date());

  const startDateFormatted = format(validStart, 'dd-MM-yy');
  const endDateFormatted = format(validEnd, 'dd-MM-yy');

  let baseFileName = `${filePrefix}-${startDateFormatted}-to-${endDateFormatted}`;
  if (name && name.trim() !== '') {
    baseFileName += `-${name.replace(/\s+/g, '_')}`;
  }
  let fileName = `${baseFileName}.xlsx`;
  let fullPath = path.join(folderPath, fileName);

  let fileNumber = 1;
  while (fs.existsSync(fullPath)) {
    fileName = `${baseFileName} (${fileNumber}).xlsx`;
    fullPath = path.join(folderPath, fileName);
    fileNumber++;
  }

  const workbook = new Excel.Workbook();
  const worksheet = workbook.addWorksheet('Expenses', { views: [{ rightToLeft: true }] });

  worksheet.columns = [
    { header: 'שם הקובץ', key: 'FileName', width: 30 },
    { header: 'שם העסק', key: 'BusinessName', width: 25 },
    { header: 'מספר עסק', key: 'BusinessNumber', width: 20 },
    { header: 'תאריך', key: 'Date', width: 15 },
    { header: 'מספר חשבונית', key: 'InvoiceNumber', width: 20 },
    { header: 'סכום ללא מע"מ', key: 'PriceWithoutVat', width: 20 },
    { header: 'מע"מ', key: 'VAT', width: 15 },
    { header: 'סכום כולל', key: 'TotalPrice', width: 20 },
    { header: 'הומר מדולרים*', key: 'OriginalTotalUSD', width: 10 },
  ];

  worksheet.getRow(1).font = { bold: true, size: 12 };
  worksheet.getRow(1).alignment = { horizontal: 'center' };

  let totalWithoutVat = 0;
  let totalVAT = 0;
  let totalPrice = 0;
  let totalOriginalUSD = 0;

  expenses.forEach((expense) => {
    const originalUSD = expense['OriginalTotalUSD'] ? parseFloat(expense['OriginalTotalUSD']) : 0;
    const priceWithoutVatValue = expense['PriceWithoutVat'] ? parseFloat(expense['PriceWithoutVat']) : 0;
    const vatValue = expense['VAT'] ? parseFloat(expense['VAT']) : 0;
    const totalPriceValue = expense['TotalPrice'] ? parseFloat(expense['TotalPrice']) : 0;

    totalOriginalUSD += originalUSD;
    totalWithoutVat += priceWithoutVatValue;
    totalVAT += vatValue;
    totalPrice += totalPriceValue;

    worksheet.addRow({
      FileName: expense['FileName'],
      BusinessName: expense['BusinessName'],
      BusinessNumber: expense['BusinessNumber'],
      Date: expense['Date'],
      InvoiceNumber: expense['InvoiceNumber'],
      OriginalTotalUSD: originalUSD > 0 ? originalUSD : '',
      PriceWithoutVat: priceWithoutVatValue,
      VAT: vatValue,
      TotalPrice: totalPriceValue,
    });
  });

  const totalsRow = worksheet.addRow({
    FileName: 'Total',
    PriceWithoutVat: totalWithoutVat,
    VAT: totalVAT,
    TotalPrice: totalPrice,
  });

  totalsRow.font = { bold: true };
  totalsRow.getCell('PriceWithoutVat').numFmt = '#,##0.00 ₪';
  totalsRow.getCell('VAT').numFmt = '#,##0.00 ₪';
  totalsRow.getCell('TotalPrice').numFmt = '#,##0.00 ₪';
  totalsRow.alignment = { horizontal: 'center' };

  worksheet.getColumn('PriceWithoutVat').numFmt = '#,##0.00 ₪';
  worksheet.getColumn('VAT').numFmt = '#,##0.00 ₪';
  worksheet.getColumn('TotalPrice').numFmt = '#,##0.00 ₪';
  worksheet.getColumn('OriginalTotalUSD').numFmt = '$#,##0.00';

  worksheet.columns.forEach((column) => {
    column.alignment = { vertical: 'middle', horizontal: 'right' };
  });

  await workbook.xlsx.writeFile(fullPath);
  console.log('Expense summary Excel file created at:', fullPath);
  return fullPath;
}

async function createZipFile(files, outputFolder, zipFileName) {
  return new Promise((resolve, reject) => {
    const zipFilePath = path.join(outputFolder, zipFileName);
    const output = fs.createWriteStream(zipFilePath);
    const archive = archiver('zip', { zlib: { level: 9 } });

    output.on('close', () => {
      console.log(`ZIP file created at: ${zipFilePath} (${archive.pointer()} bytes)`);
      resolve(zipFilePath);
    });

    archive.on('error', (err) => {
      console.error('Error creating ZIP file:', err);
      reject(err);
    });

    archive.pipe(output);
    files.forEach((filePath) => {
      const fileName = path.basename(filePath);
      archive.file(filePath, { name: fileName });
    });
    archive.finalize();
  });
}

function isPdfFile(contentType, fileName) {
  const normalizedContentType = contentType.toLowerCase();
  const normalizedFileName = fileName.toLowerCase();

  if (
    normalizedContentType.includes('pdf') ||
    normalizedFileName.endsWith('.pdf')
  ) {
    return true;
  }
  return false;
}

// Download Gmail Attachments logic (Full, no omissions)
async function downloadGmailAttachments(auth, startDate, endDate, folderPath) {
  const gmail = google.gmail({ version: 'v1', auth });

  // Positive subject keywords
  const positiveSubjectKeywords = [
    'קבלה',
    'חשבונית',
    'חשבונית מס',
    'הקבלה',
    'החשבונית',
    'החשבונית החודשית',
    'אישור תשלום',
    'receipt',
    'invoice',
    'חשבון חודשי',
  ];

  const excludedSenders = [
    'חברת חשמל לישראל',
    'עיריית תל אביב-יפו',
    'ארנונה - עיריית תל-אביב-יפו'
  ];
  const senderExceptionKeywords = [
    'קבלה',
    'חשבונית',
    'חשבונית מס',
    'הקבלה',
  ];

  endDate.setHours(23,59,59,999);
  const queryEndDate = new Date(endDate.getTime() + 24*60*60*1000);

  const startDateQuery = format(startDate, 'yyyy/MM/dd');
  const endDateQuery = format(queryEndDate, 'yyyy/MM/dd');

  const query = `after:${startDateQuery} before:${endDateQuery} has:attachment`;
  console.log('Gmail query:', query);

  let nextPageToken = null;
  const allMessageIds = [];

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

  console.log('Total messages found:', allMessageIds.length);

  function getParts(payload) {
    let parts = [];
    if (payload.parts) {
      for (const part of payload.parts) {
        if (part.parts) {
          parts = parts.concat(getParts(part));
        } else {
          parts.push(part);
        }
      }
    } else {
      parts.push(payload);
    }
    return parts;
  }

  for (const messageData of allMessageIds) {
    const msg = await gmail.users.messages.get({
      userId: 'me',
      id: messageData.id,
      format: 'full',
    });

    const headers = msg.data.payload.headers;
    const fromHeader = headers.find(h => h.name.toLowerCase() === 'from');
    const subjectHeader = headers.find(h => h.name.toLowerCase() === 'subject');

    const sender = fromHeader ? fromHeader.value : '';
    const subject = subjectHeader ? subjectHeader.value : '';

    const senderName = sender.split('<')[0].trim().toLowerCase();
    const senderEmailMatch = sender.match(/<(.+?)>/);
    const senderEmail = senderEmailMatch ? senderEmailMatch[1].toLowerCase() : sender.toLowerCase();
    const lowerCaseSubject = subject.toLowerCase();

    let isExcludedSender = false;
    for (const es of excludedSenders) {
      if (senderName.includes(es) || senderEmail.includes(es)) {
        isExcludedSender = true;
        break;
      }
    }

    if (isExcludedSender) {
      let containsExceptionKeyword = false;
      for (const keyword of senderExceptionKeywords) {
        if (lowerCaseSubject.includes(keyword.toLowerCase())) {
          containsExceptionKeyword = true;
          break;
        }
      }
      if (!containsExceptionKeyword) {
        console.log('Skipping excluded sender thread without exception keyword:', subject);
        continue;
      }
    }

    let subjectMatchesPositiveKeywords = false;
    for (const kw of positiveSubjectKeywords) {
      if (lowerCaseSubject.includes(kw.toLowerCase())) {
        subjectMatchesPositiveKeywords = true;
        break;
      }
    }

    if (!msg.data.payload) {
      console.log('No attachments in message:', subject);
      continue;
    }

    const parts = getParts(msg.data.payload);
    const attachmentKeywords = ['receipt','חשבונית','קבלה'];

    let receiptFoundInThread = false;
    for (const part of parts) {
      if (part.filename && part.filename.length > 0) {
        const normalizedFileName = part.filename.toLowerCase();
        for (const kw of attachmentKeywords) {
          if (normalizedFileName.includes(kw.toLowerCase())) {
            receiptFoundInThread = true;
            break;
          }
        }
      }
      if (receiptFoundInThread) break;
    }

    for (const part of parts) {
      if (part.filename && part.filename.length > 0) {
        const attachmentId = part.body && part.body.attachmentId;
        if (!attachmentId) continue;
        const fileName = part.filename;
        const normalizedFileName = fileName.toLowerCase();
        const isPDF = isPdfFile(part.mimeType, fileName);

        if (isPDF) {
          if (subjectMatchesPositiveKeywords) {
            if (receiptFoundInThread) {
              let matchesKeyword = false;
              for (const kw of attachmentKeywords) {
                if (normalizedFileName.includes(kw.toLowerCase())) {
                  matchesKeyword = true;
                  break;
                }
              }
              if (matchesKeyword) {
                const attachment = await gmail.users.messages.attachments.get({
                  userId: 'me',
                  messageId: messageData.id,
                  id: attachmentId,
                });
                const data = attachment.data.data;
                const buffer = Buffer.from(data, 'base64');
                const filePath = path.join(folderPath, sanitize(fileName));
                fs.writeFileSync(filePath, buffer);
                console.log(`Saved PDF attachment: ${filePath}`);
              } else {
                console.log('Skipping non-receipt PDF:', fileName);
              }
            } else {
              // No receipt found by filename, just save the PDF
              const attachment = await gmail.users.messages.attachments.get({
                userId: 'me',
                messageId: messageData.id,
                id: attachmentId,
              });
              const data = attachment.data.data;
              const buffer = Buffer.from(data, 'base64');
              const filePath = path.join(folderPath, sanitize(fileName));
              fs.writeFileSync(filePath, buffer);
              console.log(`Saved PDF attachment: ${filePath}`);
            }
          } else {
            let matchesKeyword = false;
            for (const kw of attachmentKeywords) {
              if (normalizedFileName.includes(kw.toLowerCase())) {
                matchesKeyword = true;
                break;
              }
            }
            if (matchesKeyword) {
              const attachment = await gmail.users.messages.attachments.get({
                userId: 'me',
                messageId: messageData.id,
                id: attachmentId,
              });
              const data = attachment.data.data;
              const buffer = Buffer.from(data, 'base64');
              const filePath = path.join(folderPath, sanitize(fileName));
              fs.writeFileSync(filePath, buffer);
              console.log(`Saved PDF by filename keyword: ${filePath}`);
            } else {
              console.log('Skipping attachment not matching criteria:', fileName);
            }
          }
        } else {
          console.log('Skipping non-PDF attachment:', fileName);
        }
      }
    }
  }
  return folderPath;
}

// Nodemailer
const transporter = nodemailer.createTransport({
  host: process.env.SMTP_HOST,
  port: Number(process.env.SMTP_PORT) || 587,
  secure: false,
  auth: {
    user: process.env.SMTP_USER,
    pass: process.env.SMTP_PASS,
  },
});

async function sendResultByEmail(toEmail, excelPath, zipPath) {
  if (!toEmail) return;
  const mailOptions = {
    from: '"Expenses Processor" <no-reply@example.com>',
    to: toEmail,
    subject: 'Your Processed Files',
    text: 'Your files have been processed. Please find attachments.',
    attachments: [
      { filename: path.basename(excelPath), path: excelPath },
      { filename: path.basename(zipPath), path: zipPath },
    ],
  };

  try {
    await transporter.sendMail(mailOptions);
    console.log(`Email sent to ${toEmail}`);
  } catch (error) {
    console.error('Error sending email:', error);
  }
}

const INPUT_FOLDER = path.join(__dirname, 'input_files');
fs.ensureDirSync(INPUT_FOLDER);

let lastCreatedResult = null;

// Upload tasks
let taskQueue = [];
let isProcessing = false;

// Gmail tasks
let gmailTaskQueue = [];
let isGmailProcessing = false;

const progressEmitters = new Map();

async function processFile(filePath, serviceAccountAuth, password) {
  const ext = path.extname(filePath).toLowerCase();
  const isPDF = ext === '.pdf';
  let processedFilePath = filePath;

  if (isPDF) {
    const encrypted = await isPdfEncrypted(filePath);
    if (encrypted) {
      console.log('PDF encrypted, unlocking:', filePath);
      const unlockedPath = await unlockPdf(filePath, password || PASSWORD_PROTECTED_PDF_PASSWORD);
      if (unlockedPath) {
        fs.copyFileSync(unlockedPath, filePath);
        fs.unlinkSync(unlockedPath);
        processedFilePath = filePath;
      } else {
        console.log('Failed to unlock PDF:', filePath);
        return null;
      }
    } else {
      console.log('PDF not encrypted:', filePath);
    }
  } else {
    console.log('Image file:', filePath);
  }

  const expenseData = await parseReceiptWithDocumentAI(processedFilePath, serviceAccountAuth);
  return expenseData;
}

async function processNextTask() {
  if (taskQueue.length === 0) {
    isProcessing = false;
    return;
  }

  isProcessing = true;
  const job = taskQueue.shift();
  const { sessionId, userFolder, files, pdfPassword, email } = job;
  const progressEmitter = new EventEmitter();
  progressEmitters.set(sessionId, progressEmitter);

  try {
    progressEmitter.emit('progress', [{ status: 'Processing started.', progress: 0 }]);
    const serviceAccountAuth = authenticateServiceAccount();
    await serviceAccountAuth.authorize();

    if (!files || files.length === 0) {
      progressEmitter.emit('progress', [{ status: 'No files uploaded.', progress: 100 }]);
      return;
    }

    const progressData = files.map(filePath => ({
      fileName: path.basename(filePath),
      status: 'Pending',
      progress: 0,
    }));

    const emitProgress = () => {
      progressEmitter.emit('progress', progressData);
    };

    emitProgress();

    const expenses = [];
    for (let i = 0; i < files.length; i++) {
      const filePath = files[i];
      progressData[i].status = 'Processing';
      progressData[i].progress = 25;
      emitProgress();

      const expenseData = await processFile(filePath, serviceAccountAuth, pdfPassword);
      if (expenseData) {
        expenses.push(expenseData);
        progressData[i].status = 'Completed';
        progressData[i].progress = 100;
        progressData[i].businessName = expenseData.BusinessName || 'N/A';
        progressData[i].date = expenseData.Date || 'N/A';
        progressData[i].totalPrice = expenseData.TotalPrice
          ? parseFloat(expenseData.TotalPrice).toFixed(2)
          : 'N/A';
      } else {
        progressData[i].status = 'Failed';
        progressData[i].progress = 100;
      }
      emitProgress();
    }

    if (expenses.length > 0) {
      progressEmitter.emit('progress', [
        ...progressData,
        { status: 'Creating Excel file...', progress: 80 },
      ]);
      const startDate = formatDate(new Date());
      const endDate = formatDate(new Date());
      const customPrefix = 'סיכום הוצאות';
      const excelPath = await createExpenseExcel(expenses, userFolder, customPrefix, startDate, endDate, '');

      const zipFileName = `processed_files_${Date.now()}.zip`;
      const zipFilePath = await createZipFile(files, userFolder, zipFileName);

      if (email && email.trim() !== '') {
        await sendResultByEmail(email, excelPath, zipFilePath);
      }

      lastCreatedResult = {
        excelPath,
        zipPath: zipFilePath,
        timestamp: new Date(),
        emailSent: !!email,
      };

      progressEmitter.emit('progress', [
        ...progressData,
        {
          status: 'Processing complete.',
          progress: 100,
          downloadLinks: [
            { label: 'Download Excel', url: `/download/${encodeURIComponent(path.basename(excelPath))}` },
            { label: 'Download ZIP', url: `/download/${encodeURIComponent(path.basename(zipFilePath))}` },
          ],
        },
      ]);
    } else {
      progressEmitter.emit('progress', [
        ...progressData,
        { status: 'No expenses were extracted.', progress: 100 },
      ]);
    }

    setTimeout(() => {
      fs.remove(userFolder).catch(console.error);
    }, 3600000);
  } catch (error) {
    console.error('Processing Error:', error.message);
    progressEmitter.emit('progress', [
      { status: `Processing Error: ${error.message}`, progress: 100 },
    ]);
  } finally {
    progressEmitters.delete(sessionId);
    isProcessing = false;
    processNextTask();
  }
}

// GMAIL PROCESSING LOGIC - NO OMISSIONS
// Similar to /upload, but we download from Gmail first
let gmailProgressEmitters = new Map();

async function processNextGmailTask() {
  if (gmailTaskQueue.length === 0) {
    isGmailProcessing = false;
    return;
  }

  isGmailProcessing = true;
  const job = gmailTaskQueue.shift();
  const { sessionId, userFolder, startDate, endDate, pdfPassword, email, additionalFiles } = job;

  const progressEmitter = new EventEmitter();
  gmailProgressEmitters.set(sessionId, progressEmitter);

  try {
    progressEmitter.emit('progress', [{ status: 'Processing started.', progress: 0 }]);

    const auth = new google.auth.OAuth2(
      process.env.GMAIL_CLIENT_ID,
      process.env.GMAIL_CLIENT_SECRET,
      process.env.GMAIL_REDIRECT_URI
    );
    auth.setCredentials(job.tokens);

    const customPrefix = 'סיכום הוצאות';
    console.log('Processing Gmail attachments from', startDate, 'to', endDate);

    const startDateObj = new Date(startDate);
    const endDateObj = new Date(endDate);
    endDateObj.setHours(23,59,59,999);

    progressEmitter.emit('progress', [{ status: 'Downloading Gmail attachments...', progress: 10 }]);
    await downloadGmailAttachments(auth, startDateObj, endDateObj, userFolder);

    // Get all files after download
    let files = fs.readdirSync(userFolder).map(f => path.join(userFolder, f));

    // Add additional files if any
    additionalFiles.forEach(file => {
      if (!files.includes(file)) {
        files.push(file);
      }
    });

    console.log('Files found after Gmail and additional:', files);

    if (files.length === 0) {
      progressEmitter.emit('progress', [{ status: 'No files found to process.', progress: 100 }]);
      return;
    }

    const progressData = files.map(filePath => ({
      fileName: path.basename(filePath),
      status: 'Pending',
      progress: 0,
    }));
    const emitProgress = () => {
      progressEmitter.emit('progress', progressData);
    };
    emitProgress();

    const serviceAccountAuth = authenticateServiceAccount();
    await serviceAccountAuth.authorize();
    const expenses = [];

    for (let i = 0; i < files.length; i++) {
      const filePath = files[i];
      progressData[i].status = 'Processing';
      progressData[i].progress = 25;
      emitProgress();

      const expenseData = await processFile(filePath, serviceAccountAuth, pdfPassword);
      if (expenseData) {
        expenses.push(expenseData);
        progressData[i].status = 'Completed';
        progressData[i].progress = 100;
        progressData[i].businessName = expenseData.BusinessName || 'N/A';
        progressData[i].date = expenseData.Date || 'N/A';
        progressData[i].totalPrice = expenseData.TotalPrice
          ? parseFloat(expenseData.TotalPrice).toFixed(2)
          : 'N/A';
      } else {
        progressData[i].status = 'Failed';
        progressData[i].progress = 100;
      }
      emitProgress();
    }

    progressEmitter.emit('progress', [{ status: 'Creating Excel file...', progress: 80 }, ...progressData]);
    const excelPath = await createExpenseExcel(expenses, userFolder, customPrefix, startDate, endDate, '');
    console.log('Excel file:', excelPath);

    const zipFileName = `processed_files_${Date.now()}.zip`;
    const zipFilePath = await createZipFile(files, userFolder, zipFileName);

    if (email && email.trim() !== '') {
      await sendResultByEmail(email, excelPath, zipFilePath);
    }

    lastCreatedResult = {
      excelPath,
      zipPath: zipFilePath,
      timestamp: new Date(),
      emailSent: !!email,
    };

    progressEmitter.emit('progress', [
      ...progressData,
      {
        status: 'Processing complete. Files ready.',
        progress: 100,
        downloadLinks: [
          { label: 'Download Excel', url: `/download/${encodeURIComponent(path.basename(excelPath))}` },
          { label: 'Download ZIP', url: `/download/${encodeURIComponent(path.basename(zipFilePath))}` },
        ],
      },
    ]);

    setTimeout(() => {
      fs.remove(userFolder).catch(console.error);
    }, 3600000);

  } catch (error) {
    console.error('Error processing Gmail attachments:', error);
    const pe = gmailProgressEmitters.get(sessionId);
    if (pe) {
      pe.emit('progress', [{ status: `Error: ${error.message}`, progress: 100 }]);
    }
  } finally {
    gmailProgressEmitters.delete(sessionId);
    isGmailProcessing = false;
    processNextGmailTask();
  }
}

// Routes

app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Upload routes
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    if (!req.session.sessionId) {
      req.session.sessionId = uuidv4();
    }
    const userFolder = path.join(INPUT_FOLDER, req.session.sessionId);
    fs.ensureDirSync(userFolder);
    cb(null, userFolder);
  },
  filename: (req, file, cb) => {
    const sanitized = sanitize(file.originalname) || 'unnamed_attachment';
    cb(null, sanitized);
  },
});
const upload = multer({
  storage: storage,
  fileFilter: (req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    const allowed = ['.pdf', '.jpg', '.jpeg', '.png', '.tiff', '.tif'];
    cb(null, allowed.includes(ext));
  },
  limits: { fileSize: 50 * 1024 * 1024 },
}).array('files', 100);

app.post('/upload', upload, (req, res) => {
  if (!req.session.sessionId) {
    req.session.sessionId = uuidv4();
  }
  const sessionId = req.session.sessionId;
  const userFolder = path.join(INPUT_FOLDER, sessionId);
  fs.ensureDirSync(userFolder);

  const pdfPassword = req.body.idNumber || '';
  const email = req.body.email || '';
  const files = req.files ? req.files.map(f => f.path) : [];

  if (files.length === 0) {
    return res.status(400).json({ message: 'No files uploaded.' });
  }

  const job = { sessionId, userFolder, files, pdfPassword, email, createdAt: Date.now() };
  taskQueue.push(job);

  res.json({ 
    sessionId,
    message: 'Your job has been queued. You can check progress at /upload-progress'
  });

  if (!isProcessing) {
    processNextTask();
  }
});

app.get('/upload-progress', (req, res) => {
  const sessionId = req.session.sessionId;
  if (!sessionId) {
    return res.status(404).end();
  }
  const existing = progressEmitters.get(sessionId);
  if (!existing) {
    return res.status(404).end();
  }

  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.flushHeaders();

  const onProgress = data => {
    res.write(`data: ${JSON.stringify(data)}\n\n`);
  };

  existing.on('progress', onProgress);

  const heartbeat = setInterval(() => {
    res.write(':\n\n');
  }, 30000);

  req.on('close', () => {
    existing.removeListener('progress', onProgress);
    clearInterval(heartbeat);
  });
});

app.get('/last-results', (req, res) => {
  if (!lastCreatedResult) {
    return res.send('No recent results available.');
  }
  const excelName = encodeURIComponent(path.basename(lastCreatedResult.excelPath));
  const zipName = encodeURIComponent(path.basename(lastCreatedResult.zipPath));

  res.send(`
    <h1>Last Created Results</h1>
    <p>Created at: ${lastCreatedResult.timestamp}</p>
    <p><a href="/download/${excelName}">Download Excel</a></p>
    <p><a href="/download/${zipName}">Download ZIP</a></p>
  `);
});

app.get('/download/:filename', (req, res) => {
  const { filename } = req.params;
  const decoded = decodeURIComponent(filename);
  const sessionId = req.session.sessionId;
  if (!sessionId) {
    return res.status(403).send('Access denied.');
  }

  const userFolder = path.join(INPUT_FOLDER, sessionId);
  function findFile(dir) {
    if (!fs.existsSync(dir)) return null;
    const files = fs.readdirSync(dir);
    for (const file of files) {
      const fp = path.join(dir, file);
      const stat = fs.statSync(fp);
      if (stat.isDirectory()) {
        const found = findFile(fp);
        if (found) return found;
      } else if (file === decoded) {
        return fp;
      }
    }
    return null;
  }

  const filePath = findFile(userFolder);
  if (filePath && fs.existsSync(filePath)) {
    let contentType = 'application/octet-stream';
    const ext = path.extname(filePath).toLowerCase();
    if (ext === '.xlsx') contentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    if (ext === '.zip') contentType = 'application/zip';
    res.setHeader('Content-Type', contentType);
    return res.download(filePath, decoded, (err) => {
      if (err) console.error('Download Error:', err.message);
    });
  } else {
    return res.status(404).send('File not found. It may have expired.');
  }
});

// Gmail OAuth
app.get('/is-authenticated', (req, res) => {
  res.json({ authenticated: !!req.session.tokens });
});

app.get('/start-gmail-auth', (req, res) => {
  const oAuth2Client = new google.auth.OAuth2(
    process.env.GMAIL_CLIENT_ID,
    process.env.GMAIL_CLIENT_SECRET,
    process.env.GMAIL_REDIRECT_URI
  );
  const authUrl = oAuth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: ['https://www.googleapis.com/auth/gmail.readonly'],
    prompt: 'consent',
  });
  res.redirect(authUrl);
});

app.get('/oauth2callback', async (req, res) => {
  const code = req.query.code;
  if (!code) return res.status(400).send('No code provided');
  const oAuth2Client = new google.auth.OAuth2(
    process.env.GMAIL_CLIENT_ID,
    process.env.GMAIL_CLIENT_SECRET,
    process.env.GMAIL_REDIRECT_URI
  );
  try {
    const { tokens } = await oAuth2Client.getToken(code);
    oAuth2Client.setCredentials(tokens);
    req.session.tokens = tokens;
    res.redirect('/gmail');
  } catch (error) {
    console.error('Error retrieving access token', error);
    res.status(500).send('Authentication failed');
  }
});

app.get('/gmail', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'gmail.html'));
});

// Handle Gmail Processing (no omissions)
const gmailAdditionalUpload = multer({
  storage: multer.diskStorage({
    destination: (req, file, cb) => {
      // We'll create a unique folder for Gmail too
      const sessionId = req.session.sessionId || uuidv4();
      req.session.sessionId = sessionId;
      const userFolder = path.join(INPUT_FOLDER, sessionId);
      fs.ensureDirSync(userFolder);
      cb(null, userFolder);
    },
    filename: (req, file, cb) => {
      const sanitized = sanitize(file.originalname) || 'unnamed_attachment';
      cb(null, sanitized);
    },
  }),
  fileFilter: (req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    const allowed = ['.pdf','.jpg','.jpeg','.png','.tiff','.tif'];
    cb(null, allowed.includes(ext));
  },
  limits: { fileSize: 50*1024*1024 },
}).array('additionalFiles', 100);

app.post('/process-gmail', gmailAdditionalUpload, (req, res) => {
  if (!req.session.tokens) {
    return res.status(403).send('Not authenticated with Gmail. Please login first.');
  }

  if (!req.session.sessionId) {
    req.session.sessionId = uuidv4();
  }
  const sessionId = req.session.sessionId;
  const userFolder = path.join(INPUT_FOLDER, sessionId);
  fs.ensureDirSync(userFolder);

  const startDate = req.body.startDate;
  const endDate = req.body.endDate;
  const pdfPassword = req.body.idNumber || '';
  const email = req.body.email || '';
  const additionalFiles = req.files ? req.files.map(f => f.path) : [];

  if (!startDate || !endDate) {
    return res.status(400).send('Please provide startDate and endDate.');
  }

  const job = {
    sessionId,
    userFolder,
    startDate,
    endDate,
    pdfPassword,
    email,
    additionalFiles,
    tokens: req.session.tokens,
    createdAt: Date.now()
  };

  gmailTaskQueue.push(job);

  res.json({
    sessionId,
    message: 'Your Gmail processing job has been queued. Check /gmail-progress'
  });

  if (!isGmailProcessing) {
    processNextGmailTask();
  }
});

app.get('/gmail-progress', (req, res) => {
  const sessionId = req.session.sessionId;
  if (!sessionId) {
    return res.status(404).end();
  }

  const existing = gmailProgressEmitters.get(sessionId);
  if (!existing) {
    return res.status(404).end();
  }

  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.flushHeaders();

  const onProgress = data => {
    res.write(`data: ${JSON.stringify(data)}\n\n`);
  };

  existing.on('progress', onProgress);

  const heartbeat = setInterval(() => {
    res.write(':\n\n');
  }, 30000);

  req.on('close', () => {
    existing.removeListener('progress', onProgress);
    clearInterval(heartbeat);
  });
});

// Start server
const PORT = process.env.PORT || 8080;
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
