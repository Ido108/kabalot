// server.js

require('dotenv').config(); // Load environment variables

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

const app = express();

// Middleware
app.use(cors()); // Allow CORS for all origins
app.use(express.static('public'));
app.use(express.urlencoded({ extended: true })); // To parse form data
app.use(express.json()); // To parse JSON bodies
app.set('view engine', 'ejs');

// Configure session middleware
app.use(
  session({
    secret: process.env.SESSION_SECRET || 'your_session_secret', // Replace with your own secret
    resave: false,
    saveUninitialized: true,
  })
);

// Configuration Constants
const date = new Date();
const folderName = `${date.getFullYear()}-${(date.getMonth() + 1)
  .toString()
  .padStart(2, '0')}-${date
  .getDate()
  .toString()
  .padStart(2, '0')}_${date
  .getHours()
  .toString()
  .padStart(2, '0')}-${date
  .getMinutes()
  .toString()
  .padStart(2, '0')}-${date.getSeconds().toString().padStart(2, '0')}`;

// Create folder path based on the current date and time
const INPUT_FOLDER =
  process.env.INPUT_FOLDER ||
  path.join(__dirname, 'input_files', folderName);

// Ensure the directory exists
if (!fs.existsSync(INPUT_FOLDER)) {
  fs.mkdirSync(INPUT_FOLDER, { recursive: true });
}

console.log(`Input folder: ${INPUT_FOLDER}`);

const PDFCO_API_KEY = process.env.PDFCO_API_KEY;
const PASSWORD_PROTECTED_PDF_PASSWORD =
  process.env.PASSWORD_PROTECTED_PDF_PASSWORD || 'your-default-password';

// Google Document AI Configuration (Unchanged)
const DOCUMENT_AI_CONFIG = {
  projectId: 'eighth-block-311611', // Replace with your GCP project ID
  location: 'us', // Processor location
  processorId: '78ac8067f0c37ec6', // Your actual processor ID
};

// Use base64 encoded service account credentials from environment variable
const SERVICE_ACCOUNT_BASE64 = process.env.SERVICE_ACCOUNT_BASE64;

// Initialize Google OAuth2 Client for Document AI
const SCOPES = ['https://www.googleapis.com/auth/cloud-platform'];

/**
 * Authenticate with Service Account for Document AI
 */
function authenticateServiceAccount() {
  if (!SERVICE_ACCOUNT_BASE64) {
    throw new Error('SERVICE_ACCOUNT_BASE64 is not set in environment variables.');
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

/**
 * Format Date as YYYY-MM-DD
 */
function formatDate(date) {
  return format(date, 'yyyy-MM-dd'); // Using date-fns for consistent formatting
}

/**
 * Clean and Parse Amount
 */
function cleanAndParseAmount(amountStr) {
  if (!amountStr) return 0;
  // Remove any non-numeric characters except for '.' and '-'
  amountStr = amountStr.replace(/[^0-9.\-]+/g, '');
  const parsedAmount = parseFloat(amountStr);
  return isNaN(parsedAmount) ? 0 : parsedAmount;
}

/**
 * Get USD to ILS Exchange Rate for a Specific Date
 * @param {string} date - Date in 'YYYY-MM-DD' format
 */
async function getUsdToIlsExchangeRate(date) {
  // Use exchange rate API that supports historical rates
  // For example, using exchangerate.host API
  const url = `https://api.exchangerate.host/convert?from=USD&to=ILS&date=${date}&amount=1`;
  try {
    const response = await axios.get(url);
    if (response.data && response.data.result) {
      console.log(`Exchange rate on ${date}: ${response.data.result}`);
      return response.data.result;
    } else {
      console.log(`Error retrieving exchange rate for ${date}. Defaulting to 3.7404.`);
      return 3.7404;
    }
  } catch (error) {
    console.error(`Error fetching exchange rate for ${date}:`, error.message);
    return 3.7404; // Default rate
  }
}

/**
 * Unlock Password-Protected PDF using PDF.co
 * @param {string} filePath - Path to the locked PDF
 * @param {string} password - Password to unlock the PDF
 * @returns {string|null} - Path to the unlocked PDF or null if failed
 */
async function unlockPdf(
  filePath,
  password 
) {
  if (!PDFCO_API_KEY) {
    throw new Error('PDFCO_API_KEY is not set in environment variables.');
  }
  password = password || '';
  try {
    // Step 1: Upload the PDF
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
      console.error('Error uploading PDF:', JSON.stringify(uploadResponse.data));
      return null;
    }

    const uploadedFileUrl = uploadResponse.data.url;

    // Step 2: Unlock the PDF
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
      // Download the unlocked PDF
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
      console.error('Error unlocking PDF:', JSON.stringify(unlockResponse.data));
      return null;
    }
  } catch (error) {
    if (error.response) {
      // Server responded with a status code outside 2xx
      console.error(`Error in unlockPdf: Status ${error.response.status}`);
      console.error('Response data:', JSON.stringify(error.response.data));
    } else if (error.request) {
      // No response received
      console.error('Error in unlockPdf: No response received from PDF.co API');
      console.error(error.request);
    } else {
      // Error setting up the request
      console.error('Error in unlockPdf:', error.message);
    }
    return null;
  }
}

/**
 * Check if a PDF is encrypted (password-protected)
 * @param {string} filePath - Path to the PDF file
 * @returns {Promise<boolean>} - Returns true if encrypted, false otherwise
 */
async function isPdfEncrypted(filePath) {
  try {
    const existingPdfBytes = fs.readFileSync(filePath);
    await PDFDocument.load(existingPdfBytes, { ignoreEncryption: true });
    return false; // Not encrypted
  } catch (error) {
    // PDF-lib throws an error if the PDF is encrypted
    if (
      error.message.includes('Cannot parse PDF') ||
      error.message.includes('encrypted')
    ) {
      return true; // Encrypted
    }
    // Re-throw if it's a different error
    throw error;
  }
}

/**
 * Detect Currency using regex
 * @param {string} value - Text containing the amount
 * @returns {string} - Detected currency code ('USD' or 'ILS')
 */
function detectCurrency(value) {
  const usdRegex = /\$|USD/; // Updated to detect both $ and USD
  const ilsRegex = /₪|ILS|ש"ח/;

  if (usdRegex.test(value)) {
    return 'USD';
  } else if (ilsRegex.test(value)) {
    return 'ILS';
  } else {
    return 'ILS'; // Default to ILS
  }
}

// Exchange rate cache to store rates by date
const exchangeRateCache = {};

/**
 * Parse Receipt with Google Document AI
 * @param {string} filePath - Path to the file (PDF or Image)
 * @param {object} serviceAccountAuth - Authenticated service account
 * @returns {object} - Extracted expense data
 */
async function parseReceiptWithDocumentAI(filePath, serviceAccountAuth) {
  const { projectId, location, processorId } = DOCUMENT_AI_CONFIG;
  const url = `https://${location}-documentai.googleapis.com/v1/projects/${projectId}/locations/${location}/processors/${processorId}:process`;

  try {
    const fileContent = fs.readFileSync(filePath);
    const encodedFile = fileContent.toString('base64');

    // Determine MIME type based on file extension
    const ext = path.extname(filePath).toLowerCase();
    let mimeType = 'application/pdf'; // Default

    if (ext === '.jpg' || ext === '.jpeg') {
      mimeType = 'image/jpeg';
    } else if (ext === '.png') {
      mimeType = 'image/png';
    } else if (ext === '.tiff' || ext === '.tif') {
      mimeType = 'image/tiff';
    }

    const payload = {
      rawDocument: {
        content: encodedFile,
        mimeType: mimeType,
      },
    };

    // Ensure the service account is authorized
    if (
      !serviceAccountAuth.credentials ||
      !serviceAccountAuth.credentials.access_token
    ) {
      await serviceAccountAuth.authorize();
    }

    const accessToken = serviceAccountAuth.credentials.access_token;

    const response = await axios.post(url, payload, {
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
    let originalUSD = 0;
    let convertedTotalPrice = 0;
    let invoiceDate = ''; // To store the date extracted

    // First pass to detect if USD exists and extract the date
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

      // Extract the date
      if (entity.type === 'Date') {
        result['Date'] = value;
        invoiceDate = value;
      }
    }

    // If date is not extracted, use today's date
    if (!invoiceDate) {
      invoiceDate = formatDate(new Date());
      result['Date'] = invoiceDate;
    } else {
      // Parse and format the date to 'YYYY-MM-DD'
      try {
        // Try parsing common date formats
        const parsedDate = parse(invoiceDate, 'yyyy-MM-dd', new Date());
        if (isNaN(parsedDate)) {
          // Try another format
          const altParsedDate = parse(invoiceDate, 'MM/dd/yyyy', new Date());
          if (isNaN(altParsedDate)) {
            // If parsing fails, use today's date
            invoiceDate = formatDate(new Date());
            result['Date'] = invoiceDate;
          } else {
            invoiceDate = formatDate(altParsedDate);
            result['Date'] = invoiceDate;
          }
        } else {
          invoiceDate = formatDate(parsedDate);
          result['Date'] = invoiceDate;
        }
      } catch (dateParseError) {
        console.error('Error parsing invoice date:', dateParseError.message);
        invoiceDate = formatDate(new Date());
        result['Date'] = invoiceDate;
      }
    }

    // Fetch exchange rate for the invoice date, use cache if available
    let exchangeRate = exchangeRateCache[invoiceDate];
    if (!exchangeRate) {
      exchangeRate = await getUsdToIlsExchangeRate(invoiceDate);
      exchangeRateCache[invoiceDate] = exchangeRate;
    }

    // Process entities and convert amounts
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

      switch (entity.type) {
        case 'Business-Name':
          result['BusinessName'] = value;
          break;
        case 'Business-Number':
          result['BusinessNumber'] = value;
          break;
        case 'Date':
          // Already handled
          break;
        case 'Invoice-Number':
          result['InvoiceNumber'] = value;
          break;
        case 'Price-Without-Vat':
          const priceWithoutVat = cleanAndParseAmount(value);
          if (hasUSD) {
            originalUSD += priceWithoutVat;
            result['PriceWithoutVat'] = priceWithoutVat * exchangeRate;
          } else {
            result['PriceWithoutVat'] = priceWithoutVat;
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
          const totalPrice = cleanAndParseAmount(value);
          if (hasUSD) {
            convertedTotalPrice += totalPrice * exchangeRate;
            result['TotalPrice'] = totalPrice * exchangeRate;
          } else {
            result['TotalPrice'] = totalPrice;
          }
          break;
        default:
          // Ignore other entities
          break;
      }
    }

    if (hasUSD) {
      result['OriginalTotalUSD'] = originalUSD;
      result['Currency'] = 'ILS'; // After conversion
      result['TotalPrice'] = convertedTotalPrice;
    } else {
      result['Currency'] = 'ILS';
    }

    console.log(
      `Total amount for ${result['FileName']}: ${result['TotalPrice']} (${result['Currency']})`
    );
    return result;
  } catch (error) {
    console.error('Error in parseReceiptWithDocumentAI:', error.message);
    return {};
  }
}

/**
 * Create Expense Spreadsheet (Excel) with formatting
 * @param {Array} expenses - Array of expense objects
 * @param {string} folderPath - Path to the folder where Excel file will be saved
 */
async function createExpenseExcel(expenses, folderPath) {
  const excelPath = path.join(folderPath, 'סיכום הוצאות.xlsx'); // 'Expense Summary' in Hebrew

  // Create a new workbook and add a worksheet
  const workbook = new Excel.Workbook();
  const worksheet = workbook.addWorksheet('Expenses');

  // Define columns with headers and keys
  worksheet.columns = [
    { header: 'שם הקובץ', key: 'FileName', width: 30 }, // File Name
    { header: 'שם העסק', key: 'BusinessName', width: 25 }, // Business Name
    { header: 'מספר עסק', key: 'BusinessNumber', width: 20 }, // Business Number
    { header: 'תאריך', key: 'Date', width: 15 }, // Date
    { header: 'מספר חשבונית', key: 'InvoiceNumber', width: 20 }, // Invoice Number
    {
      header: 'סכום מקורי בדולרים',
      key: 'OriginalTotalUSD',
      width: 20,
    }, // Original Amount (USD)
    { header: 'סכום ללא מע"מ', key: 'PriceWithoutVat', width: 20 }, // Price Without VAT
    { header: 'מע"מ', key: 'VAT', width: 15 }, // VAT
    { header: 'סכום כולל', key: 'TotalPrice', width: 20 }, // Total Price
    { header: 'מטבע', key: 'Currency', width: 10 }, // Currency
  ];

  // Apply styling to header row
  worksheet.getRow(1).font = { bold: true, size: 12 };
  worksheet.getRow(1).alignment = { horizontal: 'center' };

  // Prepare data rows
  let totalWithoutVat = 0;
  let totalVAT = 0;
  let totalPrice = 0;
  let totalOriginalUSD = 0;

  expenses.forEach((expense) => {
    // Sum totals for each column
    const originalUSD = expense['OriginalTotalUSD']
      ? parseFloat(expense['OriginalTotalUSD'])
      : 0;
    const priceWithoutVatValue = expense['PriceWithoutVat']
      ? parseFloat(expense['PriceWithoutVat'])
      : 0;
    const vatValue = expense['VAT'] ? parseFloat(expense['VAT']) : 0;
    const totalPriceValue = expense['TotalPrice']
      ? parseFloat(expense['TotalPrice'])
      : 0;

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
      Currency: expense['Currency'],
    });
  });

  // Add totals row
  const totalsRow = worksheet.addRow({
    FileName: 'Total',
    OriginalTotalUSD: totalOriginalUSD > 0 ? totalOriginalUSD : '',
    PriceWithoutVat: totalWithoutVat,
    VAT: totalVAT,
    TotalPrice: totalPrice,
  });

  // Apply styling to totals row
  totalsRow.font = { bold: true };
  totalsRow.getCell('PriceWithoutVat').numFmt = '#,##0.00 ₪';
  totalsRow.getCell('VAT').numFmt = '#,##0.00 ₪';
  totalsRow.getCell('TotalPrice').numFmt = '#,##0.00 ₪';
  totalsRow.alignment = { horizontal: 'center' };

  // Format currency columns
  worksheet.getColumn('PriceWithoutVat').numFmt = '#,##0.00 ₪';
  worksheet.getColumn('VAT').numFmt = '#,##0.00 ₪';
  worksheet.getColumn('TotalPrice').numFmt = '#,##0.00 ₪';
  worksheet.getColumn('OriginalTotalUSD').numFmt = '$#,##0.00';

  // Adjust alignment
  worksheet.columns.forEach((column) => {
    column.alignment = { vertical: 'middle', horizontal: 'center' };
  });

  // Save the workbook to file
  await workbook.xlsx.writeFile(excelPath);
  console.log('Expense summary Excel file created at:', excelPath);
}

/**
 * Process Files (PDFs and Images)
 * @param {Array} files - Array of file paths
 * @param {string} folderPath - Path to save the Excel file
 * @param {string} pdfPassword - Password for protected PDFs
 * @returns {Array} - Array of extracted expense data
 */
async function processFiles(files, folderPath, pdfPassword) {
  // Authenticate with Service Account for Document AI
  const serviceAccountAuth = authenticateServiceAccount();
  await serviceAccountAuth.authorize(); // Ensure the client is authorized

  const expenses = [];

  for (const filePath of files) {
    console.log(`Processing file: ${filePath}`);

    const ext = path.extname(filePath).toLowerCase();
    const isPDF = ext === '.pdf';

    let processedFilePath = filePath;

    if (isPDF) {
      // Check if the PDF is encrypted
      const isEncrypted = await isPdfEncrypted(filePath);
      if (isEncrypted) {
        console.log('PDF is encrypted. Attempting to unlock:', filePath);
        // Attempt to unlock PDF using the provided password
        const unlockedPath = await unlockPdf(filePath, pdfPassword);
        if (unlockedPath) {
          // Overwrite the original file with the unlocked PDF
          fs.copyFileSync(unlockedPath, filePath);
          fs.unlinkSync(unlockedPath); // Remove the temporary unlocked file
          console.log('PDF unlocked and overwritten:', filePath);
          processedFilePath = filePath; // Continue processing the unlocked file
        } else {
          console.log('Skipping locked PDF:', filePath);
          continue; // Skip processing if PDF is locked and couldn't be unlocked
        }
      } else {
        console.log(
          'PDF is not encrypted. Proceeding without unlocking:',
          filePath
        );
        // No action needed since we're processing in place
      }
    } else {
      console.log('File is an image. Proceeding to process:', filePath);
      // No encryption handling needed for images
    }

    // Parse the receipt with Document AI using the file path
    const expenseData = await parseReceiptWithDocumentAI(
      processedFilePath,
      serviceAccountAuth
    );
    if (Object.keys(expenseData).length > 0) {
      expenses.push(expenseData);
    }
  }

  return expenses;
}

/**
 * Download Gmail Attachments
 * Adapted for dynamic user authentication
 */
async function downloadGmailAttachments(auth, startDate, endDate) {
  const gmail = google.gmail({ version: 'v1', auth });

  // Create folder to save attachments
  const folderName = `קבלות ${formatDate(startDate)} עד ${formatDate(endDate)}`;
  const folderPath = path.join(INPUT_FOLDER, folderName);
  fs.ensureDirSync(folderPath);

  // Prepare date queries
  const afterDate = format(startDate, 'yyyy/MM/dd');
  const beforeDate = format(endDate, 'yyyy/MM/dd');

  const query = `after:${afterDate} before:${beforeDate}`;

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

    // Exclusion logic
    let excludeThread = false;
    let keywordFound = false;

    for (const excludedSender of excludedSenders) {
      if (sender.includes(excludedSender)) {
        excludeThread = true;
        for (const keyword of keywords) {
          if (subject.includes(keyword)) {
            keywordFound = true;
            break; // No need to check further if keyword is found
          }
        }
        break; // No need to check other senders
      }
    }

    // Decide whether to skip the thread
    if (excludeThread && !keywordFound) {
      // Skip this thread
      console.log('Skipping message from excluded sender:', sender);
      continue;
    }

    let receiptFoundInThread = false; // Flag to indicate if a receipt PDF has been found in this thread

    // First pass: check if there's a receipt PDF in the message
    if (msg.data.payload.parts) {
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
    }

    // Second pass: process the attachments based on whether receipt was found
    if (msg.data.payload.parts) {
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
              }
            } else {
              // If no receipt is found, collect the PDF as usual
              const filePath = path.join(folderPath, sanitize(fileName));
              fs.writeFileSync(filePath, buffer);
            }
          }
        }
      }
    }
  }

  return folderPath;
}

/**
 * Check if a file is a PDF
 */
function isPdfFile(contentType, fileName) {
  const normalizedContentType = contentType.toLowerCase();
  const normalizedFileName = fileName.toLowerCase();

  if (
    normalizedContentType === 'application/pdf' ||
    normalizedContentType === 'application/x-pdf' ||
    normalizedContentType === 'application/acrobat' ||
    normalizedContentType === 'applications/vnd.pdf' ||
    normalizedContentType === 'text/pdf' ||
    normalizedContentType === 'text/x-pdf' ||
    normalizedContentType.includes('pdf')
  ) {
    return true;
  } else if (normalizedFileName.endsWith('.pdf')) {
    return true;
  } else {
    return false;
  }
}

// Ensure input folder exists
fs.ensureDirSync(INPUT_FOLDER);

// Set up Multer for handling file uploads
const storage = multer.diskStorage({
  destination: function (req, file, cb) {
    cb(null, INPUT_FOLDER);
  },
  filename: function (req, file, cb) {
    // Sanitize filename
    const sanitized = sanitize(file.originalname) || 'unnamed_attachment';
    cb(null, sanitized);
  },
});

const upload = multer({
  storage: storage,
  fileFilter: function (req, file, cb) {
    // Accept PDF and common image files
    const ext = path.extname(file.originalname).toLowerCase();
    const allowedExtensions = ['.pdf', '.jpg', '.jpeg', '.png', '.tiff', '.tif'];
    if (allowedExtensions.includes(ext)) {
      cb(null, true);
    } else {
      console.warn(`Skipped unsupported file type: ${file.originalname}`);
      cb(null, false); // Skip the file without throwing an error
    }
  },
  limits: { fileSize: 50 * 1024 * 1024 }, // 50MB limit per file
}).array('files', 100); // Max 100 files

// Routes

// Home Route - Serve the upload form
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Handle File Upload and Processing
app.post('/upload', (req, res) => {
  upload(req, res, async function (err) {
    if (err instanceof multer.MulterError) {
      // A Multer error occurred when uploading.
      console.error('Multer Error:', err.message);
      return res.status(500).send(`Multer Error: ${err.message}`);
    } else if (err) {
      // An unknown error occurred when uploading.
      console.error('Upload Error:', err.message);
      return res.status(500).send(`Upload Error: ${err.message}`);
    }

    // Everything went fine.
    if (!req.files || req.files.length === 0) {
      return res
        .status(400)
        .send(
          'No supported files were uploaded. Please upload PDF or image files.'
        );
    }

    console.log(`Received ${req.files.length} file(s). Starting processing...`);

    // Set output folder path to INPUT_FOLDER
    const outputFolder = INPUT_FOLDER;
    const pdfPassword = req.body.pdfPassword || '';
    try {
      // Get file paths
      const files = req.files.map((file) => file.path);
      const expenses = await processFiles(files, OUTPUT_FOLDER, pdfPassword);


      // Create Expense Excel File
      if (expenses.length > 0) {
        await createExpenseExcel(expenses, outputFolder);
        console.log('Expense summary Excel file created.');

        // Provide a download link to the Excel file
        const excelFileName = encodeURIComponent('סיכום הוצאות.xlsx');
        const csvUrl = `/download/${excelFileName}`;
        res.render('result', {
          success: true,
          csvUrl: csvUrl,
          message: 'Files Uploaded and Processed Successfully!',
        });
      } else {
        res.render('result', {
          success: false,
          message: 'No expenses were extracted from the uploaded files.',
        });
      }
    } catch (processingError) {
      console.error('Processing Error:', processingError.message);
      res.status(500).send(`Processing Error: ${processingError.message}`);
    }
  });
});

// Download Route - Serve the generated Excel file
app.get('/download/:filename', (req, res) => {
  const { filename } = req.params;
  const decodedFilename = decodeURIComponent(filename);
  const filePath = path.join(INPUT_FOLDER, decodedFilename);

  if (fs.existsSync(filePath)) {
    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    );
    res.download(filePath, decodedFilename, (err) => {
      if (err) {
        console.error('Download Error:', err.message);
        res.status(500).send('Error downloading the file.');
      }
    });
  } else {
    res.status(404).send('File not found.');
  }
});

// Gmail Authentication and Processing Routes

/**
 * Authenticate with OAuth2 for Gmail API
 */
function authenticateGmail(req, res, next) {
  const oAuth2Client = new google.auth.OAuth2(
    process.env.GMAIL_CLIENT_ID,
    process.env.GMAIL_CLIENT_SECRET,
    process.env.GMAIL_REDIRECT_URI
  );

  // Check if we have an access token stored in session
  if (req.session.tokens) {
    oAuth2Client.setCredentials(req.session.tokens);

    // Set up event listener to update tokens if refreshed
    oAuth2Client.on('tokens', (tokens) => {
      if (tokens.refresh_token) {
        req.session.tokens.refresh_token = tokens.refresh_token;
      }
      req.session.tokens.access_token = tokens.access_token;
    });

    req.oAuth2Client = oAuth2Client;
    next();
  } else {
    // No tokens, redirect to Gmail authentication
    res.redirect('/');
  }
}

// Route to initiate Gmail authentication
app.post('/start-gmail-auth', (req, res) => {
  const { email } = req.body;
  if (!email) {
    return res.status(400).send('Please enter your Gmail address.');
  }

  // Store the email in session
  req.session.email = email;

  const oAuth2Client = new google.auth.OAuth2(
    process.env.GMAIL_CLIENT_ID,
    process.env.GMAIL_CLIENT_SECRET,
    process.env.GMAIL_REDIRECT_URI
  );

  // Generate an OAuth URL with login_hint to prefill the email
  const authUrl = oAuth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: ['https://www.googleapis.com/auth/gmail.readonly'],
    prompt: 'consent',
    login_hint: email,
  });

  res.redirect(authUrl);
});

// OAuth2 callback route
app.get('/oauth2callback', async (req, res) => {
  const code = req.query.code;
  if (!code) {
    return res.status(400).send('No code provided');
  }

  const oAuth2Client = new google.auth.OAuth2(
    process.env.GMAIL_CLIENT_ID,
    process.env.GMAIL_CLIENT_SECRET,
    process.env.GMAIL_REDIRECT_URI
  );

  try {
    const { tokens } = await oAuth2Client.getToken(code);
    oAuth2Client.setCredentials(tokens);
    // Store tokens in session
    req.session.tokens = tokens;
    req.oAuth2Client = oAuth2Client;
    // Redirect to the date input form
    res.redirect('/process-gmail');
  } catch (error) {
    console.error('Error retrieving access token', error);
    res.status(500).send('Authentication failed');
  }
});

// Route to display date input form
app.get('/process-gmail', authenticateGmail, (req, res) => {
  res.render('process-gmail', { email: req.session.email });
});

// Process Gmail attachments
app.post('/process-gmail', authenticateGmail, async (req, res) => {
  try {
    const auth = req.oAuth2Client;

    // Get start and end dates from the form
    const { startDate: startDateStr, endDate: endDateStr } = req.body;

    if (!startDateStr || !endDateStr) {
      return res.status(400).send('Please provide both start date and end date.');
    }

    const startDate = new Date(startDateStr);
    const endDate = new Date(endDateStr);
    endDate.setHours(23, 59, 59, 999);

    // Download attachments
    const attachmentsFolder = await downloadGmailAttachments(auth, startDate, endDate);

    // Get all files in the attachments folder
    const files = fs.readdirSync(attachmentsFolder).map((file) =>
      path.join(attachmentsFolder, file)
    );

    console.log(`Downloaded ${files.length} attachment(s). Starting processing...`);

    if (files.length === 0) {
      return res
        .status(400)
        .send('No attachments were downloaded from Gmail within the specified date range.');
    }

    // Process the downloaded files
    const expenses = await processFiles(files, attachmentsFolder);

    // Create Expense Excel File
    if (expenses.length > 0) {
      await createExpenseExcel(expenses, attachmentsFolder);
      console.log('Expense summary Excel file created.');

      // Provide a download link to the Excel file
      const excelFileName = encodeURIComponent('סיכום הוצאות.xlsx');
      const csvUrl = `/download/${excelFileName}`;
      res.render('result', {
        success: true,
        csvUrl: csvUrl,
        message: 'Gmail attachments processed successfully!',
      });
    } else {
      res.render('result', {
        success: false,
        message: 'No expenses were extracted from the Gmail attachments.',
      });
    }
  } catch (error) {
    console.error('Error processing Gmail attachments:', error.message);
    res.status(500).send(`Error: ${error.message}`);
  }
});

// User Logout Route
app.get('/logout', (req, res) => {
  req.session.destroy();
  res.redirect('/');
});

// Start the Server
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server is running on port ${PORT}`);
});
