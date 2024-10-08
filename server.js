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
const RedisStore = require('connect-redis')(session);
const redisClient = require('redis').createClient();
const session = require('express-session');
const events = require('events'); // For progress events
const archiver = require('archiver'); // For creating ZIP files

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
    store: new RedisStore({ client: redisClient }),
    secret: process.env.SESSION_SECRET || 'your_session_secret',
    resave: false,
    saveUninitialized: true,
    cookie: { maxAge: 3600000 },
  })
);
const taskQueue = [];
let isProcessing = false;
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

// Google Document AI Configuration
const DOCUMENT_AI_CONFIG = {
  projectId: process.env.DOCUMENT_AI_PROJECT_ID || 'your-project-id', // Replace with your GCP project ID
  location: process.env.DOCUMENT_AI_LOCATION || 'us', // Processor location
  processorId: process.env.DOCUMENT_AI_PROCESSOR_ID || 'your-processor-id', // Your actual processor ID
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
 * Schedule deletion of a folder after a specified delay
 * @param {string} folderPath - Path to the folder to delete
 * @param {number} delayMs - Delay in milliseconds before deletion
 */
function scheduleFolderDeletion(folderPath, delayMs) {
  setTimeout(() => {
    fs.remove(folderPath)
      .then(() => {
        console.log(`Deleted folder: ${folderPath}`);
      })
      .catch((err) => {
        console.error(`Error deleting folder ${folderPath}:`, err);
      });
  }, delayMs);
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
  // Implement your method to get the exchange rate
  // For the purpose of this example, we'll return a fixed rate
  const exchangeRate = 3.5; // Replace with actual exchange rate fetching logic
  return exchangeRate;
}

/**
 * Unlock Password-Protected PDF using PDF.co
 * @param {string} filePath - Path to the locked PDF
 * @param {string} password - Password to unlock the PDF
 * @returns {string|null} - Path to the unlocked PDF or null if failed
 */
async function unlockPdf(filePath, password) {
  if (!PDFCO_API_KEY) {
    throw new Error('PDFCO_API_KEY is not set in environment variables.');
  }

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
async function processNextTask() {
  if (taskQueue.length === 0) {
    isProcessing = false;
    return;
  }

  isProcessing = true;
  const task = taskQueue.shift(); // Get the next task
  const {
    sessionId,
    userFolder,
    files,
    name,
    idNumber,
    progressEmitter,
    req,
  } = task;

  try {
    // Notify user that processing has started
    progressEmitter.emit('progress', [
      { status: 'Processing started.', progress: 0 },
    ]);

    if (!files || files.length === 0) {
      progressEmitter.emit('progress', [{ status: 'No files uploaded.', progress: 100 }]);
      return;
    }

    console.log(`Processing ${files.length} file(s) for session ${sessionId}...`);

    const progressData = files.map((filePath) => ({
      fileName: path.basename(filePath),
      status: 'Pending',
      progress: 0,
    }));

    // Function to emit progress updates
    const emitProgress = () => {
      progressEmitter.emit('progress', progressData);
    };

    emitProgress(); // Initial emit

    const expenses = [];
    const serviceAccountAuth = authenticateServiceAccount();
    await serviceAccountAuth.authorize();

    for (let i = 0; i < files.length; i++) {
      const filePath = files[i];

      // Update status to 'Processing'
      progressData[i].status = 'Processing';
      progressData[i].progress = 25;
      emitProgress();

      // Process the file
      const expenseData = await processFile(filePath, serviceAccountAuth, idNumber);

      if (expenseData) {
        expenses.push(expenseData);

        // Update progress data with expense information
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

      // Round progress percentages
      progressData[i].progress = Math.round(progressData[i].progress);

      emitProgress();
    }

    // Create Expense Excel File
    if (expenses.length > 0) {
      const startDate = formatDate(new Date());
      const endDate = formatDate(new Date());

      const excelPath = await createExpenseExcel(
        expenses,
        userFolder,
        'סיכום הוצאות', // Default file prefix
        startDate,
        endDate,
        name // Optional name
      );
      console.log('Expense summary Excel file created.');

      // Create ZIP file of processed files
      const zipFileName = `processed_files_${Date.now()}.zip`;
      const zipFilePath = await createZipFile(files, userFolder, zipFileName);
      console.log('Processed files ZIP created.');

      // Provide download links
      const excelFileName = encodeURIComponent(path.basename(excelPath));
      const excelUrl = `/download/${excelFileName}`;

      const zipFileNameEncoded = encodeURIComponent(path.basename(zipFilePath));
      const zipUrl = `/download/${zipFileNameEncoded}`;

      progressEmitter.emit('progress', [
        ...progressData,
        {
          status: 'Processing complete. Download the files below. Files will be available for 1 hour.',
          progress: 100,
          downloadLinks: [
            { label: 'הורד קובץ אקסל', url: excelUrl },
            { label: 'הורד קבצים מעובדים (ZIP)', url: zipUrl },
          ],
        },
      ]);
    } else {
      progressEmitter.emit('progress', [
        ...progressData,
        { status: 'No expenses extracted.', progress: 100 },
      ]);
    }
  } catch (processingError) {
    console.error('Processing Error:', processingError.message);
    progressEmitter.emit('progress', [
      { status: `Processing Error: ${processingError.message}`, progress: 100 },
    ]);
  } finally {
    // Schedule deletion after 1 hour (3600000 milliseconds)
    scheduleFolderDeletion(userFolder, 3600000); // 1 hour delay

    // Clean up progressEmitter
    req.session.progressEmitter = null;

    // Process the next task
    isProcessing = false;
    processNextTask();
  }
}

/**
 * Check if a PDF is encrypted (password-protected)
 * @param {string} filePath - Path to the PDF file
 * @returns {Promise<boolean>} - Returns true if encrypted, false otherwise
 */
async function isPdfEncrypted(filePath) {
  try {
    const pdfBuffer = await fs.promises.readFile(filePath);
    const pdfDoc = await PDFDocument.load(pdfBuffer, { ignoreEncryption: true });
    
    // Check if the PDF is encrypted
    if (pdfDoc.isEncrypted) {
      return true;
    }

    // Additional check: try to access a page
    try {
      pdfDoc.getPage(0);
      return false; // If we can access a page, it's not encrypted
    } catch (error) {
      // If we can't access a page, it might be encrypted
      return true;
    }
  } catch (error) {
    console.error('Error checking PDF encryption:', error.message);
    // If there's an error, assume it's encrypted to be safe
    return true;
  }
}

/**
 * Detect Currency using regex
 * @param {string} value - Text containing the amount
 * @returns {string} - Detected currency code ('USD' or 'ILS')
 */
function detectCurrency(value) {
  const usdRegex = /\$|USD/; // Detect both $ and USD
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

function getAttachmentsFolder(req) {
  const startDate = req.body.startDate;
  const endDate = req.body.endDate;

  const startDateObj = new Date(startDate);
  const endDateObj = new Date(endDate);
  endDateObj.setHours(23, 59, 59, 999); // Set to end of day

  const folderName = `קבלות ${formatDate(startDateObj)} עד ${formatDate(endDateObj)}`;
  const folderPath = path.join(INPUT_FOLDER, folderName);
  return folderPath;
}
const additionalStorage = multer.diskStorage({
  destination: function (req, file, cb) {
    const attachmentsFolder = getAttachmentsFolder(req);
    fs.ensureDirSync(attachmentsFolder); // Ensure the directory exists
    cb(null, attachmentsFolder);
  },
  filename: function (req, file, cb) {
    // Sanitize filename
    const sanitized = sanitize(file.originalname) || 'unnamed_attachment';
    cb(null, sanitized);
  },
});

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
    let invoiceDate = '';

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

      if (entity.type === 'Date' && !invoiceDate) {
        result['Date'] = value;
        invoiceDate = value;
      }

      if (hasUSD && invoiceDate) break;
    }

    // If date is not extracted, use today's date
// Existing code...
    if (!invoiceDate) {
      invoiceDate = formatDate(new Date());
      result['Date'] = invoiceDate;
    } else {
      // Parse and format the date to 'YYYY-MM-DD'
      try {
        console.log('Attempting to parse invoice date:', invoiceDate);

        // Array of date formats to try
        const dateFormats = [
          'yyyy-MM-dd',
          'MM/dd/yyyy',
          'dd/MM/yyyy',
          'dd-MM-yyyy',
          'MM-dd-yyyy',
          'yyyy/MM/dd',
          'dd.MM.yyyy',
          'yyyy.MM.dd',
          'MMMM dd, yyyy', // e.g., October 04, 2024
          'MMM dd, yyyy',   // e.g., Oct 04, 2024
        ];

        let parsedDate;
        for (const dateFormat of dateFormats) {
          parsedDate = parse(invoiceDate, dateFormat, new Date());
          if (!isNaN(parsedDate)) {
            console.log(`Date parsed successfully with format "${dateFormat}":`, parsedDate);
            break;
          }
        }
        if (isNaN(parsedDate)) {
          // If parsing fails, use today's date
          console.error('Failed to parse invoice date:', invoiceDate);
          invoiceDate = formatDate(new Date());
          result['Date'] = invoiceDate;
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

    let exchangeRate;
    if (hasUSD) {
      // Only fetch exchange rate if USD is detected
      exchangeRate = exchangeRateCache[invoiceDate];
      if (!exchangeRate) {
        exchangeRate = await getUsdToIlsExchangeRate(invoiceDate);
        exchangeRateCache[invoiceDate] = exchangeRate;
      }
      console.log(`Using exchange rate: ${exchangeRate} for date: ${invoiceDate}`);
    }

    // Process entities and convert amounts
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
            convertedTotalPrice = totalPrice * exchangeRate;
          } else {
            result['TotalPrice'] = totalPrice;
          }
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
 * @param {string} filePrefix - Prefix for the Excel filename
 * @param {string} startDate - Start date for the period
 * @param {string} endDate - End date for the period
 * @param {string} [name] - Optional name to include in the filename
 */
async function createExpenseExcel(expenses, folderPath, filePrefix, startDate, endDate, name) {
  // Ensure valid dates
  const validStartDate = parse(startDate, 'yyyy-MM-dd', new Date());
  const validEndDate = parse(endDate, 'yyyy-MM-dd', new Date());

  // Format dates
  const startDateFormatted = format(validStartDate, 'dd-MM-yy');
  const endDateFormatted = format(validEndDate, 'dd-MM-yy');

  // Create base filename
  let baseFileName = `${filePrefix}-${startDateFormatted}-to-${endDateFormatted}`;
  if (name && name.trim() !== '') {
    baseFileName += `-${name.replace(/\s+/g, '_')}`;
  }
  let fileName = `${baseFileName}.xlsx`;
  let fullPath = path.join(folderPath, fileName);

  // Check for existing files and add numbering if necessary
  let fileNumber = 1;
  while (fs.existsSync(fullPath)) {
    fileName = `${baseFileName} (${fileNumber}).xlsx`;
    fullPath = path.join(folderPath, fileName);
    fileNumber++;
  }

  // Create a new workbook and add a worksheet
  const workbook = new Excel.Workbook();
  const worksheet = workbook.addWorksheet('Expenses', {
    views: [{ rightToLeft: true }]
  });
  // Define columns with headers and keys
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
    });
  });

  // Add totals row
  const totalsRow = worksheet.addRow({
    FileName: 'Total',
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
    column.alignment = { vertical: 'middle', horizontal: 'right' };
  });

  // Save the workbook to file
  try {
    await workbook.xlsx.writeFile(fullPath);
    console.log('Expense summary Excel file created at:', fullPath);
    return fullPath;
  } catch (error) {
    console.error('Error saving Excel file:', error);
    throw new Error('Failed to save Excel file');
  }
}

/**
 * Create a ZIP file containing the processed files
 * @param {Array} files - Array of file paths to include in the ZIP
 * @param {string} outputFolder - Path to the folder where the ZIP file will be saved
 * @param {string} zipFileName - Desired name of the ZIP file
 * @returns {Promise<string>} - Path to the created ZIP file
 */
async function createZipFile(files, outputFolder, zipFileName) {
  return new Promise((resolve, reject) => {
    const zipFilePath = path.join(outputFolder, zipFileName);
    const output = fs.createWriteStream(zipFilePath);
    const archive = archiver('zip', {
      zlib: { level: 9 }, // Sets the compression level
    });

    output.on('close', () => {
      console.log(`ZIP file created at: ${zipFilePath} (${archive.pointer()} total bytes)`);
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

/**
 * Process a single file
 */
async function processFiles(files, folderPath, filePrefix, startDate, endDate, name) {
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
      try {
        // Check if the PDF is encrypted
        const isEncrypted = await isPdfEncrypted(filePath);
        if (isEncrypted) {
          console.log('PDF is encrypted. Attempting to unlock:', filePath);
          // Attempt to unlock PDF
          const unlockedPath = await unlockPdf(filePath);
          if (unlockedPath) {
            // Overwrite the original file with the unlocked PDF
            fs.copyFileSync(unlockedPath, filePath);
            fs.unlinkSync(unlockedPath); // Remove the temporary unlocked file
            console.log('PDF unlocked and overwritten:', filePath);
            processedFilePath = filePath; // Continue processing the unlocked file
          } else {
            console.log('Failed to unlock PDF:', filePath);
            continue; // Skip processing if PDF is locked and couldn't be unlocked
          }
        } else {
          console.log('PDF is not encrypted. Proceeding without unlocking:', filePath);
        }
      } catch (error) {
        console.error('Error processing PDF:', filePath, error);
        continue; // Skip this file if there's an error
      }
    } else {
      console.log('File is an image. Proceeding to process:', filePath);
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

  if (expenses.length > 0) {
    const excelPath = await createExpenseExcel(expenses, folderPath, filePrefix, startDate, endDate, name);
    return { expenses, excelPath };
  }

  return { expenses, excelPath: null };
}

async function processFile(filePath, serviceAccountAuth, idNumber) {
  const ext = path.extname(filePath).toLowerCase();
  const isPDF = ext === '.pdf';
  let processedFilePath = filePath;

  if (isPDF) {
    try {
      const isEncrypted = await isPdfEncrypted(filePath);
      if (isEncrypted) {
        console.log('PDF is encrypted. Attempting to unlock:', filePath);
        const password = idNumber || PASSWORD_PROTECTED_PDF_PASSWORD;
        const unlockedPath = await unlockPdf(filePath, password);
        if (unlockedPath) {
          fs.copyFileSync(unlockedPath, filePath);
          fs.unlinkSync(unlockedPath);
          console.log('PDF unlocked and overwritten:', filePath);
          processedFilePath = filePath;
        } else {
          console.log('Failed to unlock PDF:', filePath);
          return null;
        }
      } else {
        console.log('PDF is not encrypted. Proceeding without unlocking:', filePath);
      }
    } catch (error) {
      console.error('Error processing PDF:', filePath, error);
      return null;
    }
  } else {
    console.log('File is an image. Proceeding to process:', filePath);
  }

  const expenseData = await parseReceiptWithDocumentAI(processedFilePath, serviceAccountAuth);
  return expenseData;
}



// Ensure input folder exists
fs.ensureDirSync(INPUT_FOLDER);

// Set up Multer for handling file uploads
const storage = multer.diskStorage({
  destination: function (req, file, cb) {
    // Generate or retrieve the session ID
    if (!req.session.sessionId) {
      req.session.sessionId = uuidv4();
    }
    const userFolder = path.join(INPUT_FOLDER, req.session.sessionId);
    fs.ensureDirSync(userFolder); // Ensure the directory exists
    cb(null, userFolder);
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
// Progress Emitters
const uploadProgressEmitters = {};
const gmailProgressEmitters = {};

// Routes

// Home Route - Serve the upload form
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Handle File Upload and Processing with Progress Logging
// Handle File Upload and Processing with Progress Logging
app.post('/upload', upload, async (req, res) => {
  // Generate or retrieve the session ID
  if (!req.session.sessionId) {
    req.session.sessionId = uuidv4();
  }
  const sessionId = req.session.sessionId;
  const userFolder = path.join(INPUT_FOLDER, sessionId);
  fs.ensureDirSync(userFolder);

  const progressEmitter = new events.EventEmitter();
  req.session.progressEmitter = progressEmitter;

  // Send a response to the client indicating the session ID
  res.json({ sessionId });

  // Create a task object
  const task = {
    sessionId,
    userFolder,
    files: req.files ? req.files.map((file) => file.path) : [],
    name: req.body.name || '',
    idNumber: req.body.idNumber || '',
    progressEmitter,
    req,
  };

  // Add task to the queue
  taskQueue.push(task);

  // Inform the user if they are in a queue
  if (isProcessing) {
    const queuePosition = taskQueue.length;
    progressEmitter.emit('progress', [
      { status: `Your task is in a queue at position ${queuePosition}. It will start processing shortly.`, progress: 0 },
    ]);
  } else {
    // Start processing immediately if no other tasks are running
    processNextTask();
  }
});



// Endpoint for Upload Progress

app.get('/upload-progress', (req, res) => {
  const progressEmitter = req.session.progressEmitter;

  if (!progressEmitter) {
    res.status(404).end();
    return;
  }

  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.flushHeaders();

  const onProgress = (data) => {
    res.write(`data: ${JSON.stringify(data)}\n\n`);
  };

  progressEmitter.on('progress', onProgress);

  req.on('close', () => {
    progressEmitter.removeListener('progress', onProgress);
  });
});


// Download Route - Serve the generated files
// Download Route - Serve the generated files
app.get('/download/:filename', (req, res) => {
  const { filename } = req.params;
  const decodedFilename = decodeURIComponent(filename);
  const sessionId = req.session.sessionId;

  if (!sessionId) {
    res.status(403).send('Access denied.');
    return;
  }

  const userFolder = path.join(INPUT_FOLDER, sessionId);

  // Search for the file in userFolder and its subfolders
  const findFile = (dir) => {
    if (!fs.existsSync(dir)) {
      return null;
    }
    const files = fs.readdirSync(dir);
    for (const file of files) {
      const filePath = path.join(dir, file);
      const stat = fs.statSync(filePath);
      if (stat.isDirectory()) {
        const found = findFile(filePath);
        if (found) return found;
      } else if (file === decodedFilename) {
        return filePath;
      }
    }
    return null;
  };

  const filePath = findFile(userFolder);

  if (filePath && fs.existsSync(filePath)) {
    const ext = path.extname(filePath).toLowerCase();
    let contentType;
    if (ext === '.xlsx') {
      contentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    } else if (ext === '.zip') {
      contentType = 'application/zip';
    } else {
      contentType = 'application/octet-stream';
    }
    res.setHeader('Content-Type', contentType);
    res.download(filePath, decodedFilename, (err) => {
      if (err) {
        console.error('Download Error:', err.message);
        res.status(500).send('Error downloading the file.');
      }
    });
  } else {
    res.status(404).send('File not found. The file may have expired and been deleted.');
  }
});

async function processNextGmailTask() {
  if (gmailTaskQueue.length === 0) {
    isGmailProcessing = false;
    return;
  }

  isGmailProcessing = true;
  const task = gmailTaskQueue.shift(); // Get the next task
  const {
    sessionId,
    userFolder,
    startDate,
    endDate,
    idNumber,
    progressEmitter,
    req,
  } = task;

  try {
    // Notify user that processing has started
    progressEmitter.emit('progress', [
      { status: 'Processing started.', progress: 0 },
    ]);

    const auth = req.oAuth2Client;
    const customPrefix = 'סיכום הוצאות';

    console.log('Processing Gmail attachments from', startDate, 'to', endDate);

    const startDateObj = new Date(startDate);
    const endDateObj = new Date(endDate);
    endDateObj.setHours(23, 59, 59, 999);

    progressEmitter.emit('progress', [{ status: 'Downloading Gmail attachments...', progress: 10 }]);

    // Download Gmail attachments into userFolder
    await downloadGmailAttachments(auth, startDateObj, endDateObj, userFolder);
    console.log('Attachments downloaded to:', userFolder);

    // Get all files from userFolder
    let files = fs.readdirSync(userFolder).map((file) => path.join(userFolder, file));
    console.log('Files found:', files);

    if (files.length === 0) {
      progressEmitter.emit('progress', [{ status: 'No attachments found.', progress: 100 }]);
      return;
    }

    const progressData = files.map((filePath) => ({
      fileName: path.basename(filePath),
      status: 'Pending',
      progress: 0,
    }));

    const expenses = [];
    const serviceAccountAuth = authenticateServiceAccount();
    await serviceAccountAuth.authorize();

    for (let i = 0; i < files.length; i++) {
      const filePath = files[i];

      // Update status to 'Processing'
      progressData[i].status = 'Processing';
      progressData[i].progress = 25;
      progressEmitter.emit('progress', progressData);

      // Process the file
      const expenseData = await processFile(filePath, serviceAccountAuth, idNumber);

      if (expenseData) {
        expenses.push(expenseData);

        // Update progress data with expense information
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

      // Round progress percentages
      progressData[i].progress = Math.round(progressData[i].progress);

      progressEmitter.emit('progress', progressData);
    }

    progressEmitter.emit('progress', [{ status: 'Creating Excel file...', progress: 80 }]);

    const excelPath = await createExpenseExcel(
      expenses,
      userFolder,
      customPrefix,
      startDate,
      endDate
    );
    console.log('Expenses extracted:', expenses.length);
    console.log('Excel file created at:', excelPath);

    // Create ZIP file of processed files
    const zipFileName = `processed_files_${Date.now()}.zip`;
    const zipFilePath = await createZipFile(files, userFolder, zipFileName);
    console.log('Processed files ZIP created.');

    // Provide download links
    const excelFileName = encodeURIComponent(path.basename(excelPath));
    const excelUrl = `/download/${excelFileName}`;

    const zipFileNameEncoded = encodeURIComponent(path.basename(zipFilePath));
    const zipUrl = `/download/${zipFileNameEncoded}`;

    progressEmitter.emit('progress', [
      ...progressData,
      {
        status: 'Processing complete. Download the files below. Files will be available for 1 hour.',
        progress: 100,
        downloadLinks: [
          { label: 'הורד קובץ אקסל', url: excelUrl },
          { label: 'הורד קבצים מעובדים (ZIP)', url: zipUrl },
        ],
      },
    ]);

    // Schedule deletion after 1 hour (3600000 milliseconds)
    scheduleFolderDeletion(userFolder, 3600000); // 1 hour delay

    // Clean up progressEmitter
    req.session.gmailProgressEmitter = null;
  } catch (error) {
    console.error('Error processing Gmail attachments:', error);
    progressEmitter.emit('progress', [{ status: `Error: ${error.message}`, progress: 100 }]);
  } finally {
    // Process the next task
    isGmailProcessing = false;
    processNextGmailTask();
  }
}


// Gmail Authentication and Processing Routes

/**
 * Authenticate with OAuth2 for Gmail API
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
function authenticateGmail(req, res, next) {
  const oAuth2Client = new google.auth.OAuth2(
    process.env.GMAIL_CLIENT_ID,
    process.env.GMAIL_CLIENT_SECRET,
    process.env.GMAIL_REDIRECT_URI
  );

  if (req.session.tokens) {
    oAuth2Client.setCredentials(req.session.tokens);

    oAuth2Client.on('tokens', (tokens) => {
      if (tokens.refresh_token) {
        req.session.tokens.refresh_token = tokens.refresh_token;
      }
      req.session.tokens.access_token = tokens.access_token;
    });

    req.oAuth2Client = oAuth2Client;
    next();
  } else {
    res.redirect('/start-gmail-auth');
  }
}


app.get('/is-authenticated', (req, res) => {
  if (req.session.tokens) {
    res.json({ authenticated: true });
  } else {
    res.json({ authenticated: false });
  }
});



// Route to initiate Gmail authentication
// Route to initiate Gmail authentication
app.get('/start-gmail-auth', (req, res) => {
  const oAuth2Client = new google.auth.OAuth2(
    process.env.GMAIL_CLIENT_ID,
    process.env.GMAIL_CLIENT_SECRET,
    process.env.GMAIL_REDIRECT_URI
  );

  // Generate an OAuth URL without login_hint
  const authUrl = oAuth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: ['https://www.googleapis.com/auth/gmail.readonly'],
    prompt: 'consent',
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
    // Redirect to the Gmail processing page
    res.redirect('/gmail');
  } catch (error) {
    console.error('Error retrieving access token', error);
    res.status(500).send('Authentication failed');
  }
});


// Route to display Gmail processing form
app.get('/gmail-form', authenticateGmail, (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'gmail-form.html'));
});


// Route to display Gmail processing page
// Route to display Gmail processing page
app.get('/gmail', authenticateGmail, (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'gmail.html'));
});


// Handle Gmail Processing with Progress Logging
const additionalUpload = multer({
  storage: additionalStorage, // Use the new storage configuration
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
}).array('additionalFiles', 100);


const gmailTaskQueue = [];
let isGmailProcessing = false;

app.post('/process-gmail', authenticateGmail, additionalUpload, (req, res) => {
  // Generate or retrieve the session ID
  const uniqueId = uuidv4();
  progressEmitters[uniqueId] = progressEmitter;
  res.json({ uniqueId });
  if (!req.session.sessionId) {
    req.session.sessionId = uuidv4();
  }
  const sessionId = req.session.sessionId;
  const userFolder = path.join(INPUT_FOLDER, sessionId);
  fs.ensureDirSync(userFolder);

  const progressEmitter = new events.EventEmitter();
  req.session.gmailProgressEmitter = progressEmitter;

  // Create a task object
  const task = {
    sessionId,
    userFolder,
    startDate: req.body.startDate,
    endDate: req.body.endDate,
    idNumber: req.body.idNumber || '',
    progressEmitter,
    req,
  };

  // Add task to the Gmail queue
  gmailTaskQueue.push(task);

  // Inform the user if they are in a queue
  if (isGmailProcessing) {
    const queuePosition = gmailTaskQueue.length;
    progressEmitter.emit('progress', [
      { status: `Your task is in a queue at position ${queuePosition}. It will start processing shortly.`, progress: 0 },
    ]);
  } else {
    // Start processing immediately if no other tasks are running
    processNextGmailTask();
  }

  // Send response to the client indicating the session ID
  res.json({ sessionId });
});




// Endpoint for Gmail Progress
app.get('/gmail-progress/:uniqueId', (req, res) => {
  const { uniqueId } = req.params;
  const progressEmitter = progressEmitters[uniqueId];

  if (!progressEmitter) {
    res.status(404).end();
    return;
  }

  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.flushHeaders();

  const onProgress = (data) => {
    res.write(`data: ${JSON.stringify(data)}\n\n`);
  };

  progressEmitter.on('progress', onProgress);

  req.on('close', () => {
    progressEmitter.removeListener('progress', onProgress);
  });
});

function formatDateForGmail(date) {
  return format(date, 'yyyy/MM/dd');
}

// Function to download Gmail attachments
async function downloadGmailAttachments(auth, startDate, endDate, folderPath) {
  const gmail = google.gmail({ version: 'v1', auth });

  const excludedSubjectKeywords = ['חשבון עסקה'];
  endDate.setHours(23, 59, 59, 999);
  const queryEndDate = new Date(endDate.getTime() + 24 * 60 * 60 * 1000);
  // Prepare date queries
  const startDateQuery = formatDateForGmail(startDate);
  const endDateQuery = formatDateForGmail(queryEndDate);

  const query = `after:${startDateQuery} before:${endDateQuery}`;
  console.log('Gmail query:', query);

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
                console.log(`Saved attachment: ${filePath}`);
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




// User Logout Route
app.get('/logout', (req, res) => {
  req.session.destroy();
  res.redirect('/');
});

// Start the Server
const PORT = process.env.PORT || 8080;
app.listen(PORT, () => {
  console.log(`Server is running on port ${PORT}`);
});
