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
const events = require('events'); // For progress events
const crypto = require('crypto'); // For file hashing

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
 */
async function createExpenseExcel(expenses, folderPath, name, startDate, endDate) {
  // Ensure valid dates
  const validStartDate = parse(startDate, 'yyyy-MM-dd', new Date());
  const validEndDate = parse(endDate, 'yyyy-MM-dd', new Date());

  // Format dates
  const startDateFormatted = format(validStartDate, 'dd-MM-yy');
  const endDateFormatted = format(validEndDate, 'dd-MM-yy');
  
  // Create base filename
  let baseFileName = `${startDateFormatted}-to-${endDateFormatted}`;
  if (name) {
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
 * Process a single file
 */
async function processFile(filePath, serviceAccountAuth, password) {
  const ext = path.extname(filePath).toLowerCase();
  const isPDF = ext === '.pdf';
  let processedFilePath = filePath;

  if (isPDF) {
    try {
      // Check if the PDF is encrypted
      const isEncrypted = await isPdfEncrypted(filePath);
      if (isEncrypted) {
        console.log('PDF is encrypted. Attempting to unlock:', filePath);
        // Attempt to unlock PDF with provided password
        const unlockedPath = await unlockPdf(filePath, password);
        if (unlockedPath) {
          // Overwrite the original file with the unlocked PDF
          fs.copyFileSync(unlockedPath, filePath);
          fs.unlinkSync(unlockedPath); // Remove the temporary unlocked file
          console.log('PDF unlocked and overwritten:', filePath);
          processedFilePath = filePath; // Continue processing the unlocked file
        } else {
          console.log('Failed to unlock PDF:', filePath);
          return null; // Skip processing if PDF is locked and couldn't be unlocked
        }
      } else {
        console.log('PDF is not encrypted. Proceeding without unlocking:', filePath);
      }
    } catch (error) {
      console.error('Error processing PDF:', filePath, error);
      return null; // Skip this file if there's an error
    }
  } else {
    console.log('File is an image. Proceeding to process:', filePath);
  }

  // Parse the receipt with Document AI
  const expenseData = await parseReceiptWithDocumentAI(processedFilePath, serviceAccountAuth);
  return expenseData;
}

/**
 * Calculate File Hash
 */
function calculateFileHash(filePath) {
  return new Promise((resolve, reject) => {
    const hash = crypto.createHash('sha256');
    const stream = fs.createReadStream(filePath);
    stream.on('data', (data) => {
      hash.update(data);
    });
    stream.on('end', () => {
      resolve(hash.digest('hex'));
    });
    stream.on('error', (err) => {
      reject(err);
    });
  });
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

// Progress Emitters
const uploadProgressEmitters = {};
const gmailProgressEmitters = {};

// Duplicate Files Set
const processedFilesSet = new Set();

// Routes

// Home Route - Serve the upload form
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Handle File Upload and Processing with Progress Logging
app.post('/upload', upload, async (req, res) => {
  const uploadId = Date.now().toString();
  const progressEmitter = new events.EventEmitter();
  uploadProgressEmitters[uploadId] = progressEmitter;

  // Send the uploadId to the client
  res.json({ uploadId });

  if (!req.files || req.files.length === 0) {
    progressEmitter.emit('progress', [{ status: 'No files uploaded', progress: 0 }]);
    delete uploadProgressEmitters[uploadId];
    return;
  }

  console.log(`Received ${req.files.length} file(s). Starting processing...`);

  try {
    const files = req.files.map((file) => file.path);
    const totalFiles = files.length;
    const name = req.body.name || ''; // Optional name
    const idNumber = req.body.idNumber || ''; // Optional ID number for PDF password

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
      const fileName = path.basename(filePath);

      // Check for duplicate files
      const fileHash = await calculateFileHash(filePath);
      if (processedFilesSet.has(fileHash)) {
        console.log(`Skipping duplicate file: ${fileName}`);
        progressData[i].status = 'Skipped (Duplicate)';
        progressData[i].progress = 100;
        emitProgress();
        continue;
      } else {
        processedFilesSet.add(fileHash);
      }

      // Update status to 'Processing'
      progressData[i].status = 'Processing';
      progressData[i].progress = 25;
      emitProgress();

      // Process the file
      const expenseData = await processFile(filePath, serviceAccountAuth, idNumber);

      if (expenseData) {
        expenses.push(expenseData);
        progressData[i].status = 'Completed';
        progressData[i].progress = 100;
      } else {
        progressData[i].status = 'Failed';
        progressData[i].progress = 100;
      }

      emitProgress();
    }

    // Create Expense Excel File
    if (expenses.length > 0) {
      const startDate = formatDate(new Date());
      const endDate = formatDate(new Date());

      const excelPath = await createExpenseExcel(
        expenses,
        INPUT_FOLDER,
        name,
        startDate,
        endDate
      );
      console.log('Expense summary Excel file created.');

      // Provide a download link to the Excel file
      const excelFileName = encodeURIComponent(path.basename(excelPath));
      const csvUrl = `/download/${excelFileName}`;
      progressEmitter.emit('progress', [
        ...progressData,
        { status: 'Processing complete. Download the file below.', progress: 100, downloadLink: csvUrl },
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
      ...progressData,
      { status: `Processing Error: ${processingError.message}`, progress: 100 },
    ]);
  } finally {
    delete uploadProgressEmitters[uploadId];
  }
});

// Endpoint for Upload Progress
app.get('/upload-progress/:uploadId', (req, res) => {
  const uploadId = req.params.uploadId;
  const progressEmitter = uploadProgressEmitters[uploadId];

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

// Download Route - Serve the generated Excel file
app.get('/download/:filename', (req, res) => {
  const { filename } = req.params;
  const decodedFilename = decodeURIComponent(filename);

  // Search for the file in INPUT_FOLDER and its subfolders
  const findFile = (dir) => {
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

  const filePath = findFile(INPUT_FOLDER);

  if (filePath && fs.existsSync(filePath)) {
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
    const email = req.session.email || req.query.email || '';
    if (!email) {
      res.redirect('/gmail-auth');
      return;
    }

    // Store the email in session
    req.session.email = email;

    // Generate an OAuth URL with login_hint to prefill the email
    const authUrl = oAuth2Client.generateAuthUrl({
      access_type: 'offline',
      scope: ['https://www.googleapis.com/auth/gmail.readonly'],
      prompt: 'consent',
      login_hint: email,
    });

    res.redirect(authUrl);
  }
}

// Route to display Gmail authentication page
app.get('/gmail-auth', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'gmail-auth.html'));
});

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
    // Redirect to the Gmail processing page
    res.redirect('/process-gmail');
  } catch (error) {
    console.error('Error retrieving access token', error);
    res.status(500).send('Authentication failed');
  }
});

// Route to display Gmail processing page
app.get('/process-gmail', authenticateGmail, (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'gmail.html'));
});

// Handle Gmail Processing with Progress Logging
const additionalUpload = multer({
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
}).array('additionalFiles', 100);

app.post('/process-gmail', authenticateGmail, additionalUpload, async (req, res) => {
  const sessionId = Date.now().toString();
  const progressEmitter = new events.EventEmitter();
  gmailProgressEmitters[sessionId] = progressEmitter;

  // Send the sessionId to the client
  res.json({ sessionId });

  try {
    const auth = req.oAuth2Client;
    const { startDate, endDate, name, idNumber } = req.body;

    console.log('Processing Gmail attachments from', startDate, 'to', endDate);

    const startDateObj = new Date(startDate);
    const endDateObj = new Date(endDate);
    endDateObj.setHours(23, 59, 59, 999); // Set to end of day

    progressEmitter.emit('progress', [{ status: 'Downloading Gmail attachments...', progress: 10 }]);

    const attachmentsFolder = await downloadGmailAttachments(auth, startDateObj, endDateObj);
    console.log('Attachments downloaded to:', attachmentsFolder);

    let files = fs.readdirSync(attachmentsFolder).map((file) =>
      path.join(attachmentsFolder, file)
    );
    console.log('Files found:', files);

    // Include additional files
    if (req.files && req.files.length > 0) {
      for (const file of req.files) {
        files.push(file.path);
      }
    }

    if (files.length === 0) {
      progressEmitter.emit('progress', [{ status: 'No attachments found.', progress: 100 }]);
      return;
    }

    progressEmitter.emit('progress', [{ status: 'Processing files...', progress: 30 }]);

    const expenses = [];
    const serviceAccountAuth = authenticateServiceAccount();
    await serviceAccountAuth.authorize();

    for (let i = 0; i < files.length; i++) {
      const filePath = files[i];
      const fileName = path.basename(filePath);

      // Check for duplicate files
      const fileHash = await calculateFileHash(filePath);
      if (processedFilesSet.has(fileHash)) {
        console.log(`Skipping duplicate file: ${fileName}`);
        progressEmitter.emit('progress', [{ status: `Skipping duplicate file: ${fileName}`, progress: 100 }]);
        continue;
      } else {
        processedFilesSet.add(fileHash);
      }

      // Update progress
      const progressPercent = 30 + ((i + 1) / files.length) * 50; // Between 30% and 80%
      progressEmitter.emit('progress', [{ status: `Processing ${fileName}...`, progress: progressPercent }]);

      // Process the file
      const expenseData = await processFile(filePath, serviceAccountAuth, idNumber);

      if (expenseData) {
        expenses.push(expenseData);
      }
    }

    progressEmitter.emit('progress', [{ status: 'Creating Excel file...', progress: 80 }]);

    const excelPath = await createExpenseExcel(
      expenses,
      attachmentsFolder,
      name,
      startDate,
      endDate
    );
    console.log('Expenses extracted:', expenses.length);
    console.log('Excel file created at:', excelPath);

    // Provide download link
    const excelFileName = encodeURIComponent(path.basename(excelPath));
    const csvUrl = `/download/${excelFileName}`;

    progressEmitter.emit('progress', [{ status: 'Processing complete. Download the file below.', progress: 100, downloadLink: csvUrl }]);

  } catch (error) {
    console.error('Error processing Gmail attachments:', error);
    progressEmitter.emit('progress', [{ status: `Error: ${error.message}`, progress: 100 }]);
  } finally {
    delete gmailProgressEmitters[sessionId];
  }
});

// Endpoint for Gmail Progress
app.get('/gmail-progress/:sessionId', (req, res) => {
  const sessionId = req.params.sessionId;
  const progressEmitter = gmailProgressEmitters[sessionId];

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

// Function to download Gmail attachments
async function downloadGmailAttachments(auth, startDate, endDate) {
  const gmail = google.gmail({ version: 'v1', auth });
  const attachmentsFolder = path.join(INPUT_FOLDER, 'attachments');
  fs.ensureDirSync(attachmentsFolder);

  const query = `after:${formatDate(startDate)} before:${formatDate(endDate)} has:attachment (filename:pdf OR filename:jpg OR filename:jpeg OR filename:png OR filename:tif OR filename:tiff)`;

  let messages = [];
  let nextPageToken = null;

  do {
    const res = await gmail.users.messages.list({
      userId: 'me',
      q: query,
      maxResults: 100,
      pageToken: nextPageToken,
    });

    if (res.data.messages) {
      messages = messages.concat(res.data.messages);
    }

    nextPageToken = res.data.nextPageToken;
  } while (nextPageToken);

  console.log(`Found ${messages.length} messages with attachments.`);

  for (const message of messages) {
    const msg = await gmail.users.messages.get({
      userId: 'me',
      id: message.id,
    });

    const parts = msg.data.payload.parts;
    if (!parts) continue;

    for (const part of parts) {
      if (part.filename && part.body && part.body.attachmentId) {
        const attachment = await gmail.users.messages.attachments.get({
          userId: 'me',
          messageId: message.id,
          id: part.body.attachmentId,
        });

        const data = attachment.data.data;
        const fileData = Buffer.from(data, 'base64');

        const sanitizedFilename = sanitize(part.filename) || 'unnamed_attachment';
        const filePath = path.join(attachmentsFolder, sanitizedFilename);

        // Check for duplicate files
        const fileHash = crypto.createHash('sha256').update(fileData).digest('hex');
        if (processedFilesSet.has(fileHash)) {
          console.log(`Skipping duplicate attachment: ${sanitizedFilename}`);
          continue;
        } else {
          processedFilesSet.add(fileHash);
        }

        fs.writeFileSync(filePath, fileData);
        console.log(`Saved attachment: ${filePath}`);
      }
    }
  }

  return attachmentsFolder;
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
