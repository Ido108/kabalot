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

const app = express();

// Middleware
app.use(cors()); // Allow CORS for all origins
app.use(express.static('public'));
app.set('view engine', 'ejs');

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
  projectId: 'your-gcp-project-id', // Replace with your GCP project ID
  location: 'us', // Processor location
  processorId: 'your-processor-id', // Your actual processor ID
};

// Use base64 encoded service account credentials from environment variable
const SERVICE_ACCOUNT_BASE64 = process.env.SERVICE_ACCOUNT_BASE64;

// Initialize Google OAuth2 Client
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
  password = PASSWORD_PROTECTED_PDF_PASSWORD
) {
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
 * Process Uploaded Files (PDFs and Images)
 * @param {Array} files - Array of uploaded file objects
 * @param {string} folderPath - Path to save the Excel file
 * @returns {Array} - Array of extracted expense data
 */
async function processUploadedFiles(files, folderPath) {
  // Authenticate with Service Account for Document AI
  const serviceAccountAuth = authenticateServiceAccount();
  await serviceAccountAuth.authorize(); // Ensure the client is authorized

  const expenses = [];

  for (const file of files) {
    const originalFilePath = file.path;
    console.log(`Processing file: ${originalFilePath}`);

    const ext = path.extname(originalFilePath).toLowerCase();
    const isPDF = ext === '.pdf';

    let processedFilePath = originalFilePath;

    if (isPDF) {
      // Check if the PDF is encrypted
      const isEncrypted = await isPdfEncrypted(originalFilePath);
      if (isEncrypted) {
        console.log('PDF is encrypted. Attempting to unlock:', originalFilePath);
        // Attempt to unlock PDF
        const unlockedPath = await unlockPdf(originalFilePath);
        if (unlockedPath) {
          // Overwrite the original file with the unlocked PDF
          fs.copyFileSync(unlockedPath, originalFilePath);
          fs.unlinkSync(unlockedPath); // Remove the temporary unlocked file
          console.log('PDF unlocked and overwritten:', originalFilePath);
          processedFilePath = originalFilePath; // Continue processing the unlocked file
        } else {
          console.log('Skipping locked PDF:', originalFilePath);
          continue; // Skip processing if PDF is locked and couldn't be unlocked
        }
      } else {
        console.log(
          'PDF is not encrypted. Proceeding without unlocking:',
          originalFilePath
        );
        // No action needed since we're processing in place
      }
    } else {
      console.log('File is an image. Proceeding to process:', originalFilePath);
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

// Ensure input folder exists
fs.ensureDirSync(INPUT_FOLDER);

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

    try {
      // Process the uploaded files
      const expenses = await processUploadedFiles(req.files, outputFolder);

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

// Start the Server
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server is running on port ${PORT}`);
});
