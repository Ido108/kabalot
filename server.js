require('dotenv').config(); // Load environment variables
const nodeBuffer = require('buffer');
if (!nodeBuffer.SlowBuffer && nodeBuffer.Buffer) {
  nodeBuffer.SlowBuffer = nodeBuffer.Buffer;
}
if (typeof global.SlowBuffer === 'undefined' && nodeBuffer.SlowBuffer) {
  global.SlowBuffer = nodeBuffer.SlowBuffer;
}
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
const EventEmitter = require('events');
const session = require('express-session');
const events = require('events');
const archiver = require('archiver');
const nodemailer = require('nodemailer'); // Added for email
let franc;
import('franc-min')
  .then((module) => {
    franc = module.franc || module.default;
  })
  .catch((error) => {
    console.error('Failed to load franc-min:', error);
  });

const app = express();

// Email transporter (Configure EMAIL_USER and EMAIL_PASS)
const transporter = nodemailer.createTransport({
  service: 'gmail',
  auth: {
    user: process.env.EMAIL_USER,
    pass: process.env.EMAIL_PASS
  }
});

app.use(cors());
app.use(express.static('public'));
app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.set('view engine', 'ejs');

app.use(
  session({
    secret: process.env.SESSION_SECRET || 'your_session_secret',
    resave: false,
    saveUninitialized: true,
    cookie: {
      maxAge: 3600000,
    },
  })
);

const taskQueue = [];
let isProcessing = false;
const progressEmitters = new Map();

const gmailTaskQueue = [];
let isGmailProcessing = false;

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

const INPUT_FOLDER =
  process.env.INPUT_FOLDER ||
  path.join(__dirname, 'input_files', folderName);

if (!fs.existsSync(INPUT_FOLDER)) {
  fs.mkdirSync(INPUT_FOLDER, { recursive: true });
}

console.log(`Input folder: ${INPUT_FOLDER}`);

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
const LANGUAGE_HINTS = ['iw', 'en'];
const MAX_CONCURRENT_FILE_PROCESSING = Math.max(
  1,
  parseInt(process.env.MAX_CONCURRENT_FILE_PROCESSING || '3', 10)
);
const PROCESSED_FILE_CACHE_LIMIT = Math.max(
  1,
  parseInt(process.env.PROCESSED_FILE_CACHE_LIMIT || '200', 10)
);
const processedFileCache = new Map();

function authenticateServiceAccount() {
  if (!SERVICE_ACCOUNT_BASE64) {
    throw new Error('SERVICE_ACCOUNT_BASE64 is not set.');
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

function getQueuePosition(queue, sessionId) {
  return queue.findIndex(task => task.sessionId === sessionId) + 1;
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
  // Mock implementation
  const exchangeRate = 3.5;
  return exchangeRate;
}

async function unlockPdf(filePath, password) {
  if (!PDFCO_API_KEY) {
    throw new Error('PDFCO_API_KEY not set');
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
      console.error('Error uploading PDF:', JSON.stringify(uploadResponse.data));
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
      console.error('Error unlocking PDF:', JSON.stringify(unlockResponse.data));
      return null;
    }
  } catch (error) {
    console.error('Error in unlockPdf:', error.message);
    return null;
  }
}

async function processNextTask() {
  if (taskQueue.length === 0) {
    isProcessing = false;
    return;
  }

  isProcessing = true;
  const task = taskQueue.shift();
  const {
    sessionId,
    userFolder,
    files,
    name,
    idNumber,
    email,
    progressEmitter,
    req,
  } = task;

  try {
    progressEmitter.emit('progress', [
      { status: 'העיבוד החל.', progress: 0 },
    ]);

    if (!files || files.length === 0) {
      progressEmitter.emit('progress', [{ status: 'לא הועלו קבצים.', progress: 100 }]);
      return;
    }

    console.log(`Processing ${files.length} file(s) for session ${sessionId}...`);

    const progressData = files.map((filePath) => ({
      fileName: path.basename(filePath),
      status: 'Pending',
      progress: 0,
    }));

    const emitProgress = () => {
      progressEmitter.emit('progress', progressData);
    };

    emitProgress();

    const expenses = [];
    const serviceAccountAuth = authenticateServiceAccount();
    await serviceAccountAuth.authorize();

    await processFilesConcurrently(files, async (filePath, index) => {
      progressData[index].status = 'Processing';
      progressData[index].progress = 25;
      emitProgress();

      const expenseData = await getExpenseDataWithCache(
        filePath,
        serviceAccountAuth,
        idNumber
      );

      if (expenseData) {
        expenses.push(expenseData);
        progressData[index].status = 'Completed';
        progressData[index].progress = 100;
        progressData[index].businessName = expenseData.BusinessName || 'N/A';
        progressData[index].date = expenseData.Date || 'N/A';
        progressData[index].totalPrice = expenseData.TotalPrice
          ? parseFloat(expenseData.TotalPrice).toFixed(2)
          : 'N/A';
      } else {
        progressData[index].status = 'Failed';
        progressData[index].progress = 100;
      }

      emitProgress();
    });

    if (expenses.length > 0) {
      const startDate = formatDate(new Date());
      const endDate = formatDate(new Date());
      const excelPath = await createExpenseExcel(
        expenses,
        userFolder,
        'סיכום הוצאות',
        startDate,
        endDate,
        name
      );

      const zipFileName = `processed_files_${Date.now()}.zip`;
      const zipFilePath = await createZipFile(files, userFolder, zipFileName);

      const excelFileName = encodeURIComponent(path.basename(excelPath));
      const zipFileNameEncoded = encodeURIComponent(path.basename(zipFilePath));
      const baseUrl = process.env.BASE_URL || 'http://localhost:8080';
      const excelUrl = `${baseUrl}/download/${excelFileName}`;
      const zipUrl = `${baseUrl}/download/${zipFileNameEncoded}`;

      // Store last results in session
      req.session.lastResults = {
        excelUrl,
        zipUrl,
        timestamp: new Date().toISOString()
      };

      progressEmitter.emit('progress', [
        ...progressData,
        {
          status: 'העיבוד הושלם. ניתן להוריד את הקבצים בשעה הקרובה.',
          progress: 100,
          downloadLinks: [
            { label: 'הורד קובץ אקסל', url: excelUrl },
            { label: 'הורד קבצים מעובדים (ZIP)', url: zipUrl },
          ],
        },
      ]);

      // Send email with results if email provided
      if (email) {
        const mailOptions = {
          from: process.env.EMAIL_USER,
          to: email,
          subject: 'Your Processed Expense Files',
          html: buildResultEmailHtml(excelUrl, zipUrl),
        };

        transporter.sendMail(mailOptions, (error, info) => {
          if (error) {
            console.error('Error sending email:', error);
          } else {
            console.log('Email sent:', info.response);
          }
        });
      }
    } else {
      progressEmitter.emit('progress', [
        ...progressData,
        { status: 'לא זוהו נתוני הוצאה.', progress: 100 },
      ]);
    }
  } catch (processingError) {
    console.error('Processing Error:', processingError.message);
    progressEmitter.emit('progress', [
      { status: `שגיאת עיבוד: ${processingError.message}`, progress: 100 },
    ]);
  } finally {
    scheduleFolderDeletion(userFolder, 3600000);
    progressEmitters.delete(sessionId);
    isProcessing = false;
    processNextTask();
  }
}

async function isPdfEncrypted(filePath) {
  try {
    const pdfBuffer = await fs.promises.readFile(filePath);
    const pdfDoc = await PDFDocument.load(pdfBuffer, { ignoreEncryption: true });

    if (pdfDoc.isEncrypted) {
      return true;
    }
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

function detectCurrency(value, preferredLanguage = 'en') {
  if (!value) {
    return preferredLanguage === 'iw' ? 'ILS' : 'USD';
  }

  const normalized = value.toUpperCase();
  const usdRegex = /\$|USD|DOLLAR|דולר/;
  const ilsRegex = /₪|ILS|NIS|שח|ש"ח/;

  if (usdRegex.test(normalized)) {
    return 'USD';
  }
  if (ilsRegex.test(normalized)) {
    return 'ILS';
  }

  return preferredLanguage === 'iw' ? 'ILS' : 'USD';
}

function detectLanguageFromText(text) {
  if (!text) return null;
  if (!franc) {
    console.warn('Language detection skipped: franc-min not loaded yet.');
    return null;
  }
  const francCode = franc(text, { whitelist: ['eng', 'heb'] });
  if (francCode === 'heb') {
    return 'iw';
  }
  if (francCode === 'eng') {
    return 'en';
  }
  return null;
}

function detectDocumentLanguages(document, fallbackText = '') {
  const languagesMap = new Map();

  (document.pages || []).forEach((page) => {
    (page.detectedLanguages || []).forEach((language) => {
      if (!language.languageCode) {
        return;
      }
      const accumulated = languagesMap.get(language.languageCode) || 0;
      languagesMap.set(
        language.languageCode,
        Math.max(accumulated, language.confidence || 0)
      );
    });
  });

  let detectedLanguages = Array.from(languagesMap.entries()).map(
    ([languageCode, confidence]) => ({
      languageCode,
      confidence,
    })
  );

  detectedLanguages.sort(
    (a, b) => (b.confidence || 0) - (a.confidence || 0)
  );

  if (detectedLanguages.length === 0 && fallbackText) {
    const fallbackLanguage = detectLanguageFromText(fallbackText);
    if (fallbackLanguage) {
      detectedLanguages = [
        { languageCode: fallbackLanguage, confidence: 1 },
      ];
    }
  }

  return detectedLanguages;
}

function getPrimaryLanguage(detectedLanguages) {
  if (!detectedLanguages || detectedLanguages.length === 0) {
    return 'en';
  }

  return detectedLanguages[0].languageCode || 'en';
}

function isRtlLanguage(languageCode) {
  return ['iw', 'he', 'ar', 'fa'].includes(languageCode);
}

function formatDetectedLanguages(languages) {
  if (!languages || languages.length === 0) {
    return 'und';
  }

  return languages
    .map((language) => {
      const confidence = language.confidence || 0;
      return `${language.languageCode}:${(confidence * 100).toFixed(1)}%`;
    })
    .join(', ');
}

function buildResultEmailHtml(excelUrl, zipUrl) {
  return `<p>שלום,</p>
  <p>הקבצים שעיבדנו זמינים להורדה בשעה הקרובה:</p>
  <ul>
    <li><a href="${excelUrl}">קובץ אקסל מסודר</a></li>
    <li><a href="${zipUrl}">ארכיון ZIP עם כל המסמכים</a></li>
  </ul>
  <p>לשמירה על פרטיותך, הקישורים יפוגו בתוך שעה.</p>
  <p>תודה,<br>Receipt Cloud</p>`;
}

async function getFileChecksum(filePath) {
  return new Promise((resolve, reject) => {
    const hash = crypto.createHash('sha256');
    const stream = fs.createReadStream(filePath);
    stream.on('data', (data) => hash.update(data));
    stream.on('error', (error) => reject(error));
    stream.on('end', () => resolve(hash.digest('hex')));
  });
}

function rememberCacheEntry(key, value) {
  if (!key || !value) {
    return;
  }
  if (processedFileCache.has(key)) {
    processedFileCache.delete(key);
  }
  processedFileCache.set(key, value);
  if (processedFileCache.size > PROCESSED_FILE_CACHE_LIMIT) {
    const oldestKey = processedFileCache.keys().next().value;
    processedFileCache.delete(oldestKey);
  }
}

async function getExpenseDataWithCache(filePath, serviceAccountAuth, idNumber) {
  try {
    const checksum = await getFileChecksum(filePath);
    if (processedFileCache.has(checksum)) {
      return processedFileCache.get(checksum);
    }
    const expenseData = await processFile(filePath, serviceAccountAuth, idNumber);
    if (expenseData) {
      rememberCacheEntry(checksum, expenseData);
    }
    return expenseData;
  } catch (error) {
    console.error('Failed to compute checksum for caching:', error.message);
    return processFile(filePath, serviceAccountAuth, idNumber);
  }
}

async function processFilesConcurrently(files, handler) {
  if (!files || files.length === 0) {
    return;
  }

  const concurrency = Math.min(MAX_CONCURRENT_FILE_PROCESSING, files.length);
  let cursor = 0;

  const workers = Array.from({ length: concurrency }, () =>
    (async () => {
      while (true) {
        let currentIndex;
        if (cursor >= files.length) {
          break;
        }
        currentIndex = cursor;
        cursor += 1;

        try {
          await handler(files[currentIndex], currentIndex);
        } catch (error) {
          console.error(
            `Error while processing file ${files[currentIndex]}:`,
            error.message
          );
        }
      }
    })()
  );

  await Promise.all(workers);
}

const exchangeRateCache = {};

const additionalStorage = multer.diskStorage({
  destination: function (req, file, cb) {
    const attachmentsFolder = getAttachmentsFolder(req);
    fs.ensureDirSync(attachmentsFolder);
    cb(null, attachmentsFolder);
  },
  filename: function (req, file, cb) {
    const sanitized = sanitize(file.originalname) || 'unnamed_attachment';
    cb(null, sanitized);
  },
});

function getAttachmentsFolder(req) {
  const startDate = req.body.startDate;
  const endDate = req.body.endDate;

  const startDateObj = new Date(startDate);
  const endDateObj = new Date(endDate);
  endDateObj.setHours(23, 59, 59, 999);

  const folderName = `קבלות ${formatDate(startDateObj)} עד ${formatDate(endDateObj)}`;
  const folderPath = path.join(INPUT_FOLDER, folderName);
  return folderPath;
}

const additionalUpload = multer({
  storage: additionalStorage,
  fileFilter: function (req, file, cb) {
    const ext = path.extname(file.originalname).toLowerCase();
    const allowedExtensions = ['.pdf', '.jpg', '.jpeg', '.png', '.tiff', '.tif'];
    if (allowedExtensions.includes(ext)) {
      cb(null, true);
    } else {
      console.warn(`Skipped unsupported file type: ${file.originalname}`);
      cb(null, false);
    }
  },
  limits: { fileSize: 50 * 1024 * 1024 },
}).array('additionalFiles', 100);

async function parseReceiptWithDocumentAI(filePath, serviceAccountAuth) {
  const { projectId, location, processorId } = DOCUMENT_AI_CONFIG;
  const url = `https://${location}-documentai.googleapis.com/v1/projects/${projectId}/locations/${location}/processors/${processorId}:process`;

  try {
    const fileContent = fs.readFileSync(filePath);
    const encodedFile = fileContent.toString('base64');
    const ext = path.extname(filePath).toLowerCase();
    let mimeType = 'application/pdf';

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
      processOptions: {
        ocrConfig: {
          languageHints: LANGUAGE_HINTS,
          enableNativePdfParsing: true,
          enableImagePreprocessing: true,
        },
      },
    };

    if (!serviceAccountAuth.credentials || !serviceAccountAuth.credentials.access_token) {
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
      Language: '',
      LanguageDirection: '',
      DetectedLanguages: '',
    };

    const documentText = document.text || '';
    const detectedLanguages = detectDocumentLanguages(document, documentText);
    const primaryLanguage = getPrimaryLanguage(detectedLanguages);
    result.Language = primaryLanguage;
    result.LanguageDirection = isRtlLanguage(primaryLanguage) ? 'rtl' : 'ltr';
    result.DetectedLanguages = formatDetectedLanguages(detectedLanguages);

    let hasUSD = false;
    let originalUSD = 0;
    let convertedTotalPrice = 0;
    let invoiceDate = '';

    for (const entity of entities) {
      let value = '';
      let currencyCode = '';

      if (entity.normalizedValue && entity.normalizedValue.moneyValue) {
        value = entity.normalizedValue.moneyValue.amount;
        currencyCode = entity.normalizedValue.moneyValue.currencyCode || '';
      } else {
        value = entity.mentionText || '';
        currencyCode = detectCurrency(value, primaryLanguage);
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

    if (!invoiceDate) {
      invoiceDate = formatDate(new Date());
      result['Date'] = invoiceDate;
    } else {
      try {
        const dateFormats = [          'yyyy-MM-dd', 'MM/dd/yyyy', 'dd/MM/yyyy', 'dd-MM-yyyy', 'MM-dd-yyyy',          'yyyy/MM/dd','dd.MM.yyyy','yyyy.MM.dd','MMMM dd, yyyy','MMM dd, yyyy'        ];
        let parsedDate;
        for (const dateFormat of dateFormats) {
          parsedDate = parse(invoiceDate, dateFormat, new Date());
          if (!isNaN(parsedDate)) {
            break;
          }
        }
        if (isNaN(parsedDate)) {
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
      exchangeRate = exchangeRateCache[invoiceDate];
      if (!exchangeRate) {
        exchangeRate = await getUsdToIlsExchangeRate(invoiceDate);
        exchangeRateCache[invoiceDate] = exchangeRate;
      }
      console.log(`Using exchange rate: ${exchangeRate} for date: ${invoiceDate}`);
    }

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
      result['Currency'] = 'ILS';
      result['TotalPrice'] = convertedTotalPrice;
    } else {
      result['Currency'] = 'ILS';
    }

    return result;
  } catch (error) {
    console.error('Error in parseReceiptWithDocumentAI:', error.message);
    return {};
  }
}

async function createExpenseExcel(expenses, folderPath, filePrefix, startDate, endDate, name) {
  const validStartDate = parse(startDate, 'yyyy-MM-dd', new Date());
  const validEndDate = parse(endDate, 'yyyy-MM-dd', new Date());

  const startDateFormatted = format(validStartDate, 'dd-MM-yy');
  const endDateFormatted = format(validEndDate, 'dd-MM-yy');

  let baseFileName = `${filePrefix}-${startDateFormatted}-to-${endDateFormatted}`;
  if (name && name.trim() !== '') {
    const sanitizedName = sanitize(name).replace(/ /g, '_');
    baseFileName += `-${sanitizedName}`;
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
  const worksheet = workbook.addWorksheet('Expenses', {
    views: [{ rightToLeft: true }],
  });

  worksheet.columns = [
    { header: 'File Name', key: 'FileName', width: 30 },
    { header: 'Business Name', key: 'BusinessName', width: 26 },
    { header: 'Business Number', key: 'BusinessNumber', width: 20 },
    { header: 'Invoice Date', key: 'Date', width: 16 },
    { header: 'Invoice Number', key: 'InvoiceNumber', width: 20 },
    { header: 'Amount (excl. VAT)', key: 'PriceWithoutVat', width: 20 },
    { header: 'VAT Amount', key: 'VAT', width: 15 },
    { header: 'Total Amount (ILS)', key: 'TotalPrice', width: 22 },
    { header: 'Original Total (USD)', key: 'OriginalTotalUSD', width: 18 },
    { header: 'Primary Language', key: 'Language', width: 18 },
    { header: 'Detected Languages', key: 'DetectedLanguages', width: 28 },
  ];

  worksheet.getRow(1).font = { bold: true, size: 12 };
  worksheet.getRow(1).alignment = { horizontal: 'center' };

  let totalWithoutVat = 0;
  let totalVAT = 0;
  let totalPrice = 0;
  let totalOriginalUSD = 0;

  expenses.forEach((expense) => {
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
      Language: expense['Language'] || '',
      DetectedLanguages: expense['DetectedLanguages'] || '',
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

  try {
    await workbook.xlsx.writeFile(fullPath);
    console.log('Expense summary Excel file created at:', fullPath);
    return fullPath;
  } catch (error) {
    console.error('Error saving Excel file:', error);
    throw new Error('Failed to save Excel file');
  }
}


async function createZipFile(files, outputFolder, zipFileName) {
  return new Promise((resolve, reject) => {
    const zipFilePath = path.join(outputFolder, zipFileName);
    const output = fs.createWriteStream(zipFilePath);
    const archive = archiver('zip', {
      zlib: { level: 9 },
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

async function processFile(filePath, serviceAccountAuth, idNumber) {
  const ext = path.extname(filePath).toLowerCase();
  const isPDF = ext === '.pdf';
  let processedFilePath = filePath;

  if (isPDF) {
    try {
      const encrypted = await isPdfEncrypted(filePath);
      if (encrypted) {
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
        console.log('PDF is not encrypted:', filePath);
      }
    } catch (error) {
      console.error('Error processing PDF:', filePath, error);
      return null;
    }
  } else {
    console.log('File is an image:', filePath);
  }

  const expenseData = await parseReceiptWithDocumentAI(processedFilePath, serviceAccountAuth);
  return expenseData;
}

fs.ensureDirSync(INPUT_FOLDER);

const storage = multer.diskStorage({
  destination: function (req, file, cb) {
    if (!req.session.sessionId) {
      req.session.sessionId = uuidv4();
    }
    const userFolder = path.join(INPUT_FOLDER, req.session.sessionId);
    fs.ensureDirSync(userFolder);
    cb(null, userFolder);
  },
  filename: function (req, file, cb) {
    const sanitized = sanitize(file.originalname) || 'unnamed_attachment';
    cb(null, sanitized);
  },
});

const upload = multer({
  storage: storage,
  fileFilter: function (req, file, cb) {
    const ext = path.extname(file.originalname).toLowerCase();
    const allowedExtensions = ['.pdf', '.jpg', '.jpeg', '.png', '.tiff', '.tif'];
    if (allowedExtensions.includes(ext)) {
      cb(null, true);
    } else {
      console.warn(`Skipped unsupported file type: ${file.originalname}`);
      cb(null, false);
    }
  },
  limits: { fileSize: 50 * 1024 * 1024 },
}).array('files', 100);


app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});


app.post('/upload', upload, async (req, res) => {
  if (!req.session.sessionId) {
    req.session.sessionId = uuidv4();
  }
  const sessionId = req.session.sessionId;
  const userFolder = path.join(INPUT_FOLDER, sessionId);
  fs.ensureDirSync(userFolder);

  const progressEmitter = new EventEmitter();
  progressEmitters.set(sessionId, progressEmitter);

  res.json({ sessionId });

  const task = {
    sessionId,
    userFolder,
    files: req.files ? req.files.map((file) => file.path) : [],
    name: req.body.name || '',
    idNumber: req.body.idNumber || '',
    email: req.body.email || '',
    progressEmitter,
    req,
  };

  taskQueue.push(task);

  if (isProcessing) {
    const queuePosition = getQueuePosition(taskQueue, sessionId);
    progressEmitter.emit('progress', [
      { status: `המשימה שלך בתור במקום ${queuePosition}. העיבוד יתחיל בקרוב.`, progress: 0, queuePosition },
    ]);
  } else {
    processNextTask();
  }
});

app.get('/upload-progress', (req, res) => {
  const sessionId = req.session.sessionId;
  const progressEmitter = progressEmitters.get(sessionId);

  if (!progressEmitter) {
    res.status(404).end();
    return;
  }

  req.setTimeout(0);
  res.setTimeout(0);

  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.flushHeaders();

  const heartbeatInterval = setInterval(() => {
    res.write(':\n\n');
  }, 30000);

  const onProgress = (data) => {
    res.write(`data: ${JSON.stringify(data)}\n\n`);
  };

  progressEmitter.on('progress', onProgress);

  req.on('close', () => {
    progressEmitter.removeListener('progress', onProgress);
    clearInterval(heartbeatInterval);
    progressEmitters.delete(sessionId);
  });
});

app.get('/download/:filename', (req, res) => {
  const { filename } = req.params;
  const decodedFilename = decodeURIComponent(filename);
  const sessionId = req.session.sessionId;

  if (!sessionId) {
    res.status(403).send('Access denied.');
    return;
  }

  const userFolder = path.join(INPUT_FOLDER, sessionId);

  function findFile(dir) {
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
  }

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

app.get('/is-authenticated', (req, res) => {
  if (req.session.tokens) {
    res.json({
      authenticated: true,
      profile: req.session.gmailProfile || null,
    });
  } else {
    res.json({ authenticated: false });
  }
});

async function authenticateGmail(req, res, next) {
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
    try {
      if (!req.session.gmailProfile) {
        const profile = await fetchGmailProfile(oAuth2Client);
        if (profile) {
          req.session.gmailProfile = profile;
        }
      }
      req.gmailProfile = req.session.gmailProfile;
    } catch (profileError) {
      console.error('Failed to load Gmail profile:', profileError.message);
    }
    next();
  } else {
    res.redirect('/start-gmail-auth');
  }
}

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
    req.session.tokens = tokens;
    try {
      const profile = await fetchGmailProfile(oAuth2Client);
      if (profile) {
        req.session.gmailProfile = profile;
      }
    } catch (profileError) {
      console.error('Unable to fetch Gmail profile after auth:', profileError.message);
    }
    res.redirect('/gmail');
  } catch (error) {
    console.error('Error retrieving access token', error);
    res.status(500).send('Authentication failed');
  }
});

app.get('/gmail-form', authenticateGmail, (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'gmail-auth.html'));
});

app.get('/gmail', authenticateGmail, (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'gmail.html'));
});

app.post('/process-gmail', authenticateGmail, additionalUpload, (req, res) => {
  if (!req.session.sessionId) {
    req.session.sessionId = uuidv4();
  }
  const sessionId = req.session.sessionId;
  const userFolder = path.join(INPUT_FOLDER, sessionId);
  fs.ensureDirSync(userFolder);
  const additionalFiles = req.files ? req.files.map(file => file.path) : [];
  const progressEmitter = new EventEmitter();
  progressEmitters.set(sessionId, progressEmitter);
  const gmailProfile = req.session.gmailProfile || null;
  const connectedEmail = gmailProfile?.emailAddress || '';

  const task = {
    sessionId,
    userFolder,
    startDate: req.body.startDate,
    endDate: req.body.endDate,
    idNumber: req.body.idNumber || '',
    email: connectedEmail,
    gmailProfile,
    progressEmitter,
    additionalFiles,
    req,
  };

  gmailTaskQueue.push(task);

  if (isGmailProcessing) {
    const queuePosition = getQueuePosition(gmailTaskQueue, sessionId);
    const queueStatus = connectedEmail
      ? `סנכרון Gmail אל ${connectedEmail} ממתין במקום ${queuePosition}.`
      : `סנכרון Gmail ממתין בתור במקום ${queuePosition}.`;
    progressEmitter.emit('progress', [
      {
        status: queueStatus,
        progress: 0,
        queuePosition,
      },
    ]);
  } else {
    processNextGmailTask();
  }

  res.json({ sessionId });
});

app.get('/gmail-progress', (req, res) => {
  const sessionId = req.session.sessionId;
  const progressEmitter = progressEmitters.get(sessionId);

  if (!progressEmitter) {
    res.status(404).end();
    return;
  }

  req.setTimeout(0);
  res.setTimeout(0);

  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.flushHeaders();

  const heartbeatInterval = setInterval(() => {
    res.write(':\n\n');
  }, 30000);

  const onProgress = (data) => {
    res.write(`data: ${JSON.stringify(data)}\n\n`);
  };

  progressEmitter.on('progress', onProgress);

  req.on('close', () => {
    progressEmitter.removeListener('progress', onProgress);
    clearInterval(heartbeatInterval);
    progressEmitters.delete(sessionId);
  });
});

async function processNextGmailTask() {
  if (gmailTaskQueue.length === 0) {
    isGmailProcessing = false;
    return;
  }

  isGmailProcessing = true;
  const task = gmailTaskQueue.shift();
  const {
    sessionId,
    userFolder,
    startDate,
    endDate,
    idNumber,
    email,
    gmailProfile,
    progressEmitter,
    req,
    additionalFiles,
  } = task;

  try {
    progressEmitter.emit('progress', [
      { status: 'העיבוד החל.', progress: 0 },
    ]);

    const auth = req.oAuth2Client;
    const connectedAccountLabel = gmailProfile?.emailAddress
      ? `Gmail - ${gmailProfile.emailAddress}`
      : 'Gmail Inbox';
    const customPrefix = connectedAccountLabel;

    progressEmitter.emit('progress', [{ status: 'מוריד קבצים מ-Gmail...', progress: 10 }]);

    await downloadGmailAttachments(auth, new Date(startDate), new Date(endDate), userFolder);

    let files = fs.readdirSync(userFolder).map((file) => path.join(userFolder, file));
    additionalFiles.forEach(file => {
      if (!files.includes(file)) {
        files.push(file);
      }
    });

    if (files.length === 0) {
      progressEmitter.emit('progress', [{ status: 'לא נמצאו קבצים בטווח המבוקש.', progress: 100 }]);
      return;
    }

    const progressData = files.map((filePath) => ({
      fileName: path.basename(filePath),
      status: 'Pending',
      progress: 0,
    }));

    const emitProgress = () => {
      progressEmitter.emit('progress', progressData);
    };

    emitProgress();

    const expenses = [];
    const serviceAccountAuth = authenticateServiceAccount();
    await serviceAccountAuth.authorize();

    await processFilesConcurrently(files, async (filePath, index) => {
      progressData[index].status = 'Processing';
      progressData[index].progress = 25;
      emitProgress();

      const expenseData = await getExpenseDataWithCache(
        filePath,
        serviceAccountAuth,
        idNumber
      );

      if (expenseData) {
        expenses.push(expenseData);
        progressData[index].status = 'Completed';
        progressData[index].progress = 100;
        progressData[index].businessName = expenseData.BusinessName || 'N/A';
        progressData[index].date = expenseData.Date || 'N/A';
        progressData[index].totalPrice = expenseData.TotalPrice
          ? parseFloat(expenseData.TotalPrice).toFixed(2)
          : 'N/A';
      } else {
        progressData[index].status = 'Failed';
        progressData[index].progress = 100;
      }

      emitProgress();
    });

    progressEmitter.emit('progress', [{ status: 'Creating Excel file...', progress: 80 }]);

    const excelPath = await createExpenseExcel(
      expenses,
      userFolder,
      customPrefix,
      startDate,
      endDate
    );
    const zipFileName = `processed_files_${Date.now()}.zip`;
    const zipFilePath = await createZipFile(files, userFolder, zipFileName);
    const excelFileName = encodeURIComponent(path.basename(excelPath));
    const zipFileNameEncoded = encodeURIComponent(path.basename(zipFilePath));
    const baseUrl = process.env.BASE_URL || 'http://localhost:8080';
    
    const excelUrl = `${baseUrl}/download/${excelFileName}`;
    const zipUrl = `${baseUrl}/download/${zipFileNameEncoded}`;
    
    req.session.lastResults = {
      excelUrl,
      zipUrl,
      timestamp: new Date().toISOString()
    };
    
    progressEmitter.emit('progress', [
      ...progressData,
      {
        status: 'העיבוד הושלם. ניתן להוריד את הקבצים בשעה הקרובה.',
        progress: 100,
        downloadLinks: [
          { label: 'הורד קובץ אקסל', url: excelUrl },
          { label: 'הורד קבצים מעובדים (ZIP)', url: zipUrl },
        ],
      },
    ]);
    
    if (email) {
      const mailOptions = {
        from: process.env.EMAIL_USER,
        to: email,
        subject: 'Your Processed Expense Files',
        html: buildResultEmailHtml(excelUrl, zipUrl),
      };
    
      transporter.sendMail(mailOptions, (error, info) => {
        if (error) {
          console.error('Error sending email:', error);
        } else {
          console.log('Email sent:', info.response);
        }
      });
    }

    scheduleFolderDeletion(userFolder, 3600000);
    req.session.gmailProgressEmitter = null;
  } catch (error) {
    console.error('Error processing Gmail attachments:', error);
    progressEmitter.emit('progress', [{ status: `Error: ${error.message}`, progress: 100 }]);
  } finally {
    scheduleFolderDeletion(userFolder, 3600000);
    progressEmitters.delete(sessionId);
    isGmailProcessing = false;
    processNextGmailTask();
  }
}

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


async function fetchGmailProfile(oAuth2Client) {
  const gmail = google.gmail({ version: 'v1', auth: oAuth2Client });
  const profileResponse = await gmail.users.getProfile({ userId: 'me' });
  return profileResponse.data;
}

async function downloadGmailAttachments(auth, startDate, endDate, folderPath) {
  const gmail = google.gmail({ version: 'v1', auth });

  const positiveSubjectKeywords = [    'קבלה',    'חשבונית',    'חשבונית מס',    'הקבלה',    'החשבונית',    'החשבונית החודשית',    'אישור תשלום',    'receipt',    'invoice',    'חשבון חודשי',  ];

  const excludedSenders = [    'חברת חשמל לישראל',    'עיריית תל אביב-יפו',    'ארנונה - עיריית תל-אביב-יפו',  ];
  const senderExceptionKeywords = [    'קבלה',    'חשבונית',    'חשבונית מס',    'הקבלה',  ];

  endDate.setHours(23, 59, 59, 999);
  const queryEndDate = new Date(endDate.getTime() + 24 * 60 * 60 * 1000);

  const startDateQuery = formatDate(startDate).replace(/-/g, '/');
  const endDateQuery = formatDate(queryEndDate).replace(/-/g, '/');

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

  for (const messageData of allMessageIds) {
    const msg = await gmail.users.messages.get({
      userId: 'me',
      id: messageData.id,
      format: 'full',
    });

    const headers = msg.data.payload.headers;
    const fromHeader = headers.find((h) => h.name.toLowerCase() === 'from');
    const subjectHeader = headers.find((h) => h.name.toLowerCase() === 'subject');

    const sender = fromHeader ? fromHeader.value : '';
    const subject = subjectHeader ? subjectHeader.value : '';

    const senderEmailMatch = sender.match(/<(.+?)>/);
    const senderEmail = senderEmailMatch ? senderEmailMatch[1] : sender;
    const senderName = sender.split('<')[0].trim();

    const lowerCaseSubject = subject.toLowerCase();

    let isExcludedSender = false;
    for (const excludedSender of excludedSenders) {
      if (
        senderName.includes(excludedSender) ||
        senderEmail.includes(excludedSender)
      ) {
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
        console.log(
          'Skipping message from excluded sender without exception keyword:',
          subject
        );
        continue;
      }
    }

    let subjectMatchesPositiveKeywords = false;
    for (const keyword of positiveSubjectKeywords) {
      if (lowerCaseSubject.includes(keyword.toLowerCase())) {
        subjectMatchesPositiveKeywords = true;
        break;
      }
    }

    if (msg.data.payload) {
      const parts = getParts(msg.data.payload);

      let receiptFoundInThread = false;

      const attachmentKeywords = [        'receipt',        'חשבונית',      ];

      for (const part of parts) {
        if (part.filename && part.filename.length > 0) {
          const fileName = part.filename;
          const normalizedFileName = fileName.toLowerCase();
          for (const keyword of attachmentKeywords) {
            if (normalizedFileName.includes(keyword.toLowerCase())) {
              receiptFoundInThread = true;
              break;
            }
          }
          if (receiptFoundInThread) {
            break;
          }
        }
      }

      for (const part of parts) {
        if (part.filename && part.filename.length > 0) {
          const attachmentId = part.body && part.body.attachmentId;
          if (!attachmentId) continue;

          const contentType = part.mimeType;
          const fileName = part.filename;
          const normalizedFileName = fileName.toLowerCase();
          const isPDF = isPdfFile(contentType, fileName);

          if (isPDF) {
            if (subjectMatchesPositiveKeywords) {
              if (receiptFoundInThread) {
                let attachmentMatchesKeyword = false;
                for (const keyword of attachmentKeywords) {
                  if (normalizedFileName.includes(keyword.toLowerCase())) {
                    attachmentMatchesKeyword = true;
                    break;
                  }
                }
                if (attachmentMatchesKeyword) {
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
                  console.log('Skipping non-receipt PDF in message with receipt:', fileName);
                }
              } else {
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
              let attachmentMatchesKeyword = false;
              for (const keyword of attachmentKeywords) {
                if (normalizedFileName.includes(keyword.toLowerCase())) {
                  attachmentMatchesKeyword = true;
                  break;
                }
              }
              if (attachmentMatchesKeyword) {
                const attachment = await gmail.users.messages.attachments.get({
                  userId: 'me',
                  messageId: messageData.id,
                  id: attachmentId,
                });

                const data = attachment.data.data;
                const buffer = Buffer.from(data, 'base64');

                const filePath = path.join(folderPath, sanitize(fileName));
                fs.writeFileSync(filePath, buffer);
                console.log(`Saved PDF attachment based on filename keyword: ${filePath}`);
              } else {
                console.log('Skipping attachment as it does not match criteria:', fileName);
              }
            }
          } else {
            console.log('Skipping non-PDF attachment:', fileName);
          }
        }
      }
    } else {
      console.log('No attachments found in message:', subject);
    }
  }

  return folderPath;
}

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

app.get('/logout', (req, res) => {
  req.session.destroy();
  res.redirect('/');
});

// Last results route
app.get('/last-results', (req, res) => {
  const lastResults = req.session.lastResults;
  if (!lastResults) {
    return res.send('No recent results found.');
  }

  res.send(`
    <h1>Last Processed Files</h1>
    <p>Processed at: ${new Date(lastResults.timestamp).toLocaleString()}</p>
    <p><a href="${lastResults.excelUrl}">Download Excel</a></p>
    <p><a href="${lastResults.zipUrl}">Download ZIP</a></p>
  `);
});

const PORT = process.env.PORT || 8080;
app.listen(PORT, () => {
  console.log(`Server is running on port ${PORT}`);
});


