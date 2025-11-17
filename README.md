# Receipt Cloud – Expense Automation Suite

Modernized Express application that ingests receipts from manual uploads or Gmail, uses Google Document AI for bilingual (Hebrew/English) OCR, and returns structured Excel/ZIP bundles with automatic email delivery.

## Features

- **Parallel OCR pipeline** – files are processed concurrently with caching to avoid duplicate uploads and to keep latency predictable.
- **Bilingual intelligence** – Document AI requests include Hebrew/English hints and every parsed file is annotated with detected language, direction (RTL/LTR), and inferred currency.
- **Smart Gmail flow** – the authenticated Gmail account is detected automatically, files are downloaded for a given date range, optional attachments can be merged, and results are sent back to the same inbox without an extra email prompt.
- **New dashboard** – redesigned RTL-friendly UI with live timelines, queue indicators, and auto-generated download buttons.

## Local Development

> **Important:** per project policy Codex agents must never run `npm start`, `npm run <script>`, or any command that launches the app without explicit user approval. Instead, list the commands so the user can execute them manually.

1. Install dependencies once:
   ```bash
   npm install
   ```
2. Provide the required environment variables (see below) in a local `.env`.
3. Start the development server manually (user action only):
   ```bash
   npm start
   ```

### Environment Variables

| Name | Description |
| --- | --- |
| `PORT` | HTTP port (defaults to `8080`). |
| `BASE_URL` | Public URL used to build download links. |
| `EMAIL_USER` / `EMAIL_PASS` | SMTP credentials for nodemailer. |
| `SERVICE_ACCOUNT_BASE64` | Base64 encoded Google service account JSON for Document AI. |
| `DOCUMENT_AI_PROJECT_ID`, `DOCUMENT_AI_LOCATION`, `DOCUMENT_AI_PROCESSOR_ID` | Document AI identifiers. |
| `GMAIL_CLIENT_ID`, `GMAIL_CLIENT_SECRET`, `GMAIL_REDIRECT_URI` | OAuth2 credentials for Gmail ingestion. |
| `SESSION_SECRET` | Secret used by `express-session`. |
| `PDFCO_API_KEY`, `PASSWORD_PROTECTED_PDF_PASSWORD` | Optional PDF unlocking service settings. |

### Scripts

| Command | Description |
| --- | --- |
| `npm start` | Runs `node server.js`. **Only the user should run this command.** |

## Frontend Overview

The UI lives in `public/index.html` with styling in `public/custom.css`. It contains:

- Manual upload workflow (name/email/id number + drag/drop input, SSE timeline, download buttons).
- Gmail workflow (auth badge, profile avatar, date filters, optional extra files, auto-send notice).
- Shared timeline renderer that consumes the `/upload-progress` and `/gmail-progress` SSE endpoints.

## Backend Overview

- `server.js` combines Express routes, Google Document AI calls, Gmail OAuth, Excel generation, and download endpoints.
- File processing uses a concurrency limiter (`MAX_CONCURRENT_FILE_PROCESSING`, default 3) with checksum-based caching.
- The Gmail route fetches the authenticated profile once, stores it in the session, and uses it both for UI state and for email delivery.

## Manual Verification

1. Start the server (`npm start`) locally.
2. Navigate to `http://localhost:8080/`.
3. Upload a few PDFs/images and watch the progress timeline update; confirm Excel/ZIP download buttons work.
4. Authenticate via Gmail, pick a date range, and ensure the page shows the connected account plus auto-send notice.
5. Check that received emails contain both download links and expire after an hour.

Record any console logs or API failures in `Handoff.md` for the next contributor.*** End Patch
