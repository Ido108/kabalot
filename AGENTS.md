# Repository Guidelines

## Project Structure & Module Organization
`server.js` is the single Express entry point: it wires uploads, Gmail OAuth, PDF/Excel processing, and download endpoints. Static client assets live under `public/` (`public/index.html` is the drag-and-drop dashboard, `public/gmail*.html` guide the Gmail auth flow, and `public/styles.css` plus `public/custom.css` contain the UI skin). Server-rendered responses are kept in `views/` (`process-gmail.ejs`, `result.ejs`) so keep shared fragments here instead of duplicating markup inside `server.js`.

## Build, Test, and Development Commands
- `npm install`: install dependencies declared in `package.json` (Node 18.x / npm 6.x as pinned in `engines`).
- `npm start`: runs `node server.js`, starting the API and static server (PORT defaults to 8080; override by exporting `PORT=...`).
- `node server.js --inspect`: optional when you need a debugger attached from VS Code/Chrome.

### Runtime Restrictions
- Never run `npm start`, `npm run <script>`, or any other command that starts the application locally unless the user explicitly instructs you to do so for that session.
- When verification is needed, describe the steps the user can run instead of executing the runtime commands yourself.

## Coding Style & Naming Conventions
Use CommonJS modules and 2-space indentation to match the existing files. Prefer `const` for dependencies and immutable helpers, `let` only when mutation is required. Route handlers live directly in `server.js`; group new middleware and helpers near the existing sections (authentication, archiving, email). Name EJS templates and HTML files with kebab-case verbs that describe the view (`process-gmail.ejs`, `gmail-auth.html`). Front-end scripts and styles belong in `public/`; keep filenames lower-case with hyphens.

## Testing Guidelines
There is no automated test suite yet. Exercise endpoints manually after each change:
1. Run `npm start`, visit `http://localhost:8080/`, upload a sample archive, and watch `/upload-progress`.
2. Trigger the Gmail flow via `/gmail-form`; verify `/gmail-progress` reflects status and that `/download/:filename` returns the processed archive.
Capture console logs for regressions and document edge cases in PRs until Jest or integration tests are introduced.

## Commit & Pull Request Guidelines
Existing commits (for example `Update package-lock.json`) use short, imperative subjects; follow that style and keep bodies concise but specific about files touched. Reference issue IDs when available (`Fix: handle oversized PDFs (#42)`). Each PR should include:
- Summary of changes and affected routes/views.
- Steps to reproduce/verify (commands + sample inputs).
- Screenshots or console excerpts when UI/upload behavior changes.

## Configuration & Security Notes
All credentials live in a `.env` file that defines `EMAIL_USER`, `EMAIL_PASS`, `GOOGLE_CLIENT_ID/SECRET`, and `SESSION_SECRET`; never commit that file or production archives. Use separate OAuth credentials for local testing and redact user data before attaching logs or test files to reviews.
