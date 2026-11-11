# Backend setup – Google Apps Script

This folder contains the starter API for the Attendance Rakah project.

## Important deployment boundary
GitHub Pages is a static frontend. This Apps Script project is a separate backend. Do not put spreadsheet IDs, passwords, session tokens, or API secrets in the public repository or frontend source.

## Setup
1. Create a new Google Apps Script project at https://script.google.com.
2. Copy the contents of `Code.gs` into the Apps Script editor.
3. Save, then run `setupAttendanceDatabase` once and approve the requested Google permissions.
4. Open the execution log. Save the generated initial admin password securely; the username is `admin`. Do not share the execution log publicly.
5. Deploy → New deployment → Web app. Execute as: Me. Who has access: Anyone (the app enforces its own username/password session checks). Keep the deployment URL private until the frontend integration is completed.
6. Test the deployment URL with a GET request; it should return the service status.

## Current API actions
- `login`
- `bootstrapStatus`
- `dashboard`
- `listEmployees`
- `saveEmployee`
- `deactivateEmployee`
- `listAttendance`
- `savePunch`
- `getSettings`
- `saveSettings`
- `listUsers`
- `createUser`

POST payloads use JSON with an `action` field. Authenticated actions also require `token`.

## Security and production checklist
- This is a starter implementation, not yet production-ready.
- Restrict the web app deployment audience if the organization has a suitable Google Workspace configuration.
- Add rate limiting, login lockout, password reset, token revocation, and stronger session storage before production.
- Review Google Apps Script quotas and execution limits.
- Validate all attendance policy rules server-side, including duplicate punches, shift boundaries, leave, holidays, and edits.
- Test cross-origin browser requests from GitHub Pages. Apps Script web-app redirects and browser CORS behavior can require an integration adapter; do not assume a successful GET from the browser means POST works.
- The initial password is printed once in the Apps Script execution log. Change it after first login using a future password-management workflow; do not reuse it elsewhere.
- Do not store fingerprint templates or face biometric templates in this spreadsheet. Store only attendance event metadata and the device's employee identifier.

## Spreadsheet tabs
`employees`, `attendance`, `users`, `settings`, `audit`.
