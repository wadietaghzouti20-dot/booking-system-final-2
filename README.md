# Booking System Final 2

Responsive, French-language clinic appointment request frontend.

## Current files

- `index.html` — accessible, five-step appointment request form.
- `styles.css` — responsive layout for mobile, tablet, and desktop.
- `app.js` — step navigation, client-side validation, review, and submission handling.
- `config.js` — public clinic settings and Apps Script endpoint placeholder.

## Current status

The frontend files are in place. **Appointment submission is intentionally disabled until a valid Google Apps Script web-app URL is configured.** The interface will show an explicit error rather than pretending a request was sent.

## Connect Google Apps Script later

1. Deploy the Apps Script project as a web app.
2. Set access according to the intended patient-facing workflow and test it with a non-sensitive test request.
3. Copy the deployed `/exec` URL into `appsScriptUrl` in `config.js`.
4. Implement server-side validation, duplicate/conflict prevention, spreadsheet storage, and optional email notification in Apps Script.
5. Test a successful request, an invalid request, a duplicate/conflicting slot, and an unavailable backend before publishing.

## Publish with GitHub Pages

In the repository, open **Settings → Pages**, choose **Deploy from a branch**, select `main` and `/(root)`, then save. GitHub Pages may take a few minutes to publish.

## Important

This is an appointment *request* form, not a real-time confirmed calendar yet. Do not enter real patient data until the backend, access controls, and privacy workflow have been configured and tested. The public `config.js` must never contain credentials or private secrets.
