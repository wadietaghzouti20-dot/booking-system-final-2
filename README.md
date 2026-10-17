# Booking System Final 2

Responsive French-language clinic booking website with a Google Apps Script + Google Sheets backend.

## Files

- `index.html`, `styles.css`, `app.js` — patient booking flow.
- `config.js` — public Apps Script web-app URL and clinic display settings.
- `cancel.html` — patient cancellation confirmation page.
- `Code.gs` — Apps Script backend source.

## Backend features in `Code.gs`

- Server-side form validation and slot-conflict prevention.
- Booking storage in the existing `Demandes` sheet.
- Automatic addition of the cancellation/reminder columns without clearing existing rows.
- A random, one-time cancellation token (a booking reference alone cannot cancel an appointment).
- A one-off time trigger targeting 24 hours before the appointment; Google may run time triggers slightly later than the target.
- Reminder email with a cancellation link, if the patient provided an email and the trigger was successfully created.
- Cancelling a booking changes its status to `Annulée` and removes its pending reminder trigger.

## Important: GitHub code and Apps Script are separate

Committing `Code.gs` here does **not** update the live Google Apps Script deployment automatically. The latest backend source must be copied into the Apps Script editor and redeployed before the new cancellation and reminder features work.

### Activate the updated backend

1. Open the Apps Script project that owns the web-app URL in `config.js`.
2. Replace the project code with the complete `Code.gs` from this repository and save it.
3. Run `setupBookingSystem()` once in the editor and approve the requested permissions. It adds the new columns while preserving existing booking rows.
4. Open **Deploy → Manage deployments**, edit the existing web-app deployment, choose **New version**, and deploy. Keep the same web-app URL so `config.js` does not need to change.
5. Test with a future appointment using an email address you control. Verify the sheet row, reminder trigger, reminder email, and cancellation link before using real patient information.

**Reminder timing:** the trigger is scheduled for the appointment's date/time minus exactly 24 hours in the Algeria UTC+1 time zone. Google controls the actual execution time and can delay it; no Apps Script trigger can guarantee delivery at an exact second. Appointments made less than 24 hours ahead cannot receive a 24-hour advance reminder. A valid email address is required.

## Publish the website

In GitHub, open **Settings → Pages**, choose **Deploy from a branch**, select `main` and `/(root)`, then save. Once published, the cancellation page URL is:

`https://wadietaghzouti20-dot.github.io/booking-system-final-2/cancel.html`

Do not assume this URL is live until GitHub Pages is enabled and the page opens successfully.

## Privacy and safety

- Test using dummy data before handling real patient information.
- Do not store detailed medical records or symptoms in this basic booking sheet.
- The cancellation link contains a private one-time token; do not share it.
- The public `config.js` must never contain credentials or private secrets.
