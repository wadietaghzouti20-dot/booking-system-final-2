# RENDEVO — Réservation en ligne pour cabinets

RENDEVO is a responsive, French-language appointment-booking system for clinics and professional practices, built with GitHub Pages, Google Apps Script, Google Sheets, and Google Calendar. The clinic name remains configurable so the booking experience can be presented under each customer's practice name.

## Live website

https://wadietaghzouti20-dot.github.io/booking-system-final-2/

## Backend features

The complete Apps Script source is in [Code.gs](./Code.gs). It includes:

- Appointment validation and server-side prevention of overlapping confirmed bookings.
- Booking storage in the existing `Demandes` sheet; existing rows are preserved.
- Confirmation email to the patient with appointment details and a secure cancellation button.
- Admin notification with patient contact details, service, date, time, status, reference, and message.
- Appointment creation in the primary Google Calendar, with a guest invitation when the patient supplies an email address.
- One-off reminder trigger scheduled for 24 hours before the appointment, with a cancellation button in the email.
- Secure cancellation using a random one-time token.
- On cancellation: status becomes `Annulée`, the reminder trigger is deleted, the calendar event is removed where accessible, cancellation emails are attempted, and the slot becomes available again.
- Formula-injection protection for user-entered spreadsheet values.

## Required activation steps

**Updating GitHub does not automatically update the live Apps Script web app.**

1. Open the Apps Script project that owns the web-app URL in `config.js`.
2. Replace all contents of `Code.gs` with the current [Code.gs file](https://github.com/wadietaghzouti20-dot/booking-system-final-2/blob/main/Code.gs) and save.
3. Run `setupBookingSystem()` once and approve Google permissions for Sheets, Gmail, Calendar, and time-driven triggers.
4. Check the execution log for the spreadsheet URL and the detected admin notification email.
5. Open **Deploy → Manage deployments → Edit → New version → Deploy**. Keep the existing web-app URL.
6. Submit a future test booking using an email address you control. Check both inboxes, Google Calendar, the sheet row, and the trigger list.
7. Test cancellation using the button in the client confirmation email. Verify the status becomes `Annulée`, the event/reminder is removed, and the same time appears again in availability.

## Admin email

The script first uses `OWNER_EMAIL` in `Code.gs` if configured. If it is blank, it attempts to use the Google account executing the web app. If admin notifications must go to a different mailbox, set `OWNER_EMAIL` to that address before deploying.

## Important limitations

- Apps Script and Google Calendar must be authorized by the account that owns the deployment.
- Google may delay time-driven triggers; a reminder cannot be guaranteed to arrive at the exact second 24 hours before the appointment.
- Appointments made less than 24 hours ahead cannot receive an advance 24-hour reminder.
- The booking form requires a valid patient email so the confirmation, reminder, and calendar invitation can be sent.
- Email and calendar operations can fail independently of saving a booking; inspect Apps Script **Executions** if a test message or event is missing.
- Test with dummy data before using real patient information. Avoid collecting detailed medical records in this basic booking system.

## Cancellation page

https://wadietaghzouti20-dot.github.io/booking-system-final-2/cancel.html
