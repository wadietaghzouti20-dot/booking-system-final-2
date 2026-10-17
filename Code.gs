/**
 * Booking System Final 2 — Google Apps Script backend
 * Time zone: Africa/Algiers
 *
 * SETUP:
 * 1. Run setupBookingSystem() once from the Apps Script editor.
 * 2. Approve Google's requested permissions.
 * 3. Open the execution log and copy the spreadsheet URL.
 * 4. Deploy as a Web app (execute as you; access: anyone).
 *
 * Do not store medical records or detailed symptoms in this basic booking sheet.
 */

const BOOKING_SETTINGS = Object.freeze({
  TIME_ZONE: "Africa/Algiers",
  SPREADSHEET_PROPERTY: "BOOKING_SPREADSHEET_ID",
  SHEET_NAME: "Demandes",
  OWNER_EMAIL: "", // Optional: add the clinic's notification email here.
  ALLOWED_SERVICES: [
    "Consultation générale",
    "Contrôle de routine",
    "Suivi / traitement",
    "Autre demande"
  ],
  SERVICE_DURATIONS_MINUTES: {
    "Consultation générale": 30,
    "Contrôle de routine": 20,
    "Suivi / traitement": 45,
    "Autre demande": 30
  },
  MAX_NAME_LENGTH: 100,
  MAX_PHONE_LENGTH: 25,
  MAX_EMAIL_LENGTH: 150,
  MAX_NOTES_LENGTH: 1000
});

const BOOKING_HEADERS = [
  "Référence",
  "Créée le",
  "Statut",
  "Motif",
  "Date souhaitée",
  "Moment préféré",
  "Date confirmée",
  "Heure confirmée",
  "Nom et prénom",
  "Téléphone",
  "E-mail",
  "Message",
  "Notes internes",
  "Jeton annulation",
  "Rappel envoyé le",
  "ID déclencheur rappel"
];

/**
 * Run once manually. Creates the spreadsheet and stores its ID in script
 * properties so the web app can find it without exposing the ID to patients.
 */
function setupBookingSystem() {
  const props = PropertiesService.getScriptProperties();
  let spreadsheetId = props.getProperty(BOOKING_SETTINGS.SPREADSHEET_PROPERTY);
  let spreadsheet;

  if (spreadsheetId) {
    try {
      spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    } catch (error) {
      spreadsheet = null;
    }
  }

  if (!spreadsheet) {
    spreadsheet = SpreadsheetApp.create("Booking System Final 2 — Demandes de rendez-vous");
    props.setProperty(BOOKING_SETTINGS.SPREADSHEET_PROPERTY, spreadsheet.getId());
  }

  spreadsheet.setSpreadsheetTimeZone(BOOKING_SETTINGS.TIME_ZONE);
  let sheet = spreadsheet.getSheetByName(BOOKING_SETTINGS.SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(BOOKING_SETTINGS.SHEET_NAME);

  ensureBookingColumns_(sheet);
  if (sheet.getLastRow() === 1) {
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, BOOKING_HEADERS.length).setFontWeight("bold");
    sheet.getRange("B:B").setNumberFormat("yyyy-mm-dd hh:mm:ss");
    sheet.getRange("E:E").setNumberFormat("@");
    sheet.getRange("G:G").setNumberFormat("@");
    sheet.getRange("O:O").setNumberFormat("yyyy-mm-dd hh:mm:ss");
    sheet.autoResizeColumns(1, BOOKING_HEADERS.length);
  }

  Logger.log("Spreadsheet created/ready: " + spreadsheet.getUrl());
  return spreadsheet.getUrl();
}

function doGet(e) {
  return jsonResponse_({
    success: true,
    service: "Booking System Final 2",
    message: "Le service de réservation est disponible."
  });
}

function doPost(e) {
  try {
    if (!e || !e.postData || typeof e.postData.contents !== "string") {
      return jsonResponse_({ success: false, message: "Requête invalide." });
    }

    const raw = JSON.parse(e.postData.contents);
    if (raw && raw.action === "cancel") {
      const cancellation = cancelAppointment_(raw.reference, raw.token);
      return jsonResponse_({ success: true, message: cancellation });
    }
    if (raw && raw.action === "availability") {
      const availableSlots = getAvailableSlots_(raw.preferredDate, raw.service);
      return jsonResponse_({
        success: true,
        availableSlots: availableSlots,
        message: availableSlots.length
          ? "Choisissez un créneau disponible."
          : "Aucun créneau disponible à cette date. Choisissez une autre date."
      });
    }

    const data = validateRequest_(raw);
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);

    let reference;
    try {
      const sheet = getBookingSheet_();
      assertSlotCanBeBooked_(sheet, data.preferredDate, data.preferredTime, data.service, "");
      reference = "RDV-" + Utilities.getUuid().replace(/-/g, "").slice(0, 8).toUpperCase();
      const cancellationToken = Utilities.getUuid().replace(/-/g, "") + Utilities.getUuid().replace(/-/g, "");
      const now = new Date();

      sheet.appendRow([
        reference,
        now,
        "Confirmé",
        data.service,
        data.preferredDate,
        data.preferredTime,
        data.preferredDate,
        data.preferredTime,
        safeCell_(data.fullName),
        safeCell_(data.phone),
        safeCell_(data.email),
        safeCell_(data.notes),
        "",
        cancellationToken,
        "",
        ""
      ]);
      try {
        scheduleAppointmentReminder_(reference, cancellationToken, data.email, data.preferredDate, data.preferredTime, data.service, data.fullName);
      } catch (reminderError) {
        // Keep the confirmed booking successful if Google rejects a reminder trigger.
        console.error("Could not schedule 24-hour reminder for " + reference + ": " + reminderError);
      }
    } finally {
      lock.releaseLock();
    }

    notifyClinic_(reference, data);
    return jsonResponse_({
      success: true,
      reference: reference,
      status: "Confirmé",
      confirmedDate: data.preferredDate,
      confirmedTime: data.preferredTime,
      message: "Votre rendez-vous est confirmé le " + data.preferredDate + " à " + data.preferredTime + "."
    });
  } catch (error) {
    const message = error && error.message ? error.message : "";
    return jsonResponse_({
      success: false,
      message: message || "Une erreur est survenue. Veuillez vérifier les informations et réessayer."
    });
  }
}

function validateRequest_(input) {
  if (!input || typeof input !== "object" || Array.isArray(input)) {
    throw new Error("Les informations envoyées sont invalides.");
  }

  const service = cleanText_(input.service, 80);
  const preferredDate = cleanText_(input.preferredDate, 10);
  const preferredTime = cleanText_(input.preferredTime, 30);
  const fullName = cleanText_(input.fullName, BOOKING_SETTINGS.MAX_NAME_LENGTH);
  const phone = cleanText_(input.phone, BOOKING_SETTINGS.MAX_PHONE_LENGTH);
  const email = cleanText_(input.email, BOOKING_SETTINGS.MAX_EMAIL_LENGTH);
  const notes = cleanText_(input.notes, BOOKING_SETTINGS.MAX_NOTES_LENGTH);

  if (BOOKING_SETTINGS.ALLOWED_SERVICES.indexOf(service) === -1) {
    throw new Error("Veuillez choisir un motif de rendez-vous valide.");
  }
  if (!/^\d{4}-\d{2}-\d{2}$/.test(preferredDate) || !isRealDate_(preferredDate)) {
    throw new Error("Veuillez choisir une date valide.");
  }
  const today = Utilities.formatDate(new Date(), BOOKING_SETTINGS.TIME_ZONE, "yyyy-MM-dd");
  if (preferredDate < today) {
    throw new Error("La date choisie est déjà passée. Veuillez en sélectionner une autre.");
  }
  if (!/^([01]\d|2[0-3]):([0-5]\d)$/.test(preferredTime)) {
    throw new Error("Veuillez choisir une heure de rendez-vous valide parmi les créneaux proposés.");
  }
  if (fullName.length < 2) {
    throw new Error("Veuillez saisir votre nom et prénom.");
  }
  const digits = phone.replace(/\D/g, "");
  if (digits.length < 8 || digits.length > 15) {
    throw new Error("Veuillez saisir un numéro de téléphone valide.");
  }
  if (email && (email.length > BOOKING_SETTINGS.MAX_EMAIL_LENGTH ||
      !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email))) {
    throw new Error("Veuillez vérifier votre adresse e-mail.");
  }
  if (notes.length > BOOKING_SETTINGS.MAX_NOTES_LENGTH) {
    throw new Error("Le message ne doit pas dépasser 1 000 caractères.");
  }
  if (input.privacyConsent !== true) {
    throw new Error("Veuillez accepter l’utilisation de vos informations pour traiter la demande.");
  }

  return {
    service: service,
    preferredDate: preferredDate,
    preferredTime: preferredTime,
    fullName: fullName,
    phone: phone,
    email: email,
    notes: notes
  };
}

/**
 * Return exact available start times for a date and service.
 * All reads use the same script lock as booking writes to avoid inconsistent checks.
 */
function getAvailableSlots_(dateValue, serviceValue) {
  const date = cleanText_(dateValue, 10);
  const service = cleanText_(serviceValue, 80);

  if (BOOKING_SETTINGS.ALLOWED_SERVICES.indexOf(service) === -1) {
    throw new Error("Veuillez choisir un motif de rendez-vous valide.");
  }
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || !isRealDate_(date)) {
    throw new Error("Veuillez choisir une date valide.");
  }
  const today = Utilities.formatDate(new Date(), BOOKING_SETTINGS.TIME_ZONE, "yyyy-MM-dd");
  if (date < today) {
    throw new Error("La date choisie est déjà passée. Veuillez en sélectionner une autre.");
  }

  const schedule = getScheduleForDate_(date);
  if (!schedule.open) return [];

  const duration = BOOKING_SETTINGS.SERVICE_DURATIONS_MINUTES[service] || 30;
  let firstStart = schedule.openingMinutes;
  if (date === today) {
    const nowText = Utilities.formatDate(new Date(), BOOKING_SETTINGS.TIME_ZONE, "HH:mm").split(":").map(Number);
    const currentMinutes = nowText[0] * 60 + nowText[1];
    firstStart = Math.max(firstStart, Math.ceil((currentMinutes + 15) / 15) * 15);
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sheet = getBookingSheet_();
    const values = sheet.getDataRange().getValues();
    const slots = [];
    for (let start = firstStart; start + duration <= schedule.closingMinutes; start += 15) {
      if (!hasBookingConflict_(values, date, start, duration, "")) {
        slots.push(minutesToTime_(start));
      }
    }
    return slots;
  } finally {
    lock.releaseLock();
  }
}

function getScheduleForDate_(date) {
  const day = new Date(date + "T12:00:00");
  const weekday = Number(Utilities.formatDate(day, BOOKING_SETTINGS.TIME_ZONE, "u"));
  if (weekday === 5) return { open: false, openingMinutes: 0, closingMinutes: 0 };
  if (weekday === 6) return { open: true, openingMinutes: 10 * 60, closingMinutes: 16 * 60 };
  return { open: true, openingMinutes: 8 * 60, closingMinutes: 17 * 60 };
}

function assertSlotCanBeBooked_(sheet, date, time, service, ignoredReference) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || !isRealDate_(date)) {
    throw new Error("Date de rendez-vous invalide.");
  }
  const today = Utilities.formatDate(new Date(), BOOKING_SETTINGS.TIME_ZONE, "yyyy-MM-dd");
  if (date < today) throw new Error("La date choisie est déjà passée. Veuillez en sélectionner une autre.");

  if (!/^([01]\d|2[0-3]):([0-5]\d)$/.test(time)) {
    throw new Error("Veuillez choisir une heure de rendez-vous valide.");
  }
  const parts = time.split(":").map(Number);
  const start = parts[0] * 60 + parts[1];
  if (parts[1] % 15 !== 0) {
    throw new Error("Choisissez un début de rendez-vous sur un quart d'heure.");
  }
  if (date === today) {
    const nowParts = Utilities.formatDate(new Date(), BOOKING_SETTINGS.TIME_ZONE, "HH:mm").split(":").map(Number);
    const currentMinutes = nowParts[0] * 60 + nowParts[1];
    const earliestStart = Math.ceil((currentMinutes + 15) / 15) * 15;
    if (start < earliestStart) {
      throw new Error("Ce créneau n’est plus disponible. Choisissez une heure future.");
    }
  }

  const schedule = getScheduleForDate_(date);
  if (!schedule.open) throw new Error("Le cabinet est fermé le vendredi.");
  const duration = BOOKING_SETTINGS.SERVICE_DURATIONS_MINUTES[service] || 30;
  const end = start + duration;
  if (start < schedule.openingMinutes || end > schedule.closingMinutes) {
    throw new Error("Ce créneau ne respecte pas les horaires du cabinet (" +
      minutesToTime_(schedule.openingMinutes) + "–" + minutesToTime_(schedule.closingMinutes) + ").");
  }

  const values = sheet.getDataRange().getValues();
  if (hasBookingConflict_(values, date, start, duration, ignoredReference)) {
    throw new Error("Ce créneau vient d'être réservé. Veuillez choisir un autre horaire.");
  }
}

function hasBookingConflict_(values, date, startMinutes, duration, ignoredReference) {
  if (!values || values.length < 2) return false;
  const headers = values[0];
  const col = {};
  headers.forEach(function (name, index) { col[name] = index; });

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (String(row[col["Référence"]]) === ignoredReference) continue;
    if (String(row[col["Statut"]]) !== "Confirmé") continue;
    if (normalizeDate_(row[col["Date confirmée"]]) !== date) continue;

    const existingMatch = normalizeTime_(row[col["Heure confirmée"]]).match(/^(\d{2}):(\d{2})$/);
    if (!existingMatch) continue;
    const existingStart = Number(existingMatch[1]) * 60 + Number(existingMatch[2]);
    const existingService = String(row[col["Motif"]] || "");
    const existingDuration = BOOKING_SETTINGS.SERVICE_DURATIONS_MINUTES[existingService] || 30;
    const existingEnd = existingStart + existingDuration;
    if (startMinutes < existingEnd && existingStart < startMinutes + duration) return true;
  }
  return false;
}

/**
 * Confirm a request against the clinic's weekly schedule and service duration.
 * Hours: Sunday-Thursday 08:00-17:00; Friday closed; Saturday 10:00-16:00.
 * Starts must be on 15-minute boundaries. Overlapping confirmed appointments
 * are rejected using each service's configured duration.
 */
function confirmAppointment(reference, confirmedDate, confirmedTime) {
  reference = cleanText_(reference, 30);
  confirmedDate = cleanText_(confirmedDate, 10);
  confirmedTime = cleanText_(confirmedTime, 5);

  if (!/^\d{4}-\d{2}-\d{2}$/.test(confirmedDate) || !isRealDate_(confirmedDate)) {
    throw new Error("Date de confirmation invalide (format attendu : AAAA-MM-JJ).");
  }
  if (!/^([01]\d|2[0-3]):([0-5]\d)$/.test(confirmedTime)) {
    throw new Error("Heure invalide. Utilisez le format HH:MM.");
  }

  const parts = confirmedTime.split(":").map(Number);
  const startMinutes = parts[0] * 60 + parts[1];
  if (parts[1] % 15 !== 0) {
    throw new Error("Choisissez un début de rendez-vous sur un quart d'heure (ex. 10:00, 10:15, 10:30).");
  }

  const day = new Date(confirmedDate + "T12:00:00");
  const weekday = Number(Utilities.formatDate(day, BOOKING_SETTINGS.TIME_ZONE, "u"));
  let openingMinutes;
  let closingMinutes;
  if (weekday === 5) {
    throw new Error("Le cabinet est fermé le vendredi.");
  } else if (weekday === 6) {
    openingMinutes = 10 * 60;
    closingMinutes = 16 * 60;
  } else {
    openingMinutes = 8 * 60;
    closingMinutes = 17 * 60;
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sheet = getBookingSheet_();
    const values = sheet.getDataRange().getValues();
    const headers = values[0];
    const col = {};
    headers.forEach(function (name, index) { col[name] = index; });

    let targetRow = -1;
    let targetData = null;
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][col["Référence"]]) === reference) {
        targetRow = i + 1;
        targetData = values[i];
        break;
      }
    }
    if (targetRow === -1) throw new Error("Référence introuvable.");

    const currentStatus = String(sheet.getRange(targetRow, col["Statut"] + 1).getValue());
    if (currentStatus === "Annulée") throw new Error("Cette demande a été annulée.");

    const service = String(targetData[col["Motif"]] || "");
    const duration = BOOKING_SETTINGS.SERVICE_DURATIONS_MINUTES[service] || 30;
    const endMinutes = startMinutes + duration;
    if (startMinutes < openingMinutes || endMinutes > closingMinutes) {
      const opening = minutesToTime_(openingMinutes);
      const closing = minutesToTime_(closingMinutes);
      throw new Error("Ce rendez-vous ne respecte pas les horaires du cabinet (" + opening + "–" + closing + ").");
    }

    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      if (String(row[col["Référence"]]) === reference) continue;
      if (String(row[col["Statut"]]) !== "Confirmé") continue;
      if (normalizeDate_(row[col["Date confirmée"]]) !== confirmedDate) continue;

      const existingStartText = normalizeTime_(row[col["Heure confirmée"]]);
      const existingMatch = existingStartText.match(/^(\d{2}):(\d{2})$/);
      if (!existingMatch) continue;
      const existingStart = Number(existingMatch[1]) * 60 + Number(existingMatch[2]);
      const existingService = String(row[col["Motif"]] || "");
      const existingDuration = BOOKING_SETTINGS.SERVICE_DURATIONS_MINUTES[existingService] || 30;
      const existingEnd = existingStart + existingDuration;

      if (startMinutes < existingEnd && existingStart < endMinutes) {
        throw new Error("Ce créneau chevauche un autre rendez-vous déjà confirmé.");
      }
    }

    sheet.getRange(targetRow, col["Date confirmée"] + 1).setNumberFormat("@").setValue(confirmedDate);
    sheet.getRange(targetRow, col["Heure confirmée"] + 1).setNumberFormat("@").setValue(confirmedTime);
    sheet.getRange(targetRow, col["Statut"] + 1).setValue("Confirmé");
    return "Rendez-vous confirmé : " + reference + " le " + confirmedDate + " à " + confirmedTime + " (" + duration + " min).";
  } finally {
    lock.releaseLock();
  }
}

function getBookingSheet_() {
  const id = PropertiesService.getScriptProperties().getProperty(BOOKING_SETTINGS.SPREADSHEET_PROPERTY);
  if (!id) throw new Error("Le système n’est pas initialisé. Exécutez setupBookingSystem() dans Apps Script.");
  const spreadsheet = SpreadsheetApp.openById(id);
  const sheet = spreadsheet.getSheetByName(BOOKING_SETTINGS.SHEET_NAME);
  if (!sheet) throw new Error("L’onglet des demandes est introuvable. Exécutez setupBookingSystem().");
  ensureBookingColumns_(sheet);
  return sheet;
}

function notifyClinic_(reference, data) {
  const recipient = String(BOOKING_SETTINGS.OWNER_EMAIL || "").trim();
  if (!recipient) return;
  try {
    const body = [
      "Nouveau rendez-vous confirmé automatiquement",
      "",
      "Référence : " + reference,
      "Motif : " + data.service,
      "Date confirmée : " + data.preferredDate,
      "Heure confirmée : " + data.preferredTime,
      "Nom : " + data.fullName,
      "Téléphone : " + data.phone,
      "E-mail : " + (data.email || "Non renseigné"),
      "Message : " + (data.notes || "Aucun")
    ].join("\n");
    MailApp.sendEmail({
      to: recipient,
      subject: "Rendez-vous confirmé — " + reference,
      body: body
    });
  } catch (error) {
    // Email notification is optional; a mail failure must not undo a saved request.
    console.error("Notification email failed: " + error);
  }
}

function jsonResponse_(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function cleanText_(value, maxLength) {
  if (typeof value !== "string" && typeof value !== "number") return "";
  return String(value).replace(/[\u0000-\u001F\u007F]/g, " ").trim().slice(0, maxLength);
}

function safeCell_(value) {
  // Prevent spreadsheet formula injection from user-submitted text.
  const text = String(value == null ? "" : value);
  return /^[=+@\-]/.test(text) ? "'" + text : text;
}

function isRealDate_(value) {
  const parts = value.split("-").map(Number);
  const date = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2]));
  return date.getUTCFullYear() === parts[0] &&
    date.getUTCMonth() === parts[1] - 1 &&
    date.getUTCDate() === parts[2];
}

function normalizeDate_(value) {
  if (value instanceof Date) return Utilities.formatDate(value, BOOKING_SETTINGS.TIME_ZONE, "yyyy-MM-dd");
  const text = String(value == null ? "" : value).trim();
  const match = text.match(/^(\d{4}-\d{2}-\d{2})/);
  return match ? match[1] : text;
}

function minutesToTime_(minutes) {
  return String(Math.floor(minutes / 60)).padStart(2, "0") + ":" +
    String(minutes % 60).padStart(2, "0");
}

function normalizeTime_(value) {
  if (value instanceof Date) return Utilities.formatDate(value, BOOKING_SETTINGS.TIME_ZONE, "HH:mm");
  const text = String(value == null ? "" : value).trim();
  const match = text.match(/^([0-2]?\d):([0-5]\d)/);
  return match ? String(match[1]).padStart(2, "0") + ":" + match[2] : text;
}


/**
 * Add new system columns without deleting or rewriting any existing booking data.
 */
function ensureBookingColumns_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, BOOKING_HEADERS.length).setValues([BOOKING_HEADERS]);
    return;
  }
  const existing = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getDisplayValues()[0];
  BOOKING_HEADERS.forEach(function (header) {
    if (existing.indexOf(header) === -1) {
      const newColumn = sheet.getLastColumn() + 1;
      sheet.getRange(1, newColumn).setValue(header).setFontWeight("bold");
      existing.push(header);
    }
  });
}

/**
 * Create a one-time reminder trigger for exactly 24 hours before the appointment.
 * Google may execute time-driven triggers a little later than the target time.
 */
function scheduleAppointmentReminder_(reference, token, email, date, time, service, fullName) {
  if (!email) return; // Email is optional in the patient form.
  const appointmentAt = appointmentDateTime_(date, time);
  const reminderAt = new Date(appointmentAt.getTime() - 24 * 60 * 60 * 1000);
  if (reminderAt.getTime() <= Date.now()) return; // Booking made less than 24 hours ahead.

  const trigger = ScriptApp.newTrigger("sendScheduledReminder")
    .timeBased()
    .at(reminderAt)
    .create();

  const sheet = getBookingSheet_();
  const row = findBookingRow_(sheet, reference);
  if (row < 2) return;
  const headers = getHeaderMap_(sheet);
  sheet.getRange(row, headers["ID déclencheur rappel"]).setValue(trigger.getUniqueId());
}

/** Trigger entry point. The trigger UID maps to one appointment in the sheet. */
function sendScheduledReminder(e) {
  const triggerUid = e && e.triggerUid ? String(e.triggerUid) : "";
  if (!triggerUid) return;
  const sheet = getBookingSheet_();
  const headers = getHeaderMap_(sheet);
  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (String(row[headers["ID déclencheur rappel"] - 1] || "") !== triggerUid) continue;

    const sheetRow = i + 1;
    const status = String(row[headers["Statut"] - 1] || "");
    const reference = String(row[headers["Référence"] - 1] || "");
    const email = String(row[headers["E-mail"] - 1] || "").trim();
    const token = String(row[headers["Jeton annulation"] - 1] || "");
    const date = normalizeDate_(row[headers["Date confirmée"] - 1]);
    const time = normalizeTime_(row[headers["Heure confirmée"] - 1]);
    const name = String(row[headers["Nom et prénom"] - 1] || "");
    const service = String(row[headers["Motif"] - 1] || "");

    if (status === "Confirmé" && email && token && date && time &&
        !row[headers["Rappel envoyé le"] - 1]) {
      const cancelUrl = getCancellationUrl_(reference, token);
      const appointmentLabel = date + " à " + time;
      MailApp.sendEmail({
        to: email,
        subject: "Rappel : votre rendez-vous est demain — " + reference,
        body: [
          "Bonjour " + name + ",",
          "",
          "Petit rappel : votre rendez-vous est prévu dans 24 heures.",
          "Date et heure : " + appointmentLabel,
          "Motif : " + service,
          "Référence : " + reference,
          "",
          "Si vous ne pouvez pas venir, ouvrez ce lien pour demander l'annulation :",
          cancelUrl,
          "",
          "Si vous souhaitez conserver le rendez-vous, aucune action n'est nécessaire.",
          "",
          "Cabinet médical"
        ].join("\\n")
      });
      sheet.getRange(sheetRow, headers["Rappel envoyé le"]).setValue(new Date());
    }
    sheet.getRange(sheetRow, headers["ID déclencheur rappel"]).clearContent();
    return;
  }
}

/** Patient cancellation requires both the booking reference and the unguessable token. */
function cancelAppointment_(referenceValue, tokenValue) {
  const reference = cleanText_(referenceValue, 30);
  const token = cleanText_(tokenValue, 100);
  if (!/^RDV-[A-F0-9]{8}$/.test(reference) || !/^[a-f0-9]{64}$/.test(token)) {
    throw new Error("Lien d’annulation invalide.");
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sheet = getBookingSheet_();
    const headers = getHeaderMap_(sheet);
    const row = findBookingRow_(sheet, reference);
    if (row < 2) throw new Error("Rendez-vous introuvable.");
    const savedToken = String(sheet.getRange(row, headers["Jeton annulation"]).getValue() || "");
    if (!savedToken || savedToken !== token) throw new Error("Lien d’annulation invalide ou expiré.");
    const statusCell = sheet.getRange(row, headers["Statut"]);
    const status = String(statusCell.getValue() || "");
    if (status === "Annulée") return "Ce rendez-vous est déjà annulé.";
    if (status !== "Confirmé") throw new Error("Ce rendez-vous ne peut plus être annulé en ligne.");

    statusCell.setValue("Annulée");
    sheet.getRange(row, headers["Jeton annulation"]).clearContent(); // one-time use
    const triggerId = String(sheet.getRange(row, headers["ID déclencheur rappel"]).getValue() || "");
    if (triggerId) {
      ScriptApp.getProjectTriggers().forEach(function (trigger) {
        if (trigger.getUniqueId() === triggerId) ScriptApp.deleteTrigger(trigger);
      });
      sheet.getRange(row, headers["ID déclencheur rappel"]).clearContent();
    }
    return "Votre rendez-vous a été annulé. Le cabinet en sera informé dans son planning.";
  } finally {
    lock.releaseLock();
  }
}

function appointmentDateTime_(date, time) {
  if (!/^\\d{4}-\\d{2}-\\d{2}$/.test(date) || !/^([01]\\d|2[0-3]):([0-5]\\d)$/.test(time) || !isRealDate_(date)) {
    throw new Error("Date ou heure de rendez-vous invalide.");
  }
  const parts = time.split(":").map(Number);
  // Use a fixed offset for Algeria (UTC+1) to avoid server-local timezone differences.
  return new Date(date + "T" + time + ":00+01:00");
}

function getCancellationUrl_(reference, token) {
  return "https://wadietaghzouti20-dot.github.io/booking-system-final-2/cancel.html?ref=" +
    encodeURIComponent(reference) + "&token=" + encodeURIComponent(token);
}

function getHeaderMap_(sheet) {
  const names = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const map = {};
  names.forEach(function (name, index) { if (name) map[name] = index + 1; });
  return map;
}

function findBookingRow_(sheet, reference) {
  const headers = getHeaderMap_(sheet);
  const refColumn = headers["Référence"];
  if (!refColumn || sheet.getLastRow() < 2) return -1;
  const found = sheet.getRange(2, refColumn, sheet.getLastRow() - 1, 1)
    .createTextFinder(reference).matchEntireCell(true).findNext();
  return found ? found.getRow() : -1;
}
