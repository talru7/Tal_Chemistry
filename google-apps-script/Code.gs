const SPREADSHEET_ID = '168Zer_Rqm9jMccYOPpDJLun5dczeOeRrOz77ZvqJyrI';
const SHEET_NAME = 'פניות מהאתר';
const NOTIFICATION_EMAIL = 'talru7@gmail.com';

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const data = e && e.parameter ? e.parameter : {};

    if (data.website) {
      return response_({ ok: true });
    }

    const name = clean_(data.name, 80);
    const phone = clean_(data.phone, 30);
    const subject = clean_(data.subject || data.course, 120);
    const message = clean_(data.message, 1000);
    const submittedId = clean_(data.leadId, 36);
    const leadId = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(submittedId)
      ? submittedId : Utilities.getUuid();

    if (!name || !phone) {
      return response_({ ok: false, error: 'missing_required_fields' });
    }

    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    if (!sheet) {
      throw new Error('Lead sheet not found');
    }

    if (findLead_(sheet, leadId)) {
      return response_({ ok: true, leadId });
    }
    const row = firstEmptyRow_(sheet);
    sheet.getRange(row, 1, 1, 8).setValues([[
      new Date(),
      name,
      phone,
      subject,
      message,
      'חדש',
      'אתר',
      leadId,
    ]]);

    const spreadsheetUrl = `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/edit`;
    const emailBody = [
      'התקבלה פנייה חדשה מעמוד הנחיתה.',
      '',
      `שם: ${name}`,
      `טלפון: ${phone}`,
      `קורס או נושא: ${subject || 'לא צוין'}`,
      `הודעה: ${message || 'לא נכתבה הודעה'}`,
      '',
      `לצפייה בגיליון: ${spreadsheetUrl}`,
      `מזהה פנייה: ${leadId}`,
    ].join('\n');

    MailApp.sendEmail({
      to: NOTIFICATION_EMAIL,
      subject: `פנייה חדשה מהאתר: ${name}`,
      body: emailBody,
      name: 'עמוד הנחיתה של טל רובין',
    });

    return response_({ ok: true, leadId });
  } catch (error) {
    console.error(error);
    return response_({ ok: false, error: 'server_error' });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  const data = e && e.parameter ? e.parameter : {};
  const leadId = clean_(data.leadId, 36);
  const callback = String(data.callback || '');
  const validId = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(leadId);
  const received = validId && findLead_(
    SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME), leadId
  );
  const payload = JSON.stringify({ received: Boolean(received) });

  if (/^[A-Za-z_$][\w$]{0,80}$/.test(callback)) {
    return ContentService.createTextOutput(`${callback}(${payload});`)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(payload)
    .setMimeType(ContentService.MimeType.JSON);
}

function findLead_(sheet, leadId) {
  if (!sheet || sheet.getLastRow() < 2) return false;
  return Boolean(sheet.getRange(2, 8, sheet.getLastRow() - 1, 1)
    .createTextFinder(leadId).matchEntireCell(true).findNext());
}

function firstEmptyRow_(sheet) {
  const values = sheet.getRange(2, 2, Math.max(sheet.getMaxRows() - 1, 1), 1).getDisplayValues();
  const emptyIndex = values.findIndex(([value]) => !value);
  if (emptyIndex !== -1) return emptyIndex + 2;

  sheet.insertRowsAfter(sheet.getMaxRows(), 100);
  return sheet.getMaxRows() - 99;
}

function clean_(value, maxLength) {
  return String(value || '').trim().replace(/[<>]/g, '').slice(0, maxLength);
}

function response_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
