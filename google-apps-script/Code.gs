const SPREADSHEET_ID = '11C2FImTZcwBycpkAv-4CjiL_f8FO6nMVRfWF2XW8fIA';
const SHEET_NAME = 'Новые ответы сайта';
// Укажите адрес, на который присылать уведомления. Пока пусто — скрипт попробует
// определить аккаунт владельца веб-приложения автоматически.
const NOTIFICATION_EMAIL = '';

/**
 * Один раз запустите setupExistingSpreadsheet() вручную в редакторе Apps Script.
 * Создаст отдельную вкладку в существующей таблице, не затрагивая старые ответы.
 */
function setupExistingSpreadsheet() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
    sheet.appendRow(['Дата и время ответа', 'Имя и фамилия', 'Присутствие', 'Комментарий']);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
    sheet.autoResizeColumns(1, 4);
  }
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', SPREADSHEET_ID);
  Logger.log('Готово. Вкладка: ' + spreadsheet.getUrl() + '#gid=' + sheet.getSheetId());
}

/** Запустите вручную один раз, чтобы выдать разрешение на отправку уведомлений. */
function authorizeNotifications() {
  const recipient = notificationRecipient_();
  if (!recipient) {
    throw new Error('Не удалось определить адрес владельца. Укажите его в NOTIFICATION_EMAIL.');
  }
  MailApp.getRemainingDailyQuota();
  Logger.log('Уведомления будут приходить на: ' + recipient);
}

/** Принимает ответы с сайта. Веб-приложение должно выполняться от имени владельца. */
function doPost(e) {
  const params = (e && e.parameter) || {};
  const token = String(params.submissionToken || '').replace(/[^a-zA-Z0-9-]/g, '').slice(0, 100);
  try {
    const name = String(params.guestName || '').trim().slice(0, 120);
    const attendance = String(params.attendance || '').trim();
    const comment = String(params.comment || '').trim().slice(0, 1500);

    if (!name || !['Да, с удовольствием буду!', 'К сожалению, не смогу присутствовать.'].includes(attendance)) {
      throw new Error('Не заполнены обязательные поля.');
    }

    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error('Сначала выполните setupExistingSpreadsheet().');

    sheet.appendRow([new Date(), name, attendance, comment]);

    // Ошибка отправки письма не отменяет уже сохранённый ответ.
    try {
      const recipient = notificationRecipient_();
      if (recipient) {
        MailApp.sendEmail({
          to: recipient,
          subject: 'Новый ответ на приглашение — свадьба Дмитрия и Яны',
          body: 'Поступил новый ответ на свадебное приглашение.\n\n' +
            'Имя: ' + name + '\n' +
            'Присутствие: ' + attendance + '\n' +
            'Комментарий: ' + (comment || 'Нет') + '\n\n' +
            'Таблица ответов: ' + spreadsheet.getUrl()
        });
      }
    } catch (mailError) {
      console.error('Ответ сохранён, но письмо не отправлено: ' + mailError);
    }

    return response_({ type: 'rsvp-success', token: token });
  } catch (error) {
    console.error(error);
    return response_({ type: 'rsvp-error', token: token });
  }
}

function notificationRecipient_() {
  return NOTIFICATION_EMAIL || Session.getEffectiveUser().getEmail();
}

function response_(payload) {
  const safeJson = JSON.stringify(payload).replace(/</g, '\\u003c');
  return HtmlService.createHtmlOutput(
    '<!doctype html><html><head><meta charset="utf-8"></head><body><script>' +
    'window.top.postMessage(' + safeJson + ', "*");' +
    '</script></body></html>'
  ).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
