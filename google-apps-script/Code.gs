const SHEET_NAME = 'Ответы гостей';

/**
 * Один раз запустите setup() вручную в редакторе Apps Script.
 * Функция создаст таблицу и сохранит её ID в настройках проекта.
 */
function setup() {
  const spreadsheet = SpreadsheetApp.create('Свадьба Дмитрия и Яны — ответы гостей');
  const sheet = spreadsheet.getActiveSheet();
  sheet.setName(SHEET_NAME);
  sheet.appendRow(['Дата и время ответа', 'Имя и фамилия', 'Присутствие', 'Комментарий']);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
  sheet.autoResizeColumns(1, 4);
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', spreadsheet.getId());
  Logger.log('Ссылка на таблицу с ответами: ' + spreadsheet.getUrl());
}

/**
 * Принимает ответы с сайта. Публикуйте проект как веб-приложение:
 * выполнять от имени владельца, доступ — все.
 */
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

    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    if (!spreadsheetId) throw new Error('Сначала выполните setup().');

    const sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error('Не найден лист с ответами.');

    sheet.appendRow([new Date(), name, attendance, comment]);
    return response_({ type: 'rsvp-success', token: token });
  } catch (error) {
    console.error(error);
    return response_({ type: 'rsvp-error', token: token });
  }
}

function response_(payload) {
  const safeJson = JSON.stringify(payload).replace(/</g, '\\u003c');
  return HtmlService.createHtmlOutput(
    '<!doctype html><html><head><meta charset="utf-8"></head><body><script>' +
    'window.top.postMessage(' + safeJson + ', "*");' +
    '</script></body></html>'
  ).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
