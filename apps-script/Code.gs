/**
 * Junction CRM — Google Apps Script, прив'язаний до Google-таблиці «Junction CRM».
 *
 * Що робить:
 *  - doPost: приймає заявку з форми сайту, записує рядок у лист «Заявки» і шле сповіщення в Telegram;
 *  - doGet:  віддає заявки сторінці crm.html і змінює статус/коментар (лише з правильним паролем).
 *
 * Секрети НЕ в коді, а у «Властивостях скрипта» (Project Settings → Script properties):
 *  TG_TOKEN   — токен бота від @BotFather
 *  TG_CHAT_ID — заповнюється сам функцією findChatId()
 *  CRM_KEY    — пароль до сторінки crm.html (придумай довгий)
 */

const SHEET = 'Заявки';
const HEAD = ['ID', 'Дата', "Ім'я", 'Бізнес', 'Instagram / сайт', 'Контакт', 'Послуга', 'Що болить',
  'Джерело (UTM)', 'Сторінка', 'Статус', 'Коментар'];
const STATUSES = ['Нова', "Зв'язались", 'Розбір', 'Пропозиція', 'Клієнт', 'Відмова'];
const COL = { status: 11, comment: 12 };

const props_ = () => PropertiesService.getScriptProperties();
const sheet_ = () => {
  const ss = SpreadsheetApp.getActive();
  return ss.getSheetByName(SHEET) || ss.insertSheet(SHEET);
};
const clip_ = (v, n) => String(v == null ? '' : v).trim().slice(0, n);
// a value starting with = + - @ would be read by Sheets as a formula
const safe_ = v => (typeof v === 'string' && /^[=+\-@]/.test(v) ? "'" + v : v);
const json_ = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);

/** 1) Запусти один раз: створює шапку таблиці й список статусів. */
function setup() {
  const sh = sheet_();
  if (sh.getLastRow() === 0) sh.appendRow(HEAD);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, HEAD.length).setFontWeight('bold').setBackground('#4b0082').setFontColor('#ffffff');
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).build();
  sh.getRange(2, COL.status, 1000, 1).setDataValidation(rule);
  sh.setColumnWidths(1, HEAD.length, 160);
}

/** 2) Напиши своєму боту в Telegram будь-яке повідомлення, потім запусти цю функцію. */
function findChatId() {
  const token = props_().getProperty('TG_TOKEN');
  if (!token) throw new Error('Спочатку додай TG_TOKEN у властивості скрипта.');
  const res = JSON.parse(UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/getUpdates').getContentText());
  const last = (res.result || []).filter(u => u.message).pop();
  if (!last) throw new Error('Бот ще не отримав повідомлень. Напиши йому /start і запусти знову.');
  props_().setProperty('TG_CHAT_ID', String(last.message.chat.id));
  Logger.log('Готово, chat id збережено: ' + last.message.chat.id);
}

/** 3) Перевірка: шле тестове сповіщення в Telegram. */
function testTelegram() {
  notify_(['TEST', new Date(), 'Тест', 'Перевірка бота', '', '@test', 'Ще не знаю', '', 'direct', '', 'Нова', '']);
}

/** Заявка з сайту. */
function doPost(e) {
  const p = (e && e.parameter) || {};
  if (p.website) return json_({ ok: true });                // honeypot: so bots think it worked
  if (!clip_(p.name, 100) || !clip_(p.contact, 100)) return json_({ ok: false, error: 'missing' });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet_();
    if (sh.getLastRow() === 0) setup();
    const now = new Date();
    const id = 'L' + Utilities.formatDate(now, 'Europe/Kyiv', 'yyMMddHHmmss') + Math.floor(10 + Math.random() * 90);
    const row = [id, now, clip_(p.name, 100), clip_(p.business, 150), clip_(p.link, 200), clip_(p.contact, 100),
      clip_(p.service, 80), clip_(p.pain, 1000), clip_(p.utm, 300) || 'direct', clip_(p.page, 300), 'Нова', ''];
    sh.appendRow(row.map(safe_));
    notify_(row);
  } finally {
    lock.releaseLock();
  }
  return json_({ ok: true });
}

/** Сторінка crm.html: список заявок і зміна статусу / коментаря. */
function doGet(e) {
  const p = (e && e.parameter) || {};
  const key = props_().getProperty('CRM_KEY');
  if (!key || p.key !== key) return json_({ ok: false, error: 'auth' });

  const sh = sheet_();
  const n = Math.max(0, sh.getLastRow() - 1);

  if (p.action === 'update') {
    const ids = n ? sh.getRange(2, 1, n, 1).getValues().map(r => String(r[0])) : [];
    const i = ids.indexOf(String(p.id));
    if (i < 0) return json_({ ok: false, error: 'not_found' });
    if (p.status && STATUSES.indexOf(p.status) >= 0) sh.getRange(i + 2, COL.status).setValue(p.status);
    if (p.comment !== undefined) sh.getRange(i + 2, COL.comment).setValue(safe_(clip_(p.comment, 1000)));
    return json_({ ok: true });
  }

  const rows = n ? sh.getRange(2, 1, n, HEAD.length).getValues() : [];
  const leads = rows.map(r => ({
    id: String(r[0]), date: r[1] instanceof Date ? r[1].toISOString() : String(r[1]),
    name: r[2], business: r[3], link: r[4], contact: r[5], service: r[6], pain: r[7],
    utm: r[8], page: r[9], status: r[10] || 'Нова', comment: r[11]
  })).reverse();                                              // newest first
  return json_({ ok: true, statuses: STATUSES, leads: leads });
}

function notify_(r) {
  const token = props_().getProperty('TG_TOKEN'), chat = props_().getProperty('TG_CHAT_ID');
  if (!token || !chat) return;
  const esc = s => String(s == null ? '' : s).replace(/[<>&]/g, c => ({ '<': '&lt;', '>': '&gt;', '&': '&amp;' }[c]));
  const text = '🟢 <b>Нова заявка — Junction</b>\n\n' +
    '<b>' + esc(r[2]) + '</b> · ' + esc(r[3]) + '\n' +
    'Контакт: ' + esc(r[5]) + '\n' +
    'Instagram / сайт: ' + esc(r[4] || '—') + '\n' +
    'Послуга: ' + esc(r[6]) + '\n' +
    'Що болить: ' + esc(r[7] || '—') + '\n' +
    'Джерело: ' + esc(r[8] || 'direct');
  UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendMessage', {
    method: 'post', payload: { chat_id: chat, text: text, parse_mode: 'HTML', disable_web_page_preview: 'true' },
    muteHttpExceptions: true
  });
}
