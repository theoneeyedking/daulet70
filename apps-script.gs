/**
 * Мерей той · 70 жас — приём ответов RSVP в Google Таблицу.
 *
 * Установка (5 минут):
 * 1. Создай новую Google Таблицу (sheets.new) → меню «Расширения» → «Apps Script».
 * 2. Удали всё в редакторе и вставь этот файл целиком. Сохрани.
 * 3. «Развернуть» → «Новое развертывание» → тип «Веб-приложение».
 *    Выполнять как: «От моего имени». Доступ: «Все» (Anyone).
 * 4. Разреши доступ, скопируй URL вида https://script.google.com/macros/s/.../exec
 * 5. Вставь этот URL в index.html → CONFIG.SCRIPT_URL.
 *
 * Telegram-уведомления (необязательно):
 * «Настройки проекта» (шестерёнка) → «Свойства скрипта» → добавь
 *   TG_TOKEN = токен бота от @BotFather
 *   TG_CHAT  = id чата (узнать: написать боту, открыть https://api.telegram.org/bot<TOKEN>/getUpdates)
 */

const SHEET = 'Жауаптар';
const SUMMARY = 'Қорытынды';

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const data = JSON.parse(e.postData.contents || '{}');
    if (data.website) return out({ ok: true }); // honeypot
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = getSheet_(ss);
    const name = String(data.name || '').slice(0, 120);
    const row = [
      new Date(),
      name,
      String(data.phone || '').slice(0, 30),
      String(data.answer || '').slice(0, 60),
      Number(data.count) || 0,
      String(data.wish || '').slice(0, 1000),
      String(data.guestLink || '').slice(0, 60)
    ];
    // Если этот же человек отвечает повторно — обновляем его строку
    const last = sh.getLastRow();
    let target = 0;
    if (last > 1 && name && data.answer !== 'Тілек') { // отдельные тілектер всегда новой строкой
      const names = sh.getRange(2, 2, last - 1, 1).getValues();
      for (let i = names.length - 1; i >= 0; i--) {
        if (String(names[i][0]).trim().toLowerCase() === name.trim().toLowerCase()) { target = i + 2; break; }
      }
    }
    if (target) sh.getRange(target, 1, 1, row.length).setValues([row]);
    else sh.appendRow(row);
    ensureSummary_(ss);
    notify_(row, !!target);
    return out({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

// Стена тілектер на сайте: ?action=wishes&callback=fn → последние тілектер (имя + текст, без телефонов)
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action !== 'wishes') return out({ ok: true, service: 'merei-toi-rsvp' });
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET);
  let wishes = [], guests = 0;
  if (sh && sh.getLastRow() > 1) {
    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues();
    guests = rows.filter(r => r[3] !== 'Тілек').reduce((n, r) => n + (Number(r[4]) || 0), 0);
    wishes = rows.filter(r => String(r[5]).trim())
      .sort((a, b) => b[0] - a[0])
      .slice(0, 150)
      .map(r => ({ name: String(r[1]), text: String(r[5]) }));
  }
  const list = { wishes, guests };
  const cb = String(p.callback || '').replace(/[^\w$]/g, '');
  if (cb) return ContentService.createTextOutput(cb + '(' + JSON.stringify(list) + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  return out(list);
}

function getSheet_(ss) {
  let sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET);
    sh.appendRow(['Уақыты', 'Аты-жөні', 'Телефон', 'Жауабы', 'Адам саны', 'Тілегі', 'Жеке сілтеме']);
    sh.setFrozenRows(1);
    sh.getRange('A1:G1').setFontWeight('bold').setBackground('#0F3B30').setFontColor('#F3EAD7');
    sh.setColumnWidths(1, 7, 180);
    sh.setColumnWidth(6, 360);
  }
  return sh;
}

function ensureSummary_(ss) {
  let s = ss.getSheetByName(SUMMARY);
  if (s) return;
  s = ss.insertSheet(SUMMARY);
  s.getRange('A1:B5').setValues([
    ['Барлығы келеді (адам)', `=SUM('${SHEET}'!E2:E)`],
    ['Жалғыз келеді', `=COUNTIF('${SHEET}'!D2:D,"Иә, жалғыз өзім барамын")`],
    ['Жұбайымен келеді', `=COUNTIF('${SHEET}'!D2:D,"Жұбайыммен бірге барамын")`],
    ['Келе алмайды', `=COUNTIF('${SHEET}'!D2:D,"Өкінішке орай, келе алмаймын")`],
    ['Жауап бергендер', `=COUNTA('${SHEET}'!B2:B)`]
  ]);
  s.getRange('A1:A5').setFontWeight('bold');
  s.setColumnWidth(1, 240);
}

function notify_(row, updated) {
  const p = PropertiesService.getScriptProperties();
  const token = p.getProperty('TG_TOKEN'), chat = p.getProperty('TG_CHAT');
  if (!token || !chat) return;
  const text = (updated ? '✏️ Жауап өзгертілді\n' : '🎉 Жаңа жауап\n') +
    '👤 ' + row[1] + (row[2] ? ' (' + row[2] + ')' : '') + '\n' +
    '✅ ' + row[3] + ' — ' + row[4] + ' адам' +
    (row[5] ? '\n💬 ' + row[5] : '');
  UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendMessage', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    payload: JSON.stringify({ chat_id: chat, text: text })
  });
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
