/** ============================================================
 *  ТАБЛО НА ЗАДАЧИТЕ — Backend (само JSON, без страници)
 *  ------------------------------------------------------------
 *  Тази версия НЕ показва никакви страници/HTML. Тя е само
 *  "склад за данни" — приема заявки и връща JSON. Визията
 *  (страниците, бутоните) се хостват отделно на GitHub Pages.
 *
 *  НАСТРОЙКА:
 *  1. Отвори своя Google Sheet (същия, който вече ползваш)
 *  2. Разширения → Apps Script
 *  3. Изтрий стария код (изцяло) и постави целия този файл
 *  4. Запази (Ctrl+S)
 *  5. Deploy → New deployment → Web app (тип "Web app")
 *       Execute as:      Me
 *       Who has access:  Anyone
 *  6. Копирай линка, който завършва на /exec
 *  7. Постави го в config.js на GitHub Pages сайта (виж README.md)
 *
 *  При промяна на кода по-късно: Deploy → Manage deployments →
 *  моливче → Version: New → Deploy (иначе промените не важат).
 * ============================================================ */

const NAME_MAP = {
  geri:  'Гери',
  zori:  'Зори',
  zara:  'Зара',
  tedi:  'Теди',
  lacho: 'Лъчо',
  drago: 'Драго',
  ivo:   'Иво',
  todor: 'Тодор'
};
const DEFAULT_PIN = '1234'; // смени тук за нов ПИН (или сложи ADMIN_PIN в Script Properties)
const SHEET_TASKS = 'Задачи';
const SHEET_ARCHIVE = 'Архив';

/** ================= ПОМОЩНИ ================= */
function getPin_() {
  return PropertiesService.getScriptProperties().getProperty('ADMIN_PIN') || DEFAULT_PIN;
}

function ensureSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let t = ss.getSheetByName(SHEET_TASKS);
  if (!t) {
    t = ss.insertSheet(SHEET_TASKS);
    t.appendRow(['ID', 'Текст', 'Изпълнител', 'Статус', 'Създадена', 'Изпълнена']);
    t.setFrozenRows(1);
  }
  let a = ss.getSheetByName(SHEET_ARCHIVE);
  if (!a) {
    a = ss.insertSheet(SHEET_ARCHIVE);
    a.appendRow(['Дата на приключване', 'ID', 'Текст', 'Изпълнител', 'Създадена', 'Изпълнена']);
    a.setFrozenRows(1);
  }
  return { t, a };
}

function readTasks_() {
  const { t } = ensureSheets_();
  const values = t.getDataRange().getValues();
  const tz = Session.getScriptTimeZone();
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    if (!r[0]) continue;
    rows.push({
      sheetRow: i + 1,
      id: String(r[0]),
      text: r[1],
      assignee: r[2],
      status: r[3],
      created: r[4] instanceof Date ? Utilities.formatDate(r[4], tz, 'dd.MM.yyyy HH:mm') : String(r[4] || ''),
      done: r[5] instanceof Date ? Utilities.formatDate(r[5], tz, 'dd.MM.yyyy HH:mm') : String(r[5] || '')
    });
  }
  return rows;
}

function stripSheetRow_(r) {
  const copy = Object.assign({}, r);
  delete copy.sheetRow;
  return copy;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** ================= GET: четене (без бисквитки, без redirect) ================= */
function doGet(e) {
  const p = e.parameter || {};

  if (p.action === 'verifyPin') {
    return json_({ ok: p.pin === getPin_() });
  }

  if (p.action === 'tasks') {
    const rows = readTasks_().map(stripSheetRow_);

    // Админ (има валиден ПИН) вижда всичко
    if (p.pin && p.pin === getPin_()) {
      return json_({ ok: true, mode: 'admin', tasks: rows });
    }

    // Личен изглед (само активните задачи за това име)
    if (p.u && NAME_MAP[p.u]) {
      const name = NAME_MAP[p.u];
      const mine = rows.filter(r => r.assignee === name && r.status !== 'Изпълнена');
      return json_({ ok: true, mode: 'person', name: name, tasks: mine });
    }

    return json_({ ok: false, error: 'Липсва ПИН или личен код (u).' });
  }

  return json_({ ok: false, error: 'Непозната заявка.' });
}

/** ================= POST: промени (изисква ПИН при всяка заявка) =================
 *  Забележка: фронтендът трябва да прави fetch(...) БЕЗ да задава
 *  header "Content-Type: application/json" — иначе браузърът праща
 *  CORS preflight (OPTIONS), който Apps Script не поддържа.
 *  Изпращай тялото като обикновен текст (default text/plain) — тук
 *  все пак го парсваме като JSON.
 */
function doPost(e) {
  let body = {};
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'Невалидна заявка (лош JSON).' });
  }

  if (body.pin !== getPin_()) {
    return json_({ ok: false, error: 'Грешен ПИН.' });
  }

  switch (body.action) {
    case 'addTask': {
      const text = (body.text || '').trim();
      const assignee = body.assignee;
      if (!text || Object.values(NAME_MAP).indexOf(assignee) === -1) {
        return json_({ ok: false, error: 'Невалидни данни за задачата.' });
      }
      const { t } = ensureSheets_();
      t.appendRow([Utilities.getUuid(), text, assignee, 'Активна', new Date(), '']);
      return json_({ ok: true, tasks: readTasks_().map(stripSheetRow_) });
    }

    case 'toggleTask': {
      const { t } = ensureSheets_();
      const row = readTasks_().find(r => r.id === body.id);
      if (row) {
        t.getRange(row.sheetRow, 4).setValue(body.done ? 'Изпълнена' : 'Активна');
        t.getRange(row.sheetRow, 6).setValue(body.done ? new Date() : '');
      }
      return json_({ ok: true, tasks: readTasks_().map(stripSheetRow_) });
    }

    case 'deleteTask': {
      const { t } = ensureSheets_();
      const row = readTasks_().find(r => r.id === body.id);
      if (row) t.deleteRow(row.sheetRow);
      return json_({ ok: true, tasks: readTasks_().map(stripSheetRow_) });
    }

    case 'closeDay': {
      const { t, a } = ensureSheets_();
      const rows = readTasks_();
      const done = rows.filter(r => r.status === 'Изпълнена');
      const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy');
      done.forEach(r => a.appendRow([today, r.id, r.text, r.assignee, r.created, r.done]));
      done.sort((x, y) => y.sheetRow - x.sheetRow).forEach(r => t.deleteRow(r.sheetRow));
      return json_({ ok: true, tasks: readTasks_().map(stripSheetRow_) });
    }

    default:
      return json_({ ok: false, error: 'Непозната команда.' });
  }
}
