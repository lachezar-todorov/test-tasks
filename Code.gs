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

const PROJECTS = [
  'Холидей парк, гр. Русе',
  'Дом за възрастни хора, гр. Момин проход',
  'Холидей парк, Симеоново',
  'Алумакс',
  'Биотрейд',
  'Белисимо',
  'Ресторант Шамот',
  'Здравко Здравков',
  'Технополис, гр. Монтана',
  'Младен, с. Лозен'
];

function normalizeProject_(project) {
  return PROJECTS.indexOf(project) !== -1 ? project : '';
}

const DEFAULT_PIN = '1234'; // смени тук за нов ПИН (или сложи ADMIN_PIN в Script Properties)
const SHEET_TASKS = 'Задачи';
const SHEET_ARCHIVE = 'Архив';
const SHEET_COMMENTS = 'Коментари';

const STATUS_ACTIVE = 'Активна';
const STATUS_DONE = 'Изпълнена';
const STATUS_PENDING = 'Чака одобрение';

/** ================= ПОМОЩНИ ================= */
function getPin_() {
  return PropertiesService.getScriptProperties().getProperty('ADMIN_PIN') || DEFAULT_PIN;
}

function ensureSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let t = ss.getSheetByName(SHEET_TASKS);
  if (!t) {
    t = ss.insertSheet(SHEET_TASKS);
    t.appendRow(['ID', 'Текст', 'Изпълнител', 'Статус', 'Създадена', 'Изпълнена', 'Важна', 'Проект']);
    t.setFrozenRows(1);
  } else {
    // добавя се към по-стари листове "Задачи", създадени преди тези колони
    if (t.getRange(1, 7).getValue() === '') t.getRange(1, 7).setValue('Важна');
    if (t.getRange(1, 8).getValue() === '') t.getRange(1, 8).setValue('Проект');
  }

  let a = ss.getSheetByName(SHEET_ARCHIVE);
  if (!a) {
    a = ss.insertSheet(SHEET_ARCHIVE);
    a.appendRow(['Дата на приключване', 'ID', 'Текст', 'Изпълнител', 'Създадена', 'Изпълнена']);
    a.setFrozenRows(1);
  }

  let c = ss.getSheetByName(SHEET_COMMENTS);
  if (!c) {
    c = ss.insertSheet(SHEET_COMMENTS);
    c.appendRow(['ID', 'ID на задача', 'Автор', 'Текст', 'Дата', 'Задача']);
    c.setFrozenRows(1);
  } else if (c.getRange(1, 6).getValue() === '') {
    // добавя се към по-стари листове "Коментари", създадени преди колоната "Задача"
    c.getRange(1, 6).setValue('Задача');
  }

  return { t, a, c };
}

/** Еднократна помощна функция — попълва колона "Задача" за коментари,
 *  добавени преди тя да съществува. Пусни я ръчно веднъж от редактора
 *  на Apps Script (избери функцията в падащото меню → Run). Не се
 *  извиква от doGet/doPost.
 */
function backfillCommentTaskNames() {
  const { t, a, c } = ensureSheets_();
  const nameById = {};
  t.getDataRange().getValues().slice(1).forEach(function (r) { if (r[0]) nameById[String(r[0])] = r[1]; });
  a.getDataRange().getValues().slice(1).forEach(function (r) { if (r[1]) nameById[String(r[1])] = r[2]; });
  const values = c.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (!row[0] || row[5]) continue;
    const name = nameById[String(row[1])];
    if (name) c.getRange(i + 1, 6).setValue(name);
  }
}

/** Чете задачите (реда тук е без значение — sortByActivity_ подрежда накрая). */
function readTasks_() {
  const { t } = ensureSheets_();
  const values = t.getDataRange().getValues();
  const tz = Session.getScriptTimeZone();
  const rows = [];
  for (let i = values.length - 1; i >= 1; i--) {
    const r = values[i];
    if (!r[0]) continue;
    const createdDate = r[4] instanceof Date ? r[4] : (r[4] ? new Date(r[4]) : null);
    rows.push({
      sheetRow: i + 1,
      id: String(r[0]),
      text: r[1],
      assignee: r[2],
      status: r[3],
      created: createdDate ? Utilities.formatDate(createdDate, tz, 'dd.MM.yyyy HH:mm') : String(r[4] || ''),
      done: r[5] instanceof Date ? Utilities.formatDate(r[5], tz, 'dd.MM.yyyy HH:mm') : String(r[5] || ''),
      important: r[6] === true,
      project: r[7] || '',
      createdRaw_: createdDate ? createdDate.getTime() : 0
    });
  }
  return rows;
}

/** Чете коментарите и ги групира по ID на задача (най-стари първи). */
function readComments_() {
  const { c } = ensureSheets_();
  const values = c.getDataRange().getValues();
  const tz = Session.getScriptTimeZone();
  const map = {};
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    if (!r[0]) continue;
    const taskId = String(r[1]);
    if (!map[taskId]) map[taskId] = [];
    const dateVal = r[4] instanceof Date ? r[4] : (r[4] ? new Date(r[4]) : null);
    map[taskId].push({
      id: String(r[0]),
      author: r[2],
      text: r[3],
      date: dateVal ? Utilities.formatDate(dateVal, tz, 'dd.MM.yyyy HH:mm') : String(r[4] || ''),
      taskText: r[5] || '',
      dateRaw_: dateVal ? dateVal.getTime() : 0
    });
  }
  return map;
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

/** Изгражда пълния отговор със задачи (+коментари), според това дали
 *  извикващият е админ (вижда всичко) или личен потребител (вижда само своите).
 *  Подрежда задачите по последна активност (създаване или последен нов/
 *  редактиран коментар) — най-скоро активната отгоре.
 */
function buildTasksPayload_(opts) {
  const rows = readTasks_().map(stripSheetRow_);
  const commentsMap = readComments_();
  rows.forEach(function (r) {
    r.comments = commentsMap[r.id] || [];
    const commentTimes = r.comments.map(function (cm) { return cm.dateRaw_ || 0; });
    r.lastActivity_ = Math.max(r.createdRaw_ || 0, commentTimes.length ? Math.max.apply(null, commentTimes) : 0);
  });
  rows.sort(function (a, b) { return b.lastActivity_ - a.lastActivity_; });
  rows.forEach(function (r) {
    delete r.createdRaw_;
    delete r.lastActivity_;
    r.comments.forEach(function (cm) { delete cm.dateRaw_; });
  });

  if (opts && opts.u && NAME_MAP[opts.u]) {
    const name = NAME_MAP[opts.u];
    const mine = rows.filter(function (r) {
      return r.assignee === name && r.status !== STATUS_DONE && r.status !== STATUS_PENDING;
    });
    const pending = rows.filter(function (r) {
      return r.assignee === name && r.status === STATUS_PENDING;
    });
    return { ok: true, mode: 'person', name: name, tasks: mine, pendingTasks: pending };
  }

  return { ok: true, mode: 'admin', tasks: rows };
}

function tasksResponse_(opts) {
  return json_(buildTasksPayload_(opts));
}

/** ================= GET: четене (без бисквитки, без redirect) ================= */
function doGet(e) {
  const p = e.parameter || {};

  if (p.action === 'verifyPin') {
    return json_({ ok: p.pin === getPin_() });
  }

  if (p.action === 'tasks') {
    // Админ (има валиден ПИН) вижда всичко
    if (p.pin && p.pin === getPin_()) {
      return json_(buildTasksPayload_());
    }

    // Личен изглед (само задачите за това име)
    if (p.u && NAME_MAP[p.u]) {
      return json_(buildTasksPayload_({ u: p.u }));
    }

    return json_({ ok: false, error: 'Липсва ПИН или личен код (u).' });
  }

  return json_({ ok: false, error: 'Непозната заявка.' });
}

/** ================= POST: промени =================
 *  Забележка: фронтендът трябва да прави fetch(...) БЕЗ да задава
 *  header "Content-Type: application/json" — иначе браузърът праща
 *  CORS preflight (OPTIONS), който Apps Script не поддържа.
 *  Изпращай тялото като обикновен текст (default text/plain) — тук
 *  все пак го парсваме като JSON.
 *
 *  Повечето действия изискват валиден админ ПИН. Изключения:
 *  'personAddTask' (лично предложение за задача, чака одобрение) и
 *  'addComment' (коментар — може и от админ, и от личен потребител),
 *  които се удостоверяват през личния код "u".
 */
function doPost(e) {
  let body = {};
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'Невалидна заявка (лош JSON).' });
  }

  const isAdmin = body.pin === getPin_();
  const personName = NAME_MAP[body.u];

  if (!isAdmin && !personName) {
    return json_({ ok: false, error: 'Грешен ПИН или невалиден личен код.' });
  }

  switch (body.action) {
    case 'addTask': {
      if (!isAdmin) return json_({ ok: false, error: 'Само от администраторския панел може да се добавят потвърдени задачи.' });
      const text = (body.text || '').trim();
      const assignee = body.assignee;
      if (!text || Object.values(NAME_MAP).indexOf(assignee) === -1) {
        return json_({ ok: false, error: 'Невалидни данни за задачата.' });
      }
      const { t } = ensureSheets_();
      t.appendRow([Utilities.getUuid(), text, assignee, STATUS_ACTIVE, new Date(), '', !!body.important, normalizeProject_(body.project)]);
      return tasksResponse_();
    }

    case 'personAddTask': {
      if (!personName) return json_({ ok: false, error: 'Невалиден личен код.' });
      const text = (body.text || '').trim();
      if (!text) return json_({ ok: false, error: 'Липсва текст на задачата.' });
      const { t } = ensureSheets_();
      t.appendRow([Utilities.getUuid(), text, personName, STATUS_PENDING, new Date(), '', !!body.important, normalizeProject_(body.project)]);
      return tasksResponse_({ u: body.u });
    }

    case 'approveTask': {
      if (!isAdmin) return json_({ ok: false, error: 'Само от администраторския панел може да се одобряват задачи.' });
      const { t } = ensureSheets_();
      const row = readTasks_().find(function (r) { return r.id === body.id; });
      if (row) t.getRange(row.sheetRow, 4).setValue(STATUS_ACTIVE);
      return tasksResponse_();
    }

    case 'editTask': {
      if (!isAdmin) return json_({ ok: false, error: 'Само от администраторския панел може да се редактират задачи.' });
      const { t } = ensureSheets_();
      const row = readTasks_().find(function (r) { return r.id === body.id; });
      if (!row) return json_({ ok: false, error: 'Задачата не е намерена.' });
      const text = (body.text || '').trim();
      const assignee = body.assignee;
      if (!text) return json_({ ok: false, error: 'Липсва текст на задачата.' });
      t.getRange(row.sheetRow, 2).setValue(text);
      if (assignee && Object.values(NAME_MAP).indexOf(assignee) !== -1) {
        t.getRange(row.sheetRow, 3).setValue(assignee);
      }
      t.getRange(row.sheetRow, 7).setValue(!!body.important);
      t.getRange(row.sheetRow, 8).setValue(normalizeProject_(body.project));
      return tasksResponse_();
    }

    case 'toggleTask': {
      if (!isAdmin) return json_({ ok: false, error: 'Само от администраторския панел може да се променя статус.' });
      const { t } = ensureSheets_();
      const row = readTasks_().find(function (r) { return r.id === body.id; });
      if (row) {
        t.getRange(row.sheetRow, 4).setValue(body.done ? STATUS_DONE : STATUS_ACTIVE);
        t.getRange(row.sheetRow, 6).setValue(body.done ? new Date() : '');
      }
      return tasksResponse_();
    }

    case 'deleteTask': {
      if (!isAdmin) return json_({ ok: false, error: 'Само от администраторския панел може да се изтриват задачи.' });
      const { t } = ensureSheets_();
      const row = readTasks_().find(function (r) { return r.id === body.id; });
      if (row) t.deleteRow(row.sheetRow);
      return tasksResponse_();
    }

    case 'addComment': {
      const text = (body.text || '').trim();
      const taskId = body.taskId;
      if (!text || !taskId) return json_({ ok: false, error: 'Липсва текст на коментара.' });
      const author = isAdmin ? 'Админ' : personName;
      const task = readTasks_().find(function (r) { return r.id === taskId; });
      const { c } = ensureSheets_();
      c.appendRow([Utilities.getUuid(), taskId, author, text, new Date(), task ? task.text : '']);
      return tasksResponse_(isAdmin ? undefined : { u: body.u });
    }

    case 'editComment': {
      const text = (body.text || '').trim();
      const commentId = body.commentId;
      if (!text || !commentId) return json_({ ok: false, error: 'Липсва текст на коментара.' });
      const { c } = ensureSheets_();
      const values = c.getDataRange().getValues();
      let targetRow = -1, author = null;
      for (let i = 1; i < values.length; i++) {
        if (String(values[i][0]) === commentId) { targetRow = i + 1; author = values[i][2]; break; }
      }
      if (targetRow === -1) return json_({ ok: false, error: 'Коментарът не е намерен.' });
      if (!isAdmin && author !== personName) {
        return json_({ ok: false, error: 'Можеш да редактираш само своите коментари.' });
      }
      c.getRange(targetRow, 4).setValue(text);
      // датата се обновява, за да преброи като нова активност (задачата отива най-отгоре)
      c.getRange(targetRow, 5).setValue(new Date());
      return tasksResponse_(isAdmin ? undefined : { u: body.u });
    }

    case 'deleteComment': {
      const commentId = body.commentId;
      if (!commentId) return json_({ ok: false, error: 'Липсва коментар.' });
      const { c } = ensureSheets_();
      const values = c.getDataRange().getValues();
      let targetRow = -1, author = null;
      for (let i = 1; i < values.length; i++) {
        if (String(values[i][0]) === commentId) { targetRow = i + 1; author = values[i][2]; break; }
      }
      if (targetRow === -1) return json_({ ok: false, error: 'Коментарът не е намерен.' });
      if (!isAdmin && author !== personName) {
        return json_({ ok: false, error: 'Можеш да изтриваш само своите коментари.' });
      }
      c.deleteRow(targetRow);
      return tasksResponse_(isAdmin ? undefined : { u: body.u });
    }

    case 'closeDay': {
      if (!isAdmin) return json_({ ok: false, error: 'Само от администраторския панел може да се приключва деня.' });
      const { t, a } = ensureSheets_();
      const rows = readTasks_();
      const done = rows.filter(function (r) { return r.status === STATUS_DONE; });
      const commentsMap = readComments_();
      const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy');
      done.forEach(function (r) {
        a.appendRow([today, r.id, r.text, r.assignee, r.created, r.done]);
        // коментарите на задачата се копират точно под нея, за четимост;
        // оригиналите остават непокътнати в листа "Коментари"
        (commentsMap[r.id] || []).forEach(function (cm) {
          a.appendRow(['', '', '↳ ' + cm.author + ': ' + cm.text, '', cm.date, '']);
          a.getRange(a.getLastRow(), 1, 1, 6).setFontStyle('italic').setFontColor('#8a95a1');
        });
      });
      done.sort(function (x, y) { return y.sheetRow - x.sheetRow; }).forEach(function (r) { t.deleteRow(r.sheetRow); });
      return tasksResponse_();
    }

    default:
      return json_({ ok: false, error: 'Непозната команда.' });
  }
}
