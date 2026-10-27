/**
 * Тренажёр тестов: сервер статистики (Google Таблица). Версия 2.
 * Сайт отправляет сюда прогресс и результаты каждого студента по его логину.
 *
 * Проверка доступа: список студентов берётся из users.js на сайте.
 * В users.js лежит SHA-256 от ключа студента. Сам ключ знает только тот, кто знает пароль.
 *
 * Листы:
 *  «Результаты»  одна строка = один пройденный тест
 *  «Студенты»    сводка по каждому логину (считается формулой)
 *  «Прогресс»    служебный: прогресс каждого логина для синхронизации между устройствами
 */

const RES = 'Результаты';
const SUM = 'Студенты';
const PROG = 'Прогресс';
const USERS_URL = 'https://abdurakhmon-sl.github.io/test-trainer/users.js';
const RES_HEAD = ['Дата', 'Логин', 'Имя', 'Предмет', 'Режим', 'Правильно', 'Всего', '%', 'Оценка', 'Темы с ошибками', 'Номера ошибок'];
const PROG_HEAD = ['Логин', 'Имя', 'Ключ', 'Данные', 'Обновлено'];

function doGet() {
  return out_({ ok: true, app: 'trainer', v: 2 });
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (x) { return out_({ ok: false, error: 'bad_json' }); }
  try {
    const login = String(req.login || '').trim().toLowerCase();
    const token = String(req.token || '');
    if (!login || !/^[0-9a-f]{64}$/.test(token)) return out_({ ok: false, error: 'auth' });
    const pub = users_()[login];
    if (!pub || sha_(token) !== pub) return out_({ ok: false, error: 'auth' });

    const ss = SpreadsheetApp.getActive();
    let sh = ss.getSheetByName(PROG), res = ss.getSheetByName(RES);
    if (!sh || !res || !ss.getSheetByName(SUM)) { locked_(() => ensure_(ss)); sh = ss.getSheetByName(PROG); res = ss.getSheetByName(RES); }

    if (req.action === 'load') {
      const row = findRow_(sh, login);
      const data = row ? String(sh.getRange(row, 4).getValue()) : '';
      return out_({ ok: true, db: data ? JSON.parse(data) : null });
    }

    if (req.action === 'save') {
      const name = String(req.name || login);
      if (req.db) {
        const vals = [login, name, '', JSON.stringify(req.db), new Date()];
        let row = findRow_(sh, login);
        if (row) sh.getRange(row, 1, 1, 5).setValues([vals]);
        else locked_(() => { row = findRow_(sh, login); if (row) sh.getRange(row, 1, 1, 5).setValues([vals]); else sh.appendRow(vals); });
      }
      const results = Array.isArray(req.results) ? req.results : (req.result ? [req.result] : []);
      results.slice(0, 50).forEach(r => {
        const mode = r.mode === 'exam' ? 'Экзамен' : (r.partial ? 'Подготовка (не до конца)' : 'Подготовка');
        const when = r.at ? new Date(r.at) : new Date();
        res.appendRow([
          isNaN(when) ? new Date() : when, login, name, String(r.subject || ''), mode,
          Number(r.ok) || 0, Number(r.total) || 0, r.total ? (Number(r.ok) || 0) / Number(r.total) : 0,
          Number(r.grade) || '', String(r.weak || ''), String(r.wrong || '')
        ]);
        const lr = res.getLastRow();
        res.getRange(lr, 1).setNumberFormat('dd.mm.yyyy hh:mm');
        res.getRange(lr, 8).setNumberFormat('0%');
      });
      return out_({ ok: true, saved: results.length });
    }
    return out_({ ok: false, error: 'action' });
  } catch (err) {
    return out_({ ok: false, error: 'server', detail: String(err).slice(0, 200) });
  }
}

/* Список студентов с сайта, кэш на 5 минут. */
function users_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('users');
  if (hit) return JSON.parse(hit);
  const txt = UrlFetchApp.fetch(USERS_URL + '?t=' + Date.now(), { muteHttpExceptions: true }).getContentText();
  const map = {};
  const re = /login:\s*"([^"]+)"[^}]*?hash:\s*"([0-9a-f]{64})"/g;
  let m;
  while ((m = re.exec(txt))) map[m[1].trim().toLowerCase()] = m[2];
  if (Object.keys(map).length) cache.put('users', JSON.stringify(map), 300);
  return map;
}

function sha_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8)
    .map(b => ('0' + (b & 255).toString(16)).slice(-2)).join('');
}

function locked_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(8000)) throw new Error('busy');
  try { return fn(); } finally { lock.releaseLock(); }
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function findRow_(sh, login) {
  const n = sh.getLastRow();
  if (n < 2) return 0;
  const vals = sh.getRange(2, 1, n - 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) if (String(vals[i][0]).toLowerCase() === login) return i + 2;
  return 0;
}

/* Создаёт листы, если их нет. */
function ensure_(ss) {
  let res = ss.getSheetByName(RES);
  if (!res) {
    res = ss.getSheets()[0];
    if (res.getLastRow() > 1 || ss.getSheetByName(SUM) || ss.getSheetByName(PROG)) res = ss.insertSheet(RES, 0); else res.setName(RES);
    res.getRange(1, 1, 1, RES_HEAD.length).setValues([RES_HEAD]);
    head_(res, RES_HEAD.length);
    res.getRange('A2:A').setNumberFormat('dd.mm.yyyy hh:mm');
    res.getRange('H2:H').setNumberFormat('0%');
  }
  if (!ss.getSheetByName(SUM)) {
    const s = ss.insertSheet(SUM, 0);
    s.getRange('A1').setFormula(
      '=IFERROR(QUERY(\'' + RES + '\'!A:K,' +
      '"select B, C, count(A), avg(H), max(H), max(A) where B is not null group by B, C order by max(A) desc ' +
      'label B \'Логин\', C \'Имя\', count(A) \'Тестов\', avg(H) \'Средний %\', max(H) \'Лучший %\', max(A) \'Последний тест\'",1),"Пока нет результатов")');
    head_(s, 6);
    s.getRange('D2:E').setNumberFormat('0%');
    s.getRange('F2:F').setNumberFormat('dd.mm.yyyy hh:mm');
  }
  if (!ss.getSheetByName(PROG)) {
    const p = ss.insertSheet(PROG);
    p.getRange(1, 1, 1, PROG_HEAD.length).setValues([PROG_HEAD]);
    head_(p, PROG_HEAD.length);
    p.getRange('D:D').setNumberFormat('@');
    p.getRange('E2:E').setNumberFormat('dd.mm.yyyy hh:mm');
    p.hideSheet();
  }
}

function head_(sh, n) {
  sh.getRange(1, 1, 1, n).setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#2445C7');
  sh.setFrozenRows(1);
}

/* Запустите вручную один раз после обновления кода: создаст листы и попросит разрешение на чтение users.js. */
function setup() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone('Asia/Tashkent');
  ss.setSpreadsheetLocale('ru_RU');
  ensure_(ss);
  CacheService.getScriptCache().remove('users');
  Logger.log('Студентов в users.js: ' + Object.keys(users_()).length);
}
