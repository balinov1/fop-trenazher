/**
 * Сервер результатів навчального тренажера «Реєстрація ФОП».
 * Дисципліна «Основи підприємницької діяльності», Сторожинецький лісовий фаховий коледж.
 *
 * Що робить цей файл. Він виконується на боці Google і приймає два види запитів:
 *   action=save — тренажер студента надсилає результат, скрипт дописує рядок у таблицю;
 *   action=list — сторінка викладача (results.html) забирає всі результати.
 *
 * Відповідь віддається у форматі JSONP (текст «callback({...})»). Це потрібно тому,
 * що сторінка тренажера живе на github.io, а скрипт — на google.com: браузер не дозволяє
 * звичайний запит між різними сайтами, а підключення скрипта тегом <script> дозволяє.
 *
 * Паролі (див. нижче): WRITE_TOKEN мусить збігатися зі значенням writeToken у config.js;
 * ADMIN_TOKEN знає тільки викладач і вводить його на сторінці результатів.
 */

var WRITE_TOKEN = 'opd2026';            // такий самий, як у config.js
var ADMIN_TOKEN = 'ЗМІНІТЬ-ЦЕЙ-ПАРОЛЬ'; // пароль викладача; у config.js його НЕМАЄ
var SHEET_NAME  = 'Результати';

var HEADERS = ['Час надходження', 'Студент', 'Група', 'Заняття', '№ картки',
               'Бал', 'Контрольний код', 'Хвилин', 'Обраний КВЕД',
               'Обрана система', 'Кроки з помилками'];

function doGet(e) {
  var p = (e && e.parameter) || {};
  var out;
  try {
    if (p.action === 'save')      out = save_(p);
    else if (p.action === 'list') out = list_(p);
    else                          out = { ok: false, error: 'Невідома дія.' };
  } catch (err) {
    out = { ok: false, error: String(err) };
  }
  return reply_(out, p.callback);
}

/** Відповідь: JSONP, якщо вказано callback, інакше звичайний JSON. */
function reply_(obj, callback) {
  var json = JSON.stringify(obj);
  if (callback && /^[A-Za-z_$][\w$]*$/.test(callback)) {
    return ContentService
      .createTextOutput(callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

/** Аркуш результатів; створює його із заголовками, якщо ще немає. */
function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

/** Запис результату студента. */
function save_(p) {
  if (p.token !== WRITE_TOKEN) return { ok: false, error: 'Невірний код доступу.' };

  var name = String(p.n || '').trim().slice(0, 80);
  if (!name) return { ok: false, error: 'Не вказано прізвище та ім\'я.' };

  var caseId = Number(p.c) || 0;
  var score  = Number(p.s);
  if (!(score >= 0 && score <= 12)) return { ok: false, error: 'Неправильний бал.' };

  // Захист від подвійного натискання: той самий студент, та сама картка,
  // той самий бал протягом останньої години — оновлюємо рядок, а не додаємо новий.
  var sh = sheet_();
  var last = sh.getLastRow();
  if (last > 1) {
    var vals = sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
    for (var i = vals.length - 1; i >= 0 && i > vals.length - 40; i--) {
      var t = vals[i][0] instanceof Date ? vals[i][0].getTime() : 0;
      if (String(vals[i][1]) === name && Number(vals[i][4]) === caseId &&
          Number(vals[i][5]) === score && (new Date().getTime() - t) < 3600000) {
        return { ok: true, dup: true };
      }
    }
  }

  sh.appendRow([
    new Date(),
    name,
    String(p.g || '').slice(0, 40),
    String(p.l || '').slice(0, 60),
    caseId,
    score,
    String(p.k || '').slice(0, 8),
    Number(p.m) || 0,
    String(p.kv || '').slice(0, 10),
    String(p.gr || ''),
    String(p.e || '').slice(0, 120)
  ]);
  return { ok: true };
}

/** Видача всіх результатів для сторінки викладача. */
function list_(p) {
  if (p.token !== ADMIN_TOKEN) return { ok: false, error: 'Невірний пароль викладача.' };
  var sh = sheet_();
  var last = sh.getLastRow();
  if (last < 2) return { ok: true, rows: [] };
  var vals = sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
  var rows = vals.map(function (r) {
    return {
      ts: r[0] instanceof Date ? Utilities.formatDate(r[0], 'Europe/Kyiv', 'dd.MM.yyyy HH:mm') : String(r[0]),
      n: r[1], g: r[2], l: r[3], c: r[4], s: r[5], k: r[6], m: r[7], kv: r[8], gr: r[9], e: r[10]
    };
  });
  return { ok: true, rows: rows };
}

/**
 * Перевірка після встановлення: у редакторі скрипта оберіть функцію test_ і натисніть
 * «Виконати». У таблиці має з'явитися рядок на ім'я «Перевірка зв'язку».
 */
function test_() {
  var r = save_({ token: WRITE_TOKEN, n: 'Перевірка зв\'язку', g: 'ОО-21', l: 'тест',
                  c: 1, s: 12, k: 'TEST', m: 1, kv: '47.81', gr: '1', e: '' });
  Logger.log(JSON.stringify(r));
}
