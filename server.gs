/**
 * Опросник собственников дома «г. Гомель, ул. Сергея Руденко, 11» — сервер (Google Apps Script).
 *
 * Версия без личных данных: номер квартиры, площадь, число собственников, способ управления,
 * готовность войти в инициативную группу. ФИО и телефоны НЕ собираются.
 * Номер квартиры нужен только чтобы засчитать один ответ на квартиру; на сайт он не отдаётся.
 *
 * Ответы пишутся в лист «Опрос». Если квартира отвечает повторно — пишется новая строка,
 * в подсчёте учитывается последний ответ квартиры.
 * Наружу (doGet) отдаётся только сводка: по дому и по подъездам.
 */
const SHEET = 'Опрос';
const HEAD = ['Получено', 'Квартира', 'Подъезд', 'Площадь, м²', 'Собственников', 'Способ управления', 'Инициативная группа'];

// Подъезды: диапазоны квартир (по проектной декларации: 14-этажная секция — кв. 1–56, три 11-этажные — по 40 квартир)
const SECTIONS = [
  { n: 1, from: 1, to: 56 },
  { n: 2, from: 57, to: 96 },
  { n: 3, from: 97, to: 136 },
  { n: 4, from: 137, to: 176 },
];
const Q = {
  manage: ['ТС', 'ЖЭУ'],
  group: ['Да', 'Нет', 'Не определился'],
};
const MIN_SHOW = 5;   // разбивку по подъезду показываем, когда ответили хотя бы 5 квартир

const sectionOf = apt => { const s = SECTIONS.find(s => apt >= s.from && apt <= s.to); return s ? s.n : 0; };

function sheet_() {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName(SHEET);
  if (!s) {
    s = ss.insertSheet(SHEET, 0);
    s.appendRow(HEAD);
    s.setFrozenRows(1);
    s.getRange(1, 1, 1, HEAD.length).setFontWeight('bold');
  }
  return s;
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return out_(stats_());
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const d = JSON.parse(e.postData.contents || '{}');
    if (d.website) return out_({ ok: true, stats: stats_() });            // ловушка для ботов
    const apt = parseInt(String(d.apt).replace(/\D/g, ''), 10);
    const sec = sectionOf(apt);
    if (!sec) return out_({ ok: false, error: 'В доме квартиры с 1 по 176 — проверьте номер.' });
    const area = parseFloat(String(d.area).replace(',', '.'));
    if (!(area >= 20 && area <= 120)) return out_({ ok: false, error: 'Проверьте площадь: от 20 до 120 м², например 67,19.' });
    const owners = parseInt(d.owners, 10);
    if (owners !== 1 && owners !== 2) return out_({ ok: false, error: 'Укажите, сколько собственников: 1 или 2.' });
    if (Q.manage.indexOf(d.manage) < 0) return out_({ ok: false, error: 'Выберите способ управления.' });
    if (d.group && Q.group.indexOf(d.group) < 0) return out_({ ok: false, error: 'Неизвестный вариант ответа.' });
    sheet_().appendRow([new Date(), apt, sec, Math.round(area * 100) / 100, owners, d.manage, d.group || '']);
    return out_({ ok: true, stats: stats_() });
  } catch (err) {
    return out_({ ok: false, error: 'Не получилось сохранить ответ. Попробуйте ещё раз.' });
  } finally {
    lock.releaseLock();
  }
}

// Сводка без номеров квартир: последний ответ каждой квартиры
function stats_() {
  const v = sheet_().getDataRange().getValues().slice(1);
  const last = {};
  v.forEach(r => { if (r[1]) last[r[1]] = r; });
  const rows = Object.keys(last).map(k => last[k]);
  const blank = () => ({ flats: 0, owners: 0, area: 0, manage: {}, group: {} });
  const fillOpts = o => { Q.manage.forEach(k => o.manage[k] = { n: 0, area: 0 }); Q.group.forEach(k => o.group[k] = { n: 0, area: 0 }); return o; };
  const all = fillOpts(blank());
  const secs = SECTIONS.map(s => Object.assign(fillOpts(blank()), { n: s.n, from: s.from, to: s.to, total: s.to - s.from + 1 }));
  rows.forEach(r => {
    const area = Number(r[3]) || 0, owners = Number(r[4]) || 0, sec = secs[(Number(r[2]) || sectionOf(Number(r[1]))) - 1];
    [all, sec].forEach(t => {
      if (!t) return;
      t.flats++; t.owners += owners; t.area += area;
      if (t.manage[r[5]]) { t.manage[r[5]].n++; t.manage[r[5]].area += area; }
      if (t.group[r[6]]) { t.group[r[6]].n++; t.group[r[6]].area += area; }
    });
  });
  // в подъездах с малым числом ответов распределение не раскрываем
  secs.forEach(s => { s.area = Math.round(s.area * 100) / 100; if (s.flats < MIN_SHOW) { s.hidden = true; s.manage = null; s.group = null; } });
  all.area = Math.round(all.area * 100) / 100;
  return { total: 176, minShow: MIN_SHOW, all: all, sections: secs, updated: new Date().toISOString() };
}
