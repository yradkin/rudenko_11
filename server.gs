/**
 * Опрос собственников дома «г. Гомель, ул. Сергея Руденко, 11» — сервер (Google Apps Script).
 *
 * Хранит ответы в листе «Ответы» этой таблицы. Наружу (на сайт) отдаёт ТОЛЬКО
 * сводку: сколько квартир и собственников ответило и проценты по вариантам.
 * ФИО и телефоны видит только владелец таблицы.
 *
 * Если квартира отвечает повторно — в таблицу пишется новая строка (история сохраняется),
 * а в подсчёте учитывается последний ответ этой квартиры.
 *
 * Развёртывание: Развернуть → Новое развёртывание → Веб-приложение,
 *   «Выполнять как: я», «Доступ: все». Ссылку /exec вставить в index.html (CONFIG.API).
 */
const SHEET = 'Ответы';
const HEAD = ['Получено', 'Квартира', 'Площадь, м²', 'Собственник 1', 'Собственник 2', 'Телефон',
  'Согласие на обработку ПД', 'Способ управления', 'Вступительный взнос', 'Членский взнос, руб/м²',
  'Готов работать в ТС', 'Инициативная группа', 'Комментарий'];

const Q = {
  manage: ['ТС', 'ЖЭУ'],
  entry: ['3 БВ (135 р.)', '4 БВ (180 р.)'],
  fee: ['0,20 руб.', '0,25 руб.'],
  group: ['Да', 'Нет'],
};
const ROLES = ['Председатель', 'Главный бухгалтер', 'Электрик', 'Сантехник', 'Теплотехник', 'Уборщик'];

function sheet_() {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName(SHEET);
  if (!s) {
    s = ss.insertSheet(SHEET);
    s.appendRow(HEAD);
    s.setFrozenRows(1);
    s.getRange(1, 1, 1, HEAD.length).setFontWeight('bold');
  }
  return s;
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// защита от формул в ячейках и обрезка длины
function clean_(v, max) {
  let s = String(v == null ? '' : v).replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, max || 200);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
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
    if (!(apt >= 1 && apt <= 999)) return out_({ ok: false, error: 'Проверьте номер квартиры.' });
    const area = parseFloat(String(d.area).replace(',', '.'));
    if (!(area >= 10 && area <= 300)) return out_({ ok: false, error: 'Проверьте площадь квартиры: от 10 до 300 м².' });
    if (!clean_(d.owner1)) return out_({ ok: false, error: 'Укажите ФИО собственника.' });
    if (d.consent !== true) return out_({ ok: false, error: 'Без согласия на обработку данных ответ не сохраняется.' });
    for (const k in Q) if (d[k] && Q[k].indexOf(d[k]) < 0) return out_({ ok: false, error: 'Неизвестный вариант ответа.' });
    const roles = (Array.isArray(d.roles) ? d.roles : []).filter(r => ROLES.indexOf(r) >= 0);
    const other = clean_(d.roleOther, 120).replace(/;/g, ',');
    if (other) roles.push('Другое: ' + other);
    sheet_().appendRow([new Date(), apt, Math.round(area * 100) / 100, clean_(d.owner1, 120), clean_(d.owner2, 120),
      clean_(d.phone, 40), 'да', d.manage || '', d.entry || '', d.fee || '', roles.join('; '), d.group || '', clean_(d.comment, 1000)]);
    return out_({ ok: true, stats: stats_() });
  } catch (err) {
    return out_({ ok: false, error: 'Не получилось сохранить ответ. Попробуйте ещё раз.' });
  } finally {
    lock.releaseLock();
  }
}

// Сводка без персональных данных: последний ответ каждой квартиры
function stats_() {
  const v = sheet_().getDataRange().getValues().slice(1);
  const last = {};
  v.forEach(r => { if (r[1]) last[r[1]] = r; });
  const rows = Object.keys(last).map(k => last[k]);
  const res = { flats: rows.length, owners: 0, area: 0, q: {}, roles: {}, updated: new Date().toISOString() };
  const col = { manage: 7, entry: 8, fee: 9, group: 11 };
  for (const k in Q) { res.q[k] = { answered: 0, answeredArea: 0, opts: {} }; Q[k].forEach(o => res.q[k].opts[o] = { n: 0, area: 0 }); }
  ROLES.concat(['Другое']).forEach(r => res.roles[r] = 0);
  rows.forEach(r => {
    const area = Number(r[2]) || 0;
    res.owners += (r[3] ? 1 : 0) + (r[4] ? 1 : 0);
    res.area += area;
    for (const k in col) {
      const a = r[col[k]];
      if (a && res.q[k].opts[a]) { res.q[k].opts[a].n++; res.q[k].opts[a].area += area; res.q[k].answered++; res.q[k].answeredArea += area; }
    }
    String(r[10] || '').split('; ').forEach(x => { if (!x) return; const key = x.indexOf('Другое') === 0 ? 'Другое' : x; if (key in res.roles) res.roles[key]++; });
  });
  res.area = Math.round(res.area * 100) / 100;
  return res;
}
