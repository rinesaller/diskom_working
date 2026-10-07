// ============================================================
// E-commerce Project 1.0 — Google Apps Script (Backend)
// Лист с данными переименован: "Data" -> "Таблица"
// Колонки листа: A = ID (уникальный), W = ID Связи (ссылка на ID родителя)
// ============================================================

const SHEET_NAME = 'Таблица';          // новое имя листа (было 'Data')
const SHEET_NAME_OLD = 'Data';         // совместимость со старыми таблицами
const REF_SHEET_NAME = 'Справочник';

// Полный ожидаемый набор колонок (A = ID, B = № структуры, W = ID Связи)
const EXPECTED_HEADERS = [
  'ID', '№', 'Этап', 'Задача', 'Под-задача', 'Владелец результата', 'Ответственный',
  'Решающий', 'Информируемые', 'Вход инф.', 'Вход (ID)', 'Выход инф.', 'Выход (ID)',
  'Gate', 'Старт', 'Финиш', 'Статус', 'Блокер', 'Кем заблок.', 'С какого числа',
  'Документ', 'Прогресс', 'Комментарий', 'ID Связи'
];

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('E-commerce Project 1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function ping() {
  return "OK: " + new Date().toLocaleTimeString();
}

// Возвращает лист с данными: сначала ищем "Таблица", затем старый "Data"
function getDataSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(SHEET_NAME) || ss.getSheetByName(SHEET_NAME_OLD);
}

function getProjectData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getDataSheet_();

    if (!sheet) {
      return { _error: `Лист "${SHEET_NAME}" не найден. Переименуйте лист "Data" в "${SHEET_NAME}".` };
    }

    // getDisplayValues() — все данные как текст ("01.07.2026"), это снимает проблемы с датами
    const data = sheet.getDataRange().getDisplayValues();

    if (data.length < 2) {
      return { _error: 'Лист пуст или содержит только заголовки.' };
    }

    const headers = data[0].map(h => String(h || '').trim());

    // Колонка B «№» — номер строки для структурирования иерархии этапов/задач/подзадач.
    // Если колонки нет или значение пустое — используем порядковый номер строки данных.
    const numIdx = headers.indexOf('№');

    const rows = data.slice(1).map((row, index) => {
      let obj = { _rowIndex: index + 2 };
      headers.forEach((h, i) => {
        obj[h] = row[i];
      });
      if (numIdx >= 0 && String(row[numIdx] || '').trim()) {
        obj._num = String(row[numIdx]).trim();
      } else {
        obj._num = String(index + 1);
      }
      return obj;
    });

    return {
      headers: headers,
      rows: rows,
      options: buildOptions(ss, rows),
      _debug: `Успех! Загружено строк: ${rows.length}`
    };

  } catch (e) {
    return { _error: 'Критическая ошибка сервера: ' + e.message };
  }
}

function buildOptions(ss, rows) {
  const defaults = { Отдел: [], ФИО: [], Ответ: [], Статус: ['Планируется', 'В процессе', 'Готово'], Прогресс: [] };

  try {
    const refSheet = ss.getSheetByName(REF_SHEET_NAME);
    if (refSheet && refSheet.getLastRow() > 1) {
      const data = refSheet.getDataRange().getDisplayValues();
      const refHeaders = data[0].map(h => String(h || '').trim());
      const body = data.slice(1);

      refHeaders.forEach((h, colIdx) => {
        if (!h || !defaults.hasOwnProperty(h)) return;
        const vals = body
          .map(r => String(r[colIdx] || '').trim())
          .filter(v => v !== '');
        defaults[h] = [...new Set(vals)];
      });
    }
  } catch (e) {
    console.error("Ошибка справочника:", e);
  }

  if (!defaults.Отдел.includes('Все')) defaults.Отдел.push('Все');

  defaults.Прогресс = defaults.Прогресс
    .filter(v => /^\d+%$/.test(v))
    .sort((a, b) => parseInt(a) - parseInt(b));
  if (!defaults.Прогресс.length) {
    defaults.Прогресс = ['0%','10%','20%','30%','40%','50%','60%','70%','80%','90%','100%'];
  }

  // Уникальные ID из колонки A — используются для выпадающего списка "ID Связи"
  const ids = [...new Set(rows.map(r => String(r['ID'] || '').trim()).filter(v => v !== ''))];

  return {
    departments: defaults.Отдел,
    employees: defaults.ФИО,
    decision: defaults.Ответ,
    statuses: defaults.Статус,
    progress: defaults.Прогресс,
    ids: ids
  };
}

function updateCell(rowIndex, columnName, newValue) {
  try {
    const sheet = getDataSheet_();
    if (!sheet) return { success: false, message: 'Лист не найден' };

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0].map(h => String(h || '').trim());
    const colIndex = headers.indexOf(columnName) + 1;

    if (colIndex > 0) {
      sheet.getRange(rowIndex, colIndex).setValue(newValue);
      return { success: true };
    }
    return { success: false, message: 'Колонка не найдена: ' + columnName };
  } catch (e) {
    return { success: false, message: e.message };
  }
}
