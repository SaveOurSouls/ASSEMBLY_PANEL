/**
 * СОЗДАНИЕ МЕНЮ ПРИ ОТКРЫТИИ ТАБЛИЦЫ
 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('⚙️ Инструменты БД')
    .addItem('🔍 Открыть панель сборщика', 'showSidebar')
    .addToUi();

  SpreadsheetApp.getUi().createMenu('🧮 Пересчет ТВР')
    .addItem('📋 Открыть панель пересчета', 'openTvrRecalcControlSheet')
    .addSeparator()
    .addItem('⏱️ Пересчитать время и стоимость работ', 'recalcTvrLaborColumns')
    .addItem('🧩 Пересчитать стоимость компонентной базы К1-К5', 'recalcTvrSpecColumns')
    .addItem('✅ Пересчитать все колонки ТВР.ЛИСТ', 'recalcTvrAllColumns')
    .addSeparator()
    .addItem('🚚 Миграция: зафиксировать расчетные колонки значениями', 'migrateTvrCalculatedColumnsToValues')
    .addToUi();
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Сборщик спецификаций')
    .setWidth(320); 
  SpreadsheetApp.getUi().showSidebar(html);
}

const TVR_CONTROL_SHEET = 'ТВР.УПРАВЛЕНИЕ';

function openTvrRecalcControlSheet() {
  const controlSheet = ensureTvrRecalcControlSheet_();
  controlSheet.activate();
}

function recalcTvrAllColumns() {
  const startedAt = new Date();
  try {
    const rates = readManualRates_();
    const laborRows = recalcTvrLaborColumnsCore_(rates.rateCh, rates.rateMag);
    const specRows = recalcTvrSpecColumnsCore_();
    const message = 'Пересчет завершен. Трудоемкость: ' + laborRows + ' строк, компонентная база: ' + specRows + ' строк.';
    writeControlStatus_(message, startedAt);
    SpreadsheetApp.getActiveSpreadsheet().toast(message, 'Пересчет ТВР', 8);
    return message;
  } catch (error) {
    writeControlStatus_('Ошибка: ' + error.message, startedAt);
    throw error;
  }
}

function recalcTvrLaborColumns() {
  const startedAt = new Date();
  try {
    const rates = readManualRates_();
    const rows = recalcTvrLaborColumnsCore_(rates.rateCh, rates.rateMag);
    const message = 'Пересчет трудоемкости завершен: ' + rows + ' строк.';
    writeControlStatus_(message, startedAt);
    SpreadsheetApp.getActiveSpreadsheet().toast(message, 'Пересчет ТВР', 8);
    return message;
  } catch (error) {
    writeControlStatus_('Ошибка: ' + error.message, startedAt);
    throw error;
  }
}

function recalcTvrSpecColumns() {
  const startedAt = new Date();
  try {
    const rows = recalcTvrSpecColumnsCore_();
    const message = 'Пересчет компонентной базы завершен: ' + rows + ' строк.';
    writeControlStatus_(message, startedAt);
    SpreadsheetApp.getActiveSpreadsheet().toast(message, 'Пересчет ТВР', 8);
    return message;
  } catch (error) {
    writeControlStatus_('Ошибка: ' + error.message, startedAt);
    throw error;
  }
}

function migrateTvrCalculatedColumnsToValues() {
  const startedAt = new Date();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const listSheet = requireSheet_(ss, 'ТВР.ЛИСТ');
    const rowsCount = Math.max(listSheet.getLastRow() - 1, 0);
    if (rowsCount === 0) {
      const emptyMessage = 'Миграция не требуется: в ТВР.ЛИСТ нет строк данных.';
      writeControlStatus_(emptyMessage, startedAt);
      SpreadsheetApp.getActiveSpreadsheet().toast(emptyMessage, 'Пересчет ТВР', 8);
      return emptyMessage;
    }

    const headerRow = listSheet.getRange(1, 1, 1, listSheet.getLastColumn()).getValues()[0];
    const listCols = mapRequiredColumns_(headerRow, [
      'Время подготовки',
      'Время производства',
      'Цена подготовительных работ',
      'Цена работы',
      'Цена работы_маг',
      'Стоимость компонентной базы К1',
      'Стоимость компонентной базы К2',
      'Стоимость компонентной базы К3',
      'Стоимость компонентной базы К4',
      'Стоимость компонентной базы К5'
    ]);

    SpreadsheetApp.flush();
    freezeColumnsAsValues_(listSheet, rowsCount, listCols, [
      'Время подготовки',
      'Время производства',
      'Цена подготовительных работ',
      'Цена работы',
      'Цена работы_маг',
      'Стоимость компонентной базы К1',
      'Стоимость компонентной базы К2',
      'Стоимость компонентной базы К3',
      'Стоимость компонентной базы К4',
      'Стоимость компонентной базы К5'
    ]);

    const message = 'Миграция завершена: расчетные колонки ТВР.ЛИСТ переведены в значения (' + rowsCount + ' строк).';
    writeControlStatus_(message, startedAt);
    SpreadsheetApp.getActiveSpreadsheet().toast(message, 'Пересчет ТВР', 8);
    return message;
  } catch (error) {
    writeControlStatus_('Ошибка: ' + error.message, startedAt);
    throw error;
  }
}

function recalcTvrLaborColumnsCore_(rateCh, rateMag) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const listSheet = requireSheet_(ss, 'ТВР.ЛИСТ');
  const bdopSheet = requireSheet_(ss, 'ТВР.БДОП');
  const rowsCount = Math.max(listSheet.getLastRow() - 1, 0);
  if (rowsCount === 0) return 0;

  const headerRow = listSheet.getRange(1, 1, 1, listSheet.getLastColumn()).getValues()[0];
  const listCols = mapRequiredColumns_(headerRow, [
    'ID Техкарты',
    'N',
    'L',
    'Время подготовки',
    'Время производства',
    'Цена подготовительных работ',
    'Цена работы',
    'Цена работы_маг'
  ]);

  const techIds = listSheet.getRange(2, listCols['ID Техкарты'] + 1, rowsCount, 1).getValues();
  const nVals = listSheet.getRange(2, listCols['N'] + 1, rowsCount, 1).getValues();
  const lVals = listSheet.getRange(2, listCols['L'] + 1, rowsCount, 1).getValues();
  const bdopData = bdopSheet.getDataRange().getValues();

  const result = CALC_LABOR(techIds, nVals, lVals, rateCh, rateMag, bdopData);
  if (!Array.isArray(result)) {
    throw new Error('CALC_LABOR вернул ошибку: ' + result);
  }
  if (result.length !== rowsCount) {
    throw new Error('CALC_LABOR вернул некорректное количество строк: ' + result.length + ' вместо ' + rowsCount);
  }

  writeColumnsByHeaders_(listSheet, listCols, [
    'Время подготовки',
    'Время производства',
    'Цена подготовительных работ',
    'Цена работы',
    'Цена работы_маг'
  ], result);

  return rowsCount;
}

function recalcTvrSpecColumnsCore_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const listSheet = requireSheet_(ss, 'ТВР.ЛИСТ');
  const specSheet = requireSheet_(ss, 'ТВР.СПЯ');
  const rowsCount = Math.max(listSheet.getLastRow() - 1, 0);
  if (rowsCount === 0) return 0;

  const headerRow = listSheet.getRange(1, 1, 1, listSheet.getLastColumn()).getValues()[0];
  const listCols = mapRequiredColumns_(headerRow, [
    'ID СПЯ',
    'N',
    'L',
    'Кол-во 1',
    'Кол-во 2',
    'Кол-во 3',
    'Кол-во 4',
    'Кол-во 5',
    'Стоимость компонентной базы К1',
    'Стоимость компонентной базы К2',
    'Стоимость компонентной базы К3',
    'Стоимость компонентной базы К4',
    'Стоимость компонентной базы К5'
  ]);

  const specIds = listSheet.getRange(2, listCols['ID СПЯ'] + 1, rowsCount, 1).getValues();
  const nVals = listSheet.getRange(2, listCols['N'] + 1, rowsCount, 1).getValues();
  const lVals = listSheet.getRange(2, listCols['L'] + 1, rowsCount, 1).getValues();
  const qMatrix = extractMatrixByHeaders_(listSheet, rowsCount, listCols, [
    'Кол-во 1',
    'Кол-во 2',
    'Кол-во 3',
    'Кол-во 4',
    'Кол-во 5'
  ]);
  const specData = specSheet.getDataRange().getValues();

  const result = CALC_SPEC(specIds, nVals, lVals, qMatrix, specData);
  if (!Array.isArray(result)) {
    throw new Error('CALC_SPEC вернул ошибку: ' + result);
  }
  if (result.length !== rowsCount) {
    throw new Error('CALC_SPEC вернул некорректное количество строк: ' + result.length + ' вместо ' + rowsCount);
  }

  writeColumnsByHeaders_(listSheet, listCols, [
    'Стоимость компонентной базы К1',
    'Стоимость компонентной базы К2',
    'Стоимость компонентной базы К3',
    'Стоимость компонентной базы К4',
    'Стоимость компонентной базы К5'
  ], result);

  return rowsCount;
}

function ensureTvrRecalcControlSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(TVR_CONTROL_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(TVR_CONTROL_SHEET);
    sheet.getRange('A1').setValue('Панель управляемого пересчета ТВР.ЛИСТ');
    sheet.getRange('A3').setValue('Ставка ЧЛ (0 = взять Уд.Цена ЧЛ из ТВР.БДОП)');
    sheet.getRange('A4').setValue('Ставка ЧЛ_МАГ (0 = взять Уд.Цена ЧЛ_МАГ из ТВР.БДОП)');
    sheet.getRange('A6').setValue('Запуск: меню "🧮 Пересчет ТВР"');
    sheet.getRange('A8').setValue('Последний статус');
    sheet.getRange('B3').setValue(0);
    sheet.getRange('B4').setValue(0);
    sheet.setColumnWidths(1, 2, 500);
    sheet.getRange('A1:A8').setFontWeight('bold');
  }
  return sheet;
}

function readManualRates_() {
  const controlSheet = ensureTvrRecalcControlSheet_();
  return {
    rateCh: parseNumber_(controlSheet.getRange('B3').getValue(), 0),
    rateMag: parseNumber_(controlSheet.getRange('B4').getValue(), 0)
  };
}

function writeControlStatus_(status, startedAt) {
  const controlSheet = ensureTvrRecalcControlSheet_();
  const finishedAt = new Date();
  const elapsedSec = ((finishedAt.getTime() - startedAt.getTime()) / 1000).toFixed(1);
  controlSheet.getRange('A9').setValue('Время запуска');
  controlSheet.getRange('B9').setValue(startedAt);
  controlSheet.getRange('A10').setValue('Время завершения');
  controlSheet.getRange('B10').setValue(finishedAt);
  controlSheet.getRange('A11').setValue('Длительность, сек');
  controlSheet.getRange('B11').setValue(elapsedSec);
  controlSheet.getRange('B8').setValue(status);
}

function requireSheet_(spreadsheet, sheetName) {
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) throw new Error('Лист не найден: ' + sheetName);
  return sheet;
}

function normalizeHeader_(value) {
  return String(value || '').toLowerCase().replace(/["']/g, '').replace(/\s+/g, ' ').trim();
}

function parseNumber_(value, fallback) {
  const parsed = parseFloat(String(value).replace(',', '.'));
  return isNaN(parsed) ? fallback : parsed;
}

function mapRequiredColumns_(headerRow, requiredHeaders) {
  const normalizedMap = new Map();
  for (let i = 0; i < headerRow.length; i++) {
    normalizedMap.set(normalizeHeader_(headerRow[i]), i);
  }
  const output = {};
  for (let i = 0; i < requiredHeaders.length; i++) {
    const key = requiredHeaders[i];
    const idx = normalizedMap.get(normalizeHeader_(key));
    if (idx === undefined) throw new Error('Колонка не найдена: ' + key);
    output[key] = idx;
  }
  return output;
}

function extractMatrixByHeaders_(sheet, rowsCount, colMap, headers) {
  const matrix = new Array(rowsCount);
  for (let row = 0; row < rowsCount; row++) matrix[row] = new Array(headers.length);
  const groups = buildColumnGroups_(colMap, headers);
  for (let g = 0; g < groups.length; g++) {
    const group = groups[g];
    const values = sheet.getRange(2, group.startCol + 1, rowsCount, group.headers.length).getValues();
    for (let row = 0; row < rowsCount; row++) {
      for (let c = 0; c < group.headers.length; c++) {
        matrix[row][group.sourceIndexes[c]] = values[row][c];
      }
    }
  }
  return matrix;
}

function writeColumnsByHeaders_(sheet, colMap, targetHeaders, values) {
  const groups = buildColumnGroups_(colMap, targetHeaders);
  for (let g = 0; g < groups.length; g++) {
    const group = groups[g];
    const blockValues = new Array(values.length);
    for (let row = 0; row < values.length; row++) {
      const outRow = new Array(group.headers.length);
      for (let c = 0; c < group.headers.length; c++) {
        outRow[c] = values[row][group.sourceIndexes[c]];
      }
      blockValues[row] = outRow;
    }
    sheet.getRange(2, group.startCol + 1, values.length, group.headers.length).setValues(blockValues);
  }
}

function freezeColumnsAsValues_(sheet, rowsCount, colMap, targetHeaders) {
  const groups = buildColumnGroups_(colMap, targetHeaders);
  for (let g = 0; g < groups.length; g++) {
    const group = groups[g];
    const range = sheet.getRange(2, group.startCol + 1, rowsCount, group.headers.length);
    range.setValues(range.getValues());
  }
}

function buildColumnGroups_(colMap, headers) {
  const columns = headers.map(function(header, sourceIndex) {
    return {
      header: header,
      sourceIndex: sourceIndex,
      col: colMap[header]
    };
  }).sort(function(a, b) { return a.col - b.col; });

  const groups = [];
  let current = null;

  for (let i = 0; i < columns.length; i++) {
    const item = columns[i];
    if (!current || item.col !== current.lastCol + 1) {
      if (current) groups.push(current);
      current = {
        startCol: item.col,
        lastCol: item.col,
        headers: [item.header],
        sourceIndexes: [item.sourceIndex]
      };
    } else {
      current.lastCol = item.col;
      current.headers.push(item.header);
      current.sourceIndexes.push(item.sourceIndex);
    }
  }

  if (current) groups.push(current);
  return groups;
}

// НОВАЯ ЗАГРУЗКА БАЗЫ (Поиск по тегам на множестве листов + Удаление дублей)
function getSearchData() {
  try {
    var sheetsToSearch = ['COMPCON', 'COMPCOAX', 'COMPTERM', 'COMPWIRE', 'COMPACCESS'];
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var allValues = [];
    var uniqueSet = {}; // Словарь для отсева дубликатов

    sheetsToSearch.forEach(function(sheetName) {
      var sheet = ss.getSheetByName(sheetName);
      if (!sheet) return; // Пропускаем, если такого листа нет

      var data = sheet.getDataRange().getValues();
      if (data.length < 2) return;

      var headers = data[0];
      var tagColIdx = -1;

      // Ищем колонку "Поисковый тег"
      for (var i = 0; i < headers.length; i++) {
        if (headers[i].toString().toLowerCase().trim() === 'поисковый тег') {
          tagColIdx = i;
          break;
        }
      }

      // Если колонка найдена, собираем из неё данные
      if (tagColIdx !== -1) {
        for (var r = 1; r < data.length; r++) {
          var val = data[r][tagColIdx] ? data[r][tagColIdx].toString().trim() : "";
          // Если значение есть и мы его еще не встречали - добавляем
          if (val && !uniqueSet[val]) {
            uniqueSet[val] = true;
            allValues.push(val);
          }
        }
      }
    });

    if (allValues.length === 0) return "ОШИБКА: Колонка 'Поисковый тег' не найдена или пуста на указанных листах.";

    return allValues.join("|||");
  } catch (e) {
    return "ОШИБКА: " + e.message;
  }
}

// Стандартная выгрузка корзины
function insertBatchIntoCell(valuesArray) {
  if (!valuesArray || valuesArray.length === 0) return "Список пуст";

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var startRow = ss.getCurrentCell().getRow();
  var startCol = ss.getCurrentCell().getColumn(); 
  
  if (startRow < 2) return "Выберите строку в таблице ниже шапки";

  var numRows = valuesArray.length;
  var matrix = valuesArray.map(function(item) { return [item]; });

  try {
    sheet.getRange(startRow, startCol, numRows, 1).setValues(matrix);
    sheet.getRange(startRow + numRows, startCol).activate();
    return "Успешно добавлено " + numRows + " поз.";
  } catch (e) {
    return "ОШИБКА: " + e.message;
  }
}

// Выгрузка реплицированных (размноженных) данных
function insertReplicatedData(valuesArray, targetCellA1) {
  if (!valuesArray || valuesArray.length === 0) return "Нет данных для выгрузки";
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var startCell;

  if (targetCellA1) {
    try {
      startCell = sheet.getRange(targetCellA1);
    } catch (e) {
      return "ОШИБКА: Неверный адрес ячейки (" + targetCellA1 + "). Ожидается формат A2, C5 и т.д.";
    }
  } else {
    startCell = sheet.getActiveCell();
  }

  var numRows = valuesArray.length;
  var matrix = valuesArray.map(function(val) { return [val]; });

  try {
    sheet.getRange(startCell.getRow(), startCell.getColumn(), numRows, 1).setValues(matrix);
    return "Успешно выгружено " + numRows + " строк.";
  } catch (e) {
    return "ОШИБКА: " + e.message;
  }
}
/**
 * МАКСИМАЛЬНО УСКОРЕННАЯ ВЕРСИЯ: Использование нативных циклов for и математического округления.
 * @customfunction
 */
function CALC_SPEC(ids, n_vals, l_vals, q_matrix, spec_all_data) {
  if (!ids || !spec_all_data || spec_all_data.length < 1) return "Нет данных";

  // 1) Находим строку заголовков по маркеру "ID СПЯ"
  let headerRowIndex = -1;
  let headers = [];
  for (let i = 0; i < spec_all_data.length; i++) {
    const rowClean = spec_all_data[i].map(normalizeHeader_);
    if (rowClean.indexOf("id спя") !== -1) {
      headerRowIndex = i;
      headers = rowClean;
      break;
    }
  }

  let isHeaderFound = true;
  if (headerRowIndex === -1) {
    headerRowIndex = 0;
    headers = spec_all_data[0].map(normalizeHeader_);
    isHeaderFound = false;
  }

  let col_id = headers.indexOf("id спя");
  let col_type = headers.indexOf("тип");
  let col_itog = headers.indexOf("итог");
  let col_prices = [
    "цена закупки 1",
    "цена закупки 2",
    "цена закупки 3",
    "цена закупки 4",
    "цена закупки 5"
  ].map(function(n) { return headers.indexOf(n); });
  let col_qs = [
    "кол-во 1",
    "кол-во 2",
    "кол-во 3",
    "кол-во 4",
    "кол-во 5"
  ].map(function(n) { return headers.indexOf(n); });

  // Резервная схема индексов на случай нестандартной/сдвинутой шапки
  const hasMissingColumns = (
    col_id === -1 || col_type === -1 || col_itog === -1 ||
    col_prices.some(function(idx) { return idx === -1; }) ||
    col_qs.some(function(idx) { return idx === -1; })
  );
  if (!isHeaderFound || hasMissingColumns) {
    col_id = 2;      // ID СПЯ
    col_type = 6;    // Тип
    col_itog = 11;   // Итог
    col_prices = [13, 14, 15, 16, 17]; // Цена закупки 1..5
    col_qs = [18, 19, 20, 21, 22];     // Кол-во 1..5
  }

  const normalizeSpecId = function(str) {
    return String(str).replace(/\s+/g, '').toLowerCase()
      .replace(/х/g, 'x').replace(/с/g, 'c').replace(/а/g, 'a').replace(/е/g, 'e').replace(/о/g, 'o');
  };

  // 2. Хешируем базу СПЯ за один проход
  const specMap = new Map();
  const dataLen = spec_all_data.length;
  const startRow = isHeaderFound ? headerRowIndex + 1 : 0;
  
  for (let i = startRow; i < dataLen; i++) {
    const r = spec_all_data[i];
    if (!r || r.length === 0) continue;
    const id = normalizeSpecId(r[col_id]);
    if (!id) continue;
    
    const t = String(r[col_type]).toLowerCase();
    let typeFlag = 0; 
    if (t.includes("провод") || t.includes("кабель") || t.includes("пв-") || t.includes("мгтф") || t.includes("awg")) {
      typeFlag = 1;
    } else if (t.includes("контакт") || t.includes("терминал") || t.includes("гнездо") || t.includes("штырь") || t.includes("terminal")) {
      typeFlag = 2;
    } else if (t.includes("лента") || t.includes("тут") || t.includes("плетенка") || t.includes("трубка") || t.includes("изолента")) {
      typeFlag = 3;
    }

    if (!specMap.has(id)) specMap.set(id, []);
    specMap.get(id).push({
      typeFlag: typeFlag,
      itog: parseFloat(String(r[col_itog]).replace(',', '.')) || 0,
      p: col_prices.map(idx => parseFloat(String(r[idx]).replace(',', '.')) || 0),
      q: col_qs.map(idx => parseFloat(String(r[idx]).replace(',', '.')) || 0)
    });
  }
  if (specMap.size === 0) {
    return "ОШИБКА CALC_SPEC: база ТВР.СПЯ не распознана (проверьте заголовки/данные ID СПЯ).";
  }

  // Заранее парсим N и L в плоские массивы чисел
  const rowsCount = ids.length;
  const parsedN = new Float64Array(rowsCount);
  const parsedL = new Float64Array(rowsCount);
  const cleanIds = new Array(rowsCount);
  let nonEmptyIds = 0;

  for (let i = 0; i < rowsCount; i++) {
    cleanIds[i] = normalizeSpecId(ids[i][0]);
    parsedN[i] = parseFloat(String(n_vals[i][0]).replace(',', '.')) || 1;
    parsedL[i] = parseFloat(String(l_vals[i][0]).replace(',', '.')) || 1;
    if (cleanIds[i]) nonEmptyIds++;
  }

  const result = new Array(rowsCount);
  let matchedRows = 0;

  // 3. Главный расчет на быстрых циклах for
  for (let rIdx = 0; rIdx < rowsCount; rIdx++) {
    const id = cleanIds[rIdx];
    if (!id || !specMap.has(id)) {
      result[rIdx] = [0, 0, 0, 0, 0];
      continue;
    }
    matchedRows++;

    const N = parsedN[rIdx];
    const L = parsedL[rIdx];
    const items = specMap.get(id);
    const itemsLen = items.length;
    const rowQs = q_matrix[rIdx];
    const qsLen = rowQs.length;
    
    const rowResult = new Array(qsLen);

    for (let qIdx = 0; qIdx < qsLen; qIdx++) {
      const orderQ = parseFloat(String(rowQs[qIdx]).replace(',', '.')) || 0;
      let totalRowCost = 0;

      for (let i = 0; i < itemsLen; i++) {
        const item = items[i];
        let consumptionPerUnit = item.itog;

        if (item.typeFlag === 1) {
          consumptionPerUnit = item.itog * L * N;
        } else if (item.typeFlag === 2) {
          consumptionPerUnit = item.itog * N;
        } else if (item.typeFlag === 3) {
          consumptionPerUnit = item.itog * L;
        }

        const totalNeededForOrder = consumptionPerUnit * orderQ;
        let price = item.p[0];

        // Быстрое каскадное сравнение цен
        const iq = item.q;
        const ip = item.p;
        if (totalNeededForOrder >= iq[4] && iq[4] > 0) price = ip[4];
        else if (totalNeededForOrder >= iq[3] && iq[3] > 0) price = ip[3];
        else if (totalNeededForOrder >= iq[2] && iq[2] > 0) price = ip[2];
        else if (totalNeededForOrder >= iq[1] && iq[1] > 0) price = ip[1];

        totalRowCost += (consumptionPerUnit * price);
      }

      // Быстрое математическое округление без перевода в строку
      rowResult[qIdx] = Math.round(totalRowCost * 100) / 100;
    }
    
    result[rIdx] = rowResult;
  }
  if (nonEmptyIds > 0 && matchedRows === 0) {
    return "ОШИБКА CALC_SPEC: ID СПЯ из ТВР.ЛИСТ не совпали с ТВР.СПЯ.";
  }

  return result;
}


/**
 * ДИНАМИЧЕСКИЙ РАСЧЕТ ТРУДОЕМКОСТИ: Полная производственная логика (версия 2.0)
 * @customfunction
 */
function CALC_LABOR(tech_ids, n_vals, l_vals, rate_ch, rate_mag, bdop_all_data) {
  if (!tech_ids || !bdop_all_data || bdop_all_data.length < 1) return "Нет данных";

  const parseScalarRate = (value) => {
    if (Array.isArray(value)) {
      const first = Array.isArray(value[0]) ? value[0][0] : value[0];
      return parseFloat(String(first).replace(',', '.')) || 0;
    }
    return parseFloat(String(value).replace(',', '.')) || 0;
  };

  const externalRateCh = parseScalarRate(rate_ch);
  const externalRateMag = parseScalarRate(rate_mag);

  // 1. Поиск строки заголовков по ключевому маркеру
  let headerRowIndex = -1;
  let headers = [];
  const bdopLen = bdop_all_data.length;
  
  for (let i = 0; i < bdopLen; i++) {
    const rowClean = bdop_all_data[i].map(normalizeHeader_);
    if (rowClean.indexOf("id техкарты") !== -1) {
      headerRowIndex = i;
      headers = rowClean;
      break;
    }
  }

  let isHeaderFound = true;
  if (headerRowIndex === -1) {
    headerRowIndex = 0; 
    headers = bdop_all_data[0].map(normalizeHeader_);
    isHeaderFound = false;
  }
  
  // Динамическое сопоставление колонок по точным именам из новой структуры
  const col = {
    id: headers.indexOf("id техкарты"),
    type: headers.indexOf("тип операции"),
    t_op: headers.indexOf("время операции"),
    t_prep: headers.indexOf("время подготовки, сек"),
    t_mach: headers.indexOf("время машины, сек/оп; сек/м"),
    p_ch: headers.indexOf("уд.цена чл, сек"),
    p_mag: headers.indexOf("уд.цена чл_маг, сек"),
    p_msh: headers.indexOf("уд.цена мш, сек"),
    n_per_product: headers.indexOf("кол-во операций на продукцию"),
    operators: headers.indexOf("кол-во операторов"),
    // Погонные параметры
    v_prokat: headers.indexOf("скорость проката"),
    v_tool: headers.indexOf("скорость работы инструмента"),
    n_tool: headers.indexOf("кол-во операций инструмента"),
    // Составляющие ручных работ для Статичного/Переменного типов
    h_take: headers.indexOf("время ручных работ для взятия полуфабриката"),
    h_work: headers.indexOf("время ручных работ"),
    h_drop: headers.indexOf("время ручных работ для снятия полуфабриката")
  };

  // Резервный план Б по точным номерам колонок, если шапка отсечена платформой Google
  if (!isHeaderFound || col.t_op === -1 || col.t_prep === -1) {
    col.id = 0;       // ID техкарты
    col.t_op = 15;    // Время Операции
    col.t_prep = 16;  // Время подготовки, сек
    col.t_mach = 19;  // Время машины, сек/оп; сек/м
    col.p_ch = 20;    // Уд.Цена ЧЛ, сек
    col.p_mag = 21;   // Уд.Цена ЧЛ_МАГ, сек
    col.p_msh = 22;   // Уд.Цена МШ, сек
    col.n_per_product = 7; // Кол-во операций на продукцию
    col.operators = 8;     // Кол-во операторов
    col.type = 23;    // Тип операции
    col.h_take = 24;  // Время ручных работ для взятия...
    col.h_work = 25;  // Время ручных работ
    col.h_drop = 26;  // Время ручных работ для снятия...
    col.v_prokat = 27;// Скорость проката
    col.v_tool = 28;  // Скорость работы инструмента
    col.n_tool = 29;  // Кол-во операций инструмента
  }

  const normalizeId = (str) => {
    return String(str).replace(/\s+/g, '').toLowerCase()
      .replace(/х/g, 'x').replace(/с/g, 'c').replace(/а/g, 'a').replace(/е/g, 'e').replace(/о/g, 'o');
  };

  // 2. Хешируем базу ТВР.БДОП в оперативную память
  const bdopMap = new Map();
  const startRow = isHeaderFound ? headerRowIndex + 1 : 0;

  for (let i = startRow; i < bdopLen; i++) {
    const r = bdop_all_data[i];
    if (!r || r.length === 0 || r[col.id] === "") continue;
    
    const id = normalizeId(r[col.id]);
    if (!id || id === "undefined" || id === "null") continue;

    // Определение типа операции (Регистронезависимо)
    let typeStr = "статичный"; 
    if (col.type !== -1 && col.type < r.length) {
      const t = String(r[col.type]).toLowerCase().trim();
      if (t.includes("погон")) typeStr = "погонный";
      else if (t.includes("перемен")) typeStr = "переменный";
    }

    // Чтение базовых числовых значений
    const t_prep = col.t_prep < r.length ? (parseFloat(String(r[col.t_prep]).replace(',', '.')) || 0) : 0;
    const t_op = col.t_op < r.length ? (parseFloat(String(r[col.t_op]).replace(',', '.')) || 0) : 0;
    const t_mach = col.t_mach < r.length ? (parseFloat(String(r[col.t_mach]).replace(',', '.')) || 0) : 0;
    
    // Тарифные сетки из БДОП
    const p_ch = col.p_ch < r.length ? (parseFloat(String(r[col.p_ch]).replace(',', '.')) || 0) : 0;
    const p_mag = col.p_mag < r.length ? (parseFloat(String(r[col.p_mag]).replace(',', '.')) || 0) : 0;
    const p_msh = col.p_msh < r.length ? (parseFloat(String(r[col.p_msh]).replace(',', '.')) || 0) : 0;
    const n_per_product = col.n_per_product < r.length ? (parseFloat(String(r[col.n_per_product]).replace(',', '.')) || 1) : 1;
    const operators = col.operators < r.length ? (parseFloat(String(r[col.operators]).replace(',', '.')) || 1) : 1;

    // Параметры для погонного расчета
    const v_prokat = col.v_prokat < r.length ? (parseFloat(String(r[col.v_prokat]).replace(',', '.')) || 0) : 0;
    const v_tool = col.v_tool < r.length ? (parseFloat(String(r[col.v_tool]).replace(',', '.')) || 0) : 0;
    const n_tool = col.n_tool < r.length ? (parseFloat(String(r[col.n_tool]).replace(',', '.')) || 0) : 0;

    // Составляющие ручной работы
    const h_take = col.h_take < r.length ? (parseFloat(String(r[col.h_take]).replace(',', '.')) || 0) : 0;
    const h_work = col.h_work < r.length ? (parseFloat(String(r[col.h_work]).replace(',', '.')) || 0) : 0;
    const h_drop = col.h_drop < r.length ? (parseFloat(String(r[col.h_drop]).replace(',', '.')) || 0) : 0;

    if (!bdopMap.has(id)) {
      bdopMap.set(id, {
        total_prep_time: 0,
        operations: []
      });
    }

    const dataObj = bdopMap.get(id);
    dataObj.total_prep_time += t_prep;

    dataObj.operations.push({
      type: typeStr,
      t_op: t_op,
      t_mach: t_mach,
      p_ch: p_ch,
      p_mag: p_mag,
      p_msh: p_msh,
      n_per_product: n_per_product,
      operators: operators,
      v_prokat: v_prokat,
      v_tool: v_tool,
      n_tool: n_tool,
      sum_manual: (h_take + h_work + h_drop)
    });
  }

  // 3. Основной расчет по строкам ТВР.ЛИСТ
  const rowsCount = tech_ids.length;
  const result = new Array(rowsCount);

  for (let i = 0; i < rowsCount; i++) {
    const id = normalizeId(tech_ids[i][0]);
    
    if (!id) {
      result[i] = ["", "", "", "", ""];
      continue;
    }
    if (!bdopMap.has(id)) {
      result[i] = ["нет техкарты", "", "", "", ""];
      continue;
    }

    const N = parseFloat(String(n_vals[i][0]).replace(',', '.')) || 1;
    const L = parseFloat(String(l_vals[i][0]).replace(',', '.')) || 1;
    const card = bdopMap.get(id);

    let total_prod_time = 0;  // Итоговое время производства (человек + машина)
    let price_work = 0;       // Итоговая цена работы
    let price_work_mag = 0;   // Итоговая цена работы_маг

    const opsLen = card.operations.length;
    
    for (let j = 0; j < opsLen; j++) {
      const op = card.operations[j];
      
      // Назначение тарифных ставок (с приоритетом внешних параметров листа)
      const current_rate_ch = externalRateCh > 0 ? externalRateCh : op.p_ch;
      const current_rate_mag = externalRateMag > 0 ? externalRateMag : op.p_mag;
      const operationsPerProduct = op.n_per_product > 0 ? op.n_per_product : 1;
      const operatorsCount = op.operators > 0 ? op.operators : 1;
      let op_time = 0;

      if (op.type === "погоннный" || op.type === "погонный") {
        op_time = ((op.v_prokat * L) + (op.v_tool * op.n_tool)) * N * operationsPerProduct;
      } else {
        op_time = op.sum_manual * N * operationsPerProduct;
      }

      total_prod_time += op_time;
      price_work += op_time * (current_rate_ch + op.p_msh) * operatorsCount;
      price_work_mag += op_time * (current_rate_mag + op.p_msh) * operatorsCount;
    }

    // Итоговая цена подготовки
    const current_prep_rate = externalRateCh > 0 ? externalRateCh : (card.operations[0]?.p_ch || 0);
    const price_prep = card.total_prep_time * current_prep_rate;

    result[i] = [
      Math.round(card.total_prep_time * 100) / 100,
      Math.round(total_prod_time * 100) / 100,
      Math.round(price_prep * 100) / 100,
      Math.round(price_work * 100) / 100,
      Math.round(price_work_mag * 100) / 100
    ];
  }

  return result;
}
