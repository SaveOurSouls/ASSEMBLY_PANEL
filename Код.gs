/**
 * СОЗДАНИЕ МЕНЮ ПРИ ОТКРЫТИИ ТАБЛИЦЫ
 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('⚙️ Инструменты БД')
    .addItem('🔍 Открыть панель сборщика', 'showSidebar')
    .addToUi();
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Сборщик спецификаций')
    .setWidth(320); 
  SpreadsheetApp.getUi().showSidebar(html);
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
  if (!ids || !spec_all_data) return "Нет данных";

  const headers = spec_all_data[0];
  const col_id = headers.indexOf("ID СПЯ");
  const col_type = headers.indexOf("Тип");
  const col_itog = headers.indexOf("Итог");
  
  const col_prices = ["Цена закупки 1", "Цена закупки 2", "Цена закупки 3", "Цена закупки 4", "Цена закупки 5"].map(n => headers.indexOf(n));
  const col_qs = ["Кол-во 1", "Кол-во 2", "Кол-во 3", "Кол-во 4", "Кол-во 5"].map(n => headers.indexOf(n));

  // 1. Хешируем базу СПЯ за один проход
  const specMap = new Map();
  const dataLen = spec_all_data.length;
  
  for (let i = 1; i < dataLen; i++) {
    const r = spec_all_data[i];
    const id = String(r[col_id]).trim();
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

  // Заранее парсим N и L в плоские массивы чисел
  const rowsCount = ids.length;
  const parsedN = new Float64Array(rowsCount);
  const parsedL = new Float64Array(rowsCount);
  const cleanIds = new Array(rowsCount);

  for (let i = 0; i < rowsCount; i++) {
    cleanIds[i] = String(ids[i][0]).trim();
    parsedN[i] = parseFloat(String(n_vals[i][0]).replace(',', '.')) || 1;
    parsedL[i] = parseFloat(String(l_vals[i][0]).replace(',', '.')) || 1;
  }

  const result = new Array(rowsCount);

  // 2. Главный расчет на быстрых циклах for
  for (let rIdx = 0; rIdx < rowsCount; rIdx++) {
    const id = cleanIds[rIdx];
    if (!id || !specMap.has(id)) {
      result[rIdx] = [0, 0, 0, 0, 0];
      continue;
    }

    const N = parsedN[rIdx];
    const L = parsedL[rIdx];
    const items = specMap.get(id);
    const itemsLen = items.length;
    const rowQs = q_matrix[rIdx];
    const qsLen = rowQs.length;
    
    const rowResult = new Array(qsLen);

    for (let qIdx = 0; qIdx < qsLen; qIdx++) {
      const orderQ = parseFloat(rowQs[qIdx]) || 0;
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

  return result;
}


/**
 * ДИНАМИЧЕСКИЙ РАСЧЕТ ТРУДОЕМКОСТИ: Полная производственная логика (версия 2.0)
 * @customfunction
 */
function CALC_LABOR(tech_ids, n_vals, l_vals, rate_ch, rate_mag, bdop_all_data) {
  if (!tech_ids || !bdop_all_data || bdop_all_data.length < 1) return "Нет данных";

  const cleanHeader = (str) => {
    return String(str).toLowerCase().replace(/["']/g, '').replace(/\s+/g, ' ').trim();
  };

  // 1. Поиск строки заголовков по ключевому маркеру
  let headerRowIndex = -1;
  let headers = [];
  const bdopLen = bdop_all_data.length;
  
  for (let i = 0; i < bdopLen; i++) {
    const rowClean = bdop_all_data[i].map(cleanHeader);
    if (rowClean.indexOf("id техкарты") !== -1) {
      headerRowIndex = i;
      headers = rowClean;
      break;
    }
  }

  let isHeaderFound = true;
  if (headerRowIndex === -1) {
    headerRowIndex = 0; 
    headers = bdop_all_data[0].map(cleanHeader);
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
      const current_rate_ch = rate_ch > 0 ? rate_ch : op.p_ch;
      const current_rate_mag = rate_mag > 0 ? rate_mag : op.p_mag;

      if (op.type === "погоннный" || op.type === "погонный") {
        // --- 1. ПОГОННАЯ ОПЕРАЦИЯ ---
        // Время = (L * Скорость проката + Скорость инструмента * Кол-во операций) * N
        const pogon_time = (L * op.v_prokat + op.v_tool * op.n_tool) * N;
        total_prod_time += pogon_time;

        // Цена = Время * Тариф_ЧЛ + Время * Уд.Цена МШ
        price_work += (pogon_time * current_rate_ch) + (pogon_time * op.p_msh);
        price_work_mag += (pogon_time * current_rate_mag) + (pogon_time * op.p_msh);

      } else if (op.type === "переменный") {
        // --- 2. ПЕРЕМЕННАЯ ОПЕРАЦИЯ ---
        // Общее время = (Сумма ручных работ * N) + (Время машины * N)
        const perm_human_time = op.sum_manual * N;
        const perm_mach_time = op.t_mach * N;
        total_prod_time += (perm_human_time + perm_mach_time);

        // Цена на тираж = (Время Операции * Тариф_ЧЛ + Время Машины * Уд.Цена МШ) * N
        const base_cost_ch = (op.t_op * current_rate_ch) + (op.t_mach * op.p_msh);
        const base_cost_mag = (op.t_op * current_rate_mag) + (op.t_mach * op.p_msh);
        
        price_work += base_cost_ch * N;
        price_work_mag += base_cost_mag * N;

      } else {
        // --- 3. СТАТИЧНАЯ ОПЕРАЦИЯ ---
        // Общее время = Сумма ручных работ + Время машины (1 раз на партию)
        total_prod_time += (op.sum_manual + op.t_mach);

        // Цена статичной операции берется 1 раз на партию (не умножается на N)
        price_work += (op.t_op * current_rate_ch) + (op.t_mach * op.p_msh);
        price_work_mag += (op.t_op * current_rate_mag) + (op.t_mach * op.p_msh);
      }
    }

    // Итоговая цена подготовки
    const current_prep_rate = rate_ch > 0 ? rate_ch : (card.operations[0]?.p_ch || 0);
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
