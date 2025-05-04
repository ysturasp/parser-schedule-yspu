/**
 * Получает список всех файлов из указанной папки Google Drive
 */
function getFilesFromFolder(folderId) {
  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFiles();
  const fileList = [];
  
  while (files.hasNext()) {
    const file = files.next();
    if (file.getMimeType() === MimeType.MICROSOFT_EXCEL) {
      fileList.push({
        id: file.getId(),
        name: file.getName()
      });
    }
  }
  
  return fileList;
}

/**
 * Получает или создает лист с направлениями в активной таблице
 */
function getDirectionsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Направления');
  
  if (!sheet) {
    sheet = ss.insertSheet('Направления');
    sheet.appendRow(['ID', 'Название', 'Последнее обновление', 'Курсы']);
  }
  
  return sheet;
}

/**
 * Сохраняет направления в таблицу
 */
function saveDirectionsToSheet(directions) {
  const sheet = getDirectionsSheet();
  const data = sheet.getDataRange().getValues();
  
  if (data.length === 0) {
    sheet.appendRow(['ID', 'Название', 'Последнее обновление', 'Курсы']);
  }
  
  const existingDirections = {};
  for (let i = 1; i < data.length; i++) {
    existingDirections[data[i][0]] = {
      name: data[i][1],
      lastUpdate: data[i][2],
      courses: data[i][3]
    };
  }
  
  directions.forEach(dir => {
    const now = new Date().toISOString();
    const coursesJson = JSON.stringify(dir.courses);
    
    if (!existingDirections[dir.id] || 
        existingDirections[dir.id].name !== dir.name || 
        existingDirections[dir.id].courses !== coursesJson) {
      
      const row = [dir.id, dir.name, now, coursesJson];
      const existingRow = data.findIndex((r, i) => i > 0 && r[0] === dir.id);
      
      if (existingRow > 0) {
        sheet.getRange(existingRow + 1, 1, 1, 4).setValues([row]);
      } else {
        sheet.appendRow(row);
      }
    }
  });
}

/**
 * Получает список всех доступных направлений
 */
function getDirections() {
  const cache = CacheService.getScriptCache();
  const cachedData = cache.get('directions_data');
  
  if (cachedData) {
    return ContentService.createTextOutput(cachedData)
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  const sheet = getDirectionsSheet();
  const data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) {
    const initialData = createInitialDirectionsData();
    saveDirectionsToSheet(initialData);
    const jsonData = JSON.stringify(initialData, null, 2);
    cache.put('directions_data', jsonData, 21600);   return ContentService.createTextOutput(jsonData)
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  const directions = [];
  for (let i = 1; i < data.length; i++) {
    directions.push({
      id: data[i][0],
      name: data[i][1],
      courses: JSON.parse(data[i][3] || '{}')
    });
  }
  
  const jsonData = JSON.stringify(directions, null, 2);
  cache.put('directions_data', jsonData, 21600); 
  scheduleUpdate();
  
  return ContentService.createTextOutput(jsonData)
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Создает триггер для обновления данных
 */
function scheduleUpdate() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'updateDirectionsData') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  ScriptApp.newTrigger('updateDirectionsData')
    .timeBased()
    .after(1000)   .create();
}

/**
 * Обновляет данные о направлениях
 */
function updateDirectionsData() {
  const folderId = "1Uz9POR8Ni66-fc3Au0YrfeYOTNXYJWID";
  const files = getFilesFromFolder(folderId);
  
  const directions = files.map(file => {
    try {
      const courses = getCoursesFromFile(file.id);
      return {
        id: file.id,
        name: file.name.replace('.xlsx', ''),
        courses: courses
      };
    } catch (e) {
      console.error(`Ошибка при получении данных для файла ${file.id}: ${e.message}`);
      return {
        id: file.id,
        name: file.name.replace('.xlsx', ''),
        courses: {}
      };
    }
  });
  
  saveDirectionsToSheet(directions);
  
  const jsonData = JSON.stringify(directions, null, 2);
  CacheService.getScriptCache().put('directions_data', jsonData, 21600);
}

/**
 * Создает начальные данные о направлениях
 */
function createInitialDirectionsData() {
  const folderId = "1Uz9POR8Ni66-fc3Au0YrfeYOTNXYJWID";
  const files = getFilesFromFolder(folderId);
  
  return files.map(file => ({
    id: file.id,
    name: file.name.replace('.xlsx', ''),
    courses: {}
  }));
}

/**
 * Получает информацию о курсах из файла
 */
function getCoursesFromFile(fileId) {
  const sheetName = "Table 1";
  
  const sheetJsUrl = "https://cdn.sheetjs.com/xlsx-latest/package/dist/xlsx.full.min.js";
  const sheetJsCode = UrlFetchApp.fetch(sheetJsUrl).getContentText();
  eval(sheetJsCode);

  const fileObj = DriveApp.getFileById(fileId);
  const blob = fileObj.getBlob();
  const bytes = blob.getBytes();
  const uint8Array = new Uint8Array(bytes);
  
  const workbook = XLSX.read(uint8Array, { type: "array" });
  const worksheet = workbook.Sheets[sheetName];
  
  if (!worksheet) {
    throw new Error(`Лист "${sheetName}" не найден в файле.`);
  }
  
  const rawData = XLSX.utils.sheet_to_json(worksheet, { defval: null });
  
  const groups = {
    first: rawData[3]["__EMPTY_1"],
    second: rawData[3]["__EMPTY_2"],
    third: rawData[3]["__EMPTY_3"],
    fourth: rawData[3]["__EMPTY_4"]
  };

  const courses = {};
  Object.entries(groups).forEach(([course, groupName]) => {
    if (groupName) {
      const match = groupName.match(/(\d+)\s*\((\d+)\s*курс\)\s*с\s*(\d{2}\.\d{2}\.\d{4})/);
      if (match) {
        courses[course] = {
          name: groupName,
          number: match[1],
          course: match[2],
          startDate: match[3]
        };
      }
    }
  });

  return courses;
}

/**
 * Получает или создает лист с курсами в активной таблице
 */
function getCoursesSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Курсы');
  
  if (!sheet) {
    sheet = ss.insertSheet('Курсы');
    sheet.appendRow(['ID', 'Название', 'Курс', 'Дата начала']);
  }
  
  return sheet;
}

/**
 * Сохраняет информацию о курсах в таблицу
 */
function saveCoursesToSheet(directionId, directionName, courses) {
  const sheet = getCoursesSheet();
  const data = sheet.getDataRange().getValues();
  
  if (data.length === 0) {
    sheet.appendRow(['ID', 'Название', 'Курс', 'Дата начала']);
  }
  
  const existingCourses = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === directionId) {
      existingCourses[data[i][2]] = {
        name: data[i][1],
        startDate: data[i][3]
      };
    }
  }
  
  const courseData = {
    first: { name: '9429 (1 курс)', startDate: '27.01.2025' },
    second: { name: '9329 (2 курс)', startDate: '03.02.2025' },
    third: { name: '9229 (3 курс)', startDate: '03.02.2025' },
    fourth: { name: '9129 (4 курс)', startDate: '24.02.2025' }
  };
  
  Object.entries(courseData).forEach(([course, info]) => {
    if (!existingCourses[course] || existingCourses[course].name !== info.name) {
      const row = [directionId, directionName, course, info.startDate];
      const existingRow = data.findIndex((r, i) => i > 0 && r[0] === directionId && r[2] === course);
      
      if (existingRow > 0) {
        sheet.getRange(existingRow + 1, 1, 1, 4).setValues([row]);
      } else {
        sheet.appendRow(row);
      }
    }
  });
}

/**
 * Получает расписание для конкретного направления
 */
function getDirectionSchedule(fileId) {
  const sheetName = "Table 1";
  
  const sheetJsUrl = "https://cdn.sheetjs.com/xlsx-latest/package/dist/xlsx.full.min.js";
  const sheetJsCode = UrlFetchApp.fetch(sheetJsUrl).getContentText();
  eval(sheetJsCode);

  const fileObj = DriveApp.getFileById(fileId);
  const blob = fileObj.getBlob();
  const bytes = blob.getBytes();
  const uint8Array = new Uint8Array(bytes);
  
  const workbook = XLSX.read(uint8Array, { type: "array" });
  const worksheet = workbook.Sheets[sheetName];
  
  if (!worksheet) {
    throw new Error(`Лист "${sheetName}" не найден в файле.`);
  }
  
  const rawData = XLSX.utils.sheet_to_json(worksheet, { defval: null });
  
  const schedule = {
    header: {
      university: rawData[0]["Министерство просвещения Российской Федерации \nФедеральное государственное бюджетное учреждение высшего образования\n«Ярославский государственный педагогический университет им. К.Д. Ушинского»\n"],
      title: rawData[1]["Министерство просвещения Российской Федерации \nФедеральное государственное бюджетное учреждение высшего образования\n«Ярославский государственный педагогический университет им. К.Д. Ушинского»\n"],
      program: rawData[2]["__EMPTY_1"],
      groups: {
        first: rawData[3]["__EMPTY_1"],
        second: rawData[3]["__EMPTY_2"],
        third: rawData[3]["__EMPTY_3"],
        fourth: rawData[3]["__EMPTY_4"]
      }
    },
    courses: {
      first: {},
      second: {},
      third: {},
      fourth: {}
    }
  };

  saveCoursesToSheet(fileId, schedule.header.program, schedule.courses);

  let currentDay = null;
  let currentDaySchedule = {
    first: [],
    second: [],
    third: [],
    fourth: []
  };

  const merges = worksheet['!merges'] || [];

  const range = XLSX.utils.decode_range(worksheet['!ref']);
  const firstExcelRowIdx = range.s.r;

  const courseColIdx = {
    first: 2,
    second: 3,
    third: 4,
    fourth: 5
  };

  function isMergedCell(row, col) {
    for (let m of merges) {
      if (col === m.s.c && row > m.s.r && row <= m.e.r) {
        return true;
      }
    }
    return false;
  }

  function pushOrMerge(scheduleArr, time, subject, excelRowIdx, colIdx) {
    if (subject) {
      const last = scheduleArr[scheduleArr.length - 1];
      if (last && last.subject === subject) {
        last.time += ", " + time;
      } else {
        scheduleArr.push({ time, subject });
      }
    } else if (scheduleArr.length > 0 && isMergedCell(excelRowIdx, colIdx)) {
      scheduleArr[scheduleArr.length - 1].time += ", " + time;
    }
  }

  for (let i = 4; i < rawData.length; i++) {
    const row = rawData[i];
    const dayName = row["Министерство просвещения Российской Федерации \nФедеральное государственное бюджетное учреждение высшего образования\n«Ярославский государственный педагогический университет им. К.Д. Ушинского»\n"];
    const excelRowIdx = i + 1;
    
    if (dayName && ["Понедельник", "Вторник", "Среда", "Четверг", "Пятница", "Суббота"].includes(dayName)) {
      if (currentDay) {
        schedule.courses.first[currentDay] = currentDaySchedule.first;
        schedule.courses.second[currentDay] = currentDaySchedule.second;
        schedule.courses.third[currentDay] = currentDaySchedule.third;
        schedule.courses.fourth[currentDay] = currentDaySchedule.fourth;
      }
      currentDay = dayName;
      currentDaySchedule = {
        first: [],
        second: [],
        third: [],
        fourth: []
      };
    }

    if (currentDay && row["__EMPTY"]) {
      const time = row["__EMPTY"];
      pushOrMerge(currentDaySchedule.first, time, row["__EMPTY_1"], excelRowIdx, courseColIdx.first);
      pushOrMerge(currentDaySchedule.second, time, row["__EMPTY_2"], excelRowIdx, courseColIdx.second);
      pushOrMerge(currentDaySchedule.third, time, row["__EMPTY_3"], excelRowIdx, courseColIdx.third);
      pushOrMerge(currentDaySchedule.fourth, time, row["__EMPTY_4"], excelRowIdx, courseColIdx.fourth);
    }
  }

  if (currentDay) {
    schedule.courses.first[currentDay] = currentDaySchedule.first;
    schedule.courses.second[currentDay] = currentDaySchedule.second;
    schedule.courses.third[currentDay] = currentDaySchedule.third;
    schedule.courses.fourth[currentDay] = currentDaySchedule.fourth;
  }

  return ContentService.createTextOutput(JSON.stringify(schedule, null, 2))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * HTTP-обработчик для веб-API
 */
function doGet(e) {
  try {
    if (!e || !e.parameter) {
      return getDirections();
    }
    
    const { action, id } = e.parameter;
    
    switch (action) {
      case 'directions':
        return getDirections();
      case 'schedule':
        if (!id) {
          throw new Error('Не указан ID направления');
        }
        return getDirectionSchedule(id);
      case 'force-update':
              updateDirectionsData();
        return ContentService.createTextOutput(JSON.stringify({ 
          status: "success",
          message: "Данные обновлены"
        })).setMimeType(ContentService.MimeType.JSON);
      default:
        throw new Error('Неизвестное действие');
    }
  } catch (e) {
    return ContentService.createTextOutput(JSON.stringify({ 
      error: "Ошибка: " + e.message 
    })).setMimeType(ContentService.MimeType.JSON);
  }
} 