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
 * Получает список всех доступных направлений
 */
function getDirections() {
  const folderId = "1Uz9POR8Ni66-fc3Au0YrfeYOTNXYJWID";
  const files = getFilesFromFolder(folderId);
  
  const directions = files.map(file => ({
    id: file.id,
    name: file.name.replace('.xlsx', '')
  }));
  
  return ContentService.createTextOutput(JSON.stringify(directions, null, 2))
    .setMimeType(ContentService.MimeType.JSON);
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
      default:
        throw new Error('Неизвестное действие');
    }
  } catch (e) {
    return ContentService.createTextOutput(JSON.stringify({ 
      error: "Ошибка: " + e.message 
    })).setMimeType(ContentService.MimeType.JSON);
  }
} 