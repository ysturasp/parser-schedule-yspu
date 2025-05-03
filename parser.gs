/**
 * Парсит XLSX с Google Диска и возвращает данные в JSON.
 * Не требует конвертации файла!
 */
function getExcelDataAsJson() {
  const fileId = "15PhdKlJ0eMO7bILaAJKVcuA3_uJgjGAy";
  const sheetName = "Table 1";
  
  const sheetJsUrl = "https://cdn.sheetjs.com/xlsx-latest/package/dist/xlsx.full.min.js";
  const sheetJsCode = UrlFetchApp.fetch(sheetJsUrl).getContentText();
  eval(sheetJsCode);

  const file = DriveApp.getFileById(fileId);
  const blob = file.getBlob();
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
function doGet() {
  try {
    return getExcelDataAsJson();
  } catch (e) {
    return ContentService.createTextOutput(JSON.stringify({ 
      error: "Ошибка: " + e.message 
    })).setMimeType(ContentService.MimeType.JSON);
  }
} 