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
    isCache: false,
    items: []
  };

  const groups = {
    first: rawData[3]["__EMPTY_1"],
    second: rawData[3]["__EMPTY_2"],
    third: rawData[3]["__EMPTY_3"],
    fourth: rawData[3]["__EMPTY_4"]
  };

  // Парсинг информации о курсах
  const courseInfo = {};
  Object.entries(groups).forEach(([course, groupName]) => {
    if (groupName) {
      const match = groupName.match(/(\d+)\s*\((\d+)\s*курс\)\s*с\s*(\d{2}\.\d{2}\.\d{4})/);
      if (match) {
        courseInfo[course] = {
          number: match[1],
          course: parseInt(match[2]),
          startDate: match[3]
        };
      }
    }
  });

  const courseColIdx = {
    first: 2,
    second: 3,
    third: 4,
    fourth: 5
  };

  const merges = worksheet['!merges'] || [];

  function isMergedCell(row, col) {
    for (let m of merges) {
      if (col === m.s.c && row > m.s.r && row <= m.e.r) {
        return true;
      }
    }
    return false;
  }

  function parseTime(timeStr) {
    if (!timeStr) return null;
    
    const timeSlots = timeStr.split(',').map(slot => slot.trim());
    const parsedSlots = timeSlots.map(slot => {
      const match = slot.match(/(\d+)\.\s*(\d{2})\.(\d{2})-(\d{2})\.(\d{2})/);
      if (!match) return null;
      
      const [_, number, startHour, startMin, endHour, endMin] = match;
      return {
        number: parseInt(number),
        startAt: `${startHour}:${startMin}`,
        endAt: `${endHour}:${endMin}`,
        timeRange: `${startHour}:${startMin}-${endHour}:${endMin}`
      };
    }).filter(slot => slot !== null);

    if (parsedSlots.length === 0) return null;
    
    return {
      ...parsedSlots[0],
      originalTimeTitle: timeStr,
      additionalSlots: parsedSlots.slice(1)
    };
  }

  function parseSubject(subjectStr) {
    if (!subjectStr) return null;
    
    const parts = subjectStr.split(',');
    const name = parts[0].trim();
    
    const type = subjectStr.toLowerCase().includes('лек') ? 'lecture' : 
                subjectStr.toLowerCase().includes('практ') ? 'practice' : 'other';
    
    const teacherMatch = subjectStr.match(/(?:доц\.|проф\.|ст\.преп\.|асс\.)\s*([^,]+)/);
    const teacher = teacherMatch ? teacherMatch[1].trim() : null;
    
    const roomMatch = subjectStr.match(/(\d+[МАБВГДЕЖЗИКЛМНОПРСТУФХЦЧШЩЭЮЯ]?)/);
    const room = roomMatch ? roomMatch[1] : null;
    
    const dateMatch = subjectStr.match(/с\s*(\d{2}\.\d{2}\.\d{4})(?:\s*по\s*(\d{2}\.\d{2}\.\d{4}))?/);
    const startDate = dateMatch ? dateMatch[1] : null;
    const endDate = dateMatch ? dateMatch[2] : null;
    
    const isDistant = subjectStr.toLowerCase().includes('дистант');
    const isStream = subjectStr.toLowerCase().includes('поток');
    const isDivision = subjectStr.toLowerCase().includes('подгруппа');
    
    return {
      lessonName: name,
      type: type,
      teacherName: teacher,
      auditoryName: room,
      isDistant: isDistant,
      isStream: isStream,
      isDivision: isDivision,
      startDate: startDate,
      endDate: endDate,
      duration: 2,
      durationMinutes: 90,
      isShort: false,
      isLecture: type === 'lecture'
    };
  }

  let currentDay = null;
  let currentDaySchedule = {
    first: [],
    second: [],
    third: [],
    fourth: []
  };

  for (let i = 4; i < rawData.length; i++) {
    const row = rawData[i];
    const dayName = row["Министерство просвещеия Российской Федерации \nФедеральное государственное бюджетное учреждение высшего образования\n«Ярославский государственный педагогический университет им. К.Д. Ушинского»\n"];
    const excelRowIdx = i + 1;
    
    if (dayName && ["Понедельник", "Вторник", "Среда", "Четверг", "Пятница", "Суббота"].includes(dayName)) {
      if (currentDay) {
        Object.entries(currentDaySchedule).forEach(([course, lessons]) => {
          if (lessons.length > 0) {
            const courseNumber = course === 'first' ? 1 : course === 'second' ? 2 : course === 'third' ? 3 : 4;
            schedule.items.push({
              number: courseNumber,
              courseInfo: courseInfo[course] || {
                number: courseNumber,
                course: courseNumber,
                startDate: null
              },
              days: [{
                info: {
                  type: ["Понедельник", "Вторник", "Среда", "Четверг", "Пятница", "Суббота"].indexOf(currentDay),
                  weekNumber: 1,
                  date: new Date().toISOString()
                },
                lessons: lessons.map(lesson => ({
                  ...parseTime(lesson.time),
                  ...parseSubject(lesson.subject)
                })).filter(lesson => lesson.lessonName)
              }]
            });
          }
        });
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
      const subjects = {
        first: row["__EMPTY_1"],
        second: row["__EMPTY_2"],
        third: row["__EMPTY_3"],
        fourth: row["__EMPTY_4"]
      };

      Object.entries(subjects).forEach(([course, subject]) => {
        if (subject) {
          const last = currentDaySchedule[course][currentDaySchedule[course].length - 1];
          if (last && last.subject === subject) {
            last.time += ", " + time;
          } else {
            currentDaySchedule[course].push({ time, subject });
          }
        } else if (currentDaySchedule[course].length > 0 && isMergedCell(excelRowIdx, courseColIdx[course])) {
          currentDaySchedule[course][currentDaySchedule[course].length - 1].time += ", " + time;
        }
      });
    }
  }

  if (currentDay) {
    Object.entries(currentDaySchedule).forEach(([course, lessons]) => {
      if (lessons.length > 0) {
        const courseNumber = course === 'first' ? 1 : course === 'second' ? 2 : course === 'third' ? 3 : 4;
        schedule.items.push({
          number: courseNumber,
          courseInfo: courseInfo[course] || {
            number: courseNumber,
            course: courseNumber,
            startDate: null
          },
          days: [{
            info: {
              type: ["Понедельник", "Вторник", "Среда", "Четверг", "Пятница", "Суббота"].indexOf(currentDay),
              weekNumber: 1,
              date: new Date().toISOString()
            },
            lessons: lessons.map(lesson => ({
              ...parseTime(lesson.time),
              ...parseSubject(lesson.subject)
            })).filter(lesson => lesson.lessonName)
          }]
        });
      }
    });
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