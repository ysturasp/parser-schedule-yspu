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
    cache.put('directions_data', jsonData, 300);   
    return ContentService.createTextOutput(jsonData)
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
  
  directions.sort((a, b) => a.name.localeCompare(b.name, 'ru'));
  
  const jsonData = JSON.stringify(directions, null, 2);
  cache.put('directions_data', jsonData, 300); 
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
  CacheService.getScriptCache().put('directions_data', jsonData, 300);
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

  const directionHeaders = {
    first: rawData[2]["__EMPTY_1"] || "",
    second: rawData[2]["__EMPTY_2"] || "",
    third: rawData[2]["__EMPTY_3"] || "",
    fourth: rawData[2]["__EMPTY_4"] || ""
  };

  const courseInfo = {};
  Object.entries(groups).forEach(([course, groupName]) => {
    if (groupName) {
      let startDate = null;
      let groupNumber = null;
      let courseNumber = null;

      const fullMatch = groupName.match(/(\d+)\s*\((\d+)\s*курс\)\s*с\s*(\d{2}\.\d{2}\.\d{4})/);
      if (fullMatch) {
        groupNumber = fullMatch[1];
        courseNumber = parseInt(fullMatch[2]);
        startDate = fullMatch[3];
      } else {
        const simpleMatch = groupName.match(/(\d+)\s*\((\d+)\s*курс\)/);
        if (simpleMatch) {
          groupNumber = simpleMatch[1];
          courseNumber = parseInt(simpleMatch[2]);
        } else {
          const numberMatch = groupName.match(/^(\d{4,5})/);
          if (numberMatch) {
            groupNumber = numberMatch[1];
            courseNumber = 1;
          }
        }
      }

      if (!startDate && groupNumber && groupNumber.match(/^9[34]/)) {
        const headerDateMatch = directionHeaders[course].match(/с\s*(\d{2}\.\d{2}\.\d{4})/);
        if (headerDateMatch) {
          startDate = headerDateMatch[1];
        }
      }

      if (groupNumber) {
        courseInfo[course] = {
          number: groupNumber,
          course: courseNumber,
          startDate: startDate
        };
      }
    }
  });

  const courses = {};
  Object.entries(courseInfo).forEach(([course, info]) => {
    courses[course] = {
      name: info.number + " (" + info.course + " курс)",
      number: info.number,
      course: info.course,
      startDate: info.startDate
    };
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
 * Получает временной слот по его номеру
 */
function getTimeSlotByNumber(number) {
  const timeSlots = {
    1: { start: "08:30", end: "10:05" },
    2: { start: "10:15", end: "11:50" },
    3: { start: "12:15", end: "13:50" },
    4: { start: "14:00", end: "15:35" },
    5: { start: "15:45", end: "17:20" },
    6: { start: "17:30", end: "19:05" },
    7: { start: "19:15", end: "20:50" }
  };
  
  return timeSlots[number] || timeSlots[1];
}

/**
 * Парсит информацию о времени занятия
 */
function parseTime(timeStr, customTimeInfo = null) {
  if (!timeStr) {
    return {
      number: 1,
      startAt: "08:30",
      endAt: "10:05",
      timeRange: "08:30-10:05",
      originalTimeTitle: "1. 08.30-10.05",
      additionalSlots: []
    };
  }
    
  const timeSlots = {
    1: { start: "08:30", end: "10:05" },
    2: { start: "10:15", end: "11:50" },
    3: { start: "12:15", end: "13:50" },
    4: { start: "14:00", end: "15:35" },
    5: { start: "15:45", end: "17:20" },
    6: { start: "17:30", end: "19:05" },
    7: { start: "19:15", end: "20:50" }
  };

  const slots = timeStr.split(',').map(s => s.trim());
  
  if (slots.length > 1) {
    const parsedSlots = slots.map(slot => {
      const numberMatch = slot.match(/^(\d+)\./);
      if (numberMatch) {
        const number = parseInt(numberMatch[1]);
        const timeSlot = timeSlots[number];
        if (timeSlot) {
          return {
            number: number,
            startAt: timeSlot.start,
            endAt: timeSlot.end,
            timeRange: `${timeSlot.start}-${timeSlot.end}`,
            originalTimeTitle: `${number}. ${timeSlot.start.replace(':', '.')}-${timeSlot.end.replace(':', '.')}`
          };
        }
      }
      return null;
    }).filter(Boolean);

    if (parsedSlots.length > 0) {
      const mainSlot = parsedSlots[0];
      return {
        ...mainSlot,
        originalTimeTitle: timeStr,
        additionalSlots: parsedSlots.slice(1)
      };
    }
  }

  const numberOnlyMatch = timeStr.match(/^(\d+)$/);
  if (numberOnlyMatch) {
    const number = parseInt(numberOnlyMatch[1]);
    const slot = timeSlots[number] || timeSlots[1];
    return {
      number: number,
      startAt: slot.start,
      endAt: slot.end,
      timeRange: `${slot.start}-${slot.end}`,
      originalTimeTitle: timeStr,
      additionalSlots: []
    };
  }

  const parenNumberMatch = timeStr.match(/\((\d+)\s*пара\)/);
  if (parenNumberMatch) {
    const number = parseInt(parenNumberMatch[1]);
    const slot = timeSlots[number] || timeSlots[1];
    return {
      number: number,
      startAt: slot.start,
      endAt: slot.end,
      timeRange: `${slot.start}-${slot.end}`,
      originalTimeTitle: timeStr,
      additionalSlots: []
    };
  }
  
  const dotNumberMatch = timeStr.match(/^(\d+)\./);
  if (dotNumberMatch) {
    const number = parseInt(dotNumberMatch[1]);
    const slot = timeSlots[number] || timeSlots[1];
    
    const timeMatch = timeStr.match(/(\d{2})\.(\d{2})-(\d{2})\.(\d{2})/);
    if (timeMatch) {
      const [_, startHour, startMin, endHour, endMin] = timeMatch;
      return {
        number: number,
        startAt: `${startHour}:${startMin}`,
        endAt: `${endHour}:${endMin}`,
        timeRange: `${startHour}:${startMin}-${endHour}:${endMin}`,
        originalTimeTitle: timeStr,
        additionalSlots: []
      };
    }
      
    return {
      number: number,
      startAt: slot.start,
      endAt: slot.end,
      timeRange: `${slot.start}-${slot.end}`,
      originalTimeTitle: timeStr,
      additionalSlots: []
    };
  }

  return {
    number: 1,
    startAt: "08:30",
    endAt: "10:05",
    timeRange: "08:30-10:05",
    originalTimeTitle: timeStr,
    additionalSlots: []
  };
}

/**
 * Парсит информацию о предмете
 */
function parseSubject(subjectStr, defaultTime = "") {
  if (!subjectStr) return null;

  const languagePattern = /([А-ЯЁ][а-яё]+(?:\s+[а-яё]+)*\s+язык),\s*(?:доц\.|проф\.|ст\.преп\.|асс\.|преп\.)/g;
  const languageMatches = [...subjectStr.matchAll(languagePattern)];
  
  if (languageMatches.length > 1) {
    const subjects = [];
    let lastIndex = 0;
    let currentText = '';
    
    for (const match of languageMatches) {
      if (match.index > lastIndex) {
        if (currentText) {
          const parsed = parseSubject(currentText, defaultTime);
          if (parsed) {
            parsed.isPartOfComposite = true;
            subjects.push(parsed);
          }
        }
      }
      
      const nextMatch = languageMatches.find(m => m.index > match.index);
      const endIndex = nextMatch ? nextMatch.index : subjectStr.length;
      currentText = subjectStr.substring(match.index, endIndex);
      lastIndex = match.index;
    }
    
    if (currentText) {
      const parsed = parseSubject(currentText, defaultTime);
      if (parsed) {
        parsed.isPartOfComposite = true;
        subjects.push(parsed);
      }
    }
    
    if (subjects.length > 0) {
      return subjects;
    }
  }

  const subjects = subjectStr.split(/\n\s*\n/).filter(Boolean);
  if (subjects.length > 1) {
    return subjects.map(s => {
      const parsed = parseSubject(s, defaultTime);
      if (parsed) {
        parsed.isPartOfComposite = true;
      }
      return parsed;
    }).filter(Boolean);
  }
  
  const lowerSubject = subjectStr.toLowerCase();
  const parts = subjectStr.split(',').map(p => p.trim());
  let name = parts[0].trim();
  
  let customStartTime = null;
  let customEndTime = null;
  
  const startTimeMatch = subjectStr.match(/с\s*(\d{2}):(\d{2})/);
  if (startTimeMatch) {
    customStartTime = `${startTimeMatch[1]}:${startTimeMatch[2]}`;
  }
  
  const endTimeMatch = subjectStr.match(/до\s*(\d{2}):(\d{2})/);
  if (endTimeMatch) {
    customEndTime = `${endTimeMatch[1]}:${endTimeMatch[2]}`;
  }
  
  const isDistant = lowerSubject.includes('дистант') || 
                   lowerSubject.includes('онлайн') ||
                   lowerSubject.includes('он-лайн') ||
                   parts.some(part => part.trim().toLowerCase() === 'онлайн') ||
                   parts.some(part => part.trim().toLowerCase() === 'он-лайн');
  
  const isStream = lowerSubject.includes('поток');
  const isDivision = lowerSubject.includes('подгруппа');
  
  const isPhysicalEducation = name.toLowerCase().includes('физ') || 
                             name.toLowerCase().includes('фк') || 
                             name.toLowerCase().includes('физическ') ||
                             name.toLowerCase().includes('элективные дисциплины по фк');
  
  let type = 'other';
  if (isPhysicalEducation) {
    type = 'practice';
  } else {
    type = lowerSubject.includes('лек') ? 'lecture' : 
           lowerSubject.includes('практ') ? 'practice' : 'other';
  }

  let teachers = [];
  const teacherRegexPatterns = [
    /(?:доц\.|проф\.|ст\.преп\.|асс\.|преп\.)?\s*([А-ЯЁ][а-яё]+)\s*([А-ЯЁ])\s*\.\s*([А-ЯЁ])\s*\./,
    /(?:доц\.|проф\.|ст\.преп\.|асс\.|преп\.)?\s*([А-ЯЁ])\s*\.\s*([А-ЯЁ])\s*\.\s*([А-ЯЁ][а-яё]+)/,
    /(?:доц\.|проф\.|ст\.преп\.|асс\.|преп\.)?\s*([А-ЯЁ])\s*\.\s*([А-ЯЁ])\s*\.\s*([А-ЯЁ][а-яё]+)/,
    /(?:доц\.|проф\.|ст\.преп\.|асс\.|преп\.)?\s*([А-ЯЁ])\s+([А-ЯЁ])\s+([А-ЯЁ][а-яё]+)/
  ];

  let remainingText = subjectStr;

  for (const pattern of teacherRegexPatterns) {
    const matches = [...remainingText.matchAll(new RegExp(pattern, 'g'))];
    for (const match of matches) {
      let teacher;
      if (match[1].length > 1) {
        teacher = `${match[1]} ${match[2]}.${match[3]}.`;
      } else {
        teacher = `${match[3]} ${match[1]}.${match[2]}.`;
      }
      
      teacher = teacher.replace(/\s*\([^)]*\)/g, '')
                      .replace(/\s+с\s+\d{2}:\d{2}/g, '')
                      .replace(/\s+до\s+\d{2}\.\d{2}\.\d{4}/g, '')
                      .replace(/\s+кроме\s+\d{2}\.\d{2}\.\d{4}/g, '')
                      .trim();
      
      if (!teachers.includes(teacher)) {
        teachers.push(teacher);
      }
    }
  }

  let subjectWithoutDates = subjectStr.replace(/(?:с|по|до)\s*\d{2}\.\d{2}\.\d{4}/g, '');
  
  const buildingMatch = subjectWithoutDates.match(/(\d+)(?:-[её]|е|ое)?\s*(?:уч\.?|учебное)?\s*зд(?:ание)?\.?/i);
  const building = buildingMatch ? buildingMatch[1] : null;
  
  let room = null;
  const roomMatch = subjectWithoutDates.match(/,\s*(?:ауд\.)?\s*(\d+[МАБВГДЕЖЗИКЛМНОПРСТУФХЦЧШЩЭЮЯ]?)\s*(?:\(|$|,|\s+|ЕГФ|гл\.з\.)/);
  const sportHallMatch = subjectWithoutDates.match(/(?:спорт\.?\s*зал|спортзал|спортивный\s*зал)/i);
  
  if (roomMatch) {
    room = roomMatch[1];
  } else if (sportHallMatch) {
    room = "спортзал";
  }
  
  if (building && room) {
    room = `${building}-${room}`;
  }
  
  const startDateMatch = subjectStr.match(/с\s*(\d{2}\.\d{2}\.\d{4})/);
  const endDateMatch = subjectStr.match(/(?:по|до)\s*(\d{2}\.\d{2}\.\d{4})/);
  const startDate = startDateMatch ? startDateMatch[1] : null;
  const endDate = endDateMatch ? endDateMatch[1] : null;
  
  const lessonNumberMatch = subjectStr.match(/\((\d+)\s*пара\)/);
  const lessonNumber = lessonNumberMatch ? parseInt(lessonNumberMatch[1]) : null;
  
  let cleanName = name;
  
  cleanName = cleanName
    .replace(/\s*\(лекции\)/i, '')
    .replace(/\s*лек\./i, '')
    .replace(/\s*практ\./i, '')
    .replace(/\s*\(\d+\s*пара\)/, '')
    .replace(/\s*с\s*\d{2}:\d{2}/, '')
    .replace(/\s*до\s*\d{2}:\d{2}/, '')
    .replace(/\s*с\s*\d{2}\.\d{2}\.\d{4}/, '')
    .replace(/\s*по\s*\d{2}\.\d{2}\.\d{4}/, '')
    .trim();

  const timeInfo = {
    customStartTime,
    customEndTime
  };
  
  return {
    lessonName: cleanName,
    type: type,
    teacherName: teachers.join(', '),
    auditoryName: room,
    isDistant: isDistant,
    isStream: isStream,
    isDivision: isDivision,
    startDate: startDate,
    endDate: endDate,
    duration: 2,
    durationMinutes: 90,
    isShort: false,
    isLecture: type === 'lecture',
    originalText: subjectStr.trim(),
    lessonNumber: lessonNumber,
    defaultTime: defaultTime,
    timeInfo: timeInfo
  };
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

  const directionHeaders = {
    first: rawData[2]["__EMPTY_1"] || "",
    second: rawData[2]["__EMPTY_2"] || "",
    third: rawData[2]["__EMPTY_3"] || "",
    fourth: rawData[2]["__EMPTY_4"] || ""
  };

  const courseInfo = {};
  Object.entries(groups).forEach(([course, groupName]) => {
    if (groupName) {
      let startDate = null;
      let groupNumber = null;
      let courseNumber = null;

      const fullMatch = groupName.match(/(\d+)\s*\((\d+)\s*курс\)\s*с\s*(\d{2}\.\d{2}\.\d{4})/);
      if (fullMatch) {
        groupNumber = fullMatch[1];
        courseNumber = parseInt(fullMatch[2]);
        startDate = fullMatch[3];
      } else {
        const simpleMatch = groupName.match(/(\d+)\s*\((\d+)\s*курс\)/);
        if (simpleMatch) {
          groupNumber = simpleMatch[1];
          courseNumber = parseInt(simpleMatch[2]);
        } else {
          const numberMatch = groupName.match(/^(\d{4,5})/);
          if (numberMatch) {
            groupNumber = numberMatch[1];
            courseNumber = 1;
          }
        }
      }

      if (!startDate && groupNumber && groupNumber.match(/^9[34]/)) {
        const headerDateMatch = directionHeaders[course].match(/с\s*(\d{2}\.\d{2}\.\d{4})/);
        if (headerDateMatch) {
          startDate = headerDateMatch[1];
        }
      }

      if (groupNumber) {
        courseInfo[course] = {
          number: groupNumber,
          course: courseNumber,
          startDate: startDate
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
      if ((col === m.s.c && row > m.s.r && row <= m.e.r) ||
          (row === m.s.r && col >= m.s.c && col <= m.e.c)) {
        return true;
      }
    }
    return false;
  }

  function getMergedCellValue(row, col) {
    for (let m of merges) {
      if (row >= m.s.r && row <= m.e.r && col >= m.s.c && col <= m.e.c) {
        const startCol = m.s.c;
        const colLetter = XLSX.utils.encode_col(startCol);
        const cellRef = colLetter + (m.s.r + 1);
        return worksheet[cellRef] ? worksheet[cellRef].v : null;
      }
    }
    return null;
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
    const dayName = row["Министерство просвещения Российской Федерации \nФедеральное государственное бюджетное учреждение высшего образования\n«Ярославский государственный педагогический университет им. К.Д. Ушинского»\n"];
    const excelRowIdx = i + 1;
    
    if (dayName && ["Понедельник", "Вторник", "Среда", "Четверг", "Пятница", "Суббота"].includes(dayName)) {
      if (currentDay) {
        Object.entries(currentDaySchedule).forEach(([course, lessons]) => {
          if (lessons.length > 0 && courseInfo[course]) {
            const processedLessons = [];
            
            lessons.forEach(lesson => {
              const parsedSubjects = parseSubject(lesson.subject, lesson.time);
              if (Array.isArray(parsedSubjects)) {
                parsedSubjects.forEach(parsedSubject => {
                  const timeInfo = parsedSubject.lessonNumber ? 
                    parseTime(`${parsedSubject.lessonNumber}`, parsedSubject.timeInfo) : 
                    parseTime(parsedSubject.defaultTime, parsedSubject.timeInfo);
                  if (timeInfo) {
                    if (timeInfo.additionalSlots && timeInfo.additionalSlots.length > 0) {
                      processedLessons.push({
                        ...timeInfo,
                        ...parsedSubject
                      });
                      
                      timeInfo.additionalSlots.forEach(slot => {
                        processedLessons.push({
                          ...slot,
                          ...parsedSubject,
                          originalTimeTitle: slot.originalTimeTitle || slot.timeRange
                        });
                      });
                    } else {
                      processedLessons.push({
                        ...timeInfo,
                        ...parsedSubject
                      });
                    }
                  }
                });
              } else if (parsedSubjects) {
                const timeInfo = parseTime(lesson.time, parsedSubjects.timeInfo);
                if (timeInfo) {
                  if (timeInfo.additionalSlots && timeInfo.additionalSlots.length > 0) {
                    processedLessons.push({
                      ...timeInfo,
                      ...parsedSubjects
                    });
                    
                    timeInfo.additionalSlots.forEach(slot => {
                      processedLessons.push({
                        ...slot,
                        ...parsedSubjects,
                        originalTimeTitle: slot.originalTimeTitle || slot.timeRange
                      });
                    });
                  } else {
                    processedLessons.push({
                      ...timeInfo,
                      ...parsedSubjects
                    });
                  }
                }
              }
            });

            processedLessons.sort((a, b) => {
              if (a.number === b.number) {
                if (a.type === 'lecture' && b.type !== 'lecture') return -1;
                if (a.type !== 'lecture' && b.type === 'lecture') return 1;
                return 0;
              }
              return a.number - b.number;
            });

            schedule.items.push({
              number: course === 'first' ? 1 : course === 'second' ? 2 : course === 'third' ? 3 : 4,
              courseInfo: courseInfo[course],
              days: [{
                info: {
                  type: ["Понедельник", "Вторник", "Среда", "Четверг", "Пятница", "Суббота"].indexOf(currentDay),
                  weekNumber: 1,
                  date: new Date().toISOString()
                },
                lessons: processedLessons
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
        const courseCol = courseColIdx[course];
        let mergedValue = null;

        if (!subject) {
          mergedValue = getMergedCellValue(excelRowIdx, courseCol);
        }

        if (subject || mergedValue) {
          const actualSubject = subject || mergedValue;
          const last = currentDaySchedule[course][currentDaySchedule[course].length - 1];
          
          if (last && last.subject === actualSubject) {
            last.time += ", " + (time || "");
          } else {
            currentDaySchedule[course].push({ 
              time: time || "", 
              subject: actualSubject
            });
          }
        } else if (currentDaySchedule[course].length > 0 && isMergedCell(excelRowIdx, courseCol)) {
          currentDaySchedule[course][currentDaySchedule[course].length - 1].time += ", " + (time || "");
        }
      });
    }
  }

  if (currentDay) {
    Object.entries(currentDaySchedule).forEach(([course, lessons]) => {
      if (lessons.length > 0 && courseInfo[course]) {
        const processedLessons = [];
        
        lessons.forEach(lesson => {
          const parsedSubjects = parseSubject(lesson.subject, lesson.time);
          if (Array.isArray(parsedSubjects)) {
            parsedSubjects.forEach(parsedSubject => {
              const timeInfo = parsedSubject.lessonNumber ? 
                parseTime(`${parsedSubject.lessonNumber}`, parsedSubject.timeInfo) : 
                parseTime(parsedSubject.defaultTime, parsedSubject.timeInfo);
              if (timeInfo) {
                if (timeInfo.additionalSlots && timeInfo.additionalSlots.length > 0) {
                  processedLessons.push({
                    ...timeInfo,
                    ...parsedSubject
                  });
                  
                  timeInfo.additionalSlots.forEach(slot => {
                    processedLessons.push({
                      ...slot,
                      ...parsedSubject,
                      originalTimeTitle: slot.originalTimeTitle || slot.timeRange
                    });
                  });
                } else {
                  processedLessons.push({
                    ...timeInfo,
                    ...parsedSubject
                  });
                }
              }
            });
          } else if (parsedSubjects) {
            const timeInfo = parseTime(lesson.time, parsedSubjects.timeInfo);
            if (timeInfo) {
              if (timeInfo.additionalSlots && timeInfo.additionalSlots.length > 0) {
                processedLessons.push({
                  ...timeInfo,
                  ...parsedSubjects
                });
                
                timeInfo.additionalSlots.forEach(slot => {
                  processedLessons.push({
                    ...slot,
                    ...parsedSubjects,
                    originalTimeTitle: slot.originalTimeTitle || slot.timeRange
                  });
                });
              } else {
                processedLessons.push({
                  ...timeInfo,
                  ...parsedSubjects
                });
              }
            }
          }
        });

        processedLessons.sort((a, b) => {
          if (a.number === b.number) {
            if (a.type === 'lecture' && b.type !== 'lecture') return -1;
            if (a.type !== 'lecture' && b.type === 'lecture') return 1;
            return 0;
          }
          return a.number - b.number;
        });

        schedule.items.push({
          number: course === 'first' ? 1 : course === 'second' ? 2 : course === 'third' ? 3 : 4,
          courseInfo: courseInfo[course],
          days: [{
            info: {
              type: ["Понедельник", "Вторник", "Среда", "Четверг", "Пятница", "Суббота"].indexOf(currentDay),
              weekNumber: 1,
              date: new Date().toISOString()
            },
            lessons: processedLessons
          }]
        });
      }
    });
  }

  return ContentService.createTextOutput(JSON.stringify(schedule, null, 2))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Получает или создает лист с преподавателями в активной таблице
 */
function getTeachersSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Преподаватели');
  
  if (!sheet) {
    sheet = ss.insertSheet('Преподаватели');
    sheet.appendRow(['ID', 'ФИО', 'Последнее обновление', 'Расписание', 'История изменений']);
  }
  
  return sheet;
}

/**
 * Получает или создает лист с аудиториями в активной таблице
 */
function getAuditoriesSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Аудитории');
  
  if (!sheet) {
    sheet = ss.insertSheet('Аудитории');
    sheet.appendRow(['ID', 'Номер', 'Последнее обновление', 'Расписание', 'История изменений']);
  }
  
  return sheet;
}

/**
 * Сравнивает два расписания и возвращает список изменений
 */
function compareSchedules(oldSchedule, newSchedule) {
  const changes = new Set();
  
  const normalizeString = (str) => str.toLowerCase().replace(/\s+/g, ' ').trim();
  const createKey = (item) => {
    const day = item.day;
    const time = item.time.split('-')[0];
    const subject = normalizeString(item.subject).replace(/[^а-яёa-z0-9]/g, '');
    return `${day}_${time}_${subject}`;
  };
  
  const oldMap = new Map(oldSchedule.map(item => [createKey(item), item]));
  const newMap = new Map(newSchedule.map(item => [createKey(item), item]));
  
  for (const [key, newItem] of newMap.entries()) {
    if (!oldMap.has(key)) {
      changes.add(`Добавлено: ${newItem.day}, ${newItem.time}, ${newItem.subject}`);
    } else {
      const oldItem = oldMap.get(key);
      const fieldsToCompare = {
        group: { label: 'группа', normalize: false },
        time: { label: 'время', normalize: false },
        subject: { label: 'предмет', normalize: true },
        teacher: { label: 'преподаватель', normalize: false },
        type: { label: 'тип', normalize: false },
        auditoryName: { label: 'аудитория', normalize: false }
      };
      
      const changedFields = [];
      for (const [field, config] of Object.entries(fieldsToCompare)) {
        const oldValue = oldItem[field];
        const newValue = newItem[field];
        
        if (oldValue && newValue && oldValue !== newValue) {
          if (config.normalize) {
            const normalizedOld = normalizeString(oldValue);
            const normalizedNew = normalizeString(newValue);
            
            if (normalizedOld !== normalizedNew) {
              if (normalizedOld.replace(/[^а-яёa-z0-9]/g, '') === normalizedNew.replace(/[^а-яёa-z0-9]/g, '')) {
                changedFields.push(`исправлена опечатка в названии: ${oldValue} → ${newValue}`);
              } else {
                changedFields.push(`${config.label}: ${oldValue} → ${newValue}`);
              }
            }
          } else if (oldValue !== newValue) {
            changedFields.push(`${config.label}: ${oldValue} → ${newValue}`);
          }
        }
      }
      
      if (changedFields.length > 0) {
        changes.add(`Изменено (${changedFields.join(', ')})`);
      }
    }
  }
  
  for (const [key, oldItem] of oldMap.entries()) {
    if (!newMap.has(key)) {
      changes.add(`Удалено: ${oldItem.day}, ${oldItem.time}, ${oldItem.subject}`);
    }
  }
  
  return Array.from(changes);
}

/**
 * Обновляет данные о преподавателях и аудиториях
 */
function updateTeachersAndAuditories() {
  const folderId = "1Uz9POR8Ni66-fc3Au0YrfeYOTNXYJWID";
  const files = getFilesFromFolder(folderId);
  
  const teachersMap = new Map();
  const auditoriesMap = new Map();
  
  const teachersSheet = getTeachersSheet();
  const teachersData = teachersSheet.getDataRange().getValues();
  const existingTeachers = new Map();
  for (let i = 1; i < teachersData.length; i++) {
    existingTeachers.set(teachersData[i][0], {
      name: teachersData[i][1],
      schedule: JSON.parse(teachersData[i][3] || '[]'),
      history: JSON.parse(teachersData[i][4] || '[]')
    });
  }
  
  const auditoriesSheet = getAuditoriesSheet();
  const auditoriesData = auditoriesSheet.getDataRange().getValues();
  const existingAuditories = new Map();
  for (let i = 1; i < auditoriesData.length; i++) {
    existingAuditories.set(String(auditoriesData[i][0]), {
      number: auditorsData[i][1],
      schedule: JSON.parse(auditoriesData[i][3] || '[]'),
      history: JSON.parse(auditoriesData[i][4] || '[]')
    });
  }
  
  files.forEach(file => {
    try {
      const schedule = JSON.parse(getDirectionSchedule(file.id).getContent());
      const directionName = file.name.replace('.xlsx', '');
      
      schedule.items.forEach(item => {
        const courseInfo = item.courseInfo;
        
        item.days.forEach(day => {
          const dayType = day.info.type;
          
          day.lessons.forEach(lesson => {
            if (lesson.teacherName) {
              const teacherId = lesson.teacherName.replace(/\s+/g, '_').toLowerCase();
              if (!teachersMap.has(teacherId)) {
                teachersMap.set(teacherId, {
                  name: lesson.teacherName,
                  schedule: []
                });
              }
              
              const teacher = teachersMap.get(teacherId);
              teacher.schedule.push({
                day: dayType,
                time: lesson.timeRange,
                subject: lesson.lessonName,
                type: lesson.type,
                auditory: lesson.auditoryName,
                group: courseInfo.number,
                direction: directionName
              });
            }
            
            if (lesson.auditoryName) {
              const auditoryId = lesson.auditoryName.replace(/\s+/g, '_').toLowerCase();
              if (!auditoriesMap.has(auditoryId)) {
                auditoriesMap.set(auditoryId, {
                  number: lesson.auditoryName,
                  schedule: []
                });
              }
              
              const auditory = auditoriesMap.get(auditoryId);
              auditory.schedule.push({
                day: dayType,
                time: lesson.timeRange,
                subject: lesson.lessonName,
                type: lesson.type,
                teacher: lesson.teacherName,
                group: courseInfo.number,
                direction: directionName
              });
            }
          });
        });
      });
    } catch (e) {
      console.error(`Ошибка при обработке файла ${file.id}: ${e.message}`);
    }
  });
  
  const now = new Date().toISOString();
  const updatedTeacherRows = [];
  const updatedAuditoryRows = [];
  
  teachersMap.forEach((teacher, id) => {
    const streamKey = (item) => `${item.day}_${item.time}_${item.subject}_${item.auditory}_${item.type}`;
    const grouped = new Map();
    
    teacher.schedule.forEach(item => {
      const key = streamKey(item);
      if (!grouped.has(key)) {
        grouped.set(key, {
          ...item,
          directions: new Set([item.direction]),
          groups: new Set([item.group])
        });
      } else {
        const existing = grouped.get(key);
        existing.directions.add(item.direction);
        existing.groups.add(item.group);
      }
    });
    
    teacher.schedule = Array.from(grouped.values()).map(item => {
      const { directions, groups, ...rest } = item;
      return {
        ...rest,
        direction: Array.from(directions).sort().join(', '),
        group: Array.from(groups).sort().join(' ')
      };
    });
    
    const existing = existingTeachers.get(id);
    const scheduleJson = JSON.stringify(teacher.schedule);
    
    if (!existing || JSON.stringify(existing.schedule) !== scheduleJson) {
      const changes = existing ? compareSchedules(existing.schedule, teacher.schedule) : ['Начальное добавление расписания'];
      let history;
      try {
        history = existing?.history ? JSON.parse(JSON.stringify(existing.history)) : [];
      } catch (e) {
        history = [];
        console.error('Ошибка при парсинге истории преподавателя:', e);
      }
      
      if (changes.length > 0) {
        history.push({
          date: now,
          changes: changes
        });
        
        while (history.length > 10) {
          history.shift();
        }
        
        updatedTeacherRows.push({
          id: id,
          name: teacher.name,
          lastUpdate: now,
          schedule: scheduleJson,
          history: JSON.stringify(history)
        });
      }
    }
  });
  
  auditoriesMap.forEach((auditory, id) => {
    const streamKey = (item) => `${item.day}_${item.time}_${item.subject}_${item.teacher}_${item.type}`;
    const grouped = new Map();
    
    auditory.schedule.forEach(item => {
      const key = streamKey(item);
      if (!grouped.has(key)) {
        grouped.set(key, {
          ...item,
          directions: new Set([item.direction]),
          groups: new Set([item.group])
        });
      } else {
        const existing = grouped.get(key);
        existing.directions.add(item.direction);
        existing.groups.add(item.group);
      }
    });
    
    auditory.schedule = Array.from(grouped.values()).map(item => {
      const { directions, groups, ...rest } = item;
      return {
        ...rest,
        direction: Array.from(directions).sort().join(', '),
        group: Array.from(groups).sort().join(' ')
      };
    });
    
    const existing = existingAuditories.get(id);
    const scheduleJson = JSON.stringify(auditory.schedule);
    
    if (!existing || JSON.stringify(existing.schedule) !== scheduleJson) {
      const changes = existing ? compareSchedules(existing.schedule, auditory.schedule) : ['Начальное добавление расписания'];
      let history;
      try {
        history = existing?.history ? JSON.parse(JSON.stringify(existing.history)) : [];
      } catch (e) {
        history = [];
        console.error('Ошибка при парсинге истории аудитории:', e);
      }
      
      if (changes.length > 0) {
        history.push({
          date: now,
          changes: changes
        });
        
        while (history.length > 10) {
          history.shift();
        }
        
        updatedAuditoryRows.push({
          id: id,
          number: auditory.number,
          lastUpdate: now,
          schedule: scheduleJson,
          history: JSON.stringify(history)
        });
      }
    }
  });
  
  if (updatedTeacherRows.length > 0) {
    updatedTeacherRows.forEach(row => {
      const existingRow = teachersData.findIndex((r, i) => i > 0 && r[0] === row.id);
      if (existingRow > 0) {
        teachersSheet.getRange(existingRow + 1, 1, 1, 5).setValues([[
          row.id,
          row.name,
          row.lastUpdate,
          row.schedule,
          row.history
        ]]);
      } else {
        teachersSheet.appendRow([
          row.id,
          row.name,
          row.lastUpdate,
          row.schedule,
          row.history
        ]);
      }
    });
  }
  
  if (updatedAuditoryRows.length > 0) {
    updatedAuditoryRows.forEach(row => {
      const existingRow = auditorsData.findIndex((r, i) => i > 0 && String(r[0]) === String(row.id));
      if (existingRow > 0) {
        auditorsSheet.getRange(existingRow + 1, 1, 1, 5).setValues([[
          row.id,
          row.number,
          row.lastUpdate,
          row.schedule,
          row.history
        ]]);
      } else {
        auditorsSheet.appendRow([
          row.id,
          row.number,
          row.lastUpdate,
          row.schedule,
          row.history
        ]);
      }
    });
  }
  
  const cache = CacheService.getScriptCache();
  cache.put('teachers_data', JSON.stringify(Array.from(teachersMap.values())), 300);
  cache.put('auditories_data', JSON.stringify(Array.from(auditoriesMap.values())), 300);
}

/**
 * Получает расписание преподавателя
 */
function getTeacherSchedule(teacherId) {
  const sheet = getTeachersSheet();
  const data = sheet.getDataRange().getValues();
  const teacherRow = data.find((r, i) => i > 0 && r[0] === teacherId);
  
  if (!teacherRow) {
    throw new Error('Преподаватель не найден');
  }
  
  const schedule = {
    isCache: false,
    items: [{
      number: 0,
      courseInfo: {
        number: teacherRow[0],
        name: teacherRow[1],
        course: 0,
        startDate: null
      },
      days: []
    }]
  };

  const lessons = JSON.parse(teacherRow[3] || '[]');
  const daysMap = new Map();
  
  lessons.forEach(lesson => {
    const dayKey = lesson.day;
    if (!daysMap.has(dayKey)) {
      daysMap.set(dayKey, {
        info: {
          type: lesson.day,
          weekNumber: 1,
          date: new Date().toISOString()
        },
        lessons: []
      });
    }
    
    const timeMatch = lesson.time.match(/(\d{2}):(\d{2})-(\d{2}):(\d{2})/);
    if (timeMatch) {
      const [_, startHour, startMin, endHour, endMin] = timeMatch;
      const timeNumber = parseInt(lesson.time.split('.')[0]) || 1;
      
      const lessonData = {
        number: timeNumber,
        startAt: `${startHour}:${startMin}`,
        endAt: `${endHour}:${endMin}`,
        timeRange: lesson.time,
        originalTimeTitle: `${timeNumber}. ${startHour}.${startMin}-${endHour}.${endMin}`,
        additionalSlots: [],
        lessonName: lesson.subject,
        type: lesson.type || 'other',
        teacherName: teacherRow[1],
        auditoryName: lesson.auditory,
        isDistant: false,
        isStream: false,
        isDivision: false,
        startDate: null,
        endDate: null,
        duration: 2,
        durationMinutes: 90,
        isShort: false,
        isLecture: lesson.type === 'lecture',
        originalText: `${lesson.subject}, ${teacherRow[1]}, ${lesson.auditory || ''}`,
        groups: lesson.group,
        direction: lesson.direction
      };

      const lowerSubject = lesson.subject.toLowerCase();
      if (lowerSubject.includes('физ') || 
          lowerSubject.includes('фк') || 
          lowerSubject.includes('физическ') ||
          lowerSubject.includes('элективные дисциплины по фк')) {
        lessonData.type = 'practice';
      } else if (lowerSubject.includes('лек')) {
        lessonData.type = 'lecture';
      } else if (lowerSubject.includes('практ')) {
        lessonData.type = 'practice';
      }

      lessonData.isDistant = lowerSubject.includes('дистант') || 
                            lowerSubject.includes('онлайн') ||
                            lowerSubject.includes('он-лайн') ||
                            !lesson.auditory;
      lessonData.isStream = lowerSubject.includes('поток');
      lessonData.isDivision = lowerSubject.includes('подгруппа');
      
      daysMap.get(dayKey).lessons.push(lessonData);
    }
  });
  
  schedule.items[0].days = Array.from(daysMap.values());
  
  return ContentService.createTextOutput(JSON.stringify(schedule))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Получает расписание аудитории
 */
function getAuditorySchedule(auditoryId) {
  const sheet = getAuditoriesSheet();
  const data = sheet.getDataRange().getValues();
  const auditoryRow = data.find((r, i) => i > 0 && String(r[0]) === String(auditoryId));
  
  if (!auditoryRow) {
    throw new Error('Аудитория не найдена');
  }
  
  const schedule = {
    isCache: false,
    items: [{
      number: 0,
      courseInfo: {
        number: auditoryRow[0],
        course: 0,
        startDate: null
      },
      days: []
    }]
  };

  const lessons = JSON.parse(auditoryRow[3] || '[]');
  const daysMap = new Map();
  
  lessons.forEach(lesson => {
    const dayKey = lesson.day;
    if (!daysMap.has(dayKey)) {
      daysMap.set(dayKey, {
        info: {
          type: lesson.day,
          weekNumber: 1,
          date: new Date().toISOString()
        },
        lessons: []
      });
    }
    
    const timeMatch = lesson.time.match(/(\d{2}):(\d{2})-(\d{2}):(\d{2})/);
    if (timeMatch) {
      const [_, startHour, startMin, endHour, endMin] = timeMatch;
      const timeNumber = parseInt(lesson.time.split('.')[0]) || 1;
      
      const lessonData = {
        number: timeNumber,
        startAt: `${startHour}:${startMin}`,
        endAt: `${endHour}:${endMin}`,
        timeRange: lesson.time,
        originalTimeTitle: `${timeNumber}. ${startHour}.${startMin}-${endHour}.${endMin}`,
        additionalSlots: [],
        lessonName: lesson.subject,
        type: lesson.type || 'other',
        teacherName: lesson.teacher,
        auditoryName: auditoryRow[1],
        isDistant: false,
        isStream: false,
        isDivision: false,
        startDate: null,
        endDate: null,
        duration: 2,
        durationMinutes: 90,
        isShort: false,
        isLecture: lesson.type === 'lecture',
        originalText: `${lesson.subject}, ${lesson.teacher || ''}, ${auditoryRow[1]}`,
        groups: lesson.group,
        direction: lesson.direction
      };

      const lowerSubject = lesson.subject.toLowerCase();
      if (lowerSubject.includes('физ') || 
          lowerSubject.includes('фк') || 
          lowerSubject.includes('физическ') ||
          lowerSubject.includes('элективные дисциплины по фк')) {
        lessonData.type = 'practice';
      } else if (lowerSubject.includes('лек')) {
        lessonData.type = 'lecture';
      } else if (lowerSubject.includes('практ')) {
        lessonData.type = 'practice';
      }

      lessonData.isDistant = lowerSubject.includes('дистант') || 
                            lowerSubject.includes('онлайн') ||
                            lowerSubject.includes('он-лайн');
      lessonData.isStream = lowerSubject.includes('поток');
      lessonData.isDivision = lowerSubject.includes('подгруппа');
      
      daysMap.get(dayKey).lessons.push(lessonData);
    }
  });
  
  schedule.items[0].days = Array.from(daysMap.values());
  
  return ContentService.createTextOutput(JSON.stringify(schedule))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Получает список всех преподавателей
 */
function getTeachers() {
  const cache = CacheService.getScriptCache();
  const cachedData = cache.get('teachers_list');
  
  if (cachedData) {
    return ContentService.createTextOutput(cachedData)
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  const sheet = getTeachersSheet();
  const data = sheet.getDataRange().getValues();
  
  const teachers = [];
  for (let i = 1; i < data.length; i++) {
    teachers.push({
      id: data[i][0],
      name: data[i][1]
    });
  }
  
  teachers.sort((a, b) => a.name.localeCompare(b.name, 'ru'));
  
  const jsonData = JSON.stringify(teachers);
  cache.put('teachers_list', jsonData, 300);
  
  return ContentService.createTextOutput(jsonData)
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Получает список всех аудиторий
 */
function getAuditories() {
  const cache = CacheService.getScriptCache();
  const cachedData = cache.get('auditories_list');
  
  if (cachedData) {
    return ContentService.createTextOutput(cachedData)
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  const sheet = getAuditoriesSheet();
  const data = sheet.getDataRange().getValues();
  
  const auditories = [];
  for (let i = 1; i < data.length; i++) {
    auditories.push({
      id: data[i][0],
      number: String(data[i][1])
    });
  }
  
  auditories.sort((a, b) => {
    const numMatchA = String(a.number).match(/\d+/);
    const numMatchB = String(b.number).match(/\d+/);
    
    if (!numMatchA || !numMatchB) {
      return String(a.number).localeCompare(String(b.number), 'ru');
    }
    
    const numA = parseInt(numMatchA[0]);
    const numB = parseInt(numMatchB[0]);
    
    if (numA === numB) {
      return String(a.number).localeCompare(String(b.number), 'ru');
    }
    return numA - numB;
  });
  
  const jsonData = JSON.stringify(auditories);
  cache.put('auditories_list', jsonData, 300);
  
  return ContentService.createTextOutput(jsonData)
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Получает список всех преподавателей с расписанием
 */
function getTeachersWithSchedule() {
  const cache = CacheService.getScriptCache();
  const cachedData = cache.get('teachers_data');
  
  if (cachedData) {
    return ContentService.createTextOutput(cachedData)
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  const sheet = getTeachersSheet();
  const data = sheet.getDataRange().getValues();
  
  const teachers = [];
  for (let i = 1; i < data.length; i++) {
    teachers.push({
      id: data[i][0],
      name: data[i][1],
      schedule: JSON.parse(data[i][3] || '[]')
    });
  }
  
  teachers.sort((a, b) => a.name.localeCompare(b.name, 'ru'));
  
  const jsonData = JSON.stringify(teachers);
  cache.put('teachers_data', jsonData, 300);
  
  return ContentService.createTextOutput(jsonData)
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Получает список всех аудиторий с расписанием
 */
function getAuditoriesWithSchedule() {
  const cache = CacheService.getScriptCache();
  const cachedData = cache.get('auditories_data');
  
  if (cachedData) {
    return ContentService.createTextOutput(cachedData)
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  const sheet = getAuditoriesSheet();
  const data = sheet.getDataRange().getValues();
  
  const auditories = [];
  for (let i = 1; i < data.length; i++) {
    auditories.push({
      id: data[i][0],
      number: String(data[i][1]),
      schedule: JSON.parse(data[i][3] || '[]')
    });
  }
  
  auditories.sort((a, b) => {
    const numMatchA = String(a.number).match(/\d+/);
    const numMatchB = String(b.number).match(/\d+/);
    
    if (!numMatchA || !numMatchB) {
      return String(a.number).localeCompare(String(b.number), 'ru');
    }
    
    const numA = parseInt(numMatchA[0]);
    const numB = parseInt(numMatchB[0]);
    
    if (numA === numB) {
      return String(a.number).localeCompare(String(b.number), 'ru');
    }
    return numA - numB;
  });
  
  const jsonData = JSON.stringify(auditories);
  cache.put('auditories_data', jsonData, 300);
  
  return ContentService.createTextOutput(jsonData)
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Очищает все кэши
 */
function clearAllCaches() {
  const cache = CacheService.getScriptCache();
  cache.removeAll([
    'directions_data',
    'teachers_data',
    'teachers_list',
    'auditories_data',
    'auditories_list'
  ]);
  
  return ContentService.createTextOutput(JSON.stringify({ 
    status: "success",
    message: "Кэши очищены"
  })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * HTTP-обработчик для веб-API
 */
function doGet(e) {
  try {
    if (!e || !e.parameter) {
      return getDirections();
    }
    
    const { action, id, full } = e.parameter;
    
    switch (action) {
      case 'directions':
        return getDirections();
      case 'schedule':
        if (!id) {
          throw new Error('Не указан ID направления');
        }
        return getDirectionSchedule(id);
      case 'teachers':
        return full === 'true' ? getTeachersWithSchedule() : getTeachers();
      case 'teacher':
        if (!id) {
          throw new Error('Не указан ID преподавателя');
        }
        return getTeacherSchedule(id);
      case 'auditories':
        return full === 'true' ? getAuditoriesWithSchedule() : getAuditories();
      case 'auditory':
        if (!id) {
          throw new Error('Не указан ID аудитории');
        }
        return getAuditorySchedule(id);
      case 'force-update':
        updateDirectionsData();
        updateTeachersAndAuditories();
        return ContentService.createTextOutput(JSON.stringify({ 
          status: "success",
          message: "Данные обновлены"
        })).setMimeType(ContentService.MimeType.JSON);
      case 'clear-cache':
        return clearAllCaches();
      default:
        throw new Error('Неизвестное действие');
    }
  } catch (e) {
    return ContentService.createTextOutput(JSON.stringify({ 
      error: "Ошибка: " + e.message 
    })).setMimeType(ContentService.MimeType.JSON);
  }
} 