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
    cache.put('directions_data', jsonData, 300);   return ContentService.createTextOutput(jsonData)
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

  function parseTime(timeStr) {
    if (!timeStr) {
      const match = timeStr && timeStr.match(/(\d+)\./);
      const number = match ? parseInt(match[1]) : 1;
      
      const timeSlots = {
        1: { start: "08:30", end: "10:05" },
        2: { start: "10:15", end: "11:50" },
        3: { start: "12:15", end: "13:50" },
        4: { start: "14:00", end: "15:35" },
        5: { start: "15:45", end: "17:20" },
        6: { start: "17:30", end: "19:05" }
      };
      
      const slot = timeSlots[number] || timeSlots[1];
      
      return {
        number: number,
        startAt: slot.start,
        endAt: slot.end,
        timeRange: `${slot.start}-${slot.end}`,
        originalTimeTitle: `${number}. ${slot.start.replace(':', '.')}-${slot.end.replace(':', '.')}`,
        additionalSlots: []
      };
    }
    
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

    if (parsedSlots.length === 0) {
      const numberMatch = timeStr.match(/(\d+)\./);
      if (numberMatch) {
        const number = parseInt(numberMatch[1]);
        const timeSlots = {
          1: { start: "08:30", end: "10:05" },
          2: { start: "10:15", end: "11:50" },
          3: { start: "12:15", end: "13:50" },
          4: { start: "14:00", end: "15:35" },
          5: { start: "15:45", end: "17:20" },
          6: { start: "17:30", end: "19:05" }
        };
        const slot = timeSlots[number] || timeSlots[1];
        
        return {
          number: number,
          startAt: slot.start,
          endAt: slot.end,
          timeRange: `${slot.start}-${slot.end}`,
          originalTimeTitle: `${number}. ${slot.start.replace(':', '.')}-${slot.end.replace(':', '.')}`,
          additionalSlots: []
        };
      }
      return null;
    }
    
    return {
      ...parsedSlots[0],
      originalTimeTitle: timeStr,
      additionalSlots: parsedSlots.slice(1)
    };
  }

  function parseSubject(subjectStr) {
    if (!subjectStr) return null;
    
    const lowerSubject = subjectStr.toLowerCase();
    const parts = subjectStr.split(',');
    const name = parts[0].trim();
    
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
      type = subjectStr.toLowerCase().includes('лек') ? 'lecture' : 
             subjectStr.toLowerCase().includes('практ') ? 'practice' : 'other';
    }

    let teacher = null;
    const teacherRegexPatterns = [
      /(?:доц\.|проф\.|ст\.преп\.|асс\.|преп\.)?\s*([А-ЯЁ][а-яё]+)\s*([А-ЯЁ])\s*\.\s*([А-ЯЁ])\s*\./,
      /(?:доц\.|проф\.|ст\.преп\.|асс\.|преп\.)?\s*([А-ЯЁ])\s*\.\s*([А-ЯЁ])\s*\.\s*([А-ЯЁ][а-яё]+)/,
      /(?:доц\.|проф\.|ст\.преп\.|асс\.|преп\.)?\s*([А-ЯЁ])\s*\.\s*([А-ЯЁ])\s*\.\s*([А-ЯЁ][а-яё]+)/,
      /(?:доц\.|проф\.|ст\.преп\.|асс\.|преп\.)?\s*([А-ЯЁ])\s+([А-ЯЁ])\s+([А-ЯЁ][а-яё]+)/
    ];

    for (const pattern of teacherRegexPatterns) {
      const match = subjectStr.match(pattern);
      if (match) {
        if (match[1].length > 1) {
          teacher = `${match[1]} ${match[2]}.${match[3]}.`;
        } else {
          teacher = `${match[3]} ${match[1]}.${match[2]}.`;
        }
        break;
      }
    }

    if (teacher) {
      teacher = teacher.replace(/\s*\([^)]*\)/g, '')
                      .replace(/\s+с\s+\d{2}:\d{2}/g, '')
                      .replace(/\s+до\s+\d{2}\.\d{2}\.\d{4}/g, '')
                      .replace(/\s+кроме\s+\d{2}\.\d{2}\.\d{4}/g, '')
                      .trim();
    }
    
    let subjectWithoutDates = subjectStr.replace(/(?:с|по)\s*\d{2}\.\d{2}\.\d{4}/g, '');
    
    const roomMatch = subjectWithoutDates.match(/,\s*(?:ауд\.)?\s*(\d+[МАБВГДЕЖЗИКЛМНОПРСТУФХЦЧШЩЭЮЯ]?)\s*(?:\(|$|,|\s+|ЕГФ)/);
    const room = roomMatch ? roomMatch[1] : null;
    
    const startDateMatch = subjectStr.match(/с\s*(\d{2}\.\d{2}\.\d{4})/);
    const endDateMatch = subjectStr.match(/по\s*(\d{2}\.\d{2}\.\d{4})/);
    const startDate = startDateMatch ? startDateMatch[1] : null;
    const endDate = endDateMatch ? endDateMatch[1] : null;
    
    let cleanName = name;
    if (startDate) {
        cleanName = cleanName.replace(/с\s*\d{2}\.\d{2}\.\d{4}/, '').trim();
    }
    if (endDate) {
        cleanName = cleanName.replace(/по\s*\d{2}\.\d{2}\.\d{4}/, '').trim();
    }
    
    return {
      lessonName: cleanName,
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
      isLecture: type === 'lecture',
      originalText: subjectStr.trim()
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
    const dayName = row["Министерство просвещения Российской Федерации \nФедеральное государственное бюджетное учреждение высшего образования\n«Ярославский государственный педагогический университет им. К.Д. Ушинского»\n"];
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
        const courseCol = courseColIdx[course];
        let mergedValue = null;

        if (!subject) {
          mergedValue = getMergedCellValue(excelRowIdx, courseCol);
        }

        if (subject || mergedValue) {
          const actualSubject = subject || mergedValue;
          const last = currentDaySchedule[course][currentDaySchedule[course].length - 1];
          const isFirstLesson = currentDaySchedule[course].length === 0;
          
          if (last && last.subject === actualSubject) {
            last.time += ", " + (time || "");
          } else {
            currentDaySchedule[course].push({ 
              time: time || "", 
              subject: actualSubject,
              isFirstLesson
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
 * Получает или создает лист с преподавателями в активной таблице
 */
function getTeachersSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Преподаватели');
  
  if (!sheet) {
    sheet = ss.insertSheet('Преподаватели');
    sheet.appendRow(['ID', 'ФИО', 'Последнее обновление', 'Расписание']);
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
    sheet.appendRow(['ID', 'Номер', 'Последнее обновление', 'Расписание']);
  }
  
  return sheet;
}

/**
 * Обновляет данные о преподавателях и аудиториях
 */
function updateTeachersAndAuditories() {
  const folderId = "1Uz9POR8Ni66-fc3Au0YrfeYOTNXYJWID";
  const files = getFilesFromFolder(folderId);
  
  const teachersMap = new Map();
  const auditoriesMap = new Map();
  
  files.forEach(file => {
    try {
      const schedule = JSON.parse(getDirectionSchedule(file.id).getContent());
      const directionName = file.name.replace('.xlsx', '');
      
      schedule.items.forEach(item => {
        item.days.forEach(day => {
          day.lessons.forEach(lesson => {
            if (lesson.teacherName) {
              const teacherId = lesson.teacherName.toLowerCase().replace(/\s+/g, '_');
              if (!teachersMap.has(teacherId)) {
                teachersMap.set(teacherId, {
                  id: teacherId,
                  name: lesson.teacherName,
                  schedule: []
                });
              }
              
              teachersMap.get(teacherId).schedule.push({
                direction: directionName,
                group: item.courseInfo.number,
                day: day.info.type,
                time: lesson.timeRange,
                subject: lesson.lessonName,
                auditory: lesson.auditoryName,
                type: lesson.type,
                isDistant: lesson.isDistant || lesson.type === 'distant'
              });
            }
            
            if (lesson.auditoryName) {
              const auditoryId = String(lesson.auditoryName);
              if (!auditoriesMap.has(auditoryId)) {
                auditoriesMap.set(auditoryId, {
                  id: auditoryId,
                  number: lesson.auditoryName,
                  schedule: []
                });
              }
              
              auditoriesMap.get(auditoryId).schedule.push({
                direction: directionName,
                group: item.courseInfo.number,
                day: day.info.type,
                time: lesson.timeRange,
                subject: lesson.lessonName,
                teacher: lesson.teacherName,
                type: lesson.type
              });
            }
          });
        });
      });
    } catch (e) {
      console.error(`Ошибка при обработке файла ${file.id}: ${e.message}`);
    }
  });
  
  teachersMap.forEach(teacher => {
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
  });
  
  auditoriesMap.forEach(auditory => {
    const streamKey = (item) => `${item.day}_${item.time}_${item.subject}_${item.teacher}_${item.type}`;
    const grouped = new Map();
    
    auditory.schedule.forEach(item => {
      const key = streamKey(item);
      if (!grouped.has(key)) {
        grouped.set(key, {
          day: item.day,
          time: item.time,
          subject: item.subject,
          teacher: item.teacher,
          type: item.type,
          directions: new Set([item.direction]),
          groups: new Set([item.group])
        });
      } else {
        const existing = grouped.get(key);
        if (!existing.directions.has(item.direction)) {
          existing.directions.add(item.direction);
        }
        if (!existing.groups.has(item.group)) {
          existing.groups.add(item.group);
        }
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
  });
  
  const teachersSheet = getTeachersSheet();
  const teachersData = teachersSheet.getDataRange().getValues();
  const now = new Date().toISOString();
  
  if (teachersData.length > 1) {
    teachersSheet.getRange(2, 1, teachersData.length - 1, teachersData[0].length).clear();
  }
  
  teachersMap.forEach(teacher => {
    const scheduleJson = JSON.stringify(teacher.schedule);
    teachersSheet.appendRow([teacher.id, teacher.name, now, scheduleJson]);
  });
  
  const auditoriesSheet = getAuditoriesSheet();
  const auditoriesData = auditoriesSheet.getDataRange().getValues();
  
  if (auditoriesData.length > 1) {
    auditoriesSheet.getRange(2, 1, auditoriesData.length - 1, auditoriesData[0].length).clear();
  }
  
  auditoriesMap.forEach(auditory => {
    const scheduleJson = JSON.stringify(auditory.schedule);
    auditoriesSheet.appendRow([auditory.id, auditory.number, now, scheduleJson]);
  });
  
  const cache = CacheService.getScriptCache();
  cache.put('teachers_data', JSON.stringify(Array.from(teachersMap.values())), 300);
  cache.put('auditories_data', JSON.stringify(Array.from(auditoriesMap.values())), 300);
}

/**
 * Получает расписание преподавателя
 */
function getTeacherSchedule(teacherId) {
  const cache = CacheService.getScriptCache();
  const cachedData = cache.get('teachers_data');
  
  // if (cachedData) {
  //   const teachers = JSON.parse(cachedData);
  //   const teacher = teachers.find(t => t.id === teacherId);
  //   if (teacher) {
  //     return ContentService.createTextOutput(JSON.stringify(teacher))
  //       .setMimeType(ContentService.MimeType.JSON);
  //   }
  // }
  
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
      const timeNumber = parseInt(lesson.time.split('.')[0]);
      
      daysMap.get(dayKey).lessons.push({
        number: timeNumber,
        startAt: `${startHour}:${startMin}`,
        endAt: `${endHour}:${endMin}`,
        timeRange: lesson.time,
        originalTimeTitle: `${timeNumber}. ${startHour}.${startMin}-${endHour}.${endMin}`,
        additionalSlots: [],
        lessonName: lesson.subject,
        type: lesson.type,
        teacherName: teacherRow[1],
        auditoryName: lesson.auditory,
        isDistant: lesson.isDistant || lesson.type === 'distant',
        isStream: lesson.type === 'stream',
        isDivision: lesson.type === 'division',
        startDate: null,
        endDate: null,
        duration: 2,
        durationMinutes: 90,
        isShort: false,
        isLecture: lesson.type === 'lecture',
        originalText: `${lesson.subject}, ${lesson.type}, ${teacherRow[1]}, ${lesson.auditory || 'онлайн'}`,
        groups: lesson.group,
        direction: lesson.direction
      });
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
  const cache = CacheService.getScriptCache();
  const cachedData = cache.get('auditories_data');
  
  // if (cachedData) {
  //   const auditories = JSON.parse(cachedData);
  //   const auditory = auditors.find(a => a.id === auditoryId);
  //   if (auditory) {
  //     return ContentService.createTextOutput(JSON.stringify(auditory))
  //       .setMimeType(ContentService.MimeType.JSON);
  //   }
  // }
  
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
      const timeNumber = parseInt(lesson.time.split('.')[0]);
      
      daysMap.get(dayKey).lessons.push({
        number: timeNumber,
        startAt: `${startHour}:${startMin}`,
        endAt: `${endHour}:${endMin}`,
        timeRange: lesson.time,
        originalTimeTitle: `${timeNumber}. ${startHour}.${startMin}-${endHour}.${endMin}`,
        additionalSlots: [],
        lessonName: lesson.subject,
        type: lesson.type,
        teacherName: lesson.teacher,
        auditoryName: auditoryRow[1],
        isDistant: lesson.isDistant || lesson.type === 'distant',
        isStream: lesson.type === 'stream',
        isDivision: lesson.type === 'division',
        startDate: null,
        endDate: null,
        duration: 2,
        durationMinutes: 90,
        isShort: false,
        isLecture: lesson.type === 'lecture',
        originalText: `${lesson.subject}, ${lesson.type}, ${lesson.teacher}, ${auditoryRow[1]}`,
        groups: lesson.group,
        direction: lesson.direction
      });
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
  
  // if (cachedData) {
  //   return ContentService.createTextOutput(cachedData)
  //     .setMimeType(ContentService.MimeType.JSON);
  // }
  
  const sheet = getTeachersSheet();
  const data = sheet.getDataRange().getValues();
  
  const teachers = [];
  for (let i = 1; i < data.length; i++) {
    teachers.push({
      id: data[i][0],
      name: data[i][1]
    });
  }
  
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
  
  // if (cachedData) {
  //   return ContentService.createTextOutput(cachedData)
  //     .setMimeType(ContentService.MimeType.JSON);
  // }
  
  const sheet = getAuditoriesSheet();
  const data = sheet.getDataRange().getValues();
  
  const auditories = [];
  for (let i = 1; i < data.length; i++) {
    auditories.push({
      id: data[i][0],
      number: data[i][1]
    });
  }
  
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
  
  // if (cachedData) {
  //   return ContentService.createTextOutput(cachedData)
  //     .setMimeType(ContentService.MimeType.JSON);
  // }
  
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
  
  // if (cachedData) {
  //   return ContentService.createTextOutput(cachedData)
  //     .setMimeType(ContentService.MimeType.JSON);
  // }
  
  const sheet = getAuditoriesSheet();
  const data = sheet.getDataRange().getValues();
  
  const auditories = [];
  for (let i = 1; i < data.length; i++) {
    auditories.push({
      id: data[i][0],
      number: data[i][1],
      schedule: JSON.parse(data[i][3] || '[]')
    });
  }
  
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