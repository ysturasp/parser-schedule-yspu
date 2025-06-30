<div align="center">
  <img src="https://ysturasp.netlify.app/images/cat.png" alt="ystuRASP Logo" width="120" height="120" style="border-radius: 20%">
  
  # 🧳 Парсер расписания ЯГПУ
  
  [![Telegram](https://img.shields.io/badge/Telegram-2CA5E0?style=for-the-badge&logo=telegram&logoColor=white)](https://t.me/ysturasp)
  [![GitHub](https://img.shields.io/badge/GitHub-100000?style=for-the-badge&logo=github&logoColor=white)](https://github.com/ysturasp)
</div>

<div align="center">
  <h3>🚀 Парсер расписания ФСУ ЯГПУ им. К.Д. Ушинского</h3>
</div>

> 📁 Оригинальные таблицы с расписанием доступны по [ссылке](https://drive.google.com/drive/folders/1Uz9POR8Ni66-fc3Au0YrfeYOTNXYJWID)

## 🎯 Основные возможности

<div align="center">
  <table>
    <tr>
      <td align="center">
        <img src="https://ysturasp.github.io/tg_emoji/Objects/Card%20Index%20Dividers.webp" width="100" alt="Parser"><br>
        <b>Парсинг Excel-файлов</b>
      </td>
      <td align="center">
        <img src="https://ysturasp.github.io/tg_emoji/Objects/Memo.webp" width="100" alt="Schedule"><br>
        <b>Обработка расписания</b>
      </td>
      <td align="center">
        <img src="https://ysturasp.github.io/tg_emoji/Objects/Ballot Box With Ballot.webp" width="100" alt="Cache"><br>
        <b>Кэширование данных</b>
      </td>
      <td align="center">
        <img src="https://ysturasp.github.io/tg_emoji/REST API.gif" width="100" alt="API"><br>
        <b>REST API</b>
      </td>
    </tr>
  </table>
</div>

## 📋 Описание

Парсер предназначен для обработки Excel-файлов с расписанием занятий ЯГПУ и предоставления данных через REST API. Он является частью проекта [ysturasp](https://github.com/ysturasp/ysturaspp) и обеспечивает бэкенд для работы с расписанием.

## 🛠 Функциональность

- 📊 Парсинг Excel-файлов с расписанием
- 👨‍🏫 Извлечение информации о преподавателях
- 🏛 Обработка данных об аудиториях
- 📅 Анализ расписания по дням недели
- 💾 Кэширование данных для быстрого доступа
- 🔄 Автоматическое обновление информации
- 🌐 REST API для получения данных

## 🔧 Технологии

<div align="center">
  
  ![Google Apps Script](https://img.shields.io/badge/Google_Apps_Script-4285F4?style=for-the-badge&logo=google&logoColor=white)
  ![JavaScript](https://img.shields.io/badge/JavaScript-F7DF1E?style=for-the-badge&logo=javascript&logoColor=black)
  ![Google Sheets](https://img.shields.io/badge/Google_Sheets-34A853?style=for-the-badge&logo=google-sheets&logoColor=white)
  ![REST API](https://img.shields.io/badge/REST_API-FF6C37?style=for-the-badge&logo=postman&logoColor=white)
  
</div>

## 📥 API Endpoints

- `?action=directions` - Получение списка направлений
- `?action=schedule&id={ID}` - Расписание конкретного направления
- `?action=teachers` - Список преподавателей
- `?action=teacher&id={ID}` - Расписание преподавателя
- `?action=auditories` - Список аудиторий
- `?action=auditory&id={ID}` - Расписание аудитории
- `?action=force-update` - Принудительное обновление данных
- `?action=clear-cache` - Очистка кэша

## 💻 Использование

1. Укажите актуальный ID папки Google Drive с файлами расписания
2. Настройте доступ к файлам через Google Apps Script
3. Разверните скрипт как веб-приложение для ВСЕХ
4. Используйте API endpoints для получения данных

ps предварительно таблицы создавать не нужно, они сами генерируются если у их нет по необходимости их использования парсером

## 🌟 Поддержите проект

<div align="center">
  <a href="https://boosty.to/ysturasp.me/donate">
    <img src="https://img.shields.io/badge/Поддержать_проект-F15B2A?style=for-the-badge&logo=boosty&logoColor=white" alt="Support Project">
  </a>
</div>

---

<div align="center">
  <h3>✨ Спасибо за интерес к нашему проекту! ✨</h3>
</div>

<div align="center">
  <p>Если вам понравился проект, не забудьте поставить ⭐ звезду на GitHub!</p>
  
  <a href="https://github.com/ysturasp/parser-schedule-yspu">
    <img src="https://img.shields.io/github/stars/ysturasp/parser-schedule-yspu?style=social" alt="GitHub Stars">
  </a>
</div>

<div align="center">
  <sub>Made with ❤️ by ystuRASP © 2024</sub>
</div> 