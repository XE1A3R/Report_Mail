# Report_Mail

[![wakatime](https://wakatime.com/badge/github/lameRER/Report_Mail.svg)](https://wakatime.com/badge/github/lameRER/Report_Mail)<br>
![GitHub commit activity](https://img.shields.io/github/commit-activity/m/lamerer/Report_Mail)<br>

**Master:**

![.NET Core Desktop](https://img.shields.io/github/workflow/status/lamerer/Report_Mail/.NET%20Core%20Desktop)<br>
![GitHub branch checks state](https://img.shields.io/github/checks-status/lameRER/Report_Mail/master)<br>
![GitHub issues](https://img.shields.io/github/issues/lamerer/Report_Mail)<br>
![GitHub pull requests](https://img.shields.io/github/issues-pr/lamerer/Report_Mail)

## Общие сведения

Report-Mail — это программное обеспечение для автоматической отправки ежедневных отчетов с вложенными файлами Microsoft Excel. Приложение использует конфигурационные файлы формата JSON для настройки параметров отчетов, что обеспечивает гибкость и возможность расширения функциональности без глубокого изменения кода.

## Архитектура и технологии

*   **GUI:** Windows Forms (WinForms)
*   **База данных:** ODBC (через System.Data.Odbc)
*   **Формат конфигурации:** JSON
*   **Работа с Excel:** EPPlus, Spire.XLS, Microsoft.Office.Interop.Excel
*   **Почта:** SMTP (реализация в коде)
*   **Платформа:** .NET Framework 4.8

## Функциональные возможности

*   **Автоматическая рассылка:** Отправка отчетов по расписанию через Task Scheduler.
*   **Поддержка вложений:** Создание и отправка неограниченного количества вложенных файлов Excel.
*   **Настройка через JSON:** Конфигурация параметров отчетов (название файла, лист, SQL-запрос, расположение данных, форматирование) через JSON-файлы.
*   **Гибкие настройки Excel:** Возможность создания умных таблиц, закрепления строк/столбцов, условного форматирования ячеек.
*   **Управление получателями:** Настройка основных получателей и копий (CC) для рассылки.

## Цель проекта

Полностью автоматизировать процесс создания и отправки ежедневных отчетов, минимизировав необходимость ручного вмешательства.

## Установка и настройка

### Требования

*   Операционная система: Microsoft Windows
*   .NET Framework 4.8
*   Доступ к базе данных через ODBC
*   Microsoft Excel (для работы с Excel-файлами, возможно, требуется установленный Office или библиотеки)

### Установка из исходного кода

1.  Клонируйте репозиторий:
    ```bash
    git clone https://github.com/lameRER/Report_Mail.git
    ```
2.  Откройте решение `Report_Mail.sln` в Visual Studio.
3.  Восстановите NuGet-пакеты. В Visual Studio: `Средства` -> `Диспетчер пакетов NuGet` -> `Восстановить`.
4.  Скомпилируйте проект в конфигурации Release.
5.  Скопируйте исполняемый файл `Report_Mail.exe` и все сопутствующие библиотеки (например, `Prof.dll`, `EPPlus.dll`, `System.Data.Odbc.dll`, `System.Text.Json.dll` и др.) в целевую директорию.

### Запуск через Task Scheduler

Программа предназначена для запуска по расписанию. Для настройки Task Scheduler:

1.  Создайте новую задачу.
2.  В поле "Программа и аргументы" укажите путь к `Report_Mail.exe` и путь к файлу конфигурации в качестве аргумента:
    ```
    C:\path\to\Report_Mail.exe C:\path\to\your\config.json
    ```
3.  Настройте расписание задачи (например, ежедневно в определенное время).

### Конфигурация

Конфигурация программы осуществляется через JSON-файлы. Пример структуры конфигурационного файла (`config.json`):

```json
{
  "FileName": "Отчет_{{date}}.xlsx",
  "Sheets": [
    {
      "Name": "Лист1",
      "SQLQuery": "SELECT col1, col2 FROM table WHERE date = '{{date}}'",
      "TableLocation": "A1",
      "EnableSmartTable": true,
      "FreezePane": true,
      "ConditionalFormatting": [
        {
          "Range": "A:A",
          "ColorScale": {
            "MinColor": "#FF0000",
            "MaxColor": "#00FF00"
          }
        }
      ]
    }
  ],
  "Recipients": [
    "recipient1@example.com"
  ],
  "CC": [
    "cc_recipient1@example.com"
  ],
  "Subject": "Ежедневный отчет {{date}}",
  "Body": "Текст сообщения.",
  "ConnectionString": "Dsn=your_dsn_name"
}

```

Где:

*   `FileName` - имя выходного файла. `{{date}}` заменяется текущей датой.
*   `Sheets` - массив конфигураций для листов Excel.
*   `Sheets[].Name` - имя листа.
*   `Sheets[].SQLQuery` - SQL-запрос для получения данных.
*   `Sheets[].TableLocation` - ячейка, с которой начинается вставка данных.
*   `Sheets[].EnableSmartTable` - создавать ли умную таблицу.
*   `Sheets[].FreezePane` - закреплять ли области.
*   `Sheets[].ConditionalFormatting` - правила условного форматирования.
*   `Recipients` - список основных получателей.
*   `CC` - список получателей копии.
*   `Subject` - тема письма.
*   `Body` - текст письма.
*   `ConnectionString` - строка подключения к базе данных (альтернатива настройке в App.config).

Конфигурация может быть также частично задана в файле `App.config` (например, строка подключения `connectionStrings`, настройки почты `userSettings`).

## Структура проекта

*   `Report_Mail.sln` - файл решения Visual Studio.
*   `Report_Mail/` - основная директория проекта.
    *   `Controller/` - классы, отвечающие за логику (база данных, Excel, почта, файлы и т.д.).
    *   `Interface/` - интерфейсы для классов.
    *   `Model/` - классы, представляющие данные (конфигурация, почта, Excel и т.д.).
    *   `Form1.cs` - основная форма приложения (WinForms).
    *   `Program.cs` - точка входа в программу.
    *   `App.config` - файл конфигурации приложения .NET.
    *   `packages.config` - список управляемых NuGet-зависимостей.
    *   `Report_Mail.csproj` - файл проекта MSBuild.


## Лицензия

Этот проект распространяется под лицензией MIT. Подробности см. в файле [LICENSE](./LICENSE).
