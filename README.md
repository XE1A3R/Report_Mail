# Report_Mail
[![wakatime](https://wakatime.com/badge/github/lameRER/Report_Mail.svg)](https://wakatime.com/badge/github/lameRER/Report_Mail)<br>
![GitHub commit activity](https://img.shields.io/github/commit-activity/m/lamerer/Report_Mail)<br>
<br>
Master:<br>
![.NET Core Desktop](https://img.shields.io/github/workflow/status/lamerer/Report_Mail/.NET%20Core%20Desktop)<br>
![GitHub branch checks state](https://img.shields.io/github/checks-status/lameRER/Report_Mail/master)<br>
![GitHub issues](https://img.shields.io/github/issues/lamerer/Report_Mail)<br>
![GitHub pull requests](https://img.shields.io/github/issues-pr/lamerer/Report_Mail)

Report-Mail:

        Описание: 
          Автоматическая рассылка ежедневных отчетов с во вложенным Excel файлом(и). 
          Данное ПО построено на конфигурационных файлах (JSON), что позволяет привлечь разработчика только для расширения функционала. 
          ПО запускается через Task Scheduler с аргументом конфигурационного файла. 
          Файл конфигурации позволяет создавать отчеты с неограниченным количеством вложений. 
          В файл конфигурации входит: 
          Имя файла, название листа, SQL запрос, расположение таблицы, возможность создание умной таблицы, закрепление таблицы, окрашивание ячеек по условию, получатели, копия получателей и т.д.

        Цели: Полностью автоматизировать процесс рассылки ежедневных отчетов

        Stack: WinForms, ODBC, JSON.
