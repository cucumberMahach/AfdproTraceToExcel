# AfdproTraceToExcel
Утилита конвертации файла трассировки из AFDPRO в Excel
![AfdproTraceToExcel result](/assets/result_pic.png "AfdproTraceToExcel result")
## Использование
1. Включить запись трассировки в AFDPRO командой `TR ON CLR`.
2. Выполнить нужное кол-во шагов.
3. Выполнить команду `PT {start, lenght, {filename}}`, где `start` - с какой команды начать запись (я использую 0), `lenght` - кол-во команд для записи, `filename` - имя файла.
4. Полученный файл передать как аргумент в утилиту.
## Установка
Программа мультиплатформенная, так как написана на Java.
