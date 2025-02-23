# ExcelToPDF
Автоматизация экспорта Excel листов в PDF ver 1.1

Что делает скрипт:
1. Анализ папок с инвойсами
2. Копирует xlsx-файлы инвойсов (это необходимо для работы с файлом без его порчи)
3. Экспорт первого и второго листов инвойса.xlsx/xls по указанному диапазону, указанному через "," или "-" в файле "path.txt"
4. После экспорта всех Excel-файлов в PDF происходит общее скрепление файлов в папке "PDF"
5. Автоматическое удаление экспортированных PDF-файлов 1-го и 2-го листов, а также откопированных xlsx/xls файлов

Перед началом запуска скрипта запускаем файл "Start.bat" для того, чтобы были установлены все необходимые библиотеки. Это необходимо сделать только в 1-ый раз.

Для работы скрипта в файле path.txt указываем по порядку 3 строки

Путь директории, где находятся папки с инвойсами xlsx;
Путь куда копируются Excel-файлы и экспортируются в PDF, а также где будет создана папка "PDF" с скрепленными документами;
Диапазон номеров инвойсов (Например, 436,439 (скрепит 2 шт.), 436-439 (скрепит 5 шт.) или 436-439,451-453,455 (скрепит 7 шт.)).

Важно в самом инвойсе.xlsx в ячеке R1 первого и второго листов указать видимый диапазон для печати.
