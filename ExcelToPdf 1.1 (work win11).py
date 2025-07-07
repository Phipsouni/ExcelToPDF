import os
import shutil
import win32com.client as win32
import PyPDF2
import pythoncom  # Добавляем для обработки COM-исключений


def read_paths_and_range(file_path):
    try:
        with open(file_path, 'r', encoding='utf-8') as f:  # Указываем кодировку для надежности
            lines = [line.strip() for line in f if line.strip()]  # Убираем пустые строки
            if len(lines) < 3:
                raise ValueError(
                    "Файл path.txt должен содержать три строки: путь к исходной папке, путь к целевой папке и диапазон.")
            source_folder = lines[0]
            destination_folder = lines[1]
            range_str = lines[2]
            return source_folder, destination_folder, range_str
    except FileNotFoundError:
        print(f"Ошибка: Файл '{file_path}' не найден. Убедитесь, что он существует.")
        return None, None, None
    except Exception as e:
        print(f"Ошибка при чтении path.txt: {e}")
        return None, None, None


def save_two_sheets_as_pdf(file_path, pdf_path):
    excel = None  # Инициализируем excel вне блока try, чтобы он был доступен в finally
    wb = None  # Инициализируем wb вне блока try

    try:
        # Попытка получить уже запущенный экземпляр Excel или создать новый
        try:
            excel = win32.GetActiveObject("Excel.Application")
            print("Используется активный экземпляр Excel.")
        except pythoncom.com_error:
            excel = win32.DispatchEx('Excel.Application')
            print("Запущен новый экземпляр Excel.")

        excel.Visible = False
        excel.DisplayAlerts = False  # Отключаем всплывающие окна Excel, которые могут блокировать скрипт

        # Открываем книгу. ReadOnly=True может помочь с проблемами блокировки, но
        # может быть несовместимо, если Excel вносит изменения в файл, хотя в вашем случае он только читает PrintArea.
        # Если возникнут проблемы, попробуйте убрать ReadOnly=True.
        wb = excel.Workbooks.Open(file_path, ReadOnly=True)

        # Проверяем количество листов
        if wb.Sheets.Count < 2:
            print(f"Предупреждение: Файл {os.path.basename(file_path)} содержит менее двух листов. Пропускаем.")
            return None

        sheet1 = wb.Sheets(1)
        sheet2 = wb.Sheets(2)

        # Применяем PrintArea, если она указана в ячейке R1
        # Важно: конвертируем значение в строку, т.к. Range("R1").Value может быть нестроковым
        print_area1 = str(sheet1.Range("R1").Value).strip()
        if print_area1:
            try:
                sheet1.PageSetup.PrintArea = print_area1
            except Exception as pa_e:
                print(f"Предупреждение: Не удалось установить PrintArea для листа 1 из '{print_area1}': {pa_e}")

        print_area2 = str(sheet2.Range("R1").Value).strip()
        if print_area2:
            try:
                sheet2.PageSetup.PrintArea = print_area2
            except Exception as pa_e:
                print(f"Предупреждение: Не удалось установить PrintArea для листа 2 из '{print_area2}': {pa_e}")

        # Выбираем оба листа для экспорта
        # Это более надёжный способ выбора нескольких листов через COM
        sheet1.Select(Replace=True)  # Выбираем первый лист
        sheet2.Select(Replace=False)  # Добавляем второй лист к выбору

        # Экспортируем выбранные листы как PDF
        # 0 = xlTypePDF (стандартная константа для формата PDF в Excel)
        wb.ActiveSheet.ExportAsFixedFormat(0, pdf_path)
        print(f"Листы 1 и 2 файла {os.path.basename(file_path)} сохранены как PDF: {pdf_path}")
        return pdf_path

    except pythoncom.com_error as com_e:
        # Более детальная обработка COM-ошибок
        print(f"\n!!! КРИТИЧЕСКАЯ COM-ОШИБКА при работе с Excel:")
        print(f"!!! Файл: {os.path.basename(file_path)}")
        print(f"!!! Ошибка: {com_e}")
        print("!!! Возможные причины:")
        print("!!! 1. Excel не установлен или поврежден.")
        print(
            "!!! 2. Проблемы с регистрацией COM-объектов (попробуйте переустановить pywin32 или восстановить Office).")
        print("!!! 3. Недостаточные права доступа (попробуйте запустить скрипт от имени администратора).")
        print("!!! 4. 'Зависшие' процессы Excel (закройте все процессы Excel вручную через Диспетчер задач).")
        print("=" * 50 + "\n")
        return None
    except Exception as e:
        print(f"Ошибка при сохранении PDF для {os.path.basename(file_path)}: {e}")
        return None
    finally:
        # Убедимся, что рабочая книга и приложение Excel корректно закрываются
        if wb:
            try:
                wb.Close(SaveChanges=False)
            except Exception as close_e:
                print(f"Предупреждение: Ошибка при закрытии рабочей книги: {close_e}")
        if excel:
            try:
                excel.Quit()
            except Exception as quit_e:
                print(f"Предупреждение: Ошибка при завершении процесса Excel: {quit_e}")
            # Дополнительно: освобождаем COM-объекты
            # import gc; gc.collect() # Иногда помогает, но может быть избыточно
            # excel = None # Отвязываемся от объекта


def parse_range(range_str):
    ranges = []
    for part in range_str.split(','):
        part = part.strip()
        if '-' in part:
            try:
                start, end = map(int, part.split('-'))
                if start > end:
                    print(f"Предупреждение: Неверный диапазон '{part}'. Начало должно быть меньше или равно концу.")
                    continue
                ranges.extend(range(start, end + 1))
            except ValueError:
                print(f"Предупреждение: Неверный формат диапазона '{part}'. Пропускаем.")
        else:
            try:
                ranges.append(int(part))
            except ValueError:
                print(f"Предупреждение: Неверный формат числа '{part}'. Пропускаем.")
    return sorted(list(set(ranges)))  # Убираем дубликаты и сортируем


def main():
    script_directory = os.path.dirname(os.path.realpath(__file__))
    path_file = os.path.join(script_directory, 'path.txt')

    source_folder, destination_folder, range_str = read_paths_and_range(path_file)

    if not (source_folder and destination_folder and range_str):
        print("Ошибка: Некорректные данные в path.txt. Завершение.")
        return

    # Проверяем существование исходной и целевой папок
    if not os.path.isdir(source_folder):
        print(f"Ошибка: Исходная папка '{source_folder}' не найдена или недоступна. Завершение.")
        return
    if not os.path.exists(destination_folder):
        os.makedirs(destination_folder)
        print(f"Создана целевая папка: {destination_folder}")

    file_numbers = parse_range(range_str)
    if not file_numbers:
        print("Не найдено номеров файлов для обработки в диапазоне. Завершение.")
        return

    pdf_folder = os.path.join(destination_folder, "PDF")
    if not os.path.exists(pdf_folder):
        os.makedirs(pdf_folder)
        print(f"Создана папка для PDF: {pdf_folder}")

    pdf_files_to_merge = []  # Переименовал для ясности
    exported_file_numbers = []  # Список для хранения номеров экспортированных файлов

    print(f"Начинаем поиск и обработку файлов в папке: {source_folder}")
    print(f"Будут обработаны файлы с номерами: {file_numbers}")

    for root, dirs, files in os.walk(source_folder):
        for file in files:
            if "Invoice" in file and file.lower().endswith(('.xlsx', '.xls')):  # .lower() для надежности
                file_num_str = ''.join(filter(str.isdigit, file))
                if not file_num_str:
                    # print(f"Предупреждение: Не удалось извлечь номер из файла '{file}'. Пропускаем.")
                    continue  # Пропускаем файлы, из которых нельзя извлечь номер

                try:
                    file_num = int(file_num_str)
                except ValueError:
                    # print(f"Предупреждение: Извлеченный номер '{file_num_str}' из файла '{file}' не является числом. Пропускаем.")
                    continue

                if file_num in file_numbers:
                    source_file_path = os.path.join(root, file)
                    destination_file_path = os.path.join(destination_folder, file)

                    try:
                        shutil.copy2(source_file_path, destination_file_path)
                        print(f"Файл '{file}' скопирован в '{destination_folder}'")

                        pdf_file_name = os.path.splitext(file)[0] + ".pdf"
                        pdf_path = os.path.join(pdf_folder, pdf_file_name)
                        result_pdf = save_two_sheets_as_pdf(destination_file_path, pdf_path)

                        if result_pdf:
                            pdf_files_to_merge.append(result_pdf)
                            exported_file_numbers.append(file_num)  # Добавляем номер файла в список
                        else:
                            print(f"Не удалось создать PDF для файла '{file}'. Пропускаем его.")
                    except Exception as copy_e:
                        print(f"Ошибка при копировании или обработке файла '{file}': {copy_e}. Пропускаем.")

    if pdf_files_to_merge:
        pdf_merger = PyPDF2.PdfMerger()
        print("\nНачинаем объединение PDF-файлов...")
        for pdf_file in pdf_files_to_merge:
            try:
                # Открываем файл в бинарном режиме для большей надежности
                with open(pdf_file, 'rb') as f:
                    pdf_merger.append(f)
                print(f"-> Добавлен временный PDF-файл: {os.path.basename(pdf_file)}")
            except Exception as e:
                print(
                    f"!!! ОШИБКА при добавлении временного PDF-файла '{os.path.basename(pdf_file)}' в объединение: {e}. Файл будет пропущен.")
                continue

        # Формируем правильное название с учетом экспортированных файлов
        exported_file_numbers_sorted = sorted(exported_file_numbers)
        ranges = []

        if exported_file_numbers_sorted:  # Проверяем, что список не пуст
            start = exported_file_numbers_sorted[0]
            for i in range(1, len(exported_file_numbers_sorted)):
                if exported_file_numbers_sorted[i] == exported_file_numbers_sorted[i - 1] + 1:
                    pass  # Продолжаем диапазон
                else:
                    if start == exported_file_numbers_sorted[i - 1]:
                        ranges.append(str(start))
                    else:
                        ranges.append(f"{start}-{exported_file_numbers_sorted[i - 1]}")
                    start = exported_file_numbers_sorted[i]
            # Добавляем последний диапазон
            if start == exported_file_numbers_sorted[-1]:
                ranges.append(str(start))
            else:
                ranges.append(f"{start}-{exported_file_numbers_sorted[-1]}")

        # Формируем итоговое название
        formatted_range = ', '.join(ranges) if ranges else "NoFiles"
        output_pdf_name = f"Inv.+Spec. {formatted_range} {len(pdf_files_to_merge)} pcs..pdf"
        output_pdf_path = os.path.join(pdf_folder, output_pdf_name)

        try:
            with open(output_pdf_path, 'wb') as f_out:
                pdf_merger.write(f_out)
            pdf_merger.close()
            print(f"\nОбъединенный файл сохранен как: '{output_pdf_name}' в папке '{pdf_folder}'")
        except Exception as write_e:
            print(f"\n!!! КРИТИЧЕСКАЯ ОШИБКА при сохранении объединенного PDF-файла: {write_e}")
            print(f"!!! Проверьте права доступа к папке '{pdf_folder}' или наличие места на диске.")
            print("=" * 50 + "\n")

        # Удаляем временные PDF-файлы
        print("\nУдаление временных PDF-файлов...")
        for pdf_file in pdf_files_to_merge:
            try:
                os.remove(pdf_file)
                print(f"Удален: {os.path.basename(pdf_file)}")
            except Exception as e:
                print(f"Ошибка при удалении временного PDF-файла '{os.path.basename(pdf_file)}': {e}")
    else:
        print("Нет PDF-файлов для объединения.")

    # Удаляем временные XLSX/XLS файлы из destination_folder
    print("\nУдаление временных XLSX/XLS файлов из целевой папки...")
    for file in os.listdir(destination_folder):
        file_path = os.path.join(destination_folder, file)
        if file.lower().endswith(('.xlsx', '.xls')) and os.path.isfile(file_path):
            try:
                os.remove(file_path)
                print(f"Удален: {file}")
            except Exception as e:
                print(f"Ошибка при удалении временного XLSX/XLS файла '{file}': {e}")

    print("\nОбработка завершена.")

else:
print("Ошибка: Некорректные данные в path.txt. Проверьте содержимое файла.")

if __name__ == "__main__":
    main()