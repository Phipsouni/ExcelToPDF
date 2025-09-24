import os
import win32com.client as win32  # Используем win32com для работы с Excel


# --- Функция для чтения путей и диапазона ---
def read_paths_and_range(file_path):
    """
    Читает пути к исходной папке и диапазон номеров файлов из файла path.txt.
    """
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            lines = [line.strip() for line in f if line.strip()]
            if len(lines) < 2:
                raise ValueError(
                    "Файл path.txt должен содержать две строки: путь к исходной папке и диапазон."
                )
            source_folder = lines[0]
            range_str = lines[1]
            return source_folder, range_str
    except FileNotFoundError:
        print(f"Ошибка: Файл '{file_path}' не найден. Убедитесь, что он существует.")
        return None, None
    except Exception as e:
        print(f"Ошибка при чтении path.txt: {e}")
        return None, None


# --- ИЗМЕНЕНИЕ: Функция была переработана для экспорта дополнительных листов ---
def save_sheets_as_pdf(file_path, pdf_path):
    """
    Сохраняет первые два листа и, если найдены и видимы, листы с именами
    'Weight certificate (LI)' и 'Weight certificate (Y)' в один PDF-файл.
    Также обрабатывает PrintArea из ячейки R1 для каждого экспортируемого листа.
    """
    excel = None
    wb = None
    # Константа для проверки видимости листа (xlSheetVisible = -1)
    XL_SHEET_VISIBLE = -1

    try:
        excel = win32.DispatchEx('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False

        wb = excel.Workbooks.Open(file_path, ReadOnly=True)

        if wb.Sheets.Count < 2:
            print(f"Предупреждение: Файл {os.path.basename(file_path)} содержит менее двух листов. Пропускаем.")
            return None

        # --- ИЗМЕНЕНИЕ: Формируем список листов для экспорта ---
        sheets_to_export = [wb.Sheets(1), wb.Sheets(2)]
        sheet_names_to_export = [wb.Sheets(1).Name, wb.Sheets(2).Name]

        # Имена целевых листов
        target_sheet_names = ["Weight certificate (LI)", "Weight certificate (Y)"]

        # Ищем дополнительные листы
        for sheet in wb.Sheets:
            # Проверяем имя, видимость и то, что лист еще не в списке
            if sheet.Name in target_sheet_names and sheet.Visible == XL_SHEET_VISIBLE and sheet.Name not in sheet_names_to_export:
                sheets_to_export.append(sheet)
                sheet_names_to_export.append(sheet.Name)
                
        # --- ИЗМЕНЕНИЕ: Применяем PrintArea ко всем выбранным листам ---
        for sheet in sheets_to_export:
            print_area = sheet.Range("R1").Value
            if print_area:
                try:
                    sheet.PageSetup.PrintArea = str(print_area)
                except Exception as pa_e:
                    print(f"Предупреждение: Не удалось установить PrintArea для листа '{sheet.Name}' из {print_area}: {pa_e}")
        
        # --- ИЗМЕНЕНИЕ: Выбираем все необходимые листы для экспорта ---
        # wb.Worksheets(sheet_names_to_export).Select() # Этот метод выбирает листы, делая первый активным
        
        # Более надежный способ выбора нескольких листов для некоторых версий Excel
        wb.Sheets(sheet_names_to_export[0]).Select()
        for i in range(1, len(sheet_names_to_export)):
            wb.Sheets(sheet_names_to_export[i]).Select(False) # Replace=False добавляет лист к выделению

        # Экспортируем выбранные листы как PDF
        wb.ActiveSheet.ExportAsFixedFormat(0, pdf_path)

        print(f"Листы {sheet_names_to_export} файла {os.path.basename(file_path)} сохранены как PDF: {pdf_path}")
        return pdf_path

    except Exception as e:
        print(f"Ошибка при сохранении PDF для {os.path.basename(file_path)}: {e}")
        return None
    finally:
        if wb:
            try:
                wb.Close(SaveChanges=False)
            except Exception as close_e:
                print(f"Ошибка при закрытии рабочей книги: {close_e}")
        if excel:
            try:
                excel.Quit()
            except Exception as quit_e:
                print(f"Ошибка при завершении процесса Excel: {quit_e}")


# --- Функция для парсинга диапазона чисел ---
def parse_range(range_str):
    """
    Парсит строку диапазона чисел (например, '1-3, 5, 10-12') в список уникальных чисел.
    """
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
    return sorted(list(set(ranges)))


# --- Главная функция скрипта ---
def main():
    script_directory = os.path.dirname(os.path.realpath(__file__))
    path_file = os.path.join(script_directory, 'path.txt')

    source_folder, range_str = read_paths_and_range(path_file)

    if not (source_folder and range_str):
        print("Ошибка: Не удалось получить необходимые данные из path.txt. Завершение.")
        return

    if not os.path.isdir(source_folder):
        print(f"Ошибка: Исходная папка '{source_folder}' не найдена или недоступна. Завершение.")
        return

    file_numbers_to_process = parse_range(range_str)

    if not file_numbers_to_process:
        print("Не найдено номеров файлов для обработки в диапазоне. Завершение.")
        return

    print(f"Будут обработаны файлы с номерами: {file_numbers_to_process}")

    found_and_processed_files = 0

    for root, _, files in os.walk(source_folder):
        for file in files:
            if "Invoice" in file and file.endswith(('.xlsx', '.xls')):
                file_num_str = ''.join(filter(str.isdigit, file))
                if not file_num_str:
                    print(f"Предупреждение: Не удалось извлечь номер из файла '{file}'. Пропускаем.")
                    continue

                try:
                    file_num = int(file_num_str)
                except ValueError:
                    print(f"Предупреждение: Некорректный номер файла '{file_num_str}' в '{file}'. Пропускаем.")
                    continue

                if file_num in file_numbers_to_process:
                    source_file_path = os.path.join(root, file)
                    pdf_file_name = os.path.splitext(file)[0] + ".pdf"
                    pdf_path = os.path.join(root, pdf_file_name)

                    print(f"Обработка файла: {source_file_path}")
                    # --- ИЗМЕНЕНИЕ: Вызываем обновленную функцию ---
                    result_pdf = save_sheets_as_pdf(source_file_path, pdf_path)
                    if result_pdf:
                        found_and_processed_files += 1
                    else:
                        print(f"Не удалось создать PDF для файла: {file}")

    if found_and_processed_files > 0:
        print(f"Обработка завершена. Создано {found_and_processed_files} PDF-файлов.")
    else:
        print("Не найдено подходящих файлов для обработки в указанном диапазоне.")


if __name__ == "__main__":
    main()