import os
import win32com.client as win32  # Используем win32com для работы с Excel


# --- Функция для чтения путей и диапазона ---
def read_paths_and_range(file_path):
    """
    Читает пути к исходной папке и диапазон номеров файлов из файла path.txt.
    Теперь не нужна целевая папка, так как PDF будут сохраняться в исходной.
    """
    try:
        with open(file_path, 'r', encoding='utf-8') as f:  # Указываем кодировку для надежности
            lines = [line.strip() for line in f if line.strip()]  # Убираем пустые строки
            if len(lines) < 2:  # Теперь нужно только 2 строки: исходная папка и диапазон
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


# --- Функция для сохранения двух листов Excel в PDF ---
def save_two_sheets_as_pdf(file_path, pdf_path):
    """
    Сохраняет первые два листа указанного Excel-файла в PDF.
    Обрабатывает PrintArea из ячейки R1.
    """
    excel = None  # Инициализируем excel вне блока try, чтобы он был доступен в finally
    try:
        excel = win32.DispatchEx('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False  # Отключаем всплывающие окна Excel

        # Открываем книгу в режиме только для чтения, чтобы избежать проблем с блокировкой
        wb = excel.Workbooks.Open(file_path, ReadOnly=True)

        # Проверяем, что в книге достаточно листов
        if wb.Sheets.Count < 2:
            print(f"Предупреждение: Файл {os.path.basename(file_path)} содержит менее двух листов. Пропускаем.")
            wb.Close(SaveChanges=False)
            return None

        sheet1 = wb.Sheets(1)
        sheet2 = wb.Sheets(2)

        # Применяем PrintArea, если она указана в ячейке R1
        print_area1 = sheet1.Range("R1").Value
        if print_area1:
            try:
                sheet1.PageSetup.PrintArea = str(print_area1)
            except Exception as pa_e:
                print(f"Предупреждение: Не удалось установить PrintArea для листа 1 из {print_area1}: {pa_e}")

        print_area2 = sheet2.Range("R1").Value
        if print_area2:
            try:
                sheet2.PageSetup.PrintArea = str(print_area2)
            except Exception as pa_e:
                print(f"Предупреждение: Не удалось установить PrintArea для листа 2 из {print_area2}: {pa_e}")

        # Выбираем оба листа для экспорта
        wb.Worksheets(1).Select()
        wb.Worksheets(2).Select(Replace=False)  # Select(Replace=False) для выбора нескольких листов

        # Экспортируем выбранные листы как PDF
        wb.ActiveSheet.ExportAsFixedFormat(0, pdf_path)  # 0 = xlTypePDF
        print(f"Листы 1 и 2 файла {os.path.basename(file_path)} сохранены как PDF: {pdf_path}")
        return pdf_path

    except Exception as e:
        print(f"Ошибка при сохранении PDF для {os.path.basename(file_path)}: {e}")
        return None
    finally:
        # Убедимся, что Excel закрывается, даже если произошла ошибка
        if 'wb' in locals() and wb:
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
    return sorted(list(set(ranges)))  # Убираем дубликаты и сортируем


# --- Главная функция скрипта ---
def main():
    script_directory = os.path.dirname(os.path.realpath(__file__))
    path_file = os.path.join(script_directory, 'path.txt')

    source_folder, range_str = read_paths_and_range(path_file)

    if not (source_folder and range_str):
        print("Ошибка: Не удалось получить необходимые данные из path.txt. Завершение.")
        return

    # Проверяем существование исходной папки
    if not os.path.isdir(source_folder):
        print(f"Ошибка: Исходная папка '{source_folder}' не найдена или недоступна. Завершение.")
        return

    file_numbers_to_process = parse_range(range_str)

    if not file_numbers_to_process:
        print("Не найдено номеров файлов для обработки в диапазоне. Завершение.")
        return

    print(f"Будут обработаны файлы с номерами: {file_numbers_to_process}")

    found_and_processed_files = 0

    # Проходим по всем подпапкам в исходной директории
    for root, _, files in os.walk(source_folder):
        for file in files:
            # Ищем файлы, содержащие "Invoice" и имеющие расширение xlsx/xls
            if "Invoice" in file and file.endswith(('.xlsx', '.xls')):
                # Извлекаем номер из имени файла
                # Предполагаем, что номер - это первая последовательность цифр в имени файла
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
                    # Сохраняем PDF в ту же папку, откуда взят XLSX-файл
                    pdf_path = os.path.join(root, pdf_file_name)

                    print(f"Обработка файла: {source_file_path}")
                    result_pdf = save_two_sheets_as_pdf(source_file_path, pdf_path)
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