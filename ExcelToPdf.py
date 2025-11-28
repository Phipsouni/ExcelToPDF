import os
import win32com.client as win32
import sys


# ==========================================
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ==========================================

def get_clean_path(prompt_text):
    """Запрашивает путь у пользователя и удаляет кавычки."""
    path = input(f"{prompt_text}: ").strip()
    # Удаляем кавычки в начале и конце
    if (path.startswith('"') and path.endswith('"')) or (path.startswith("'") and path.endswith("'")):
        path = path[1:-1]
    return path


def parse_range(range_str):
    """Парсит строку диапазона (например, '1-3, 5') в список чисел."""
    ranges = []
    for part in range_str.split(','):
        part = part.strip()
        if not part: continue
        if '-' in part:
            try:
                start, end = map(int, part.split('-'))
                if start > end:
                    print(f"⚠ Предупреждение: Неверный диапазон '{part}'.")
                    continue
                ranges.extend(range(start, end + 1))
            except ValueError:
                print(f"⚠ Предупреждение: Неверный формат диапазона '{part}'.")
        else:
            try:
                ranges.append(int(part))
            except ValueError:
                print(f"⚠ Предупреждение: Неверный формат числа '{part}'.")
    return sorted(list(set(ranges)))


# ==========================================
# ОСНОВНАЯ ЛОГИКА EXCEL
# ==========================================

def process_excel_files(source_folder, file_numbers, mode):
    """
    mode 1: Инвойс и спецификация (Первые 2 листа)
    mode 2: Инвойс, спец. и весовой (Первые 2 листа + Weight certificate)
    """
    excel = None
    try:
        print("\n🚀 Запуск Excel... Пожалуйста, подождите.")
        excel = win32.DispatchEx('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False

        count_success = 0

        # Проход по файлам
        for root, _, files in os.walk(source_folder):
            for file in files:
                # Фильтр по расширениям и имени
                if "invoice" in file.lower() and file.lower().endswith(('.xlsx', '.xls', '.xlsm')):

                    # Извлечение номера файла
                    file_num_str = ''.join(filter(str.isdigit, file))
                    if not file_num_str:
                        continue

                    try:
                        file_num = int(file_num_str)
                    except ValueError:
                        continue

                    if file_num in file_numbers:
                        full_path = os.path.join(root, file)

                        # Путь сохранения - ВСЕГДА рядом с исходным файлом
                        pdf_name = os.path.splitext(file)[0] + ".pdf"
                        save_path = os.path.join(root, pdf_name)

                        print(f"➡️ Обработка: {file}")

                        # --- КОНВЕРТАЦИЯ ---
                        if convert_workbook(excel, full_path, save_path, mode):
                            count_success += 1
                            print(f"   ✅ Готово: {save_path}")
                        else:
                            print(f"   ❌ Ошибка конвертации")

        print(f"\n🏁 ИТОГ: Успешно создано файлов: {count_success}")
        print("-" * 30)  # Разделитель для визуальной чистоты перед возвратом в меню

    except Exception as e:
        print(f"🔥 Критическая ошибка Excel: {e}")
    finally:
        if excel:
            try:
                excel.Quit()
                print("Excel процесс закрыт.")
            except:
                pass


def convert_workbook(excel_app, file_path, pdf_path, mode):
    wb = None
    try:
        wb = excel_app.Workbooks.Open(file_path, ReadOnly=True)

        if wb.Sheets.Count < 2:
            print("   ⚠ В файле меньше 2 листов.")
            return False

        # 1. Формируем список листов для экспорта
        # Всегда берем первые два листа
        sheets_to_export = [wb.Sheets(1), wb.Sheets(2)]
        sheet_names = [wb.Sheets(1).Name, wb.Sheets(2).Name]

        # Если выбран режим 2 (с весовыми сертификатами)
        if mode == '2':
            target_names = ["Weight certificate (LI)", "Weight certificate (Y)"]
            XL_SHEET_VISIBLE = -1

            for sheet in wb.Sheets:
                if sheet.Name in target_names and sheet.Visible == XL_SHEET_VISIBLE:
                    # Проверяем, чтобы не добавить дубликат
                    if sheet.Name not in sheet_names:
                        sheets_to_export.append(sheet)
                        sheet_names.append(sheet.Name)

        # 2. Обработка PrintArea (Ячейка R1)
        for sheet in sheets_to_export:
            try:
                print_area = sheet.Range("R1").Value
                if print_area:
                    sheet.PageSetup.PrintArea = str(print_area)
            except:
                pass  # Если ошибка в R1, просто игнорируем

        # 3. Выделение листов
        # Сначала снимаем выделение со всего, выбрав первый целевой лист
        wb.Sheets(sheet_names[0]).Select()
        # Добавляем остальные к выделению
        for i in range(1, len(sheet_names)):
            wb.Sheets(sheet_names[i]).Select(False)  # False = добавить к текущему выделению

        # 4. Экспорт
        wb.ActiveSheet.ExportAsFixedFormat(0, pdf_path)  # 0 = PDF
        return True

    except Exception as e:
        print(f"   Ошибка внутри файла: {e}")
        return False
    finally:
        if wb:
            try:
                wb.Close(SaveChanges=False)
            except:
                pass


# ==========================================
# ГЛАВНОЕ МЕНЮ
# ==========================================

def main():
    while True:
        print("\n" + "=" * 50)
        print("   МАСТЕР ЭКСПОРТА EXCEL -> PDF")
        print("=" * 50)

        print("Выберите действие:")
        print("1. Инвойс и спецификация")
        print("2. Инвойс, спецификация и весовой сертификат")
        print("0. Выход из программы")

        mode_choice = input("\nВаш выбор (0-2): ").strip()

        # Обработка выхода
        if mode_choice == '0':
            print("Всего доброго!")
            break

        # Проверка корректности ввода
        if mode_choice not in ['1', '2']:
            print("❌ Ошибка: Неверный выбор. Пожалуйста, введите 1, 2 или 0.")
            continue  # Возврат в начало цикла

        # Шаг 2: Путь к инвойсам
        source_path = get_clean_path("\nУкажите путь к директории (или введите 'menu' для отмены)")

        # Возможность вернуться в меню, если передумали на этапе ввода пути
        if source_path.lower() == 'menu':
            continue

        if not os.path.isdir(source_path):
            print("❌ Ошибка: Указанная папка не существует.")
            continue  # Возврат в начало цикла

        # Шаг 3: Диапазон
        range_input = input("Укажите диапазон номеров (например: 3550-3553,3560): ").strip()
        file_numbers = parse_range(range_input)
        if not file_numbers:
            print("❌ Не указан корректный диапазон.")
            continue  # Возврат в начало цикла

        # Запуск процесса
        process_excel_files(source_path, file_numbers, mode_choice)

        # После завершения функции процесс не умирает, а цикл while True возвращает нас в начало


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nПрограмма остановлена пользователем.")