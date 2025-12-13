import os
import win32com.client as win32
import sys
import re
import time
import json

# ==========================================
# ЦВЕТА КОНСОЛИ (ANSI)
# ==========================================

YELLOW = "\033[33m"
RESET = "\033[0m"


# ==========================================
# КОНФИГУРАЦИЯ
# ==========================================

CONFIG_FILE = "config.json"

def load_config():
    """Загружает конфигурацию из JSON."""
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def save_config(config):
    """Сохраняет конфигурацию в JSON."""
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=4)
    except Exception as e:
        print(f"⚠ Не удалось сохранить конфиг: {e}")

# ==========================================
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ==========================================

def get_clean_path(prompt_text, saved_path=None):
    """Запрашивает путь у пользователя и удаляет кавычки."""
    try:
        print(prompt_text)

        if saved_path:
            print(f"{YELLOW}{saved_path}{RESET}")
            print(f"{YELLOW}Для продолжения нажмите Enter или введите другой путь{RESET}")

        path = input("> ").strip()

        if not path and saved_path:
            return saved_path

        if (path.startswith('"') and path.endswith('"')) or (path.startswith("'") and path.endswith("'")):
            path = path[1:-1]

        return path

    except EOFError:
        return ""

def parse_range(range_str):
    """Парсит строку диапазона (например, '1-3, 5') в список чисел."""
    ranges = []
    for part in range_str.split(','):
        part = part.strip()
        if not part:
            continue
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
    return sorted(set(ranges))

# ==========================================
# ОСНОВНАЯ ЛОГИКА EXCEL
# ==========================================

def process_excel_files(source_folder, file_numbers, mode):
    excel = None
    try:
        print("\n🚀 Запуск Excel... Пожалуйста, подождите.")
        excel = win32.DispatchEx('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False

        count_success = 0

        for root, _, files in os.walk(source_folder):
            for file in files:

                if not file.lower().endswith(('.xlsx', '.xls', '.xlsm')):
                    continue

                name_body = os.path.splitext(file)[0]

                if not re.fullmatch(r'invoice\s+\d+', name_body, re.IGNORECASE):
                    continue

                file_num_str = ''.join(filter(str.isdigit, file))
                if not file_num_str:
                    continue

                try:
                    file_num = int(file_num_str)
                except ValueError:
                    continue

                if file_num in file_numbers:
                    full_path = os.path.join(root, file)
                    pdf_name = os.path.splitext(file)[0] + ".pdf"
                    save_path = os.path.join(root, pdf_name)

                    print(f"➡️ Обработка: {file}")

                    if convert_workbook(excel, full_path, save_path, mode):
                        count_success += 1
                        print(f"   ✅ Готово: {save_path}")
                    else:
                        print(f"   ❌ Ошибка конвертации")

        print(f"\n🏁 ИТОГ: Успешно создано файлов: {count_success}")
        print("-" * 30)

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

        sheets_to_export = [wb.Sheets(1), wb.Sheets(2)]
        sheet_names = [wb.Sheets(1).Name, wb.Sheets(2).Name]

        if mode == '2':
            target_names = ["Weight certificate (LI)", "Weight certificate (Y)"]
            XL_SHEET_VISIBLE = -1

            for sheet in wb.Sheets:
                if sheet.Name in target_names and sheet.Visible == XL_SHEET_VISIBLE:
                    if sheet.Name not in sheet_names:
                        sheets_to_export.append(sheet)
                        sheet_names.append(sheet.Name)

        for sheet in sheets_to_export:
            try:
                print_area = sheet.Range("R1").Value
                if print_area:
                    sheet.PageSetup.PrintArea = str(print_area)
            except:
                pass

        wb.Sheets(sheet_names[0]).Select()
        for i in range(1, len(sheet_names)):
            wb.Sheets(sheet_names[i]).Select(False)

        wb.ActiveSheet.ExportAsFixedFormat(0, pdf_path)
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
    config = load_config()
    last_path = config.get("source_path")

    while True:
        print("\n" + "=" * 50)
        print("   УТИЛИТА ЭКСПОРТА EXCEL -> PDF")
        print("=" * 50)

        print("Выберите действие:")
        print("1. Инвойс и спецификация")
        print("2. Инвойс, спецификация и весовой сертификат")
        print("0. Выход из программы")

        mode_choice = input("\nВаш выбор (0-2): ").strip()

        if mode_choice == '0':
            print("Всего доброго!")
            break

        if mode_choice not in ['1', '2']:
            print("❌ Ошибка: Неверный выбор.")
            continue

        print()
        source_path = get_clean_path(
            "Укажите путь к директории (или 'menu' для отмены):",
            last_path
        )

        if source_path.lower() == 'menu':
            continue

        if not os.path.isdir(source_path):
            print("❌ Ошибка: Указанная папка не существует.")
            continue

        # сохраняем путь
        config["source_path"] = source_path
        save_config(config)
        last_path = source_path

        range_input = input("Укажите диапазон номеров (например: 3550-3553,3560): ").strip()
        file_numbers = parse_range(range_input)

        if not file_numbers:
            print("❌ Не указан корректный диапазон.")
            continue

        process_excel_files(source_path, file_numbers, mode_choice)

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nПрограмма остановлена пользователем.")
    except Exception as e:
        print("\n" + "!"*50)
        print(f"КРИТИЧЕСКАЯ ОШИБКА: {e}")
        print("!"*50)
        import traceback
        traceback.print_exc()
    finally:
        print("\nРабота завершена.")
        input("Нажмите Enter, чтобы закрыть окно...")
