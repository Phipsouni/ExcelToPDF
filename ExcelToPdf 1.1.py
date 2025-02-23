import os
import shutil
import win32com.client as win32
import PyPDF2

def read_paths_and_range(file_path):
    try:
        with open(file_path, 'r') as f:
            lines = f.read().strip().split("\n")
            if len(lines) < 3:
                raise ValueError(
                    "Файл path.txt должен содержать три строки: путь к исходной папке, путь к целевой папке и диапазон.")
            source_folder = lines[0].strip()
            destination_folder = lines[1].strip()
            range_str = lines[2].strip()
            return source_folder, destination_folder, range_str
    except Exception as e:
        print(f"Ошибка при чтении path.txt: {e}")
        return None, None, None

def save_two_sheets_as_pdf(file_path, pdf_path):
    try:
        excel = win32.DispatchEx('Excel.Application')
        excel.Visible = False
        wb = excel.Workbooks.Open(file_path)
        sheet1 = wb.Sheets(1)
        sheet2 = wb.Sheets(2)
        print_area1 = sheet1.Range("R1").Value
        if print_area1:
            sheet1.PageSetup.PrintArea = print_area1
        print_area2 = sheet2.Range("R1").Value
        if print_area2:
            sheet2.PageSetup.PrintArea = print_area2
        wb.Worksheets([1, 2]).Select()
        wb.ActiveSheet.ExportAsFixedFormat(0, pdf_path)
        print(f"Листы 1 и 2 файла {os.path.basename(file_path)} сохранены как PDF: {pdf_path}")
        wb.Close(SaveChanges=False)
        excel.Quit()
        return pdf_path
    except Exception as e:
        print(f"Ошибка при сохранении PDF: {e}")
        if 'excel' in locals():
            excel.Quit()
    return None


def parse_range(range_str):
    ranges = []
    for part in range_str.split(','):
        part = part.strip()
        if '-' in part:
            start, end = map(int, part.split('-'))
            ranges.extend(range(start, end + 1))
        else:
            ranges.append(int(part))
    # Убираем дубликаты и сортируем
    return sorted(set(ranges))


def main():
    script_directory = os.path.dirname(os.path.realpath(__file__))
    path_file = os.path.join(script_directory, 'path.txt')

    source_folder, destination_folder, range_str = read_paths_and_range(path_file)

    if source_folder and destination_folder and range_str:
        # Получаем все файлы, которые нужно обработать
        file_numbers = parse_range(range_str)
        pdf_folder = os.path.join(destination_folder, "PDF")
        if not os.path.exists(pdf_folder):
            os.makedirs(pdf_folder)

        pdf_files = []
        exported_files = []  # Список для хранения номеров экспортированных файлов
        for root, dirs, files in os.walk(source_folder):
            for file in files:
                if "Invoice" in file and file.endswith(('.xlsx', '.xls')):
                    file_num = int(''.join(filter(str.isdigit, file)))
                    if file_num in file_numbers:
                        source_file_path = os.path.join(root, file)
                        destination_file_path = os.path.join(destination_folder, file)
                        shutil.copy2(source_file_path, destination_file_path)
                        print(f"Файл {file} скопирован в {destination_folder}")

                        pdf_file_name = os.path.splitext(file)[0] + ".pdf"
                        pdf_path = os.path.join(pdf_folder, pdf_file_name)
                        result_pdf = save_two_sheets_as_pdf(destination_file_path, pdf_path)
                        if result_pdf:
                            pdf_files.append(result_pdf)
                            exported_files.append(file_num)  # Добавляем номер файла в список

        if pdf_files:
            pdf_merger = PyPDF2.PdfMerger()
            for pdf_file in pdf_files:
                pdf_merger.append(pdf_file)

            # Формируем правильное название с учетом экспортированных файлов
            exported_files_sorted = sorted(exported_files)
            ranges = []
            start = exported_files_sorted[0]
            for i in range(1, len(exported_files_sorted)):
                if exported_files_sorted[i] != exported_files_sorted[i - 1] + 1:
                    if start == exported_files_sorted[i - 1]:
                        ranges.append(str(start))
                    else:
                        ranges.append(f"{start}-{exported_files_sorted[i - 1]}")
                    start = exported_files_sorted[i]
            # Добавляем последний диапазон
            if start == exported_files_sorted[-1]:
                ranges.append(str(start))
            else:
                ranges.append(f"{start}-{exported_files_sorted[-1]}")

            # Формируем итоговое название
            formatted_range = ', '.join(ranges)
            output_pdf_name = f"Inv.+Spec. {formatted_range} {len(pdf_files)} pcs..pdf"
            output_pdf_path = os.path.join(pdf_folder, output_pdf_name)
            pdf_merger.write(output_pdf_path)
            pdf_merger.close()
            print(f"Объединенный файл сохранен как: {output_pdf_name}")

            for pdf_file in pdf_files:
                os.remove(pdf_file)
                print(f"Удален временный PDF-файл: {pdf_file}")

        for file in os.listdir(destination_folder):
            file_path = os.path.join(destination_folder, file)
            if file.endswith(('.xlsx', '.xls')) and os.path.isfile(file_path):
                os.remove(file_path)
                print(f"Удален временный файл: {file_path}")

        print("Обработка завершена.")
    else:
        print("Ошибка: Некорректные данные в path.txt.")


if __name__ == "__main__":
    main()
