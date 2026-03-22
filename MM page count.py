import os
import pandas as pd
from pypdf import PdfReader
import win32com.client #pip install pywin32


def count_stats():
    raw_path = input("Введите путь к ГЛАВНОЙ папке: ")
    main_folder_path = os.path.normpath(raw_path.strip().strip('"').strip("'"))

    if not os.path.exists(main_folder_path):
        print(f"❌ Путь не существует: {main_folder_path}")
        return

    print(f"\n--- Сканирование папки и всех подпапок: {main_folder_path} ---\n")

    data_list = []

    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0

    for root, dirs, files in os.walk(main_folder_path):

        for banned in ("Исходник для проверки", "архив"):
            if banned in dirs:
                dirs.remove(banned)

        if os.path.basename(root) in ("Исходник для проверки", "архив"):
            continue

        responsible_person = (
            f"{os.path.basename(root)} (Корень)"
            if root == main_folder_path
            else os.path.basename(root)
        )

        for filename in files:
            if filename.startswith("~$") or filename == "Statistics.xlsx":
                continue

            file_path = os.path.join(root, filename)
            file_stat = None
            type_str = ""

            try:
                # ---------- PDF ----------
                if filename.lower().endswith(".pdf"):
                    reader = PdfReader(file_path)
                    pages = len(reader.pages)

                    print(f"[{responsible_person}] 📄 PDF: {filename} -> {pages} стр.")
                    file_stat = pages
                    type_str = "PDF (стр)"

                # ---------- DOCX (Word COM) ----------
                elif filename.lower().endswith(".docx"):
                    doc = word.Documents.Open(
                        file_path,
                        ConfirmConversions=False,
                        ReadOnly=True,
                        AddToRecentFiles=False,
                        Revert=False
                    )

                    doc.Activate()

                    # ⚠️ Пытаемся создать окно, но не обязаны
                    try:
                        word.Windows.Add(doc)
                    except Exception:
                        pass

                    doc.Content.Select()
                    word.Selection.WholeStory()
                    word.Selection.Fields.Update()
                    doc.Repaginate()

                    char_count = doc.ComputeStatistics(5)  #5 Знаков с пробелами 3 знаки без пр 2 стр 1 строки 0 слова
                    uchet_pages = char_count #/ 1800
                    uchet = char_count/ 1800
                    print(
                        f"[{responsible_person}] 📝 DOCX: {filename} "
                        f"-> {uchet:.2f} уч. стр. (знаков: {char_count})"
                    )

                    doc.Close(False)

                    file_stat = round(uchet_pages, 2)
                    type_str = "DOCX"

                if file_stat is not None:
                    data_list.append({
                        "Название файла": filename,
                        "Объем": file_stat,
                        "Единица измерения": type_str,
                        "Ответственный (Папка)": responsible_person,
                    })

            except Exception as e:
                print(f"❌ Ошибка в файле {filename}: {e}")
                data_list.append({
                    "Название файла": filename,
                    "Объем": "ОШИБКА",
                    "Единица измерения": str(e),
                    "Ответственный (Папка)": responsible_person,
                })

    word.Quit()

    if data_list:
        df = pd.DataFrame(data_list)
        output_file = os.path.join(main_folder_path, "Statistics.xlsx")
        df.to_excel(output_file, index=False)

        print(f"\n✅ Готово! Отчет сохранен: {output_file}")
        print(f"Всего обработано файлов: {len(data_list)}")
    else:
        print("\nФайлы не найдены.")

    input("\nНажмите Enter, чтобы выйти...")


if __name__ == "__main__":
    count_stats()


#   python "d:/MM Python/MM page count.py"

# D:\AndroSync\AndroSync\MM\in_process\для учета
