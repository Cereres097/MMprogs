#pip install pandas pypdf pywin32 tkinterdnd2
#tkinter обычно уже есть в Python
#tkinterdnd2 нужен для drag & drop

import os
import threading
import pandas as pd
from pypdf import PdfReader
import win32com.client
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinterdnd2 import TkinterDnD, DND_FILES


# ---------- Парсер пути----------

def parse_dnd_paths(data: str) -> list[str]:
    paths = []
    buff = ""
    in_braces = False

    for ch in data:
        if ch == "{":
            in_braces = True
            buff = ""
        elif ch == "}":
            in_braces = False
            paths.append(buff)
            buff = ""
        elif ch == " " and not in_braces:
            if buff:
                paths.append(buff)
                buff = ""
        else:
            buff += ch

    if buff:
        paths.append(buff)

    return [os.path.normpath(p.strip().strip('"').strip("'")) for p in paths]


# ---------- GUI HELPERS ----------

def log(msg: str):
    log_box.insert(tk.END, msg + "\n")
    log_box.see(tk.END)
    root.update_idletasks()


def choose_folders():
    folder = filedialog.askdirectory()
    if folder:
        folders = folder_var.get().split(";") if folder_var.get() else []
        folders.append(folder)
        folder_var.set(";".join(dict.fromkeys(folders)))


def drop_event(event):
    paths = parse_dnd_paths(event.data)
    folders = folder_var.get().split(";") if folder_var.get() else []
    folders.extend(paths)
    folder_var.set(";".join(dict.fromkeys(folders)))


def run_count():
    threading.Thread(target=count_stats, daemon=True).start()


# ---------- CORE LOGIC ----------

def count_stats():
    folders = [p for p in folder_var.get().split(";") if p]

    if not folders:
        messagebox.showerror("Ошибка", "Папки не указаны")
        return

    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0

    for main_folder_path in folders:

        if not os.path.exists(main_folder_path):
            log(f"❌ Путь не существует: {main_folder_path}")
            continue

        log(f"\n📂 Сканирование: {main_folder_path}\n")
        data_list = []

        for root_dir, dirs, files in os.walk(main_folder_path):

            for banned in ("Исходник для проверки", "архив"):
                if banned in dirs:
                    dirs.remove(banned)

            if os.path.basename(root_dir) in ("Исходник для проверки", "архив"):
                continue

            responsible_person = (
                f"{os.path.basename(root_dir)} (Корень)"
                if root_dir == main_folder_path
                else os.path.basename(root_dir)
            )

            for filename in files:
                if filename.startswith("~$") or filename == "Statistics.xlsx":
                    continue

                file_path = os.path.join(root_dir, filename)

                try:
                    # ---------- PDF ----------
                    if filename.lower().endswith(".pdf"):
                        reader = PdfReader(file_path)
                        pages = len(reader.pages)

                        log(f"[{responsible_person}] PDF: {filename} → {pages} стр.")
                        data_list.append({
                            "Название файла": filename,
                            "Объем": pages,
                            "Единица измерения": "PDF (стр)",
                            "Ответственный (Папка)": responsible_person,
                        })

                    # ---------- DOCX ----------
                    elif filename.lower().endswith(".docx"):
                        doc = word.Documents.Open(
                            file_path,
                            ConfirmConversions=False,
                            ReadOnly=True,
                            AddToRecentFiles=False,
                            Revert=False
                        )

                        doc.Content.Select()
                        word.Selection.Fields.Update()
                        doc.Repaginate()

                        char_count = doc.ComputeStatistics(5)
                        pages = round(char_count / 1800, 2)

                        log(f"[{responsible_person}] DOCX: {filename} → {pages} уч. стр.")

                        doc.Close(False)

                        data_list.append({
                            "Название файла": filename,
                            "Объем": pages,
                            "Единица измерения": "DOCX (1800 зн)",
                            "Ответственный (Папка)": responsible_person,
                        })

                except Exception as e:
                    log(f"❌ Ошибка: {filename} → {e}")
                    data_list.append({
                        "Название файла": filename,
                        "Объем": "ОШИБКА",
                        "Единица измерения": str(e),
                        "Ответственный (Папка)": responsible_person,
                    })

        if data_list:
            df = pd.DataFrame(data_list)
            output = os.path.join(main_folder_path, "Statistics.xlsx")
            df.to_excel(output, index=False)
            log(f"✅ Отчет сохранён: {output}")
        else:
            log("Файлы не найдены")

    word.Quit()


# ---------- GUI ----------

root = TkinterDnD.Tk()
root.title("Подсчёт объёма документов")
root.geometry("900x500")

folder_var = tk.StringVar()

top = ttk.Frame(root, padding=10)
top.pack(fill=tk.X)

entry = ttk.Entry(top, textvariable=folder_var)
entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
entry.drop_target_register(DND_FILES)
entry.dnd_bind("<<Drop>>", drop_event)

ttk.Button(top, text="Добавить папку", command=choose_folders).pack(side=tk.LEFT, padx=5)
ttk.Button(top, text="Старт", command=run_count).pack(side=tk.LEFT, padx=5)

log_box = tk.Text(root, wrap=tk.WORD)
log_box.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

root.mainloop()


# Сборка в EXE (PyInstaller)
# pip install pyinstaller
# pyinstaller --onefile --noconsole --hidden-import=tkinterdnd2 page_counter_gui.py
# dist/page_counter_gui.exe

# python "d:/MM Python/page_counter_gui.py"