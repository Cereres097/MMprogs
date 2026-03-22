import os
import threading
import pandas as pd
from pypdf import PdfReader
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinterdnd2 import TkinterDnD, DND_FILES

# ---------- ПАРСЕР ПУТЕЙ ----------

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

# ---------- GUI helpers ----------

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

# ---------- Основная логика ----------

def count_stats():
    folders = [p for p in folder_var.get().split(";") if p]

    if not folders:
        messagebox.showerror("Ошибка", "Папки не указаны")
        return

    for main_folder_path in folders:
        if not os.path.exists(main_folder_path):
            log(f"❌ Путь не существует: {main_folder_path}")
            continue

        log(f"\n📂 Сканирование: {main_folder_path}\n")
        
        processed_files = []

        for root_dir, dirs, files in os.walk(main_folder_path):
            for banned in (""): # папки для исключения, строки через запятую
                if banned in dirs:
                    dirs.remove(banned)

            responsible_person = (
                f"{os.path.basename(root_dir)}"
                if root_dir == main_folder_path
                else os.path.basename(root_dir)
            )

            for filename in files:
                if filename.lower().endswith(".pdf"):
                    file_path = os.path.join(root_dir, filename)
                    try:
                        reader = PdfReader(file_path)
                        pages = len(reader.pages)
                        
                        log(f"📄 {filename} → {pages} стр.")
                        
                        processed_files.append({
                            "full_path": file_path,
                            "root_dir": root_dir,
                            "old_filename": filename,
                            "pages": pages,
                            "responsible": responsible_person
                        })
                    except Exception as e:
                        log(f"❌ Ошибка в файле {filename}: {e}")

        if not processed_files:
            log(f"В папке {main_folder_path} PDF файлы не найдены.")
            continue


        # --- ВОПРОС 1: ПЕРЕИМЕНОВАНИЕ ---
        ans_rename = messagebox.askyesno("Переименование", 
                                        f"Найдено PDF: {len(processed_files)} в папке {os.path.basename(main_folder_path)}.\n"
                                        f"Добавить количество страниц в названия файлов?")
        
        updated_data_for_excel = []
        
        if ans_rename:
            for file_info in processed_files:
                old_path = file_info["full_path"]
                name_part, ext = os.path.splitext(file_info["old_filename"])
                
                # Проверка, чтобы не добавлять число дважды
                suffix = f"_{file_info['pages']}"
                if name_part.endswith(suffix):
                    new_filename = file_info["old_filename"]
                else:
                    new_filename = f"{name_part}{suffix}{ext}"
                
                new_path = os.path.join(file_info["root_dir"], new_filename)
                
                try:
                    if old_path != new_path:
                        # Если файл с таким именем уже есть, пробуем переименовать
                        os.rename(old_path, new_path)
                    
                    updated_data_for_excel.append({
                        "Название файла": new_filename,
                        "Объем": file_info["pages"],
                        "Единица измерения": "PDF (стр)",
                        "Папка": file_info["responsible"]
                    })
                except Exception as e:
                    log(f"❌ Не удалось переименовать {file_info['old_filename']}: {e}")
            log("✅ Переименование завершено.")
        else:
            for file_info in processed_files:
                updated_data_for_excel.append({
                    "Название файла": file_info["old_filename"],
                    "Объем": file_info["pages"],
                    "Единица измерения": "PDF (стр)",
                    "Папка": file_info["responsible"]
                })

        # --- ВОПРОС 2: ЭКСЕЛЬ ---
        ans_excel = messagebox.askyesno("Отчет", "Вывести отчет в формате Excel?")
        if ans_excel:
            try:
                df = pd.DataFrame(updated_data_for_excel)
                output = os.path.join(main_folder_path, "Statistics_PDF.xlsx")
                df.to_excel(output, index=False)
                log(f"✅ Отчет сохранён: {output}")
            except Exception as e:
                log(f"❌ Ошибка сохранения Excel: {e}")

    log("\n--- Работа завершена ---")

# ---------- GUI ----------

root = TkinterDnD.Tk()
root.title("Счетчик PDF страниц")
root.geometry("800x500")

folder_var = tk.StringVar() 

top = ttk.Frame(root, padding=10)
top.pack(fill=tk.X)

entry = ttk.Entry(top, textvariable=folder_var)
entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
entry.drop_target_register(DND_FILES)
entry.dnd_bind("<<Drop>>", drop_event)

ttk.Button(top, text="Обзор папок", command=choose_folders).pack(side=tk.LEFT, padx=5)
ttk.Button(top, text="Начать", command=run_count).pack(side=tk.LEFT, padx=5)

log_box = tk.Text(root, wrap=tk.WORD, bg="#f8f8f8")
log_box.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

root.mainloop()


# python "d:/MM Python/pdf_only_counter.py"
# pyinstaller --onefile --noconsole --hidden-import=tkinterdnd2 pdf_only_counter.py