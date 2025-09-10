import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import threading
import pandas as pd
import easyocr
import re
import difflib
from pathlib import Path
from PIL import Image
import string
import sys

# ---------------------- Настройки ---------------------- #
DD_list = ['Lnl', 'Nebovesna', 'Runbott', 'Trpvz', 'Pesdaliss', 'Oguricap', 'Revanx',
           'Luthicx', 'Olven', 'Скуфнатраппере', 'Владосхристос', 'Zshturmovik', 'Арбузбек',
           'Rabbittt', 'Срал', 'Sheeeshh', 'Shzs', 'Невсегдасвятой','Хорошиймальчик',
           'Гламурныйахэгао', 'Вожакстаданегрилл', 'Pesdexely', 'Hikikomorri','Secretquest', 'Iletyouhide',
           'Сомнительнополезен', 'Dimonishzv', 'Слабейшеебедствие', 'Особонеопасен','Пиуупиу', 'Ксюшасуперкрутая',
           'Ксюшаоченькрутая']

stop_flag = False
df_global = pd.DataFrame()
LOG_FILE = Path(sys.executable).parent / "logs.txt"

# ---------------------- Помощники ---------------------- #
def safe_filename(name):
    valid_chars = string.ascii_letters + string.digits + "_"
    return ''.join(c if c in valid_chars else "_" for c in name)

def correct_nick(text_block):
    words = re.findall(r'\b\w+\b', text_block)
    corrected_nick = None
    for word in words:
        matches = difflib.get_close_matches(word, DD_list, n=1, cutoff=0.6)
        if matches:
            corrected_nick = matches[0]
            text_block = re.sub(r'\b' + re.escape(word) + r'\b', '', text_block, count=1)
            break
    if corrected_nick:
        text_block = corrected_nick + ' ' + text_block.strip()
    return text_block

def finalize_block(block_text):
    nick_match = re.match(r'^\s*(\w+)', block_text)
    nick = nick_match.group(1) if nick_match else ""

    class_match = re.search(r'Класс[:;\s]*([^\s(]+)', block_text)
    cls = class_match.group(1) if class_match else ""

    numbers = []
    pvp_match = re.search(r'Очки.*', block_text, re.IGNORECASE)
    if pvp_match:
        tail = pvp_match.group(0)
        numbers = re.findall(r'\d+', tail)

    first_two = []
    i = 0
    while i < len(numbers) and len(first_two) < 2:
        num = numbers[i]
        if len(num) < 5 and i + 1 < len(numbers):
            combined = num + numbers[i + 1]
            first_two.append(int(combined))
            i += 2
        else:
            first_two.append(int(num))
            i += 1

    cleaned = f"{nick} Класс: {cls} {' '.join(map(str, first_two))}".strip()
    return cleaned

def write_log(filename, folder_name, raw_text, final_text):
    with LOG_FILE.open('a', encoding='utf-8') as f:
        f.write(f"--- {folder_name}/{filename} ---\n")
        f.write("OCR текст:\n")
        f.write(raw_text + "\n")
        f.write("После обработки:\n")
        f.write(final_text + "\n\n")

def resize_if_needed(image_path, min_width=1800, min_height=600):
    with Image.open(image_path) as img:
        width, height = img.size
        if width != min_width or height != min_height:
            img = img.resize((min_width, min_height), Image.LANCZOS)
            img.save(image_path)
            print(f"Resized {image_path} to {min_width}x{min_height}")

# ---------------------- Лог ---------------------- #
def clear_log():
    """Удаляем старый лог и создаем пустой файл"""
    if LOG_FILE.exists():
        LOG_FILE.unlink()
    LOG_FILE.touch()

# ---------------------- Обработка файлов ---------------------- #
def process_files(file_list, progress_var, folder_name):
    global stop_flag
    reader = easyocr.Reader(['en', 'ru'], gpu=True)
    results = {}

    for i, file_str in enumerate(file_list):
        if stop_flag:
            break
        file_path = Path(file_str)
        # Переименование для безопасных путей
        new_name = safe_filename(file_path.stem) + file_path.suffix
        safe_path = file_path.parent / new_name
        if safe_path != file_path:
            file_path.rename(safe_path)
        file_path = safe_path

        resize_if_needed(file_path)
        text_result = reader.readtext(str(file_path), detail=0, paragraph=True)
        full_text = " ".join(text_result)
        corrected = correct_nick(full_text)
        cleaned = finalize_block(corrected)

        write_log(file_path.name, folder_name, full_text, cleaned)
        results[file_path.name] = cleaned
        progress_var.set(int((i+1)/len(file_list)*100))

    return results

# ---------------------- Поток обработки ---------------------- #
def start_processing():
    global stop_flag, df_global
    stop_flag = False

    # Очистка лога перед началом обработки
    clear_log()

    before_files = before_listbox.get(0, tk.END)
    after_files = after_listbox.get(0, tk.END)
    if not before_files or not after_files:
        messagebox.showwarning("Ошибка", "Добавьте файлы до и после")
        return

    progress_var.set(0)
    table_text.delete(1.0, tk.END)

    def worker():
        global df_global
        before_data = process_files(before_files, progress_var, folder_name='before')
        after_data = process_files(after_files, progress_var, folder_name='after')

        table = []
        for key in before_data:
            b = before_data[key].split()
            nick_b = b[0]
            a_values = None
            for a_key, a_val in after_data.items():
                if a_val.startswith(nick_b):
                    a_values = a_val.split()
                    break
            if a_values and len(b) >= 4 and len(a_values) >= 4:
                nick = b[0]
                cls = b[2]
                honor_diff = int(a_values[3]) - int(b[3])
                kills_diff = int(a_values[4]) - int(b[4])
                table.append([nick, cls, honor_diff, kills_diff])

        df_global = pd.DataFrame(table, columns=["Ник", "Класс", "Хонор", "Киллы"])
        df_global.sort_values("Киллы", ascending=False, inplace=True)
        table_text.insert(tk.END, df_global.to_string(index=False))
        messagebox.showinfo("Готово", "Обработка завершена, таблица готова")

    threading.Thread(target=worker).start()

def stop_processing():
    global stop_flag
    stop_flag = True

# ---------------------- Работа с файлами ---------------------- #
def add_files(listbox):
    files = filedialog.askopenfilenames(filetypes=[("PNG Files","*.png")])
    for f in files:
        listbox.insert(tk.END, str(Path(f)))  # поддержка кириллицы

def remove_selected(listbox):
    selected = listbox.curselection()
    for i in reversed(selected):
        listbox.delete(i)

def save_table_excel():
    global df_global
    if df_global.empty:
        messagebox.showwarning("Внимание", "Таблица пуста, нечего сохранять")
        return
    file_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")],
        title="Сохранить таблицу как..."
    )
    if file_path:
        df_global.to_excel(file_path, index=False)
        messagebox.showinfo("Готово", f"Таблица сохранена: {file_path}")

def save_log_file():
    if not LOG_FILE.exists():
        messagebox.showwarning("Внимание", "Лог еще пуст")
        return
    file_path = filedialog.asksaveasfilename(
        defaultextension=".txt",
        filetypes=[("Text files", "*.txt")],
        title="Сохранить лог как..."
    )
    if file_path:
        with open(LOG_FILE, 'r', encoding='utf-8') as src, open(file_path, 'w', encoding='utf-8') as dst:
            dst.write(src.read())
        messagebox.showinfo("Готово", f"Лог сохранен: {file_path}")

# ---------------------- GUI ---------------------- #
root = tk.Tk()
root.title("ArcheAge PlayersComparer v1.3.0")
root.resizable(False, False)

frame = tk.Frame(root)
frame.pack(padx=10, pady=10)

tk.Label(frame, text="До:").grid(row=0, column=0)
before_listbox = tk.Listbox(frame, width=60)
before_listbox.grid(row=1, column=0)
tk.Button(frame, text="Добавить файлы", command=lambda: add_files(before_listbox)).grid(row=2, column=0)
tk.Button(frame, text="Удалить выбранное", command=lambda: remove_selected(before_listbox)).grid(row=3, column=0)

tk.Label(frame, text="После:").grid(row=0, column=1)
after_listbox = tk.Listbox(frame, width=60)
after_listbox.grid(row=1, column=1)
tk.Button(frame, text="Добавить файлы", command=lambda: add_files(after_listbox)).grid(row=2, column=1)
tk.Button(frame, text="Удалить выбранное", command=lambda: remove_selected(after_listbox)).grid(row=3, column=1)

tk.Button(root, text="Старт", command=start_processing).pack(pady=5)
tk.Button(root, text="Стоп", command=stop_processing).pack(pady=5)
tk.Button(root, text="Выгрузить в Excel", command=save_table_excel).pack(pady=5)
tk.Button(root, text="Сохранить лог", command=save_log_file).pack(pady=5)
tk.Button(root, text="Выход", command=root.quit).pack(pady=5)

progress_var = tk.IntVar()
progress = ttk.Progressbar(root, orient="horizontal", length=400, mode="determinate", variable=progress_var)
progress.pack(pady=5)

# Поле для таблицы с прокруткой
table_frame = tk.Frame(root)
table_frame.pack(pady=5)

scrollbar = tk.Scrollbar(table_frame)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

table_text = tk.Text(table_frame, width=100, height=15, yscrollcommand=scrollbar.set)
table_text.pack(side=tk.LEFT)

scrollbar.config(command=table_text.yview)

pd.set_option("display.max_rows", None)
pd.set_option("display.max_columns", None)

root.mainloop()
