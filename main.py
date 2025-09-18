import threading
import screenshot_mode
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
import easyocr
import re
import difflib
from pathlib import Path
from PIL import Image
import string
import sys

__version__ = "v1.7.0"

# ---------------------- Настройки ---------------------- #
DD_list = ['Lnl', 'Nebovesna', 'Runbott', 'Trpvz', 'Pesdaliss', 'Oguricap', 'Revanx',
           'Luthicx', 'Olven', 'Скуфнатраппере', 'Владосхристос', 'Zshturmovik', 'Арбузбек',
           'Rabbittt', 'Срал', 'Sheeeshh', 'Shzs', 'Невсегдасвятой','Хорошиймальчик',
           'Гламурныйахэгао', 'Вожакстаданегрилл', 'Pesdexely', 'Hikikomorri','Secretquest',
           'Iletyouhide', 'Сомнительнополезен', 'Dimonishzv', 'Слабейшеебедствие', 'Особонеопасен',
           'Пиуупиу', 'Ксюшасуперкрутая', 'Ксюшаоченькрутая', 'Стараятварь', 'Знатокпоражений',
           'Слабейшееоружее', 'Ssoptymysvprame', 'Pugl', 'Zyhzv', 'Chuccky', 'Преломилось',
           'Незнающийпобеды']

stop_flag = False
df_global = pd.DataFrame()
LOG_FILE = Path(sys.executable).parent / "logs.txt"

# ---------------------- Помощники ---------------------- #
def safe_filename(name):
    valid_chars = string.ascii_letters + string.digits + "_"
    return ''.join(c if c in valid_chars else "_" for c in name)

def correct_nick(text_block):
    cleaned_text = re.sub(r'[^\w\s]', ' ', text_block)
    words = cleaned_text.split()
    corrected_nick = None
    for word in words:
        matches = difflib.get_close_matches(word, DD_list, n=1, cutoff=0.6)
        if matches:
            corrected_nick = matches[0]
            break
    if corrected_nick:
        text_block = re.sub(re.escape(word), '', text_block, count=1)
        text_block = corrected_nick + ' ' + text_block.strip()
    return text_block

def finalize_block(block_text):
    cleaned_text = re.sub(r'[^A-Za-zА-Яа-я0-9\s]', ' ', block_text)
    words = re.findall(r'\b[A-Za-zА-Яа-я0-9]{3,20}\b', cleaned_text)
    best_match = None
    best_ratio = 0
    for word in words:
        matches = difflib.get_close_matches(word, DD_list, n=1, cutoff=0.6)
        if matches:
            ratio = difflib.SequenceMatcher(None, word, matches[0]).ratio()
            if ratio > best_ratio:
                best_ratio = ratio
                best_match = matches[0]
    nick = best_match if best_match else "Unknown"
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
            if len(first_two) == 0 and len(num) == 5:
                if i + 1 < len(numbers) and len(numbers[i + 1]) == 1:
                    num = num + numbers[i + 1]
                    i += 1
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

def resize_if_needed(image_path, min_width=1500, min_height=500):
    with Image.open(image_path) as img:
        width, height = img.size
        if width != min_width or height != min_height:
            img = img.resize((min_width, min_height), Image.LANCZOS)
            img.save(image_path)
            print(f"Resized {image_path} to {min_width}x{min_height}")

# ---------------------- Лог ---------------------- #
def clear_log():
    if LOG_FILE.exists():
        LOG_FILE.unlink()
    LOG_FILE.touch()

# ---------------------- Обработка файлов ---------------------- #
def process_files(file_list, folder_name, current, total, progress_var, reader):
    global stop_flag
    results = {}
    for file_str in file_list:
        if stop_flag:
            break
        file_path = Path(file_str)
        new_name = safe_filename(file_path.stem) + file_path.suffix
        safe_path = file_path.parent / new_name
        if safe_path != file_path:
            try:
                file_path.rename(safe_path)
            except Exception as e:
                print("Rename error:", e)
        file_path = safe_path
        resize_if_needed(file_path)
        text_result = reader.readtext(str(file_path), detail=0, paragraph=True)
        full_text = " ".join(text_result)
        corrected = correct_nick(full_text)
        cleaned = finalize_block(corrected)
        write_log(file_path.name, folder_name, full_text, cleaned)
        results[file_path.name] = cleaned
        current[0] += 1
        progress_var.set(int(current[0] / total * 100))
    return results

# ---------------------- Поток обработки ---------------------- #
def start_processing():
    global stop_flag, df_global
    stop_flag = False
    df_global = pd.DataFrame()  # сброс перед запуском
    clear_log()
    before_files = before_listbox.get(0, tk.END)
    after_files = after_listbox.get(0, tk.END)
    if not before_files or not after_files:
        messagebox.showwarning("Ошибка", "Добавьте файлы до и после")
        return
    progress_var.set(0)
    table_text.delete(1.0, tk.END)
    start_button.config(state="disabled")  # блокируем кнопку

    def worker():
        global df_global
        try:
            total = len(before_files) + len(after_files)
            current = [0]
            reader = easyocr.Reader(['en', 'ru'], gpu=True)
            before_data = process_files(before_files, 'before', current, total, progress_var, reader)
            after_data = process_files(after_files, 'after', current, total, progress_var, reader)
            table = []
            for key in before_data:
                b = before_data[key].split()
                nick_b = b[0]
                after_nicks = [v.split()[0] for v in after_data.values()]
                best_match = difflib.get_close_matches(nick_b, after_nicks, n=1, cutoff=0.8)
                if best_match:
                    for a_val in after_data.values():
                        if a_val.startswith(best_match[0]):
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
        finally:
            start_button.config(state="normal")  # возвращаем кнопку
            progress_var.set(0)  # сброс прогрессбара
    threading.Thread(target=worker).start()

def stop_processing():
    global stop_flag
    stop_flag = True

# ---------------------- Работа с файлами ---------------------- #
def add_files(listbox):
    files = filedialog.askopenfilenames(filetypes=[("PNG Files","*.png")])
    for f in files:
        listbox.insert(tk.END, str(Path(f)))

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

# ---------------------- Создание таблицы из файла лог ---------------------- #
def create_table_from_log():
    global df_global

    log_file_path = filedialog.askopenfilename(
        title="Выберите файл лог",
        filetypes=[("Text files", "*.txt")]
    )
    if not log_file_path:
        return

    with open(log_file_path, 'r', encoding='utf-8') as f:
        content = f.read()

    blocks = content.split('--- ')
    before_entries = []
    after_entries = []

    # Парсим блоки
    for block in blocks:
        block = block.strip()
        if not block:
            continue
        if 'После обработки:' not in block:
            continue

        try:
            after_text = block.split('После обработки:')[1].strip()
        except IndexError:
            continue

        parts = after_text.split()
        if len(parts) < 4:
            continue

        nick_raw = parts[0]
        cls = parts[2]
        try:
            honor = int(parts[3])
            kills = int(parts[4])
        except (IndexError, ValueError):
            honor = 0
            kills = 0

        # Корректировка ника через DD_list
        matches = difflib.get_close_matches(nick_raw, DD_list, n=1, cutoff=0.6)
        nick = matches[0] if matches else nick_raw

        entry = {'nick': nick, 'class': cls, 'honor': honor, 'kills': kills}

        if block.startswith('before/'):
            before_entries.append(entry)
        elif block.startswith('after/'):
            after_entries.append(entry)

    if not before_entries or not after_entries:
        messagebox.showwarning("Внимание", "Не найдено данных в лог")
        return

    rows = []
    matched_after_idx = set()
    after_nick_list = [a['nick'] for a in after_entries]

    # Сравниваем before и after
    for b in before_entries:
        bnick = b['nick']
        bcls = b['class']
        bhonor = b['honor']
        bkills = b['kills']

        candidate = None
        candidate_idx = None

        # 1) Точное совпадение
        for idx, a in enumerate(after_entries):
            if idx in matched_after_idx:
                continue
            if a['nick'] == bnick:
                candidate = a
                candidate_idx = idx
                break

        # 2) Нечеткое совпадение через DD_list
        if candidate is None:
            matches = difflib.get_close_matches(bnick, after_nick_list, n=5, cutoff=0.6)
            for mname in matches:
                for idx, a in enumerate(after_entries):
                    if idx in matched_after_idx:
                        continue
                    if a['nick'] == mname:
                        candidate = a
                        candidate_idx = idx
                        break
                if candidate:
                    break

        # 3) Если пары нет, пропускаем
        if candidate is None:
            continue

        matched_after_idx.add(candidate_idx)

        dhonor = (candidate['honor'] or 0) - (bhonor or 0)
        dkills = (candidate['kills'] or 0) - (bkills or 0)

        use_class = bcls if bcls else candidate.get('class', '')

        rows.append([b['nick'], use_class, dhonor, dkills])

    if not rows:
        messagebox.showwarning("Внимание", "Не удалось найти совпадения для игроков")
        return

    df_global = pd.DataFrame(rows, columns=["Ник", "Класс", "Δ Хонор", "Δ Киллы"])
    df_global.sort_values("Δ Киллы", ascending=False, inplace=True)

    table_text.delete(1.0, tk.END)
    table_text.insert(tk.END, df_global.to_string(index=False))

    messagebox.showinfo("Готово", f"Таблица создана. Найдено {len(rows)} совпадений.")

# ---------------------- screenshot_mod ---------------------- #
def run_screenshot_mode():
    threading.Thread(
        target=lambda: screenshot_mode.start_screenshot_mode(root, before_listbox, after_listbox),
        daemon=True
    ).start()

# ---------------------- GUI ---------------------- #
root = tk.Tk()
root.title(f"ArcheAge PlayersComparer "+__version__)
root.minsize(800, 600)
root.maxsize(1920, 1080)
root.resizable(True, True)

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

start_button = tk.Button(root, text="Старт", command=start_processing)
start_button.pack(pady=5)
tk.Button(root, text="Стоп", command=stop_processing).pack(pady=5)
tk.Button(root, text="Выгрузить в Excel", command=save_table_excel).pack(pady=5)
tk.Button(root, text="Сохранить лог", command=save_log_file).pack(pady=5)
tk.Button(root, text="Создать таблицу из лога", command=create_table_from_log).pack(pady=5)
# кнопка выхода — теперь вызывает on_close
def on_close():
    # останавливаем observer (если запущен)
    try:
        screenshot_mode.stop_screenshot_mode()
    except Exception as e:
        print("Ошибка при остановке screenshot_mode:", e)
    # останавливаем процессы обработки
    try:
        stop_processing()
    except Exception:
        pass
    root.destroy()

#tk.Button(root, text="Выход", command=on_close).pack(pady=5)
tk.Button(root, text="Режим скриншот", command=run_screenshot_mode).pack(pady=5)

# также привязываем крестик окна к on_close
root.protocol("WM_DELETE_WINDOW", on_close)

progress_var = tk.IntVar()
progress = ttk.Progressbar(root, orient="horizontal", length=400, mode="determinate", variable=progress_var)
progress.pack(pady=5)

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
