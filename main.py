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
import json

__version__ = "v2.0.0"

# ---------------------- Настройки ---------------------- #
DD_list = ['Lnl', 'Nebovesna', 'Runbott', 'Trpvz', 'Pesdaliss', 'Oguricap', 'Revanx',
           'Luthicx', 'Olven', 'Скуфнатраппере', 'Владосхристос', 'Zshturmovik', 'Арбузбек',
           'Rabbittt', 'Срал', 'Sheeeshh', 'Shzs', 'Невсегдасвятой','Хорошиймальчик',
           'Гламурныйахэгао', 'Вожакстаданегрилл', 'Pesdexely', 'Hikikomorri','Secretquest',
           'Iletyouhide', 'Сомнительнополезен', 'Dimonishzv', 'Слабейшеебедствие', 'Особонеопасен',
           'Пиуупиу', 'Ксюшасуперкрутая', 'Ксюшаоченькрутая', 'Стараятварь', 'Знатокпоражений',
           'Слабейшееоружее', 'Ssoptymysvprame', 'Pugl', 'Zyhzv', 'Chuccky', 'Преломилось',
           'Незнающийпобеды', 'Россияскоростная', 'Skripkazv','Бесполезный','Есычь','Chillingtouch',
           'Данотеламедиа', 'Maestrozv', 'Ineedhelp', 'Kazakhx', 'Lasttry', 'Крольчашь', 'Корольлоутаба',
           'Fieakinexcellent', 'Starbust']

stop_flag = False
df_global = pd.DataFrame()
LOG_FILE = Path(sys.executable).parent / "logs.txt"
CONFIG_FILE = Path("crop_config.json")
DD_FILE = Path("dd_list.json")

def load_dd_list():
    if DD_FILE.exists():
        with open(DD_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    else:
        # создаём файл с дефолтным списком
        default_list = ['Lnl', 'Nebovesna', 'Runbott', 'Trpvz', 'Pesdaliss', 'Oguricap', 'Revanx',
                        'Luthicx', 'Olven', 'Скуфнатраппере', 'Владосхристос', 'Zshturmovik', 'Арбузбек',
                        'Rabbittt', 'Срал', 'Sheeeshh', 'Shzs', 'Невсегдасвятой','Хорошиймальчик',
                        'Гламурныйахэгао', 'Вожакстаданегрилл', 'Pesdexely', 'Hikikomorri','Secretquest',
                        'Iletyouhide', 'Сомнительнополезен', 'Dimonishzv', 'Слабейшеебедствие', 'Особонеопасен',
                        'Пиуупиу', 'Ксюшасуперкрутая', 'Ксюшаоченькрутая', 'Стараятварь', 'Знатокпоражений',
                        'Слабейшееоружее', 'Ssoptymysvprame', 'Pugl', 'Zyhzv', 'Chuccky', 'Преломилось',
                        'Незнающийпобеды', 'Россияскоростная', 'Skripkazv','Бесполезный','Есычь','Chillingtouch',
                        'Данотеламедиа', 'Maestrozv', 'Ineedhelp', 'Kazakhx', 'Lasttry', 'Крольчашь', 'Корольлоутаба',
                        'Fieakinexcellent', 'Starbust']

        with open(DD_FILE, "w", encoding="utf-8") as f:
            json.dump(default_list, f, ensure_ascii=False, indent=2)
        return default_list

def save_dd_list(dd_list):
    with open(DD_FILE, "w", encoding="utf-8") as f:
        json.dump(dd_list, f, ensure_ascii=False, indent=2)

# Загружаем при запуске
DD_list = load_dd_list()

# --- Загружаем или создаём настройки обрезки ---
def load_crop_region():
    if CONFIG_FILE.exists():
        try:
            data = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
            if all(k in data for k in ["x1", "y1", "x2", "y2"]):
                return (data["x1"], data["y1"], data["x2"], data["y2"])
        except Exception:
            pass
    # Значения по умолчанию
    default = (1170, 15, 1890, 200)
    save_crop_region(*default)
    return default

def save_crop_region(x1, y1, x2, y2):
    data = {"x1": x1, "y1": y1, "x2": x2, "y2": y2}
    CONFIG_FILE.write_text(json.dumps(data, indent=4, ensure_ascii=False), encoding="utf-8")

# --- Загружаем актуальные координаты ---
CROP_REGION = load_crop_region()

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

# ---------------------- Обрезка изображений ---------------------- #
CROP_REGION = (1170, 15, 1890, 200)  # область для обрезки

def crop_image_to_region(image_path, save_over=True):
    """Обрезает изображение по заданной области, если оно ещё не обрезано."""
    try:
        img = Image.open(image_path)
        w, h = img.size

        if w <= (CROP_REGION[2] - CROP_REGION[0]) and h <= (CROP_REGION[3] - CROP_REGION[1]):
            return image_path  # уже обрезано

        cropped = img.crop(CROP_REGION)
        if save_over:
            cropped.save(image_path)
            print(f"[OK] Обрезано: {image_path}")
            return image_path
        else:
            new_path = Path(image_path).with_stem(Path(image_path).stem + "_cropped")
            cropped.save(new_path)
            print(f"[OK] Создан: {new_path}")
            return str(new_path)
    except Exception as e:
        print(f"[ERR] Ошибка обрезки {image_path}: {e}")
        return image_path

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
        crop_image_to_region(file_path)  # 👈 заменили resize_if_needed
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
    df_global = pd.DataFrame()
    clear_log()
    before_files = before_listbox.get(0, tk.END)
    after_files = after_listbox.get(0, tk.END)
    if not before_files or not after_files:
        messagebox.showwarning("Ошибка", "Добавьте файлы до и после")
        return
    progress_var.set(0)
    table_text.delete(1.0, tk.END)

    btn_start.config(state="disabled")

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
            messagebox.showinfo("Готово", "Обработка завершена")
        finally:
            btn_start.config(state="normal")
            progress_var.set(0)
    threading.Thread(target=worker).start()

def stop_processing():
    global stop_flag
    stop_flag = True

# ---------------------- Работа с файлами ---------------------- #

def add_nick_to_ddlist():
    def save_nick():
        nick = entry.get().strip()
        if not nick:
            messagebox.showwarning("Ошибка", "Введите ник")
            return
        if nick in DD_list:
            messagebox.showinfo("Инфо", "Такой ник уже есть в базе")
            return
        DD_list.append(nick)
        save_dd_list(DD_list)
        messagebox.showinfo("Готово", f"Ник '{nick}' добавлен в базу!")
        win.destroy()

    win = tk.Toplevel(root)
    win.title("Добавить ник")
    tk.Label(win, text="Введите ник для добавления:").pack(pady=10)
    entry = tk.Entry(win, width=40)
    entry.pack(padx=10)
    tk.Button(win, text="Сохранить", command=save_nick).pack(pady=10)

def open_crop_settings():
    """Открывает окно для изменения координат CROP_REGION."""
    def save_and_close():
        try:
            x1 = int(entry_x1.get())
            y1 = int(entry_y1.get())
            x2 = int(entry_x2.get())
            y2 = int(entry_y2.get())
            if x2 <= x1 or y2 <= y1:
                messagebox.showerror("Ошибка", "x2/y2 должны быть больше x1/y1!")
                return
            save_crop_region(x1, y1, x2, y2)
            global CROP_REGION
            CROP_REGION = (x1, y1, x2, y2)
            messagebox.showinfo("Сохранено", f"Координаты обновлены:\n{x1}, {y1}, {x2}, {y2}")
            win.destroy()
        except ValueError:
            messagebox.showerror("Ошибка", "Введите только числа!")

    win = tk.Toplevel()
    win.title("Настройки области обрезки")
    win.geometry("300x220")
    win.resizable(False, False)

    tk.Label(win, text="Введите координаты обрезки:").pack(pady=5)

    frm = tk.Frame(win)
    frm.pack(pady=10)

    tk.Label(frm, text="x1:").grid(row=0, column=0)
    entry_x1 = tk.Entry(frm, width=8)
    entry_x1.grid(row=0, column=1, padx=5)

    tk.Label(frm, text="y1:").grid(row=1, column=0)
    entry_y1 = tk.Entry(frm, width=8)
    entry_y1.grid(row=1, column=1, padx=5)

    tk.Label(frm, text="x2:").grid(row=2, column=0)
    entry_x2 = tk.Entry(frm, width=8)
    entry_x2.grid(row=2, column=1, padx=5)

    tk.Label(frm, text="y2:").grid(row=3, column=0)
    entry_y2 = tk.Entry(frm, width=8)
    entry_y2.grid(row=3, column=1, padx=5)

    # Подставляем текущие значения
    entry_x1.insert(0, CROP_REGION[0])
    entry_y1.insert(0, CROP_REGION[1])
    entry_x2.insert(0, CROP_REGION[2])
    entry_y2.insert(0, CROP_REGION[3])

    tk.Button(win, text="Сохранить", command=save_and_close).pack(pady=10)
    tk.Button(win, text="Отмена", command=win.destroy).pack()

def add_files(listbox):
    files = filedialog.askopenfilenames(
        filetypes=[("Изображения", "*.png;*.jpg;*.jpeg"), ("Все файлы", "*.*")]
    )
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
            matches = difflib.get_close_matches(bnick, after_nick_list, n=5, cutoff=0.6) # % совпадения с ником
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

# ---------------------- GUI (Тёмная тема) ---------------------- #
BG_COLOR = "#2E2E2E"        # фон окна
FRAME_BG = "#383838"         # фон фреймов
BTN_COLOR = "#4A90E2"        # кнопки
BTN_HOVER = "#357ABD"        # при наведении
BTN_TEXT = "#FFFFFF"
LBL_COLOR = "#E0E0E0"
TEXT_BG = "#1E1E1E"
TEXT_FG = "#E0E0E0"
PROG_COLOR = "#4A90E2"

root = tk.Tk()
root.title(f"ArcheAge PlayersComparer {__version__}")
root.minsize(900, 600)
root.configure(bg=BG_COLOR)

# --------- Tooltips ----------
class ToolTip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tipwindow = None
        widget.bind("<Enter>", self.show)
        widget.bind("<Leave>", self.hide)

    def show(self, event=None):
        if self.tipwindow or not self.text:
            return
        x, y, cx, cy = self.widget.bbox("insert") or (0,0,0,0)
        x = x + self.widget.winfo_rootx() + 25
        y = y + self.widget.winfo_rooty() + 25
        self.tipwindow = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.configure(bg="#333333")
        label = tk.Label(tw, text=self.text, bg="#333333", fg="#FFFFFF",
                         relief='solid', borderwidth=1, font=("Arial", 9))
        label.pack()
        tw.wm_geometry(f"+{x}+{y}")

    def hide(self, event=None):
        if self.tipwindow:
            self.tipwindow.destroy()
            self.tipwindow = None

# --------- Hover ----------
def on_enter(e):
    e.widget['bg'] = BTN_HOVER
def on_leave(e):
    e.widget['bg'] = BTN_COLOR

# --------- Frames ----------
file_frame = tk.Frame(root, bg=FRAME_BG, bd=2, relief="groove")
file_frame.pack(padx=10, pady=10, fill=tk.X)

# До/После списки
left_frame = tk.Frame(file_frame, bg=FRAME_BG)
left_frame.grid(row=0, column=0, padx=10, pady=5)
tk.Label(left_frame, text="До:", bg=FRAME_BG, fg=LBL_COLOR).pack(anchor="w")
before_listbox = tk.Listbox(left_frame, width=50, height=6, bg=TEXT_BG, fg=TEXT_FG,
                            selectbackground="#555555")
before_listbox.pack(pady=5)
btn_before_add = tk.Button(left_frame, text="Добавить файлы", bg=BTN_COLOR, fg=BTN_TEXT,
                           command=lambda: add_files(before_listbox))
btn_before_add.pack(pady=2, fill=tk.X)
btn_before_add.bind("<Enter>", on_enter)
btn_before_add.bind("<Leave>", on_leave)
ToolTip(btn_before_add, "Добавить файлы для сравнения 'До'")

btn_before_remove = tk.Button(left_frame, text="Удалить выбранное", bg=BTN_COLOR, fg=BTN_TEXT,
                              command=lambda: remove_selected(before_listbox))
btn_before_remove.pack(pady=2, fill=tk.X)
btn_before_remove.bind("<Enter>", on_enter)
btn_before_remove.bind("<Leave>", on_leave)
ToolTip(btn_before_remove, "Удалить выбранные файлы из списка 'До'")

right_frame = tk.Frame(file_frame, bg=FRAME_BG)
right_frame.grid(row=0, column=1, padx=10, pady=5)
tk.Label(right_frame, text="После:", bg=FRAME_BG, fg=LBL_COLOR).pack(anchor="w")
after_listbox = tk.Listbox(right_frame, width=50, height=6, bg=TEXT_BG, fg=TEXT_FG,
                           selectbackground="#555555")
after_listbox.pack(pady=5)
btn_after_add = tk.Button(right_frame, text="Добавить файлы", bg=BTN_COLOR, fg=BTN_TEXT,
                          command=lambda: add_files(after_listbox))
btn_after_add.pack(pady=2, fill=tk.X)
btn_after_add.bind("<Enter>", on_enter)
btn_after_add.bind("<Leave>", on_leave)
ToolTip(btn_after_add, "Добавить файлы для сравнения 'После'")

btn_after_remove = tk.Button(right_frame, text="Удалить выбранное", bg=BTN_COLOR, fg=BTN_TEXT,
                             command=lambda: remove_selected(after_listbox))
btn_after_remove.pack(pady=2, fill=tk.X)
btn_after_remove.bind("<Enter>", on_enter)
btn_after_remove.bind("<Leave>", on_leave)
ToolTip(btn_after_remove, "Удалить выбранные файлы из списка 'После'")

# ---------------------- Кнопки действий ---------------------- #
action_frame = tk.Frame(root, bg=BG_COLOR)
action_frame.pack(pady=10)

btn_start = tk.Button(action_frame, text="▶", font=("Arial", 14, "bold"), bg=BTN_COLOR, fg=BTN_TEXT,
                      width=4, height=1, command=start_processing)
btn_start.grid(row=0, column=0, padx=5)
btn_start.bind("<Enter>", on_enter)
btn_start.bind("<Leave>", on_leave)
ToolTip(btn_start, "Запустить обработку файлов")

btn_stop = tk.Button(action_frame, text="■", font=("Arial", 14, "bold"), bg=BTN_COLOR, fg=BTN_TEXT,
                     width=4, height=1, command=stop_processing)
btn_stop.grid(row=0, column=1, padx=5)
btn_stop.bind("<Enter>", on_enter)
btn_stop.bind("<Leave>", on_leave)
ToolTip(btn_stop, "Остановить текущую обработку")

btn_excel = tk.Button(action_frame, text="Excel", bg=BTN_COLOR, fg=BTN_TEXT, width=10,
                      command=save_table_excel)
btn_excel.grid(row=0, column=2, padx=5)
btn_excel.bind("<Enter>", on_enter)
btn_excel.bind("<Leave>", on_leave)
ToolTip(btn_excel, "Сохранить таблицу в Excel")

btn_log = tk.Button(action_frame, text="Лог", bg=BTN_COLOR, fg=BTN_TEXT, width=10,
                    command=save_log_file)
btn_log.grid(row=0, column=3, padx=5)
btn_log.bind("<Enter>", on_enter)
btn_log.bind("<Leave>", on_leave)
ToolTip(btn_log, "Сохранить лог обработки")

btn_from_log = tk.Button(action_frame, text="Из лога", bg=BTN_COLOR, fg=BTN_TEXT, width=10,
                         command=create_table_from_log)
btn_from_log.grid(row=0, column=4, padx=5)
btn_from_log.bind("<Enter>", on_enter)
btn_from_log.bind("<Leave>", on_leave)
ToolTip(btn_from_log, "Создать таблицу из файла лога")

btn_screenshot = tk.Button(action_frame, text="📷", bg=BTN_COLOR, fg=BTN_TEXT, width=4,
                           command=run_screenshot_mode)
btn_screenshot.grid(row=0, column=5, padx=5)
btn_screenshot.bind("<Enter>", on_enter)
btn_screenshot.bind("<Leave>", on_leave)
ToolTip(btn_screenshot, "Режим скриншота")

btn_crop = tk.Button(action_frame, text="✂", bg=BTN_COLOR, fg=BTN_TEXT, width=4,
                     command=open_crop_settings)
btn_crop.grid(row=0, column=6, padx=5)
btn_crop.bind("<Enter>", on_enter)
btn_crop.bind("<Leave>", on_leave)
ToolTip(btn_crop, "Настройки области обрезки")

btn_add_nick = tk.Button(action_frame, text="+", bg=BTN_COLOR, fg=BTN_TEXT, width=4,
                         command=add_nick_to_ddlist)
btn_add_nick.grid(row=0, column=7, padx=5)
btn_add_nick.bind("<Enter>", on_enter)
btn_add_nick.bind("<Leave>", on_leave)
ToolTip(btn_add_nick, "Добавить ник в базу DD_list")

# --------- Прогресс ---------
progress_var = tk.IntVar()
style = ttk.Style()
style.theme_use('clam')
style.configure("blue.Horizontal.TProgressbar", troughcolor="#555555", background=PROG_COLOR)
progress = ttk.Progressbar(root, style="blue.Horizontal.TProgressbar", orient="horizontal",
                           length=600, mode="determinate", variable=progress_var)
progress.pack(pady=5)

# --------- Таблица ---------
table_frame = tk.Frame(root, bg=BG_COLOR)
table_frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
scrollbar = tk.Scrollbar(table_frame)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
table_text = tk.Text(table_frame, width=120, height=15, yscrollcommand=scrollbar.set,
                     bg=TEXT_BG, fg=TEXT_FG, font=("Consolas", 10))
table_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
scrollbar.config(command=table_text.yview)

# Закрытие приложения
def on_close():
    try:
        screenshot_mode.stop_screenshot_mode()
    except Exception as e:
        print("Ошибка при остановке screenshot_mode:", e)
    stop_processing()
    root.destroy()

root.protocol("WM_DELETE_WINDOW", on_close)
root.mainloop()
