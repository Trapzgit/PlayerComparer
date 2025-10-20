import time
from pathlib import Path
import tkinter as tk
from tkinter import messagebox
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from PIL import Image

# Папка, где ArcheAge сохраняет скриншоты
WATCH_DIR = Path(r"C:\ArcheAge\Documents\ScreenShots")
SAVE_DIR = Path(r"C:\ArcheAge\Documents\ScreenShots")

observer = None
_control_window = None

# --- общая функция для обрезки изображений (та же, что в main.py) ---
CROP_REGION = (1170, 15, 1890, 200)

def crop_image_to_region(image_path, save_path):
    """Обрезает изображение по заданной области, если ещё не обрезано."""
    try:
        img = Image.open(image_path)
        w, h = img.size
        if w <= (CROP_REGION[2] - CROP_REGION[0]) and h <= (CROP_REGION[3] - CROP_REGION[1]):
            img.save(save_path)
            return
        cropped = img.crop(CROP_REGION)
        cropped.save(save_path)
        print(f"[OK] Обрезан {image_path} → {save_path}")
    except Exception as e:
        print(f"[ERR] Ошибка обработки {image_path}: {e}")

class ScreenshotHandler(FileSystemEventHandler):
    def __init__(self, category, listbox):
        self.category = category
        self.listbox = listbox
        (SAVE_DIR / category).mkdir(parents=True, exist_ok=True)

    def on_created(self, event):
        if event.is_directory:
            return
        if event.src_path.lower().endswith((".jpg", ".png", ".jpeg")):
            src = Path(event.src_path)
            time.sleep(0.5)
            dest = SAVE_DIR / self.category / (src.stem + "_crop.png")
            crop_image_to_region(src, dest)
            self.listbox.insert(tk.END, str(dest))

def start_screenshot_mode(root, before_listbox, after_listbox):
    global observer, _control_window
    if observer is not None:
        messagebox.showwarning("Режим скриншота", "Уже запущен! Сначала остановите текущий режим.")
        return

    def choose_category():
        win = tk.Toplevel(root)
        win.title("Выбор категории")
        tk.Label(win, text="Куда сохранять новые скриншоты?").pack(pady=10)

        def start(category, listbox):
            nonlocal win
            win.destroy()
            run_observer(category, listbox)

        tk.Button(win, text="Before", command=lambda: start("before", before_listbox)).pack(pady=5)
        tk.Button(win, text="After", command=lambda: start("after", after_listbox)).pack(pady=5)
        tk.Button(win, text="Отмена", command=win.destroy).pack(pady=5)

    def run_observer(category, listbox):
        global observer, _control_window
        handler = ScreenshotHandler(category, listbox)
        observer = Observer()
        observer.schedule(handler, str(WATCH_DIR), recursive=False)
        observer.start()

        _control_window = tk.Toplevel(root)
        _control_window.title("Режим скриншота")
        tk.Label(
            _control_window,
            text=f"Новые скриншоты F9 будут попадать в {category.upper()}\n"
                 f"Нажмите 'Стоп', чтобы выйти"
        ).pack(pady=10)
        tk.Button(_control_window, text="Стоп", command=lambda: stop_screenshot_mode(_control_window)).pack(pady=5)

    choose_category()

def stop_screenshot_mode(win=None):
    global observer, _control_window
    if observer:
        observer.stop()
        observer.join()
        observer = None
    if win:
        win.destroy()
    if _control_window:
        try:
            _control_window.destroy()
        except:
            pass
        _control_window = None
    messagebox.showinfo("Режим скриншота", "Остановлен")
