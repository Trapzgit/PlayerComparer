import time
from pathlib import Path
import tkinter as tk
from tkinter import messagebox
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from PIL import Image
import json

# Папка, где ArcheAge сохраняет скриншоты
WATCH_DIR = Path(r"C:\ArcheAge\Documents\ScreenShots")
# Папка внутри проекта для сохраняемых скриншотов (before/after)
SAVE_BASE = Path("screenshots")

# конфиг (подгружаем те же координаты, что и main)
CONFIG_FILE = Path("crop_config.json")
DEFAULT_CROP = (1170, 15, 1890, 200)
TARGET_WIDTH = 1500
TARGET_HEIGHT = 500

def load_crop_region():
    if CONFIG_FILE.exists():
        try:
            data = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
            if all(k in data for k in ("x1","y1","x2","y2")):
                return (int(data["x1"]), int(data["y1"]), int(data["x2"]), int(data["y2"]))
        except Exception:
            pass
    return DEFAULT_CROP

CROP_REGION = load_crop_region()

observer = None
_control_window = None

def crop_and_save_scaled(src_path: Path, dest_path: Path):
    """Обрезает по CROP_REGION и сохраняет масштабированное изображение в dest_path."""
    try:
        with Image.open(src_path) as im:
            im = im.convert("RGB")
            cropped = im.crop(CROP_REGION)
            resized = cropped.resize((TARGET_WIDTH, TARGET_HEIGHT), Image.LANCZOS)
            dest_path.parent.mkdir(parents=True, exist_ok=True)
            resized.save(dest_path)
            print(f"[OK] {src_path} -> {dest_path}")
    except Exception as e:
        print(f"[ERR] crop_and_save_scaled {src_path}: {e}")

class ScreenshotHandler(FileSystemEventHandler):
    def __init__(self, category, listbox):
        self.category = category
        self.listbox = listbox
        (SAVE_BASE / category).mkdir(parents=True, exist_ok=True)

    def on_created(self, event):
        if event.is_directory:
            return
        # поддерживаем jpg/png/jpeg
        if event.src_path.lower().endswith((".jpg", ".jpeg", ".png")):
            src = Path(event.src_path)
            # ждём записи файла
            time.sleep(0.5)
            dest = (SAVE_BASE / self.category / (src.stem + "_crop.png"))
            crop_and_save_scaled(src, dest)
            # добавляем в listbox (строка)
            try:
                self.listbox.insert(tk.END, str(dest))
            except Exception:
                pass

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
        tk.Label(_control_window, text=f"Новые скриншоты F9 будут попадать в {category.upper()}\nНажмите 'Стоп', чтобы выйти").pack(pady=10)
        tk.Button(_control_window, text="Стоп", command=lambda: stop_screenshot_mode(_control_window)).pack(pady=5)

    choose_category()

def stop_screenshot_mode(win=None):
    global observer, _control_window
    if observer:
        observer.stop()
        observer.join()
        observer = None
    if win:
        try:
            win.destroy()
        except Exception:
            pass
    if _control_window:
        try:
            _control_window.destroy()
        except Exception:
            pass
        _control_window = None
    try:
        messagebox.showinfo("Режим скриншота", "Остановлен")
    except Exception:
        pass
