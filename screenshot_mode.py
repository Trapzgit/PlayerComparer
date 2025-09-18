import time
import shutil
from pathlib import Path
import tkinter as tk
from tkinter import messagebox
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

WATCH_DIR = Path(r"C:\Users\Trpz\Pictures\Screenshots")  # где Windows хранит скриншоты
SAVE_DIR = Path("screenshots")  # куда сохраняем внутри программы

stop_flag = False
observer = None
_control_window = None


class ScreenshotHandler(FileSystemEventHandler):
    def __init__(self, category, listbox):
        self.category = category
        self.listbox = listbox
        (SAVE_DIR / category).mkdir(parents=True, exist_ok=True)

    def on_created(self, event):
        if event.is_directory:
            return
        if event.src_path.lower().endswith(".png"):
            src = Path(event.src_path)
            # ждём, пока файл полностью запишется
            time.sleep(0.5)
            dest = SAVE_DIR / self.category / src.name
            shutil.copy2(src, dest)
            self.listbox.insert(tk.END, str(dest))
            print(f"[OK] Перемещён {src} → {dest}")


def start_screenshot_mode(root, before_listbox, after_listbox):
    global stop_flag
    stop_flag = False

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
        tk.Label(_control_window, text=f"Новые скриншоты будут попадать в {category.upper()}\n"
                                       f"Нажмите 'Стоп', чтобы выйти").pack(pady=10)
        tk.Button(_control_window, text="Стоп", command=lambda: stop_screenshot_mode(_control_window)).pack(pady=5)

    choose_category()


def stop_screenshot_mode(win=None):
    global observer, stop_flag, _control_window
    stop_flag = True
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
