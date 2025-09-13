import tkinter as tk
from tkinter import filedialog, messagebox
import json
import subprocess
import os

CONFIG_FILE = "config.json"
DOCKER_IMAGE = "my_container"  # название твоего докер-образа


def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"classifiers": "", "data": "", "blanks": ""}


def save_config(config):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)


def choose_folder(var, title):
    folder = filedialog.askdirectory(title=title)
    if folder:
        var.set(folder)

def run_docker():
    classifiers = classifiers_var.get()
    data = data_var.get()
    blanks = blanks_var.get()

    if not classifiers or not data or not blanks:
        messagebox.showerror("Ошибка", "Укажи все три папки!")
        return

    # сохраняем выбранные пути
    save_config({"classifiers": classifiers, "data": data, "blanks": blanks})

    # формируем команду для Docker
    cmd = [
        "docker", "run", "--rm",
        "-v", f"{classifiers}:/app/classifiers",
        "-v", f"{data}:/app/data",
        "-v", f"{blanks}:/app/blanks",
        DOCKER_IMAGE
    ]

    try:
        # запуск Docker и ожидание окончания
        subprocess.run(cmd, check=True)
    except subprocess.CalledProcessError as e:
        messagebox.showerror("Ошибка", f"Не удалось запустить контейнер:\n{e}")
        return

    # После завершения Docker запускаем add_headers.exe из папки с классификаторами
    try:
        exe_path = os.path.join(classifiers, "add_headers.exe")
        media_folder = os.path.join(classifiers, "media")
        if os.path.exists(exe_path):
            subprocess.run([exe_path, blanks, media_folder], check=True)
        else:
            messagebox.showwarning(
                "Внимание",
                f"Файл {exe_path} не найден. Header/Footer не будут добавлены."
            )
    except subprocess.CalledProcessError as e:
        messagebox.showerror("Ошибка", f"Ошибка при запуске add_headers.exe:\n{e}")
        return

    messagebox.showinfo("Готово", "Анализ завершён и header/footer добавлены!")


# === GUI ===
config = load_config()

root = tk.Tk()
root.title("Запуск анализа (Docker лаунчер)")
root.geometry("550x300")

# --- Классификаторы ---
frame1 = tk.Frame(root)
frame1.pack(fill="x", pady=5)
tk.Label(frame1, text="Папка с классификаторами:").pack(anchor="w", padx=10, pady=5)
classifiers_var = tk.StringVar(value=config.get("classifiers", ""))
tk.Entry(frame1, textvariable=classifiers_var, width=50).pack(side="left", padx=10)
tk.Button(frame1, text="Выбрать", command=lambda: choose_folder(classifiers_var, "Выбери папку с классификаторами")).pack(side="left")

# --- Данные ---
frame2 = tk.Frame(root)
frame2.pack(fill="x", pady=5)
tk.Label(frame2, text="Папка с файлами для анализа:").pack(anchor="w", padx=10, pady=5)
data_var = tk.StringVar(value=config.get("data", ""))
tk.Entry(frame2, textvariable=data_var, width=50).pack(side="left", padx=10)
tk.Button(frame2, text="Выбрать", command=lambda: choose_folder(data_var, "Выбери папку с данными")).pack(side="left")

# --- Бланки ---
frame3 = tk.Frame(root)
frame3.pack(fill="x", pady=5)
tk.Label(frame3, text="Папка с бланками заключений:").pack(anchor="w", padx=10, pady=5)
blanks_var = tk.StringVar(value=config.get("blanks", ""))
tk.Entry(frame3, textvariable=blanks_var, width=50).pack(side="left", padx=10)
tk.Button(frame3, text="Выбрать", command=lambda: choose_folder(blanks_var, "Выбери папку с бланками")).pack(side="left")

# --- Запуск ---
tk.Button(root, text="Запуск анализа", command=run_docker, bg="lightgreen", height=2, width=20).pack(pady=20)

root.mainloop()
