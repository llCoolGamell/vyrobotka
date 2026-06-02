"""
Современный интерфейс для формирования выработки по сотрудникам.

Слева  — зона загрузки исходного файла (перетащить мышью или выбрать кнопкой).
Справа — выбор папки для выгрузки и большая зелёная кнопка запуска.

Запуск:  python app.py
Сборка в .exe описана в README.md
"""

import os
import threading
import traceback

import customtkinter as ctk
from tkinter import filedialog, messagebox

import processor

# Перетаскивание мышью (необязательная зависимость)
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_AVAILABLE = True
except Exception:
    DND_AVAILABLE = False

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("green")

GREEN = "#2E7D32"
GREEN_HOVER = "#1B5E20"
CARD = "#1f2933"


class App(ctk.CTk if not DND_AVAILABLE else type("Base", (ctk.CTk,), {})):
    pass


# Базовый класс с поддержкой drag-and-drop, если доступна tkinterdnd2
if DND_AVAILABLE:
    class _Base(ctk.CTk, TkinterDnD.DnDWrapper):
        def __init__(self, *a, **k):
            super().__init__(*a, **k)
            self.TkdndVersion = TkinterDnD._require(self)
else:
    class _Base(ctk.CTk):
        pass


class WorkloadApp(_Base):
    def __init__(self):
        super().__init__()

        self.input_path = None
        self.output_dir = None

        self.title("Формирование выработки по сотрудникам")
        self.geometry("920x560")
        self.minsize(820, 500)

        # ---- Заголовок ----
        header = ctk.CTkLabel(
            self, text="Сводный запрос  →  выработка по сотрудникам",
            font=ctk.CTkFont(size=22, weight="bold"),
        )
        header.pack(pady=(22, 4))
        sub = ctk.CTkLabel(
            self,
            text="Загрузите «сводный запрос», выберите папку — программа создаст "
                 "отдельный Excel на каждого сотрудника.",
            font=ctk.CTkFont(size=13), text_color="#9aa5b1",
        )
        sub.pack(pady=(0, 16))

        # ---- Основная область: две колонки ----
        body = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=24, pady=4)
        body.grid_columnconfigure(0, weight=1)
        body.grid_columnconfigure(1, weight=1)
        body.grid_rowconfigure(0, weight=1)

        # === ЛЕВО: загрузка файла ===
        left = ctk.CTkFrame(body, corner_radius=16, fg_color=CARD)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        ctk.CTkLabel(left, text="1. Исходный файл",
                     font=ctk.CTkFont(size=16, weight="bold")).pack(pady=(18, 8))

        self.drop = ctk.CTkFrame(left, corner_radius=14, fg_color="#111820",
                                 border_width=2, border_color="#3a4a5a")
        self.drop.pack(fill="both", expand=True, padx=18, pady=8)

        self.drop_label = ctk.CTkLabel(
            self.drop,
            text=("⬇\n\nПеретащите сюда\nфайл Excel"
                  if DND_AVAILABLE else "📄\n\nВыберите файл Excel\nкнопкой ниже"),
            font=ctk.CTkFont(size=15), justify="center", text_color="#7b8a9a",
        )
        self.drop_label.pack(expand=True)

        if DND_AVAILABLE:
            self.drop.drop_target_register(DND_FILES)
            self.drop.dnd_bind("<<Drop>>", self._on_drop)
            self.drop_label.drop_target_register(DND_FILES)
            self.drop_label.dnd_bind("<<Drop>>", self._on_drop)

        ctk.CTkButton(left, text="Загрузить файл…", height=40,
                      command=self._choose_file).pack(fill="x", padx=18, pady=(8, 18))

        # === ПРАВО: папка + запуск ===
        right = ctk.CTkFrame(body, corner_radius=16, fg_color=CARD)
        right.grid(row=0, column=1, sticky="nsew", padx=(12, 0))
        ctk.CTkLabel(right, text="2. Папка для выгрузки",
                     font=ctk.CTkFont(size=16, weight="bold")).pack(pady=(18, 8))

        self.folder_label = ctk.CTkLabel(
            right, text="Папка не выбрана", font=ctk.CTkFont(size=13),
            text_color="#9aa5b1", wraplength=340, justify="center",
        )
        self.folder_label.pack(padx=18, pady=(4, 8))

        ctk.CTkButton(right, text="Выбрать папку…", height=40,
                      command=self._choose_folder).pack(fill="x", padx=18, pady=(0, 12))

        self.progress = ctk.CTkProgressBar(right)
        self.progress.set(0)
        self.progress.pack(fill="x", padx=18, pady=(4, 4))

        self.status = ctk.CTkLabel(right, text="", font=ctk.CTkFont(size=12),
                                   text_color="#9aa5b1")
        self.status.pack(pady=(2, 8))

        self.run_btn = ctk.CTkButton(
            right, text="СОЗДАТЬ СВОДНЫЙ ЗАПРОС", height=64,
            font=ctk.CTkFont(size=18, weight="bold"),
            fg_color=GREEN, hover_color=GREEN_HOVER,
            command=self._run,
        )
        self.run_btn.pack(fill="x", padx=18, pady=(6, 20))

        # Нижняя строка с выбранным файлом
        self.file_label = ctk.CTkLabel(
            self, text="Файл не выбран", font=ctk.CTkFont(size=12),
            text_color="#7b8a9a",
        )
        self.file_label.pack(pady=(0, 12))

    # ---------- обработчики ----------
    def _set_input(self, path):
        if not path:
            return
        if not path.lower().endswith((".xls", ".xlsx", ".xlsm")):
            messagebox.showwarning("Неверный файл",
                                   "Выберите файл Excel (.xls / .xlsx).")
            return
        self.input_path = path
        name = os.path.basename(path)
        self.file_label.configure(text=f"Файл: {name}", text_color="#66bb6a")
        self.drop_label.configure(text=f"✅\n\n{name}", text_color="#66bb6a")

    def _on_drop(self, event):
        path = event.data.strip().strip("{}")
        self._set_input(path)

    def _choose_file(self):
        path = filedialog.askopenfilename(
            title="Выберите сводный запрос",
            filetypes=[("Excel", "*.xls *.xlsx *.xlsm"), ("Все файлы", "*.*")],
        )
        self._set_input(path)

    def _choose_folder(self):
        d = filedialog.askdirectory(title="Куда сохранить файлы сотрудников")
        if d:
            self.output_dir = d
            self.folder_label.configure(text=d, text_color="#66bb6a")

    def _run(self):
        if not self.input_path:
            messagebox.showwarning("Нет файла", "Сначала загрузите исходный файл.")
            return
        if not self.output_dir:
            messagebox.showwarning("Нет папки", "Выберите папку для выгрузки.")
            return
        self.run_btn.configure(state="disabled", text="Обработка…")
        self.progress.set(0)
        threading.Thread(target=self._worker, daemon=True).start()

    def _worker(self):
        try:
            def prog(done, total, name):
                self.after(0, lambda: (
                    self.progress.set(done / total),
                    self.status.configure(text=f"{done}/{total}: {name}"),
                ))

            files = processor.process(self.input_path, self.output_dir, progress=prog)
            self.after(0, lambda: self._done(len(files)))
        except Exception as e:
            tb = traceback.format_exc()
            self.after(0, lambda: self._error(str(e), tb))

    def _done(self, n):
        self.progress.set(1)
        self.run_btn.configure(state="normal", text="СОЗДАТЬ СВОДНЫЙ ЗАПРОС")
        self.status.configure(text=f"Готово: создано файлов — {n}")
        messagebox.showinfo("Готово",
                            f"Создано файлов: {n}\nПапка: {self.output_dir}")

    def _error(self, msg, tb):
        self.run_btn.configure(state="normal", text="СОЗДАТЬ СВОДНЫЙ ЗАПРОС")
        self.status.configure(text="Ошибка")
        messagebox.showerror("Ошибка", msg + "\n\n" + tb)


if __name__ == "__main__":
    WorkloadApp().mainloop()
