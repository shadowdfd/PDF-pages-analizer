import tkinter as tk
from tkinter import ttk, messagebox
import yaml
from pathlib import Path
from typing import Dict, Any

font_face = "Calibri"

class ConfigEditor(tk.Toplevel):  # ✅ Наследуемся от Toplevel!

    def __init__(self, parent=None):
        # ✅ ТОЛЬКО ОДНО окно создается!
        self._own_root = None
        print(f"ConfigEditor: parent={parent}")
        if parent is None:
            self._own_root = tk.Tk()
            self._own_root.withdraw()
            parent = self._own_root
        print(f"ConfigEditor: own_root={self._own_root}")
        print(f"ConfigEditor: parent={parent}")
        self._is_main = (self._own_root is not None)
        print(f"ConfigEditor: is_main={self._is_main}")
        super().__init__(parent)
        
        self.title("Редактор конфигурации")
        self.geometry("500x500")
        self.resizable(True, True)
        self.changes_is_saved = True
        self.config = self._load_config()

        self._build_ui()
        self._populate_formats()
        
        # Автосохранение каждые 30 сек при изменениях
        self.auto_save_enabled = tk.BooleanVar(value=False)
        self._start_auto_save_timer()

        # Кнопка закрытия
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def on_closing(self):
        """Закрытие: сохраняем и закрываем корректно."""
        if self.changes_is_saved is False:
            # Сохранение не удалось / отменено — спросим пользователя закрывать ли без сохранения
            if not messagebox.askyesno("Закрыть без сохранения?", "Сохранение не выполнено. Закрыть без сохранения?"):
                return

        # Если это был главный экземпляр — завершаем приложение (уничтожаем root)
        if getattr(self, "_is_main", False) and getattr(self, "_own_root", None):
            #try:
                # попытаться уничтожить корневой Tk (если он есть)
                #root = self.master if isinstance(self.master, tk.Tk) else self.winfo_toplevel()
            try:
                self._own_root.destroy()
                return
            except Exception:
                pass

        # Для обычного диалога — просто уничтожаем сам диалог
        self.destroy()

    def _load_config(self) -> Dict[str, Any]:
        """Автоматически находит config.yaml"""
        
        # 1. ПУТИ ПОИСКА (по приоритету)
        possible_paths = [
            Path("config.yaml"),  # текущая папка
            Path(__file__).parent / "config.yaml",  # папка скрипта
            Path.cwd() / "config.yaml",  # рабочая папка
            Path.home() / "config.yaml",  # домашняя папка
        ]
        
        # 2. Ищем существующий файл
        self.config_path = None
        for path in possible_paths:
            if path.exists():
                self.config_path = path
                break
        
        # 3. Если НЕ НАЙДЕН - создаем в текущей папке
        if self.config_path is None:
            self.config_path = Path("config.yaml")
            print(f"⚠️ config.yaml не найден. Создан: {self.config_path}")
        
        # 4. Загружаем/создаем конфиг
        default_config = {
            "tolerance_mm": 5.0,
            "compress_ranges": True,
            "formats": {
                "A0": [841, 1189],
                "A0×2": [1189, 1682],
                "A0×3": [1189, 2523],
                "A1": [594, 841],
                "A1×3": [841, 1783],
                "A1×4": [841, 2378],
                "A2": [420, 594],
                "A2×3": [594, 1261],
                "A2×4": [594, 1682],
                "A2×5": [594, 2102],
                "A3": [297, 420],
                "A3×3": [420, 891],
                "A3×4": [420, 1189],
                "A3×5": [420, 1486],
                "A3×6": [420, 1783],
                "A3×7": [420, 2080],
                "A4": [210, 297],
                "A4×3": [297, 630],
                "A4×4": [297, 841],
                "A4×5": [297, 1051],
                "A4×6": [297, 1261],
                "A4×7": [297, 1471],
                "A4×8": [297, 1682],
                "A4×9": [297, 1892],
                "A5": [148, 210]
            }
        }
        
        try:
            with open(self.config_path, 'r', encoding='utf-8') as f:
                self.config = yaml.safe_load(f) or {}
        except Exception:
            self.config = {}
        
        # 5. Дополняем дефолтными значениями
        self.config = {**default_config, **self.config}
        
        print(f"✅ Загружен: {self.config_path}")
        print(f"📐 tolerance_mm: {self.config['tolerance_mm']}")
        print(f"📦 compress_ranges: {self.config['compress_ranges']}")
        
        return self.config

    def _build_ui(self):

        # Главный notebook с вкладками
        notebook = ttk.Notebook(self)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Вкладка "Основные настройки" — ГАРАНТИРОВАННО РАБОТАЕТ
        main_frame = ttk.Frame(notebook)
        notebook.add(main_frame, text="Основные")

        # ✅ 100% БЕЗОПАСНЫЙ ВАРИАНТ - StringVar + Entry.get()
        tolerance_frame = ttk.Frame(main_frame)
        tolerance_frame.pack(fill=tk.X, pady=5)

        self.tolerance_label = ttk.Label(tolerance_frame, text="Допуск распознавания форматов (мм):", font=(font_face, 10))
        self.tolerance_label.pack(side=tk.LEFT, pady=(10, 5), padx=(10, 8))

        # StringVar НЕ выбрасывает ошибки!
        self.tolerance_var = tk.StringVar(value=str(self.config["tolerance_mm"]))
        self.tolerance_entry = tk.Entry(  # Сохраняем ссылку!
            tolerance_frame, 
            textvariable=self.tolerance_var,
            width=8,
            font=(font_face, 10),
            justify="center",
            bg="white"
        )
        self.tolerance_var.trace_add('write', self._on_tolerance_change)
           
        self.tolerance_entry.pack(side=tk.LEFT, pady=2)
        self.tolerance_entry.focus_set()
        #self.tolerance_entry.select_range(0, tk.END)

        ttk.Label(tolerance_frame, text="мм", font=(font_face, 10)).pack(side=tk.LEFT, padx=(8, 0))

        self.tolerance_status_label = ttk.Label(tolerance_frame, text="✓ Корректно", foreground="green", font=(font_face, 10))
        self.tolerance_status_label.pack(side=tk.LEFT, padx=(15, 0))

        # ✅ ТОЛЬКО bind к Entry, НЕ к Var!
        self.tolerance_entry.bind("<KeyRelease>", self._check_tolerance_safe)
        self.tolerance_entry.bind("<FocusOut>", self._check_tolerance_safe)

        # Разделитель
        ttk.Separator(main_frame, orient="horizontal").pack(fill=tk.X, pady=2)
        
        # Флажок сжатия диапазонов
        ttk.Label(main_frame, text="Сжатие диапазонов страниц:", 
                font=(font_face, 10)).pack(anchor=tk.W, pady=(10, 5), padx=10)

        compress_frame = ttk.Frame(main_frame)
        compress_frame.pack(fill=tk.X, pady=5)

        self.compress_ranges_var = tk.BooleanVar(value=self.config.get("compress_ranges", True))
        compress_check = ttk.Checkbutton(
            compress_frame,
            text="Сжимать последовательные страницы (1,2,3,4 → 1-4)",
            variable=self.compress_ranges_var,
            state="normal",
            command=self.changes_made
        )
        compress_check.pack(anchor=tk.W, padx=20)

        ttk.Label(compress_frame, text="Пример: 1,2,3,45,46,47 → 1-3,45-47", 
                font=(font_face, 9), foreground="gray").pack(anchor=tk.W, pady=(2, 0), padx=40)

        # Вкладка "Форматы"
        formats_frame = ttk.Frame(notebook)
        notebook.add(formats_frame, text="Форматы")

        # Левая панель — список форматов
        left_frame = ttk.Frame(formats_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, padx=(10, 5))

        ttk.Label(left_frame, text="Используемые форматы:", font=(font_face, 10)).pack(anchor=tk.W)

        # Listbox с форматами
        list_frame = ttk.Frame(left_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        formats_scroll = ttk.Scrollbar(list_frame)
        formats_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.formats_listbox = tk.Listbox(list_frame, yscrollcommand=formats_scroll.set, height=10, font=(font_face, 10))
        self.formats_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        formats_scroll.config(command=self.formats_listbox.yview)

        self.formats_listbox.bind("<<ListboxSelect>>", self.on_format_selected)
        self.formats_listbox.bind("<Double-Button-1>", lambda e: self.edit_format())  # опционально

        # Кнопки управления форматами
        btn_frame_left = ttk.Frame(left_frame)
        btn_frame_left.pack(fill=tk.X, pady=5)

        self.add_btn = ttk.Button(btn_frame_left, text="Добавить", command=self.add_format)
        self.add_btn.pack(side=tk.LEFT, padx=(0, 5))
        self.delete_btn = ttk.Button(btn_frame_left, text="Удалить", command=self.delete_format)
        self.delete_btn.pack(side=tk.LEFT)

        # Правая панель — редактор текущего формата
        right_frame = ttk.LabelFrame(formats_frame, text="Редактировать формат", padding=10)
        right_frame.pack(side=tk.RIGHT, anchor=tk.NE, padx=(5, 10), fill=tk.X, expand=True)

        ttk.Label(right_frame, text="Название:", font=(font_face, 10)).pack(anchor=tk.W)
        self.format_name_var = tk.StringVar()
        ttk.Entry(right_frame, textvariable=self.format_name_var, font=(font_face, 10), width=20).pack(fill=tk.X, pady=(0, 10))

        dimensions_frame = ttk.Frame(right_frame)
        dimensions_frame.pack(fill=tk.X)

        names_frame = ttk.Frame(dimensions_frame)
        names_frame.pack(fill=tk.X, side=tk.LEFT)
        ttk.Label(names_frame, justify=tk.RIGHT, text="Ширина (мм):", font=(font_face, 10)).pack(side=tk.TOP, anchor=tk.E, pady=(0, 5))
        ttk.Label(names_frame, justify=tk.RIGHT, text="Высота (мм):", font=(font_face, 10)).pack(side=tk.TOP, anchor=tk.E, pady=(0, 5))

        entrys_frame = ttk.Frame(dimensions_frame)
        entrys_frame.pack(fill=tk.X, expand=True, side=tk.RIGHT)
        self.format_width_var = tk.DoubleVar()
        ttk.Entry(entrys_frame, textvariable=self.format_width_var, width=12, font=(font_face, 10)).pack(side=tk.TOP, padx=(5, 0), pady=(0, 5))
        self.format_height_var = tk.DoubleVar()
        ttk.Entry(entrys_frame, textvariable=self.format_height_var, width=12, font=(font_face, 10)).pack(side=tk.TOP, padx=(5, 0), pady=(0, 5))
        
        self.edit_btn = ttk.Button(right_frame, text="Сохранить", command=self.edit_format)
        self.edit_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        # по умолчанию выключаем, если список пуст
        self.edit_btn.config(state="disabled")
        self.delete_btn.config(state="disabled")

        # Нижняя панель кнопок
        bottom_frame = ttk.Frame(self)
        bottom_frame.pack(fill=tk.X, padx=10, pady=(0, 0))

        ttk.Button(bottom_frame, text="💾 Сохранить", command=self.save_config).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(bottom_frame, text="💾 Сохранить и закрыть", command=self.save_and_close).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(bottom_frame, text="Закрыть без сохранения", command=self.on_closing).pack(side=tk.RIGHT)

        #Статус бар
        status_frame = ttk.Frame(self)
        status_frame.pack(fill=tk.X, side=tk.BOTTOM,padx=10, pady=(2, 2))

        self.status_label = tk.Label(status_frame, text="Готов", font=(font_face, 10), width=30, justify=tk.LEFT, anchor=tk.W)
        self.status_label.pack(side=tk.LEFT, pady=5, anchor=tk.W, padx=(0, 10) )       

    def _on_tolerance_change(self, var, index, mode):
            self.changes_made()

    def _validate_number(self, value):
        """Валидация чисел для tolerance_entry"""
        if value == "":
            return True
        try:
            float(value)
            return 0.1 <= float(value) <= 50.0  # допуск 0.1-50мм
        except ValueError:
            return False

    def _populate_formats(self):
        """Заполняет список форматов"""
        self.formats_listbox.delete(0, tk.END)
        for name, (w, h) in self.config["formats"].items():
            self.formats_listbox.insert(tk.END, f"{name}: {w}×{h} мм")

    def add_format(self):
        """Добавляет новый формат"""
        name = self.format_name_var.get().strip()
        width = self.format_width_var.get()
        height = self.format_height_var.get()

        if not name or width <= 0 or height <= 0:
            messagebox.showerror("Ошибка", "Введите корректное название и размеры!")
            return

        if name in self.config["formats"]:
            messagebox.showerror("Ошибка", "Формат с таким именем уже существует!")
            return

        self.config["formats"][name] = [int(width), int(height)]
        self._populate_formats()
        self.format_name_var.set("")
        self.changes_made()
        messagebox.showinfo("Успех", f"Добавлен формат: {name}")

    def on_format_selected(self, event):
        sel = self.formats_listbox.curselection()
        if not sel:
            # при отсутствии выбора — очистить поля / выключить кнопки
            self.edit_btn.config(state="disabled")
            self.delete_btn.config(state="disabled")
            return

        idx = sel[0]
        # безопасно получить имя формата (используем список ключей)
        name = list(self.config["formats"].keys())[idx]
        w, h = self.config["formats"][name]

        self.format_name_var.set(name)
        self.format_width_var.set(w)
        self.format_height_var.set(h)

        self.edit_btn.config(state="normal")
        self.delete_btn.config(state="normal")

    def edit_format(self):
        """Изменяет выбранный формат"""
        name = self.format_name_var.get().strip()
        width = self.format_width_var.get()
        height = self.format_height_var.get()

        if not name or width <= 0 or height <= 0:
            messagebox.showerror("Ошибка", "Введите корректное название и размеры!")
            return

        if name not in self.config["formats"]:
            messagebox.showerror("Ошибка", "Формат с таким именем не существует!")
            return

        if name in self.config["formats"]:
            if not messagebox.askyesno("Сохранить изменения?", "Обновить размеры стандарта?"):
                return
        
        self.config["formats"][name] = [int(width), int(height)]
        self._populate_formats()
        self.format_name_var.set("")
        self.changes_made()
        messagebox.showinfo("Успех", f"Обновлён формат: {name}")

    def delete_format(self):
        """Удаляет выбранный формат"""
        selection = self.formats_listbox.curselection()
        if not selection:
            messagebox.showwarning("Предупреждение", "Выберите формат для удаления!")
            return

        index = selection[0]
        name = list(self.config["formats"].keys())[index]

        if messagebox.askyesno("Подтверждение", f"Удалить формат '{name}'?"):
            del self.config["formats"][name]
            self._populate_formats()
            self.changes_made()
            messagebox.showinfo("Успех", f"Удалён формат: {name}")

    def changes_made(self):
        """Отмечает, что были внесены изменения"""
        self.status_label.config(text="Изменения не сохранены", foreground="black")
        self.changes_is_saved = False

    def _check_tolerance_safe(self, event=None):
        """100% БЕЗОПАСНО - читает Entry напрямую"""
        try:
            # ЧИТАЕМ ТОЛЬКО Entry.get() - НИКОГДА не ломается!
            text = self.tolerance_entry.get().strip()
            
            if not text:
                self.tolerance_status_label.config(text="✗ Введите число", foreground="red")
                return
                
            # Запятая → точка
            text = text.replace(",", ".")
            
            num = float(text)
            if 0.1 <= num <= 50.0:
                self.tolerance_status_label.config(text="✓ Корректно", foreground="green")
            else:
                self.tolerance_status_label.config(text="✗ 0.1–50.0 мм", foreground="red")
                
        except ValueError:
            self.tolerance_status_label.config(text="✗ Только число", foreground="red")

    def save_config(self):
        try:
            # ЧИТАЕМ Entry напрямую!
            text = self.tolerance_entry.get().strip().replace(",", ".")
            if not text:
                messagebox.showerror("Ошибка", "Введите допуск!")
                return
                
            num = float(text)
            if not (0.1 <= num <= 50.0):
                messagebox.showerror("Ошибка", "Допуск: 0.1–50.0 мм!")
                return
            
            self.config["tolerance_mm"] = num
            self.config["compress_ranges"] = self.compress_ranges_var.get()
            
            with open(self.config_path, 'w', encoding='utf-8') as f:
                yaml.dump(self.config, f, default_flow_style=None, sort_keys=False, allow_unicode=True)
            self.status_label.config(text="Сохранено!", foreground="green")
            #messagebox.showinfo("✅ Сохранено!", f"Файл:\n{self.config_path}")
            self._populate_formats()
            self.changes_is_saved = True

        except ValueError:
            messagebox.showerror("Ошибка", "Допуск должен быть числом!")
            self.changes_is_saved = False
        except Exception as e:
            messagebox.showerror("Ошибка", f"Сохранение:\n{str(e)}")
            self.changes_is_saved = False

    def save_and_close(self):
        """Сохраняет и закрывает"""
        self.save_config()
        self.on_closing()

    def _start_auto_save_timer(self):
        """Автосохранение каждые 30 секунд"""
        def auto_save():
            if self.auto_save_enabled.get():
                self.save_config()
            self.after(30000, auto_save)  # 30 секунд
        
        auto_save()

    def run(self):
        self.mainloop()

if __name__ == "__main__":
    # Запускаем ConfigEditor как главное окно!
    app = ConfigEditor()  # parent=None → создаст свое Tk()
    app.mainloop()

