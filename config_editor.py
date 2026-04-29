import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox
import yaml
from pathlib import Path
from typing import Dict, Any

from defaults import DEFAULT_SETTINGS, DEFAULT_FORMATS, DEFAULT_ADDFORMATS
import setting

pt_to_mm = setting.pt_to_mm
font_face = setting.font_face
version = setting.version
config_file_name = setting.config_file_name
config_path=setting.config_path

class ConfigEditor(tk.Toplevel):  # ✅ Наследуемся от Toplevel!

# Инициализация: поддержка как диалога, так и отдельного окна

    def __init__(self, parent=None):
        # ✅ ТОЛЬКО ОДНО окно создается!
        self._own_root = None
        print(f"ConfigEditor: parent={parent}")
        if parent is None:
            self._own_root = tk.Tk()
            self._own_root.withdraw()
            parent = self._own_root
            # Загружаем иконку только в режиме отдельного приложения с fallback
            try:
                icon_path = self.resource_path("img/icon_conf.png")
                icon = tk.PhotoImage(file=icon_path)
                self._own_root.iconphoto(True, icon)
            except Exception:
                # Если иконка не найдена, продолжаем без неё
                pass
        print(f"ConfigEditor: own_root={self._own_root}")
        print(f"ConfigEditor: parent={parent}")
        self._is_main = (self._own_root is not None)
        print(f"ConfigEditor: is_main={self._is_main}")
        super().__init__(parent)
        
        # Настройки окна
        self.title(f"Редактор конфигурации {version}")
        self.geometry("500x500") # Ширина x Высота
        self.resizable(True, True)
        self.changes_is_saved = True
        self.config = self._load_config()

        # Построение интерфейса
        self._build_ui()

        # Заполняем форматы в Listbox
        self._populate_formats()
        
        # Автосохранение каждые 30 сек при изменениях
        self.auto_save_enabled = tk.BooleanVar(value=False)
        self._start_auto_save_timer()

        # Обработка закрытия окна
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def on_closing(self):
        """Закрытие: сохраняем и закрываем корректно."""
        if self.changes_is_saved is False:
            # Сохранение не удалось / отменено — спросим пользователя закрывать ли без сохранения
            if not messagebox.askyesno("Закрыть без сохранения?", "Сохранение не выполнено. Закрыть без сохранения?"):
                return

        # Если это был главный экземпляр — завершаем приложение (уничтожаем root)
        if getattr(self, "_is_main", False) and getattr(self, "_own_root", None):
            try:
                self._own_root.destroy()
                return
            except Exception:
                messagebox.showerror("Ошибка", "Не удалось корректно закрыть приложение!")
                pass

        # Для обычного диалога — просто уничтожаем сам диалог
        self.destroy()

    def _load_config(self) -> Dict[str, Any]:
        """Находит существующий config.yaml или создает новый с дефолтными НАСТРОЙКАМИ
        (без форматов - они загружаются только по запросу)"""
        
        # 1. ПУТИ ПОИСКА (по приоритету)
        possible_paths = [
            Path(config_path),  # глобальный путь из setting.py
            Path(config_file_name),  # текущая папка
            Path(__file__).parent / config_file_name,  # папка скрипта
            Path.cwd() / config_file_name,  # рабочая папка
            Path.home() / config_file_name,  # домашняя папка
        ]
        
        # 2. Ищем существующий файл
        self.config_path = None
        for path in possible_paths:
            if path.exists():
                self.config_path = path
                break
        
        # 3. Если НЕ НАЙДЕН - создаем новый с дефолтными НАСТРОЙКАМИ
        if self.config_path is None:
            self.config_path = Path(config_path)
            # Создаем конфиг только с настройками (форматы - пусто)
            self.config = {
                **DEFAULT_SETTINGS,
                "formats": {}
            }
            self._save_config_to_file()
            self._ask_load_default_formats()
            print(f"✅ Создан новый конфиг: {self.config_path}")
            print(f"📐 tolerance_mm: {self.config['tolerance_mm']}")
            print(f"📦 compress_ranges: {self.config['compress_ranges']}")
            return self.config
        
        # 4. Загружаем существующий конфиг
        try:
            with open(self.config_path, 'r', encoding='utf-8') as f:
                self.config = yaml.safe_load(f) or {}
        except yaml.YAMLError as e:
            print(f"❌ Синтаксическая ошибка в {self.config_path}: {e}")
            # При ошибке парсинга - используем дефолты
            self.config = {**DEFAULT_SETTINGS, "formats": {}}
            messagebox.showerror("Ошибка конфигурации", 
                f"Ошибка синтаксиса в config.yaml:\n{e}\n\nИспользованы дефолтные настройки.")
        except Exception as e:
            print(f"⚠️ Ошибка при чтении {self.config_path}: {e}")
            # При других ошибках - используем дефолты
            self.config = {**DEFAULT_SETTINGS, "formats": {}}
            messagebox.showerror("Ошибка", 
                f"Не удалось прочитать config.yaml:\n{e}\n\nИспользованы дефолтные настройки.")
        
        # 5. Убеждаемся, что есть ключи "formats" и основные настройки
        if "formats" not in self.config:
            self.config["formats"] = {}
        if "tolerance_mm" not in self.config:
            self.config["tolerance_mm"] = DEFAULT_SETTINGS["tolerance_mm"]
        if "compress_ranges" not in self.config:
            self.config["compress_ranges"] = DEFAULT_SETTINGS["compress_ranges"]
        
        print(f"✅ Загружен: {self.config_path}")
        print(f"📐 tolerance_mm: {self.config['tolerance_mm']}")
        print(f"📦 compress_ranges: {self.config['compress_ranges']}")
        print(f"📋 Форматов: {len(self.config['formats'])}")
        
        return self.config

    def _save_config_to_file(self):
        """Сохраняет текущий config в файл"""
        try:
            with open(self.config_path, 'w', encoding='utf-8') as f:
                yaml.dump(self.config, f, default_flow_style=None, sort_keys=False, allow_unicode=True)
        except Exception as e:
            print(f"⚠️ Ошибка при сохранении конфига: {e}")

    def _ask_load_default_formats(self):
        """При создании нового конфига - предлагает загрузить дефолтные форматы"""
        if messagebox.askyesno("Загрузить форматы?", 
            "Конфигурация создана с дефолтными настройками.\n\n"
            "Загрузить стандартные форматы бумаги (A0–A5 и их комбинации)?"):
            self.config["formats"] = DEFAULT_FORMATS.copy()
            self._save_config_to_file()
            print(f"✅ Загружены {len(DEFAULT_FORMATS)} стандартных форматов")
        else:
            print("ℹ️ Форматы не загружены, можно добавить их вручную позже")

    def resource_path(self, relative_path):
        """Получает абсолютный путь к ресурсу, работает как в dev, так и в PyInstaller"""
        try:
            # PyInstaller создаёт временную папку _MEIPASS
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")
        
        return os.path.join(base_path, relative_path)

# Построение интерфейса

    def _build_ui(self):

        # Главный notebook с вкладками
        notebook = ttk.Notebook(self)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Вкладка "Основные настройки"
        main_frame = ttk.Frame(notebook)
        notebook.add(main_frame, text="Основные настройки")

        # ✅ 100% БЕЗОПАСНЫЙ ВАРИАНТ - StringVar + Entry.get()
        tolerance_frame = ttk.Frame(main_frame)
        tolerance_frame.pack(fill=tk.X, pady=(10, 0))

        self.tolerance_label = ttk.Label(tolerance_frame, text="Допуск распознавания форматов, мм:", font=(font_face, 10))
        self.tolerance_label.pack(side=tk.LEFT, pady=(0, 0), padx=(10, 8))

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

        self.tolerance_status_label = ttk.Label(tolerance_frame, font=(font_face, 10))
        self.tolerance_status_label.pack(side=tk.LEFT, padx=(15, 0))

        # ✅ ТОЛЬКО bind к Entry, НЕ к Var!
        self.tolerance_entry.bind("<KeyRelease>", self._check_tolerance_safe)
        self.tolerance_entry.bind("<FocusOut>", self._check_tolerance_safe)

        # Разделитель
        # ttk.Separator(main_frame, orient="horizontal").pack(fill=tk.X, pady=2)
        
        # Флажок сжатия диапазонов
        compress_frame = ttk.Frame(main_frame)
        compress_frame.pack(fill=tk.X, pady=(10, 0))

        self.compress_ranges_var = tk.BooleanVar(value=self.config.get("compress_ranges", True))
        compress_check = ttk.Checkbutton(
            compress_frame,
            text="Сжимать последовательные страницы (1,2,3,4 → 1-4)",
            variable=self.compress_ranges_var,
            state="normal",
            command=self.changes_made
        )
        compress_check.pack(anchor=tk.W, padx=10)

        ttk.Label(compress_frame, text="Пример: 1,2,3,45,46,47 → 1-3,45-47",
                font=(font_face, 9), foreground="gray").pack(anchor=tk.W, pady=(2, 0), padx=40)

        # Вкладка "Форматы"
        formats_frame = ttk.Frame(notebook)
        notebook.add(formats_frame, text="Форматы")

        # Левая панель — список форматов
        left_frame = ttk.LabelFrame(formats_frame, text="Используемые форматы")
        left_frame.pack(side=tk.LEFT, anchor=tk.NW, fill=tk.BOTH, pady=(5,5), padx=(5, 5), expand=True)

        # Listbox с форматами
        list_frame = ttk.Frame(left_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        formats_scroll = ttk.Scrollbar(list_frame)
        formats_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.formats_listbox = tk.Listbox(list_frame, yscrollcommand=formats_scroll.set, height=10, font=(font_face, 10))
        self.formats_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        formats_scroll.config(command=self.formats_listbox.yview)

        self.formats_listbox.bind("<<ListboxSelect>>", self.on_format_selected)
        #self.formats_listbox.bind("<Double-Button-1>", lambda e: self.edit_format())  # опционально

        # Кнопки управления форматами
        btn_frame_left = ttk.Frame(left_frame)
        btn_frame_left.pack(fill=tk.X, pady=5)

        self.delete_btn = ttk.Button(btn_frame_left, text="Удалить", command=self.delete_format)
        self.delete_btn.pack(side=tk.LEFT, padx=(5, 0))
        self.delete_all_btn = ttk.Button(btn_frame_left, text="Удалить все", command=self.delete_all_formats)
        self.delete_all_btn.pack(side=tk.LEFT, padx=(5, 0))

        # Правая панель — редактор текущего формата
        right_frame = ttk.LabelFrame(formats_frame, text="Редактировать формат", padding=10)
        right_frame.pack(side=tk.RIGHT, anchor=tk.NE, padx=(5, 5), pady=(5,5), fill=tk.BOTH, expand=False)

        ttk.Label(right_frame, text="Название:", font=(font_face, 10)).pack(anchor=tk.W, side=tk.TOP, pady=(0, 5))
        self.format_name_var = tk.StringVar()
        ttk.Entry(right_frame, textvariable=self.format_name_var, font=(font_face, 10), width=30).pack(fill=tk.X, pady=(0, 5))

        dimensions_frame = ttk.LabelFrame(right_frame, text="Размеры", padding=5)
        dimensions_frame.pack(fill=tk.X, side=tk.TOP)

        ttk.Label(dimensions_frame, text="Высота, мм:", font=(font_face, 10), width=13, anchor=tk.E).grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        self.format_height_var = tk.DoubleVar()
        ttk.Entry(dimensions_frame, textvariable=self.format_height_var, width=13, font=(font_face, 10)).grid(row=0, column=1, sticky=tk.W, padx=(5, 0), pady=(0, 5))

        ttk.Label(dimensions_frame, text="Ширина, мм:", font=(font_face, 10), width=13, anchor=tk.E).grid(row=1, column=0, sticky=tk.W, pady=(0, 5))
        self.format_width_var = tk.DoubleVar()
        ttk.Entry(dimensions_frame, textvariable=self.format_width_var, width=13, font=(font_face, 10)).grid(row=1, column=1, sticky=tk.E, padx=(5, 0), pady=(0, 5))

        button_frame = ttk.Frame(right_frame)
        button_frame.pack(fill=tk.X,  side=tk.TOP)
        self.edit_btn = ttk.Button(button_frame, text="Сохранить", command=self.edit_format)
        self.edit_btn.pack(side=tk.LEFT, padx=(0, 5), pady=(10, 0))
        self.add_btn = ttk.Button(button_frame, text="Добавить", command=self.add_format)
        self.add_btn.pack(side=tk.LEFT, padx=(0, 5), pady=(10, 0))
        
        # по умолчанию выключаем, если список пуст
        self.edit_btn.config(state="disabled")
        self.delete_btn.config(state="disabled")

        right_bottomframe = ttk.Frame(right_frame, padding=0)
        right_bottomframe.pack(side=tk.BOTTOM, anchor=tk.SW,  fill=tk.X, expand=False)

        self.load_gost_btn = ttk.Button(right_bottomframe, text="Добавить форматы из ГОСТ", command=self.add_gost_formats, width=30)
        self.load_gost_btn.pack(side=tk.TOP, padx=(5, 0), pady=(5, 5))
        self.load_gost_addformats_btn = ttk.Button(right_bottomframe, text="Добавить доп. форматы", command=self.add_gost_addformats, width=30)
        self.load_gost_addformats_btn.pack(side=tk.TOP, padx=(5, 0), pady=(5, 5), after=self.load_gost_btn)

        # Нижняя панель кнопок
        bottom_frame = ttk.Frame(self)
        bottom_frame.pack(fill=tk.X, padx=10, pady=(0, 0))

        ttk.Button(bottom_frame, text="💾 Сохранить", command=self.save_config).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(bottom_frame, text="💾 Сохранить и закрыть", command=self.save_and_close).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(bottom_frame, text="Закрыть без сохранения", command=self.on_closing).pack(side=tk.RIGHT)

        #Статус бар
        status_frame = ttk.Frame(self)
        status_frame.pack(fill=tk.X, side=tk.BOTTOM,padx=10, pady=(2, 2))

        self.status_label = tk.Label(status_frame, text="Готов", font=(font_face, 10), justify=tk.LEFT, anchor=tk.W, borderwidth=1, relief=tk.SUNKEN)
        self.status_label.pack(side=tk.LEFT, fill=tk.X, expand=True, pady=5, padx=(0, 5), ipadx=5, ipady=2)
        self.info_label = tk.Label(status_frame, text="-", font=(font_face, 10), justify=tk.LEFT, anchor=tk.W, borderwidth=1, relief=tk.SUNKEN)
        self.info_label.pack(side=tk.RIGHT, fill=tk.X, expand=True, pady=5, padx=(5, 5), ipadx=5, ipady=2)               

# Основные настройки

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

    def _check_tolerance_safe(self, event=None):
        """Проверяет значение в поле tolerance_entry и обновляет статус, безопасно обрабатывая любые ошибки"""
        try:
            # ЧИТАЕМ ТОЛЬКО Entry.get() - НИКОГДА не ломается!
            text = self.tolerance_entry.get().strip()
            
            if not text:
                self.tolerance_status_label.config(text="✗ Введите число", foreground="red")
                return
                
            # Запятая → точка
            text = text.replace(",", ".")
            
            num = float(text)
            if 0.01 <= num <= 50.0:
                self.tolerance_status_label.config(text="✓ Корректно", foreground="green")
            else:
                self.tolerance_status_label.config(text="✗ 0.01–50.0 мм", foreground="red")
                
        except ValueError:
            self.tolerance_status_label.config(text="✗ Только число", foreground="red")

# Управление форматами

    def _populate_formats(self):
        """Заполняет список форматов (отсортировано по имени)"""
        self.formats_listbox.delete(0, tk.END)
        for name in sorted(self.config["formats"].keys()):
            w, h = self.config["formats"][name]
            self.formats_listbox.insert(tk.END, f"{name}: {w} x {h}")

    def add_format(self):
        """Добавляет новый формат"""
        name = self.format_name_var.get().strip()
        
        try:
            width = self.format_width_var.get()
            height = self.format_height_var.get()
        except tk.TclError:
            messagebox.showerror("Ошибка", "Ширина и высота должны быть числами!")
            return

        if not name or width <= 0 or height <= 0:
            messagebox.showerror("Ошибка", "Введите корректное название и размеры!")
            return

        if name in self.config["formats"]:
            messagebox.showerror("Ошибка: Формат существует", "Формат с таким именем уже существует, для обновления используйте кнопку 'Сохранить'!")
            return

        # Сохранение нового формата
        self.config["formats"][name] = [int(width), int(height)]
        # Обновляем список форматов и очищаем поля
        self._populate_formats()
        self.format_name_var.set("")
        self.format_width_var.set(0.0)
        self.format_height_var.set(0.0)
        
        self.changes_made() # отмечаем, что были изменения
        self.info_label.config(text=f"Добавлен формат: {name}", foreground="black")

    def on_format_selected(self, event):
        """ При выборе формата — заполняет поля для редактирования и включает кнопки"""
        sel = self.formats_listbox.curselection()
        if not sel:
            # при отсутствии выбора — очистить поля / выключить кнопки
            self.edit_btn.config(state="disabled")
            self.delete_btn.config(state="disabled")
            return

        idx = sel[0]
        # безопасно получить имя формата (используем список ключей)
        name = list(self.config["formats"].keys())[idx]
        h, w = self.config["formats"][name]

        # безопасно установить значения в поля
        self.format_name_var.set(name)
        self.format_width_var.set(w)
        self.format_height_var.set(h)

        # включаем кнопки редактирования/удаления при выборе формата
        self.edit_btn.config(state="normal")
        self.delete_btn.config(state="normal")

    def edit_format(self):
        """Изменяет выбранный формат"""
        name = self.format_name_var.get().strip()
        
        try:
            width = self.format_width_var.get()
            height = self.format_height_var.get()
        except tk.TclError:
            messagebox.showerror("Ошибка", "Ширина и высота должны быть числами!")
            return

        if not name or width <= 0 or height <= 0:
            messagebox.showerror("Ошибка", "Введите корректное название и размеры!")
            return

        if name not in self.config["formats"]:
            messagebox.showerror("Ошибка", "Формат с таким именем не существует, для добавления используйте кнопку 'Добавить'!")
            return

        if name in self.config["formats"]:
            if not messagebox.askyesno("Сохранить изменения?", "Обновить размеры стандарта?"):
                return
        
        # Сохранение изменений формата
        self.config["formats"][name] = [int(width), int(height)]
        # Обновляем список форматов и очищаем поля
        self._populate_formats()
        self.format_name_var.set("")
        self.format_width_var.set(0.0)
        self.format_height_var.set(0.0)
        self.changes_made() # отмечаем, что были изменения
        self.info_label.config(text=f"Обновлён формат: {name}", foreground="black")

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
            self.format_name_var.set("")
            self.format_width_var.set(0.0)
            self.format_height_var.set(0.0)
            self.changes_made()
            self.info_label.config(text=f"Удалён формат: {name}", foreground="black")

    def delete_all_formats(self):
        """Удаляет все форматы с подтверждением"""
        if not self.config["formats"]:
            messagebox.showwarning("Предупреждение", "Нет форматов для удаления!")
            return

        count = len(self.config["formats"])
        if messagebox.askyesno("Подтверждение", f"Удалить все {count} форматов?"):
            self.config["formats"].clear()
            self._populate_formats()
            self.format_name_var.set("")
            self.format_width_var.set(0.0)
            self.format_height_var.set(0.0)
            self.edit_btn.config(state="disabled")
            self.delete_btn.config(state="disabled")
            self.changes_made()
            self.info_label.config(text=f"Удалены все {count} форматов", foreground="black")

    def add_gost_formats(self):
        """Загрузить форматы по ГОСТу (добавляются к существующим, не заменяют их)"""
        # Добавляем дефолтные форматы, не заменяя существующие
        added_count = 0
        for name, dimensions in DEFAULT_FORMATS.items():
            if name not in self.config["formats"]:
                self.config["formats"][name] = dimensions
                added_count += 1
        
        self._populate_formats()
        self.changes_made() # отмечаем, что были изменения
        if added_count > 0:
            self.info_label.config(text=f"Добавлено {added_count} новых форматов", foreground="green")
        else:
            self.info_label.config(text=f"Все {len(DEFAULT_FORMATS)} форматы уже присутствуют", foreground="blue")
    
    def add_gost_addformats(self):
        """Загрузить кратные форматы отсутствующие в ГОСТе (добавляются к существующим, не заменяют их)"""
        # Добавляем дефолтные форматы, не заменяя существующие
        added_count = 0
        for name, dimensions in DEFAULT_ADDFORMATS.items():
            if name not in self.config["formats"]:
                self.config["formats"][name] = dimensions
                added_count += 1
        
        self._populate_formats()
        self.changes_made() # отмечаем, что были изменения
        if added_count > 0:
            self.info_label.config(text=f"Добавлено {added_count} новых форматов", foreground="green")
        else:
            self.info_label.config(text=f"Все {len(DEFAULT_ADDFORMATS)} форматы уже присутствуют", foreground="blue")

# Сохранение и автосохранение

    def changes_made(self):
        """Отмечает, что были внесены изменения"""
        self.status_label.config(text="Изменения не сохранены", foreground="black")
        self.changes_is_saved = False

    def save_config(self):
        try:
            # ЧИТАЕМ Entry напрямую!
            text = self.tolerance_entry.get().strip().replace(",", ".")
            if not text:
                messagebox.showerror("Ошибка", "Введите допуск!")
                return
                
            num = float(text)
            if not (0.01 <= num <= 50.0):
                messagebox.showerror("Ошибка", "Допуск: 0.01–50.0 мм!")
                return
            
            # Сохраняем в конфиг
            self.config["tolerance_mm"] = num
            self.config["compress_ranges"] = self.compress_ranges_var.get()
            
            # Сохраняем в файл
            with open(self.config_path, 'w', encoding='utf-8') as f:
                yaml.dump(self.config, f, default_flow_style=None, sort_keys=False, allow_unicode=True)
            self.status_label.config(text="Сохранено!", foreground="green") # обновляем статус

            # обновляем форматы в интерфейсе (на случай, если были изменения)
            self._populate_formats()
            # отмечаем, что изменения сохранены
            self.changes_is_saved = True

        except ValueError:
            messagebox.showerror("Ошибка", "Допуск должен быть числом!")
            self.changes_is_saved = False
        except Exception as e:
            messagebox.showerror("Ошибка", f"При сохранении возникла ошибка:\n{str(e)}")
            self.changes_is_saved = False

    def save_and_close(self):
        """Сохраняет и закрывает"""
        self.save_config()
        if self.changes_is_saved:
            self.on_closing()
        else:
            messagebox.showerror("Ошибка", "Сохранение не выполнено, проверьте настройки!") 

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

