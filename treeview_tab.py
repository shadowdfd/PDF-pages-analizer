"""
Модуль вкладки "Сводка по форматам" с древовидной структурой (Концепция 1)
Использует ttk.Treeview для иерархического отображения данных
"""

import tkinter as tk
from tkinter import ttk
import pandas as pd
from typing import Dict, List, Tuple, Optional

Debug_mode = False

class TreeviewTab:
    """Вкладка с древовидным представлением данных анализа PDF"""

    def __init__(self, parent, font_face="Segoe UI", app=None):
        self.parent = parent
        self.font_face = font_face
        self.app = app
        self.df = None
        self.summary = None
        self.roll_summary = pd.DataFrame()
        self.search_var = tk.StringVar()
        self.filter_var = tk.StringVar(value="Все")

        # Создаем основной фрейм вкладки
        self.frame = ttk.Frame(parent)

        # Инициализация интерфейса
        self._create_toolbar()
        self._create_treeview()
        self._create_context_menu()

    def _create_toolbar(self):
        """Создает панель инструментов с поиском и фильтрами"""
        toolbar = ttk.Frame(self.frame)
        toolbar.pack(fill=tk.X, padx=10, pady=5)

        # Поиск
        ttk.Label(toolbar, text="Поиск:", font=(self.font_face, 10)).pack(side=tk.LEFT, padx=(0, 5))
        self.search_entry = ttk.Entry(toolbar, textvariable=self.search_var, width=30, font=(self.font_face, 10), state='normal')
        self.search_entry.pack(side=tk.LEFT, padx=(0, 15))
        self.search_var.trace_add("write", lambda *args: self._apply_filter())

        # Фильтр
        ttk.Label(toolbar, text="Фильтр:", font=(self.font_face, 10)).pack(side=tk.LEFT, padx=(0, 5))
        self.filter_combo = ttk.Combobox(
            toolbar,
            textvariable=self.filter_var,
            values=["Все", "Только форматы", "Только по устройствам", "Только Ч/Б", "Только цветные"],
            state="readonly",
            width=15,
            font=(self.font_face, 10)
        )
        self.filter_combo.pack(side=tk.LEFT, padx=(0, 15))
        self.filter_combo.bind("<<ComboboxSelected>>", lambda e: self._apply_filter())

        # Кнопки управления
        ttk.Button(
            toolbar,
            text="Развернуть всё",
            command=self._expand_all,
            width=15
        ).pack(side=tk.LEFT, padx=5)

        ttk.Button(
            toolbar,
            text="Свернуть всё",
            command=self._collapse_all,
            width=15
        ).pack(side=tk.LEFT, padx=5)

        # Легенда цветовой маркировки строк форматов
        legend = ttk.Frame(toolbar)
        legend.pack(side=tk.RIGHT, padx=(12, 0))
        # ttk.Label(legend, text="Цветность:", font=(self.font_face, 9, "bold")).pack(
        #     side=tk.LEFT, padx=(0, 6)
        # )
        for label, color in (
            ("Ч/Б", "#a5f7ac"),
            ("Цвет.", "#95d2fd"),
            ("Смеш.", "#f8ee98"),
        ):
            swatch = tk.Label(
                legend,
                text=label,
                bg=color,
                fg="#202020",
                font=(self.font_face, 9),
                padx=5,
                pady=2,
                relief=tk.GROOVE,
                bd=1,
            )
            swatch.pack(side=tk.LEFT, padx=2)

    def _create_treeview(self):
        """Создает древовидное представление"""
        # Контейнер с прокруткой
        tree_container = ttk.Frame(self.frame)
        tree_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        # Scrollbars
        vsb = ttk.Scrollbar(tree_container, orient="vertical")
        hsb = ttk.Scrollbar(tree_container, orient="horizontal")

        # Treeview с колонками
        columns = ("count", "bw", "color", "pages")
        self.tree = ttk.Treeview(
            tree_container,
            columns=columns,
            show="tree headings",
            yscrollcommand=vsb.set,
            xscrollcommand=hsb.set,
            height=20
        )

        vsb.config(command=self.tree.yview)
        hsb.config(command=self.tree.xview)

        # Заголовки колонок
        self.tree.heading("#0", text="Файл/Формат", anchor=tk.W)
        self.tree.heading("count", text="Количество", anchor=tk.CENTER)
        self.tree.heading("bw", text="Ч/Б", anchor=tk.CENTER)
        self.tree.heading("color", text="Цветных", anchor=tk.CENTER)
        self.tree.heading("pages", text="Страницы", anchor=tk.W)

        # Ширина колонок
        self.tree.column("#0", width=400, minwidth=200, stretch=True)
        self.tree.column("count", width=100, minwidth=80, anchor=tk.CENTER)
        self.tree.column("bw", width=80, minwidth=60, anchor=tk.CENTER)
        self.tree.column("color", width=80, minwidth=60, anchor=tk.CENTER)
        self.tree.column("pages", width=300, minwidth=150, stretch=True)

        # Упаковка
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.pack(fill=tk.BOTH, expand=True)

        # Настройка тегов для цветового кодирования
        self.tree.tag_configure("bw_only", background="#a5f7ac")  # Светло-зеленый
        self.tree.tag_configure("color_only", background="#95d2fd")  # Светло-синий
        self.tree.tag_configure("mixed", background="#f8ee98")  # Светло-желтый
        self.tree.tag_configure("neutral", background="")  # Без фона (для лазера и рулонов)
        self.tree.tag_configure("file", font=(self.font_face, 10, "bold"))
        self.tree.tag_configure("highlight", background="#ffeb3b", foreground="#000000")  # Подсветка поиска

        # Привязка событий
        self.tree.bind("<Button-3>", self._show_context_menu)
        # Двойной клик - полностью переопределяем встроенное поведение
        self.tree.bind("<Double-Button-1>", self._on_double_click)

    def _create_context_menu(self):
        """Создает контекстное меню"""
        self.context_menu = tk.Menu(self.frame, tearoff=0)
        self.context_menu.add_command(label="Копировать диапазон", command=self._copy_page_range)
        self.context_menu.add_command(label="Копировать строку", command=self._copy_selected_row)
        self.context_menu.add_command(label="Копировать дочерние элементы", command=self._copy_children)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="Развернуть/Свернуть", command=self._toggle_node)
        self.context_menu.add_command(label="Развернуть всё", command=self._expand_all)
        self.context_menu.add_command(label="Свернуть всё", command=self._collapse_all)

    def update_data(self, data: Dict = None, df: pd.DataFrame = None, summary: pd.DataFrame = None, roll_summary: pd.DataFrame = None):
        """Обновляет данные в древовидном представлении

        Args:
            data: Словарь с ключами 'format_data' и 'roll_summary' (новый формат)
            df: Полный DataFrame с данными по всем страницам (устаревший)
            summary: Сводка по форматам (устаревший)
            roll_summary: Сводка по рулонам (устаревший)
        """
        # Определяем формат данных
        if data is not None and isinstance(data, dict):
            # Новый формат - словарь
            self.data = data
            self.df = None
            self.summary = None
            self.roll_summary = None
        elif df is not None or summary is not None:
            # Старый формат - DataFrame
            self.data = None
            self.df = df
            self.summary = summary
            self.roll_summary = roll_summary if roll_summary is not None else pd.DataFrame()
        else:
            # Нет данных
            print("[TreeviewTab] Нет данных для обновления")
            return

        self._populate_tree()

    def _populate_tree(self):
        """Заполняет дерево данными"""
        try:
            # Очищаем существующие данные
            for item in self.tree.get_children():
                self.tree.delete(item)

            # Проверяем наличие данных в новом или старом формате
            has_data = False
            if self.data is not None:
                has_data = bool(self.data.get("format_data"))
            elif self.df is not None and self.summary is not None:
                has_data = not self.df.empty and not self.summary.empty

            if not has_data:
                # Показываем placeholder
                self.tree.insert("", "end", text="Нет данных для отображения", values=("", "", "", ""))
                return

            # Группируем данные по файлам
            files_data = self._prepare_tree_data()

            # Заполняем дерево
            for file_info in files_data:
                self._add_file_node(file_info)
        except Exception as e:
            print(f"[TreeviewTab] ОШИБКА в _populate_tree: {e}")
            import traceback
            traceback.print_exc()
            raise

    def _prepare_tree_data_from_dict(self) -> List[Dict]:
        """Подготавливает данные из словаря (новый формат)"""
        format_data = self.data.get("format_data", [])
        roll_summary = self.data.get("roll_summary", [])

        # Группируем по файлам
        files_dict = {}

        # Обрабатываем format_data
        for row in format_data:
            file_name = row["Файл"]
            if file_name not in files_dict:
                files_dict[file_name] = {
                    "name": file_name,
                    "total_pages": row["Итого страниц"],
                    "total_bw": 0,
                    "total_color": 0,
                    "formats": [],
                    "laser": None,
                    "rolls": []
                }

            # Накапливаем счетчики
            files_dict[file_name]["total_bw"] += row["Ч/Б"]
            files_dict[file_name]["total_color"] += row["Цветных"]

            # Проверяем, это лазерный формат?
            is_laser = "A4" in row["Формат"] or "A3" in row["Формат"]

            if is_laser:
                # Добавляем в лазерные форматы
                if files_dict[file_name]["laser"] is None:
                    files_dict[file_name]["laser"] = {
                        "count": 0,
                        "bw": 0,
                        "color": 0,
                        "bw_range": "",
                        "color_range": "",
                        "formats": []
                    }

                laser = files_dict[file_name]["laser"]
                laser["count"] += row["Количество"]
                laser["bw"] += row["Ч/Б"]
                laser["color"] += row["Цветных"]

                # Парсим страницы для диапазонов
                bw_pages = []
                color_pages = []
                if row["Ч/Б"] > 0:
                    # Предполагаем, что первые страницы - Ч/Б
                    page_str = str(row["Страницы"])
                    for part in page_str.split(","):
                        part = part.strip()
                        if not part or part == "-":
                            continue
                        if "-" in part:
                            try:
                                start, end = map(int, part.split("-"))
                                bw_pages.extend(range(start, min(end + 1, start + row["Ч/Б"])))
                            except ValueError:
                                pass
                        else:
                            try:
                                if len(bw_pages) < row["Ч/Б"]:
                                    bw_pages.append(int(part))
                            except ValueError:
                                pass

                if row["Цветных"] > 0:
                    # Цветные страницы идут после Ч/Б
                    page_str = str(row["Страницы"])
                    all_pages_in_format = []
                    for part in page_str.split(","):
                        part = part.strip()
                        if not part or part == "-":
                            continue
                        if "-" in part:
                            try:
                                start, end = map(int, part.split("-"))
                                all_pages_in_format.extend(range(start, end + 1))
                            except ValueError:
                                pass
                        else:
                            try:
                                all_pages_in_format.append(int(part))
                            except ValueError:
                                pass
                    color_pages = all_pages_in_format[row["Ч/Б"]:]

                # Добавляем к общим диапазонам
                if "bw_pages" not in laser:
                    laser["bw_pages"] = []
                if "color_pages" not in laser:
                    laser["color_pages"] = []
                laser["bw_pages"].extend(bw_pages)
                laser["color_pages"].extend(color_pages)

                # Добавляем формат
                laser["formats"].append({
                    "name": row["Формат"].split()[0],  # "A4" или "A3"
                    "count": row["Количество"],
                    "bw": row["Ч/Б"],
                    "color": row["Цветных"],
                    "bw_pages": row["Страницы"] if row["Ч/Б"] > 0 else "-",
                    "color_pages": row["Страницы"] if row["Цветных"] > 0 else "-",
                    "all_pages": row["Страницы"]
                })
            else:
                # Обычный формат (не лазерный, не рулонный)
                files_dict[file_name]["formats"].append({
                    "name": row["Формат"],
                    "count": row["Количество"],
                    "bw": row["Ч/Б"],
                    "color": row["Цветных"],
                    "bw_pages": row["Страницы"] if row["Ч/Б"] > 0 else "-",
                    "color_pages": row["Страницы"] if row["Цветных"] > 0 else "-",
                    "pages": row["Страницы"]
                })

        # Обрабатываем roll_summary
        for roll in roll_summary:
            file_name = roll["Файл"]
            if file_name in files_dict:
                # Преобразуем детали рулона в нужный формат
                details = []
                for detail in roll.get("Форматы", []):
                    detail_info = {
                        "name": detail["Формат"],
                        "size": "",  # У рулонных форматов размер уже в названии
                        "count": detail["Количество"],
                        "bw": detail.get("Ч/Б", 0),
                        "color": detail.get("Цветных", 0),
                        "pages": detail.get("Страницы", "-")
                    }
                    details.append(detail_info)

                # Собираем общий диапазон страниц для рулона из деталей
                all_roll_pages = []
                for detail in roll.get("Форматы", []):
                    page_str = detail.get("Страницы", "")
                    if page_str and page_str != "-":
                        for part in str(page_str).split(","):
                            part = part.strip()
                            if "-" in part:
                                try:
                                    start, end = map(int, part.split("-"))
                                    all_roll_pages.extend(range(start, end + 1))
                                except ValueError:
                                    pass
                            else:
                                try:
                                    all_roll_pages.append(int(part))
                                except ValueError:
                                    pass

                roll_pages_range = self._format_page_ranges(all_roll_pages) if all_roll_pages else "-"

                roll_info = {
                    "name": roll["Формат рулона"],
                    "count": roll["Количество"],
                    "bw": roll["Ч/Б"],
                    "color": roll["Цветных"],
                    "pages": roll_pages_range,
                    "details": details
                }
                files_dict[file_name]["rolls"].append(roll_info)
                print(f"[_prepare_tree_data_from_dict] Рулон добавлен: {roll['Формат рулона']}, details={len(details)}")
                for detail in details:
                    print(f"[_prepare_tree_data_from_dict]   Деталь: {detail}")

        # Форматируем диапазоны для лазерных форматов
        for file_info in files_dict.values():
            if file_info["laser"] is not None:
                laser = file_info["laser"]
                if "bw_pages" in laser and laser["bw_pages"]:
                    laser["bw_range"] = self._format_page_ranges(laser["bw_pages"])
                    del laser["bw_pages"]
                else:
                    laser["bw_range"] = "-"

                if "color_pages" in laser and laser["color_pages"]:
                    laser["color_range"] = self._format_page_ranges(laser["color_pages"])
                    del laser["color_pages"]
                else:
                    laser["color_range"] = "-"

        return list(files_dict.values())

    def _prepare_tree_data(self) -> List[Dict]:
        """Подготавливает данные для отображения в древовидной структуре"""

        # Если используется новый формат (словарь), возвращаем данные напрямую
        if self.data is not None:
            return self._prepare_tree_data_from_dict()

        # Иначе используем старый формат (DataFrame)
        files_data = []


        for file_name in self.summary["Файл"].unique():
            file_summary = self.summary[self.summary["Файл"] == file_name]
            file_df = self.df[self.df["Файл"] == file_name]

            total_pages = len(file_df)
            total_bw = (file_df["Цветность"] == "Ч/Б").sum()
            total_color = (file_df["Цветность"] == "Цветная").sum()

            # Лазерную группу определяем по габаритам, а не по имени:
            # так в неё попадут любые листы, помещающиеся в A3.
            laser_page_mask = (
                (file_df[["Ширина (мм)", "Высота (мм)"]].min(axis=1) <= 297 + 1)
                & (file_df[["Ширина (мм)", "Высота (мм)"]].max(axis=1) <= 420 + 1)
            )
            laser_format_names = set(
                file_df.loc[laser_page_mask, "Стандартный формат"].unique()
            )

            # Собираем форматы вне лазерной группы, чтобы листы не дублировались.
            formats = []
            for _, row in file_summary.iterrows():
                fmt_name = row["Стандартный формат"]

                if fmt_name == "A4 + A3" or fmt_name in laser_format_names:
                    continue

                fmt_count = row["Количество"]
                fmt_bw = row["Ч/Б"]
                fmt_color = row["Цветных"]
                bw_pages = row["Ч/Б страницы"]
                color_pages = row["Цветные страницы"]
                fmt_size = row.get("Размер стандарта", "")

                # Получаем общий диапазон страниц для этого формата
                fmt_pages_df = file_df[file_df["Стандартный формат"] == fmt_name]
                if len(fmt_pages_df) > 0:
                    page_numbers = sorted(fmt_pages_df["Страница"].tolist())
                    all_pages = self._format_page_ranges(page_numbers)
                else:
                    all_pages = "-"

                formats.append({
                    "name": fmt_name,
                    "size": fmt_size,
                    "count": fmt_count,
                    "bw": fmt_bw,
                    "color": fmt_color,
                    "bw_pages": bw_pages,
                    "color_pages": color_pages,
                    "all_pages": all_pages
                })

            # В лазерную группу включаем все страницы, помещающиеся на A3,
            # а не форматы, в названии которых случайно встречается A3/A4
            # (например, длинные A3xN относятся к рулонам).
            laser_formats = []

            for fmt_name in sorted(laser_format_names):
                fmt_row = file_summary[file_summary["Стандартный формат"] == fmt_name]
                if not fmt_row.empty:
                    row = fmt_row.iloc[0]
                    laser_formats.append({
                        "name": fmt_name,
                        "size": row.get("Размер стандарта", ""),
                        "count": row["Количество"],
                        "bw": row["Ч/Б"],
                        "color": row["Цветных"],
                        "bw_pages": row["Ч/Б страницы"],
                        "color_pages": row["Цветные страницы"]
                    })

            # Вычисляем общее количество для узла "Лазерная печать"
            laser_total_count = sum(f["count"] for f in laser_formats)
            laser_total_bw = sum(f["bw"] for f in laser_formats)
            laser_total_color = sum(f["color"] for f in laser_formats)

            # Собираем общие диапазоны страниц для Лазерной печати
            all_laser_bw_pages = []
            all_laser_color_pages = []
            for fmt in laser_formats:
                # Ч/Б страницы
                if fmt["bw_pages"] and fmt["bw_pages"] != "-":
                    pages_str = str(fmt["bw_pages"]).replace(",", ", ")  # Нормализуем разделители
                    for part in pages_str.split(", "):
                        part = part.strip()
                        if not part or part == "-":
                            continue
                        if "-" in part:
                            try:
                                start, end = map(int, part.split("-"))
                                all_laser_bw_pages.extend(range(start, end + 1))
                            except ValueError:
                                pass  # Пропускаем невалидные диапазоны
                        else:
                            try:
                                all_laser_bw_pages.append(int(part))
                            except ValueError:
                                pass  # Пропускаем невалидные номера

                # Цветные страницы
                if fmt["color_pages"] and fmt["color_pages"] != "-":
                    pages_str = str(fmt["color_pages"]).replace(",", ", ")  # Нормализуем разделители
                    for part in pages_str.split(", "):
                        part = part.strip()
                        if not part or part == "-":
                            continue
                        if "-" in part:
                            try:
                                start, end = map(int, part.split("-"))
                                all_laser_color_pages.extend(range(start, end + 1))
                            except ValueError:
                                pass  # Пропускаем невалидные диапазоны
                        else:
                            try:
                                all_laser_color_pages.append(int(part))
                            except ValueError:
                                pass  # Пропускаем невалидные номера

            # Форматируем диапазоны
            laser_bw_range = self._format_page_ranges(all_laser_bw_pages) if all_laser_bw_pages else "-"
            laser_color_range = self._format_page_ranges(all_laser_color_pages) if all_laser_color_pages else "-"

            # Собираем информацию о рулонах для этого файла
            rolls = []
            if not self.roll_summary.empty and file_name in self.roll_summary["Файл"].values:
                file_rolls = self.roll_summary[self.roll_summary["Файл"] == file_name]

                # Группируем по формату рулона
                for roll_format in file_rolls["Формат рулона"].unique():
                    roll_data = file_rolls[file_rolls["Формат рулона"] == roll_format]

                    # Строка ИТОГО для общей информации
                    itogo_row = roll_data[roll_data["Стандартный формат"] == "ИТОГО"]
                    if itogo_row.empty:
                        continue

                    itogo = itogo_row.iloc[0]

                    # Детальные форматы внутри рулона (исключаем строку ИТОГО)
                    detail_formats = []
                    detail_rows = roll_data[roll_data["Стандартный формат"] != "ИТОГО"]

                    if Debug_mode:
                        print(f"[DEBUG] Файл: {file_name}, Рулон: {roll_format}")
                        print(f"[DEBUG] Всего строк в roll_data: {len(roll_data)}")
                        print(f"[DEBUG] Детальных строк (не ИТОГО): {len(detail_rows)}")
                        if not detail_rows.empty:
                            print(f"[DEBUG] Форматы деталей: {detail_rows['Стандартный формат'].tolist()}")

                    for _, detail_row in detail_rows.iterrows():
                        # Собираем страницы для детального формата
                        detail_pages = []
                        bw_pages_str = detail_row["Ч/Б страницы"]
                        color_pages_str = detail_row["Цветные страницы"]

                        for pages_str in [bw_pages_str, color_pages_str]:
                            if pages_str and pages_str != "-":
                                pages_normalized = str(pages_str).replace(",", ", ")
                                for part in pages_normalized.split(", "):
                                    part = part.strip()
                                    if not part or part == "-":
                                        continue
                                    if "-" in part:
                                        try:
                                            start, end = map(int, part.split("-"))
                                            detail_pages.extend(range(start, end + 1))
                                        except ValueError:
                                            pass
                                    else:
                                        try:
                                            detail_pages.append(int(part))
                                        except ValueError:
                                            pass

                        detail_dict = {
                            "name": detail_row["Стандартный формат"],
                            "size": detail_row["Размер стандарта"],
                            "count": detail_row["Количество"],
                            "bw": detail_row["Ч/Б"],
                            "color": detail_row["Цветных"],
                            "pages": self._format_page_ranges(detail_pages) if detail_pages else "-"
                        }
                        detail_formats.append(detail_dict)

                    # Собираем все страницы для ИТОГО рулона
                    all_roll_pages = []
                    bw_itogo = itogo["Ч/Б страницы"]
                    color_itogo = itogo["Цветные страницы"]

                    for pages_str in [bw_itogo, color_itogo]:
                        if pages_str != "-":
                            pages_normalized = str(pages_str).replace(",", ", ")  # Нормализуем разделители
                            for part in pages_normalized.split(", "):
                                part = part.strip()
                                if not part or part == "-":
                                    continue
                                if "-" in part:
                                    try:
                                        start, end = map(int, part.split("-"))
                                        all_roll_pages.extend(range(start, end + 1))
                                    except ValueError:
                                        pass
                                else:
                                    try:
                                        all_roll_pages.append(int(part))
                                    except ValueError:
                                        pass

                    rolls.append({
                        "name": itogo["Формат рулона"],  # Уже содержит "🧻 Рулон 297мм"
                        "count": itogo["Количество"],
                        "bw": itogo["Ч/Б"],
                        "color": itogo["Цветных"],
                        "pages": self._format_page_ranges(all_roll_pages) if all_roll_pages else "-",
                        "details": detail_formats
                    })
                    if Debug_mode:
                        print(f"[_prepare_tree_data] Рулон добавлен: {itogo['Формат рулона']}, details={len(detail_formats)}")
                        for detail in detail_formats:
                            print(f"[_prepare_tree_data]   Деталь: {detail}")

            files_data.append({
                "name": file_name,
                "total_pages": total_pages,
                "total_bw": total_bw,
                "total_color": total_color,
                "formats": formats,
                "laser": {
                    "count": laser_total_count,
                    "bw": laser_total_bw,
                    "color": laser_total_color,
                    "bw_range": laser_bw_range,
                    "color_range": laser_color_range,
                    "formats": laser_formats
                } if laser_formats else None,
                "rolls": rolls
            })

        return files_data


    def _add_file_node(self, file_info: Dict):
        """Добавляет узел файла в дерево"""
        file_name = file_info["name"]
        total_pages = file_info["total_pages"]
        total_bw = file_info["total_bw"]
        total_color = file_info["total_color"]

        # Определяем тег для цветового кодирования
        if total_color == 0:
            tag = "bw_only"
        elif total_bw == 0:
            tag = "color_only"
        else:
            tag = "mixed"

        # Добавляем узел файла
        file_node = self.tree.insert(
            "",
            "end",
            text=f" {file_name}",
            image=self.app.icons["pdf_file16"],
            values=(total_pages, total_bw, total_color, f"1-{total_pages}"),
            tags=(tag, "file"),
            open=True  # Развернуто по умолчанию
        )

        # Добавляем дочерние узлы форматов
        for fmt in file_info["formats"]:
            self._add_format_node(file_node, fmt)

        # Добавляем узел "Лазерная печать" (если есть A4 или A3)
        if file_info.get("laser"):
            self._add_laser_node(file_node, file_info["laser"])

        # Добавляем дочерние узлы рулонов (если есть)
        for roll in file_info.get("rolls", []):
            self._add_roll_node(file_node, roll)

    def _add_format_node(self, parent, format_info: Dict):
        """Добавляет узел формата в дерево"""
        fmt_name = format_info["name"]
        fmt_size = format_info.get("size", "")
        fmt_count = format_info["count"]
        fmt_bw = format_info["bw"]
        fmt_color = format_info["color"]
        bw_pages = format_info["bw_pages"]
        color_pages = format_info["color_pages"]
        all_pages = format_info.get("all_pages", "-")

        # Определяем тег для цветового кодирования
        if fmt_color == 0:
            tag = "bw_only"
        elif fmt_bw == 0:
            tag = "color_only"
        else:
            tag = "mixed"

        # Формируем отображаемое имя с размером
        display_name = f" {fmt_name}"
        if fmt_size and fmt_size.strip() and fmt_size != "-":
            display_name += f" ({fmt_size})"

        # Добавляем узел формата с общим диапазоном страниц
        format_node = self.tree.insert(
            parent,
            "end",
            text=display_name,
            image=self.app.icons["sheet16"],
            values=(fmt_count, fmt_bw, fmt_color, all_pages),
            tags=(tag,),
            open=False
        )

        # Добавляем детализацию по цветности (если есть и Ч/Б и цветные)
        if fmt_bw > 0 and bw_pages != "-":
            self.tree.insert(
                format_node,
                "end",
                text=" Ч/Б страницы",
                image=self.app.icons["bw16"],
                values=("", fmt_bw, "", bw_pages),
                tags=("bw_only",)
            )

        if fmt_color > 0 and color_pages != "-":
            self.tree.insert(
                format_node,
                "end",
                text=" Цветные страницы",
                image=self.app.icons["color16"],
                values=("", "", fmt_color, color_pages),
                tags=("color_only",)
            )

    def _add_laser_node(self, parent, laser_info: Dict):
        """Добавляет узел 'Лазерная печать (A4+A3)' в дерево

        Структура:
        🖨️ Лазерная печать (A4+A3) [нейтральный цвет]
           ├─ Ч/Б страницы
           └─ Цветные страницы
        """
        laser_count = laser_info["count"]
        laser_bw = laser_info["bw"]
        laser_color = laser_info["color"]
        laser_bw_range = laser_info.get("bw_range", "-")
        laser_color_range = laser_info.get("color_range", "-")

        # Формируем общий диапазон в порядке страниц, независимо от цветности.
        # Нельзя просто склеивать строки Ч/Б и цветных диапазонов:
        # например, "1-3" + "5-6" должно оставаться "1-3, 5-6".
        all_laser_pages = self._expand_page_ranges(laser_bw_range)
        all_laser_pages.extend(self._expand_page_ranges(laser_color_range))
        laser_range = self._format_page_ranges(all_laser_pages)

        # Добавляем главный узел без цветового кодирования, но с общим диапазоном
        laser_node = self.tree.insert(
            parent,
            "end",
            text=" Лазерная печать (A4+A3)",
            image=self.app.icons["laser16"],
            values=(laser_count, laser_bw, laser_color, laser_range),
            tags=("neutral",),
            open=False
        )

        # A4 и A3 объединены в одну группу. Детализация только по цветности.
        if laser_bw > 0 and laser_bw_range != "-":
            self.tree.insert(
                laser_node,
                "end",
                text=" Ч/Б страницы",
                image=self.app.icons["bw16"],
                values=(laser_bw, laser_bw, 0, laser_bw_range),
                tags=("bw_only",),
            )

        if laser_color > 0 and laser_color_range != "-":
            self.tree.insert(
                laser_node,
                "end",
                text=" Цветные страницы",
                image=self.app.icons["color16"],
                values=(laser_color, 0, laser_color, laser_color_range),
                tags=("color_only",),
            )

    def _add_laser_format_node(self, parent, format_info: Dict):
        """Добавляет узел формата A4/A3 под узлом 'Лазерная печать'

        Каждый формат имеет цветовое кодирование и раскрывается на Ч/Б и Цвет
        """
        fmt_name = format_info["name"]
        fmt_size = format_info.get("size", "")
        fmt_count = format_info["count"]
        fmt_bw = format_info["bw"]
        fmt_color = format_info["color"]
        bw_pages = format_info.get("bw_pages", "-")
        color_pages = format_info.get("color_pages", "-")

        # Определяем тег для цветового кодирования
        if fmt_color == 0:
            tag = "bw_only"
        elif fmt_bw == 0:
            tag = "color_only"
        else:
            tag = "mixed"

        # Формируем отображаемое имя с размером
        display_name = f"{fmt_name}"
        if fmt_size and fmt_size.strip() and fmt_size != "-":
            display_name += f" ({fmt_size})"

        # Добавляем узел формата
        format_node = self.tree.insert(
            parent,
            "end",
            text=display_name,
            image=self.app.icons["sheet16"],
            values=(fmt_count, fmt_bw, fmt_color, "-"),
            tags=(tag,),
            open=False
        )

        # Добавляем детализацию по цветности (если есть и Ч/Б и цветные)
        if fmt_bw > 0 and bw_pages != "-":
            self.tree.insert(
                format_node,
                "end",
                text="   Ч/Б страницы",
                values=(fmt_bw, fmt_bw, 0, bw_pages),
                tags=("bw_only",)
            )

        if fmt_color > 0 and color_pages != "-":
            self.tree.insert(
                format_node,
                "end",
                text="   Цветные страницы",
                values=(fmt_color, 0, fmt_color, color_pages),
                tags=("color_only",)
            )

    def _add_roll_node(self, parent, roll_info: Dict):
        """Добавляет узел рулона в дерево

        Структура:
        🧻 Рулон 297мм [нейтральный цвет]
           ├─ 📄 A1L (размер) [без цветового кодирования]
           └─ 📄 A0L (размер) [без цветового кодирования]
        """
        roll_name = roll_info["name"]  # Уже содержит "Рулон 297мм"
        roll_count = roll_info["count"]
        roll_pages = roll_info.get("pages", "-")
        detail_formats = roll_info.get("details", [])

        if Debug_mode:
            print(f"[_add_roll_node] Рулон: {roll_name}, pages={roll_pages}, details={len(detail_formats)}")
            for detail in detail_formats:
                print(f"[_add_roll_node]   Деталь: {detail}")

        # Добавляем узел рулона без цветового кодирования (без столбцов Ч/Б и Цвет)
        roll_node = self.tree.insert(
            parent,
            "end",
            text=" "+roll_name,
            image=self.app.icons["roll16"],
            values=(roll_count, "-", "-", roll_pages),
            tags=("neutral", "roll"),
            open=False  # Развёрнуто по умолчанию, чтобы было видно детализацию
        )

        # Добавляем детальные форматы внутри рулона (без деления на Ч/Б и Цвет)
        for detail in detail_formats:
            detail_name = detail["name"]
            detail_size = detail.get("size", "")
            detail_count = detail["count"]
            detail_pages = detail.get("pages", "-")

            # Формируем отображаемое имя
            display_name = f" {detail_name}"
            if detail_size and detail_size.strip() and detail_size != "-":
                display_name += f" ({detail_size})"

            # Добавляем детальный формат без цветового кодирования
            detail_node = self.tree.insert(
                roll_node,
                "end",
                text=display_name,
                image=self.app.icons["sheet16"],
                values=(detail_count, "-", "-", detail_pages),
                tags=("neutral",)
            )

    def _format_page_ranges(self, pages):
        """
        Форматирует список номеров страниц в диапазоны.
        Например: [1, 2, 3, 5, 7, 8, 9] -> "1-3, 5, 7-9"
        """
        if not pages:
            return "-"

        # Преобразуем в целые числа (на случай если пришли float)
        pages = sorted(set(int(p) for p in pages))
        ranges = []
        start = pages[0]
        end = pages[0]

        for i in range(1, len(pages)):
            if pages[i] == end + 1:
                end = pages[i]
            else:
                if start == end:
                    ranges.append(str(start))
                else:
                    ranges.append(f"{start}-{end}")
                start = pages[i]
                end = pages[i]

        # Добавляем последний диапазон
        if start == end:
            ranges.append(str(start))
        else:
            ranges.append(f"{start}-{end}")

        return ", ".join(ranges)

    @staticmethod
    def _expand_page_ranges(page_range):
        """Раскрывает строку диапазонов страниц в список номеров."""
        if not page_range or page_range == "-":
            return []

        pages = []
        for part in str(page_range).replace(",", ", ").split(", "):
            part = part.strip()
            if not part or part == "-":
                continue
            try:
                if "-" in part:
                    start, end = map(int, part.split("-", 1))
                    pages.extend(range(start, end + 1))
                else:
                    pages.append(int(part))
            except ValueError:
                continue
        return pages

    def _apply_filter(self):
        """Применяет фильтрацию и поиск"""
        search_text = self.search_var.get().lower()
        filter_type = self.filter_var.get()

        # Очищаем дерево
        for item in self.tree.get_children():
            self.tree.delete(item)

        # Проверяем наличие данных
        has_data = False
        if self.data is not None:
            has_data = bool(self.data.get("format_data"))
        elif self.df is not None and self.summary is not None:
            has_data = not self.df.empty and not self.summary.empty

        if not has_data:
            self.tree.insert("", "end", text="Нет данных для отображения", values=("", "", "", ""))
            return

        # Получаем исходные данные
        files_data = self._prepare_tree_data()

        # Применяем фильтры
        filtered_files = self._filter_files(files_data, search_text, filter_type)

        # Заполняем дерево отфильтрованными данными
        for file_info in filtered_files:
            self._add_file_node(file_info)

        # Если после фильтрации ничего не найдено
        if not filtered_files:
            self.tree.insert("", "end", text="Ничего не найдено", values=("", "", "", ""))
        elif search_text:
            # Подсвечиваем результаты поиска
            self._highlight_search_results(search_text)
            # Разворачиваем все узлы, чтобы показать результаты
            self._expand_all()

    def _filter_files(self, files_data: List[Dict], search_text: str, filter_type: str) -> List[Dict]:
        """Фильтрует обычные форматы, лазерную печать и рулоны."""
        if not search_text and filter_type == "Все":
            return files_data

        result = []
        for file_info in files_data:
            file_match = bool(search_text) and search_text in file_info["name"].lower()
            query = "" if file_match else search_text
            formats = self._filter_formats(file_info.get("formats", []), query, filter_type)
            laser = self._filter_laser(file_info.get("laser"), query, filter_type)
            rolls = self._filter_rolls(file_info.get("rolls", []), query, filter_type)
            if not (formats or laser or rolls):
                continue

            item = file_info.copy()
            item.update(formats=formats, laser=laser, rolls=rolls)
            result.append(item)
        return result

    @staticmethod
    def _matches_filter_type(item: Dict, filter_type: str) -> bool:
        if filter_type == "Все":
            return True
        if filter_type == "Только форматы":
            return item.get("_kind") == "format"
        if filter_type == "Только по устройствам":
            return item.get("_kind") in {"laser", "roll"}
        if filter_type == "Только Ч/Б":
            return item.get("bw", 0) > 0
        if filter_type == "Только цветные":
            return item.get("color", 0) > 0
        return True

    def _filter_laser(self, laser_info: Optional[Dict], search_text: str, filter_type: str):
        if not laser_info:
            return None
        group_match = not search_text or search_text in "лазерная печать"
        formats = []
        for fmt in laser_info.get("formats", []):
            match = not search_text or search_text in fmt.get("name", "").lower()
            child = fmt.copy()
            child["_kind"] = "laser"
            if (group_match or match) and self._matches_filter_type(child, filter_type):
                formats.append(child)
        group = laser_info.copy()
        group["_kind"] = "laser"
        group_matches_type = self._matches_filter_type(group, filter_type)
        if filter_type in {"Только Ч/Б", "Только цветные"}:
            group_matches_type = group_matches_type or bool(formats)
        if not group_matches_type or not formats:
            return None
        if search_text and not group_match and not any(
            search_text in fmt.get("name", "").lower() for fmt in formats
        ):
            return None
        group["formats"] = formats
        return group

    def _filter_rolls(self, rolls: List[Dict], search_text: str, filter_type: str) -> List[Dict]:
        result = []
        for source in rolls:
            name_match = not search_text or search_text in source.get("name", "").lower()
            details = []
            for detail in source.get("details", []):
                match = not search_text or search_text in detail.get("name", "").lower()
                child = detail.copy()
                child["_kind"] = "roll"
                if (name_match or match) and self._matches_filter_type(child, filter_type):
                    details.append(child)
            roll = source.copy()
            roll["_kind"] = "roll"
            roll_matches_type = self._matches_filter_type(roll, filter_type)
            # Рулон может содержать и Ч/Б, и цветные листы: в таком случае
            # фильтр должен оставить его как родитель найденных подходящих деталей.
            if filter_type in {"Только Ч/Б", "Только цветные"}:
                roll_matches_type = roll_matches_type or bool(details)
            if not roll_matches_type:
                continue
            if search_text and not name_match and not details:
                continue
            if name_match and filter_type in {"Все", "Только по устройствам"}:
                details = [dict(d, _kind="roll") for d in source.get("details", [])]
            roll["details"] = details
            result.append(roll)
        return result

    def _filter_formats(self, formats: List[Dict], search_text: str, filter_type: str) -> List[Dict]:
        """Фильтрует обычные форматы листов."""
        result = []
        for fmt in formats:
            if search_text and search_text not in fmt.get("name", "").lower():
                continue
            item = fmt.copy()
            item["_kind"] = "format"
            if self._matches_filter_type(item, filter_type):
                result.append(item)
        return result
    def _highlight_search_results(self, search_text: str):
        """Подсвечивает результаты поиска в дереве"""
        def highlight_node(item):
            # Получаем текст узла
            text = self.tree.item(item, "text").lower()

            # Если найдено совпадение, добавляем тег подсветки
            if search_text in text:
                current_tags = list(self.tree.item(item, "tags"))
                if "highlight" not in current_tags:
                    current_tags.append("highlight")
                    self.tree.item(item, tags=current_tags)

            # Рекурсивно обрабатываем дочерние узлы
            for child in self.tree.get_children(item):
                highlight_node(child)

        # Обрабатываем все узлы верхнего уровня
        for item in self.tree.get_children():
            highlight_node(item)

    def _expand_all(self):
        """Разворачивает все узлы"""
        def expand_recursive(item):
            self.tree.item(item, open=True)
            for child in self.tree.get_children(item):
                expand_recursive(child)

        for item in self.tree.get_children():
            expand_recursive(item)

    def _collapse_all(self):
        """Сворачивает все узлы"""
        def collapse_recursive(item):
            self.tree.item(item, open=False)
            for child in self.tree.get_children(item):
                collapse_recursive(child)

        for item in self.tree.get_children():
            collapse_recursive(item)

    def _toggle_node(self):
        """Переключает состояние выбранного узла"""
        selected = self.tree.selection()
        if not selected:
            return

        item = selected[0]
        is_open = self.tree.item(item, "open")
        self.tree.item(item, open=not is_open)

    def _on_double_click(self, event):
        """Обработчик двойного клика - разворачивает/сворачивает узел"""
        # Определяем элемент под курсором
        item = self.tree.identify_row(event.y)
        if not item:
            return "break"  # Останавливаем дальнейшую обработку

        # Проверяем, есть ли у узла дочерние элементы
        children = self.tree.get_children(item)
        if not children:
            return "break"  # Нет детей - ничего не делаем

        # Переключаем состояние
        is_open = self.tree.item(item, "open")
        self.tree.item(item, open=not is_open)

        return "break"  # Останавливаем встроенную обработку ttk.Treeview

    def _show_context_menu(self, event):
        """Показывает контекстное меню"""
        # Выделяем элемент под курсором
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            try:
                self.context_menu.tk_popup(event.x_root, event.y_root)
            finally:
                self.context_menu.grab_release()

    def _copy_selected_row(self):
        """Копирует выбранную строку в буфер обмена"""
        selected = self.tree.selection()
        if not selected:
            return

        item = selected[0]
        text = self.tree.item(item, "text")
        values = self.tree.item(item, "values")

        # Формируем текст для копирования
        row_text = f"{text}\t" + "\t".join(str(v) for v in values)

        # Копируем в буфер обмена
        self.parent.clipboard_clear()
        self.parent.clipboard_append(row_text)

    def _copy_children(self):
        """Копирует все дочерние элементы выбранного узла"""
        selected = self.tree.selection()
        if not selected:
            return

        item = selected[0]
        children = self.tree.get_children(item)

        if not children:
            return

        lines = []
        for child in children:
            text = self.tree.item(child, "text")
            values = self.tree.item(child, "values")
            line = f"{text}\t" + "\t".join(str(v) for v in values)
            lines.append(line)

        # Копируем в буфер обмена
        result = "\n".join(lines)
        self.parent.clipboard_clear()
        self.parent.clipboard_append(result)

    def _copy_page_range(self):
        """Копирует диапазон страниц выбранной строки в буфер обмена"""
        selected = self.tree.selection()
        if not selected:
            return

        item = selected[0]
        values = self.tree.item(item, "values")

        # Диапазон страниц находится в последней колонке (индекс 3)
        if len(values) > 3 and values[3]:
            page_range = str(values[3])
            if page_range and page_range.strip() and page_range != "-":
                self.parent.clipboard_clear()
                self.parent.clipboard_append(page_range)
