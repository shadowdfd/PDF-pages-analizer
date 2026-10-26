import os
import sys
import re
import hashlib
from typing import Dict, Any
import subprocess
import platform
import time
import fitz  # PyMuPDF
import pandas as pd
import yaml
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
#import threading
import webbrowser
import setting
try:
    from tkinterdnd2 import TkinterDnD
except ImportError:
    TkinterDnD = None
import matplotlib
matplotlib.use('TkAgg')
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure

from defaults import DEFAULT_CONFIG, DEFAULT_SETTINGS, DEFAULT_FORMATS
from treeview_tab import TreeviewTab

pt_to_mm = setting.pt_to_mm
font_face = setting.font_face
version = setting.version
config_path=setting.config_path
# use_sheet_rotation=False

class PDFAnalyzer:
    
    def __init__(self, config_path="config.yaml"):
        # Инициализация пустых структур
        self.config = {}
        self.tolerance = DEFAULT_SETTINGS["tolerance_mm"]
        self.compress_ranges_y = DEFAULT_SETTINGS["compress_ranges"]
        self.pt_to_mm = DEFAULT_SETTINGS["pt_to_mm"]
        self.formats = {}  # Будут загружены из конфига в apply_config()

        # Обнуляем статистику
        self.stats = {
            "files_processed": 0,
            "pages_processed": 0,
            "errors": [],
            "files_skipped": 0,
            "start_time": None,
            "end_time": None
        }
        self._custom_counter = {}
        
        # Загружаем и применяем конфиг
        loaded_config = self.load_config(config_path)
        self.apply_config(loaded_config)

    def reset_stats(self):
        """Сбрасывает статистику перед новым анализом."""
        self.stats = {
            "files_processed": 0,
            "pages_processed": 0,
            "errors": [],
            "files_skipped": 0,
            "start_time": None,
            "end_time": None
        }
        self._custom_counter = {}

    def apply_config(self, config):
        """Применяет конфиг к рабочим полям анализатора.
        Дефолтные НАСТРОЙКИ используются только если отсутствуют в config.
        Форматы берутся только из config (дефолтные не подставляются автоматом)."""
        # Слияние: дефолтные настройки + загруженные настройки
        merged_config = {**DEFAULT_SETTINGS, **(config or {})}
        self.config = merged_config
        self.tolerance = merged_config.get("tolerance_mm", DEFAULT_SETTINGS["tolerance_mm"])
        self.compress_ranges_y = merged_config.get("compress_ranges", DEFAULT_SETTINGS["compress_ranges"])
        self.pt_to_mm = float(merged_config.get("pt_to_mm", DEFAULT_SETTINGS["pt_to_mm"]))
        # Форматы - только из config, без дефолтов
        self.formats = {k: tuple(v) for k, v in (config or {}).get("formats", {}).items()}
        self.config["formats"] = self.formats

    def resolve_config_path(self, config_path="config.yaml"):
        """Возвращает путь к config.yaml с учётом dev/EXE-окружения."""
        possible_paths = [
            Path(config_path),
            Path(__file__).parent / config_path,
            Path.cwd() / config_path,
        ]

        for path in possible_paths:
            if path.exists():
                return path

        return Path(config_path)

    def load_config(self, config_path):
        """Загружает конфигурацию из YAML, возвращает только загруженные значения.
        Слияние с DEFAULT_SETTINGS происходит в apply_config().
        Форматы используются только если присутствуют в config."""

        resolved_path = self.resolve_config_path(config_path)

        try:
            with open(resolved_path, 'r', encoding='utf-8') as f:
                config = yaml.safe_load(f)
                if config is None:
                    print(f"⚠️ Пустой или невалидный config.yaml, использую дефолт")
                    return
                return config
        except FileNotFoundError:
            # Файл не найден - создаем его с дефолтными значениями
            print(f"⚠️ config.yaml не найден: {resolved_path}")
            print(f"✅ Создаю новый config.yaml с дефолтными значениями...")

            try:
                # Создаем конфиг с дефолтными значениями
                with open(resolved_path, 'w', encoding='utf-8') as f:
                    yaml.dump(DEFAULT_CONFIG, f, default_flow_style=False, allow_unicode=True, sort_keys=False)

                messagebox.showinfo(
                    "Файл конфигурации создан",
                    f"Файл конфигурации config.yaml не найден.\n\n"
                    f"Создан новый файл с настройками по умолчанию:\n{resolved_path}"
                )
                print(f"✅ config.yaml успешно создан: {resolved_path}")
                return DEFAULT_CONFIG

            except Exception as create_error:
                print(f"❌ Ошибка при создании config.yaml: {create_error}")
                messagebox.showerror(
                    "Ошибка создания конфигурации",
                    f"Не удалось создать config.yaml:\n{str(create_error)}\n\nИспользуются значения по умолчанию."
                )
                return {}

        except yaml.YAMLError as e:
            print(f"❌ Синтаксическая ошибка в config.yaml: {e}")
            messagebox.showerror("Ошибка конфигурации", f"Синтаксическая ошибка в config.yaml:\n{str(e)}\n\nИспользуются значения по умолчанию.")
            return {}
        except Exception as e:
            print(f"❌ Ошибка при чтении config.yaml: {e}")
            messagebox.showerror("Файл конфигурации", f"Ошибка при попытке чтения config.yaml:\n{str(e)}\n\nИспользуются значения по умолчанию.")
            return {}

    def get_standard_format(self, w_mm: float, h_mm: float) -> tuple[str, str]:
        """Определяет формат из настроек либо создаёт новый Custom1, Custom2... для нестандартных размеров"""
        
        # Стандартные форматы
        for name, (sh, sw) in self.formats.items():
            if (abs(w_mm - sw) <= self.tolerance and abs(h_mm - sh) <= self.tolerance) or \
               (abs(w_mm - sh) <= self.tolerance and abs(h_mm - sw) <= self.tolerance):
                if w_mm <= h_mm:
                    return name, f"{int(sw)}x{int(sh)}"
                return name, f"{int(sh)}x{int(sw)}"
        
        # НЕСТАНДАРТНЫЙ формат
        size_key = f"{int(w_mm)}x{int(h_mm)}"
        
        # Счётчик уникальных нестандартных размеров
        if not hasattr(self, '_custom_counter'):
            self._custom_counter = {}
        
        if size_key not in self._custom_counter:
            self._custom_counter[size_key] = len(self._custom_counter) + 1
        
        return f"Custom{self._custom_counter[size_key]}", size_key

    def analyze_page_color(self, page: fitz.Page) -> str:
        """Определяет цветность страницы"""
        try:
            for img in page.get_images(full=True):
                pix = None
                try:
                    pix = fitz.Pixmap(page.parent, img[0])
                    if pix.colorspace and pix.colorspace.n > 1:
                        return "Цветная"
                finally:
                    # Правильное освобождение ресурсов Pixmap
                    if pix is not None:
                        del pix
            for draw in page.get_drawings():
                # PyMuPDF exposes the outline color as "color" (not "stroke").
                # Keep "stroke" as a fallback for compatible drawing objects.
                for col in (draw.get("fill"), draw.get("color", draw.get("stroke"))):
                    if col and len(col) >= 3:
                        r, g, b = col[:3]
                        if not (abs(r - g) < 1e-3 and abs(g - b) < 1e-3):
                            return "Цветная"
        except Exception as e:
            # Ошибка при анализе цвета - считаем ч/б и логируем
            print(f"⚠️ Ошибка анализа цвета страницы: {e}")
            pass
        return "Ч/Б"

    def compress_ranges(self, input_str):
        # Парсим строку в список чисел
        if not input_str or input_str == "-":
            return "-"
        
        nums = [int(x.strip()) for x in input_str.split(',')]
        
        nums.sort()  # Сортируем на всякий случай
        
        if not nums:
            return ""
        
        ranges = []
        start = nums[0]
        prev = nums[0]
        
        for num in nums[1:]:
            if num != prev + 1:  # Разрыв последовательности
                if start == prev:
                    ranges.append(str(start))  # Одиночное число
                else:
                    ranges.append(f"{start}-{prev}")
                start = num
            prev = num
        
        # Добавляем последний диапазон
        if start == prev:
            ranges.append(str(start))
        else:
            ranges.append(f"{start}-{prev}")
        
        return ','.join(ranges)

    def expand_ranges(self, input_str: str) -> list:
        """Расширяет сжатый диапазон в список чисел. Например: '1-2,6-8' -> [1,2,6,7,8]"""
        if not input_str or input_str == "-":
            return []
        
        result = []
        for part in input_str.split(','):
            part = part.strip()
            if '-' in part:
                # Это диапазон - используем rsplit для правильного парсинга
                try:
                    parts = part.rsplit('-', 1)
                    if len(parts) == 2:
                        start, end = map(int, parts)
                        result.extend(range(start, end + 1))
                except (ValueError, IndexError):
                    pass
            else:
                # Это одиночное число
                try:
                    result.append(int(part))
                except ValueError:
                    pass
        
        return sorted(set(result))

    def classify_by_roll_format(self, w_mm: float, h_mm: float, std_format: str) -> str:
        """Классифицирует лист по формату рулона, определяя по меньшей стороне листа.
        
        Параметры:
            w_mm: ширина листа (мм)
            h_mm: высота листа (мм)
            std_format: название стандартного формата (для исключений A4 и A3)
        
        Исключения: A4 и A3 возвращают 'Нестандартный'
        
        Группы по меньшей стороне формата:
        - 297мм: листы с меньшей стороной ~297мм
        - 420мм: листы с меньшей стороной ~420мм
        - 594мм: листы с меньшей стороной ~594мм
        - 841мм: листы с меньшей стороной ~841мм
        
        Возвращает: '297мм', '420мм', '594мм', '841мм' или 'Нестандартный'
        """
        # Исключаем только A4 и A3
        if std_format in ("A4", "A3"):
            return "Отдельный лист"
        
        # Определяем меньшую сторону
        min_side = min(w_mm, h_mm)
        
        # Классифицируем по меньшей стороне с допуском
        tolerance = self.tolerance
        
        if abs(min_side - 297) <= tolerance:
            return "297мм"
        elif abs(min_side - 420) <= tolerance:
            return "420мм"
        elif abs(min_side - 594) <= tolerance:
            return "594мм"
        elif abs(min_side - 841) <= tolerance:
            return "841мм"
        else:
            return "Нестандартный"

    def build_roll_summary(self, df: pd.DataFrame) -> pd.DataFrame:
        """Строит сводку по форматам рулонов."""

        # Добавляем классификацию по рулону (передаём размеры и формат)
        df_filtered = df.copy()
        df_filtered["Формат рулона"] = df_filtered.apply(
            lambda row: self.classify_by_roll_format(
                row["Ширина (мм)"],
                row["Высота (мм)"],
                row["Стандартный формат"]
            ),
            axis=1
        )

        # Исключаем "Нестандартный" из сводки
        df_filtered = df_filtered[~df_filtered["Формат рулона"].isin(["Отдельный лист", "Нестандартный"])]

        if df_filtered.empty:
            return pd.DataFrame()

        # Группировка: Файл + Формат рулона + Стандартный формат
        def get_page_list(group: pd.DataFrame, color: str) -> str:
            pages = sorted(group[group["Цветность"] == color]["Страница"].tolist())
            return ", ".join(map(str, pages)) if pages else "-"

        roll_summary_data = []

        # Группировка по файлу и формату рулона для добавления итоговых строк
        for (file_name, roll_format), file_roll_group in df_filtered.groupby(["Файл", "Формат рулона"]):

            # Добавляем итоговую строку по рулону
            all_pages_bw = sorted(file_roll_group[file_roll_group["Цветность"] == "Ч/Б"]["Страница"].tolist())
            all_pages_col = sorted(file_roll_group[file_roll_group["Цветность"] == "Цветная"]["Страница"].tolist())

            roll_summary_data.append({
                "Файл": file_name,
                "Формат рулона": f"Рулон {roll_format}",
                "Стандартный формат": "ИТОГО",
                "Размер стандарта": "-",
                "Количество": len(file_roll_group),
                "Ч/Б страницы": ", ".join(map(str, all_pages_bw)) if all_pages_bw else "-",
                "Цветные страницы": ", ".join(map(str, all_pages_col)) if all_pages_col else "-",
                "Цветных": (file_roll_group["Цветность"] == "Цветная").sum(),
                "Ч/Б": (file_roll_group["Цветность"] == "Ч/Б").sum(),
            })

            # Добавляем детальные строки по каждому формату внутри рулона
            for std_format, group in file_roll_group.groupby("Стандартный формат"):
                pages_list_bw = get_page_list(group, "Ч/Б")
                pages_list_col = get_page_list(group, "Цветная")

                roll_summary_data.append({
                    "Файл": file_name,
                    "Формат рулона": f"Рулон {roll_format}",  # Полное название, как в ИТОГО
                    "Стандартный формат": std_format,
                    "Размер стандарта": group["Размер стандарта"].iloc[0],
                    "Количество": len(group),
                    "Ч/Б страницы": pages_list_bw,
                    "Цветные страницы": pages_list_col,
                    "Цветных": (group["Цветность"] == "Цветная").sum(),
                    "Ч/Б": (group["Цветность"] == "Ч/Б").sum(),
                })

        return pd.DataFrame(roll_summary_data).reset_index(drop=True)

    def process_pdf(self, pdf_path: str, all_data: list) -> None:
        """Обрабатывает один PDF файл"""
        
        try:
            self.stats["files_processed"] += 1
            # print(f"Обрабатываем: {os.path.basename(pdf_path)}")
            doc = fitz.open(pdf_path)
            self.stats["pages_processed"] += len(doc)

            for i, page in enumerate(doc, start=1): # type: ignore
                rotation = page.rotation
                rect = page.rect
                w_pt, h_pt = rect.width, rect.height

                # if use_sheet_rotation and rotation in (90, 270):
                #     width_mm = round(h_pt * pt_to_mm, 1)
                #     height_mm = round(w_pt * pt_to_mm, 1)
                # else:
                if w_pt < h_pt:
                    temp1 = w_pt
                    w_pt = h_pt
                    h_pt = temp1
                width_mm = round(w_pt * pt_to_mm, 1)
                height_mm = round(h_pt * pt_to_mm, 1)

                std_format, std_size = self.get_standard_format(width_mm, height_mm)
                color_type = self.analyze_page_color(page)

                all_data.append({
                    "Файл": os.path.basename(pdf_path),
                    "Страница": i,
                    "Поворот": rotation,
                    "Ширина (мм)": width_mm,
                    "Высота (мм)": height_mm,
                    "Стандартный формат": std_format,
                    "Размер стандарта": std_size,
                    "Цветность": color_type,
                })
            doc.close()
        
        except Exception as e:
            self.stats["errors"].append(f"{pdf_path}: {str(e)}")
            print(f"Ошибка в {pdf_path}: {e}")

    def save_excel_safe(self, out_path: Path, df: pd.DataFrame, summary: pd.DataFrame, roll_summary: pd.DataFrame) -> str:
        """Безопасное сохранение Excel с обработкой ошибок при занятом файле"""
        max_attempts = 3
        attempt = 0
        current_path = out_path

        while attempt < max_attempts:
            try:
                with pd.ExcelWriter(current_path, engine="openpyxl") as writer:
                    df.to_excel(writer, index=False, sheet_name="Все страницы")
                    summary.to_excel(writer, index=False, sheet_name="Сводка ЕСКД")
                    if not roll_summary.empty:
                        roll_summary.to_excel(writer, index=False, sheet_name="Сводка по рулонам")

                    # Автоподбор ширины столбцов для всех листов
                    for sheet_name in writer.sheets:
                        worksheet = writer.sheets[sheet_name]
                        # Включаем выпадающий фильтр в заголовке каждой колонки.
                        # Диапазон ограничен фактически заполненной таблицей.
                        if worksheet.max_row >= 1 and worksheet.max_column >= 1:
                            worksheet.auto_filter.ref = worksheet.dimensions

                        for column in worksheet.columns:
                            max_length = 0
                            column_letter = column[0].column_letter
                            for cell in column:
                                try:
                                    if cell.value:
                                        cell_length = len(str(cell.value))
                                        max_length = max(max_length, cell_length)
                                except:
                                    pass
                            adjusted_width = min(max_length + 2, 50)
                            worksheet.column_dimensions[column_letter].width = adjusted_width

                # Успешно сохранено
                return str(current_path)

            except PermissionError:
                attempt += 1
                if attempt < max_attempts:
                    # Ищем свободное имя файла с автоинкрементом (1), (2), (3)...
                    counter = 1
                    new_path = out_path.parent / f"{out_path.stem}({counter}).xlsx"
                    while new_path.exists():
                        counter += 1
                        new_path = out_path.parent / f"{out_path.stem}({counter}).xlsx"

                    # Создаем кастомное окно с понятными кнопками
                    dialog = tk.Toplevel()
                    dialog.title("Файл занят")
                    dialog.geometry("400x150")
                    dialog.resizable(False, False)
                    dialog.grab_set()

                    # Центрируем окно
                    dialog.update_idletasks()
                    x = (dialog.winfo_screenwidth() // 2) - (400 // 2)
                    y = (dialog.winfo_screenheight() // 2) - (150 // 2)
                    dialog.geometry(f"+{x}+{y}")

                    # Текст сообщения
                    message = tk.Label(
                        dialog,
                        text=f"Файл '{out_path.name}' открыт в другой программе.\n\n"
                             f"Закройте файл и нажмите 'Повторить'\n"
                             f"или сохраните как '{new_path.name}'",
                        justify="left",
                        padx=20,
                        pady=20
                    )
                    message.pack()

                    # Переменная для хранения выбора
                    user_choice = tk.StringVar(value="")

                    # Кнопки
                    button_frame = tk.Frame(dialog)
                    button_frame.pack(pady=10)

                    retry_btn = tk.Button(
                        button_frame,
                        text="🔄 Повторить",
                        width=15,
                        command=lambda: [user_choice.set("retry"), dialog.destroy()]
                    )
                    retry_btn.pack(side="left", padx=5)

                    rename_btn = tk.Button(
                        button_frame,
                        text="💾 Переименовать",
                        width=15,
                        command=lambda: [user_choice.set("rename"), dialog.destroy()]
                    )
                    rename_btn.pack(side="left", padx=5)

                    # Ждем выбора пользователя
                    dialog.wait_window()

                    if user_choice.get() == "retry":
                        # Повторяем попытку с тем же именем
                        continue
                    else:
                        # Сохраняем с новым именем
                        current_path = new_path
                        continue

            except Exception as e:
                messagebox.showerror(
                    "Ошибка сохранения Excel",
                    f"Не удалось сохранить файл: {str(e)}"
                )
                raise
        
        # Если все попытки исчерпаны
        messagebox.showerror(
            "Ошибка сохранения",
            f"Не удалось сохранить файл после {max_attempts} попыток. Пожалуйста, закройте файл вручную."
        )
        raise PermissionError(f"Невозможно сохранить {current_path}: файл остаётся занят")

    def process_path(self, path: str, progress_callback=None, save_excel: bool = True) -> tuple[pd.DataFrame, pd.DataFrame, str, pd.DataFrame]:
        """Главная функция обработки"""
        
        all_data = []
        self.reset_stats()
        self.stats["start_time"] = time.time()  # Засекаем время начала
        self.apply_config(self.load_config("config.yaml"))  # Загружаем свежий конфиг
        path_obj = Path(path)

        if path_obj.is_file() and path_obj.suffix.lower() == ".pdf":
            self.process_pdf(str(path_obj), all_data)
            if progress_callback:
                progress_callback(1, 1, str(path_obj))
            base_name = path_obj.stem
            out_dir = path_obj.parent

        elif path_obj.is_dir():
            base_name = path_obj.name
            out_dir = path_obj
            pdf_files = sorted(path_obj.rglob("*.pdf"))
            self.stats["files_skipped"] = len([
                f for f in path_obj.rglob("*") if f.is_file() and f.suffix.lower() != ".pdf"
            ])

            # Обрабатываем каждый PDF файл в папке с обновлением прогресса
            for idx, pdf_path in enumerate(pdf_files, start=1):
                self.process_pdf(str(pdf_path), all_data)
                # Обновляем прогресс
                if progress_callback:
                    progress_callback(idx, len(pdf_files), str(pdf_path))

        else:
            raise ValueError(f"'{path}' не является PDF или папкой")

        if not all_data:
            raise ValueError("PDF файлы не найдены")

        df = pd.DataFrame(all_data)

        # СВОДКА
        def get_page_list(group: pd.DataFrame, color: str) -> str:
            pages = sorted(group[group["Цветность"] == color]["Страница"].tolist())
            return ", ".join(map(str, pages)) if pages else "-"

        summary_data = []
        for (format_name, std_size, file_name), group in df.groupby(["Стандартный формат", "Размер стандарта", "Файл"]):
            pages_list_bw = get_page_list(group, "Ч/Б")
            pages_list_col = get_page_list(group, "Цветная")

            if self.compress_ranges_y:
                if pages_list_bw != "-": pages_list_bw = self.compress_ranges(pages_list_bw)
                if pages_list_col != "-": pages_list_col = self.compress_ranges(pages_list_col)

            summary_data.append({
                "Файл": file_name,
                "Стандартный формат": format_name,
                "Количество": len(group),
                "Размер стандарта": std_size,
                "Ч/Б страницы": pages_list_bw,
                "Цветные страницы": pages_list_col,
                "Цветных": (group["Цветность"] == "Цветная").sum(),
                "Ч/Б": (group["Цветность"] == "Ч/Б").sum(),
            })

        # Добавляем строку для объединенных форматов A4+A3 по каждому файлу
        for file_name, file_group in df.groupby("Файл"):
            a4a3_group = file_group[file_group["Стандартный формат"].isin(["A4", "A3"])]
            if not a4a3_group.empty:
                pages_list_bw = get_page_list(a4a3_group, "Ч/Б")
                pages_list_col = get_page_list(a4a3_group, "Цветная")

                if self.compress_ranges_y:
                    if pages_list_bw != "-": pages_list_bw = self.compress_ranges(pages_list_bw)
                    if pages_list_col != "-": pages_list_col = self.compress_ranges(pages_list_col)

                summary_data.append({
                    "Файл": file_name,
                    "Стандартный формат": "A4 + A3",
                    "Количество": len(a4a3_group),
                    "Размер стандарта": "-",
                    "Ч/Б страницы": pages_list_bw,
                    "Цветные страницы": pages_list_col,
                    "Цветных": (a4a3_group["Цветность"] == "Цветная").sum(),
                    "Ч/Б": (a4a3_group["Цветность"] == "Ч/Б").sum(),
                })

        summary = pd.DataFrame(summary_data).sort_values(["Файл", "Стандартный формат"]).reset_index(drop=True)

        # СВОДКА ПО РУЛОНАМ (исключая A4)
        roll_summary = self.build_roll_summary(df)

        out_path = out_dir / f"{base_name}_all_sizes.xlsx"

        # Безопасное сохранение Excel с обработкой ошибок при занятом файле
        saved_path = str(out_path)
        if save_excel:
            saved_path = self.save_excel_safe(out_path, df, summary, roll_summary)

        self.stats["end_time"] = time.time()  # Засекаем время конца обработки
        return df, summary, saved_path, roll_summary

    def build_report_text(self, df: pd.DataFrame, summary: pd.DataFrame, out_path: str, roll_summary: pd.DataFrame | None) -> str:
        """Формирует текстовый отчёт для вывода в GUI/копирования."""
        
        lines: list[str] = []

        # Заголовок и общая статистика
        lines.append(f"=== ОТЧЁТ АНАЛИЗА PDF ===  Дата: {pd.Timestamp.now().strftime('%d.%m.%Y %H:%M')}")
        report_dir = str(Path(out_path).parent)
        unique_files = int(df["Файл"].nunique()) if not df.empty and "Файл" in df.columns else 0
        lines.append(f"Файл отчёта: {out_path}")
        lines.append(f"Каталог отчёта: {report_dir}")
        lines.append(f"PDF-файлов в выборке: {unique_files}")
        lines.append("")

        lines.append("📊 ОБЩАЯ СТАТИСТИКА:")
        lines.append(f"    Файлов обработано: {self.stats.get('files_processed', 0)}")
        lines.append(f"    Листов обработано: {self.stats.get('pages_processed', 0)}")
        lines.append(f"    Файлов пропущено (не PDF): {self.stats.get('files_skipped', 0)}")
        
        # Вычисляем время обработки
        if self.stats.get('start_time') and self.stats.get('end_time'):
            elapsed_time = self.stats['end_time'] - self.stats['start_time']
            minutes = int(elapsed_time // 60)
            seconds = int(elapsed_time % 60)
            time_str = f"{minutes}м {seconds}с" if minutes > 0 else f"{seconds}с"
            lines.append(f"    ⏱️ Время обработки: {time_str}")
        
        #lines.append(f"    Допуск распознавания форматов: {self.tolerance} мм")
        #lines.append(f"    Стандартных форматов в базе: {len(self.formats)}")
        lines.append("")

        # 📁 Статистика по файлам
        lines.append("📁 СТАТИСТИКА ПО ФАЙЛАМ:")
        file_stats = df.groupby("Файл").agg({
            "Страница": "count",
            "Цветность": lambda x: (x == "Цветная").sum()
        }).reset_index()

        for _, row in file_stats.iterrows():
            file_name = row["Файл"]
            total_pages = int(row["Страница"])
            color_pages = int(row["Цветность"])
            bw_pages = total_pages - color_pages

            pages = df[df["Файл"] == file_name]["Страница"].tolist()
            if pages:
                page_range = f"{min(pages)}–{max(pages)}"
            else:
                page_range = "-"

            lines.append(
                f"    {file_name}: {total_pages} стр. "
                f"(Ч/Б: {bw_pages}, цветных: {color_pages}), диапазон: {page_range}"
            )

        lines.append("")

        # 📐 Форматы по файлам (как вы просили)
        lines.append("📐 ФОРМАТЫ ПО ФАЙЛАМ:")

        for file_name, file_group in df.groupby("Файл"):
            lines.append(f"\n    📄 {file_name}:")

            # --- ОТДЕЛЬНЫЕ ФОРМАТЫ ВНУТРИ ФАЙЛА ---
            file_formats = file_group.groupby("Стандартный формат").agg({
                "Страница": "count"
            }).reset_index()

            file_format_details = file_group.groupby("Стандартный формат")["Страница"].apply(
                lambda x: ",".join(map(str, sorted(x.tolist())))
            ).to_dict()

            for _, frow in file_formats.iterrows():
                fmt = frow["Стандартный формат"]          # A4, A3, A3×4, Custom1 ...
                total_pages = int(frow["Страница"])
                # print(f"self.compress_ranges_y = {self.compress_ranges_y}")
                if self.compress_ranges_y:
                    pages_list = self.compress_ranges(file_format_details.get(fmt, "-"))
                else:
                    pages_list = file_format_details.get(fmt, "-")                
                
                sample_size = file_group[file_group["Стандартный формат"] == fmt]["Размер стандарта"].iloc[0]

                # Формат строки: "A4 210x297 (45 стр.): 1,2,3,..."
                lines.append(f"        {fmt} {sample_size} ({total_pages} стр.): {pages_list}")
                
            # --- СУММАРНАЯ СТРОКА ПО A4 + A3 ---
            a4a3_group = file_group[file_group["Стандартный формат"].isin(["A4", "A3"])]
            if not a4a3_group.empty:
                total_a4a3 = int(a4a3_group["Страница"].count())
                color_a4a3 = int((a4a3_group["Цветность"] == "Цветная").sum())
                
                if self.compress_ranges_y:
                    pages_a4a3 = self.compress_ranges(",".join(map(str, sorted(a4a3_group["Страница"].tolist()))))
                else:
                    pages_a4a3 = ",".join(map(str, sorted(a4a3_group["Страница"].tolist())))                 
               
                lines.append(f"        A4 + A3 ({total_a4a3} стр.): {pages_a4a3}")

        lines.append("")

        # 🧻 СВОДКА ПО РУЛОНАМ
        if roll_summary is not None and not roll_summary.empty:
            lines.append("🧻 СВОДКА ПО РУЛОНАМ:")

            # Фильтруем итоговые строки (они только для Excel)
            roll_summary_filtered = roll_summary[~roll_summary["Стандартный формат"].str.contains("ИТОГО", na=False)].copy()

            # Группировка сначала по файлу, потом по формату рулона
            for file_name, file_group in roll_summary_filtered.groupby("Файл"):
                lines.append(f"\n    📄 {file_name}:")

                for roll_fmt, roll_group in file_group.groupby("Формат рулона"):
                    # Пропускаем "Нестандартный" рулон
                    if roll_fmt == "Нестандартный":
                        continue

                    total_in_roll = int(roll_group["Количество"].sum())

                    # Собираем все страницы для этого рулона (расширяя диапазоны)
                    all_pages_for_roll = []
                    for _, row in roll_group.iterrows():
                        bw_pages = row["Ч/Б страницы"]
                        all_pages_for_roll.extend(self.expand_ranges(bw_pages))

                        col_pages = row["Цветные страницы"]
                        all_pages_for_roll.extend(self.expand_ranges(col_pages))

                    # Сортируем и форматируем список страниц
                    all_pages_for_roll = sorted(set(all_pages_for_roll))
                    if self.compress_ranges_y and all_pages_for_roll:
                        pages_for_roll_str = self.compress_ranges(",".join(map(str, all_pages_for_roll)))
                    else:
                        pages_for_roll_str = ",".join(map(str, all_pages_for_roll)) if all_pages_for_roll else "-"

                    lines.append(f"        🧻 {roll_fmt}: {total_in_roll} стр.: {pages_for_roll_str}")

                    for _, row in roll_group.iterrows():
                        fmt = row["Стандартный формат"]
                        size = row["Размер стандарта"]
                        count = int(row["Количество"])
                        bw_pages = row["Ч/Б страницы"]
                        col_pages = row["Цветные страницы"]

                        # Объединяем ч/б и цветные страницы (расширяя диапазоны)
                        fmt_pages = self.expand_ranges(bw_pages) + self.expand_ranges(col_pages)
                        fmt_pages = sorted(set(fmt_pages))

                        if self.compress_ranges_y and fmt_pages:
                            fmt_pages_str = self.compress_ranges(",".join(map(str, fmt_pages)))
                        else:
                            fmt_pages_str = ",".join(map(str, fmt_pages)) if fmt_pages else "-"

                        lines.append(
                            f"                {fmt} {size} ({count} стр.): {fmt_pages_str}"
                        )

            lines.append("")

        # Ошибки (если были)
        errors = self.stats.get("errors", [])
        if errors:
            lines.append(f"❌ ОШИБКИ ({len(errors)}):")
            for i, err in enumerate(errors, 1):
                lines.append(f"  {i}. {err}")
        else:
            lines.append("✅ Ошибок не обнаружено")

        return "\n".join(lines)

    def calculate_roll_metrics(self, df: pd.DataFrame, roll_summary: pd.DataFrame | None) -> Dict[str, Dict[str, float]]:
        """Рассчитывает метрики для каждого типа рулона: площадь и длину.

        Возвращает словарь вида:
        {
            "420мм": {"area_m2": 12.5, "length_m": 30.0, "sheets_count": 15},
            "594мм": {"area_m2": 8.3, "length_m": 14.0, "sheets_count": 7},
            ...
        }
        """
        if roll_summary is None or roll_summary.empty:
            return {}

        # Фильтруем только строки форматов (не итоговые строки рулонов)
        roll_data = roll_summary[roll_summary["Стандартный формат"] != "ИТОГО"].copy()

        if roll_data.empty:
            return {}

        metrics = {}

        # Группируем по формату рулона
        for roll_format, group in roll_data.groupby("Формат рулона"):
            if roll_format == "Отдельный лист":
                continue

            total_area_m2 = 0.0
            total_length_m = 0.0
            total_sheets = 0

            for _, row in group.iterrows():
                count = row["Количество"]
                size_str = row["Размер стандарта"]

                # Парсим размер (например "594x420")
                try:
                    parts = size_str.split('x')
                    if len(parts) == 2:
                        w_mm = float(parts[0])
                        h_mm = float(parts[1])

                        # Площадь одного листа в м²
                        area_one_m2 = (w_mm * h_mm) / 1_000_000
                        total_area_m2 += area_one_m2 * count

                        # Длина = большая сторона
                        length_one_m = max(w_mm, h_mm) / 1000
                        total_length_m += length_one_m * count

                        total_sheets += count
                except (ValueError, IndexError):
                    continue

            if total_sheets > 0:
                metrics[roll_format] = {
                    "area_m2": round(total_area_m2, 2),
                    "length_m": round(total_length_m, 2),
                    "sheets_count": total_sheets
                }

        return metrics

# ================== TOOLBAR ==================
class Toolbar(ttk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app

    def add_button(self, *, icon, command, tooltip=None, state="normal", text_name=None, side_str = tk.LEFT):
        btn = ttk.Button(
            self,
            image=self.app.icons[icon],
            command=command,
            state=state,
        )
        if text_name:
            btn.configure(text=" "+text_name+" ", compound="left")

        btn.pack(side=side_str, padx=2) # type: ignore

        if tooltip:
            self.app.tooltip(btn, tooltip)

        return btn

    def add_separator(self, side_str=tk.LEFT):
        ttk.Separator(self, orient="vertical").pack(
            side=side_str, fill="y", padx=5 # type: ignore
        )

class MainWindow:
    
    def __init__(self, analyzer, initial_result=None):
        """
        initial_result:
          - None → обычный режим (ждём, пока пользователь нажмёт Открыть файл/папку)
          - (df, summary, out_path) → показываем уже готовый отчёт (CLI‑режим)
        """
        self.analyzer = analyzer
        if TkinterDnD is not None:
            self.root = TkinterDnD.Tk()
        else:
            self.root = tk.Tk()
        
        # Загружаем иконку с fallback
        try:
            icon_path = self.resource_path("img/icon.png")
            icon = tk.PhotoImage(file=icon_path)
            self.root.iconphoto(True, icon)
        except Exception:
            # Если иконка не найдена, продолжаем без неё
            pass
       
        self.root.title("Анализатор PDF файлов "+version)
        self.root.geometry("900x650")
        self.root.update_idletasks()
        x = max(0, (self.root.winfo_screenwidth() - self.root.winfo_width()) // 2)
        y = max(0, (self.root.winfo_screenheight() - self.root.winfo_height()) // 2)
        self.root.geometry(f"+{x}+{y}")
        self.root.resizable(True, True)

        # Иконки
        self.load_icons()

        self.last_result = initial_result  # (df, summary, out_path) или None

        self._build_ui()

        # если уже есть результат (CLI-сценарий) — сразу показываем его
        if self.last_result is not None:
            df, summary, out_path, roll_summary = self.last_result
            report_text = self.analyzer.build_report_text(df, summary, out_path, roll_summary)
            self._set_stats_text(report_text)
            self._update_roll_visualization(df, roll_summary)
            self.treeview_tab.update_data(df=df, summary=summary, roll_summary=roll_summary)  # Обновляем древовидную вкладку
        else:
            self._set_stats_text("Выберите файл или папку для анализа PDF документов.")

    # ---------- построение интерфейса ----------

    def load_icons(self):
        icon_names = [
            "pdf_file", "folder", "settings", "help", "copy", "save", "order", "cur_settings", "formats16", "margin16", "order16", "icon", "icon48",
            "pdf_file16", "plotter16", "roll16", "sheet16", "laser16", "bw16", "color16",
            "openfolder", "formats", "margin","select","txt_report","exit", "xlsx", "diagram", "report", "diagram16", "report16", "text-file"
        ]
        self.icons = {
            name: tk.PhotoImage(
                file=self.resource_path(f"img/{name}.png")
            )
            for name in icon_names
        }

    def tooltip(self, widget, text, delay=600):
        tip = tk.Toplevel(widget)
        tip.withdraw()
        tip.overrideredirect(True)

        label = ttk.Label(
            tip,
            text=text,
            background="#ffffe0",
            relief="solid",
            borderwidth=1,
            padding=(6, 3)
        )
        label.pack()

        after_id = None

        def show():
            x = widget.winfo_rootx() + 20
            y = widget.winfo_rooty() + widget.winfo_height() + 5
            tip.geometry(f"+{x}+{y}")
            tip.deiconify()

        def on_enter(event):
            nonlocal after_id
            after_id = widget.after(delay, show)

        def on_leave(event):
            nonlocal after_id
            if after_id:
                widget.after_cancel(after_id)
                after_id = None
            tip.withdraw()

        widget.bind("<Enter>", on_enter)
        widget.bind("<Leave>", on_leave)

    def image_label(self, parent, icon, ttext: str="", im_padx=2, im_pady=0, im_side = 'left' ):
        lbl = tk.Label(parent, image=self.icons[icon], bd=0)
        if ttext:
            lbl.configure(text=ttext, compound="left")
        #lbl.image = self.icons['filter']   # важно! чтобы не удалилось сборщиком мусора
        lbl.pack(side=im_side, padx=im_padx, pady=im_pady)   # type: ignore

    def _build_ui(self):
        root = self.root

        # Верхняя панель: статус конфигурации + кнопки
        top_frame = ttk.Frame(root, padding=10)
        top_frame.pack(fill=tk.X)

        # Панель кнопок сверху слева
        top_toolbar = Toolbar(top_frame, self)
        top_toolbar.pack(side=tk.LEFT, anchor=tk.NW, fill=tk.X, expand=True, padx=0, pady=0)

        top_toolbar.add_button(icon="pdf_file", command=self.on_open_file, tooltip="Открыть PDF файл")
        top_toolbar.add_button(icon="openfolder", command=self.on_open_folder, tooltip="Открыть папку")
        top_toolbar.add_separator()
        top_toolbar.add_button(icon="help", command=self.show_help_window, tooltip="Помощь <F1>", side_str='right')
        top_toolbar.add_button(icon="settings", command=self.open_config_editor, tooltip="Настройки", side_str='right')
       
        self.select_all_button = top_toolbar.add_button(
            icon="select",
            command=self.select_all,
            tooltip="Выделить всё",
            state="disabled",
        )
        self.copy_report_button = top_toolbar.add_button(
            icon="copy",
            command=self.copy_report,
            tooltip="Копировать отчёт",
            state="disabled",
        )
        top_toolbar.add_separator()
        top_toolbar.add_button(icon="xlsx", command=self.open_excel_report, tooltip="Excel отчет")
        top_toolbar.add_button(icon="folder", command=self.open_reports_folder, tooltip="Папка отчетов")
        top_toolbar.add_button(icon="save", command=self.save_report_to_file, tooltip="Сохранить текстовый отчёт")
        # top_toolbar.add_button(icon="exit", command=root.destroy, tooltip="Выход", side_str='right')

        # Центральная область — вкладки
        center_frame = ttk.Frame(root)
        center_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        # Создаем Notebook для вкладок
        self.notebook = ttk.Notebook(center_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        # Вкладка 1: Сводка по форматам (древовидная таблица)
        self.treeview_tab = TreeviewTab(self.notebook, font_face=font_face, app=self)
        self.notebook.add(self.treeview_tab.frame, text=" Сводка по форматам", image=self.icons["report16"], compound="left")

        # Вкладка 2: Визуализация рулонов
        viz_tab = ttk.Frame(self.notebook)
        self.notebook.add(viz_tab, text=" Загрузка рулонов", image=self.icons["diagram16"], compound="left")

        # Вкладка 3: Текстовый отчет
        report_tab = ttk.Frame(self.notebook)
        self.notebook.add(report_tab, text=" Отчёт", image=self.icons["report16"], compound="left")
        self.report_tab = report_tab
        self.notebook.bind("<<NotebookTabChanged>>", self._update_report_toolbar_state)

        self.stats_text = tk.Text(
            report_tab,
            wrap=tk.WORD,
            font=(font_face, 10),
            bg="#f8f9fa",
            bd=1,
            selectbackground="#405361",
            selectforeground="white",
            padx=10,
            pady=10,
        )

        scroll_report = ttk.Scrollbar(report_tab, orient=tk.VERTICAL, command=self.stats_text.yview)
        self.stats_text.configure(yscrollcommand=scroll_report.set)
        scroll_report.pack(side=tk.RIGHT, fill=tk.Y)
        self.stats_text.pack(fill=tk.BOTH, expand=True)
        self._update_report_toolbar_state()

        # Canvas с прокруткой для графиков
        self.viz_canvas = tk.Canvas(viz_tab, bg="#f8f9fa")
        viz_scrollbar = ttk.Scrollbar(viz_tab, orient=tk.VERTICAL, command=self.viz_canvas.yview)

        # Контейнер для графиков внутри Canvas
        self.viz_frame = ttk.Frame(self.viz_canvas)

        # Создаем окно внутри Canvas
        self.viz_canvas_window = self.viz_canvas.create_window((0, 0), window=self.viz_frame, anchor="nw")

        # Настраиваем прокрутку
        self.viz_canvas.configure(yscrollcommand=viz_scrollbar.set)

        # Упаковываем элементы
        viz_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.viz_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Обновление размеров при изменении содержимого
        self.viz_frame.bind("<Configure>", lambda e: self.viz_canvas.configure(scrollregion=self.viz_canvas.bbox("all")))

        # Обновление ширины окна при изменении размера Canvas
        self.viz_canvas.bind("<Configure>", lambda e: self.viz_canvas.itemconfig(self.viz_canvas_window, width=e.width))

        # Прокрутка колесиком мыши
        def _on_mousewheel(event):
            self.viz_canvas.yview_scroll(int(-1*(event.delta/120)), "units")

        self.viz_canvas.bind_all("<MouseWheel>", _on_mousewheel)

        # Placeholder для графиков
        self.viz_placeholder = ttk.Label(
            self.viz_frame,
            text="Выполните анализ для просмотра визуализации загрузки рулонов",
            font=(font_face, 12),
            foreground="#666"
        )
        self.viz_placeholder.pack(expand=True)

        #Статус бар
        status_frame = ttk.Frame(root, padding=2)
        status_frame.pack(fill=tk.X)
        
        self.status_label = ttk.Label(status_frame, text="Готов", font=(font_face, 10), justify=tk.LEFT, relief=tk.SUNKEN, borderwidth=1)
        self.status_label.pack(side=tk.LEFT, padx=5, pady=5, anchor=tk.W, fill=tk.BOTH, expand=True)

        self.image_label(status_frame, icon="margin16", im_padx=5)
        self.analyzer.tolerance_var = tk.StringVar()
        self.analyzer.tolerance_var.set("...")
        self.tolerance_status_entry = tk.Entry(status_frame, 
                                       textvariable=self.analyzer.tolerance_var,
                                       font=(font_face, 10),
                                       background="#ffffff",
                                       width=4,
                                       justify=tk.CENTER, 
                                       state="readonly")
        self.tolerance_status_entry.pack(side=tk.LEFT, anchor=tk.W, fill=tk.Y, padx=0, pady=5)        
        self.tooltip(self.tolerance_status_entry, "Допуск распознавания форматов (мм)")

        self.image_label(status_frame, icon="order16", im_padx=5)
        self.analyzer.compress_ranges_var= tk.StringVar()
        self.analyzer.compress_ranges_var.set("...")
        self.compress_status_entry = tk.Entry(status_frame, 
                                       textvariable=self.analyzer.compress_ranges_var,
                                       font=(font_face, 10),
                                       background="#ffffff",
                                       width=5,
                                       justify=tk.CENTER, 
                                       state="readonly")
        self.compress_status_entry.pack(side=tk.LEFT, anchor=tk.W, fill=tk.Y, padx=0, pady=5)
        self.tooltip(self.compress_status_entry, "Сжатие диапазонов страниц")
                
        self.image_label(status_frame, icon="formats16", im_padx=5)
        self.analyzer.formats_count_var = tk.StringVar()
        self.analyzer.formats_count_var.set("...")
        self.formats_count_status_entry = tk.Entry(status_frame, 
                                       textvariable=self.analyzer.formats_count_var,
                                       font=(font_face, 10),
                                       background="#ffffff",
                                       width=4,
                                       justify=tk.CENTER,
                                       state="readonly")
        self.formats_count_status_entry.pack(side=tk.LEFT, anchor=tk.W, fill=tk.Y, padx=0, pady=5)
        self.tooltip(self.formats_count_status_entry, "Количество распознаваемых форматов")

        # Реальный прогресс-бар (determinate)
        self.progress = ttk.Progressbar(status_frame, mode='determinate', maximum=100, length=200)
        self.progress.pack(pady=5, side=tk.RIGHT, anchor=tk.E, padx=5)
        self.progress_value = 0
        self.progress_max = 0

        # Обновляем статус в интерфейсе (после инициализации stats_text)
        self.refresh_config()

        if TkinterDnD is not None:
            self.stats_text.drop_target_register("DND_Files")
            self.stats_text.dnd_bind("<<Drop>>", self._on_drop)
            self.stats_text.dnd_bind("<<DropEnter>>", lambda event: self._set_drop_zone_active())
            self.stats_text.dnd_bind("<<DropLeave>>", lambda event: self._set_drop_zone_normal())
            # Приём PDF также непосредственно на вкладке сводки и её Treeview.
            for target in (self.treeview_tab.frame, self.treeview_tab.tree):
                target.drop_target_register("DND_Files")
                target.dnd_bind("<<Drop>>", self._on_drop)
                target.dnd_bind("<<DropEnter>>", lambda event: self._set_drop_zone_active())
                target.dnd_bind("<<DropLeave>>", lambda event: self._set_drop_zone_normal())
            print("✅ Drag & Drop активен")
        else:
            print("⚠️ Drag & Drop недоступен (tkinterdnd2 не установлен)")
            print("   Установите: pip install tkinterdnd2")
            # Fallback для визуальной индикации при наведении мыши
            self.stats_text.bind("<Enter>", lambda event: self._set_drop_zone_active())
            self.stats_text.bind("<Leave>", lambda event: self._set_drop_zone_normal())

        # Контекстное меню (ПКМ)
        self.context_menu = tk.Menu(self.stats_text, tearoff=0)
        self.context_menu.add_command(label="Копировать", command=self.copy_selection)
        self.context_menu.add_command(label="Выделить всё", command=self.select_all)
        
        def show_context_menu(event):
            try:
                self.context_menu.tk_popup(event.x_root, event.y_root)
            finally:
                self.context_menu.grab_release()
        
        self.stats_text.bind("<Button-3>", show_context_menu)
        
        # Горячие клавиши на самом виджете stats_text (работают при любой раскладке)
        self.stats_text.bind("<Control-a>", lambda e: (self.select_all(), "break")[1])
        self.stats_text.bind("<Control-c>", lambda e: (self.copy_selection(), "break")[1])
        
        # Горячие клавиши на root (общие для приложения)
        self.root.bind("<Control-s>", lambda e: self.save_report_to_file())
        self.root.bind("<Escape>", lambda e: root.destroy())
        self.root.bind("<F1>", lambda e: self.show_help_window())

        self.stats_text.focus_set() # Фокус на тексовый отчет

    def _update_report_toolbar_state(self, event=None):
        """Включает кнопки работы с текстом только на вкладке отчёта."""
        if not hasattr(self, "notebook") or not hasattr(self, "report_tab"):
            return

        report_active = self.notebook.select() == str(self.report_tab)
        state = "normal" if report_active else "disabled"
        self.select_all_button.configure(state=state)
        self.copy_report_button.configure(state=state)

    def _set_drop_zone_normal(self):
        try:
            self.stats_text.configure(bg="#f8f9fa", highlightbackground="#cbd5e1", highlightthickness=0)
        except Exception:
            pass
        try:
            self.treeview_tab.tree.configure(style="Treeview")
        except Exception:
            pass
        try:
            self.notebook.tab(self.report_tab, state="normal")
        except Exception:
            pass

    def _set_drop_zone_active(self):
        try:
            self.stats_text.configure(bg="#dbeafe", highlightbackground="#A0DAB1", highlightthickness=2)
        except Exception:
            pass
        try:
            style_name = "DropZone.Treeview"
            style = ttk.Style(self.root)
            style.configure(style_name, background="#dbeafe")
            self.treeview_tab.tree.configure(style=style_name)
        except Exception:
            pass
        try:
            self.notebook.tab(self.report_tab, state="active")
        except Exception:
            pass

    def _extract_drop_paths(self, raw_data) -> list[str]:
        if raw_data is None:
            return []

        if isinstance(raw_data, (list, tuple, set)):
            paths: list[str] = []
            for item in raw_data:
                paths.extend(self._extract_drop_paths(item))
            return paths

        text = str(raw_data).strip()
        if not text:
            return []

        text = text.replace("\\n", "\n")
        text = text.replace("}\n{", "\n").replace("}\r\n{", "\n")
        text = text.replace("}{", "\n").replace(";", "\n")
        text = text.replace("\r\n", "\n").replace("\r", "\n")
        text = text.replace("\x00", "\n")

        matches = re.findall(r"(?i)(?:[A-Za-z]:)?(?:[^\r\n;{}\[\]()\"']+?\.pdf)", text)
        if matches:
            return [match.strip().strip("'\"{}[]() ") for match in matches if match.strip().strip("'\"{}[]() ")]

        fallback = []
        for block in re.split(r"\n+", text):
            block = block.strip().strip("'\"{}[]() ")
            if block:
                fallback.append(block)
        return fallback

    def _on_drop(self, event):
        raw_data = getattr(event, "data", event)
        paths = self._extract_drop_paths(raw_data)
        if not paths:
            return

        self._set_drop_zone_active()
        try:
            self._run_analysis(paths)
        except ValueError as e:
            messagebox.showwarning("Предупреждение", str(e))
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка обработки:\n{e}")
        finally:
            self._set_drop_zone_normal()

    def _normalize_input_paths(self, path):
        """Нормализует входные пути из drag&drop или выбора файла/папки."""
        if path is None:
            return []

        if isinstance(path, (list, tuple, set)):
            items = []
            for item in path:
                items.extend(self._normalize_input_paths(item))
            return items

        text = str(path).strip()
        if not text:
            return []

        text = text.replace("\\n", "\n")
        text = text.replace("}\n{", "\n").replace("}\r\n{", "\n")
        text = text.replace("}{", "\n").replace(";", "\n")
        text = text.replace("\r\n", "\n").replace("\r", "\n")
        text = text.replace("\x00", "\n")

        matches = re.findall(r"(?i)(?:[A-Za-z]:)?(?:[^\r\n;{}\[\]()\"']+?\.pdf)", text)
        if matches:
            return [match.strip().strip("'\"{}[]() ") for match in matches if match.strip().strip("'\"{}[]() ")]

        fallback = []
        for block in re.split(r"\n+", text):
            block = block.strip().strip("'\"{}[]() ")
            if block:
                fallback.append(block)
        return fallback

    def _set_stats_text(self, text: str):
        self.stats_text.config(state=tk.NORMAL)
        self.stats_text.delete("1.0", tk.END)
        self.stats_text.insert(tk.END, text)
        self.stats_text.config(state=tk.DISABLED)

    def _update_roll_visualization(self, df: pd.DataFrame, roll_summary: pd.DataFrame):
        """Обновляет визуализацию загрузки рулонов."""

        # Очищаем старые графики
        for widget in self.viz_frame.winfo_children():
            widget.destroy()

        if roll_summary is None or roll_summary.empty:
            no_data_label = ttk.Label(
                self.viz_frame,
                text="Нет данных по рулонам для визуализации",
                font=(font_face, 12),
                foreground="#999"
            )
            no_data_label.pack(expand=True)
            return

        # Рассчитываем метрики
        metrics = self.analyzer.calculate_roll_metrics(df, roll_summary)

        if not metrics:
            no_data_label = ttk.Label(
                self.viz_frame,
                text="Нет данных по рулонам для визуализации",
                font=(font_face, 12),
                foreground="#999"
            )
            no_data_label.pack(expand=True)
            return

        # Создаем фигуру с двумя круговыми диаграммами
        fig = Figure(figsize=(12, 5), dpi=100)
        fig.patch.set_facecolor('#f8f9fa')

        # Диаграмма 1: Площадь печати
        ax1 = fig.add_subplot(121)
        roll_names = list(metrics.keys())
        areas = [metrics[roll]["area_m2"] for roll in roll_names]

        colors = ['#38BDF8', '#22D3EE', '#4FD1C5', '#F97316', '#F59E0B']

        # Убираем эмодзи из меток для графика (чтобы избежать предупреждений)
        labels_without_emoji = [roll.replace("🧻 ", "") for roll in roll_names]

        wedges1, texts1, autotexts1 = ax1.pie( # type: ignore
            areas,
            labels=[f"{label}\n{metrics[roll_names[i]]['sheets_count']} листов"
                    for i, label in enumerate(labels_without_emoji)],
            autopct='%1.1f%%',
            startangle=90,
            colors=colors[:len(roll_names)],
            textprops={'fontsize': 9, 'color': '#1E293B'}
        )

        for autotext in autotexts1:
            autotext.set_color('white')
            autotext.set_fontweight('bold')
            autotext.set_fontsize(10)

        ax1.set_title('Распределение площади печати по рулонам',
                     fontsize=12, fontweight='bold', color='#0F172A', pad=20)

        # Добавляем легенду с площадями
        legend_labels1 = [f"{label}: {metrics[roll_names[i]]['area_m2']} м²"
                          for i, label in enumerate(labels_without_emoji)]
        ax1.legend(legend_labels1, loc='upper left', bbox_to_anchor=(0, -0.1),
                  fontsize=9, frameon=False)

        # Диаграмма 2: Длина рулона
        ax2 = fig.add_subplot(122)
        lengths = [metrics[roll]["length_m"] for roll in roll_names]

        wedges2, texts2, autotexts2 = ax2.pie( # type: ignore
            lengths,
            labels=[f"{label}\n{metrics[roll_names[i]]['sheets_count']} листов"
                    for i, label in enumerate(labels_without_emoji)],
            autopct='%1.1f%%',
            startangle=90,
            colors=colors[:len(roll_names)],
            textprops={'fontsize': 9, 'color': '#1E293B'}
        )

        for autotext in autotexts2:
            autotext.set_color('white')
            autotext.set_fontweight('bold')
            autotext.set_fontsize(10)

        ax2.set_title('Распределение длины рулонов',
                     fontsize=12, fontweight='bold', color='#0F172A', pad=20)

        # Добавляем легенду с длинами
        legend_labels2 = [f"{label}: {metrics[roll_names[i]]['length_m']} м"
                          for i, label in enumerate(labels_without_emoji)]
        ax2.legend(legend_labels2, loc='upper left', bbox_to_anchor=(0, -0.1),
                  fontsize=9, frameon=False)

        fig.tight_layout(pad=3.0)

        # Встраиваем в Tkinter
        canvas = FigureCanvasTkAgg(fig, master=self.viz_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        # Добавляем итоговую информацию внизу
        info_frame = ttk.Frame(self.viz_frame)
        info_frame.pack(fill=tk.X, pady=10)

        total_area = sum(areas)
        total_length = sum(lengths)
        total_sheets = sum(metrics[roll]['sheets_count'] for roll in roll_names)

        info_text = f"📊 ИТОГО:  {total_sheets} листов  |  {total_area:.2f} м²  |  {total_length:.2f} м"
        info_label = ttk.Label(info_frame, text=info_text, font=(font_face, 11, 'bold'),
                              foreground='#0F172A')
        info_label.pack()

        # Обновляем область прокрутки после добавления графиков
        self.viz_frame.update_idletasks()
        self.viz_canvas.configure(scrollregion=self.viz_canvas.bbox("all"))

    # ---------- обработчики кнопок ----------

    def on_open_file(self):
        path = filedialog.askopenfilename(
            title="Выберите PDF-файл",
            filetypes=[("PDF files", "*.pdf")]
        )
        if not path:
            return
        self._run_analysis(path)

    def on_open_folder(self):
        path = filedialog.askdirectory(
            title="Выберите папку с PDF-файлами"
        )
        if not path:
            return
        self._run_analysis(path)

    def _run_analysis(self, path):
        clean_paths = self._normalize_input_paths(path)
        if not clean_paths:
            return

        pdf_files = []
        folders = []
        for item in clean_paths:
            p = Path(item)
            if p.is_dir():
                folders.append(str(p))
            elif p.suffix.lower() == ".pdf":
                pdf_files.append(str(p))

        if not pdf_files and not folders:
            return

        if folders:
            recursive_pdf_files = []
            for folder in folders:
                recursive_pdf_files.extend(str(p) for p in Path(folder).rglob("*.pdf") if p.is_file())
            pdf_files = list(dict.fromkeys(pdf_files + recursive_pdf_files))

        if not pdf_files:
            raise ValueError("Не выбраны корректные PDF-файлы или папка")

        if len(pdf_files) == 1:
            path = pdf_files[0]
        else:
            all_dfs = []
            all_summaries = []
            all_roll_summaries = []
            base_name = "multi_pdf_analysis"
            combined_start = time.time()
            combined_files_processed = 0
            combined_pages_processed = 0
            combined_files_skipped = 0

            folder_paths = [Path(item).resolve() for item in clean_paths if Path(item).is_dir()]
            if folder_paths:
                root_dir = Path(os.path.commonpath([str(p) for p in folder_paths])) if len(folder_paths) > 1 else folder_paths[0]
            else:
                root_dir = Path(pdf_files[0]).resolve().parent

            for idx, pdf_path in enumerate(pdf_files, start=1):
                self.progress_value = 0
                self.progress_max = len(pdf_files)
                self.progress['maximum'] = max(100, len(pdf_files) * 20)
                self.progress['value'] = 0
                self.status_label.config(text=f"Обработка PDF ({idx}/{len(pdf_files)})...")
                self.root.update_idletasks()

                try:
                    df, summary, out_path, roll_summary = self.analyzer.process_path(
                        pdf_path,
                        progress_callback=self._update_progress,
                        save_excel=False,
                    )
                    combined_files_processed += self.analyzer.stats.get("files_processed", 0)
                    combined_pages_processed += self.analyzer.stats.get("pages_processed", 0)
                    combined_files_skipped += self.analyzer.stats.get("files_skipped", 0)
                    all_dfs.append(df)
                    all_summaries.append(summary)
                    all_roll_summaries.append(roll_summary)
                except Exception as e:
                    messagebox.showerror("Ошибка", f"Ошибка обработки файла {pdf_path}:\n{e}")
                    return

            self.analyzer.stats = {
                "files_processed": combined_files_processed,
                "pages_processed": combined_pages_processed,
                "files_skipped": combined_files_skipped,
                "errors": [],
                "start_time": combined_start,
                "end_time": time.time(),
            }

            combined_df = pd.concat(all_dfs, ignore_index=True) if all_dfs else pd.DataFrame()
            combined_summary = pd.concat(all_summaries, ignore_index=True) if all_summaries else pd.DataFrame()
            combined_roll = pd.concat(all_roll_summaries, ignore_index=True) if all_roll_summaries else pd.DataFrame()

            combined_out = root_dir / f"{base_name}_all_sizes.xlsx"
            self.analyzer.save_excel_safe(combined_out, combined_df, combined_summary, combined_roll)
            self.last_result = (combined_df, combined_summary, str(combined_out), combined_roll)
            report_text = self.analyzer.build_report_text(combined_df, combined_summary, str(combined_out), combined_roll)
            self._set_stats_text(report_text)
            self._update_roll_visualization(combined_df, combined_roll)
            self.treeview_tab.update_data(df=combined_df, summary=combined_summary, roll_summary=combined_roll)  # Обновляем древовидную вкладку
            self._save_report_auto(combined_df, combined_summary, str(combined_out), report_text, combined_roll)
            self.progress['value'] = self.progress['maximum']
            self.status_label.config(text="Готово!")
            self.root.update_idletasks()
            return

        # Одиночный PDF или папка: обрабатываем как один объект, без рекурсии.
        self.progress_value = 0
        self.progress_max = 0
        self.progress['value'] = 0

        self.status_label.config(text="Подсчёт файлов...")
        self.root.update_idletasks()

        path_obj = Path(path)
        if path_obj.is_file() and path_obj.suffix.lower() == ".pdf":
            self.progress_max = 1
        elif path_obj.is_dir():
            pdf_count = len(list(path_obj.rglob("*.pdf")))
            self.progress_max = max(1, pdf_count)
        else:
            raise ValueError("Не выбраны корректные PDF-файлы или папка")

        self.progress['maximum'] = max(100, self.progress_max * 20)
        self.progress_value = 0

        self.status_label.config(text="Обработка PDF...")
        self.root.update_idletasks()

        try:
            df, summary, out_path, roll_summary = self.analyzer.process_path(
                path,
                progress_callback=self._update_progress
            )
            self.last_result = (df, summary, out_path, roll_summary)
            report_text = self.analyzer.build_report_text(df, summary, out_path, roll_summary)
            self._set_stats_text(report_text)
            self._update_roll_visualization(df, roll_summary)
            self.treeview_tab.update_data(df=df, summary=summary, roll_summary=roll_summary)  # Обновляем древовидную вкладку
            # Автосохранение отчёта в текстовый файл рядом с Excel-отчётом
            # self._save_report_auto(df, summary, out_path, report_text, roll_summary)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка обработки:\n{e}")

        self.progress['value'] = self.progress['maximum']
        self.status_label.config(text="Готово!")
        self.root.update_idletasks()

    def _update_progress(self, current: int, total: int, filename: str = ""):
        """Обновляет прогресс-бар при обработке файлов"""
        if total > 0:
            # 80% на обработку файлов, 20% на финальные операции
            self.progress_value = int((current / total) * 80)
            self.progress['value'] = self.progress_value
            
            if filename:
                self.status_label.config(text=f"Обработка: {Path(filename).name} ({current}/{total})")
            
            self.root.update_idletasks()

    def _save_report_auto(self, df, summary, out_path: str, report_text: str, roll_summary=None):
        """Автоматически сохраняет отчёт после анализа"""
        base_name = Path(out_path).stem
        txt_path = Path(out_path).parent / f"{base_name}_report.txt"
        
        try:
            with open(txt_path, 'w', encoding='utf-8-sig') as f:
                f.write(report_text)
            print(f"📄 Автосохранение: {txt_path}")
        except Exception as e:
            print(f"⚠️ Автосохранение не удалось: {e}")

    def resource_path(self, relative_path):
        """Получает абсолютный путь к ресурсу, работает как в dev, так и в PyInstaller"""
        try:
            # PyInstaller создаёт временную папку _MEIPASS
            base_path = sys._MEIPASS # type: ignore
        except Exception:
            base_path = os.path.abspath(".")
        
        return os.path.join(base_path, relative_path)

    def show_help_window(self):
        win = tk.Toplevel(self.root)
        win.title("Помощь")
        win.geometry("550x660") # ширина x высота
        win.minsize(550, 660)
        win.configure(bg="#eef3f9")
        win.update_idletasks()
        x = (win.winfo_screenwidth() - 550) // 2
        y = (win.winfo_screenheight() - 660) // 2
        win.geometry(f"550x660+{x}+{y}")

        header = tk.Frame(win, bg="#405361", padx=5, pady=12)
        header.pack(fill=tk.X, padx=0, pady=0)

        logoframe = tk.Frame(header, bg="#405361", padx=0, pady=0)
        logoframe.pack(side=tk.TOP, padx=0, pady=0, fill=tk.X)
        
        logo_image = tk.Label(logoframe, image=self.icons["icon48"], bg="#405361")
        logo_image.pack(side=tk.LEFT, anchor=tk.W, padx=0, pady=0)

        header_title = tk.Label(
            logoframe,
            text="Анализатор страниц PDF-документов \nPDF Page Analyzer",
            fg="#f8fafc",
            bg="#405361",
            font=(font_face, 18, "bold"),
            justify=tk.LEFT
        )
        header_title.pack(anchor=tk.W, side=tk.LEFT, padx=(10, 0), pady=(0, 0))

        header_subtitle = tk.Label(
            header,
            text="Анализ форматов, размеров и цветности PDF",
            fg="#cbd5e1",
            bg="#405361",
            font=(font_face, 10)
        )
        header_subtitle.pack(side=tk.TOP, anchor=tk.W, pady=(4, 0))

        main = tk.Frame(win, bg="#ffffff", padx=5, pady=5)
        main.pack(fill=tk.BOTH, expand=True)

        info_frame = ttk.Frame(main, padding=(0, 0))
        info_frame.pack(fill=tk.X, pady=(0, 10), expand=True)

        about_heading = tk.Label(
            info_frame,
            text="О программе",
            justify=tk.LEFT,
            anchor="w",
            fg="#405361",
            bg="#ffffff",
            font=(font_face, 12, "bold"),
            wraplength=540
        )
        about_heading.pack(fill=tk.X, pady=(0, 0))

        info_text = (
            "Главная цель — быстро определить размеры, форматы и цветность страниц PDF,\n"
            "а затем сформировать сводку и Excel-отчёт для дальнейшей обработки."
            "\n"
            "Инструмент разработан для отдела выпуска компании СП-Инновация.\n"
            "\n"            
            f"Версия программы: {version}\n"
            "Автор: Родионов Вадим"
        )

        info_label = tk.Label(
            info_frame,
            text=info_text,
            justify=tk.LEFT,
            anchor="w",
            font=(font_face, 10),
            bg="#ffffff",
            fg="#0f172a",
            padx=14,
            pady=0, 
            wraplength=540
        )
        info_label.pack(fill=tk.X)

        icon_hyperlink_label = tk.Label(
            info_frame,
            text="Иконки взяты с сайта flaticon.com автора Magnific",
            justify=tk.LEFT,
            anchor="w",
            font=(font_face, 10, "underline"),
            bg="#ffffff",
            fg="#0f172a",
            padx=14,
            pady=0,
            wraplength=540,
            cursor="hand2"
        )
        icon_hyperlink_label.pack(fill=tk.X)
        icon_hyperlink_label.bind("<Button-1>", lambda e: webbrowser.open("https://www.flaticon.com/ru/authors/magnific"))

        steps_frame = ttk.Frame(main, padding=(0, 0))
        steps_frame.pack(fill=tk.X, expand=True)

        steps_heading = tk.Label(
            steps_frame,
            text="Как работать",
            justify=tk.LEFT,
            anchor="w",
            fg="#405361",
            bg="#ffffff",
            font=(font_face, 12, "bold"),
            wraplength=540
        )
        steps_heading.pack(fill=tk.X, pady=(0, 0))

        steps = (
            "1. Откройте файлы PDF для анализа одним из способов:\n"
            "       а. Нажмите «Открыть файл» для анализа одного PDF.\n"
            "       б. Нажмите «Открыть папку» для анализа всех PDF в выбранной папке\n"
            "           включая вложенные подпапки.\n"
            "       в. Забосьте файлы или папки прямо в окно программы.\n"
            "2. После обработки программа создаёт Excel-файл с листами: «Все страницы»,\n"
            "   «Сводка ЕСКД» и «Сводка по рулонам».\n"
            "3. Форматы распознаются по ГОСТ 2.301-68 (A0, A1, A4×3 и т.д.).\n"
            "   А также по собственным стандартам, настроенными в настройках.\n"
            "4. В основном окне отображается текстовый отчёт; его можно выделить и скопировать.\n"
            "5. Параметры допуска и список стандартных форматов задаются в настройках,\n"
            "   а также могут быть заданы в config.yaml (текстовый файл).\n"
            "6. При запуске с параметром пути (файл или папка) обработка выполняется сразу\n"
            "   (режим CLI).\n"
            "7. Если станет скучно, то попробуй вспомнить чит-коды из DOOM и вводить их на \n"
            "   клавиатуре, пока открыто окно помощи."
        )

        steps_label = tk.Label(
            steps_frame,
            text=steps,
            justify=tk.LEFT,
            anchor="w",
            font=(font_face, 10),
            bg="#ffffff",
            fg="#1f2937",
            padx=14,
            pady=0,
            wraplength=540
        )
        steps_label.pack(fill=tk.BOTH, expand=True)

        footer = tk.Frame(win, bg="#BFCDD2", padx=18, pady=10)
        footer.pack(fill=tk.X)

        github_btn = tk.Button(
            footer,
            text="Посетить GitHub",
            command=lambda: webbrowser.open("https://github.com/shadowdfd/PDF-pages-analizer"),
            bg="#405361",
            fg="#ffffff",
            activebackground="#1d4f7c",
            activeforeground="#ffffff",
            relief=tk.RAISED,
            padx=9,
            pady=9,
            bd=1,
            highlightbackground="#0c2942",
            highlightthickness=1,
            cursor="hand2",
            font=(font_face, 10, "bold")
        )
        github_btn.pack(side=tk.LEFT, padx=(0, 10))

        close_btn = tk.Button(
            footer,
            text="Закрыть",
            command=win.destroy,
            bg="#f1f5f9",
            fg="#0f172a",
            activebackground="#dfe7f0",
            activeforeground="#0f172a",
            relief=tk.RAISED,
            padx=9,
            pady=9,
            bd=1,
            highlightbackground="#94a3b8",
            highlightthickness=1,
            cursor="hand2",
            font=(font_face, 10, "bold")
        )
        close_btn.pack(side=tk.RIGHT)

        try:
            from PIL import Image, ImageTk

            logo_path = self.resource_path("img/logo.png")
            img = Image.open(logo_path)
            logo_img = ImageTk.PhotoImage(img)

            logo_container = tk.Frame(
                footer,
                bg="#BFCDD2",
                bd=0,
                relief=tk.SOLID,
                padx=6,
                pady=6
            )
            logo_container.pack(side=tk.RIGHT, padx=(10, 10))

            logo_label = tk.Label(
                logo_container,
                image=logo_img,
                borderwidth=0,
                bg="#BFCDD2"
            )
            logo_label.pack(padx=5, pady=0)
        except (ImportError, FileNotFoundError):
            logo_label = tk.Label(
                footer,
                text="🏢 СП-Инновация",
                font=(font_face, 12, "bold"),
                fg="#BFCDD2",
                bg="#405361",
                padx=10,
                pady=7,
                bd=1,
                relief=tk.SOLID
            )
            logo_label.pack(side=tk.RIGHT, padx=(10, 0))

        # 🎮 Пасхалка: отслеживание ввода чит-кодов
        cheat_buffer = []

        # Храним только SHA-256 хеши кодов, чтобы сами последовательности
        # не были видны при обычном просмотре исходного кода.
        cheat_salt = "pdf-pages-easter-egg-v1:"
        cheat_hashes = {
            'f44d2cc7bd669a80b4e58da86b7fc3a986b6a47b8f8bf86f109807bab53fd192': 'snake',
            '675dff614dd069ddd14a5f1cec810d123328d67c5a9a9af54d193e4d36bcc3af': 'snake',
            '96be471287e32979669b41439da5cb661d4d2e08997349c4f7abea28d8f157a4': 'tetris',
            'ae13cbbde434821f038e98a2dcae6fb56987d01e250eac67e203d1b9f4d60c61': 'tetris',
            '18afa4f742d0db24384b18c8b98ef93204b9819998118d5e06554bd4075e15ad': 'arkanoid',
            '220c527b2997e4c4d3445683f20c454fe5af1ea73d8892f48b1058420c11e87d': 'arkanoid',
            '2dd7147369a84f5067f83c00c237bf9903dfa47c7de664f93197e8eb21201290': 'galaga',
            '572c474a9a8fab307a834dade8e2d7a9a8639d02d19ad7f75a2fc8a4d9388a43': 'galaga',
        }

        def on_key_press(event):
            # Добавляем символ в буфер
            char = (event.char or '').lower()
            if char in 'idkfaqlcpsoшвйлафсдзыщ':
                cheat_buffer.append(char)
                # Оставляем только последние 10 символов (макс. длина кода)
                if len(cheat_buffer) > 10:
                    cheat_buffer.pop(0)

                buffer_str = ''.join(cheat_buffer)
                game_type = None
                for code_length in (5, 6, 10):
                    if len(buffer_str) < code_length:
                        continue
                    digest = hashlib.sha256(
                        (cheat_salt + buffer_str[-code_length:]).encode('utf-8')
                    ).hexdigest()
                    game_type = cheat_hashes.get(digest)
                    if game_type is not None:
                        break
                if game_type is not None:
                    # Очищаем буфер, чтобы код не сработал повторно.
                    cheat_buffer.clear()
                    win.destroy()
                    self._launch_easter_egg(game_type)
        win.bind("<Escape>", lambda e: win.destroy())
        win.bind("<Key>", on_key_press)
        win.transient(self.root)
        win.grab_set()
        win.lift()
        win.after_idle(win.focus_force)
        win.wait_window()

    def open_config_editor(self):
        """Открывает редактор конфигурации"""
        try:
            from config_editor import ConfigEditor
            editor = ConfigEditor(parent=self.root)
            editor.grab_set()  # модальное окно
            editor.wait_window()  # ждем закрытия
            print("✅ ConfigEditor закрыт")
            self.refresh_config()  # перезагружаем настройки
        except ImportError as e:
            messagebox.showerror("Ошибка", f"Не найден config_editor.py:\n{str(e)}")

    def select_all(self):
        self.stats_text.config(state=tk.NORMAL)
        self.stats_text.tag_add("sel", "1.0", "end")
        self.stats_text.config(state=tk.DISABLED)
        self.stats_text.focus_set()
        self.status_label.config(text=f"Всё выделено ({len(self.stats_text.get('1.0', 'end').strip())} симв.)")

    def copy_selection(self):
        """Копирует выделенный текст или весь текст, если ничего не выделено"""
        try:
            # Пытаемся скопировать выделенный текст
            text = self.stats_text.get(tk.SEL_FIRST, tk.SEL_LAST)
        except tk.TclError:
            # Если ничего не выделено - копируем всё
            text = self.stats_text.get("1.0", tk.END).strip()
                
        if text:
            self.root.clipboard_clear()
            self.root.clipboard_append(text)
            self.status_label.config(text=f"Скопировано {len(text)} символов")
        else:
            self.status_label.config(text="Нет текста для копирования")

    def copy_report(self):
        """Копирует весь текст отчёта"""
        text = self.stats_text.get("1.0", tk.END).strip()
        if not text:
            self.status_label.config(text="Нет текста для копирования")
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.status_label.config(text=f"Отчет скопирован. {len(text)} символов")

    def save_report_to_file(self):
        """Сохраняет текущий текст отчёта в файл рядом с XLSX"""
        if not self.last_result:
            messagebox.showwarning("Предупреждение", "Сначала выполните анализ!")
            return

        df, summary, out_path, roll_summary = self.last_result
        report_text = self.stats_text.get("1.0", tk.END).strip()
        
        # Имя файла: тот же base_name + _report.txt
        base_name = Path(out_path).stem
        txt_path = Path(out_path).parent / f"{base_name}_report.txt"
        
        try:
            with open(txt_path, 'w', encoding='utf-8-sig') as f:
                f.write(report_text)
            messagebox.showinfo("Сохранено", f"Отчёт сохранён:\n{txt_path}")
            print(f"📄 Текстовый отчёт: {txt_path}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить:\n{e}")

    def open_excel_report(self):
        """Открывает последний Excel отчёт"""
        if not self.last_result:
            messagebox.showwarning("Предупреждение", "Сначала выполните анализ!")
            return
        
        _, _, out_path, _ = self.last_result
        
        try:
            # Открываем файл стандартным обработчиком
            if platform.system() == "Windows":
                os.startfile(out_path)
            elif platform.system() == "Darwin":
                subprocess.run(['open', out_path])
            else:
                subprocess.run(['xdg-open', out_path])
            print(f"📊 Открыт Excel: {out_path}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось открыть файл:\n{e}")

    def open_reports_folder(self):
        """Открывает папку с отчётами в проводнике"""
        if not self.last_result:
            messagebox.showwarning("Предупреждение", "Сначала выполните анализ!")
            return
        
        _, _, out_path, _ = self.last_result
        folder_path = str(Path(out_path).parent)
        
        try:
            # Открываем папку в стандартном проводнике
            if platform.system() == "Windows":
                os.startfile(folder_path)
            elif platform.system() == "Darwin":
                subprocess.run(['open', folder_path])
            else:
                subprocess.run(['xdg-open', folder_path])
            print(f"📁 Открыта папка: {folder_path}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось открыть папку:\n{e}")

    def refresh_config(self):
        """Перезагружает config.yaml и обновляет состояние приложения"""
        try:
            # 1. Перезагружаем конфиг
            self.config = self._load_config()
            self.analyzer.apply_config(self.config)
            print(f"self.analyzer.compress_ranges_y = {self.analyzer.compress_ranges_y}")
            print(f"✅ Конфиг перезагружен:")
            print(f"   📐 tolerance_mm: {self.config.get('tolerance_mm', 5.0)}")
            print(f"   📦 compress_ranges: {self.config.get('compress_ranges', True)}")
            print(f"   📚 форматов загружено: {len(self.config.get('formats', {}))}")
            
            # 2. Обновляем статус в интерфейсе
            self._update_config_status()
            
            # 3. Если есть результаты анализа - пересчитываем
            if self.last_result:
                self._update_results_display()
                
            #messagebox.showinfo("✅ Конфигурация", 
            #                f"Обновлено:\n"
            #                f"• Допуск: {self.config['tolerance_mm']} мм\n"
            #                f"• Сжатие диапазонов: {'Вкл' if self.config['compress_ranges'] else 'Выкл'}")
            #                f"• Форматов загружено: {len(self.config['formats'])}")

        except Exception as e:
            print(f"❌ Ошибка refresh_config: {e}")
            messagebox.showerror("Ошибка", f"Не удалось обновить конфиг:\n{str(e)}")

    def _load_config(self) -> Dict[str, Any]:
        """Загружает config.yaml.
        Возвращает только то, что загружено из файла.
        Дефолтные НАСТРОЙКИ применяются в apply_config при необходимости."""
        possible_paths = [
            Path("config.yaml"),
            Path(__file__).parent / "config.yaml",
            Path.cwd() / "config.yaml",
        ]
        
        config_path = None
        for path in possible_paths:
            if path.exists():
                config_path = path
                break
        
        if config_path is None:
            config_path = Path("config.yaml")
        
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                config = yaml.safe_load(f)
                if config is None:
                    print(f"⚠️ Пустой или невалидный config.yaml, использую дефолтные настройки")
                    config = {}
        except yaml.YAMLError as e:
            print(f"❌ Синтаксическая ошибка в config.yaml: {e}")
            print(f"ℹ️ Используются дефолтные настройки")
            config = {}
        except Exception as e:
            print(f"⚠️ Не могу прочитать config.yaml: {e}")
            print(f"ℹ️ Используются дефолтные настройки")
            config = {}
        
        # Возвращаем только то, что загружено из файла (без автоматических дефолтов)
        return config

    def _update_config_status(self):
        """Обновляет индикаторы конфигурации в интерфейсе"""
        try:
            # Статус допуска
            tolerance_status = self.config.get('tolerance_mm', 5.0)
            self.analyzer.tolerance_var.set(f"{tolerance_status}")
            
            # Статус сжатия диапазонов
            compress_status = self.config.get('compress_ranges', True)
            self.analyzer.compress_ranges_var.set(f"{'Вкл' if compress_status else 'Выкл'}")

            # Статус количества форматов
            formats_count_status = len(self.config.get('formats', {}))
            self.analyzer.formats_count_var.set(f"{formats_count_status}")

        except Exception as e:
            print(f"⚠️ Ошибка обновления статуса: {e}")

    def _update_results_display(self):
        """Пересчитывает отображение результатов с новым compress_ranges"""
        if not self.last_result:
            return

        # Проверяем, инициализирован ли stats_text (на случай вызова до _build_ui)
        if not hasattr(self, 'stats_text'):
            return

        # Пересчитываем с новыми настройками
        df, summary, out_path, roll_summary = self.last_result
        report = self.analyzer.build_report_text(df, summary, out_path, roll_summary)

        # Обновляем Text виджет
        self._set_stats_text(report)

        # Обновляем визуализацию
        self._update_roll_visualization(df, roll_summary)

        # Обновляем древовидную вкладку
        self.treeview_tab.update_data(df=df, summary=summary, roll_summary=roll_summary)

    # ---------- пасхалка ----------

    def _launch_easter_egg(self, game_type='snake'):
        """Запускает пасхалку - мини-игры

        Args:
            game_type: тип игры ('snake', 'tetris', 'arkanoid' или 'galaga')
        """
        try:
            if game_type == 'snake':
                from snake_game import launch_snake_game
                launch_snake_game(parent=self.root)
            elif game_type == 'tetris':
                from tetris_game import launch_tetris_game
                launch_tetris_game(parent=self.root)
            elif game_type == 'arkanoid':
                from arkanoid_game import launch_arkanoid_game
                launch_arkanoid_game(parent=self.root)
            elif game_type == 'galaga':
                from galaga_game import launch_galaga_game
                launch_galaga_game(parent=self.root)
        except ImportError as e:
            messagebox.showerror("Ошибка", f"Не найден модуль игры:\n{str(e)}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось запустить игру:\n{str(e)}")

    # ---------- запуск ----------

    def run(self):
        self.root.mainloop()

def main():
    analyzer = PDFAnalyzer()

    if len(sys.argv) >= 2:
        input_path = sys.argv[1]
        try:
            df, summary, out_path, roll_summary = analyzer.process_path(input_path)
            # передаём готовый результат в GUI-класс
            app = MainWindow(analyzer, initial_result=(df, summary, out_path, roll_summary))
            app.run()
        except Exception as e:
            # даже в CLI-сценарии покажем нормальное окно об ошибке
            #import tkinter as tk
            #from tkinter import messagebox
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("Ошибка", f"Ошибка обработки:\n{e}")
            root.destroy()
    else:
        # GUI-режим по умолчанию
        app = MainWindow(analyzer)
        app.run()

if __name__ == "__main__":
    main()
 
