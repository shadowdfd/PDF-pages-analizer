import os
import sys
import fitz  # PyMuPDF
import pandas as pd
import yaml
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
#import threading
import webbrowser

pt_to_mm = 0.3528
font_face = "Calibri"
version = "v.1.3.0"
config_path="config.yaml"

class PDFAnalyzer:
    
    def __init__(self, config_path="config.yaml"):
        self.config = self.load_config(config_path)
        self.tolerance = self.config["tolerance_mm"]
        self.compress_ranges_y = self.config["compress_ranges"]
        self.formats = {k: tuple(v) for k, v in self.config["formats"].items()}
        self.stats = {
            "files_processed": 0,
            "pages_processed": 0,
            "errors": [],
            "files_skipped": 0
        }

    def load_config(self, config_path):
        """Загружает конфигурацию из YAML"""
        # Конфиг по умолчанию
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
            with open(config_path, 'r', encoding='utf-8') as f:
                config = yaml.safe_load(f)
                config["fileload"] = "загружен"
            return {**default_config, **config}
        except FileNotFoundError:
            print(f"config.yaml не найден, используются значения по умолчанию")
            messagebox.showerror("Файл конфигурации", "Файл конфигурации config.yaml не найден, используются значения по умолчанию.")
            return default_config
        except Exception as e:
            print(f"Ошибка чтения config.yaml: {e}")
            messagebox.showerror("Файл конфигурации", "Ошибка чтения config.yaml: "+ str(e) + "\n Используются значения по умолчанию.")
            return default_config

    def get_standard_format(self, w_mm: float, h_mm: float) -> tuple[str, str]:
        """Custom1, Custom2... для нестандартных размеров"""
        
        # Стандартные форматы
        for name, (sw, sh) in self.formats.items():
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
                pix = fitz.Pixmap(page.parent, img[0])
                try:
                    if pix.colorspace and pix.colorspace.n > 1:
                        return "Цветная"
                finally:
                    pix = None
            for draw in page.get_drawings():
                for col in (draw.get("fill"), draw.get("stroke")):
                    if col and len(col) >= 3:
                        r, g, b = col[:3]
                        if not (abs(r - g) < 1e-3 and abs(g - b) < 1e-3):
                            return "Цветная"
        except:
            pass
        return "Ч/Б"

    def process_pdf(self, pdf_path: str, all_data: list) -> None:
        
        """Обрабатывает один PDF файл"""
        try:
            self.stats["files_processed"] += 1
            print(f"Обрабатываем: {os.path.basename(pdf_path)}")
            doc = fitz.open(pdf_path)
            self.stats["pages_processed"] += len(doc)

            for i, page in enumerate(doc, start=1):
                rotation = page.rotation
                rect = page.rect
                w_pt, h_pt = rect.width, rect.height

                if rotation in (90, 270):
                    width_mm = round(h_pt * pt_to_mm, 1)
                    height_mm = round(w_pt * pt_to_mm, 1)
                else:
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

    def process_path(self, path: str) -> tuple[pd.DataFrame, pd.DataFrame, str]:
        """Главная функция обработки"""
        all_data = []
        self.config = self.load_config(config_path)
        path_obj = Path(path)


        if path_obj.is_file() and path_obj.suffix.lower() == ".pdf":
            self.process_pdf(str(path_obj), all_data)
            base_name = path_obj.stem
            out_dir = path_obj.parent

        elif path_obj.is_dir():
            base_name = path_obj.name
            out_dir = path_obj
            pdf_files = list(path_obj.glob("*.pdf"))
            self.stats["files_skipped"] = len([f for f in path_obj.iterdir() if f.suffix.lower() != ".pdf"])
            
            for pdf_path in pdf_files:
                self.process_pdf(str(pdf_path), all_data)

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
        for (format_name, file_name), group in df.groupby(["Стандартный формат", "Файл"]):
            summary_data.append({
                "Файл": file_name,
                "Стандартный формат": format_name,
                "Количество": len(group),
                "Ч/Б страницы": get_page_list(group, "Ч/Б"),
                "Цветные страницы": get_page_list(group, "Цветная"),
                "Цветных": (group["Цветность"] == "Цветная").sum(),
                "Ч/Б": (group["Цветность"] == "Ч/Б").sum(),
            })

        summary = pd.DataFrame(summary_data).sort_values(["Файл", "Стандартный формат"]).reset_index(drop=True)
        
        out_path = out_dir / f"{base_name}_all_sizes.xlsx"

        with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Все страницы")
            summary.to_excel(writer, index=False, sheet_name="Сводка ЕСКД")

        
        return df, summary, str(out_path)

    def build_report_text(self, df: pd.DataFrame, summary: pd.DataFrame, out_path: str) -> str:
        """Формирует текстовый отчёт для вывода в GUI/копирования."""
        lines: list[str] = []

        # Заголовок и общая статистика
        lines.append(f"=== ОТЧЁТ АНАЛИЗА PDF ===  Дата: {pd.Timestamp.now().strftime('%d.%m.%Y %H:%M')}")
        lines.append(f"Файл отчёта: {out_path}")
        lines.append("")

        lines.append("📊 ОБЩАЯ СТАТИСТИКА:")
        lines.append(f"    Файлов обработано: {self.stats.get('files_processed', 0)}")
        lines.append(f"    Листов обработано: {self.stats.get('pages_processed', 0)}")
        lines.append(f"    Файлов пропущено (не PDF): {self.stats.get('files_skipped', 0)}")
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

        def compress_ranges(input_str):
            # Парсим строку в список чисел
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
        
        for file_name, file_group in df.groupby("Файл"):
            lines.append(f"\n    {file_name}:")

            # --- ОТДЕЛЬНЫЕ ФОРМАТЫ ВНУТРИ ФАЙЛА ---
            file_formats = file_group.groupby("Стандартный формат").agg({
                "Страница": "count"
            }).reset_index()

            file_format_details = file_group.groupby("Стандартный формат")["Страница"].apply(
                lambda x: ", ".join(map(str, sorted(x.tolist())))
            ).to_dict()

            for _, frow in file_formats.iterrows():
                fmt = frow["Стандартный формат"]          # A4, A3, A3×4, Custom1 ...
                total_pages = int(frow["Страница"])
                print(f"self.compress_ranges_y = {self.compress_ranges_y}")
                if self.compress_ranges_y:
                    pages_list = compress_ranges(file_format_details.get(fmt, "-"))
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
                    pages_a4a3 = compress_ranges(",".join(map(str, sorted(a4a3_group["Страница"].tolist()))))
                else:
                    pages_a4a3 = ",".join(map(str, sorted(a4a3_group["Страница"].tolist())))                 
               
                lines.append(f"        A4 + A3 ({total_a4a3} стр.): {pages_a4a3}")

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

class MainWindow:
    
    def __init__(self, analyzer, initial_result=None):
        """
        initial_result:
          - None → обычный режим (ждём, пока пользователь нажмёт Открыть файл/папку)
          - (df, summary, out_path) → показываем уже готовый отчёт (CLI‑режим)
        """
        self.analyzer = analyzer
        self.root = tk.Tk()
        
        icon_path = self.resource_path("icon.png")
        icon = tk.PhotoImage(file=icon_path)
        
        # True — применить ко всем будущим Toplevel
        self.root.iconphoto(True, icon)
       
        self.root.title("Анализатор PDF файлов "+version)
        self.root.geometry("900x650")
        self.root.resizable(True, True)



        self.last_result = initial_result  # (df, summary, out_path) или None

        self._build_ui()
       
        # если уже есть результат (CLI-сценарий) — сразу показываем его
        if self.last_result is not None:
            df, summary, out_path = self.last_result
            report_text = self.analyzer.build_report_text(df, summary, out_path)
            self._set_stats_text(report_text)
        else:
            self._set_stats_text("Выберите файл или папку для анализа PDF документов.")

    # ---------- построение интерфейса ----------

    def _build_ui(self):
        root = self.root

        # Верхняя панель: статус конфигурации + кнопки
        top_frame = ttk.Frame(root, padding=10)
        top_frame.pack(fill=tk.X)

        # Кнопки управления слева
        btn_top_frame1 = ttk.Frame(top_frame)
        btn_top_frame1.pack(side=tk.LEFT, anchor=tk.NW)

        ttk.Button(btn_top_frame1, text="📂 Открыть файл", command=self.on_open_file, width=25).pack(side=tk.TOP, padx=5, pady=5)
        ttk.Button(btn_top_frame1, text="📂 Открыть папку", command=self.on_open_folder, width=25).pack(side=tk.TOP, padx=5, pady=5)
        
        btn_top_frame2 = ttk.Frame(btn_top_frame1)
        btn_top_frame2.pack(side=tk.TOP)

        ttk.Button(btn_top_frame2, text="?", command=self.show_help_window, width=7).pack(side=tk.LEFT, padx=5, pady=5)
        ttk.Button(btn_top_frame2, text="⚙️ Настройки", command=self.open_config_editor, width=15).pack(side=tk.RIGHT, padx=5, pady=5)

        # Статус config.yaml справа
        status_frame = ttk.LabelFrame(top_frame, text="Текущие настройки", padding=10)
        status_frame.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=10)

        status_names_frame = ttk.Frame(status_frame)
        status_names_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=10)

        self.tolerance_name_label = ttk.Label(status_names_frame, text="Допуск:", font=(font_face, 10))
        self.tolerance_name_label.pack(anchor=tk.E)
        self.compress_name_label = ttk.Label(status_names_frame, text="Диапазоны:", font=(font_face, 10))
        self.compress_name_label.pack(anchor=tk.E)
        self.formats_count_name_label = ttk.Label(status_names_frame, text="Форматов загруженно:", font=(font_face, 10))
        self.formats_count_name_label.pack(anchor=tk.E)

        status_values_frame = ttk.Frame(status_frame)
        status_values_frame.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=10)

        self.tolerance_status_label = ttk.Label(status_values_frame, 
                                       text="...", 
                                       font=(font_face, 10))
        self.tolerance_status_label.pack(anchor=tk.W)

        self.compress_status_label = ttk.Label(status_values_frame, 
                                            text="...", 
                                            font=(font_face, 10))
        self.compress_status_label.pack(anchor=tk.W)

        self.formats_count_status_label = ttk.Label(status_values_frame, 
                                            text="...", 
                                            font=(font_face, 10))
        self.formats_count_status_label.pack(anchor=tk.W)

        # 2. Обновляем статус в интерфейсе
        self.refresh_config()
         
        # Центральная область — текст отчёта
        center_frame = ttk.LabelFrame(root, text="Отчёт", padding=10)
        center_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        self.stats_text = tk.Text(
            center_frame,
            wrap=tk.WORD,
            font=(font_face, 10),
            bg="#f8f9fa",
            relief="solid",
            bd=1,
            selectbackground="#4CAF50",
            selectforeground="white",
            padx=10,
            pady=10,
        )
        scroll = ttk.Scrollbar(center_frame, orient=tk.VERTICAL, command=self.stats_text.yview)
        self.stats_text.configure(yscrollcommand=scroll.set)

        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.stats_text.pack(fill=tk.BOTH, expand=True)

        # Нижняя панель — сервисные кнопки
        bottom_frame = ttk.Frame(root, padding=10)
        bottom_frame.pack(fill=tk.X)

        ttk.Button(bottom_frame, text="Выделить всё", command=self.select_all).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(bottom_frame, text="Копировать отчёт", command=self.copy_report).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(bottom_frame, text="Выход", command=root.destroy).pack(side=tk.RIGHT, padx=(5, 0))       
        ttk.Button(bottom_frame, text="💾 Сохранить отчёт", command=self.save_report_to_file).pack(side=tk.RIGHT, padx=(5, 0))

        def hotkeys(event):  # Обработчик нажатий горячих клавиш
            CTRL_MASK = 0x0004  # Control_L
            SHIFT_MASK = 0x0001 # Shift_L
            
            #Debug string
            #print(f"keycode={event.keycode}, state={hex(event.state)}, keysym={event.keysym}")
            
            # Ctrl+A (keycode=65)
            if (event.state & CTRL_MASK) and event.keycode == 65:
                self.select_all()
                return "break"
            # Ctrl+S (keycode=83)
            elif (event.state & CTRL_MASK) and event.keycode == 83:
                self.save_report_to_file
                return "break"
            # Ctrl+C (keycode=67)
            elif (event.state & CTRL_MASK) and event.keycode == 67:
                self.stats_text.event_generate("<<Copy>>")
                return "break"
            return None
        
        
        #Статус бар
        status_frame = ttk.Frame(root, padding=10)
        status_frame.pack(fill=tk.X)
        
        self.status_label = tk.Label(status_frame, text="Готов", font=(font_face, 10), width=15, justify=tk.LEFT)
        self.status_label.pack(side=tk.LEFT, pady=5, anchor="nw")
        
        self.progress = ttk.Progressbar(status_frame, mode='indeterminate')
        self.progress.pack(pady=10, fill='x', side=tk.RIGHT, expand=True)
        
        
        # Горячие клавиши
        self.root.bind("<KeyPress>", hotkeys)
        self.root.bind("<Escape>", lambda e: root.destroy())
        
        self.stats_text.focus_set() # Фокус на тексовый отчет
        
        #root.bind("<Control-a>", lambda e: self.select_all())
        #root.bind("<Control-c>", lambda e: self.stats_text.event_generate("<<Copy>>"))

    def _set_stats_text(self, text: str):
        self.stats_text.config(state=tk.NORMAL)
        self.stats_text.delete("1.0", tk.END)
        self.stats_text.insert(tk.END, text)
        self.stats_text.config(state=tk.DISABLED)

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

    def _run_analysis(self, path: str):
        
        self.progress.start(10)
        self.status_label.config(text="Обработка PDF...")
        self.root.update_idletasks()
        
        try:
            df, summary, out_path = self.analyzer.process_path(path)
            self.last_result = (df, summary, out_path)
            report_text = self.analyzer.build_report_text(df, summary, out_path)
            self._set_stats_text(report_text)
             # 🆕 АВТОМАТИЧЕСКОЕ СОХРАНЕНИЕ ТЕКСТОВОГО ОТЧЁТА
            self._save_report_auto(df, summary, out_path, report_text)           
            
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка обработки:\n{e}")
            
        self.progress.stop()
        self.status_label.config(text="Готово!")

    def _save_report_auto(self, df, summary, out_path: str, report_text: str):
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
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")
        
        return os.path.join(base_path, relative_path)

    def show_help_window(self):
        win = tk.Toplevel(self.root)
        win.title("Помощь")
        win.geometry("500x400")

        instr_frame = ttk.LabelFrame(win, text="Инструкция", padding="10")
        instr_frame.pack(fill=tk.X, anchor="nw", expand=True, padx=5, pady=5)       

        instructions = (
            "\n"
            " 1. Нажмите «Открыть файл» для анализа одного PDF.\n"
            " 2. Нажмите «Открыть папку» для анализа всех PDF в выбранной папке.\n"
            " 3. После обработки создаётся Excel-файл с листами:\n"
            "   - «Все страницы» — детальная информация по страницам\n"
            "   - «Сводка ЕСКД» — сводка по форматам и цветности\n"
            "   Также создаётся текстовый файл с отчетом.\n"          
            " 4. Форматы распознаются по ГОСТ 2.301-68 (A0, A1, A4×3 и т.д.)\n"
            " 5. В основном окне отображается текстовый отчёт — его можно частично или полностью выделить и скопировать.\n"
            " 6. Параметры допуска и список стандартных форматов задаются в файле config.yaml.\n"
            " 7. Программу можно запускать с параметром пути (файл/папка) через командную строку — в этом случае сразу выполняется обработка и открывается окно с отчётом.\n"
            "\n"
            "Инструмент разработан для Отдела выпуска компании СП-Инновация\n"
            "Версия программы: "+version+
            "\nАвтор: Родионов Вадим\n"
        )
 
        instr_text = tk.Label(instr_frame, font=(font_face, 10), text=instructions, justify="left", wraplength=450 )
        instr_text.pack(side=tk.LEFT, anchor="nw", expand=True, padx=10, ipadx=0)    
               
        # Группа для лототипа
        Logo_frame = ttk.Frame(win)
        Logo_frame.pack(fill=tk.BOTH, pady=5, expand=True)

        def open_link():
            webbrowser.open("https://github.com/shadowdfd/PDF-pages-analizer")
       
        Gitbutton = ttk.Button(Logo_frame, text="Посетить GitHub", command=open_link)
        Gitbutton.pack(side=tk.LEFT, padx=(20,0), pady=5)

        # ЛОГОТИП - ПРАВЫЙ НИЖНИЙ УГОЛ (работает в EXE)
        try:
            from PIL import Image, ImageTk
            
            logo_path = self.resource_path("logo.png")
            img = Image.open(logo_path)
            #img = img.resize((200, 33), Image.Resampling.LANCZOS)
            logo_img = ImageTk.PhotoImage(img)
            
            logo_label = tk.Label(Logo_frame, image=logo_img, borderwidth=0)
            logo_label.image = logo_img
            logo_label.pack(anchor="ne", padx=5, pady=5)
            #logo_label.place(relx=1.0, rely=0.02, anchor="ne", x=-5, y=5)
            
        except (ImportError, FileNotFoundError):
            # Текстовый логотип как fallback
            logo_label = tk.Label(Logo_frame, text="🏢 СП-Инновация", 
                                 font=(font_face, 14, "bold"), fg="#2E86AB")
            logo_label.pack(anchor="ne", padx=5, pady=5)
            #logo_label.place(relx=1.0, rely=0.02, anchor="se", x=-5, y=5)

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

    def copy_report(self):
        text = self.stats_text.get("1.0", tk.END).strip()
        if not text:
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        messagebox.showinfo("Копирование", "Текст отчёта скопирован в буфер обмена.")

    def save_report_to_file(self):
        """Сохраняет текущий текст отчёта в файл рядом с XLSX"""
        if not self.last_result:
            messagebox.showwarning("Предупреждение", "Сначала выполните анализ!")
            return
        
        df, summary, out_path = self.last_result
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

    def refresh_config(self):
        """Перезагружает config.yaml и обновляет состояние приложения"""
        try:
            # 1. Перезагружаем конфиг
            self.config = self._load_config()
            self.analyzer.compress_ranges_y = self.config.get('compress_ranges', True)
            print(f"self.analyzer.compress_ranges_y = {self.analyzer.compress_ranges_y}")
            print(f"✅ Конфиг перезагружен:")
            print(f"   📐 tolerance_mm: {self.config.get('tolerance_mm', 5.0)}")
            print(f"   📦 compress_ranges: {self.config.get('compress_ranges', True)}")
            print(f"   📚 форматов загружено: {len(self.config.get('formats', {}))}")
            
            # 2. Обновляем статус в интерфейсе
            self._update_config_status()
            
            # 3. Если есть результаты анализа - пересчитываем
            if hasattr(self, 'results_text') and self.results_text.get(1.0, tk.END).strip():
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
        """Загружает config.yaml (тот же код что в ConfigEditor)"""
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
            with open(config_path, 'r', encoding='utf-8') as f:
                config = yaml.safe_load(f) or {}
        except Exception:
            config = {}
        
        return {**default_config, **config}

    def _update_config_status(self):
        """Обновляет индикаторы конфигурации в интерфейсе"""
        try:
            # Статус допуска
            tolerance_status = self.config.get('tolerance_mm', 5.0)
            tolerance_label = getattr(self, 'tolerance_status_label', None)
            if tolerance_label:
                tolerance_label.config(
                    text=f"{tolerance_status} мм",
                    foreground="green"
                )
            
            # Статус сжатия диапазонов
            compress_status = self.config.get('compress_ranges', True)

            compress_label = getattr(self, 'compress_status_label', None)
            if compress_label:
                compress_label.config(
                    text=f"{'Сжатие ВКЛ' if compress_status else 'По отдельности'}",
                    foreground="green" if compress_status else "orange"
                )
            
            # Статус количества форматов
            formats_count_status = len(self.config.get('formats', {}))
            formats_count_label = getattr(self, 'formats_count_status_label', None)
            if formats_count_label:
                formats_count_label.config(
                    text=f"{formats_count_status}",
                    foreground="green"
                )

        except Exception as e:
            print(f"⚠️ Ошибка обновления статуса: {e}")

    def _update_results_display(self):
        """Пересчитывает отображение результатов с новым compress_ranges"""
        if not hasattr(self, 'analyzer') or not self.analyzer:
            return
            
        # Пересчитываем с новыми настройками
        self.analyzer.config = self.config
        formats_data = self.analyzer.analyze_all()
        
        # Обновляем Text виджет
        self.results_text.delete(1.0, tk.END)
        report = self.analyzer.build_report_text(formats_data)
        self.results_text.insert(1.0, report)

    # ---------- запуск ----------

    def run(self):
        self.root.mainloop()

def main():
    analyzer = PDFAnalyzer()

    if len(sys.argv) >= 2:
        input_path = sys.argv[1]
        try:
            df, summary, out_path = analyzer.process_path(input_path)
            # передаём готовый результат в GUI-класс
            app = MainWindow(analyzer, initial_result=(df, summary, out_path))
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
