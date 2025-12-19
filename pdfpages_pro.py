import os
import sys
import fitz  # PyMuPDF
import pandas as pd
import yaml
import tkinter as tk
from tkinter import messagebox, ttk
from pathlib import Path
import threading

pt_to_mm = 0.3528
font_face = "Calibri"

class PDFAnalyzer:
    def __init__(self, config_path="config.yaml"):
        self.config = self.load_config(config_path)
        self.tolerance = self.config["tolerance_mm"]
        self.formats = {k: tuple(v) for k, v in self.config["formats"].items()}
        self.stats = {
            "files_processed": 0,
            "pages_processed": 0,
            "errors": [],
            "files_skipped": 0
        }

    def load_config(self, config_path):
        """Загружает конфигурацию из YAML"""
        default_config = {
            "tolerance_mm": 5.0,
            "formats": {
                "A4": [210, 297],
                "A3": [297, 420],
                # ... остальные форматы
            }
        }
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                config = yaml.safe_load(f)
            return {**default_config, **config}
        except FileNotFoundError:
            print(f"config.yaml не найден, используются значения по умолчанию")
            return default_config
        except Exception as e:
            print(f"Ошибка чтения config.yaml: {e}")
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

        
        self.last_result = (df, summary, str(out_path))
        return self.last_result

    def show_report(self):
        """Улучшенный GUI-отчёт с логотипом, инструкцией и копированием"""
        root = tk.Tk()
        root.title("Отчёт анализа PDF")
        root.geometry("700x600")
        root.resizable(True, True)

        # Главный фрейм
        main_frame = ttk.Frame(root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        def resource_path(relative_path):
            """Получает абсолютный путь к ресурсу, работает как в dev, так и в PyInstaller"""
            try:
                # PyInstaller создаёт временную папку _MEIPASS
                base_path = sys._MEIPASS
            except Exception:
                base_path = os.path.abspath(".")
            
            return os.path.join(base_path, relative_path)

        # ЛОГОТИП - ПРАВЫЙ ВЕРХНИЙ УГОЛ (работает в EXE)
        try:
            from PIL import Image, ImageTk
            
            logo_path = resource_path("logo.png")
            img = Image.open(logo_path)
            #img = img.resize((200, 33), Image.Resampling.LANCZOS)
            logo_img = ImageTk.PhotoImage(img)
            
            logo_label = tk.Label(root, image=logo_img, borderwidth=0)
            logo_label.image = logo_img
            logo_label.place(relx=1.0, rely=0.02, anchor="ne", x=-5, y=5)
            
        except (ImportError, FileNotFoundError):
            # Текстовый логотип как fallback
            logo_label = tk.Label(root, text="🏢 PDF Analyzer", 
                                 font=(font_face, 14, "bold"), fg="#2E86AB")
            logo_label.place(relx=1.0, rely=0.02, anchor="ne", x=-5, y=5)

        # Заголовок
        Label_frame = ttk.Frame(main_frame)
        Label_frame.pack(fill=tk.X, pady=(0,5))
        ttk.Label(Label_frame, text="ОТЧЁТ", font=(font_face, 16, "bold")).pack(side=tk.LEFT, pady=(0, 0))
        ttk.Label(Label_frame, text="анализа содержимого файлов PDF", font=(font_face, 12)).pack(side=tk.LEFT, pady=(0, 5))
        
        # ИНСТРУКЦИЯ
        instr_frame = ttk.LabelFrame(main_frame, text="Инструкция", padding="10")
        instr_frame.pack(fill=tk.X, pady=(0, 15))
        
        instructions = """• Бросьте PDF файл или папку с PDF на ярлык для запуска анализа
• Создаётся Excel-файл с двумя листами: "Все страницы" и "Сводка ЕСКД"
• Форматы распознаются по ГОСТ 2.301-68 (A0, A1, A4×3 и т.д.)
• Цветность: Ч/Б или Цветная (для подбора принтера)
• Настройки хранятся в config.yaml (допуск, форматы)
    
    Инструмент разработан для Отдела выпуска компании СП-Инновация
    Автор: Родионов Вадим

Совет: выделите текст отчёта и нажмите Ctrl+C для копирования!"""
        
        instr_text = tk.Text(instr_frame, height=4, wrap=tk.WORD, font=(font_face, 9))
        scrollbar = ttk.Scrollbar(instr_frame, orient=tk.VERTICAL, command=instr_text.yview)
        instr_text.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        instr_text.insert(tk.END, instructions)
        instr_text.config(state=tk.DISABLED, bg="lightyellow")
        instr_text.pack(fill=tk.X)

        # СТАТИСТИКА (копируемый текст с выделением)
        stats_frame = ttk.LabelFrame(main_frame, text="Статистика обработки", padding="15")
        stats_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))

        # Текстовое поле С ВЫДЕЛЕНИЕМ
        stats_text = tk.Text(stats_frame, wrap=tk.WORD, font=(font_face, 10), 
                            height=12, bg="#f8f9fa", relief="solid", bd=1,
                            selectbackground="#4CAF50", selectforeground="white",
                            padx=10, pady=10)
        df, summary, out_path = self.last_result  # Сохраняем результат process_path в self.last_result                    
        scrollbar = ttk.Scrollbar(stats_frame, orient=tk.VERTICAL, command=stats_text.yview)
        stats_text.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        stats_text.pack(fill=tk.BOTH, expand=True)
    
        # Формируем отчёт с детальной сводкой форматов
        report_lines = []

        # Заголовок и базовая статистика
        report_lines.append("=== ОТЧЁТ АНАЛИЗА PDF ===")
        report_lines.append(f"Дата: {pd.Timestamp.now().strftime('%d.%m.%Y %H:%M')}")
        report_lines.append("")

        report_lines.append("📊 СТАТИСТИКА:")
        report_lines.append(f"  Файлов обработано: {self.stats['files_processed']}")
        report_lines.append(f"  Листов обработано: {self.stats['pages_processed']}")
        report_lines.append(f"  Файлов пропущено: {self.stats['files_skipped']}")
        report_lines.append(f"  Допуск форматов: {self.tolerance} мм")
        report_lines.append(f"  Форматов в БД: {len(self.formats)}")
        report_lines.append("")

        # СТАТИСТИКА ПО ФАЙЛАМ
        report_lines.append("📁 СТАТИСТИКА ПО ФАЙЛАМ:")

        file_stats = df.groupby("Файл").agg({
            "Страница": "count",
            "Цветность": lambda x: (x == "Цветная").sum()
        }).round(0).astype(int).reset_index()

        file_details = df.groupby("Файл")["Страница"].apply(
            lambda x: f"{len(x)} стр. (1-{max(x)})"
        ).to_dict()

        for _, row in file_stats.iterrows():
            file_name = row["Файл"]
            total_pages = row["Страница"]
            color_pages = row["Цветность"]
            page_range = file_details.get(file_name, "-")
            
            report_lines.append(f"  {file_name}: Всего {total_pages} стр., цветных: {color_pages}, диапазон: {page_range}")

        report_lines.append("")

        # СВОДКА ПО ФОРМАТАМ ПО ФАЙЛАМ
        report_lines.append("")
        report_lines.append("📐 ФОРМАТЫ ПО ФАЙЛАМ:")

        # Группируем по Файл
        for file_name, file_group in df.groupby("Файл"):

            report_lines.append(f"\n    ФАЙЛ: {file_name}:")        

             # Форматы внутри файла
            file_formats = file_group.groupby("Стандартный формат").agg({
                "Страница": "count"
            }).round(0).astype(int).reset_index()
            
            file_format_details = file_group.groupby("Стандартный формат")["Страница"].apply(
                lambda x: ",".join(map(str, sorted(x.tolist())))
            ).to_dict()
            
            for _, row in file_formats.iterrows():
                fmt = row["Стандартный формат"]
                total_pages = row["Страница"]
                pages_list = file_format_details.get(fmt, "-")
                sample_size = file_group[file_group["Стандартный формат"] == fmt]["Размер стандарта"].iloc[0]
                
                report_lines.append(f"      {fmt} {sample_size} ({total_pages} стр.):")       
                report_lines.append(f"          Страницы: {pages_list}")
                report_lines.append("")

        report_lines.append("")

        # Ошибки
        if self.stats["errors"]:
            report_lines.append(f"❌ ОШИБКИ ({len(self.stats['errors'])}):")
            for i, error in enumerate(self.stats["errors"], 1):
                report_lines.append(f"  {i}. {error}")
        else:
            report_lines.append("✅ Ошибок не обнаружено")
        report_lines.append("\n Подробный отчет в файле Excel")
        report_text = "\n".join(report_lines)

        # Вставляем отчёт
        stats_text.insert(tk.END, report_text)
        
        # ✅ РАЗРЕШАЕМ выделение и копирование
        stats_text.bind("<Control-c>", lambda e: stats_text.event_generate("<<Copy>>"))      # Ctrl+C
        stats_text.bind("<Button-3>", lambda e: stats_text.event_generate("<<Copy>>"))       # ПКМ меню
        stats_text.config(state=tk.DISABLED)    

        # Кнопки
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=10)

        def select_all():
            stats_text.config(state=tk.NORMAL)
            stats_text.tag_add("sel", "1.0", "end")
            stats_text.config(state=tk.DISABLED)

        ttk.Button(btn_frame, text="Выделить всё", command=select_all).pack(side=tk.LEFT, padx=(0, 5))
        
        def copy_to_clipboard():
            root.clipboard_clear()
            root.clipboard_append(report_text)
            messagebox.showinfo("Копирование", "Отчёт скопирован в буфер обмена!")

        ttk.Button(btn_frame, text="📋 Копировать отчёт", 
                  command=copy_to_clipboard).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(btn_frame, text="✅ Закрыть", 
                  command=root.destroy).pack(side=tk.RIGHT)

        # Горячие клавиши
        #root.bind("<Control-c>", lambda e: copy_to_clipboard())
        root.bind("<Escape>", lambda e: root.destroy())

        root.mainloop()

def main():
    if len(sys.argv) < 2:
        print("Использование: python pdfpages_pro.py путь_к_pdf_или_папке")
        return

    analyzer = PDFAnalyzer()
    input_path = sys.argv[1]

    try:
        df, summary, out_path = analyzer.process_path(input_path)
        print(f"\nГотово! XLSX: {out_path}")
        analyzer.show_report()
    except Exception as e:
        messagebox.showerror("Ошибка", f"Ошибка обработки: {str(e)}")

if __name__ == "__main__":
    main()
