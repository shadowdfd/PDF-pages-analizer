"""
Пасхалка: Игра "Тетрис" с форматами PDF листов
Активация: чит-код в окне помощи
"""
import tkinter as tk
from tkinter import messagebox
import random
from typing import List, Tuple

class PDFTetrisGame:
    """Тетрис с форматами PDF листов"""

    # Форматы PDF: название, цвет, форма (относительные координаты блоков)
    PDF_FORMATS = {
        "A4": {
            "color": "#38BDF8",
            "shape": [(0, 0), (1, 0), (0, 1), (1, 1)],  # Квадрат 2x2
            "points": 10
        },
        "A3": {
            "color": "#22D3EE",
            "shape": [(0, 0), (1, 0), (2, 0), (3, 0)],  # Линия 4x1
            "points": 40
        },
        "A2": {
            "color": "#4FD1C5",
            "shape": [(1, 0), (0, 1), (1, 1), (2, 1)],  # T-образная
            "points": 30
        },
        "A1": {
            "color": "#F97316",
            "shape": [(0, 0), (0, 1), (1, 1), (1, 2)],  # Z-образная
            "points": 30
        },
        "A0": {
            "color": "#F59E0B",
            "shape": [(1, 0), (1, 1), (0, 1), (0, 2)],  # S-образная
            "points": 30
        },
        "A3×3": {
            "color": "#C9A24B",
            "shape": [(0, 0), (1, 0), (1, 1), (2, 1)],  # L-образная
            "points": 30
        },
        "A3×4": {
            "color": "#A78BFA",
            "shape": [(2, 0), (0, 1), (1, 1), (2, 1)],  # J-образная
            "points": 30
        }
    }

    def __init__(self, parent=None):
        self.window = tk.Toplevel(parent) if parent else tk.Tk()
        self.window.title("📄 PDF Tetris - Форматы листов!")
        self.window.resizable(False, False)

        # Параметры игры
        self.cell_size = 30
        self.grid_width = 10
        self.grid_height = 20
        self.game_speed = 500  # мс между падениями

        # Состояние игры
        self.grid = [[None for _ in range(self.grid_width)] for _ in range(self.grid_height)]
        self.current_piece = None
        self.current_format = None
        self.current_pos = [0, 0]  # [x, y]
        self.score = 0
        self.lines_cleared = 0
        self.level = 1
        self.game_running = False
        self.game_over = False
        self.paused = False

        self.fall_timer = None

        self._build_ui()
        self._center_window()

        # Привязка клавиш
        self.window.bind("<Key>", self._on_key_press)
        self.window.bind("<Escape>", lambda e: self.window.destroy())

    def _build_ui(self):
        """Строит интерфейс игры"""

        main_frame = tk.Frame(self.window, bg="#0F172A")
        main_frame.pack()

        # Левая панель - игровое поле
        game_frame = tk.Frame(main_frame, bg="#0F172A")
        game_frame.pack(side=tk.LEFT, padx=10, pady=10)

        # Верхняя панель с информацией
        info_frame = tk.Frame(game_frame, bg="#0F172A", padx=10, pady=8)
        info_frame.pack(fill=tk.X)

        self.score_label = tk.Label(
            info_frame,
            text="Счёт: 0",
            font=("Segoe UI", 14, "bold"),
            fg="#F8FAFC",
            bg="#0F172A"
        )
        self.score_label.pack(side=tk.LEFT)

        self.level_label = tk.Label(
            info_frame,
            text="Уровень: 1",
            font=("Segoe UI", 12),
            fg="#CBD5E1",
            bg="#0F172A"
        )
        self.level_label.pack(side=tk.RIGHT)

        # Игровое поле
        canvas_width = self.grid_width * self.cell_size
        canvas_height = self.grid_height * self.cell_size

        self.canvas = tk.Canvas(
            game_frame,
            width=canvas_width,
            height=canvas_height,
            bg="#1E293B",
            highlightthickness=2,
            highlightbackground="#334155"
        )
        self.canvas.pack()

        # Нижняя панель с подсказками
        help_frame = tk.Frame(game_frame, bg="#0F172A", padx=10, pady=6)
        help_frame.pack(fill=tk.X)

        help_label = tk.Label(
            help_frame,
            text="← → поворот | ↓ ускорить | Space пауза",
            font=("Segoe UI", 9),
            fg="#94A3B8",
            bg="#0F172A"
        )
        help_label.pack()

        # Правая панель - статистика и следующая фигура
        right_frame = tk.Frame(main_frame, bg="#0F172A", padx=10, pady=10)
        right_frame.pack(side=tk.LEFT, fill=tk.Y)

        # Следующая фигура
        next_label = tk.Label(
            right_frame,
            text="Следующий формат:",
            font=("Segoe UI", 11, "bold"),
            fg="#F8FAFC",
            bg="#0F172A"
        )
        next_label.pack(pady=(0, 5))

        self.next_canvas = tk.Canvas(
            right_frame,
            width=120,
            height=120,
            bg="#1E293B",
            highlightthickness=1,
            highlightbackground="#334155"
        )
        self.next_canvas.pack(pady=(0, 20))

        # Статистика
        stats_frame = tk.Frame(right_frame, bg="#0F172A")
        stats_frame.pack(fill=tk.BOTH, expand=True)

        self.lines_label = tk.Label(
            stats_frame,
            text="Линий: 0",
            font=("Segoe UI", 11),
            fg="#CBD5E1",
            bg="#0F172A",
            anchor="w"
        )
        self.lines_label.pack(fill=tk.X, pady=2)

        # Легенда форматов
        legend_label = tk.Label(
            stats_frame,
            text="\nФорматы:",
            font=("Segoe UI", 10, "bold"),
            fg="#F8FAFC",
            bg="#0F172A",
            anchor="w"
        )
        legend_label.pack(fill=tk.X, pady=(10, 5))

        for name, data in list(self.PDF_FORMATS.items())[:5]:
            format_frame = tk.Frame(stats_frame, bg="#0F172A")
            format_frame.pack(fill=tk.X, pady=1)

            color_box = tk.Label(
                format_frame,
                text="  ",
                bg=data["color"],
                width=2
            )
            color_box.pack(side=tk.LEFT)

            format_label = tk.Label(
                format_frame,
                text=f" {name} = {data['points']}⭐",
                font=("Segoe UI", 9),
                fg="#94A3B8",
                bg="#0F172A",
                anchor="w"
            )
            format_label.pack(side=tk.LEFT, fill=tk.X)

        # Футер
        footer_label = tk.Label(
            right_frame,
            text="🎮 PDF Tetris",
            font=("Segoe UI", 9),
            fg="#64748B",
            bg="#0F172A"
        )
        footer_label.pack(side=tk.BOTTOM, pady=(20, 0))

    def _center_window(self):
        """Центрирует окно на экране"""
        self.window.update_idletasks()
        width = self.window.winfo_width()
        height = self.window.winfo_height()
        x = (self.window.winfo_screenwidth() // 2) - (width // 2)
        y = (self.window.winfo_screenheight() // 2) - (height // 2)
        self.window.geometry(f"+{x}+{y}")

    def start_game(self):
        """Запускает новую игру"""
        self.grid = [[None for _ in range(self.grid_width)] for _ in range(self.grid_height)]
        self.score = 0
        self.lines_cleared = 0
        self.level = 1
        self.game_speed = 500
        self.game_running = True
        self.game_over = False
        self.paused = False

        self.next_piece = random.choice(list(self.PDF_FORMATS.keys()))
        self._spawn_piece()
        self._update_stats()
        self._game_loop()

    def _spawn_piece(self):
        """Создаёт новую фигуру"""
        self.current_format = self.next_piece
        self.current_piece = self.PDF_FORMATS[self.current_format]["shape"][:]
        self.current_pos = [self.grid_width // 2 - 1, 0]

        # Выбираем следующую фигуру
        self.next_piece = random.choice(list(self.PDF_FORMATS.keys()))
        self._draw_next_piece()

        # Проверка game over
        if self._check_collision(self.current_piece, self.current_pos):
            self._game_over()

    def _draw_next_piece(self):
        """Отрисовывает следующую фигуру"""
        self.next_canvas.delete("all")

        next_shape = self.PDF_FORMATS[self.next_piece]["shape"]
        next_color = self.PDF_FORMATS[self.next_piece]["color"]

        # Центрируем фигуру в canvas
        min_x = min(x for x, y in next_shape)
        max_x = max(x for x, y in next_shape)
        min_y = min(y for x, y in next_shape)
        max_y = max(y for x, y in next_shape)

        width = (max_x - min_x + 1) * 20
        height = (max_y - min_y + 1) * 20

        offset_x = (120 - width) // 2 - min_x * 20
        offset_y = (120 - height) // 2 - min_y * 20

        for x, y in next_shape:
            self.next_canvas.create_rectangle(
                x * 20 + offset_x,
                y * 20 + offset_y,
                (x + 1) * 20 + offset_x,
                (y + 1) * 20 + offset_y,
                fill=next_color,
                outline="#F8FAFC",
                width=2
            )

    def _game_loop(self):
        """Основной игровой цикл"""
        if not self.game_running or self.game_over or self.paused:
            return

        # Двигаем фигуру вниз
        if self._move_piece(0, 1):
            # Фигура упала
            self._lock_piece()
            self._check_lines()
            self._spawn_piece()

        self._draw_game()

        # Следующий кадр
        self.fall_timer = self.window.after(self.game_speed, self._game_loop)

    def _move_piece(self, dx, dy):
        """Двигает текущую фигуру. Возвращает True если фигура не может двигаться дальше"""
        new_pos = [self.current_pos[0] + dx, self.current_pos[1] + dy]

        if self._check_collision(self.current_piece, new_pos):
            return dy != 0  # Возвращаем True только если двигались вниз

        self.current_pos = new_pos
        self._draw_game()
        return False

    def _rotate_piece(self):
        """Поворачивает текущую фигуру"""
        # Поворот на 90 градусов: (x, y) -> (-y, x)
        rotated = [(-y, x) for x, y in self.current_piece]

        # Нормализуем координаты (чтобы минимальные были 0)
        min_x = min(x for x, y in rotated)
        min_y = min(y for x, y in rotated)
        rotated = [(x - min_x, y - min_y) for x, y in rotated]

        if not self._check_collision(rotated, self.current_pos):
            self.current_piece = rotated
            self._draw_game()

    def _check_collision(self, piece, pos):
        """Проверяет столкновение фигуры с границами и другими блоками"""
        for dx, dy in piece:
            x = pos[0] + dx
            y = pos[1] + dy

            # Проверка границ
            if x < 0 or x >= self.grid_width or y >= self.grid_height:
                return True

            # Проверка столкновения с другими блоками
            if y >= 0 and self.grid[y][x] is not None:
                return True

        return False

    def _lock_piece(self):
        """Фиксирует фигуру на поле"""
        color = self.PDF_FORMATS[self.current_format]["color"]

        for dx, dy in self.current_piece:
            x = self.current_pos[0] + dx
            y = self.current_pos[1] + dy
            if y >= 0:
                self.grid[y][x] = (self.current_format, color)

    def _check_lines(self):
        """Проверяет и удаляет заполненные линии"""
        lines_to_clear = []

        for y in range(self.grid_height):
            if all(cell is not None for cell in self.grid[y]):
                lines_to_clear.append(y)

        if lines_to_clear:
            # Показываем анимацию мигания перед удалением
            self._animate_line_clear(lines_to_clear)

    def _animate_line_clear(self, lines_to_clear, blink_count=0):
        """Анимирует мигание линий перед их удалением"""
        max_blinks = 6  # 3 полных цикла мигания (вкл-выкл)

        if blink_count < max_blinks:
            # Переключаем видимость линий
            if blink_count % 2 == 0:
                # Подсвечиваем линии белым
                for y in lines_to_clear:
                    for x in range(self.grid_width):
                        if self.grid[y][x] is not None:
                            format_name, _ = self.grid[y][x]
                            self.grid[y][x] = (format_name, "#F8FAFC")  # Белый цвет
            else:
                # Возвращаем оригинальные цвета
                for y in lines_to_clear:
                    for x in range(self.grid_width):
                        if self.grid[y][x] is not None:
                            format_name, _ = self.grid[y][x]
                            original_color = self.PDF_FORMATS[format_name]["color"]
                            self.grid[y][x] = (format_name, original_color)

            self._draw_game()

            # Следующий кадр анимации через 80 мс
            self.window.after(80, lambda: self._animate_line_clear(lines_to_clear, blink_count + 1))
        else:
            # Анимация завершена - удаляем линии
            self._remove_lines(lines_to_clear)

    def _remove_lines(self, lines_to_clear):
        """Удаляет линии и обновляет счёт"""
        # Удаляем линии
        for y in lines_to_clear:
            del self.grid[y]
            self.grid.insert(0, [None for _ in range(self.grid_width)])

        # Обновляем счёт
        self.lines_cleared += len(lines_to_clear)

        # Очки за линии (с бонусом за несколько линий сразу)
        points = [0, 100, 300, 500, 800]
        self.score += points[min(len(lines_to_clear), 4)] * self.level

        # Повышаем уровень каждые 10 линий
        new_level = (self.lines_cleared // 10) + 1
        if new_level > self.level:
            self.level = new_level
            self.game_speed = max(100, 500 - (self.level - 1) * 50)

        self._update_stats()
        self._draw_game()

    def _draw_game(self):
        """Отрисовывает игровое поле"""
        self.canvas.delete("all")

        # Рисуем сетку
        for i in range(self.grid_width + 1):
            x = i * self.cell_size
            self.canvas.create_line(x, 0, x, self.grid_height * self.cell_size, fill="#334155", width=1)
        for i in range(self.grid_height + 1):
            y = i * self.cell_size
            self.canvas.create_line(0, y, self.grid_width * self.cell_size, y, fill="#334155", width=1)

        # Рисуем зафиксированные блоки
        for y in range(self.grid_height):
            for x in range(self.grid_width):
                if self.grid[y][x] is not None:
                    format_name, color = self.grid[y][x]
                    self._draw_block(x, y, color)

        # Рисуем текущую фигуру
        if self.current_piece:
            color = self.PDF_FORMATS[self.current_format]["color"]
            for dx, dy in self.current_piece:
                x = self.current_pos[0] + dx
                y = self.current_pos[1] + dy
                if y >= 0:
                    self._draw_block(x, y, color)

    def _draw_block(self, x, y, color):
        """Рисует один блок"""
        self.canvas.create_rectangle(
            x * self.cell_size + 2,
            y * self.cell_size + 2,
            (x + 1) * self.cell_size - 2,
            (y + 1) * self.cell_size - 2,
            fill=color,
            outline="#F8FAFC",
            width=1
        )

    def _update_stats(self):
        """Обновляет статистику"""
        self.score_label.config(text=f"Счёт: {self.score}")
        self.level_label.config(text=f"Уровень: {self.level}")
        self.lines_label.config(text=f"Линий: {self.lines_cleared}")

    def _game_over(self):
        """Конец игры"""
        self.game_over = True
        self.game_running = False

        if self.fall_timer:
            self.window.after_cancel(self.fall_timer)

        # Затемнение поля
        self.canvas.create_rectangle(
            0, 0,
            self.grid_width * self.cell_size,
            self.grid_height * self.cell_size,
            fill="#0F172A",
            stipple="gray50",
            tags="overlay"
        )

        # Сообщение Game Over
        cx = (self.grid_width * self.cell_size) // 2
        cy = (self.grid_height * self.cell_size) // 2

        self.canvas.create_text(
            cx, cy - 60,
            text="GAME OVER",
            font=("Segoe UI", 32, "bold"),
            fill="#F87171",
            tags="overlay"
        )

        self.canvas.create_text(
            cx, cy - 10,
            text=f"Счёт: {self.score}",
            font=("Segoe UI", 18),
            fill="#F8FAFC",
            tags="overlay"
        )

        self.canvas.create_text(
            cx, cy + 20,
            text=f"Линий: {self.lines_cleared}",
            font=("Segoe UI", 14),
            fill="#CBD5E1",
            tags="overlay"
        )

        self.canvas.create_text(
            cx, cy + 50,
            text=f"Уровень: {self.level}",
            font=("Segoe UI", 14),
            fill="#CBD5E1",
            tags="overlay"
        )

        self.canvas.create_text(
            cx, cy + 90,
            text="Нажмите R для рестарта",
            font=("Segoe UI", 12),
            fill="#94A3B8",
            tags="overlay"
        )

    def _on_key_press(self, event):
        """Обработка нажатий клавиш"""
        if self.game_over:
            if event.keysym.lower() == "r":
                self.start_game()
            return

        if not self.game_running:
            return

        key = event.keysym

        # Управление
        if key == "Left":
            self._move_piece(-1, 0)
        elif key == "Right":
            self._move_piece(1, 0)
        elif key == "Down":
            self._move_piece(0, 1)

        # Поворот фигуры
        elif key == "space":
            self._rotate_piece()

        # Пауза
        elif key == "Return":
            if self.paused:
                self.paused = False
                self.canvas.delete("pause")
                self._game_loop()
            else:
                self.paused = True
                if self.fall_timer:
                    self.window.after_cancel(self.fall_timer)
                self.canvas.create_text(
                    (self.grid_width * self.cell_size) // 2,
                    (self.grid_height * self.cell_size) // 2,
                    text="ПАУЗА",
                    font=("Segoe UI", 32, "bold"),
                    fill="#F8FAFC",
                    tags="pause"
                )

    def run(self):
        """Запускает окно игры"""
        self._show_welcome()
        self.window.mainloop()

    def _show_welcome(self):
        """Показывает приветственный экран"""
        cx = (self.grid_width * self.cell_size) // 2
        cy = (self.grid_height * self.cell_size) // 2

        self.canvas.create_text(
            cx, cy - 80,
            text="📄 PDF TETRIS",
            font=("Segoe UI", 28, "bold"),
            fill="#38BDF8"
        )

        self.canvas.create_text(
            cx, cy - 30,
            text="Собирай форматы PDF листов!",
            font=("Segoe UI", 12),
            fill="#CBD5E1"
        )

        self.canvas.create_text(
            cx, cy + 10,
            text="← → движение | ↓ ускорить",
            font=("Segoe UI", 10),
            fill="#94A3B8"
        )

        self.canvas.create_text(
            cx, cy + 30,
            text="Space поворот | Enter пауза",
            font=("Segoe UI", 10),
            fill="#94A3B8"
        )

        self.canvas.create_text(
            cx, cy + 70,
            text="Нажмите любую стрелку для старта",
            font=("Segoe UI", 12, "bold"),
            fill="#F8FAFC"
        )

        # Ждём нажатия стрелки для старта
        def start_on_arrow(event):
            if event.keysym in ["Up", "Down", "Left", "Right"]:
                self.window.unbind("<Key>", bind_id)
                self.window.bind("<Key>", self._on_key_press)
                self.start_game()

        bind_id = self.window.bind("<Key>", start_on_arrow)


def launch_tetris_game(parent=None):
    """Запускает игру тетрис"""
    game = PDFTetrisGame(parent)
    game.run()


if __name__ == "__main__":
    launch_tetris_game()
