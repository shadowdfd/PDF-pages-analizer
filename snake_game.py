"""
Пасхалка: Игра "Змейка" - собирай форматы PDF листов
Активация: Ctrl+Shift+S в главном окне
"""
import tkinter as tk
from tkinter import messagebox
import random
from typing import List, Tuple

class PDFSnakeGame:
    """Змейка собирает форматы PDF листов"""

    # Форматы PDF для сбора (цвет, название, очки)
    PDF_FORMATS = [
        ("#38BDF8", "A4", 10),   # Sky blue
        ("#22D3EE", "A3", 20),   # Cyan
        ("#4FD1C5", "A2", 30),   # Teal
        ("#F97316", "A1", 40),   # Orange
        ("#F59E0B", "A0", 50),   # Amber
    ]

    def __init__(self, parent=None):
        self.window = tk.Toplevel(parent) if parent else tk.Tk()
        self.window.title("🐍 PDF Snake Game - Собирай форматы листов!")
        self.window.resizable(False, False)

        # Параметры игры
        self.cell_size = 20
        self.grid_width = 30
        self.grid_height = 25
        self.game_speed = 150  # мс между ходами

        # Состояние игры
        self.snake: List[Tuple[int, int]] = []
        self.direction = "Right"
        self.next_direction = "Right"
        self.food_pos = (0, 0)
        self.food_format = None
        self.score = 0
        self.game_running = False
        self.game_over = False

        self._build_ui()
        self._center_window()

        # Привязка клавиш
        self.window.bind("<Key>", self._on_key_press)
        self.window.bind("<Escape>", lambda e: self.window.destroy())

    def _build_ui(self):
        """Строит интерфейс игры"""

        # Верхняя панель с информацией
        info_frame = tk.Frame(self.window, bg="#0F172A", padx=10, pady=8)
        info_frame.pack(fill=tk.X)

        self.score_label = tk.Label(
            info_frame,
            text="Счёт: 0",
            font=("Segoe UI", 14, "bold"),
            fg="#F8FAFC",
            bg="#0F172A"
        )
        self.score_label.pack(side=tk.LEFT)

        self.speed_label = tk.Label(
            info_frame,
            text="Скорость: 1x",
            font=("Segoe UI", 10),
            fg="#CBD5E1",
            bg="#0F172A"
        )
        self.speed_label.pack(side=tk.LEFT, padx=(20, 0))

        help_label = tk.Label(
            info_frame,
            text="Стрелки - управление | Space - пауза | R - рестарт",
            font=("Segoe UI", 9),
            fg="#94A3B8",
            bg="#0F172A"
        )
        help_label.pack(side=tk.RIGHT)

        # Игровое поле
        canvas_width = self.grid_width * self.cell_size
        canvas_height = self.grid_height * self.cell_size

        self.canvas = tk.Canvas(
            self.window,
            width=canvas_width,
            height=canvas_height,
            bg="#1E293B",
            highlightthickness=0
        )
        self.canvas.pack()

        # Нижняя панель
        footer_frame = tk.Frame(self.window, bg="#0F172A", padx=10, pady=6)
        footer_frame.pack(fill=tk.X)

        footer_label = tk.Label(
            footer_frame,
            text="🎮 Пасхалка от разработчиков PDF Analyzer",
            font=("Segoe UI", 9),
            fg="#64748B",
            bg="#0F172A"
        )
        footer_label.pack()

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
        # Инициализация змейки в центре
        center_x = self.grid_width // 2
        center_y = self.grid_height // 2
        self.snake = [
            (center_x, center_y),
            (center_x - 1, center_y),
            (center_x - 2, center_y)
        ]

        self.direction = "Right"
        self.next_direction = "Right"
        self.score = 0
        self.game_speed = 150
        self.game_running = True
        self.game_over = False

        self._spawn_food()
        self._update_score()
        self._game_loop()

    def _spawn_food(self):
        """Создаёт новую еду (формат PDF)"""
        while True:
            x = random.randint(0, self.grid_width - 1)
            y = random.randint(0, self.grid_height - 1)
            if (x, y) not in self.snake:
                self.food_pos = (x, y)
                self.food_format = random.choice(self.PDF_FORMATS)
                break

    def _game_loop(self):
        """Основной игровой цикл"""
        if not self.game_running or self.game_over:
            return

        self._move_snake()
        self._check_collision()
        self._draw_game()

        # Следующий кадр
        self.window.after(self.game_speed, self._game_loop)

    def _move_snake(self):
        """Двигает змейку"""
        # Обновляем направление
        self.direction = self.next_direction

        # Вычисляем новую позицию головы
        head_x, head_y = self.snake[0]

        if self.direction == "Up":
            new_head = (head_x, head_y - 1)
        elif self.direction == "Down":
            new_head = (head_x, head_y + 1)
        elif self.direction == "Left":
            new_head = (head_x - 1, head_y)
        else:  # Right
            new_head = (head_x + 1, head_y)

        # Проверка выхода за границы (телепорт на другую сторону)
        new_head = (
            new_head[0] % self.grid_width,
            new_head[1] % self.grid_height
        )

        # Добавляем новую голову
        self.snake.insert(0, new_head)

        # Проверяем, съели ли еду
        if new_head == self.food_pos:
            color, format_name, points = self.food_format
            self.score += points
            self._update_score()
            self._spawn_food()

            # Увеличиваем скорость каждые 100 очков
            if self.score % 100 == 0 and self.game_speed > 50:
                self.game_speed = max(50, self.game_speed - 15)
                speed_mult = round((150 / self.game_speed), 1)
                self.speed_label.config(text=f"Скорость: {speed_mult}x")
        else:
            # Убираем хвост, если не съели еду
            self.snake.pop()

    def _check_collision(self):
        """Проверяет столкновения"""
        head = self.snake[0]

        # Столкновение с собой
        if head in self.snake[1:]:
            self._game_over()

    def _game_over(self):
        """Конец игры"""
        self.game_over = True
        self.game_running = False

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
            cx, cy - 40,
            text="GAME OVER",
            font=("Segoe UI", 36, "bold"),
            fill="#F87171",
            tags="overlay"
        )

        self.canvas.create_text(
            cx, cy + 10,
            text=f"Счёт: {self.score}",
            font=("Segoe UI", 20),
            fill="#F8FAFC",
            tags="overlay"
        )

        self.canvas.create_text(
            cx, cy + 50,
            text="Нажмите R для рестарта",
            font=("Segoe UI", 14),
            fill="#CBD5E1",
            tags="overlay"
        )

    def _draw_game(self):
        """Отрисовывает игровое поле"""
        self.canvas.delete("all")

        # Рисуем сетку (опционально)
        # for i in range(self.grid_width):
        #     x = i * self.cell_size
        #     self.canvas.create_line(x, 0, x, self.grid_height * self.cell_size, fill="#334155", width=1)
        # for i in range(self.grid_height):
        #     y = i * self.cell_size
        #     self.canvas.create_line(0, y, self.grid_width * self.cell_size, y, fill="#334155", width=1)

        # Рисуем еду (формат PDF)
        if self.food_format:
            fx, fy = self.food_pos
            color, format_name, points = self.food_format

            # Квадрат формата
            self.canvas.create_rectangle(
                fx * self.cell_size + 2,
                fy * self.cell_size + 2,
                (fx + 1) * self.cell_size - 2,
                (fy + 1) * self.cell_size - 2,
                fill=color,
                outline="#F8FAFC",
                width=2
            )

            # Название формата
            self.canvas.create_text(
                fx * self.cell_size + self.cell_size // 2,
                fy * self.cell_size + self.cell_size // 2,
                text=format_name,
                font=("Segoe UI", 9, "bold"),
                fill="#0F172A"
            )

        # Рисуем змейку
        for i, (x, y) in enumerate(self.snake):
            if i == 0:
                # Голова - ярче
                fill_color = "#38BDF8"
                outline_color = "#F8FAFC"
            else:
                # Тело - градиент
                alpha = 1 - (i / len(self.snake)) * 0.5
                fill_color = "#22D3EE" if alpha > 0.7 else "#4FD1C5"
                outline_color = "#1E293B"

            self.canvas.create_rectangle(
                x * self.cell_size + 1,
                y * self.cell_size + 1,
                (x + 1) * self.cell_size - 1,
                (y + 1) * self.cell_size - 1,
                fill=fill_color,
                outline=outline_color,
                width=1
            )

    def _update_score(self):
        """Обновляет счёт"""
        self.score_label.config(text=f"Счёт: {self.score}")

    def _on_key_press(self, event):
        """Обработка нажатий клавиш"""
        key = event.keysym

        # Управление направлением
        if key == "Up" and self.direction != "Down":
            self.next_direction = "Up"
        elif key == "Down" and self.direction != "Up":
            self.next_direction = "Down"
        elif key == "Left" and self.direction != "Right":
            self.next_direction = "Left"
        elif key == "Right" and self.direction != "Left":
            self.next_direction = "Right"

        # Пауза
        elif key == "space":
            if self.game_running and not self.game_over:
                self.game_running = False
                self.canvas.create_text(
                    (self.grid_width * self.cell_size) // 2,
                    (self.grid_height * self.cell_size) // 2,
                    text="ПАУЗА",
                    font=("Segoe UI", 32, "bold"),
                    fill="#F8FAFC",
                    tags="pause"
                )
            elif not self.game_over:
                self.canvas.delete("pause")
                self.game_running = True
                self._game_loop()

        # Рестарт
        elif key.lower() == "r":
            self.start_game()

    def run(self):
        """Запускает окно игры"""
        # Показываем приветствие
        self._show_welcome()
        self.window.mainloop()

    def _show_welcome(self):
        """Показывает приветственный экран"""
        cx = (self.grid_width * self.cell_size) // 2
        cy = (self.grid_height * self.cell_size) // 2

        self.canvas.create_text(
            cx, cy - 60,
            text="🐍 PDF SNAKE",
            font=("Segoe UI", 32, "bold"),
            fill="#38BDF8"
        )

        self.canvas.create_text(
            cx, cy - 10,
            text="Собирай форматы PDF листов!",
            font=("Segoe UI", 14),
            fill="#CBD5E1"
        )

        formats_text = "  ".join([f"{name} = {pts}⭐" for _, name, pts in self.PDF_FORMATS])
        self.canvas.create_text(
            cx, cy + 30,
            text=formats_text,
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


def launch_snake_game(parent=None):
    """Запускает игру змейка"""
    game = PDFSnakeGame(parent)
    game.run()


if __name__ == "__main__":
    launch_snake_game()
