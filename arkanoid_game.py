"""
PDF Arkanoid - арканойд с тематикой PDF форматов
Пасхалка для PDF Pages Analyzer
"""

import tkinter as tk
from tkinter import messagebox
import random
import math
from highscores import get_highscore, save_highscore


def _normalize_key(event):
    """Возвращает имя клавиши независимо от раскладки клавиатуры."""
    keysym = getattr(event, "keysym", "") or ""
    if keysym in {"Up", "Down", "Left", "Right", "space", "Return", "Escape"}:
        return keysym

    key = keysym.lower()
    char = (getattr(event, "char", "") or "").lower()
    return {
        "r": "r",
        "к": "r",
        "cyrillic_ka": "r",
    }.get(char or key, char or key)


class PDFArkanoidGame:
    def __init__(self, parent=None):
        self.parent = parent

        # Создаём окно
        self.window = tk.Toplevel(parent) if parent else tk.Tk()
        self.window.title("📄 PDF Arkanoid")
        self.window.configure(bg="#0F172A")
        self.window.resizable(False, False)

        # Размеры игрового поля
        self.width = 635
        self.height = 700
        self.paddle_width = 100
        self.paddle_height = 15
        self.ball_size = 12
        self.block_rows = 6
        self.block_cols = 10
        self.block_width = 54
        self.block_height = 25
        self.block_padding = 6

        # Игровые переменные
        self.paddle_x = self.width // 2 - self.paddle_width // 2
        self.paddle_y = self.height - 50
        self.ball_x = self.width // 2
        self.ball_y = self.height - 70
        self.ball_dx = 4
        self.ball_dy = -4
        self.score = 0
        self.highscore = get_highscore("arkanoid")
        self.lives = 3
        self.level = 1
        self.game_running = False
        self.game_over = False
        self.paused = False
        self.blocks = []
        self.game_timer = None
        self.life_explosion_timer = None
        self.life_restart_timer = None
        self.life_explosion_frame = 0

        # Форматы PDF для блоков (цвета и очки)
        self.PDF_FORMATS = [
            {"name": "A0", "color": "#F87171", "points": 50},  # Красный
            {"name": "A1", "color": "#FB923C", "points": 40},  # Оранжевый
            {"name": "A2", "color": "#FBBF24", "points": 30},  # Жёлтый
            {"name": "A3", "color": "#34D399", "points": 20},  # Зелёный
            {"name": "A4", "color": "#38BDF8", "points": 15},  # Голубой
            {"name": "A5", "color": "#A78BFA", "points": 10},  # Фиолетовый
        ]

        # Создаём интерфейс
        self._create_ui()

        # Центрируем окно
        self._center_window()

        # Привязываем события
        self.window.bind("<Key>", self._on_key_press)
        self.window.bind("<Motion>", self._on_mouse_move)
        self.canvas.bind("<Button-1>", self._on_click)

        # Обработка закрытия окна
        self.window.protocol("WM_DELETE_WINDOW", self._on_close)

    def _create_ui(self):
        """Создаёт интерфейс игры"""
        # Фрейм для статистики
        stats_frame = tk.Frame(self.window, bg="#1E293B", padx=20, pady=10)
        stats_frame.pack(fill=tk.X)

        # Счёт
        self.score_label = tk.Label(
            stats_frame,
            text=f"Счёт: {self.score}",
            font=("Segoe UI", 14, "bold"),
            fg="#F8FAFC",
            bg="#1E293B"
        )
        self.score_label.pack(side=tk.LEFT, padx=10)

        self.highscore_label = tk.Label(
            stats_frame,
            text=f"Рекорд: {self.highscore}",
            font=("Segoe UI", 10),
            fg="#FBBF24",
            bg="#1E293B",
        )
        self.highscore_label.pack(side=tk.LEFT, padx=10)

        # Жизни
        self.lives_label = tk.Label(
            stats_frame,
            text=f"❤️ × {self.lives}",
            font=("Segoe UI", 14, "bold"),
            fg="#F87171",
            bg="#1E293B"
        )
        self.lives_label.pack(side=tk.LEFT, padx=10)

        # Уровень
        self.level_label = tk.Label(
            stats_frame,
            text=f"Уровень: {self.level}",
            font=("Segoe UI", 14, "bold"),
            fg="#38BDF8",
            bg="#1E293B"
        )
        self.level_label.pack(side=tk.RIGHT, padx=10)

        # Canvas для игры
        self.canvas = tk.Canvas(
            self.window,
            width=self.width,
            height=self.height,
            bg="#0F172A",
            highlightthickness=0
        )
        self.canvas.pack()

    def _center_window(self):
        """Центрирует окно на экране"""
        self.window.update_idletasks()
        width = self.window.winfo_width()
        height = self.window.winfo_height()
        x = (self.window.winfo_screenwidth() // 2) - (width // 2)
        y = (self.window.winfo_screenheight() // 2) - (height // 2)
        self.window.geometry(f"+{x}+{y}")

    def _on_close(self):
        """Обработка закрытия окна"""
        self.game_running = False
        self._cancel_game_timer()
        self._cancel_life_explosion()
        self.window.destroy()

    def start_game(self):
        """Запускает новую игру"""
        # Отменяем старый callback перед запуском новой игровой цепочки.
        self._cancel_game_timer()
        self._cancel_life_explosion()
        self.score = 0
        self.lives = 3
        self.level = 1
        self.game_over = False
        self.paused = False
        self._start_level()

    def _cancel_game_timer(self):
        """Отменяет запланированный кадр игры, если он есть."""
        if self.game_timer is not None:
            try:
                self.window.after_cancel(self.game_timer)
            except tk.TclError:
                # Таймер уже мог выполниться или окно могло закрыться.
                pass
            finally:
                self.game_timer = None

    def _start_level(self):
        """Запускает новый уровень"""
        self.game_running = False
        self.paddle_x = self.width // 2 - self.paddle_width // 2
        self.ball_x = self.width // 2
        self.ball_y = self.height - 70

        # Скорость шарика увеличивается с уровнем
        speed = 8 + (self.level - 2) * 0.5
        angle = random.uniform(-60, -120) * math.pi / 180
        self.ball_dx = speed * math.cos(angle)
        self.ball_dy = speed * math.sin(angle)

        # Создаём блоки
        self._create_blocks()

        # Обновляем статистику
        self._update_stats()

        # Показываем сообщение о старте уровня
        self._draw_game()
        cx = self.width // 2
        cy = self.height // 2

        self.canvas.create_text(
            cx, cy - 30,
            text=f"УРОВЕНЬ {self.level}",
            font=("Segoe UI", 32, "bold"),
            fill="#38BDF8",
            tags="level_msg"
        )

        self.canvas.create_text(
            cx, cy + 20,
            text="Кликните для старта",
            font=("Segoe UI", 14),
            fill="#CBD5E1",
            tags="level_msg"
        )

    def _create_blocks(self):
        """Создаёт блоки для уровня"""
        self.blocks = []
        start_y = 80

        for row in range(self.block_rows):
            format_data = self.PDF_FORMATS[row % len(self.PDF_FORMATS)]
            for col in range(self.block_cols):
                x = col * (self.block_width + self.block_padding) + 20
                y = row * (self.block_height + self.block_padding) + start_y

                block = {
                    "x": x,
                    "y": y,
                    "width": self.block_width,
                    "height": self.block_height,
                    "color": format_data["color"],
                    "format": format_data["name"],
                    "points": format_data["points"],
                    "hits": 1 + (self.level - 1) // 3  # Прочность растёт с уровнем
                }
                self.blocks.append(block)

    def _game_loop(self):
        """Главный игровой цикл"""
        # Callback уже сработал; следующий будет назначен в конце кадра.
        self.game_timer = None
        if not self.game_running or self.paused:
            return

        # Движение шарика
        self.ball_x += self.ball_dx
        self.ball_y += self.ball_dy

        # Отражение от стен
        if self.ball_x <= self.ball_size // 2:
            self.ball_x = self.ball_size // 2
            self.ball_dx = abs(self.ball_dx)
        elif self.ball_x >= self.width - self.ball_size // 2:
            self.ball_x = self.width - self.ball_size // 2
            self.ball_dx = -abs(self.ball_dx)

        if self.ball_y <= self.ball_size // 2:
            self.ball_y = self.ball_size // 2
            self.ball_dy = abs(self.ball_dy)

        # Проверка падения шарика
        if self.ball_y >= self.height:
            self._lose_life()
            return

        # Столкновение с платформой
        if (self.paddle_y <= self.ball_y + self.ball_size // 2 <= self.paddle_y + self.paddle_height and
            self.paddle_x <= self.ball_x <= self.paddle_x + self.paddle_width):

            # Отражаем шарик
            self.ball_dy = -abs(self.ball_dy)
            self.ball_y = self.paddle_y - self.ball_size // 2

            # Изменяем угол отражения в зависимости от места удара
            hit_pos = (self.ball_x - self.paddle_x) / self.paddle_width  # 0 to 1
            angle_offset = (hit_pos - 0.5) * 60  # -30 to +30 degrees
            speed = math.sqrt(self.ball_dx**2 + self.ball_dy**2)
            angle = -90 + angle_offset
            self.ball_dx = speed * math.cos(angle * math.pi / 180)
            self.ball_dy = speed * math.sin(angle * math.pi / 180)

        # Столкновение с блоками
        self._check_block_collision()

        # Проверка победы на уровне
        if not self.blocks:
            self._next_level()
            return

        # Отрисовка
        self._draw_game()

        # Следующий кадр
        if self.game_running and not self.paused and not self.game_over:
            self.game_timer = self.window.after(16, self._game_loop)  # ~60 FPS

    def _check_block_collision(self):
        """Проверяет столкновение с блоками"""
        ball_rect = {
            "left": self.ball_x - self.ball_size // 2,
            "right": self.ball_x + self.ball_size // 2,
            "top": self.ball_y - self.ball_size // 2,
            "bottom": self.ball_y + self.ball_size // 2
        }

        for block in self.blocks[:]:
            block_rect = {
                "left": block["x"],
                "right": block["x"] + block["width"],
                "top": block["y"],
                "bottom": block["y"] + block["height"]
            }

            # Проверка пересечения
            if (ball_rect["right"] >= block_rect["left"] and
                ball_rect["left"] <= block_rect["right"] and
                ball_rect["bottom"] >= block_rect["top"] and
                ball_rect["top"] <= block_rect["bottom"]):

                # Определяем сторону столкновения
                overlap_left = ball_rect["right"] - block_rect["left"]
                overlap_right = block_rect["right"] - ball_rect["left"]
                overlap_top = ball_rect["bottom"] - block_rect["top"]
                overlap_bottom = block_rect["bottom"] - ball_rect["top"]

                min_overlap = min(overlap_left, overlap_right, overlap_top, overlap_bottom)

                # Отражаем шарик
                if min_overlap == overlap_left or min_overlap == overlap_right:
                    self.ball_dx = -self.ball_dx
                else:
                    self.ball_dy = -self.ball_dy

                # Уменьшаем прочность блока
                block["hits"] -= 1

                if block["hits"] <= 0:
                    # Блок разрушен
                    self.score += block["points"] * self.level
                    self.blocks.remove(block)
                else:
                    # Блок ещё держится, меняем цвет
                    block["color"] = self._darken_color(block["color"])

                self._update_stats()
                break

    def _darken_color(self, color):
        """Затемняет цвет для повреждённого блока"""
        # Простое затемнение через hex
        r = int(color[1:3], 16)
        g = int(color[3:5], 16)
        b = int(color[5:7], 16)

        r = max(0, r - 40)
        g = max(0, g - 40)
        b = max(0, b - 40)

        return f"#{r:02x}{g:02x}{b:02x}"

    def _lose_life(self):
        """Потеря жизни"""
        if self.life_explosion_timer is not None:
            return
        self.life_explosion_frame = 0
        self._draw_life_explosion()
        self.lives -= 1
        self._update_stats()

        # Дожидаемся окончания взрыва и затем либо завершаем игру,
        # либо создаём новый шарик.
        self.game_running = False
        self.life_restart_timer = self.window.after(
            280, self._restart_after_life_loss
        )

    def _cancel_life_explosion(self):
        """Отменяет анимацию взрыва после потери жизни."""
        for timer_name in ("life_explosion_timer", "life_restart_timer"):
            timer_id = getattr(self, timer_name)
            if timer_id is None:
                continue
            try:
                self.window.after_cancel(timer_id)
            except tk.TclError:
                pass
            finally:
                setattr(self, timer_name, None)

    def _draw_life_explosion(self):
        """Рисует взрыв в точке потери шарика."""
        self.canvas.delete("game")
        cx = self.ball_x
        cy = self.ball_y
        radius = 10 + self.life_explosion_frame * 5
        colors = ["#FBBF24", "#FB923C", "#F87171", "#7C2D12"]
        color = colors[min(self.life_explosion_frame // 2, len(colors) - 1)]

        self.canvas.create_oval(
            cx - radius, cy - radius,
            cx + radius, cy + radius,
            fill=color,
            outline="#F8FAFC",
            width=2,
            tags=("game", "life_explosion")
        )

        if self.life_explosion_frame < 5:
            self.life_explosion_frame += 1
            self.life_explosion_timer = self.window.after(
                45, self._draw_life_explosion
            )
        else:
            self.life_explosion_timer = None

    def _restart_after_life_loss(self):
        """Возобновляет игру после анимации потери жизни."""
        self.life_restart_timer = None
        self.life_explosion_timer = None
        self.canvas.delete("life_explosion")
        if self.lives <= 0:
            self._game_over()
            return

        self.ball_x = self.width // 2
        self.ball_y = self.height - 70
        self.paddle_x = self.width // 2 - self.paddle_width // 2

        angle = random.uniform(-60, -120) * math.pi / 180
        speed = 8 + (self.level - 2) * 0.5
        self.ball_dx = speed * math.cos(angle)
        self.ball_dy = speed * math.sin(angle)
        self._draw_game()

    def _next_level(self):
        """Переход на следующий уровень"""
        self.level += 1
        self.lives = min(5, self.lives + 1)  # Бонусная жизнь, макс 5
        self._start_level()

    def _draw_game(self):
        """Отрисовывает игровое поле"""
        self.canvas.delete("game")

        # Рисуем блоки
        for block in self.blocks:
            self.canvas.create_rectangle(
                block["x"], block["y"],
                block["x"] + block["width"],
                block["y"] + block["height"],
                fill=block["color"],
                outline="#F8FAFC",
                width=1,
                tags="game"
            )

            # Текст формата на блоке
            self.canvas.create_text(
                block["x"] + block["width"] // 2,
                block["y"] + block["height"] // 2,
                text=block["format"],
                font=("Segoe UI", 9, "bold"),
                fill="#0F172A",
                tags="game"
            )

        # Рисуем платформу
        self.canvas.create_rectangle(
            self.paddle_x, self.paddle_y,
            self.paddle_x + self.paddle_width,
            self.paddle_y + self.paddle_height,
            fill="#38BDF8",
            outline="#F8FAFC",
            width=2,
            tags="game"
        )

        # Рисуем шарик
        self.canvas.create_oval(
            self.ball_x - self.ball_size // 2,
            self.ball_y - self.ball_size // 2,
            self.ball_x + self.ball_size // 2,
            self.ball_y + self.ball_size // 2,
            fill="#FBBF24",
            outline="#F8FAFC",
            width=2,
            tags="game"
        )

    def _update_stats(self):
        """Обновляет статистику"""
        self.score_label.config(text=f"Счёт: {self.score}")
        if self.score > self.highscore:
            self.highscore = save_highscore("arkanoid", self.score)
            self.highscore_label.config(text=f"Рекорд: {self.highscore}")
        self.lives_label.config(text=f"❤️ × {self.lives}")
        self.level_label.config(text=f"Уровень: {self.level}")

    def _game_over(self):
        """Конец игры"""
        self.game_over = True
        self.game_running = False
        self._cancel_game_timer()
        self.highscore = save_highscore("arkanoid", self.score)
        self.highscore_label.config(text=f"Рекорд: {self.highscore}")

        # Затемнение поля
        self.canvas.create_rectangle(
            0, 0, self.width, self.height,
            fill="#0F172A",
            stipple="gray50",
            tags="overlay"
        )

        # Сообщение Game Over
        cx = self.width // 2
        cy = self.height // 2

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
            cx, cy + 18,
            text=f"Рекорд: {self.highscore}",
            font=("Segoe UI", 13),
            fill="#FBBF24",
            tags="overlay"
        )

        self.canvas.create_text(
            cx, cy + 44,
            text=f"Достигнут уровень: {self.level}",
            font=("Segoe UI", 14),
            fill="#CBD5E1",
            tags="overlay"
        )

        self.canvas.create_text(
            cx, cy + 72,
            text="Нажмите R для рестарта",
            font=("Segoe UI", 12),
            fill="#94A3B8",
            tags="overlay"
        )

    def _on_key_press(self, event):
        """Обработка нажатий клавиш"""
        key = _normalize_key(event)

        # Если ждём старта игры
        if hasattr(self, 'waiting_for_start') and self.waiting_for_start:
            self.waiting_for_start = False
            self.canvas.delete("welcome")
            self.start_game()
            return

        if self.game_over:
            if key == "r":
                self.canvas.delete("overlay")
                self.start_game()
            return

        # Пауза
        if key == "Return":
            if self.paused:
                self.paused = False
                self.canvas.delete("pause")
                self._game_loop()
            else:
                self.paused = True
                self._cancel_game_timer()
                self.canvas.create_text(
                    self.width // 2,
                    self.height // 2,
                    text="ПАУЗА",
                    font=("Segoe UI", 32, "bold"),
                    fill="#F8FAFC",
                    tags="pause"
                )

        # Движение платформы (резервное управление)
        elif key == "Left":
            self.paddle_x = max(0, self.paddle_x - 30)
        elif key == "Right":
            self.paddle_x = min(self.width - self.paddle_width, self.paddle_x + 30)

    def _on_mouse_move(self, event):
        """Движение платформы мышью"""
        if not self.game_over:
            self.paddle_x = max(0, min(self.width - self.paddle_width, event.x - self.paddle_width // 2))

    def _on_click(self, event):
        """Клик для старта"""
        if (
            not self.game_running
            and not self.game_over
            and self.life_restart_timer is None
        ):
            self.canvas.delete("level_msg")
            self.game_running = True
            self._game_loop()

    def run(self):
        """Запускает окно игры"""
        self._show_welcome()
        self.window.mainloop()

    def _show_welcome(self):
        """Показывает приветственный экран"""
        cx = self.width // 2
        cy = self.height // 2

        self.canvas.create_text(
            cx, cy - 100,
            text="📄 PDF ARKANOID",
            font=("Segoe UI", 32, "bold"),
            fill="#38BDF8",
            tags="welcome"
        )

        self.canvas.create_text(
            cx, cy - 40,
            text="Разбивай блоки форматов PDF!",
            font=("Segoe UI", 12),
            fill="#CBD5E1",
            tags="welcome"
        )

        self.canvas.create_text(
            cx, cy,
            text="🖱 Управление мышью",
            font=("Segoe UI", 11),
            fill="#94A3B8",
            tags="welcome"
        )

        self.canvas.create_text(
            cx, cy + 25,
            text="Enter — пауза",
            font=("Segoe UI", 11),
            fill="#94A3B8",
            tags="welcome"
        )

        self.canvas.create_text(
            cx, cy + 70,
            text="Нажмите любую кнопку или кликните",
            font=("Segoe UI", 14, "bold"),
            fill="#F8FAFC",
            tags="welcome"
        )

        # Флаг для отслеживания первого запуска
        self.waiting_for_start = True

        # Ждём клика для старта
        def start_on_click(event):
            if self.waiting_for_start:
                self.canvas.unbind("<Button-1>", click_bind_id)
                self.canvas.bind("<Button-1>", self._on_click)
                self.waiting_for_start = False
                self.canvas.delete("welcome")
                self.start_game()

        click_bind_id = self.canvas.bind("<Button-1>", start_on_click)


def launch_arkanoid_game(parent=None):
    """Запускает игру арканойд"""
    game = PDFArkanoidGame(parent)
    game.run()


if __name__ == "__main__":
    launch_arkanoid_game()
