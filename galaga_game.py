"""
PDF Galaga - классический космический шутер с тематикой PDF форматов
Пасхалка для PDF Pages Analyzer
"""

import tkinter as tk
from tkinter import messagebox
import random
import math


class PDFGalagaGame:
    def __init__(self, parent=None):
        self.parent = parent

        # Создаём окно
        self.window = tk.Toplevel(parent) if parent else tk.Tk()
        self.window.title("🚀 PDF Galaga")
        self.window.configure(bg="#0F172A")
        self.window.resizable(False, False)

        # Размеры игрового поля
        self.width = 600
        self.height = 700
        self.player_width = 40
        self.player_height = 40
        self.enemy_width = 35
        self.enemy_height = 30
        self.bullet_width = 4
        self.bullet_height = 12

        # Игровые переменные
        self.player_x = self.width // 2 - self.player_width // 2
        self.player_y = self.height - 80
        self.player_speed = 6
        self.score = 0
        self.lives = 3
        self.level = 1
        self.game_running = False
        self.game_over = False
        self.paused = False

        # Управление
        self.keys_pressed = set()

        # Игровые объекты
        self.bullets = []
        self.enemies = []
        self.enemy_bullets = []
        self.explosions = []

        # Параметры врагов
        self.enemy_formation_y = 80
        self.enemy_move_down = False
        self.enemy_direction = 1  # 1 = вправо, -1 = влево
        self.enemy_speed = 1
        self.enemy_shoot_chance = 0.002

        # Форматы PDF для врагов (типы и цвета)
        self.PDF_FORMATS = [
            {"name": "A0", "color": "#F87171", "points": 50, "hp": 3},   # Красный - босс
            {"name": "A1", "color": "#FB923C", "points": 40, "hp": 2},   # Оранжевый - средний
            {"name": "A2", "color": "#FBBF24", "points": 30, "hp": 2},   # Жёлтый - средний
            {"name": "A3", "color": "#34D399", "points": 20, "hp": 1},   # Зелёный - лёгкий
            {"name": "A4", "color": "#38BDF8", "points": 15, "hp": 1},   # Голубой - лёгкий
            {"name": "A5", "color": "#A78BFA", "points": 10, "hp": 1},   # Фиолетовый - лёгкий
        ]

        # Создаём интерфейс
        self._create_ui()

        # Центрируем окно
        self._center_window()

        # Привязываем события
        self.window.bind("<KeyPress>", self._on_key_press)
        self.window.bind("<KeyRelease>", self._on_key_release)
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

        # Жизни
        self.lives_label = tk.Label(
            stats_frame,
            text=f"🚀 × {self.lives}",
            font=("Segoe UI", 14, "bold"),
            fg="#38BDF8",
            bg="#1E293B"
        )
        self.lives_label.pack(side=tk.LEFT, padx=10)

        # Уровень
        self.level_label = tk.Label(
            stats_frame,
            text=f"Уровень: {self.level}",
            font=("Segoe UI", 14, "bold"),
            fg="#FBBF24",
            bg="#1E293B"
        )
        self.level_label.pack(side=tk.RIGHT, padx=10)

        # Canvas для игры
        self.canvas = tk.Canvas(
            self.window,
            width=self.width,
            height=self.height,
            bg="#0B0E14",
            highlightthickness=0
        )
        self.canvas.pack()

        # Рисуем звёзды на фоне
        self._draw_stars()

    def _draw_stars(self):
        """Рисует звёзды на фоне"""
        for _ in range(100):
            x = random.randint(0, self.width)
            y = random.randint(0, self.height)
            size = random.choice([1, 1, 1, 2])
            brightness = random.choice(["#4B5563", "#6B7280", "#9CA3AF", "#D1D5DB"])
            self.canvas.create_oval(
                x, y, x + size, y + size,
                fill=brightness,
                outline="",
                tags="stars"
            )

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
        self.window.destroy()

    def start_game(self):
        """Запускает новую игру"""
        self.score = 0
        self.lives = 3
        self.level = 1
        self.game_over = False
        self.paused = False
        self._start_level()

    def _start_level(self):
        """Запускает новый уровень"""
        self.game_running = False
        self.player_x = self.width // 2 - self.player_width // 2
        self.bullets = []
        self.enemy_bullets = []
        self.explosions = []

        # Параметры врагов увеличиваются с уровнем
        self.enemy_speed = 1 + (self.level - 1) * 0.3
        self.enemy_shoot_chance = 0.002 + (self.level - 1) * 0.0005
        self.enemy_direction = 1
        self.enemy_formation_y = 80

        # Создаём формацию врагов
        self._create_enemy_formation()

        # Обновляем статистику
        self._update_stats()

        # Показываем сообщение о старте уровня
        self._draw_game()
        cx = self.width // 2
        cy = self.height // 2

        self.canvas.create_text(
            cx, cy - 30,
            text=f"ВОЛНА {self.level}",
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

    def _create_enemy_formation(self):
        """Создаёт формацию врагов"""
        self.enemies = []
        rows = 4
        cols = 10
        spacing_x = 50
        spacing_y = 45
        start_x = (self.width - (cols - 1) * spacing_x) // 2

        for row in range(rows):
            # Разные типы врагов по рядам
            if row == 0:
                format_idx = 0  # A0 - боссы
            elif row == 1:
                format_idx = random.choice([1, 2])  # A1, A2 - средние
            else:
                format_idx = random.choice([3, 4, 5])  # A3, A4, A5 - лёгкие

            format_data = self.PDF_FORMATS[format_idx]

            for col in range(cols):
                x = start_x + col * spacing_x
                y = self.enemy_formation_y + row * spacing_y

                enemy = {
                    "x": x,
                    "y": y,
                    "base_x": x,
                    "base_y": y,
                    "row_offset": row * spacing_y,  # Смещение в формации
                    "width": self.enemy_width,
                    "height": self.enemy_height,
                    "color": format_data["color"],
                    "format": format_data["name"],
                    "points": format_data["points"],
                    "hp": format_data["hp"],
                    "max_hp": format_data["hp"],
                    "dive": False,
                    "dive_angle": 0,
                    "dive_speed": 0
                }
                self.enemies.append(enemy)

    def _game_loop(self):
        """Главный игровой цикл"""
        if not self.game_running or self.paused or self.game_over:
            return

        # Движение игрока
        if "Left" in self.keys_pressed or "a" in self.keys_pressed:
            self.player_x = max(0, self.player_x - self.player_speed)
        if "Right" in self.keys_pressed or "d" in self.keys_pressed:
            self.player_x = min(self.width - self.player_width, self.player_x + self.player_speed)

        # Обновление пуль игрока
        for bullet in self.bullets[:]:
            bullet["y"] -= bullet["speed"]
            if bullet["y"] < 0:
                self.bullets.remove(bullet)

        # Обновление пуль врагов
        for bullet in self.enemy_bullets[:]:
            bullet["y"] += bullet["speed"]
            if bullet["y"] > self.height:
                self.enemy_bullets.remove(bullet)
            # Проверка попадания в игрока
            elif self._check_player_hit(bullet):
                if bullet in self.enemy_bullets:  # Проверяем, что пуля всё ещё в списке
                    self.enemy_bullets.remove(bullet)
                self._lose_life()
                if self.game_over:
                    return  # Прерываем цикл только если игра закончилась

        # Обновление врагов
        self._update_enemies()

        # Проверка столкновений
        self._check_bullet_collisions()

        # Обновление взрывов
        for explosion in self.explosions[:]:
            explosion["frame"] += 1
            if explosion["frame"] > 10:
                self.explosions.remove(explosion)

        # Проверка победы на уровне
        if not self.enemies:
            self._next_level()
            return

        # Отрисовка
        self._draw_game()

        # Следующий кадр
        self.window.after(16, self._game_loop)  # ~60 FPS

    def _update_enemies(self):
        """Обновляет позиции и поведение врагов"""
        # Движение формации
        formation_enemies = [e for e in self.enemies if not e["dive"]]

        if formation_enemies:
            # Проверка границ для смены направления
            edge_hit = False
            for enemy in formation_enemies:
                if self.enemy_direction > 0 and enemy["x"] + self.enemy_width >= self.width - 10:
                    edge_hit = True
                elif self.enemy_direction < 0 and enemy["x"] <= 10:
                    edge_hit = True

            if edge_hit:
                self.enemy_direction *= -1
                self.enemy_formation_y += 20

            # Движение формации
            for enemy in formation_enemies:
                enemy["x"] += self.enemy_direction * self.enemy_speed
                enemy["y"] = self.enemy_formation_y + enemy["row_offset"]
                enemy["base_x"] = enemy["x"]
                enemy["base_y"] = enemy["y"]

                # Случайная стрельба
                if random.random() < self.enemy_shoot_chance:
                    self._enemy_shoot(enemy)

                # Случайная атака пикированием (реже на ранних уровнях)
                dive_chance = 0.0005 * self.level
                if random.random() < dive_chance:
                    enemy["dive"] = True
                    enemy["dive_angle"] = random.uniform(-30, 30)
                    enemy["dive_speed"] = 3 + self.level * 0.5

        # Обновление пикирующих врагов
        for enemy in self.enemies[:]:
            if enemy["dive"]:
                angle_rad = enemy["dive_angle"] * math.pi / 180
                enemy["x"] += enemy["dive_speed"] * math.sin(angle_rad)
                enemy["y"] += enemy["dive_speed"]

                # Стрельба во время пикирования
                if random.random() < 0.02:
                    self._enemy_shoot(enemy)

                # Проверка столкновения с игроком
                if self._check_enemy_player_collision(enemy):
                    self.enemies.remove(enemy)
                    self._create_explosion(enemy["x"], enemy["y"], enemy["color"])
                    self._lose_life()
                    if self.game_over:
                        return  # Прерываем цикл только если игра закончилась
                    continue

                # Возврат в формацию или вылет за экран
                if enemy["y"] > self.height:
                    # Возвращаем врага в исходную позицию формации
                    enemy["dive"] = False
                    enemy["x"] = enemy["base_x"]
                    enemy["y"] = enemy["base_y"]

    def _enemy_shoot(self, enemy):
        """Враг стреляет"""
        bullet = {
            "x": enemy["x"] + enemy["width"] // 2,
            "y": enemy["y"] + enemy["height"],
            "speed": 5 + self.level * 0.3
        }
        self.enemy_bullets.append(bullet)

    def _check_bullet_collisions(self):
        """Проверяет столкновения пуль с врагами"""
        for bullet in self.bullets[:]:
            for enemy in self.enemies[:]:
                if (enemy["x"] <= bullet["x"] <= enemy["x"] + enemy["width"] and
                    enemy["y"] <= bullet["y"] <= enemy["y"] + enemy["height"]):

                    # Попадание!
                    self.bullets.remove(bullet)
                    enemy["hp"] -= 1

                    if enemy["hp"] <= 0:
                        # Враг уничтожен
                        self.score += enemy["points"] * self.level
                        self.enemies.remove(enemy)
                        self._create_explosion(enemy["x"], enemy["y"], enemy["color"])
                    else:
                        # Враг повреждён, меняем цвет
                        enemy["color"] = self._darken_color(enemy["color"])

                    self._update_stats()
                    break

    def _check_player_hit(self, bullet):
        """Проверяет попадание пули врага в игрока"""
        return (self.player_x <= bullet["x"] <= self.player_x + self.player_width and
                self.player_y <= bullet["y"] <= self.player_y + self.player_height)

    def _check_enemy_player_collision(self, enemy):
        """Проверяет столкновение врага с игроком"""
        return (self.player_x < enemy["x"] + enemy["width"] and
                self.player_x + self.player_width > enemy["x"] and
                self.player_y < enemy["y"] + enemy["height"] and
                self.player_y + self.player_height > enemy["y"])

    def _create_explosion(self, x, y, color):
        """Создаёт эффект взрыва"""
        explosion = {
            "x": x,
            "y": y,
            "color": color,
            "frame": 0
        }
        self.explosions.append(explosion)

    def _darken_color(self, color):
        """Затемняет цвет для повреждённого врага"""
        r = int(color[1:3], 16)
        g = int(color[3:5], 16)
        b = int(color[5:7], 16)

        r = max(0, r - 50)
        g = max(0, g - 50)
        b = max(0, b - 50)

        return f"#{r:02x}{g:02x}{b:02x}"

    def _shoot(self):
        """Игрок стреляет"""
        if self.game_running and not self.paused:
            bullet = {
                "x": self.player_x + self.player_width // 2,
                "y": self.player_y,
                "speed": 10
            }
            self.bullets.append(bullet)

    def _lose_life(self):
        """Потеря жизни"""
        self.lives -= 1
        self._update_stats()

        if self.lives <= 0:
            self._game_over()
        else:
            # Короткая пауза и респавн
            self.player_x = self.width // 2 - self.player_width // 2
            self.enemy_bullets.clear()

    def _next_level(self):
        """Переход на следующий уровень"""
        self.level += 1
        self.lives = min(5, self.lives + 1)  # Бонусная жизнь
        self._start_level()

    def _draw_game(self):
        """Отрисовывает игровое поле"""
        self.canvas.delete("game")

        # Рисуем врагов
        for enemy in self.enemies:
            # Тело врага (треугольник/ромб)
            self.canvas.create_polygon(
                enemy["x"] + enemy["width"] // 2, enemy["y"],
                enemy["x"] + enemy["width"], enemy["y"] + enemy["height"] // 2,
                enemy["x"] + enemy["width"] // 2, enemy["y"] + enemy["height"],
                enemy["x"], enemy["y"] + enemy["height"] // 2,
                fill=enemy["color"],
                outline="#F8FAFC",
                width=1,
                tags="game"
            )

            # Текст формата
            self.canvas.create_text(
                enemy["x"] + enemy["width"] // 2,
                enemy["y"] + enemy["height"] // 2,
                text=enemy["format"],
                font=("Segoe UI", 7, "bold"),
                fill="#0B0E14",
                tags="game"
            )

            # Индикатор HP для врагов с несколькими HP
            if enemy["max_hp"] > 1:
                hp_width = 20
                hp_height = 3
                hp_x = enemy["x"] + (enemy["width"] - hp_width) // 2
                hp_y = enemy["y"] - 6

                # Фон HP
                self.canvas.create_rectangle(
                    hp_x, hp_y,
                    hp_x + hp_width, hp_y + hp_height,
                    fill="#1E293B",
                    outline="",
                    tags="game"
                )

                # Текущий HP
                hp_fill_width = (enemy["hp"] / enemy["max_hp"]) * hp_width
                self.canvas.create_rectangle(
                    hp_x, hp_y,
                    hp_x + hp_fill_width, hp_y + hp_height,
                    fill="#F87171",
                    outline="",
                    tags="game"
                )

        # Рисуем взрывы
        for explosion in self.explosions:
            size = 10 + explosion["frame"] * 3
            alpha_colors = ["#FBBF24", "#FB923C", "#F87171", "#7C2D12"]
            color_idx = min(explosion["frame"] // 3, len(alpha_colors) - 1)

            self.canvas.create_oval(
                explosion["x"] - size, explosion["y"] - size,
                explosion["x"] + size, explosion["y"] + size,
                fill=alpha_colors[color_idx],
                outline="",
                tags="game"
            )

        # Рисуем пули игрока
        for bullet in self.bullets:
            self.canvas.create_rectangle(
                bullet["x"] - self.bullet_width // 2, bullet["y"],
                bullet["x"] + self.bullet_width // 2, bullet["y"] + self.bullet_height,
                fill="#38BDF8",
                outline="",
                tags="game"
            )

        # Рисуем пули врагов
        for bullet in self.enemy_bullets:
            self.canvas.create_oval(
                bullet["x"] - 3, bullet["y"] - 3,
                bullet["x"] + 3, bullet["y"] + 3,
                fill="#F87171",
                outline="",
                tags="game"
            )

        # Рисуем игрока (корабль-треугольник)
        self.canvas.create_polygon(
            self.player_x + self.player_width // 2, self.player_y,
            self.player_x + self.player_width, self.player_y + self.player_height,
            self.player_x, self.player_y + self.player_height,
            fill="#38BDF8",
            outline="#F8FAFC",
            width=2,
            tags="game"
        )

        # Огоньки двигателя
        self.canvas.create_oval(
            self.player_x + 8, self.player_y + self.player_height - 5,
            self.player_x + 12, self.player_y + self.player_height + 5,
            fill="#FBBF24",
            outline="",
            tags="game"
        )
        self.canvas.create_oval(
            self.player_x + 28, self.player_y + self.player_height - 5,
            self.player_x + 32, self.player_y + self.player_height + 5,
            fill="#FBBF24",
            outline="",
            tags="game"
        )

    def _update_stats(self):
        """Обновляет статистику"""
        self.score_label.config(text=f"Счёт: {self.score}")
        self.lives_label.config(text=f"🚀 × {self.lives}")
        self.level_label.config(text=f"Уровень: {self.level}")

    def _game_over(self):
        """Конец игры"""
        self.game_over = True
        self.game_running = False

        # Затемнение поля
        self.canvas.create_rectangle(
            0, 0, self.width, self.height,
            fill="#0B0E14",
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
            cx, cy + 30,
            text=f"Волна: {self.level}",
            font=("Segoe UI", 14),
            fill="#CBD5E1",
            tags="overlay"
        )

        self.canvas.create_text(
            cx, cy + 70,
            text="Нажмите R для рестарта",
            font=("Segoe UI", 12),
            fill="#94A3B8",
            tags="overlay"
        )

    def _on_key_press(self, event):
        """Обработка нажатий клавиш"""
        key = event.keysym

        # Если ждём старта игры
        if hasattr(self, 'waiting_for_start') and self.waiting_for_start:
            if key in ["space", "Left", "Right", "a", "d", "A", "D", "Return"]:
                self.waiting_for_start = False
                self.canvas.delete("welcome")
                self.start_game()
            return

        if self.game_over:
            if key.lower() == "r":
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
                self.canvas.create_text(
                    self.width // 2,
                    self.height // 2,
                    text="ПАУЗА",
                    font=("Segoe UI", 32, "bold"),
                    fill="#F8FAFC",
                    tags="pause"
                )

        # Стрельба
        elif key == "space":
            self._shoot()

        # Движение
        self.keys_pressed.add(key)

    def _on_key_release(self, event):
        """Обработка отпускания клавиш"""
        key = event.keysym
        self.keys_pressed.discard(key)

    def _on_click(self, event):
        """Клик для старта"""
        if not self.game_running and not self.game_over:
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
            text="🚀 PDF GALAGA",
            font=("Segoe UI", 32, "bold"),
            fill="#38BDF8",
            tags="welcome"
        )

        self.canvas.create_text(
            cx, cy - 40,
            text="Защищай от вторжения PDF-форматов!",
            font=("Segoe UI", 12),
            fill="#CBD5E1",
            tags="welcome"
        )

        self.canvas.create_text(
            cx, cy,
            text="← → / A D — движение",
            font=("Segoe UI", 11),
            fill="#94A3B8",
            tags="welcome"
        )

        self.canvas.create_text(
            cx, cy + 25,
            text="Space — огонь",
            font=("Segoe UI", 11),
            fill="#94A3B8",
            tags="welcome"
        )

        self.canvas.create_text(
            cx, cy + 50,
            text="Enter — пауза",
            font=("Segoe UI", 11),
            fill="#94A3B8",
            tags="welcome"
        )

        self.canvas.create_text(
            cx, cy + 90,
            text="Нажмите любую игровую кнопку",
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


def launch_galaga_game(parent=None):
    """Запускает игру Galaga"""
    game = PDFGalagaGame(parent)
    game.run()


if __name__ == "__main__":
    launch_galaga_game()
