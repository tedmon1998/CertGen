import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageDraw, ImageFont, ImageTk
import pandas as pd
import os
from pathlib import Path
import threading
import matplotlib.font_manager as fm
from datetime import datetime
import json

class CertificateGenerator:
    def __init__(self, root):
        self.root = root
        self.root.title("Генератор сертификатов")
        self.root.geometry("1200x800")  # Размер окна по умолчанию: 1200px ширина, 800px высота
        
        # Переменные для хранения файлов
        self.template_path = None
        self.excel_path = None
        self.output_folder = os.getcwd()  # Текущая папка по умолчанию
        
        # Изображение и его размеры
        self.original_image = None
        self.display_image = None
        self.image_scale = 1.0
        
        # Координаты для размещения ФИО (старый способ - одна точка)
        self.text_x = tk.IntVar(value=400)
        self.text_y = tk.IntVar(value=300)
        
        # Область для размещения ФИО (новый способ - прямоугольник)
        self.text_area_x1 = tk.IntVar(value=300)
        self.text_area_y1 = tk.IntVar(value=250)
        self.text_area_x2 = tk.IntVar(value=500)
        self.text_area_y2 = tk.IntVar(value=350)
        
        # Выравнивание текста
        self.text_alignment = tk.StringVar(value="center")
        
        # Режим размещения (точка или область)
        self.text_mode = tk.StringVar(value="area")
        
        # Инструмент для работы с областью
        self.tool_mode = tk.StringVar(value="resize")  # "move" или "resize"
        
        # Отступы текста внутри области
        self.text_padding_left = tk.IntVar(value=10)
        self.text_padding_right = tk.IntVar(value=10)
        self.text_padding_top = tk.IntVar(value=10)
        self.text_padding_bottom = tk.IntVar(value=10)
        
        # Настройки окна
        self.window_width = tk.IntVar(value=1200)  # Ширина окна по умолчанию: 1200px
        self.window_height = tk.IntVar(value=800)
        self.left_panel_width = tk.IntVar(value=400)  # Ширина левой панели по умолчанию: 400px (33% от 1200px)
        
        # Настройки шрифта
        self.font_size = tk.IntVar(value=50)
        self.font_color = tk.StringVar(value="#000000")
        self.selected_font = tk.StringVar(value="Arial")
        self.line_spacing = tk.IntVar(value=5)  # Межстрочный интервал
        
        # Тестовый текст для предварительного просмотра
        self.preview_text = tk.StringVar(value="Иванов Иван Иванович")
        
        # Список доступных шрифтов
        self.available_fonts = []
        self.load_system_fonts()
        
        # Переменные для отслеживания изменений
        self.last_update_time = 0
        self.update_delay = 100  # миллисекунды
        
        # Переменные для перетаскивания области
        self.dragging = False
        self.drag_start_x = 0
        self.drag_start_y = 0
        self.drag_type = None  # 'move', 'resize_tl', 'resize_tr', 'resize_bl', 'resize_br'
        self.original_area = None  # Сохраняет исходные координаты при начале перетаскивания
        
        self.setup_ui()
        
    def load_system_fonts(self):
        """Загружает список доступных системных шрифтов"""
        try:
            # Получаем список шрифтов через matplotlib
            font_list = [f.name for f in fm.fontManager.ttflist]
            # Убираем дубликаты и сортируем
            self.available_fonts = sorted(list(set(font_list)))
            
            # Добавляем популярные шрифты, если их нет в списке
            popular_fonts = ['Arial', 'Times New Roman', 'Calibri', 'Verdana', 'Tahoma', 'Georgia']
            for font in popular_fonts:
                if font not in self.available_fonts:
                    self.available_fonts.insert(0, font)
                    
        except Exception as e:
            print(f"Ошибка при загрузке шрифтов: {e}")
            # Fallback к базовым шрифтам
            self.available_fonts = ['Arial', 'Times New Roman', 'Calibri', 'Verdana', 'Tahoma']
        
    def setup_ui(self):
        # Главный фрейм с разделением на левую и правую части
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Разделитель между панелями (PanedWindow)
        self.paned_window = ttk.PanedWindow(main_frame, orient=tk.HORIZONTAL)
        self.paned_window.pack(fill=tk.BOTH, expand=True)
        
        # ШИРИНА ПОЛЯ НАСТРОЕК: здесь задается начальная ширина левой панели (400px = 33% от 1200px)
        # Левая панель с настройками и прокруткой
        left_frame = ttk.Frame(self.paned_window, width=600)
        left_frame.pack_propagate(False)  # Запрещаем изменение размера содержимым
        self.left_panel = left_frame  # Сохраняем ссылку для обновления
        self.paned_window.add(left_frame, weight=1)  # Левая панель: меньший приоритет (33%)
        
        # Создаем Canvas и Scrollbar для прокрутки (вертикальная + горизонтальная)
        canvas = tk.Canvas(left_frame)
        v_scrollbar = ttk.Scrollbar(left_frame, orient="vertical", command=canvas.yview)
        h_scrollbar = ttk.Scrollbar(left_frame, orient="horizontal", command=canvas.xview)
        scrollable_frame = ttk.Frame(canvas)
        
        # ШИРИНА КОНТЕНТА: здесь задается ширина содержимого панели настроек (1000px)
        # Устанавливаем минима# ШИРИНА КОНТЕНТА: здесь задается ширина содержимого панели настроек (1000px)
        # льную ширину для горизонтальной прокрутки
        scrollable_frame.configure(width=1000)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True, padx=(0, 5))  # Добавляем отступ справа
        v_scrollbar.pack(side="right", fill="y")
        h_scrollbar.pack(side="bottom", fill="x")
        
        # Привязываем колесо мыши к прокрутке
        def _on_mousewheel(event):
            if event.state & 0x1:  # Shift + колесо = горизонтальная прокрутка
                canvas.xview_scroll(int(-1*(event.delta/120)), "units")
            else:  # Обычное колесо = вертикальная прокрутка
                canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        # Принудительно обновляем размер левой панели
        left_frame.update_idletasks()
        
        # Правая панель с предварительным просмотром
        # ШИРИНА ПРАВОЙ ПАНЕЛИ: здесь задается начальная ширина правой панели (800px = 67% от 1200px)
        right_frame = ttk.Frame(self.paned_window, width=800)
        right_frame.pack_propagate(False)  # Запрещаем изменение размера содержимым
        self.paned_window.add(right_frame, weight=2)  # Правая панель: больший приоритет (67%)
        
        # Заголовок
        title_label = ttk.Label(scrollable_frame, text="Генератор сертификатов", 
                              font=("Arial", 16, "bold"))
        title_label.pack(pady=(0, 10))
        
        # Инструкция по прокрутке
        scroll_info_label = ttk.Label(scrollable_frame, 
                                    text="💡 Shift + колесо мыши = горизонтальная прокрутка",
                                    font=("Arial", 9, "italic"), foreground="blue")
        scroll_info_label.pack(pady=(0, 10))
        
        # Секция загрузки файлов
        files_frame = ttk.LabelFrame(scrollable_frame, text="Файлы", padding="10")
        files_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Шаблон сертификата
        ttk.Label(files_frame, text="Шаблон сертификата:").pack(anchor=tk.W, pady=2)
        template_frame = ttk.Frame(files_frame)
        template_frame.pack(fill=tk.X, pady=2)
        self.template_label = ttk.Label(template_frame, text="Не выбран", foreground="gray")
        self.template_label.pack(side=tk.LEFT)
        ttk.Button(template_frame, text="Выбрать", 
                  command=self.select_template).pack(side=tk.RIGHT)
        
        # Excel/CSV файл с ФИО
        ttk.Label(files_frame, text="Excel/CSV файл с ФИО:").pack(anchor=tk.W, pady=2)
        excel_frame = ttk.Frame(files_frame)
        excel_frame.pack(fill=tk.X, pady=2)
        self.excel_label = ttk.Label(excel_frame, text="Не выбран", foreground="gray")
        self.excel_label.pack(side=tk.LEFT)
        ttk.Button(excel_frame, text="Выбрать", 
                  command=self.select_excel).pack(side=tk.RIGHT)
        
        # Папка для сохранения (автоматическая)
        ttk.Label(files_frame, text="Папка для сохранения:").pack(anchor=tk.W, pady=2)
        output_frame = ttk.Frame(files_frame)
        output_frame.pack(fill=tk.X, pady=2)
        self.output_label = ttk.Label(output_frame, text="Создается автоматически", foreground="blue")
        self.output_label.pack(side=tk.LEFT)
        ttk.Button(output_frame, text="Выбрать вручную", 
                  command=self.select_output_folder).pack(side=tk.RIGHT)
        
        # Секция настроек текста
        settings_frame = ttk.LabelFrame(scrollable_frame, text="Настройки текста", padding="10")
        settings_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Тестовый текст
        ttk.Label(settings_frame, text="Тестовый текст:").pack(anchor=tk.W, pady=2)
        preview_entry = ttk.Entry(settings_frame, textvariable=self.preview_text, width=30)
        preview_entry.pack(fill=tk.X, pady=2)
        
        # Режим размещения текста - вертикальное расположение
        mode_frame = ttk.Frame(settings_frame)
        mode_frame.pack(fill=tk.X, pady=5)
        
        # Первая строка
        row1_frame = ttk.Frame(mode_frame)
        row1_frame.pack(fill=tk.X, pady=2)
        
        ttk.Label(row1_frame, text="Режим:").pack(side=tk.LEFT)
        mode_combo = ttk.Combobox(row1_frame, textvariable=self.text_mode, 
                                 values=["area", "point"], state="readonly", width=8)
        mode_combo.pack(side=tk.LEFT, padx=(5, 15))
        
        ttk.Label(row1_frame, text="Инструмент:").pack(side=tk.LEFT)
        tool_combo = ttk.Combobox(row1_frame, textvariable=self.tool_mode, 
                                 values=["move", "resize"], state="readonly", width=8)
        tool_combo.pack(side=tk.LEFT, padx=(5, 0))
        
        # Вторая строка
        row2_frame = ttk.Frame(mode_frame)
        row2_frame.pack(fill=tk.X, pady=2)
        
        ttk.Label(row2_frame, text="Выравнивание:").pack(side=tk.LEFT)
        align_combo = ttk.Combobox(row2_frame, textvariable=self.text_alignment, 
                                  values=["left", "center", "right"], state="readonly", width=8)
        align_combo.pack(side=tk.LEFT, padx=(5, 0))
        
        # Область для текста (новый способ)
        area_frame = ttk.LabelFrame(settings_frame, text="Область для ФИО", padding="5")
        area_frame.pack(fill=tk.X, pady=5)
        
        # Точное позиционирование области
        pos_frame = ttk.Frame(area_frame)
        pos_frame.pack(fill=tk.X, pady=2)
        
        ttk.Label(pos_frame, text="Позиция области:").pack(anchor=tk.W)
        
        # Горизонтальное позиционирование
        h_frame = ttk.Frame(pos_frame)
        h_frame.pack(fill=tk.X, pady=2)
        
        ttk.Label(h_frame, text="X:").pack(side=tk.LEFT)
        x_scale = ttk.Scale(h_frame, from_=0, to=1000, orient=tk.HORIZONTAL, 
                           length=200, variable=self.text_area_x1, command=self.on_area_scale_change)
        x_scale.pack(side=tk.LEFT, padx=(5, 5), fill=tk.X, expand=True)
        ttk.Label(h_frame, textvariable=self.text_area_x1, width=4).pack(side=tk.RIGHT, padx=(0, 25))
        
        # Вертикальное позиционирование
        v_frame = ttk.Frame(pos_frame)
        v_frame.pack(fill=tk.X, pady=2)
        
        ttk.Label(v_frame, text="Y:").pack(side=tk.LEFT)
        y_scale = ttk.Scale(v_frame, from_=0, to=1000, orient=tk.HORIZONTAL, 
                           length=200, variable=self.text_area_y1, command=self.on_area_scale_change)
        y_scale.pack(side=tk.LEFT, padx=(5, 5), fill=tk.X, expand=True)
        ttk.Label(v_frame, textvariable=self.text_area_y1, width=4).pack(side=tk.RIGHT, padx=(0, 25))
        
        # Размер области
        size_frame = ttk.Frame(area_frame)
        size_frame.pack(fill=tk.X, pady=2)
        
        ttk.Label(size_frame, text="Размер области:").pack(anchor=tk.W)
        
        # Ширина области
        w_frame = ttk.Frame(size_frame)
        w_frame.pack(fill=tk.X, pady=2)
        
        ttk.Label(w_frame, text="Ширина:").pack(side=tk.LEFT)
        width_scale = ttk.Scale(w_frame, from_=50, to=500, orient=tk.HORIZONTAL, 
                               length=200, command=self.on_width_scale_change)
        width_scale.pack(side=tk.LEFT, padx=(5, 5), fill=tk.X, expand=True)
        width_label = ttk.Label(w_frame, text="200", width=4)
        width_label.pack(side=tk.RIGHT, padx=(0, 25))
        
        # Высота области
        h_size_frame = ttk.Frame(size_frame)
        h_size_frame.pack(fill=tk.X, pady=2)
        
        ttk.Label(h_size_frame, text="Высота:").pack(side=tk.LEFT)
        height_scale = ttk.Scale(h_size_frame, from_=20, to=200, orient=tk.HORIZONTAL, 
                                length=200, command=self.on_height_scale_change)
        height_scale.pack(side=tk.LEFT, padx=(5, 5), fill=tk.X, expand=True)
        height_label = ttk.Label(h_size_frame, text="100", width=4)
        height_label.pack(side=tk.RIGHT, padx=(0, 25))
        
        # Отступы текста внутри области
        padding_frame = ttk.LabelFrame(area_frame, text="Отступы текста", padding="5")
        padding_frame.pack(fill=tk.X, pady=(10, 0))
        
        # Отступы по горизонтали
        h_padding_frame = ttk.Frame(padding_frame)
        h_padding_frame.pack(fill=tk.X, pady=2)
        
        ttk.Label(h_padding_frame, text="Слева:").pack(side=tk.LEFT)
        left_padding_scale = ttk.Scale(h_padding_frame, from_=0, to=50, orient=tk.HORIZONTAL, 
                                     length=120, variable=self.text_padding_left, command=self.schedule_update)
        left_padding_scale.pack(side=tk.LEFT, padx=(5, 10), fill=tk.X, expand=True)
        ttk.Label(h_padding_frame, textvariable=self.text_padding_left, width=3).pack(side=tk.RIGHT, padx=(0, 25))
        
        # Отступы по вертикали
        v_padding_frame = ttk.Frame(padding_frame)
        v_padding_frame.pack(fill=tk.X, pady=2)
        
        ttk.Label(v_padding_frame, text="Справа:").pack(side=tk.LEFT)
        right_padding_scale = ttk.Scale(v_padding_frame, from_=0, to=50, orient=tk.HORIZONTAL, 
                                      length=120, variable=self.text_padding_right, command=self.schedule_update)
        right_padding_scale.pack(side=tk.LEFT, padx=(5, 10), fill=tk.X, expand=True)
        ttk.Label(v_padding_frame, textvariable=self.text_padding_right, width=3).pack(side=tk.RIGHT, padx=(0, 25))
        
        # Отступы по вертикали
        v_padding_frame2 = ttk.Frame(padding_frame)
        v_padding_frame2.pack(fill=tk.X, pady=2)
        
        ttk.Label(v_padding_frame2, text="Сверху:").pack(side=tk.LEFT)
        top_padding_scale = ttk.Scale(v_padding_frame2, from_=0, to=50, orient=tk.HORIZONTAL, 
                                    length=120, variable=self.text_padding_top, command=self.schedule_update)
        top_padding_scale.pack(side=tk.LEFT, padx=(5, 10), fill=tk.X, expand=True)
        ttk.Label(v_padding_frame2, textvariable=self.text_padding_top, width=3).pack(side=tk.RIGHT, padx=(0, 25))
        
        # Отступы по вертикали
        v_padding_frame3 = ttk.Frame(padding_frame)
        v_padding_frame3.pack(fill=tk.X, pady=2)
        
        ttk.Label(v_padding_frame3, text="Снизу:").pack(side=tk.LEFT)
        bottom_padding_scale = ttk.Scale(v_padding_frame3, from_=0, to=50, orient=tk.HORIZONTAL, 
                                       length=120, variable=self.text_padding_bottom, command=self.schedule_update)
        bottom_padding_scale.pack(side=tk.LEFT, padx=(5, 10), fill=tk.X, expand=True)
        ttk.Label(v_padding_frame3, textvariable=self.text_padding_bottom, width=3).pack(side=tk.RIGHT, padx=(0, 25))
        
        # Старый способ (точка) - скрыт по умолчанию
        point_frame = ttk.LabelFrame(settings_frame, text="Точка для ФИО", padding="5")
        point_frame.pack(fill=tk.X, pady=5)
        
        coords_frame = ttk.Frame(point_frame)
        coords_frame.pack(fill=tk.X, pady=2)
        
        ttk.Label(coords_frame, text="X:").pack(side=tk.LEFT)
        x_spinbox = ttk.Spinbox(coords_frame, from_=0, to=2000, width=8, 
                               textvariable=self.text_x, command=self.update_preview)
        x_spinbox.pack(side=tk.LEFT, padx=(5, 10))
        
        ttk.Label(coords_frame, text="Y:").pack(side=tk.LEFT)
        y_spinbox = ttk.Spinbox(coords_frame, from_=0, to=2000, width=8, 
                               textvariable=self.text_y, command=self.update_preview)
        y_spinbox.pack(side=tk.LEFT, padx=(5, 0))
        
        # Шрифт
        font_name_frame = ttk.Frame(settings_frame)
        font_name_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(font_name_frame, text="Шрифт:").pack(side=tk.LEFT)
        font_combo = ttk.Combobox(font_name_frame, textvariable=self.selected_font, 
                                 values=self.available_fonts, width=20, state="readonly")
        font_combo.pack(side=tk.LEFT, padx=(5, 10))
        
        ttk.Button(font_name_frame, text="Загрузить файл", 
                  command=self.load_custom_font).pack(side=tk.RIGHT)
        
        # Размер шрифта
        font_frame = ttk.Frame(settings_frame)
        font_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(font_frame, text="Размер:").pack(side=tk.LEFT)
        font_scale = ttk.Scale(font_frame, from_=10, to=200, orient=tk.HORIZONTAL, 
                              length=200, variable=self.font_size, command=self.schedule_update)
        font_scale.pack(side=tk.LEFT, padx=(5, 10), fill=tk.X, expand=True)
        ttk.Label(font_frame, textvariable=self.font_size, width=4).pack(side=tk.RIGHT, padx=(0, 25))
        
        # Цвет шрифта
        color_frame = ttk.Frame(settings_frame)
        color_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(color_frame, text="Цвет:").pack(side=tk.LEFT)
        color_entry = ttk.Entry(color_frame, textvariable=self.font_color, width=15)
        color_entry.pack(side=tk.LEFT, padx=(5, 0))
        
        # Межстрочный интервал
        spacing_frame = ttk.Frame(settings_frame)
        spacing_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(spacing_frame, text="Межстрочный интервал:").pack(side=tk.LEFT)
        spacing_scale = ttk.Scale(spacing_frame, from_=0, to=50, orient=tk.HORIZONTAL, 
                                 length=200, variable=self.line_spacing, command=self.schedule_update)
        spacing_scale.pack(side=tk.LEFT, padx=(5, 0), fill=tk.X, expand=True)
        ttk.Label(spacing_frame, textvariable=self.line_spacing, width=4).pack(side=tk.RIGHT, padx=(0, 25))
        
        # Привязка событий для автоматического обновления
        self.text_x.trace('w', self.schedule_update)
        self.text_y.trace('w', self.schedule_update)
        self.text_area_x1.trace('w', self.schedule_update)
        self.text_area_y1.trace('w', self.schedule_update)
        self.text_area_x2.trace('w', self.schedule_update)
        self.text_area_y2.trace('w', self.schedule_update)
        self.text_alignment.trace('w', self.schedule_update)
        self.text_mode.trace('w', self.schedule_update)
        self.tool_mode.trace('w', self.schedule_update)
        self.font_size.trace('w', self.schedule_update)
        self.font_color.trace('w', self.schedule_update)
        self.line_spacing.trace('w', self.schedule_update)
        self.text_padding_left.trace('w', self.schedule_update)
        self.text_padding_right.trace('w', self.schedule_update)
        self.text_padding_top.trace('w', self.schedule_update)
        self.text_padding_bottom.trace('w', self.schedule_update)
        self.preview_text.trace('w', self.schedule_update)
        self.selected_font.trace('w', self.schedule_update)
        
        # Секция настроек окна
        window_frame = ttk.LabelFrame(scrollable_frame, text="Настройки окна", padding="10")
        window_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Размеры окна
        size_frame = ttk.Frame(window_frame)
        size_frame.pack(fill=tk.X, pady=2)
        
        ttk.Label(size_frame, text="Ширина окна:").pack(side=tk.LEFT)
        window_width_scale = ttk.Scale(size_frame, from_=800, to=2000, orient=tk.HORIZONTAL, 
                                     length=200, variable=self.window_width, command=self.on_window_size_change)
        window_width_scale.pack(side=tk.LEFT, padx=(5, 10), fill=tk.X, expand=True)
        ttk.Label(size_frame, textvariable=self.window_width, width=4).pack(side=tk.RIGHT, padx=(0, 25))
        
        height_frame = ttk.Frame(window_frame)
        height_frame.pack(fill=tk.X, pady=2)
        
        ttk.Label(height_frame, text="Высота окна:").pack(side=tk.LEFT)
        window_height_scale = ttk.Scale(height_frame, from_=600, to=1200, orient=tk.HORIZONTAL, 
                                      length=200, variable=self.window_height, command=self.on_window_size_change)
        window_height_scale.pack(side=tk.LEFT, padx=(5, 10), fill=tk.X, expand=True)
        ttk.Label(height_frame, textvariable=self.window_height, width=4).pack(side=tk.RIGHT, padx=(0, 25))
        
        # Ширина панели настроек
        panel_frame = ttk.Frame(window_frame)
        panel_frame.pack(fill=tk.X, pady=2)
        
        ttk.Label(panel_frame, text="Ширина панели:").pack(side=tk.LEFT)
        panel_width_scale = ttk.Scale(panel_frame, from_=500, to=1200, orient=tk.HORIZONTAL, 
                                    length=200, variable=self.left_panel_width, command=self.on_panel_width_change)
        panel_width_scale.pack(side=tk.LEFT, padx=(5, 10), fill=tk.X, expand=True)
        ttk.Label(panel_frame, textvariable=self.left_panel_width, width=4).pack(side=tk.RIGHT, padx=(0, 25))
        
        # Секция сохранения/загрузки настроек
        settings_buttons_frame = ttk.LabelFrame(scrollable_frame, text="Настройки проекта", padding="10")
        settings_buttons_frame.pack(fill=tk.X, pady=(0, 10))
        
        buttons_frame = ttk.Frame(settings_buttons_frame)
        buttons_frame.pack(fill=tk.X)
        
        ttk.Button(buttons_frame, text="Сохранить настройки", 
                  command=self.save_settings).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(buttons_frame, text="Загрузить настройки", 
                  command=self.load_settings).pack(side=tk.LEFT)
        
        # Секция генерации
        generate_frame = ttk.Frame(scrollable_frame)
        generate_frame.pack(fill=tk.X, pady=10)
        
        self.generate_button = ttk.Button(generate_frame, text="Генерировать сертификаты", 
                                         command=self.generate_certificates)
        self.generate_button.pack(fill=tk.X)
        
        # Прогресс бар
        self.progress = ttk.Progressbar(scrollable_frame, mode='determinate')
        self.progress.pack(fill=tk.X, pady=5)
        
        # Статус
        self.status_label = ttk.Label(scrollable_frame, text="Готов к работе")
        self.status_label.pack(pady=5)
        
        # Создаем правую панель с предварительным просмотром
        preview_frame = ttk.LabelFrame(right_frame, text="Предварительный просмотр", padding="10")
        preview_frame.pack(fill=tk.BOTH, expand=True)
        
        # Canvas для отображения изображения с границей
        self.canvas = tk.Canvas(preview_frame, bg="white", cursor="crosshair", 
                               relief=tk.SUNKEN, bd=2)
        self.canvas.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Привязка событий мыши
        self.canvas.bind("<Button-1>", self.on_canvas_click)
        self.canvas.bind("<B1-Motion>", self.on_canvas_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_canvas_release)
        self.canvas.bind("<Motion>", self.on_canvas_motion)
        
        # Инструкция
        instruction_label = ttk.Label(preview_frame, 
                                    text="Выберите инструмент: move (перемещение) или resize (изменение размера)",
                                    font=("Arial", 10, "italic"))
        instruction_label.pack(pady=5)
        
        # Принудительно обновляем размеры всех элементов
        self.root.update_idletasks()
        self.root.geometry(f"{self.window_width.get()}x{self.window_height.get()}")
        
        # Привязываем событие изменения размера окна
        self.root.bind('<Configure>', self.on_window_configure)
        
        # Принудительно обновляем размеры всех панелей
        self.root.after(100, self.force_update_layout)
        
    def select_template(self):
        file_path = filedialog.askopenfilename(
            title="Выберите шаблон сертификата",
            filetypes=[("Изображения", "*.png *.jpg *.jpeg *.bmp *.gif"), ("Все файлы", "*.*")]
        )
        if file_path:
            self.template_path = file_path
            self.template_label.config(text=os.path.basename(file_path), foreground="black")
            self.load_template_image()
            
    def load_template_image(self):
        """Загружает изображение шаблона и отображает его в canvas"""
        try:
            self.original_image = Image.open(self.template_path)
            self.display_image_in_canvas()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить изображение: {str(e)}")
            
    def display_image_in_canvas(self):
        """Отображает изображение в canvas с масштабированием"""
        if not self.original_image:
            return
        
        # Принудительно обновляем размеры canvas
        self.canvas.update_idletasks()
            
        # Получаем размеры canvas
        canvas_width = self.canvas.winfo_width()
        canvas_height = self.canvas.winfo_height()
        
        if canvas_width <= 1 or canvas_height <= 1:
            # Canvas еще не инициализирован, повторим попытку через 100мс
            self.root.after(100, self.display_image_in_canvas)
            return
            
        # Вычисляем масштаб для подгонки изображения под canvas
        img_width, img_height = self.original_image.size
        scale_x = canvas_width / img_width
        scale_y = canvas_height / img_height
        self.image_scale = min(scale_x, scale_y, 1.0)  # Не увеличиваем изображение
        
        # Масштабируем изображение
        new_width = int(img_width * self.image_scale)
        new_height = int(img_height * self.image_scale)
        resized_image = self.original_image.resize((new_width, new_height), Image.Resampling.LANCZOS)
        
        # Конвертируем в PhotoImage для tkinter
        self.display_image = ImageTk.PhotoImage(resized_image)
        
        # Очищаем canvas и отображаем изображение
        self.canvas.delete("all")
        self.canvas.create_image(canvas_width//2, canvas_height//2, image=self.display_image, anchor=tk.CENTER)
        
        # Обновляем предварительный просмотр
        self.update_preview()
        
    def on_canvas_click(self, event):
        """Обработчик клика по canvas для установки координат"""
        if not self.original_image:
            return
            
        # Конвертируем координаты canvas в координаты изображения
        original_x, original_y = self.canvas_to_image_coords(event.x, event.y)
        
        # Проверяем, что клик был по изображению
        img_width, img_height = self.original_image.size
        if 0 <= original_x < img_width and 0 <= original_y < img_height:
            if self.text_mode.get() == "point":
                # Старый способ - одна точка
                self.text_x.set(original_x)
                self.text_y.set(original_y)
            else:
                # Проверяем, не кликнули ли мы по существующей области
                drag_type = self.get_drag_type(original_x, original_y)
                if drag_type:
                    # Начинаем перетаскивание
                    self.dragging = True
                    self.drag_type = drag_type
                    self.drag_start_x = original_x
                    self.drag_start_y = original_y
                    self.original_area = (
                        self.text_area_x1.get(),
                        self.text_area_y1.get(),
                        self.text_area_x2.get(),
                        self.text_area_y2.get()
                    )
                else:
                    # Создаем новую область
                    self.text_area_x1.set(original_x)
                    self.text_area_y1.set(original_y)
                    self.text_area_x2.set(original_x + 200)
                    self.text_area_y2.set(original_y + 100)
            
    def on_canvas_motion(self, event):
        """Обработчик движения мыши для отображения координат"""
        if not self.original_image:
            return
            
        # Получаем размеры canvas и изображения
        canvas_width = self.canvas.winfo_width()
        canvas_height = self.canvas.winfo_height()
        img_width, img_height = self.original_image.size
        
        # Вычисляем позицию мыши относительно изображения
        mouse_x = event.x - (canvas_width - img_width * self.image_scale) // 2
        mouse_y = event.y - (canvas_height - img_height * self.image_scale) // 2
        
        # Конвертируем в координаты оригинального изображения
        original_x = int(mouse_x / self.image_scale)
        original_y = int(mouse_y / self.image_scale)
        
        # Обновляем статус с координатами
        if 0 <= original_x < img_width and 0 <= original_y < img_height:
            self.status_label.config(text=f"Координаты: {original_x}, {original_y}")
        else:
            self.status_label.config(text="Готов к работе")
    
    def on_canvas_drag(self, event):
        """Обработчик перетаскивания мыши"""
        if not self.dragging or not self.original_image:
            return
            
        # Конвертируем координаты canvas в координаты изображения
        current_x, current_y = self.canvas_to_image_coords(event.x, event.y)
        
        # Вычисляем смещение
        dx = current_x - self.drag_start_x
        dy = current_y - self.drag_start_y
        
        # Получаем исходные координаты
        x1, y1, x2, y2 = self.original_area
        
        if self.drag_type == 'move':
            # Перемещаем всю область
            self.text_area_x1.set(x1 + dx)
            self.text_area_y1.set(y1 + dy)
            self.text_area_x2.set(x2 + dx)
            self.text_area_y2.set(y2 + dy)
            
        elif self.drag_type == 'resize_tl':
            # Изменяем размер от верхнего левого угла
            self.text_area_x1.set(x1 + dx)
            self.text_area_y1.set(y1 + dy)
            
        elif self.drag_type == 'resize_tr':
            # Изменяем размер от верхнего правого угла
            self.text_area_x2.set(x2 + dx)
            self.text_area_y1.set(y1 + dy)
            
        elif self.drag_type == 'resize_bl':
            # Изменяем размер от нижнего левого угла
            self.text_area_x1.set(x1 + dx)
            self.text_area_y2.set(y2 + dy)
            
        elif self.drag_type == 'resize_br':
            # Изменяем размер от нижнего правого угла
            self.text_area_x2.set(x2 + dx)
            self.text_area_y2.set(y2 + dy)
        
        # Мгновенное обновление предварительного просмотра без задержки
        self.update_preview()
    
    def on_canvas_release(self, event):
        """Обработчик отпускания мыши"""
        self.dragging = False
        self.drag_type = None
        self.original_area = None
            
    def schedule_update(self, *args):
        """Планирует обновление предварительного просмотра с задержкой"""
        current_time = self.root.after_idle(lambda: None)
        if hasattr(self, '_update_job'):
            self.root.after_cancel(self._update_job)
        # Уменьшаем задержку для более отзывчивого интерфейса
        self._update_job = self.root.after(50, self.update_preview)
        
    def update_preview(self):
        """Обновляет предварительный просмотр текста на изображении"""
        if not self.original_image or not self.display_image:
            return
            
        try:
            # Создаем копию изображения для предварительного просмотра
            preview_img = self.original_image.copy()
            draw = ImageDraw.Draw(preview_img)
            
            # Получаем шрифт
            font = self.get_font(self.font_size.get())
            
            # Добавляем границу вокруг сертификата
            img_width, img_height = preview_img.size
            border_width = 3
            draw.rectangle([0, 0, img_width-1, img_height-1], outline="#CCCCCC", width=border_width)
            
            # Добавляем текст
            text = self.preview_text.get()
            if text:
                if self.text_mode.get() == "point":
                    # Старый способ - одна точка
                    x, y = self.calculate_text_position(text, font, self.text_alignment.get())
                    print(f"Добавляем текст: '{text}' в позицию ({x}, {y}) с выравниванием {self.text_alignment.get()}")
                    draw.text((x, y), text, fill=self.font_color.get(), font=font)
                else:
                    # Новый способ - область с переносом строк
                    x, y, max_width = self.calculate_text_position(text, font, self.text_alignment.get())
                    print(f"Добавляем многострочный текст: '{text}' в позицию ({x}, {y}) с выравниванием {self.text_alignment.get()}, ширина области: {max_width}")
                    self.draw_multiline_text(draw, text, font, x, y, self.text_alignment.get(), 
                                           max_width, self.line_spacing.get())
                
                # Если режим "область", рисуем рамку области и маркеры
                if self.text_mode.get() == "area":
                    x1, y1 = self.text_area_x1.get(), self.text_area_y1.get()
                    x2, y2 = self.text_area_x2.get(), self.text_area_y2.get()
                    
                    # Цвет рамки зависит от инструмента
                    if self.tool_mode.get() == "move":
                        outline_color = "#00AA00"  # Зеленый для перемещения
                    else:
                        outline_color = "#FF0000"  # Красный для изменения размера
                    
                    # Рисуем внешнюю рамку области
                    draw.rectangle([x1, y1, x2, y2], outline=outline_color, width=2)
                    
                    # Рисуем внутреннюю рамку отступов (синяя)
                    text_x1 = x1 + self.text_padding_left.get()
                    text_y1 = y1 + self.text_padding_top.get()
                    text_x2 = x2 - self.text_padding_right.get()
                    text_y2 = y2 - self.text_padding_bottom.get()
                    
                    if text_x1 < text_x2 and text_y1 < text_y2:  # Проверяем что отступы не превышают размер области
                        draw.rectangle([text_x1, text_y1, text_x2, text_y2], outline="#0066CC", width=1)
                    
                    # Рисуем маркеры только для инструмента "resize"
                    if self.tool_mode.get() == "resize":
                        handle_size = 4
                        handle_color = "#FF0000"
                        # Верхний левый угол
                        draw.rectangle([x1-handle_size, y1-handle_size, x1+handle_size, y1+handle_size], 
                                     fill=handle_color, outline="#FFFFFF", width=1)
                        # Верхний правый угол
                        draw.rectangle([x2-handle_size, y1-handle_size, x2+handle_size, y1+handle_size], 
                                     fill=handle_color, outline="#FFFFFF", width=1)
                        # Нижний левый угол
                        draw.rectangle([x1-handle_size, y2-handle_size, x1+handle_size, y2+handle_size], 
                                     fill=handle_color, outline="#FFFFFF", width=1)
                        # Нижний правый угол
                        draw.rectangle([x2-handle_size, y2-handle_size, x2+handle_size, y2+handle_size], 
                                     fill=handle_color, outline="#FFFFFF", width=1)
            
            # Масштабируем для отображения
            img_width, img_height = preview_img.size
            new_width = int(img_width * self.image_scale)
            new_height = int(img_height * self.image_scale)
            resized_preview = preview_img.resize((new_width, new_height), Image.Resampling.LANCZOS)
            
            # Конвертируем в PhotoImage
            preview_photo = ImageTk.PhotoImage(resized_preview)
            
            # Обновляем canvas
            self.canvas.delete("all")
            canvas_width = self.canvas.winfo_width()
            canvas_height = self.canvas.winfo_height()
            self.canvas.create_image(canvas_width//2, canvas_height//2, image=preview_photo, anchor=tk.CENTER)
            
            # Сохраняем ссылку на изображение, чтобы оно не было удалено сборщиком мусора
            self.display_image = preview_photo
            
        except Exception as e:
            print(f"Ошибка при обновлении предварительного просмотра: {e}")
            
    def select_excel(self):
        file_path = filedialog.askopenfilename(
            title="Выберите Excel или CSV файл с ФИО",
            filetypes=[("Excel файлы", "*.xlsx *.xls"), ("CSV файлы", "*.csv"), ("Все файлы", "*.*")]
        )
        if file_path:
            self.excel_path = file_path
            self.excel_label.config(text=os.path.basename(file_path), foreground="black")
            
    def select_output_folder(self):
        folder_path = filedialog.askdirectory(title="Выберите папку для сохранения сертификатов")
        if folder_path:
            self.output_folder = folder_path
            self.output_label.config(text=os.path.basename(folder_path), foreground="black")
            
    def load_custom_font(self):
        """Загружает пользовательский шрифт из файла"""
        file_path = filedialog.askopenfilename(
            title="Выберите файл шрифта",
            filetypes=[("Шрифты", "*.ttf *.otf *.ttc"), ("Все файлы", "*.*")]
        )
        if file_path:
            try:
                # Пробуем загрузить шрифт
                font = ImageFont.truetype(file_path, 20)
                font_name = os.path.splitext(os.path.basename(file_path))[0]
                
                # Добавляем в список доступных шрифтов
                if font_name not in self.available_fonts:
                    self.available_fonts.append(font_name)
                    # Обновляем combobox
                    font_combo = None
                    for child in self.root.winfo_children():
                        font_combo = self.find_font_combo(child)
                        if font_combo:
                            break
                    if font_combo:
                        font_combo['values'] = self.available_fonts
                
                self.selected_font.set(font_name)
                messagebox.showinfo("Успех", f"Шрифт '{font_name}' успешно загружен!")
                
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось загрузить шрифт: {str(e)}")
                
    def find_font_combo(self, widget):
        """Рекурсивно ищет combobox для шрифтов"""
        if isinstance(widget, ttk.Combobox) and widget.cget('values'):
            return widget
        for child in widget.winfo_children():
            result = self.find_font_combo(child)
            if result:
                return result
        return None
            
    def get_font(self, size):
        """Получает шрифт с указанным размером"""
        font_name = self.selected_font.get()
        
        # Список возможных путей к шрифтам
        font_paths = [
            # Популярные шрифты Windows
            f"C:/Windows/Fonts/{font_name}.ttf",
            f"C:/Windows/Fonts/{font_name}.otf",
            f"C:/Windows/Fonts/{font_name}.ttc",
            # Популярные шрифты с разными регистрами
            f"C:/Windows/Fonts/{font_name.lower()}.ttf",
            f"C:/Windows/Fonts/{font_name.upper()}.ttf",
            f"C:/Windows/Fonts/{font_name.title()}.ttf",
            # Альтернативные названия
            f"C:/Windows/Fonts/arial.ttf",
            f"C:/Windows/Fonts/Arial.ttf",
            f"C:/Windows/Fonts/calibri.ttf",
            f"C:/Windows/Fonts/Calibri.ttf",
            f"C:/Windows/Fonts/times.ttf",
            f"C:/Windows/Fonts/Times.ttf",
        ]
        
        # Пробуем загрузить шрифт по разным путям
        for font_path in font_paths:
            try:
                if os.path.exists(font_path):
                    print(f"Загружаем шрифт: {font_path}")
                    return ImageFont.truetype(font_path, size)
            except Exception as e:
                print(f"Ошибка загрузки шрифта {font_path}: {e}")
                continue
        
        # Если не нашли файл шрифта, пробуем загрузить по имени
        try:
            print(f"Пробуем загрузить шрифт по имени: {font_name}")
            return ImageFont.truetype(font_name, size)
        except Exception as e:
            print(f"Ошибка загрузки шрифта по имени {font_name}: {e}")
            pass
        
        # Fallback к системному шрифту по умолчанию
        try:
            print("Используем fallback шрифт: arial")
            return ImageFont.truetype("arial", size)
        except Exception as e:
            print(f"Ошибка загрузки fallback шрифта: {e}")
            print("Используем шрифт по умолчанию")
            return ImageFont.load_default()
    
    def save_settings(self):
        """Сохраняет все настройки в JSON файл"""
        try:
            settings = {
                "template_path": self.template_path,
                "excel_path": self.excel_path,
                "output_folder": self.output_folder,
                "text_mode": self.text_mode.get(),
                "tool_mode": self.tool_mode.get(),
                "text_alignment": self.text_alignment.get(),
                "text_x": self.text_x.get(),
                "text_y": self.text_y.get(),
                "text_area_x1": self.text_area_x1.get(),
                "text_area_y1": self.text_area_y1.get(),
                "text_area_x2": self.text_area_x2.get(),
                "text_area_y2": self.text_area_y2.get(),
            "text_padding_left": self.text_padding_left.get(),
            "text_padding_right": self.text_padding_right.get(),
            "text_padding_top": self.text_padding_top.get(),
            "text_padding_bottom": self.text_padding_bottom.get(),
            "window_width": self.window_width.get(),
            "window_height": self.window_height.get(),
            "left_panel_width": self.left_panel_width.get(),
            "font_size": self.font_size.get(),
                "font_color": self.font_color.get(),
                "selected_font": self.selected_font.get(),
                "line_spacing": self.line_spacing.get(),
                "preview_text": self.preview_text.get(),
                "available_fonts": self.available_fonts
            }
            
            file_path = filedialog.asksaveasfilename(
                title="Сохранить настройки проекта",
                defaultextension=".json",
                filetypes=[("JSON файлы", "*.json"), ("Все файлы", "*.*")]
            )
            
            if file_path:
                with open(file_path, 'w', encoding='utf-8') as f:
                    json.dump(settings, f, ensure_ascii=False, indent=2)
                messagebox.showinfo("Успех", f"Настройки сохранены в файл:\n{file_path}")
                
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить настройки: {str(e)}")
    
    def load_settings(self):
        """Загружает настройки из JSON файла"""
        try:
            file_path = filedialog.askopenfilename(
                title="Загрузить настройки проекта",
                filetypes=[("JSON файлы", "*.json"), ("Все файлы", "*.*")]
            )
            
            if file_path:
                with open(file_path, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
                
                # Загружаем настройки
                self.template_path = settings.get("template_path")
                self.excel_path = settings.get("excel_path")
                self.output_folder = settings.get("output_folder", os.getcwd())
                
                # Обновляем переменные интерфейса
                self.text_mode.set(settings.get("text_mode", "area"))
                self.tool_mode.set(settings.get("tool_mode", "resize"))
                self.text_alignment.set(settings.get("text_alignment", "center"))
                self.text_x.set(settings.get("text_x", 400))
                self.text_y.set(settings.get("text_y", 300))
                self.text_area_x1.set(settings.get("text_area_x1", 300))
                self.text_area_y1.set(settings.get("text_area_y1", 250))
                self.text_area_x2.set(settings.get("text_area_x2", 500))
                self.text_area_y2.set(settings.get("text_area_y2", 350))
            self.text_padding_left.set(settings.get("text_padding_left", 10))
            self.text_padding_right.set(settings.get("text_padding_right", 10))
            self.text_padding_top.set(settings.get("text_padding_top", 10))
            self.text_padding_bottom.set(settings.get("text_padding_bottom", 10))
            self.window_width.set(settings.get("window_width", 1800))
            self.window_height.set(settings.get("window_height", 800))
            self.left_panel_width.set(settings.get("left_panel_width", 900))
            self.font_size.set(settings.get("font_size", 50))
            self.font_color.set(settings.get("font_color", "#000000"))
            self.selected_font.set(settings.get("selected_font", "Arial"))
            self.line_spacing.set(settings.get("line_spacing", 5))
            self.preview_text.set(settings.get("preview_text", "Иванов Иван Иванович"))
            
            # Обновляем список шрифтов
            if "available_fonts" in settings:
                self.available_fonts = settings["available_fonts"]
            
            # Обновляем интерфейс
            self.update_interface_labels()
            
            # Обновляем размеры окна
            self.root.geometry(f"{self.window_width.get()}x{self.window_height.get()}")
            
            # Обновляем ширину левой панели
            if hasattr(self, 'left_panel'):
                self.left_panel.configure(width=self.left_panel_width.get())
            
            # Загружаем шаблон если путь указан
            if self.template_path and os.path.exists(self.template_path):
                self.load_template_image()
            
            messagebox.showinfo("Успех", f"Настройки загружены из файла:\n{file_path}")
                
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить настройки: {str(e)}")
    
    def update_interface_labels(self):
        """Обновляет метки интерфейса после загрузки настроек"""
        if self.template_path:
            self.template_label.config(text=os.path.basename(self.template_path), foreground="black")
        else:
            self.template_label.config(text="Не выбран", foreground="gray")
            
        if self.excel_path:
            self.excel_label.config(text=os.path.basename(self.excel_path), foreground="black")
        else:
            self.excel_label.config(text="Не выбран", foreground="gray")
            
        if self.output_folder:
            self.output_label.config(text=os.path.basename(self.output_folder), foreground="black")
        else:
            self.output_label.config(text="Создается автоматически", foreground="blue")
    
    def on_area_scale_change(self, value):
        """Обработчик изменения позиции области ползунками"""
        # Обновляем X2 и Y2 чтобы сохранить размер области
        current_width = self.text_area_x2.get() - self.text_area_x1.get()
        current_height = self.text_area_y2.get() - self.text_area_y1.get()
        
        self.text_area_x2.set(self.text_area_x1.get() + current_width)
        self.text_area_y2.set(self.text_area_y1.get() + current_height)
    
    def on_width_scale_change(self, value):
        """Обработчик изменения ширины области"""
        width = int(float(value))
        self.text_area_x2.set(self.text_area_x1.get() + width)
    
    def on_height_scale_change(self, value):
        """Обработчик изменения высоты области"""
        height = int(float(value))
        self.text_area_y2.set(self.text_area_y1.get() + height)
    
    def on_window_size_change(self, value):
        """Обновляет размер окна при изменении через ползунки"""
        self.root.geometry(f"{self.window_width.get()}x{self.window_height.get()}")
        # Принудительно обновляем все панели
        self.root.update_idletasks()
    
    def on_panel_width_change(self, value):
        """Обновляет ширину левой панели при изменении через ползунок"""
        # PanedWindow автоматически управляет размерами
        pass
    
    def on_window_configure(self, event):
        """Обработчик изменения размера окна"""
        # PanedWindow автоматически управляет размерами панелей
        if event.widget == self.root:
            self.root.update_idletasks()
    
    def force_update_layout(self):
        """Принудительно обновляет макет всех панелей"""
        self.root.update_idletasks()
    
    def force_update_right_panel(self):
        """Принудительно обновляет правую панель"""
        self.root.update_idletasks()
            
    def calculate_text_position(self, text, font, alignment):
        """Вычисляет позицию текста в зависимости от выравнивания"""
        if self.text_mode.get() == "point":
            # Старый способ - одна точка
            return self.text_x.get(), self.text_y.get()
        
        # Новый способ - область с отступами
        x1, y1 = self.text_area_x1.get(), self.text_area_y1.get()
        x2, y2 = self.text_area_x2.get(), self.text_area_y2.get()
        
        # Применяем отступы
        text_x1 = x1 + self.text_padding_left.get()
        text_y1 = y1 + self.text_padding_top.get()
        text_x2 = x2 - self.text_padding_right.get()
        text_y2 = y2 - self.text_padding_bottom.get()
        
        # Вычисляем максимальную ширину области с учетом отступов
        max_width = text_x2 - text_x1
        
        # Вычисляем позицию по Y (по центру области с отступами)
        y = text_y1 + (text_y2 - text_y1) // 2
        
        # Для многострочного текста всегда возвращаем левый край области как x
        # Выравнивание будет обрабатываться в draw_multiline_text
        x = text_x1
            
        return x, y, max_width
    
    def wrap_text_to_lines(self, text, font, max_width):
        """Разбивает текст на строки, чтобы поместиться в заданную ширину"""
        words = text.split()
        lines = []
        current_line = ""
        
        for word in words:
            # Проверяем, поместится ли слово в текущую строку
            test_line = current_line + (" " if current_line else "") + word
            bbox = font.getbbox(test_line)
            line_width = bbox[2] - bbox[0]
            
            if line_width <= max_width:
                current_line = test_line
            else:
                # Если текущая строка не пустая, добавляем её
                if current_line:
                    lines.append(current_line)
                    current_line = word
                else:
                    # Если даже одно слово не помещается, добавляем его как есть
                    lines.append(word)
                    current_line = ""
        
        # Добавляем последнюю строку
        if current_line:
            lines.append(current_line)
            
        return lines
    
    def draw_multiline_text(self, draw, text, font, x, y, alignment, max_width, line_spacing):
        """Рисует многострочный текст с выравниванием"""
        lines = self.wrap_text_to_lines(text, font, max_width)
        print(f"Разбивка текста на строки: {lines}")
        
        # Получаем высоту строки
        bbox = font.getbbox("Ay")  # Используем символы с верхними и нижними выносами
        line_height = bbox[3] - bbox[1] + line_spacing
        
        # Вычисляем общую высоту текста
        total_height = len(lines) * line_height - line_spacing
        
        # Начинаем рисовать с верхней позиции
        start_y = y - total_height // 2
        
        for i, line in enumerate(lines):
            line_y = start_y + i * line_height
            
            # Вычисляем позицию X для каждой строки относительно левого края области
            bbox = font.getbbox(line)
            line_width = bbox[2] - bbox[0]
            
            if alignment == "left":
                line_x = x  # x - это левый край области
            elif alignment == "right":
                line_x = x + max_width - line_width  # x + ширина - ширина строки
            else:  # center
                line_x = x + (max_width - line_width) // 2  # x + половина свободного места
            
            print(f"Строка {i+1}: '{line}', позиция ({line_x}, {line_y}), ширина строки: {line_width}, ширина области: {max_width}")
            draw.text((line_x, line_y), line, fill=self.font_color.get(), font=font)
    
    def get_drag_type(self, x, y):
        """Определяет тип перетаскивания по позиции мыши"""
        if self.text_mode.get() != "area":
            return None
            
        x1, y1 = self.text_area_x1.get(), self.text_area_y1.get()
        x2, y2 = self.text_area_x2.get(), self.text_area_y2.get()
        
        # Если выбран инструмент "move", всегда перемещаем
        if self.tool_mode.get() == "move":
            if x1 <= x <= x2 and y1 <= y <= y2:
                return 'move'
            return None
        
        # Если выбран инструмент "resize", проверяем углы
        if self.tool_mode.get() == "resize":
            handle_size = 8
            
            # Проверяем углы для изменения размера
            if abs(x - x1) <= handle_size and abs(y - y1) <= handle_size:
                return 'resize_tl'  # top-left
            elif abs(x - x2) <= handle_size and abs(y - y1) <= handle_size:
                return 'resize_tr'  # top-right
            elif abs(x - x1) <= handle_size and abs(y - y2) <= handle_size:
                return 'resize_bl'  # bottom-left
            elif abs(x - x2) <= handle_size and abs(y - y2) <= handle_size:
                return 'resize_br'  # bottom-right
            
            # Если кликнули внутри области, тоже перемещаем
            if x1 <= x <= x2 and y1 <= y <= y2:
                return 'move'
                
        return None
    
    def canvas_to_image_coords(self, canvas_x, canvas_y):
        """Конвертирует координаты canvas в координаты изображения"""
        if not self.original_image:
            return 0, 0
            
        canvas_width = self.canvas.winfo_width()
        canvas_height = self.canvas.winfo_height()
        img_width, img_height = self.original_image.size
        
        # Вычисляем позицию относительно изображения
        click_x = canvas_x - (canvas_width - img_width * self.image_scale) // 2
        click_y = canvas_y - (canvas_height - img_height * self.image_scale) // 2
        
        # Конвертируем в координаты оригинального изображения
        original_x = int(click_x / self.image_scale)
        original_y = int(click_y / self.image_scale)
        
        return original_x, original_y
            
    def create_output_folder(self):
        """Создает папку с именем дата-время-сертификат"""
        now = datetime.now()
        folder_name = now.strftime("%Y-%m-%d_%H-%M-%S_сертификаты")
        folder_path = os.path.join(os.getcwd(), folder_name)
        
        try:
            os.makedirs(folder_path, exist_ok=True)
            return folder_path
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось создать папку: {str(e)}")
            return None
            
    def generate_certificates(self):
        if not all([self.template_path, self.excel_path]):
            messagebox.showerror("Ошибка", "Выберите шаблон сертификата и файл с ФИО")
            return
            
        if not self.original_image:
            messagebox.showerror("Ошибка", "Сначала загрузите шаблон сертификата")
            return
            
        # Создаем папку для сохранения
        output_folder = self.create_output_folder()
        if not output_folder:
            return
            
        try:
            # Читаем файл (Excel или CSV)
            if self.excel_path.endswith('.csv'):
                df = pd.read_csv(self.excel_path, encoding='utf-8-sig')
            else:
                df = pd.read_excel(self.excel_path)
            
            # Ищем только колонку "ФИО"
            name_column = None
            
            for col in df.columns:
                if col.strip() == 'ФИО':
                    name_column = col
                    break
                    
            if name_column is None:
                messagebox.showerror("Ошибка", "Не найдена колонка 'ФИО' в файле. Убедитесь, что в файле есть колонка с точным названием 'ФИО'.")
                return
                
            names = df[name_column].dropna().tolist()
            
            if not names:
                messagebox.showerror("Ошибка", "Не найдены данные в Excel файле")
                return
                
            # Настраиваем прогресс бар
            self.progress['maximum'] = len(names)
            self.progress['value'] = 0
            
            # Получаем шрифт
            font = self.get_font(self.font_size.get())
                
            # Генерируем сертификаты
            for i, name in enumerate(names):
                # Создаем копию оригинального изображения
                cert_img = self.original_image.copy()
                draw = ImageDraw.Draw(cert_img)
                
                # Добавляем границу вокруг сертификата
                img_width, img_height = cert_img.size
                border_width = 3
                draw.rectangle([0, 0, img_width-1, img_height-1], outline="#CCCCCC", width=border_width)
                
                # Добавляем ФИО
                if self.text_mode.get() == "point":
                    # Старый способ - одна точка
                    x, y = self.calculate_text_position(str(name), font, self.text_alignment.get())
                    draw.text((x, y), str(name), fill=self.font_color.get(), font=font)
                else:
                    # Новый способ - область с переносом строк
                    x, y, max_width = self.calculate_text_position(str(name), font, self.text_alignment.get())
                    self.draw_multiline_text(draw, str(name), font, x, y, self.text_alignment.get(), 
                                           max_width, self.line_spacing.get())
                
                # Сохраняем сертификат
                safe_name = "".join(c for c in str(name) if c.isalnum() or c in (' ', '-', '_')).rstrip()
                output_path = os.path.join(output_folder, f"certificate_{i+1}_{safe_name}.png")
                cert_img.save(output_path)
                
                # Обновляем прогресс
                self.progress['value'] = i + 1
                self.status_label.config(text=f"Обработано: {i+1}/{len(names)}")
                self.root.update()
                
            messagebox.showinfo("Успех", f"Сгенерировано {len(names)} сертификатов в папке:\n{output_folder}")
            self.status_label.config(text="Готово!")
            
            # Обновляем отображение папки в интерфейсе
            self.output_folder = output_folder
            self.output_label.config(text=os.path.basename(output_folder), foreground="black")
            
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при генерации сертификатов: {str(e)}")
            self.status_label.config(text="Ошибка")

def main():
    root = tk.Tk()
    app = CertificateGenerator(root)
    root.mainloop()

if __name__ == "__main__":
    main()
