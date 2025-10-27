import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageDraw, ImageFont, ImageTk
import pandas as pd
import os
from pathlib import Path
import threading
import matplotlib.font_manager as fm
from datetime import datetime

class CertificateGenerator:
    def __init__(self, root):
        self.root = root
        self.root.title("Генератор сертификатов")
        self.root.geometry("1200x800")
        
        # Переменные для хранения файлов
        self.template_path = None
        self.excel_path = None
        self.output_folder = os.getcwd()  # Текущая папка по умолчанию
        
        # Изображение и его размеры
        self.original_image = None
        self.display_image = None
        self.image_scale = 1.0
        
        # Координаты для размещения ФИО
        self.text_x = tk.IntVar(value=400)
        self.text_y = tk.IntVar(value=300)
        
        # Настройки шрифта
        self.font_size = tk.IntVar(value=50)
        self.font_color = tk.StringVar(value="#000000")
        self.selected_font = tk.StringVar(value="Arial")
        
        # Тестовый текст для предварительного просмотра
        self.preview_text = tk.StringVar(value="Иванов Иван Иванович")
        
        # Список доступных шрифтов
        self.available_fonts = []
        self.load_system_fonts()
        
        # Переменные для отслеживания изменений
        self.last_update_time = 0
        self.update_delay = 100  # миллисекунды
        
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
        
        # Левая панель с настройками
        left_frame = ttk.Frame(main_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        
        # Правая панель с предварительным просмотром
        right_frame = ttk.Frame(main_frame)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # Заголовок
        title_label = ttk.Label(left_frame, text="Генератор сертификатов", 
                               font=("Arial", 16, "bold"))
        title_label.pack(pady=(0, 20))
        
        # Секция загрузки файлов
        files_frame = ttk.LabelFrame(left_frame, text="Файлы", padding="10")
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
        settings_frame = ttk.LabelFrame(left_frame, text="Настройки текста", padding="10")
        settings_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Тестовый текст
        ttk.Label(settings_frame, text="Тестовый текст:").pack(anchor=tk.W, pady=2)
        preview_entry = ttk.Entry(settings_frame, textvariable=self.preview_text, width=30)
        preview_entry.pack(fill=tk.X, pady=2)
        
        # Координаты
        coords_frame = ttk.Frame(settings_frame)
        coords_frame.pack(fill=tk.X, pady=5)
        
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
        font_spinbox = ttk.Spinbox(font_frame, from_=10, to=200, width=8, 
                                  textvariable=self.font_size, command=self.update_preview)
        font_spinbox.pack(side=tk.LEFT, padx=(5, 10))
        
        # Цвет шрифта
        ttk.Label(font_frame, text="Цвет:").pack(side=tk.LEFT)
        color_entry = ttk.Entry(font_frame, textvariable=self.font_color, width=10)
        color_entry.pack(side=tk.LEFT, padx=(5, 0))
        
        # Привязка событий для автоматического обновления
        self.text_x.trace('w', self.schedule_update)
        self.text_y.trace('w', self.schedule_update)
        self.font_size.trace('w', self.schedule_update)
        self.font_color.trace('w', self.schedule_update)
        self.preview_text.trace('w', self.schedule_update)
        self.selected_font.trace('w', self.schedule_update)
        
        # Секция генерации
        generate_frame = ttk.Frame(left_frame)
        generate_frame.pack(fill=tk.X, pady=10)
        
        self.generate_button = ttk.Button(generate_frame, text="Генерировать сертификаты", 
                                         command=self.generate_certificates)
        self.generate_button.pack(fill=tk.X)
        
        # Прогресс бар
        self.progress = ttk.Progressbar(left_frame, mode='determinate')
        self.progress.pack(fill=tk.X, pady=5)
        
        # Статус
        self.status_label = ttk.Label(left_frame, text="Готов к работе")
        self.status_label.pack(pady=5)
        
        # Правая панель с предварительным просмотром
        preview_frame = ttk.LabelFrame(right_frame, text="Предварительный просмотр", padding="10")
        preview_frame.pack(fill=tk.BOTH, expand=True)
        
        # Canvas для отображения изображения с границей
        self.canvas = tk.Canvas(preview_frame, bg="white", cursor="crosshair", 
                               relief=tk.SUNKEN, bd=2)
        self.canvas.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Привязка событий мыши
        self.canvas.bind("<Button-1>", self.on_canvas_click)
        self.canvas.bind("<Motion>", self.on_canvas_motion)
        
        # Инструкция
        instruction_label = ttk.Label(preview_frame, 
                                    text="Кликните на изображении, чтобы установить позицию текста",
                                    font=("Arial", 10, "italic"))
        instruction_label.pack(pady=5)
        
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
            
        # Получаем размеры canvas и изображения
        canvas_width = self.canvas.winfo_width()
        canvas_height = self.canvas.winfo_height()
        img_width, img_height = self.original_image.size
        
        # Вычисляем позицию клика относительно изображения
        click_x = event.x - (canvas_width - img_width * self.image_scale) // 2
        click_y = event.y - (canvas_height - img_height * self.image_scale) // 2
        
        # Конвертируем в координаты оригинального изображения
        original_x = int(click_x / self.image_scale)
        original_y = int(click_y / self.image_scale)
        
        # Проверяем, что клик был по изображению
        if 0 <= original_x < img_width and 0 <= original_y < img_height:
            self.text_x.set(original_x)
            self.text_y.set(original_y)
            
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
            
    def schedule_update(self, *args):
        """Планирует обновление предварительного просмотра с задержкой"""
        current_time = self.root.after_idle(lambda: None)
        if hasattr(self, '_update_job'):
            self.root.after_cancel(self._update_job)
        self._update_job = self.root.after(self.update_delay, self.update_preview)
        
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
                print(f"Добавляем текст: '{text}' в позицию ({self.text_x.get()}, {self.text_y.get()})")
                draw.text((self.text_x.get(), self.text_y.get()), text, 
                         fill=self.font_color.get(), font=font)
            
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
                draw.text((self.text_x.get(), self.text_y.get()), str(name), 
                         fill=self.font_color.get(), font=font)
                
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
