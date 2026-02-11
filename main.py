import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.utils import ImageReader
from reportlab.graphics.barcode import eanbc
from reportlab.graphics.shapes import Drawing
from reportlab.graphics import renderPDF
from PIL import Image
import os
import re
from typing import Dict, Optional
from LabelsMarksGenerator.barcode.writer import ImageWriter
from io import BytesIO
import logging
import threading
import shutil
import sys
from datetime import datetime
from barcode.writer import ImageWriter
import barcode

try:
    from log import Log
except ImportError:
    print(
        "Библиотека логирования не установлена. Установите через: pip install git+ssh://git@github.com/AlexMayka/logging_python.git")

    class Log:
        def __init__(self, token=None, **kwargs):
            self.token = token
            print("Логгер инициализирован в режиме заглушки (библиотека не установлена)")

        def info(self, msg):
            print(f"INFO: {msg}")
            return type('Response', (), {'status_code': 201})()

        def error(self, msg):
            print(f"ERROR: {msg}")
            return type('Response', (), {'status_code': 201})()

        def warning(self, msg):
            print(f"WARNING: {msg}")
            return type('Response', (), {'status_code': 201})()

        def debug(self, msg):
            print(f"DEBUG: {msg}")
            return type('Response', (), {'status_code': 201})()

        def critical(self, msg):
            print(f"CRITICAL: {msg}")
            return type('Response', (), {'status_code': 201})()

        def __enter__(self):
            return self

        def __exit__(self, exc_type, exc_val, exc_tb):
            if exc_type:
                self.error(f"Исключение в контекстном менеджере: {exc_val}")
            return False

        def finish_success(self, period_from, period_to, **kwargs):
            print(f"SUCCESS: Завершено с {period_from} по {period_to}")
            return type('Response', (), {'status_code': 201})()

        def finish_error(self, period_from, period_to, **kwargs):
            print(f"ERROR: Завершено с ошибкой с {period_from} по {period_to}")
            return type('Response', (), {'status_code': 201})()


# Пытаемся импортировать tkinter для GUI, но не падаем, если его нет (например, на macOS без Tk)
try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
    TK_AVAILABLE = True
except Exception:
    tk = None
    filedialog = None
    messagebox = None
    ttk = None
    TK_AVAILABLE = False

# Токен для логирования - замените на ваш действительный токен
TOKEN = "10619b22-ca9f-404c-bf91-977530ac0e1a"

# Настройка логирования
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class Config:
    LABEL_SIZE_PX = (472, 472)  # 40mm x 40mm at 300 DPI
    PAGE_SIZE = (40 * mm, 40 * mm)


class ResourceManager:
    def __init__(self):
        self.logger = Log(token=TOKEN, silent_errors=True)

    @staticmethod
    def get_image(path: str) -> Optional[Image.Image]:
        try:
            image = Image.open(path)
            # Конвертируем в RGB если нужно, убираем прозрачность
            if image.mode in ('RGBA', 'LA', 'P'):
                # Создаем белый фон
                background = Image.new('RGB', image.size, (255, 255, 255))
                if image.mode == 'P':
                    image = image.convert('RGBA')
                background.paste(image, mask=image.split()[-1] if image.mode == 'RGBA' else None)
                return background
            else:
                return image.convert('RGB')
        except Exception as e:
            Log(token=TOKEN, silent_errors=True).error(f"Image load error {path}: {e}")
            return None


class MarkGenerator:
    def __init__(self):
        self.config = Config
        self.resource_manager = ResourceManager()
        self.logger = Log(token=TOKEN, silent_errors=True)

    def generate_pdf(self, data: Dict, output_pdf_path: str) -> bool:
        try:
            # Создаем PDF canvas
            c = canvas.Canvas(output_pdf_path, pagesize=self.config.PAGE_SIZE)

            # Получаем данные
            article = data.get('артикул', '')
            code = data.get('код', '')
            logo_name = data.get('лого', '').strip().lower()

            # Загружаем изображение марки
            mark_image = None
            mark_image_dir = "LabelsMarksGenerator/img/mark_images"
            extensions = ['.png', '.jpg', '.jpeg', '.bmp']
            for ext in extensions:
                test_path = os.path.join(mark_image_dir, f"mark_images{ext}")
                if os.path.exists(test_path):
                    mark_image = self.resource_manager.get_image(test_path)
                    if mark_image:
                        break

            # Рисуем изображение марки
            if mark_image:
                try:
                    # Масштабируем изображение (максимальная ширина 270px)
                    original_width, original_height = mark_image.size
                    scale_factor = 270.0 / original_width
                    new_width = 270
                    new_height = int(original_height * scale_factor)

                    # Ресайзим изображение
                    mark_image_resized = mark_image.resize((new_width, new_height), Image.Resampling.LANCZOS)

                    # Конвертируем в RGB для PDF
                    if mark_image_resized.mode != 'RGB':
                        mark_image_resized = mark_image_resized.convert('RGB')

                    # Сохраняем во временный буфер
                    temp_buffer = BytesIO()
                    mark_image_resized.save(temp_buffer, format='JPEG',
                                            quality=95)  # Используем JPEG для надежности
                    temp_buffer.seek(0)

                    # Рассчитываем размеры для PDF (конвертируем пиксели в мм)
                    mark_pdf_width = new_width * (40 * mm / 472)
                    mark_pdf_height = new_height * (40 * mm / 472)

                    # Позиционируем вверху слева с небольшими отступами
                    x_pos = 0.5 * mm
                    y_pos = self.config.PAGE_SIZE[1] - mark_pdf_height - 0.5 * mm

                    # Рисуем изображение в PDF
                    c.drawImage(ImageReader(temp_buffer), x_pos, y_pos,
                                mark_pdf_width, mark_pdf_height)

                except Exception as e:
                    # Продолжаем без изображения марки
                    pass

            # Загружаем и рисуем логотип
            logo_image = None
            if logo_name:
                logo_dir = "LabelsMarksGenerator/img/logos"
                extensions = ['.png', '.jpg', '.jpeg', '.bmp']
                for ext in extensions:
                    logo_path = os.path.join(logo_dir, f"{logo_name}{ext}")
                    if os.path.exists(logo_path):
                        try:
                            logo_image = self.resource_manager.get_image(logo_path)
                            if logo_image:

                                # Масштабируем логотип (максимальный размер 200px)
                                logo_image.thumbnail((200, 200), Image.Resampling.LANCZOS)
                                logo_width, logo_height = logo_image.size

                                # Конвертируем в RGB
                                if logo_image.mode != 'RGB':
                                    logo_image = logo_image.convert('RGB')

                                # Сохраняем во временный буфер
                                temp_buffer_logo = BytesIO()
                                logo_image.save(temp_buffer_logo, format='JPEG', quality=95)
                                temp_buffer_logo.seek(0)

                                # Рассчитываем размеры для PDF
                                logo_pdf_width = logo_width * (40 * mm / 472)
                                logo_pdf_height = logo_height * (40 * mm / 472)

                                # Позиционируем логотип под маркой
                                logo_x = 0.5 * mm
                                if mark_image:
                                    logo_y = self.config.PAGE_SIZE[1] - mark_pdf_height - logo_pdf_height - 1 * mm
                                else:
                                    logo_y = self.config.PAGE_SIZE[1] - logo_pdf_height - 0.5 * mm

                                # Рисуем логотип
                                c.drawImage(ImageReader(temp_buffer_logo), logo_x, logo_y,
                                            logo_pdf_width, logo_pdf_height)
                                break

                        except Exception as e:
                            continue

            # Настраиваем шрифты
            try:
                # Пробуем зарегистрировать шрифты
                font_paths = [
                    'arial.ttf',
                    'arialbd.ttf',
                    os.path.join(os.environ.get('WINDIR', ''), 'Fonts', 'arial.ttf'),
                    os.path.join(os.environ.get('WINDIR', ''), 'Fonts', 'arialbd.ttf'),
                    '/usr/share/fonts/truetype/freefont/FreeSans.ttf',
                    '/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf'
                ]

                arial_registered = False
                arial_bold_registered = False

                for font_path in font_paths:
                    if os.path.exists(font_path):
                        try:
                            if 'bd' in font_path.lower() or 'bold' in font_path.lower():
                                pdfmetrics.registerFont(TTFont('Arial-Bold', font_path))
                                arial_bold_registered = True
                            else:
                                pdfmetrics.registerFont(TTFont('Arial', font_path))
                                arial_registered = True
                        except:
                            continue

                if arial_registered and arial_bold_registered:
                    font_title = 'Arial-Bold'
                    font_regular = 'Arial'
                else:
                    # Используем стандартные шрифты
                    font_title = 'Helvetica-Bold'
                    font_regular = 'Helvetica'

            except Exception as e:
                font_title = 'Helvetica-Bold'
                font_regular = 'Helvetica'

            # Рассчитываем стартовую позицию для текста
            if logo_image:
                start_y_offset = logo_pdf_height + 15 * mm  # Отступ под логотипом(кажется лого не учитывается)
            else:
                start_y_offset = mark_pdf_height + 5 * mm  # Отступ от верха

            y_position = self.config.PAGE_SIZE[1] - start_y_offset

            # Текст артикула
            if code:
                c.setFont(font_title, 5)
                c.drawString(3 * mm, y_position, f"Артикул: {code}")

            # Остальной текст
            y_position -= 3 * mm
            c.setFont(font_regular, 5)
            c.drawString(3 * mm, y_position, "Количество:")

            y_position -= 3 * mm
            c.drawString(3 * mm, y_position, "Вес нетто:")

            # Текст "кг" справа
            kg_text = "кг"
            kg_width = c.stringWidth(kg_text, font_regular, 5)
            c.drawString(self.config.PAGE_SIZE[0] - 3 * mm - kg_width,
                         y_position, kg_text)

            y_position -= 3 * mm
            c.drawString(3 * mm, y_position, "Вес брутто:")
            c.drawString(self.config.PAGE_SIZE[0] - 3 * mm - kg_width,
                         y_position, kg_text)

            c.save()
            return True

        except Exception as e:
            return False


class PDFLabelGenerator:
    def __init__(self):
        self.page_width = 40 * mm
        self.page_height = 40 * mm
        self.logo_width = 18 * mm
        self.logo_height = 5.76 * mm
        self.cert_sign_size = 4 * mm
        self.logger = Log(token=TOKEN, silent_errors=True)
        self._register_fonts()

    def _register_fonts(self):
        """Регистрирует шрифты Arial с fallback на стандартные шрифты."""
        try:
            # Пробуем найти Arial на разных платформах
            font_paths = [
                'arial.ttf',
                'arialbd.ttf',
                os.path.join(os.environ.get('WINDIR', ''), 'Fonts', 'arial.ttf'),
                os.path.join(os.environ.get('WINDIR', ''), 'Fonts', 'arialbd.ttf'),
                '/System/Library/Fonts/Supplemental/Arial.ttf',
                '/System/Library/Fonts/Supplemental/Arial Bold.ttf',
                '/Library/Fonts/Arial.ttf',
                '/Library/Fonts/Arial Bold.ttf',
            ]

            arial_registered = False
            arial_bold_registered = False

            for font_path in font_paths:
                if os.path.exists(font_path):
                    try:
                        if 'bd' in font_path.lower() or 'bold' in font_path.lower():
                            if not arial_bold_registered:
                                pdfmetrics.registerFont(TTFont('Arial-Bold', font_path))
                                arial_bold_registered = True
                        else:
                            if not arial_registered:
                                pdfmetrics.registerFont(TTFont('Arial', font_path))
                                arial_registered = True
                    except Exception as e:
                        logger.debug(f"Не удалось зарегистрировать шрифт {font_path}: {e}")
                        continue

            # Если не удалось зарегистрировать, используем стандартные шрифты
            if not arial_registered or not arial_bold_registered:
                logger.info("Arial не найден, будут использованы стандартные шрифты Helvetica")
        except Exception as e:
            logger.debug(f"Ошибка при регистрации шрифтов: {e}")

    def _get_fonts(self):
        """Возвращает имена шрифтов с fallback на стандартные."""
        try:
            # Проверяем, зарегистрированы ли шрифты
            if pdfmetrics.getFont('Arial-Bold') and pdfmetrics.getFont('Arial'):
                return 'Arial-Bold', 'Arial-Bold', 'Arial', 'Arial'
        except:
            pass
        
        # Fallback на стандартные шрифты
        return 'Helvetica-Bold', 'Helvetica-Bold', 'Helvetica', 'Helvetica'

    def normalize_column_name(self, col_name: str) -> str:
        if pd.isna(col_name):
            return ""
        col_name = str(col_name).strip()
        col_name = re.sub(r'\s+', ' ', col_name)
        return col_name.lower()

    def normalize_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        column_mapping = {
            'наименование': ['название', 'product', 'name', 'товар'],
            'артикул': ['арт', 'article', 'sku', 'код товара'],
            'штрихкод': ['barcode', 'штрих-код', 'штрих код'],
            'сертификация': ['сертификат', 'certification'],
            'тип сертификации': ['тип сертификата', 'certification type'],
            'лого': ['logo', 'логотип'],
            'назначение': ['purpose', 'применение'],
            'материал': ['material', 'состав'],
            'производитель': ['manufacturer', 'producer'],
            'импортер': ['importer'],
            'страна происхождения': ['country', 'страна'],
            'дата изготовления': ['production date', 'дата'],
            'код': ['code', 'код товара']
        }

        df.columns = [self.normalize_column_name(col) for col in df.columns]

        for standard_name, variants in column_mapping.items():
            for col in df.columns:
                if col in variants:
                    df.rename(columns={col: standard_name}, inplace=True)

        return df

    def read_excel(self, file_path: str) -> Optional[pd.DataFrame]:
        try:
            engines = ['openpyxl', 'xlrd']
            df = None
            for engine in engines:
                try:
                    df = pd.read_excel(file_path, engine=engine)
                    break
                except Exception as e:
                    continue

            if df is None:
                return None

            df = self.normalize_columns(df)
            df = df.fillna('')

            return df
        except Exception as e:
            return None

    def wrap_text(self, text: str, font_name: str, font_size: int, max_width: float, canvas_obj: canvas.Canvas) -> list:
        if not text:
            return []

        words = text.split()
        lines = []
        current_line = []

        for word in words:
            test_line = ' '.join(current_line + [word])
            text_width = canvas_obj.stringWidth(test_line, font_name, font_size)

            if text_width <= max_width:
                current_line.append(word)
            else:
                if current_line:
                    lines.append(' '.join(current_line))
                current_line = [word]

        if current_line:
            lines.append(' '.join(current_line))

        return lines

    def draw_ean13_barcode(
        self,
        canvas_obj: canvas.Canvas,
        barcode_value: str,
        x: float,
        y: float,
        width: float,
        height: float,
    ) -> bool:
        """
        Рисует векторный штрихкод EAN-13 непосредственно на canvas.
        Штрихкод и цифры под ним полностью векторные - масштабируются вместе с PDF
        и адаптируются к изменению размера страницы.
        """
        try:
            if not barcode_value:
                return False

            digits = re.sub(r'\D', '', str(barcode_value).strip())

            if len(digits) == 13:
                digits = digits[:12]
            if len(digits) != 12:
                logger.warning(f"Штрихкод должен содержать 12 или 13 цифр, получено {len(digits)}: {digits} из {barcode_value}")
                return False

            # Создаем виджет штрихкода с включенными цифрами (векторный текст)
            barcode_widget = eanbc.Ean13BarcodeWidget(digits)
            barcode_widget.humanReadable = True  # Цифры под штрихкодом как векторный текст

            # Получаем естественные размеры виджета (включая цифры)
            bounds = barcode_widget.getBounds()
            bw = bounds[2] - bounds[0]
            bh = bounds[3] - bounds[1]
            if bw <= 0 or bh <= 0:
                return False

            # Вычисляем коэффициент масштабирования для сохранения пропорций
            # Это гарантирует, что и штрихкод, и цифры масштабируются одинаково
            scale_x = width / bw
            scale_y = height / bh
            scale = min(scale_x, scale_y)  # Сохраняем пропорции
            
            # Создаем векторный drawing с исходными размерами
            # Все элементы (штрихкод + цифры) будут масштабироваться вместе
            drawing = Drawing(bw, bh)
            drawing.add(barcode_widget)

            # Масштабируем весь drawing (штрихкод + цифры) пропорционально
            # При изменении размера PDF все будет масштабироваться вместе
            drawing.scale(scale, scale)

            # Рисуем на canvas - координаты в мм, поэтому масштабируются вместе с PDF
            # Цифры остаются векторным текстом и масштабируются вместе со штрихкодом
            renderPDF.draw(drawing, canvas_obj, x, y)
            return True
        except Exception as e:
            logger.error(f"Ошибка при рисовании векторного штрихкода: {e}")
            return False

    def get_logo_image(self, logo_name: str) -> Optional[ImageReader]:
        if not logo_name:
            return None

        logo_dir = "LabelsMarksGenerator/img/logos"
        extensions = ['.png', '.jpg', '.jpeg', '.bmp']

        for ext in extensions:
            logo_path = os.path.join(logo_dir, f"{logo_name}{ext}")
            if os.path.exists(logo_path):
                try:
                    return ImageReader(logo_path)
                except Exception as e:
                    continue

        return None

    def get_certification_icon(self, certification_type: str) -> Optional[str]:
        if not certification_type:
            return None

        cert_dir = "LabelsMarksGenerator/img/certificates"
        os.makedirs(cert_dir, exist_ok=True)

        cert_type_lower = str(certification_type).lower().strip()

        if cert_type_lower in ['рст', 'rct', 'rst']:
            possible_paths = [
                os.path.join(cert_dir, "рст.png"),
                os.path.join(cert_dir, "rct.png"),
                os.path.join(cert_dir, "rst.png"),
                os.path.join(cert_dir, "рст.jpg"),
                os.path.join(cert_dir, "rct.jpg"),
                os.path.join(cert_dir, "rst.jpg")
            ]
        elif cert_type_lower in ['eac', 'еас']:
            possible_paths = [
                os.path.join(cert_dir, "eac.png"),
                os.path.join(cert_dir, "еас.png"),
                os.path.join(cert_dir, "eac.jpg"),
                os.path.join(cert_dir, "еас.jpg")
            ]
        else:
            return None

        for path in possible_paths:
            if os.path.exists(path):
                return path

        return None

    def create_label_pdf(self, data: Dict, output_path: str) -> bool:
        try:
            c = canvas.Canvas(output_path, pagesize=(self.page_width, self.page_height))

            # Регистрируем шрифты с fallback на стандартные
            font_large_bold, font_medium_bold, font_medium, font_small = self._get_fonts()

            large_font_size = 5
            medium_font_size = 4
            small_font_size = 3

            line_height = 1.6 * mm
            field_spacing = 0.3 * mm

            name = data.get('наименование', '')
            purpose = data.get('назначение', '')
            material = data.get('материал', '')
            manufacturer = data.get('производитель', '')
            importer = data.get('импортер', '')
            country = data.get('страна происхождения', '')
            production_date = data.get('дата изготовления', '')
            code = data.get('код', '')
            article = data.get('артикул', '')
            barcode_value = data.get('штрихкод', '')
            certification = data.get('сертификация', '')
            certification_type = data.get('тип сертификации', '')

            logo_name = data.get('лого', '').strip().lower()
            logo_img = self.get_logo_image(logo_name)
            has_logo = logo_img is not None

            if has_logo:
                try:
                    logo_x = self.page_width - self.logo_width - 1 * mm
                    logo_y = self.page_height - self.logo_height
                    c.drawImage(logo_img, logo_x, logo_y, self.logo_width, self.logo_height)
                except Exception as e:
                    pass

            x_position = 1 * mm
            y_position = self.page_height - 2 * mm

            if name:
                if has_logo:
                    text_width_for_name = (self.page_width - self.logo_width - 3 * mm)
                else:
                    text_width_for_name = self.page_width - 2 * mm

                lines_beside_logo = self.wrap_text(name, font_large_bold, large_font_size,
                                                   text_width_for_name, c)

                lines_used = 0
                for i, line in enumerate(lines_beside_logo):
                    current_y = y_position - (i * line_height)
                    if has_logo and current_y > (self.page_height - self.logo_height):
                        c.setFont(font_large_bold, large_font_size)
                        c.drawString(x_position, current_y, line)
                        lines_used += 1
                    elif not has_logo:
                        c.setFont(font_large_bold, large_font_size)
                        c.drawString(x_position, current_y, line)
                        lines_used += 1
                    else:
                        break

                next_line_y = y_position - (lines_used * line_height)

                if has_logo and lines_used < len(lines_beside_logo):
                    remaining_text = ' '.join(
                        name.split()[sum(len(line.split()) for line in lines_beside_logo[:lines_used]):])

                    if remaining_text:
                        full_width = self.page_width - 2 * mm
                        remaining_lines = self.wrap_text(remaining_text, font_large_bold, large_font_size,
                                                         full_width, c)

                        start_y = next_line_y

                        for j, line in enumerate(remaining_lines):
                            if start_y - (j * line_height) < 15 * mm:
                                break
                            c.setFont(font_large_bold, large_font_size)
                            c.drawString(x_position, start_y - (j * line_height), line)

                        y_position = start_y - len(remaining_lines) * line_height - field_spacing
                    else:
                        y_position = next_line_y - field_spacing
                else:
                    y_position = next_line_y - field_spacing

            priority_fields = [
                ("Назначение", purpose, font_medium),
                ("Материал", material, font_medium),
                ("Производитель", manufacturer, font_medium),
                ("Импортер", importer, font_medium),
                ("Страна происхождения", country, font_medium),
                ("Дата изготовления", production_date, font_medium)
            ]

            for field_name, field_value, font in priority_fields:
                if not field_value:
                    continue

                if y_position - line_height < 12 * mm:
                    break

                c.setFont(font_medium_bold, medium_font_size)
                field_text = f"{field_name}:"
                c.drawString(x_position, y_position, field_text)

                field_name_width = c.stringWidth(field_text, font_medium_bold, medium_font_size)
                value_x = x_position + field_name_width + 0.5 * mm
                value_max_width = self.page_width - value_x - 2 * mm

                value_lines = self.wrap_text(field_value, font, medium_font_size, value_max_width, c)
                if value_lines:
                    c.setFont(font, medium_font_size)
                    c.drawString(value_x, y_position, value_lines[0])

                    for i in range(1, len(value_lines)):
                        if y_position - (i * line_height) < 12 * mm:
                            break
                        c.drawString(x_position, y_position - (i * line_height), value_lines[i])

                    y_position -= max(1, len(value_lines)) * line_height
                else:
                    y_position -= line_height

                y_position -= field_spacing

            if code or article:
                code_article_text = ""
                if code:
                    code_article_text += f"Код: {code}"
                if article:
                    if code_article_text:
                        code_article_text += " "
                    code_article_text += f"Арт: {article}"

                if code_article_text:
                    c.setFont(font_small, small_font_size)
                    c.drawString(22 * mm, 10 * mm, code_article_text)

            if certification:
                cert_area_width = 18 * mm
                cert_area_x = 1 * mm

                cert_icon_path = None
                if certification_type:
                    cert_icon_path = self.get_certification_icon(certification_type)

                cert_text_y = 3 * mm

                if cert_icon_path:
                    try:
                        cert_icon = ImageReader(cert_icon_path)
                        cert_x = cert_area_x
                        cert_y = cert_text_y + 2 * mm
                        c.drawImage(cert_icon, cert_x, cert_y, self.cert_sign_size, self.cert_sign_size,
                                    mask='auto')
                    except Exception as e:
                        pass

                cert_text = str(certification)
                cert_lines = self.wrap_text(cert_text, font_small, small_font_size,
                                            cert_area_width, c)

                cert_lines = cert_lines[:2]

                c.setFont(font_small, small_font_size)
                for i, line in enumerate(cert_lines):
                    cert_y = cert_text_y - (i * 1 * mm)
                    if cert_y > 1 * mm:
                        c.drawString(cert_area_x, cert_y, line)

            if barcode_value:
                # Адаптивный размер штрихкода - занимает 50% ширины страницы
                # Минимальная ширина 15mm, максимальная 25mm для читаемости
                barcode_width = max(15 * mm, min(25 * mm, self.page_width * 0.5))
                # Высота пропорциональна ширине (соотношение 2.5:1 для EAN-13)
                barcode_height = barcode_width / 2.5
                
                # Позиция справа с небольшим отступом
                barcode_x = self.page_width - barcode_width - 1 * mm
                barcode_y = 1 * mm
                
                try:
                    self.draw_ean13_barcode(
                        c,
                        barcode_value,
                        barcode_x,
                        barcode_y,
                        barcode_width,
                        barcode_height,
                    )
                except Exception as e:
                    # Логируем ошибку, но продолжаем создание PDF
                    logger.warning(f"Ошибка при рисовании штрихкода: {e}")

            c.save()
            return True

        except Exception as e:
            logger.error(f"Ошибка при создании этикетки {output_path}: {e}")
            import traceback
            logger.error(traceback.format_exc())
            return False


class CombinedGenerator:
    def __init__(self):
        self.mark_generator = MarkGenerator()
        self.label_generator = PDFLabelGenerator()
        self.logger = Log(token=TOKEN, silent_errors=True)

    def normalize_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        return self.label_generator.normalize_columns(df)

    def read_excel(self, file_path: str) -> Optional[pd.DataFrame]:
        return self.label_generator.read_excel(file_path)

    def process_excel_file(self, excel_file_path: str, output_dir: str = "output"):
        try:
            df = self.read_excel(excel_file_path)
            if df is None or df.empty:
                self.logger.error("No data found in Excel file")
                return False

            total_rows = len(df)

            if 'штрихкод' in df.columns:
                df['штрихкод'] = df['штрихкод'].apply(lambda x:
                                                      str(int(x)) if pd.notna(x) and x != '' and str(x).replace(
                                                          '.0',
                                                          '').isdigit()
                                                      else str(x) if pd.notna(x) and x != ''
                                                      else '')

            os.makedirs(output_dir, exist_ok=True)
            os.makedirs(os.path.join(output_dir, "marks"), exist_ok=True)
            os.makedirs(os.path.join(output_dir, "labels"), exist_ok=True)

            success_count_marks = 0
            success_count_labels = 0

            for idx, row in df.iterrows():
                if not row.get('наименование'):
                    continue

                row_data = row.to_dict()

                if 'штрихкод' in row_data and row_data['штрихкод']:
                    barcode_val = row_data['штрихкод']
                    if isinstance(barcode_val, float) and barcode_val.is_integer():
                        row_data['штрихкод'] = str(int(barcode_val))
                    else:
                        row_data['штрихкод'] = str(barcode_val)

                article = str(row.get('артикул', '')).strip()
                code = str(row.get('код', '')).strip()
                article_clean = re.sub(r'[\\/*?:"<>|]', "_", article)
                code_clean = re.sub(r'[\\/*?:"<>|]', "_", code)

                base_filename = f"{article_clean}_{code_clean}" if article_clean or code_clean else f"row_{idx}"

                # Generate mark PDF directly
                mark_pdf_path = os.path.join(output_dir, "marks", f"mark_{base_filename}.pdf")
                if self.mark_generator.generate_pdf(row_data, mark_pdf_path):
                    success_count_marks += 1

                # Generate label PDF
                label_pdf_path = os.path.join(output_dir, "labels", f"label_{base_filename}.pdf")
                if self.label_generator.create_label_pdf(row_data, label_pdf_path):
                    success_count_labels += 1

                # Логируем прогресс каждые 25 записей или на последней записи
                current_progress = idx + 1
                if current_progress % 25 == 0 or current_progress == total_rows:
                    self.logger.info(f"Обработано {current_progress} из {total_rows} записей")

            # Финальное сообщение о результатах
            self.logger.info(f"Обработано файлов: 1, создано этикеток: {success_count_labels}, марок: {success_count_marks}")

            return success_count_marks > 0 or success_count_labels > 0

        except Exception as e:
            self.logger.error(f"Ошибка при формировании запроса: {e}")
            return False


class Application:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Генератор этикеток и марок")
        self.root.geometry("600x400")
        self.root.resizable(True, True)

        # Создаем необходимые директории
        os.makedirs('LabelsMarksGenerator/input', exist_ok=True)
        os.makedirs('LabelsMarksGenerator/output', exist_ok=True)
        os.makedirs('LabelsMarksGenerator/img/logos', exist_ok=True)
        os.makedirs('LabelsMarksGenerator/img/certificates', exist_ok=True)
        os.makedirs('LabelsMarksGenerator/img/mark_images', exist_ok=True)

        self.generator = CombinedGenerator()
        self.setup_ui()

    def setup_ui(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        title_label = ttk.Label(main_frame, text="Генератор этикеток и марок", font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))

        instruction_text = """
        Инструкция:
        1. Выберите Excel-файл с данными
        2. Нажмите кнопку 'Обработать файл'
        3. Результаты появятся в папке 'output'

        Структура папки output:
        - labels/ - этикетки в формате PDF
        - marks/ - маркировки в формате PDF
        """
        instruction_label = ttk.Label(main_frame, text=instruction_text, justify=tk.LEFT)
        instruction_label.grid(row=1, column=0, columnspan=2, pady=(0, 20))

        select_btn = ttk.Button(main_frame, text="Выбрать файл", command=self.select_file)
        select_btn.grid(row=2, column=0, pady=5, padx=5, sticky=tk.EW)

        process_btn = ttk.Button(main_frame, text="Обработать файл", command=self.process_file)
        process_btn.grid(row=2, column=1, pady=5, padx=5, sticky=tk.EW)

        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=3, column=0, columnspan=2, pady=10, sticky=tk.EW)

        self.status_var = tk.StringVar(value="Готов к работе")
        status_label = ttk.Label(main_frame, textvariable=self.status_var)
        status_label.grid(row=4, column=0, columnspan=2, pady=5)

        log_frame = ttk.LabelFrame(main_frame, text="Лог выполнения", padding="5")
        log_frame.grid(row=5, column=0, columnspan=2, pady=10, sticky=(tk.W, tk.E, tk.N, tk.S))

        self.log_text = tk.Text(log_frame, height=10, width=70)
        log_scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scrollbar.set)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        log_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))

        self.file_var = tk.StringVar(value="Файл не выбран")
        file_label = ttk.Label(main_frame, textvariable=self.file_var)
        file_label.grid(row=6, column=0, columnspan=2, pady=5)

        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(5, weight=1)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        self.redirect_logging()

    def redirect_logging(self):
        class TextHandler(logging.Handler):
            def __init__(self, text_widget):
                super().__init__()
                self.text_widget = text_widget

            def emit(self, record):
                msg = self.format(record)
                self.text_widget.configure(state='normal')
                self.text_widget.insert(tk.END, msg + '\n')
                self.text_widget.configure(state='disabled')
                self.text_widget.see(tk.END)

        text_handler = TextHandler(self.log_text)
        text_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        logger.addHandler(text_handler)

    def select_file(self):
        file_path = filedialog.askopenfilename(
            title="Выберите Excel файл",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file_path:
            try:
                # Копируем файл в папку input
                filename = os.path.basename(file_path)
                input_path = os.path.join('LabelsMarksGenerator/input', filename)
                shutil.copy2(file_path, input_path)
                self.file_var.set(f"Выбран файл: {filename}")
                self.status_var.set("Файл готов к обработке")
                logger.info(f"Файл скопирован: {filename}")

                # Логируем событие выбора файла
                log = Log(token=TOKEN, silent_errors=True)
                log.info(f"Пользователь выбрал файл: {filename}")

            except Exception as e:
                logger.error(f"Ошибка копирования файла: {e}")
                messagebox.showerror("Ошибка", f"Не удалось скопировать файл: {e}")

                # Логируем ошибку
                log = Log(token=TOKEN, silent_errors=True)
                log.error(f"Ошибка копирования файла: {e}")

    def process_file(self):
        input_dir = 'LabelsMarksGenerator/input'
        excel_files = [f for f in os.listdir(input_dir) if f.endswith(('.xlsx', '.xls'))]

        if not excel_files:
            messagebox.showwarning("Предупреждение", "В папке 'input' нет Excel файлов")

            # Логируем предупреждение
            log = Log(token=TOKEN, silent_errors=True)
            log.warning("Попытка обработки при отсутствии файлов в папке input")
            return

        # Запускаем обработку в отдельном потоке
        thread = threading.Thread(target=self.process_files_thread)
        thread.daemon = True
        thread.start()

    def process_files_thread(self):
        start_time = datetime.now()

        try:
            self.progress.start()
            self.status_var.set("Обработка файлов...")

            input_dir = 'LabelsMarksGenerator/input'
            excel_files = [f for f in os.listdir(input_dir) if f.endswith(('.xlsx', '.xls'))]

            # Логируем начало обработки
            log = Log(token=TOKEN, silent_errors=True)
            log.info(f"Начало обработки {len(excel_files)} файлов")

            total_files = len(excel_files)
            processed_files = 0

            for excel_file in excel_files:
                excel_file_path = os.path.join(input_dir, excel_file)
                logger.info(f"Обработка файла: {excel_file}")

                success = self.generator.process_excel_file(excel_file_path, "output")

                if success:
                    logger.info(f"Файл {excel_file} успешно обработан")
                else:
                    logger.error(f"Ошибка при обработке файла {excel_file}")

                processed_files += 1
                logger.info(f"Обработано файлов: {processed_files} из {total_files}")

            logger.info("Обработка завершена!")
            self.status_var.set("Обработка завершена успешно!")
            messagebox.showinfo("Успех", "Обработка файлов завершена!\n\nРезультаты в папке 'output'")

            # Логируем успешное завершение в eff_runs
            end_time = datetime.now()
            log.finish_success(
                period_from=start_time,
                period_to=end_time,
                files_processed=total_files,
                duration_seconds=(end_time - start_time).total_seconds(),
                message=f"Обработано {total_files} файлов, созданы этикетки и марки"
            )

        except Exception as e:
            logger.error(f"Ошибка обработки: {e}")
            self.status_var.set(f"Ошибка: {e}")
            messagebox.showerror("Ошибка", f"Ошибка обработки: {e}")

            # Логируем ошибку завершения
            end_time = datetime.now()
            log = Log(token=TOKEN, silent_errors=True)
            log.finish_error(
                period_from=start_time,
                period_to=end_time,
                error=str(e),
                error_type=type(e).__name__,
                files_processed=processed_files if 'processed_files' in locals() else 0
            )
        finally:
            self.progress.stop()

    def run(self):
        # Логируем запуск приложения
        log = Log(token=TOKEN, silent_errors=True)
        log.info("Приложение генератора этикеток и марок запущено")

        self.root.mainloop()


def main():
    # Логируем запуск программы
    log = Log(token=TOKEN, silent_errors=True)
    log.info("Программа генератора этикеток и марок запущена")

    def run_console_mode():
        """Запуск обработки в консольном режиме (без GUI)."""
        log.info("Запуск в консольном режиме")

        generator = CombinedGenerator()
        input_dir = "LabelsMarksGenerator/input"
        output_dir = "LabelsMarksGenerator/output"

        if not os.path.exists(input_dir):
            log.error(f"Input directory '{input_dir}' not found")
            print(f"Input directory '{input_dir}' not found")
            return

        excel_files = []
        for file in os.listdir(input_dir):
            if file.endswith(('.xlsx', '.xls')):
                excel_files.append(os.path.join(input_dir, file))

        if not excel_files:
            log.warning("No Excel files found in input directory")
            print("No Excel files found in input directory")
            return

        start_time = datetime.now()
        total_files = len(excel_files)
        log.info(f"Начало обработки {total_files} файлов в консольном режиме")

        processed_files = 0
        for excel_file in excel_files:
            print(f"Processing: {excel_file}")
            generator.process_excel_file(excel_file, output_dir)
            processed_files += 1
            print(f"Обработано файлов: {processed_files} из {total_files}")

        end_time = datetime.now()
        log.finish_success(
            period_from=start_time,
            period_to=end_time,
            files_processed=total_files,
            duration_seconds=(end_time - start_time).total_seconds(),
            mode="console",
            message=f"Обработано {total_files} файлов, созданы этикетки и марки"
        )

    # Если явно указан консольный режим
    if len(sys.argv) > 1 and sys.argv[1] == '--console':
        run_console_mode()
    else:
        # Пытаемся запустить графический режим, если доступен tkinter
        if not TK_AVAILABLE:
            log.warning("Tkinter недоступен, запускаю в консольном режиме")
            print("Внимание: Tkinter недоступен, запуск в консольном режиме.")
            run_console_mode()
            return

        log.info("Запуск в графическом режиме")
        app = Application()
        app.run()


if __name__ == "__main__":
    main()