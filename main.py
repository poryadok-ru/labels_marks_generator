import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import barcode
from barcode.writer import ImageWriter
import os
import textwrap
import re
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import mm
from reportlab.lib.utils import ImageReader
import shutil
import logging
import stat
from typing import Dict, List, Optional, Tuple
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import sys
import threading

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

barcode_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'barcode')
if os.path.exists(barcode_path):
    sys.path.insert(0, barcode_path)

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('app.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


class Config:
    DPI = 300
    LABEL_SIZE_MM = (40, 40)
    LABEL_SIZE_PX = (472, 472)
    FONT_PATHS = {
        'regular': 'arial.ttf',
        'bold': 'arialbd.ttf'
    }
    OUTPUT_DIRS = [
        'output/labels_png',
        'output/labels_pdf',
        'output/marks_png',
        'output/marks_pdf'
    ]
    IMAGE_DIRS = [
        'img/certificates',
        'img/logos',
        'img/mark_images'
    ]


class ResourceManager:
    _cache = {}

    @classmethod
    def get_font(cls, font_type: str, size: int) -> ImageFont.FreeTypeFont:
        key = (font_type, size)
        if key not in cls._cache:
            try:
                font_path = Config.FONT_PATHS['bold'] if font_type == 'bold' else Config.FONT_PATHS['regular']

                if not os.path.exists(font_path):
                    if sys.platform == 'win32':
                        font_dir = os.path.join(os.environ['WINDIR'], 'Fonts')
                        font_path = os.path.join(font_dir, 'arialbd.ttf' if font_type == 'bold' else 'arial.ttf')

                    if not os.path.exists(font_path):
                        raise FileNotFoundError(f"Font not found: {font_path}")

                cls._cache[key] = ImageFont.truetype(font_path, size)
            except Exception as e:
                logger.warning(f"Font load error: {e}. Using default font.")
                cls._cache[key] = ImageFont.load_default()
        return cls._cache[key]

    @classmethod
    def get_image(cls, image_path: str) -> Optional[Image.Image]:
        if image_path not in cls._cache:
            try:
                if not os.path.exists(image_path):
                    logger.warning(f"Image not found: {image_path}")
                    return None

                img = Image.open(image_path)
                if img.mode != 'RGBA':
                    img = img.convert('RGBA')
                cls._cache[image_path] = img
            except Exception as e:
                logger.error(f"Image load error {image_path}: {e}")
                return None
        return cls._cache[image_path]


def handle_remove_readonly(func, path, exc_info):
    if not os.access(path, os.W_OK):
        os.chmod(path, stat.S_IWUSR)
        func(path)
    else:
        raise


def normalize_column_name(col_name: str) -> str:
    if pd.isna(col_name):
        return ""
    col_name = str(col_name).strip()
    col_name = re.sub(r'\s+', ' ', col_name)
    return col_name.lower()


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
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

    df.columns = [normalize_column_name(col) for col in df.columns]

    for standard_name, variants in column_mapping.items():
        for col in df.columns:
            if col in variants:
                df.rename(columns={col: standard_name}, inplace=True)

    return df


def read_excel(file_path: str) -> Optional[pd.DataFrame]:
    try:
        engines = ['openpyxl', 'xlrd']
        df = None

        for engine in engines:
            try:
                df = pd.read_excel(file_path, engine=engine)
                logger.info(f"File read with engine {engine}")
                break
            except Exception as e:
                logger.warning(f"Engine {engine} failed: {e}")
                continue

        if df is None:
            logger.error("All engines failed to read the file")
            return None

        df = normalize_columns(df)
        logger.info(f"Normalized columns: {list(df.columns)}")
        df = df.fillna('')

        reference_row = None
        for idx, row in df.iterrows():
            if any(row.values):
                reference_row = row
                break

        if reference_row is None:
            logger.error("No data found in file!")
            return df

        excluded_fields = ['код', 'артикул', 'штрихкод', 'сертификация', 'тип сертификации', 'лого']
        for idx, row in df.iterrows():
            for col in df.columns:
                if col not in excluded_fields and (row[col] == '' or pd.isna(row[col])):
                    df.at[idx, col] = reference_row[col]

        return df
    except Exception as e:
        logger.error(f"File read error: {e}")
        return None


def wrap_text(text: str, font: ImageFont.FreeTypeFont, max_width: int) -> List[str]:
    lines = []
    if not text:
        return lines

    words = text.split()
    current_line = []

    for word in words:
        test_line = ' '.join(current_line + [word])
        try:
            bbox = font.getbbox(test_line)
            text_width = bbox[2] - bbox[0]
        except AttributeError:
            text_width = font.getsize(test_line)[0]

        if text_width <= max_width:
            current_line.append(word)
        else:
            if current_line:
                lines.append(' '.join(current_line))
            current_line = [word]

    if current_line:
        lines.append(' '.join(current_line))

    return lines


def validate_data(row: pd.Series) -> bool:
    required_fields = ['наименование']
    for field in required_fields:
        if not row.get(field):
            logger.error(f"Missing required field {field} in row {row.name}")
            return False
    return True


class LabelGenerator:
    def __init__(self):
        self.config = Config
        self.resource_manager = ResourceManager

    def generate_barcode(self, barcode_value: str) -> Optional[Image.Image]:
        try:
            if not barcode_value:
                return None

            barcode_str = str(barcode_value).strip()
            if not barcode_str:
                return None

            import barcode
            from io import BytesIO

            code128 = barcode.codex.Code128(barcode_str, writer=barcode.writer.ImageWriter())
            buffer = BytesIO()
            code128.write(buffer)

            barcode_img = Image.open(buffer)
            barcode_img.load()

            if barcode_img.mode != 'RGB':
                barcode_img = barcode_img.convert('RGB')

            return barcode_img

        except Exception as e:
            logger.error(f"Barcode generation error {barcode_value}: {e}")
            return None

    def generate(self, data: Dict, output_path: str, output_pdf_path: Optional[str] = None) -> bool:
        try:
            img = Image.new('RGB', self.config.LABEL_SIZE_PX, color='white')
            draw = ImageDraw.Draw(img)

            font_large_bold = self.resource_manager.get_font('bold', 20)
            font_medium_bold = self.resource_manager.get_font('bold', 16)
            font_medium = self.resource_manager.get_font('regular', 16)

            logo_name = data.get('лого', '').strip().lower()
            logo_img = None
            logo_width = 0
            logo_height = 0

            if logo_name:
                logo_dir = "img/logos"
                extensions = ['.png', '.jpg', '.jpeg', '.bmp']
                for ext in extensions:
                    logo_path = os.path.join(logo_dir, f"{logo_name}{ext}")
                    if os.path.exists(logo_path):
                        try:
                            logo_img = self.resource_manager.get_image(logo_path)
                            if logo_img:
                                max_logo_size = 200
                                logo_img.thumbnail((max_logo_size, max_logo_size))
                                logo_width, logo_height = logo_img.size
                                img.paste(logo_img, (472 - logo_width - 5, 5), logo_img)
                                break
                        except Exception as e:
                            logger.error(f"Logo load error {logo_name}: {e}")
                            continue

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

            y_position = 10
            max_y = 470

            if logo_img:
                left_area_width = 472 - logo_width - 15
                under_logo_area_width = 462
            else:
                left_area_width = 462
                under_logo_area_width = 462

            name_lines = wrap_text(name, font_large_bold, left_area_width)
            for line in name_lines:
                if y_position + 20 > max_y:
                    break
                draw.text((10, y_position), line, font=font_large_bold, fill='black')
                y_position += 20

            y_position += 5

            fields = [
                ("Назначение", purpose),
                ("Материал", material),
                ("Производитель", manufacturer),
                ("Импортер", importer),
                ("Страна происхождения", country),
                ("Дата изготовления", production_date)
            ]

            for field_name, field_value in fields:
                if not field_value:
                    continue

                if y_position > max_y - 50:
                    break

                current_width = left_area_width if y_position < (5 + logo_height) else under_logo_area_width

                draw.text((10, y_position), f"{field_name}:", font=font_medium_bold, fill='black')

                try:
                    bbox = font_medium_bold.getbbox(field_name + ":")
                    field_name_width = bbox[2] - bbox[0]
                except AttributeError:
                    field_name_width = font_medium_bold.getsize(field_name + ":")[0]

                value_max_width = current_width - field_name_width - 5
                value_lines = wrap_text(field_value, font_medium, value_max_width)

                if value_lines:
                    draw.text((10 + field_name_width + 5, y_position), value_lines[0], font=font_medium, fill='black')

                    for i in range(1, len(value_lines)):
                        if y_position + i * 15 > max_y - 50:
                            break
                        if y_position + i * 15 >= (5 + logo_height):
                            line_width = under_logo_area_width
                        else:
                            line_width = left_area_width

                        sub_lines = wrap_text(value_lines[i], font_medium, line_width)
                        for sub_line in sub_lines:
                            if y_position + i * 15 > max_y - 50:
                                break
                            draw.text((10, y_position + i * 15), sub_line, font=font_medium, fill='black')
                            i += 1

                    y_position += len(value_lines) * 15

                y_position += 8

            if y_position < 350:
                code_article_parts = []
                if code:
                    code_article_parts.append(f"Код: {code}")
                if article:
                    code_article_parts.append(f"Артикул: {article}")

                if code_article_parts:
                    code_article_text = "\n".join(code_article_parts)
                    draw.text((250, 350), code_article_text, font=font_medium, fill='black')

            if barcode_value and y_position < 400:
                try:
                    barcode_img = self.generate_barcode(barcode_value)
                    if barcode_img:
                        logger.info(f"Barcode generated: {barcode_value}")
                        barcode_img = barcode_img.resize((200, 80))
                        img.paste(barcode_img, (250, 390))
                    else:
                        logger.warning(f"Barcode generation failed for: {barcode_value}")
                except Exception as e:
                    logger.error(f"Barcode insertion error: {e}")
            else:
                logger.warning(f"Barcode conditions not met: barcode_value={barcode_value}, y_position={y_position}")

            if certification and y_position < 420:
                cert_text = certification
                cert_lines = wrap_text(cert_text, font_medium, 250)

                for i, line in enumerate(cert_lines):
                    if 400 + i * 15 > max_y:
                        break
                    draw.text((10, 430 + i * 15), line, font=font_medium, fill='black')

                if certification_type and 380 < max_y:
                    if certification_type.lower() == 'рст':
                        cert_icon_path = "img/certificates/рст.png"
                    elif certification_type.lower() == 'eac':
                        cert_icon_path = "img/certificates/eac.png"
                    else:
                        cert_icon_path = None

                    if cert_icon_path and os.path.exists(cert_icon_path):
                        try:
                            cert_icon = self.resource_manager.get_image(cert_icon_path)
                            if cert_icon:
                                cert_icon = cert_icon.resize((80, 80))
                                img.paste(cert_icon, (10, 360), cert_icon)
                        except Exception as e:
                            logger.error(f"Certification icon load error: {e}")

            os.makedirs(os.path.dirname(output_path), exist_ok=True)
            img.save(output_path)
            logger.info(f"Label saved: {output_path}")

            if output_pdf_path:
                self.create_pdf_from_image(output_path, output_pdf_path)

            return True

        except Exception as e:
            logger.error(f"Label generation error: {e}")
            return False

    def create_pdf_from_image(self, image_path: str, pdf_path: str) -> bool:
        try:
            c = canvas.Canvas(pdf_path, pagesize=(40 * mm, 40 * mm))
            img_reader = ImageReader(image_path)
            c.drawImage(img_reader, 0, 0, 40 * mm, 40 * mm)
            c.save()
            logger.info(f"PDF created: {pdf_path}")
            return True
        except Exception as e:
            logger.error(f"PDF creation error: {e}")
            return False


class MarkGenerator:
    def __init__(self):
        self.config = Config
        self.resource_manager = ResourceManager

    def generate(self, data: Dict, output_path: str, output_pdf_path: Optional[str] = None) -> bool:
        try:
            img = Image.new('RGB', self.config.LABEL_SIZE_PX, color='white')
            draw = ImageDraw.Draw(img)

            font_title = self.resource_manager.get_font('bold', 20)
            font_regular = self.resource_manager.get_font('regular', 20)

            mark_image_path = None
            mark_image_dir = "img/mark_images"
            extensions = ['.png', '.jpg', '.jpeg', '.bmp']

            for ext in extensions:
                test_path = os.path.join(mark_image_dir, f"mark_images{ext}")
                if os.path.exists(test_path):
                    mark_image_path = test_path
                    break

            mark_img = None
            mark_height = 0

            if mark_image_path:
                try:
                    mark_img = self.resource_manager.get_image(mark_image_path)
                    if mark_img:
                        width_percent = 270 / float(mark_img.size[0])
                        new_height = int(float(mark_img.size[1]) * float(width_percent))
                        mark_img = mark_img.resize((270, new_height))
                        mark_height = new_height
                        img.paste(mark_img, (5, 3), mark_img)
                except Exception as e:
                    logger.error(f"Mark image load error: {e}")

            logo_name = data.get('лого', '').strip().lower()
            logo_img = None

            if logo_name:
                logo_dir = "img/logos"
                extensions = ['.png', '.jpg', '.jpeg', '.bmp']
                for ext in extensions:
                    logo_path = os.path.join(logo_dir, f"{logo_name}{ext}")
                    if os.path.exists(logo_path):
                        try:
                            logo_img = self.resource_manager.get_image(logo_path)
                            if logo_img:
                                max_logo_size = 200
                                logo_img.thumbnail((max_logo_size, max_logo_size))
                                logo_width, logo_height = logo_img.size
                                x_position = 10
                                y_logo_position = mark_height + 20
                                img.paste(logo_img, (x_position, y_logo_position), logo_img)
                                break
                        except Exception as e:
                            logger.error(f"Logo load error {logo_name}: {e}")
                            continue

            article = data.get('артикул', '')
            code = data.get('код', '')

            y_position = mark_height + 120 if mark_img else 120

            if code:
                draw.text((50, y_position), f"Артикул: {code}", font=font_title, fill='black')
            y_position += 40

            draw.text((50, y_position), "Количество:", font=font_regular, fill='black')
            y_position += 40

            draw.text((50, y_position), "Вес нетто:", font=font_regular, fill='black')
            kg_text = "кг"
            try:
                bbox = font_regular.getbbox(kg_text)
                kg_width = bbox[2] - bbox[0]
            except AttributeError:
                kg_width = font_regular.getsize(kg_text)[0]
            draw.text((472 - 50 - kg_width, y_position), kg_text, font=font_regular, fill='black')
            y_position += 40

            draw.text((50, y_position), "Вес брутто:", font=font_regular, fill='black')
            draw.text((472 - 50 - kg_width, y_position), kg_text, font=font_regular, fill='black')

            os.makedirs(os.path.dirname(output_path), exist_ok=True)
            img.save(output_path)
            logger.info(f"Mark saved: {output_path}")

            if output_pdf_path:
                self.create_pdf_from_image(output_path, output_pdf_path)

            return True

        except Exception as e:
            logger.error(f"Mark generation error: {e}")
            return False

    def create_pdf_from_image(self, image_path: str, pdf_path: str) -> bool:
        try:
            c = canvas.Canvas(pdf_path, pagesize=(40 * mm, 40 * mm))
            img_reader = ImageReader(image_path)
            c.drawImage(img_reader, 0, 0, 40 * mm, 40 * mm)
            c.save()
            logger.info(f"PDF created: {pdf_path}")
            return True
        except Exception as e:
            logger.error(f"PDF creation error: {e}")
            return False


class Application:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Генератор этикеток и марок")
        self.root.geometry("600x400")
        self.root.resizable(True, True)

        os.makedirs('input', exist_ok=True)
        for dir_path in Config.IMAGE_DIRS:
            os.makedirs(dir_path, exist_ok=True)

        self.setup_ui()

    def setup_ui(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        title_label = ttk.Label(main_frame, text="Генератор этикеток и марок", font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))

        instruction_text = """
        Инструкция:
        1. Поместите Excel-файлы в папку 'input' или выберите файлы кнопкой ниже
        2. Нажмите кнопку 'Обработать файлы'
        3. Результаты появятся в папке 'output'

        Структура папки output:
        - labels_png/ - этикетки в формате PNG
        - labels_pdf/ - этикетки в формате PDF  
        - marks_png/ - маркировки в формате PNG
        - marks_pdf/ - маркировки в формате PDF
        """

        instruction_label = ttk.Label(main_frame, text=instruction_text, justify=tk.LEFT)
        instruction_label.grid(row=1, column=0, columnspan=2, pady=(0, 20))

        select_btn = ttk.Button(main_frame, text="Выбрать файлы", command=self.select_files)
        select_btn.grid(row=2, column=0, pady=5, padx=5, sticky=tk.EW)

        process_btn = ttk.Button(main_frame, text="Обработать файлы", command=self.process_files)
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

    def select_files(self):
        files = filedialog.askopenfilenames(
            title="Выберите Excel файлы",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )

        if files:
            for file_path in files:
                try:
                    shutil.copy2(file_path, 'input')
                    logger.info(f"File copied: {os.path.basename(file_path)}")
                except Exception as e:
                    logger.error(f"File copy error {file_path}: {e}")

            self.status_var.set(f"Copied {len(files)} files to input")

    def process_files(self):
        for dir_path in Config.OUTPUT_DIRS:
            if os.path.exists(dir_path):
                shutil.rmtree(dir_path, onerror=handle_remove_readonly)
            os.makedirs(dir_path, exist_ok=True)

        thread = threading.Thread(target=self.process_files_thread)
        thread.daemon = True
        thread.start()

    def process_files_thread(self):
        try:
            self.progress.start()
            self.status_var.set("Processing files...")

            label_generator = LabelGenerator()
            mark_generator = MarkGenerator()

            input_dir = 'input'
            excel_files = []
            for file in os.listdir(input_dir):
                if file.endswith(('.xlsx', '.xls')):
                    excel_files.append(os.path.join(input_dir, file))

            if not excel_files:
                logger.warning("No Excel files found in input folder")
                self.status_var.set("No Excel files found in input folder")
                return

            for file_path in excel_files:
                logger.info(f"Processing file: {file_path}")
                df = read_excel(file_path)

                if df is None or df.empty:
                    logger.error(f"Failed to read data from {file_path}")
                    continue

                for idx, row in df.iterrows():
                    if not validate_data(row):
                        logger.warning(f"Skipping row {idx} due to data errors")
                        continue

                    article = str(row.get('артикул', '')).strip()
                    code = str(row.get('код', '')).strip()

                    article_clean = re.sub(r'[\\/*?:"<>|]', "_", article)
                    code_clean = re.sub(r'[\\/*?:"<>|]', "_", code)

                    filename = f"label_{article_clean}_{code_clean}"
                    mark_filename = f"mark_{article_clean}_{code_clean}"

                    logger.info(f"Processing row {idx}: {row.get('наименование', '')}")

                    label_png_path = f'output/labels_png/{filename}.png'
                    label_pdf_path = f'output/labels_pdf/{filename}.pdf'
                    label_generator.generate(row, label_png_path, label_pdf_path)

                    mark_path = f'output/marks_png/{mark_filename}.png'
                    mark_pdf_path = f'output/marks_pdf/{mark_filename}.pdf'
                    mark_generator.generate(row, mark_path, mark_pdf_path)

            logger.info("Generation completed!")
            self.status_var.set("Processing completed successfully!")

            messagebox.showinfo("Success", "File processing completed!\n\nResults are in the output folder.")

        except Exception as e:
            logger.error(f"Processing error: {e}")
            self.status_var.set(f"Error: {e}")
            messagebox.showerror("Error", f"Processing error: {e}")
        finally:
            self.progress.stop()

    def run(self):
        self.root.mainloop()


def main():
    if len(sys.argv) > 1 and sys.argv[1] == '--console':
        console_main()
    else:
        app = Application()
        app.run()


def console_main():
    os.makedirs('input', exist_ok=True)
    for dir_path in Config.IMAGE_DIRS:
        os.makedirs(dir_path, exist_ok=True)

    for dir_path in Config.OUTPUT_DIRS:
        if os.path.exists(dir_path):
            shutil.rmtree(dir_path, onerror=handle_remove_readonly)
        os.makedirs(dir_path, exist_ok=True)

    label_generator = LabelGenerator()
    mark_generator = MarkGenerator()

    input_dir = 'input'
    excel_files = []
    for file in os.listdir(input_dir):
        if file.endswith(('.xlsx', '.xls')):
            excel_files.append(os.path.join(input_dir, file))

    if not excel_files:
        logger.warning("No Excel files found in input folder")
        return

    for file_path in excel_files:
        logger.info(f"Processing file: {file_path}")
        df = read_excel(file_path)

        if df is None or df.empty:
            logger.error(f"Failed to read data from {file_path}")
            continue

        for idx, row in df.iterrows():
            if not validate_data(row):
                logger.warning(f"Skipping row {idx} due to data errors")
                continue

            article = str(row.get('артикул', '')).strip()
            code = str(row.get('код', '')).strip()

            article_clean = re.sub(r'[\\/*?:"<>|]', "_", article)
            code_clean = re.sub(r'[\\/*?:"<>|]', "_", code)

            filename = f"label_{article_clean}_{code_clean}"
            mark_filename = f"mark_{article_clean}_{code_clean}"

            logger.info(f"Processing row {idx}: {row.get('наименование', '')}")

            label_png_path = f'output/labels_png/{filename}.png'
            label_pdf_path = f'output/labels_pdf/{filename}.pdf'
            label_generator.generate(row, label_png_path, label_pdf_path)

            mark_path = f'output/marks_png/{mark_filename}.png'
            mark_pdf_path = f'output/marks_pdf/{mark_filename}.pdf'
            mark_generator.generate(row, mark_path, mark_pdf_path)

    logger.info("Generation completed!")
    print("Processing completed. Results in output folder.")


if __name__ == "__main__":
    main()