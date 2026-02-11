"""
Bitrix24 Bot для генерации этикеток и марок
Принимает архив формата: data.xlsx + img/ (с подпапками logos, certificates, mark_images)
Возвращает архив с результатами (labels/ и marks/)

Архитектура:
1. Bitrix24 отправляет события (ONIMBOTMESSAGEADD) на наш /webhook
2. Мы обрабатываем события и отвечаем через REST API (imbot.message.add)
"""

import os
import sys
import tempfile
import shutil
import zipfile
import json
import logging
from typing import Optional, Dict, List
from datetime import datetime
from pathlib import Path

import requests
from flask import Flask, request, jsonify

from main import CombinedGenerator

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

BITRIX_WEBHOOK = os.getenv("BITRIX_WEBHOOK", "https://poryadok.bitrix24.ru/rest/159096/1g7d9dxu9rd1kpxc/")
PUBLIC_URL = os.getenv("PUBLIC_URL", "https://your-server.com")
BOT_CODE = "labels_generator_bot"
BOT_NAME = "Генератор Этикеток"

app = Flask(__name__)

BOT_ID = None


class BitrixBot:
    """Класс для работы с Bitrix24 API"""

    def __init__(self, webhook_url: str):
        self.webhook_url = webhook_url.rstrip('/')
        self.bot_id = None

    def call_method(self, method: str, params: Dict = None) -> Dict:
        url = f"{self.webhook_url}/{method}"
        try:
            response = requests.post(url, json=params or {}, timeout=30)
            response.raise_for_status()
            return response.json()
        except Exception as e:
            logger.error(f"Ошибка вызова метода {method}: {e}")
            return {"error": str(e)}

    def register_bot(self, bot_code: str, bot_name: str, event_handler_url: str) -> Optional[int]:
        """
        Регистрация бота в Bitrix24
        Возвращает BOT_ID или None в случае ошибки
        """
        try:
            params = {
                "CODE": bot_code,
                "TYPE": "B",  # Bot
                "EVENT_MESSAGE_ADD": event_handler_url,
                "EVENT_WELCOME_MESSAGE": event_handler_url,
                "EVENT_BOT_DELETE": event_handler_url,
                "PROPERTIES": {
                    "NAME": bot_name,
                    "WORK_POSITION": "Генератор этикеток и марок из Excel"
                }
            }

            logger.info(f"Регистрируем бота: {bot_code}")
            result = self.call_method("imbot.register", params)

            if "result" in result:
                bot_id = result["result"]
                logger.info(f"✅ Бот зарегистрирован! BOT_ID: {bot_id}")
                self.bot_id = bot_id
                return bot_id
            else:
                logger.error(f"❌ Ошибка регистрации: {result}")
                return None

        except Exception as e:
            logger.error(f"❌ Исключение при регистрации: {e}")
            return None

    def bind_event(self, event: str, handler_url: str) -> bool:
        """
        Подписка на событие
        event: ONIMBOTMESSAGEADD, ONIMBOTJOINCHAT, etc.
        """
        try:
            params = {
                "event": event,
                "handler": handler_url
            }

            logger.info(f"Подписываемся на событие: {event}")
            result = self.call_method("event.bind", params)

            if "result" in result:
                logger.info(f"✅ Подписка создана: {event}")
                return True
            else:
                logger.warning(f"⚠️ Не удалось подписаться: {result}")
                return False

        except Exception as e:
            logger.error(f"❌ Ошибка подписки на событие: {e}")
            return False

    def get_events(self) -> List[Dict]:
        """Получить список подписок на события"""
        try:
            result = self.call_method("event.get")
            if "result" in result:
                return result["result"]
            return []
        except Exception as e:
            logger.error(f"Ошибка получения списка событий: {e}")
            return []

    def send_message(self, dialog_id: str, message: str) -> bool:
        """
        Отправка сообщения в чат
        dialog_id: ID диалога (обычно совпадает с user_id для личных сообщений)
        """
        try:
            params = {
                "DIALOG_ID": dialog_id,
                "MESSAGE": message
            }

            result = self.call_method("imbot.message.add", params)

            if "result" in result:
                logger.info(f"✅ Сообщение отправлено в диалог {dialog_id}")
                return True
            else:
                logger.error(f"❌ Ошибка отправки сообщения: {result}")
                return False

        except Exception as e:
            logger.error(f"❌ Исключение при отправке сообщения: {e}")
            return False

    def send_file(self, dialog_id: str, file_path: str, message: str = "") -> bool:
        """
        Отправка файла в чат
        Bitrix24 поддерживает загрузку файлов через disk.folder.uploadfile
        """
        try:
            with open(file_path, 'rb') as f:
                files = {
                    'fileContent': (os.path.basename(file_path), f, 'application/octet-stream')
                }

                upload_params = {
                    "id": "shared_folder",  # ID папки бота
                    "data": {"NAME": os.path.basename(file_path)},
                }

                params = {
                    "DIALOG_ID": dialog_id,
                    "MESSAGE": message,
                }

                result = self.call_method("imbot.message.add", params)

                if "result" in result:
                    logger.info(f"✅ Файл отправлен в диалог {dialog_id}")
                    return True
                else:
                    logger.error(f"❌ Ошибка отправки файла: {result}")
                    return False

        except Exception as e:
            logger.error(f"❌ Исключение при отправке файла: {e}")
            return False

    def download_file(self, file_id: str, save_path: str) -> bool:
        """Скачивание файла из Bitrix24"""
        try:
            params = {"ID": file_id}
            result = self.call_method("disk.file.get", params)

            if "result" not in result:
                logger.error(f"Не удалось получить информацию о файле: {result}")
                return False

            download_url = result["result"].get("DOWNLOAD_URL")
            if not download_url:
                logger.error("URL для скачивания не найден")
                return False

            response = requests.get(download_url, timeout=60)
            response.raise_for_status()

            with open(save_path, 'wb') as f:
                f.write(response.content)

            logger.info(f"Файл скачан: {save_path}")
            return True

        except Exception as e:
            logger.error(f"Ошибка скачивания файла: {e}")
            return False


class LabelsProcessor:
    """Класс для обработки архивов и генерации этикеток"""

    def __init__(self):
        self.generator = CombinedGenerator()
        self.temp_dir = None

    def process_archive(self, archive_path: str) -> Optional[str]:
        """
        Обработка архива в формате: data.xlsx + img/ (с подпапками)
        Структура архива:
          - data.xlsx (обязательно)
          - img/logos/ (опционально)
          - img/certificates/ (опционально)
          - img/mark_images/ (опционально)

        Возвращает путь к результирующему архиву или None в случае ошибки
        """
        try:
            self.temp_dir = tempfile.mkdtemp(prefix="labels_")
            logger.info(f"Создана временная директория: {self.temp_dir}")

            input_dir = os.path.join(self.temp_dir, "input")
            os.makedirs(input_dir, exist_ok=True)

            with zipfile.ZipFile(archive_path, 'r') as zip_ref:
                zip_ref.extractall(input_dir)
            logger.info(f"Архив распакован в: {input_dir}")

            excel_files = []
            img_dir_path = None

            for root, dirs, files in os.walk(input_dir):
                for file in files:
                    if file.endswith(('.xlsx', '.xls')):
                        excel_files.append(os.path.join(root, file))

                if 'img' in dirs:
                    img_dir_path = os.path.join(root, 'img')
                    logger.info(f"Найдена папка img: {img_dir_path}")

            logger.info(f"Найдено Excel файлов: {len(excel_files)}")

            if not excel_files:
                logger.error("Excel файлы не найдены в архиве")
                return None

            if img_dir_path and os.path.exists(img_dir_path):
                self._copy_img_folder(img_dir_path)
            else:
                logger.warning("Папка img не найдена в архиве, используем существующие изображения")

            output_dir = os.path.join(self.temp_dir, "output")
            os.makedirs(output_dir, exist_ok=True)

            for excel_file in excel_files:
                logger.info(f"Обработка файла: {excel_file}")
                success = self.generator.process_excel_file(excel_file, output_dir)

                if success:
                    logger.info(f"Файл успешно обработан: {excel_file}")
                else:
                    logger.warning(f"Ошибка при обработке файла: {excel_file}")

            result_archive = os.path.join(self.temp_dir, "result.zip")
            self._create_result_archive(output_dir, result_archive)

            logger.info(f"Результирующий архив создан: {result_archive}")
            return result_archive

        except Exception as e:
            logger.error(f"Ошибка обработки архива: {e}", exc_info=True)
            return None

    def _copy_img_folder(self, img_dir_path: str):
        """Копирование папки img целиком в LabelsMarksGenerator/img/"""
        target_img_dir = "LabelsMarksGenerator/img"

        # Очищаем старые изображения (опционально)
        # for subdir in ['logos', 'certificates', 'mark_images']:
        #     subdir_path = os.path.join(target_img_dir, subdir)
        #     if os.path.exists(subdir_path):
        #         shutil.rmtree(subdir_path)

        # Копируем содержимое папки img
        for item in os.listdir(img_dir_path):
            source_path = os.path.join(img_dir_path, item)
            target_path = os.path.join(target_img_dir, item)

            if os.path.isdir(source_path):
                # Копируем подпапку целиком (logos, certificates, mark_images)
                if os.path.exists(target_path):
                    shutil.rmtree(target_path)
                shutil.copytree(source_path, target_path)
                logger.info(f"Скопирована папка: {item}/")
            else:
                # Копируем отдельный файл
                shutil.copy2(source_path, target_path)
                logger.info(f"Скопирован файл: {item}")

    def _create_result_archive(self, output_dir: str, archive_path: str):
        """Создание архива с результатами"""
        with zipfile.ZipFile(archive_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, dirs, files in os.walk(output_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, output_dir)
                    zipf.write(file_path, arcname)

        logger.info(f"Архив создан: {archive_path}")

    def cleanup(self):
        """Очистка временных файлов"""
        if self.temp_dir and os.path.exists(self.temp_dir):
            try:
                shutil.rmtree(self.temp_dir)
                logger.info(f"Временная директория удалена: {self.temp_dir}")
            except Exception as e:
                logger.error(f"Ошибка при удалении временной директории: {e}")


bot = BitrixBot(BITRIX_WEBHOOK)
processor = LabelsProcessor()


@app.route('/webhook', methods=['POST'])
def webhook():
    """Обработчик вебхуков от Bitrix24"""
    try:
        data = request.json
        logger.info(f"Получен вебхук: {json.dumps(data, ensure_ascii=False)}")

        event = data.get('event')

        if event == 'ONIMBOTMESSAGEADD':
            message_data = data.get('data', {})
            dialog_id = message_data.get('PARAMS', {}).get('DIALOG_ID')
            message_text = message_data.get('PARAMS', {}).get('MESSAGE', '')
            files = message_data.get('FILES', [])

            if message_text.startswith('/'):
                handle_command(dialog_id, message_text)

            elif files:
                handle_files(dialog_id, files)

            else:
                bot.send_message(dialog_id,
                    "Привет! Я бот для генерации этикеток и марок.\n\n"
                    "Отправь мне архив (ZIP) с Excel файлом и картинками, "
                    "и я создам для тебя этикетки и марки.\n\n"
                    "Команды:\n"
                    "/help - помощь\n"
                    "/start - начать работу"
                )

        return jsonify({"status": "ok"})

    except Exception as e:
        logger.error(f"Ошибка обработки вебхука: {e}", exc_info=True)
        return jsonify({"status": "error", "message": str(e)}), 500


def handle_command(dialog_id: str, command: str):
    """Обработка команд бота"""
    command = command.lower().strip()

    if command == '/start':
        bot.send_message(dialog_id,
            "Привет! Я готов к работе.\n\n"
            "Отправь мне ZIP архив со структурой:\n"
            "📂 archive.zip\n"
            "   ├── data.xlsx (обязательно)\n"
            "   └── img/ (опционально)\n"
            "       ├── logos/\n"
            "       ├── certificates/\n"
            "       └── mark_images/\n\n"
            "Я обработаю данные и отправлю архив с этикетками и марками."
        )

    elif command == '/help':
        bot.send_message(dialog_id,
            "📋 Инструкция:\n\n"
            "1. Создай ZIP архив со структурой:\n"
            "   📂 archive.zip\n"
            "      ├── data.xlsx\n"
            "      └── img/\n"
            "          ├── logos/ (логотипы .png/.jpg)\n"
            "          ├── certificates/ (еас.png, рст.png)\n"
            "          └── mark_images/ (mark_images.png)\n\n"
            "2. Отправь архив сюда\n\n"
            "3. Получишь result.zip с:\n"
            "   - labels/ (этикетки PDF)\n"
            "   - marks/ (марки PDF)\n\n"
            "Команды:\n"
            "/start - начать\n"
            "/help - справка"
        )

    else:
        bot.send_message(dialog_id, "Неизвестная команда. Используй /help для справки.")


def handle_files(dialog_id: str, files: List[Dict]):
    """Обработка загруженных файлов"""
    try:
        bot.send_message(dialog_id, "📦 Получен архив, начинаю обработку...")

        zip_file = None
        for file_info in files:
            if file_info.get('name', '').lower().endswith('.zip'):
                zip_file = file_info
                break

        if not zip_file:
            bot.send_message(dialog_id,
                "❌ Ошибка: отправь ZIP архив с данными.\n"
                "Используй /help для справки."
            )
            return

        file_id = zip_file.get('id')
        temp_archive = os.path.join(tempfile.gettempdir(), f"input_{file_id}.zip")

        if not bot.download_file(file_id, temp_archive):
            bot.send_message(dialog_id, "❌ Ошибка скачивания архива")
            return

        bot.send_message(dialog_id, "⚙️ Обработка данных...")

        result_archive = processor.process_archive(temp_archive)

        if result_archive:
            bot.send_message(dialog_id, "✅ Обработка завершена! Отправляю результаты...")

            success = bot.send_file(dialog_id, result_archive,
                "🎉 Готово! Этикетки и марки в архиве.")

            if success:
                bot.send_message(dialog_id,
                    "Архив содержит:\n"
                    "- labels/ - этикетки в PDF\n"
                    "- marks/ - марки в PDF"
                )
            else:
                bot.send_message(dialog_id, "❌ Ошибка отправки результатов")

        else:
            bot.send_message(dialog_id,
                "❌ Ошибка обработки архива.\n"
                "Проверь структуру архива и наличие Excel файла."
            )

        processor.cleanup()
        if os.path.exists(temp_archive):
            os.remove(temp_archive)

    except Exception as e:
        logger.error(f"Ошибка обработки файлов: {e}", exc_info=True)
        bot.send_message(dialog_id, f"❌ Ошибка: {str(e)}")


@app.route('/health', methods=['GET'])
def health():
    """Проверка работоспособности сервера"""
    return jsonify({"status": "ok", "timestamp": datetime.now().isoformat()})


if __name__ == '__main__':
    # bot.register_bot()

    logger.info("Запуск Bitrix24 бота...")
    app.run(host='0.0.0.0', port=5000, debug=True)