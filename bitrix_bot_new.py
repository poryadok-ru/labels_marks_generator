"""
Bitrix24 Bot для генерации этикеток и марок
Основано на официальной документации: https://apidocs.bitrix24.com/api-reference/chat-bots/

Архитектура:
1. Регистрация бота через imbot.register
2. Bitrix24 отправляет события на /webhook (ONIMBOTMESSAGEADD)
3. Обрабатываем события и отвечаем через imbot.message.add
4. Файлы загружаем через disk.folder.uploadfile
"""

import os
import tempfile
import shutil
import zipfile
import json
import logging
from typing import Optional, Dict

import requests
from flask import Flask, request, jsonify

from main import CombinedGenerator

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Конфигурация из .env
BITRIX_WEBHOOK = os.getenv("BITRIX_WEBHOOK", "https://poryadok.bitrix24.ru/rest/159096/1g7d9dxu9rd1kpxc/")
PUBLIC_URL = os.getenv("PUBLIC_URL", "https://your-ngrok-url.ngrok.io")  # Укажи URL от ngrok
BOT_CODE = "labels_generator_bot"
BOT_NAME = "Генератор Этикеток"

app = Flask(__name__)
BOT_ID = None


class BitrixAPI:
    """Класс для работы с Bitrix24 REST API"""

    def __init__(self, webhook_url: str):
        self.webhook_url = webhook_url.rstrip('/')

    def call(self, method: str, params: Dict = None) -> Dict:
        """Вызов REST API метода"""
        url = f"{self.webhook_url}/{method}"
        try:
            logger.info(f"API call: {method}")
            response = requests.post(url, json=params or {}, timeout=30)
            response.raise_for_status()
            result = response.json()
            logger.debug(f"Response: {result}")
            return result
        except Exception as e:
            logger.error(f"API error {method}: {e}")
            return {"error": str(e)}

    def register_bot(self, code: str, name: str, handler_url: str) -> Optional[int]:
        """
        Регистрация бота
        Документация: https://apidocs.bitrix24.com/api-reference/chat-bots/imbot-register.html
        """
        params = {
            "CODE": code,
            "TYPE": "B",  # B = обычный бот
            "EVENT_MESSAGE_ADD": handler_url,
            "EVENT_WELCOME_MESSAGE": handler_url,
            "EVENT_BOT_DELETE": handler_url,
            "PROPERTIES": {
                "NAME": name,
                "COLOR": "AZURE",
                "WORK_POSITION": "Генератор этикеток и марок"
            }
        }

        result = self.call("imbot.register", params)

        if "result" in result:
            bot_id = result["result"]
            logger.info(f"✅ Бот зарегистрирован! BOT_ID: {bot_id}")
            return bot_id
        else:
            logger.error(f"❌ Ошибка регистрации: {result}")
            return None

    def send_message(self, dialog_id: str, message: str) -> bool:
        """
        Отправка сообщения
        Документация: https://apidocs.bitrix24.com/api-reference/chat-bots/messages/imbot-message-add.html
        """
        params = {
            "DIALOG_ID": dialog_id,
            "MESSAGE": message
        }

        result = self.call("imbot.message.add", params)

        if "result" in result:
            logger.info(f"✅ Сообщение отправлено в {dialog_id}")
            return True
        else:
            logger.error(f"❌ Ошибка отправки: {result}")
            return False

    def upload_file(self, folder_id: str, file_path: str) -> Optional[int]:
        """
        Загрузка файла на диск Bitrix24
        Документация: https://apidocs.bitrix24.com/api-reference/disk/folder/disk-folder-upload-file.html
        """
        try:
            url = f"{self.webhook_url}/disk.folder.uploadfile"

            with open(file_path, 'rb') as f:
                files = {
                    'fileContent': (os.path.basename(file_path), f, 'application/zip')
                }
                data = {
                    'id': folder_id,
                    'data[NAME]': os.path.basename(file_path)
                }

                response = requests.post(url, files=files, data=data, timeout=60)
                response.raise_for_status()
                result = response.json()

                if "result" in result:
                    file_id = result["result"]["ID"]
                    logger.info(f"✅ Файл загружен: ID={file_id}")
                    return file_id
                else:
                    logger.error(f"❌ Ошибка загрузки файла: {result}")
                    return None

        except Exception as e:
            logger.error(f"❌ Исключение при загрузке файла: {e}")
            return None

    def send_file_message(self, dialog_id: str, file_id: int, message: str = "") -> bool:
        """Отправка сообщения с прикрепленным файлом"""
        params = {
            "DIALOG_ID": dialog_id,
            "MESSAGE": message,
            "FILE_ID": [file_id]  # Массив ID файлов
        }

        result = self.call("imbot.message.add", params)

        if "result" in result:
            logger.info(f"✅ Файл отправлен в {dialog_id}")
            return True
        else:
            logger.error(f"❌ Ошибка отправки файла: {result}")
            return False

    def get_bot_storage_folder(self) -> Optional[str]:
        """Получить ID папки хранилища бота"""
        # Используем общее хранилище или создаем папку
        result = self.call("disk.storage.getlist")

        if "result" in result and len(result["result"]) > 0:
            # Берем первое доступное хранилище
            storage_id = result["result"][0]["ID"]
            logger.info(f"Storage ID: {storage_id}")
            return storage_id

        return None


class LabelsProcessor:
    """Обработчик архивов для генерации этикеток"""

    def __init__(self):
        self.generator = CombinedGenerator()
        self.temp_dir = None

    def process_archive(self, archive_path: str) -> Optional[str]:
        """
        Обработка архива: data.xlsx + img/
        Возвращает путь к result.zip или None
        """
        try:
            self.temp_dir = tempfile.mkdtemp(prefix="labels_")
            logger.info(f"Temp dir: {self.temp_dir}")

            # Распаковываем
            input_dir = os.path.join(self.temp_dir, "input")
            os.makedirs(input_dir)

            with zipfile.ZipFile(archive_path, 'r') as zip_ref:
                zip_ref.extractall(input_dir)

            # Ищем Excel и img
            excel_files = []
            img_dir_path = None

            for root, dirs, files in os.walk(input_dir):
                for file in files:
                    if file.endswith(('.xlsx', '.xls')):
                        excel_files.append(os.path.join(root, file))

                if 'img' in dirs:
                    img_dir_path = os.path.join(root, 'img')

            if not excel_files:
                logger.error("Excel не найден")
                return None

            # Копируем img если есть
            if img_dir_path:
                self._copy_img(img_dir_path)

            # Генерируем
            output_dir = os.path.join(self.temp_dir, "output")
            os.makedirs(output_dir)

            for excel_file in excel_files:
                logger.info(f"Processing: {excel_file}")
                self.generator.process_excel_file(excel_file, output_dir)

            # Создаем result.zip
            result_archive = os.path.join(self.temp_dir, "result.zip")
            with zipfile.ZipFile(result_archive, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for root, dirs, files in os.walk(output_dir):
                    for file in files:
                        file_path = os.path.join(root, file)
                        arcname = os.path.relpath(file_path, output_dir)
                        zipf.write(file_path, arcname)

            logger.info(f"✅ Результат: {result_archive}")
            return result_archive

        except Exception as e:
            logger.error(f"❌ Ошибка обработки: {e}", exc_info=True)
            return None

    def _copy_img(self, img_dir_path: str):
        """Копирование папки img в LabelsMarksGenerator/img/"""
        target_img_dir = "LabelsMarksGenerator/img"

        for item in os.listdir(img_dir_path):
            source_path = os.path.join(img_dir_path, item)
            target_path = os.path.join(target_img_dir, item)

            if os.path.isdir(source_path):
                if os.path.exists(target_path):
                    shutil.rmtree(target_path)
                shutil.copytree(source_path, target_path)
                logger.info(f"Copied folder: {item}/")
            else:
                shutil.copy2(source_path, target_path)
                logger.info(f"Copied file: {item}")

    def cleanup(self):
        """Очистка временных файлов"""
        if self.temp_dir and os.path.exists(self.temp_dir):
            shutil.rmtree(self.temp_dir)
            logger.info("Temp cleaned")


# Инициализация
api = BitrixAPI(BITRIX_WEBHOOK)
processor = LabelsProcessor()


@app.route('/webhook', methods=['POST'])
def webhook():
    """
    Обработчик событий от Bitrix24
    Документация: https://apidocs.bitrix24.com/api-reference/chat-bots/messages/events/on-imbot-message-add.html
    """
    try:
        data = request.json
        logger.info(f"📨 Event: {json.dumps(data, ensure_ascii=False)[:500]}")

        event = data.get('event')

        # Новое сообщение боту
        if event == 'ONIMBOTMESSAGEADD':
            message_data = data.get('data', {}).get('PARAMS', {})
            dialog_id = message_data.get('DIALOG_ID')
            message_text = message_data.get('MESSAGE', '').strip()
            files = data.get('data', {}).get('FILES', [])

            logger.info(f"Dialog: {dialog_id}, Message: {message_text}, Files: {len(files)}")

            # Команды
            if message_text.startswith('/'):
                handle_command(dialog_id, message_text)

            # Файлы (архив)
            elif files:
                handle_files(dialog_id, files)

            # Обычное сообщение
            else:
                api.send_message(dialog_id,
                    "👋 Привет! Отправь мне ZIP архив:\n\n"
                    "📂 archive.zip\n"
                    "   ├── data.xlsx\n"
                    "   └── img/ (опционально)\n\n"
                    "Команды: /start, /help"
                )

        # Приветствие
        elif event == 'ONIMBOTJOINCHAT':
            dialog_id = data.get('data', {}).get('PARAMS', {}).get('DIALOG_ID')
            api.send_message(dialog_id, "👋 Привет! Я помогу создать этикетки. Отправь /help для справки.")

        return jsonify({"status": "ok"})

    except Exception as e:
        logger.error(f"❌ Webhook error: {e}", exc_info=True)
        return jsonify({"status": "error", "message": str(e)}), 500


def handle_command(dialog_id: str, command: str):
    """Обработка команд"""
    command = command.lower().strip()

    if command == '/start':
        api.send_message(dialog_id,
            "🚀 Я готов!\n\n"
            "Отправь ZIP архив:\n"
            "📂 archive.zip\n"
            "   ├── data.xlsx (обязательно)\n"
            "   └── img/ (опционально)\n"
            "       ├── logos/\n"
            "       ├── certificates/\n"
            "       └── mark_images/\n\n"
            "Получишь result.zip с этикетками и марками."
        )

    elif command == '/help':
        api.send_message(dialog_id,
            "📋 Справка:\n\n"
            "1. Создай ZIP:\n"
            "   - data.xlsx\n"
            "   - img/logos/ (.png/.jpg)\n"
            "   - img/certificates/ (еас.png, рст.png)\n"
            "   - img/mark_images/ (mark_images.png)\n\n"
            "2. Отправь архив сюда\n\n"
            "3. Получишь result.zip с PDF\n\n"
            "/start - начать"
        )

    else:
        api.send_message(dialog_id, "❓ Неизвестная команда. Используй /help")


def handle_files(dialog_id: str, files: list):
    """Обработка загруженных файлов"""
    try:
        # Ищем ZIP файл
        zip_file = None
        for file in files:
            if file.get('name', '').lower().endswith('.zip'):
                zip_file = file
                break

        if not zip_file:
            api.send_message(dialog_id, "❌ Нужен ZIP архив. Используй /help")
            return

        api.send_message(dialog_id, "📦 Получен архив, обрабатываю...")

        # Скачиваем файл через Bitrix24 API
        file_id = zip_file.get('id')
        file_url_id = zip_file.get('urlDownload')  # Или FILE_ID

        # Для упрощения пока просим пользователя отправить через внешний способ
        api.send_message(dialog_id,
            "⚠️ Прямая обработка файлов в разработке.\n\n"
            "Пока что:\n"
            "1. Загрузи архив через веб-интерфейс\n"
            "2. Или отправь мне ссылку на архив\n\n"
            "Скоро добавлю автоматическую обработку!"
        )

    except Exception as e:
        logger.error(f"❌ File handling error: {e}", exc_info=True)
        api.send_message(dialog_id, f"❌ Ошибка: {str(e)}")


@app.route('/health', methods=['GET'])
def health():
    """Проверка работоспособности"""
    return jsonify({"status": "ok", "bot_id": BOT_ID})


@app.route('/setup', methods=['GET'])
def setup():
    """Регистрация бота (вызвать один раз)"""
    global BOT_ID

    handler_url = f"{PUBLIC_URL}/webhook"

    bot_id = api.register_bot(BOT_CODE, BOT_NAME, handler_url)

    if bot_id:
        BOT_ID = bot_id
        return jsonify({
            "status": "success",
            "bot_id": bot_id,
            "message": f"Бот зарегистрирован! BOT_ID: {bot_id}"
        })
    else:
        return jsonify({
            "status": "error",
            "message": "Ошибка регистрации"
        }), 500


if __name__ == '__main__':
    logger.info("=" * 60)
    logger.info("🤖 Bitrix24 Bot для генерации этикеток")
    logger.info("=" * 60)
    logger.info(f"Webhook URL: {BITRIX_WEBHOOK}")
    logger.info(f"Public URL: {PUBLIC_URL}")
    logger.info("")
    logger.info("📝 Шаги для запуска:")
    logger.info("1. Запусти ngrok: ngrok http 5000")
    logger.info("2. Укажи PUBLIC_URL в .env (https://xxx.ngrok.io)")
    logger.info("3. Открой /setup для регистрации бота")
    logger.info("4. Бот готов принимать сообщения!")
    logger.info("=" * 60)

    app.run(host='0.0.0.0', port=5000, debug=True)
