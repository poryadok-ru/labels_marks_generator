"""
Bitrix24 Bot с Polling (опросом)
Не требует публичного URL - просто запусти и бот работает!

Как работает:
1. Бот периодически проверяет новые сообщения через im.dialog.messages.get
2. Обрабатывает команды и файлы
3. Отвечает через imbot.message.add

Документация:
- https://apidocs.bitrix24.com/api-reference/chat-bots/
- https://training.bitrix24.com/support/training/course/?COURSE_ID=115
"""

import os
import time
import tempfile
import shutil
import zipfile
import json
import logging
from typing import Optional, Dict, List
from datetime import datetime

import requests

from main import CombinedGenerator

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Конфигурация
BITRIX_WEBHOOK = os.getenv("BITRIX_WEBHOOK", "https://poryadok.bitrix24.ru/rest/159096/1g7d9dxu9rd1kpxc/")
BOT_CODE = "labels_generator_bot"
BOT_NAME = "Генератор Этикеток"
POLL_INTERVAL = 3  # Проверять каждые 3 секунды

BOT_ID = None
LAST_MESSAGE_ID = {}  # Храним ID последнего обработанного сообщения для каждого диалога


class BitrixAPI:
    """Класс для работы с Bitrix24 REST API"""

    def __init__(self, webhook_url: str):
        self.webhook_url = webhook_url.rstrip('/')

    def call(self, method: str, params: Dict = None) -> Dict:
        """Вызов REST API метода"""
        url = f"{self.webhook_url}/{method}"
        try:
            response = requests.post(url, json=params or {}, timeout=30)
            response.raise_for_status()
            result = response.json()
            return result
        except Exception as e:
            logger.error(f"API error {method}: {e}")
            return {"error": str(e)}

    def register_bot(self, code: str, name: str) -> Optional[int]:
        """
        Регистрация бота (без вебхуков, т.к. используем polling)
        Документация: https://apidocs.bitrix24.com/api-reference/chat-bots/imbot-register.html
        """
        params = {
            "CODE": code,
            "TYPE": "B",
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
            # Возможно бот уже зарегистрирован
            if "error_description" in result and "already exists" in result.get("error_description", "").lower():
                logger.warning("⚠️ Бот уже зарегистрирован")
                # Пробуем получить список ботов
                bots = self.get_bot_list()
                for bot in bots:
                    if bot.get("CODE") == code:
                        bot_id = bot.get("ID")
                        logger.info(f"✅ Найден существующий бот! BOT_ID: {bot_id}")
                        return bot_id

            logger.error(f"❌ Ошибка регистрации: {result}")
            return None

    def get_bot_list(self) -> List[Dict]:
        """Получить список ботов"""
        result = self.call("imbot.bot.list")
        if "result" in result:
            return result["result"]
        return []

    def get_dialogs(self) -> List[Dict]:
        """
        Получить список диалогов
        Документация: https://training.bitrix24.com/rest_help/im/im_dialog_get.php
        """
        result = self.call("im.dialog.get")
        if "result" in result:
            return result["result"]
        return []

    def get_messages(self, dialog_id: str, last_id: int = 0) -> List[Dict]:
        """
        Получить сообщения из диалога
        Документация: https://training.bitrix24.com/rest_help/im/im_dialog_messages_get.php
        """
        params = {
            "DIALOG_ID": dialog_id,
            "LIMIT": 20
        }

        if last_id > 0:
            params["LAST_ID"] = last_id

        result = self.call("im.dialog.messages.get", params)

        if "result" in result and "messages" in result["result"]:
            return result["result"]["messages"]
        return []

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

    def download_file(self, file_id: str, save_path: str) -> bool:
        """Скачать файл из Bitrix24"""
        try:
            # Получаем информацию о файле
            result = self.call("disk.file.get", {"id": file_id})

            if "result" not in result:
                logger.error(f"Не удалось получить информацию о файле {file_id}")
                return False

            download_url = result["result"].get("DOWNLOAD_URL")
            if not download_url:
                logger.error("URL для скачивания не найден")
                return False

            # Скачиваем
            response = requests.get(download_url, timeout=60)
            response.raise_for_status()

            with open(save_path, 'wb') as f:
                f.write(response.content)

            logger.info(f"✅ Файл скачан: {save_path}")
            return True

        except Exception as e:
            logger.error(f"❌ Ошибка скачивания файла: {e}")
            return False


class LabelsProcessor:
    """Обработчик архивов"""

    def __init__(self):
        self.generator = CombinedGenerator()
        self.temp_dir = None

    def process_archive(self, archive_path: str) -> Optional[str]:
        """Обработка архива: data.xlsx + img/"""
        try:
            self.temp_dir = tempfile.mkdtemp(prefix="labels_")
            logger.info(f"📁 Temp: {self.temp_dir}")

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
                logger.error("❌ Excel не найден")
                return None

            # Копируем img
            if img_dir_path:
                self._copy_img(img_dir_path)
            else:
                logger.warning("⚠️ Папка img не найдена")

            # Генерируем
            output_dir = os.path.join(self.temp_dir, "output")
            os.makedirs(output_dir)

            for excel_file in excel_files:
                logger.info(f"⚙️ Обработка: {os.path.basename(excel_file)}")
                self.generator.process_excel_file(excel_file, output_dir)

            # Создаем result.zip
            result_archive = os.path.join(self.temp_dir, "result.zip")
            with zipfile.ZipFile(result_archive, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for root, dirs, files in os.walk(output_dir):
                    for file in files:
                        file_path = os.path.join(root, file)
                        arcname = os.path.relpath(file_path, output_dir)
                        zipf.write(file_path, arcname)

            logger.info(f"✅ Результат готов: {os.path.basename(result_archive)}")
            return result_archive

        except Exception as e:
            logger.error(f"❌ Ошибка обработки: {e}", exc_info=True)
            return None

    def _copy_img(self, img_dir_path: str):
        """Копирование img в LabelsMarksGenerator/img/"""
        target_img_dir = "LabelsMarksGenerator/img"

        for item in os.listdir(img_dir_path):
            source_path = os.path.join(img_dir_path, item)
            target_path = os.path.join(target_img_dir, item)

            if os.path.isdir(source_path):
                if os.path.exists(target_path):
                    shutil.rmtree(target_path)
                shutil.copytree(source_path, target_path)
                logger.info(f"📂 Скопирована папка: {item}/")
            else:
                shutil.copy2(source_path, target_path)
                logger.info(f"📄 Скопирован файл: {item}")

    def cleanup(self):
        """Очистка"""
        if self.temp_dir and os.path.exists(self.temp_dir):
            shutil.rmtree(self.temp_dir)


class BitrixBot:
    """Основной класс бота с polling"""

    def __init__(self, api: BitrixAPI):
        self.api = api
        self.processor = LabelsProcessor()
        self.bot_id = None
        self.running = False

    def setup(self):
        """Инициализация бота"""
        logger.info("🔧 Настройка бота...")

        # Регистрируем бота
        self.bot_id = self.api.register_bot(BOT_CODE, BOT_NAME)

        if not self.bot_id:
            logger.error("❌ Не удалось зарегистрировать бота")
            return False

        logger.info(f"✅ Бот готов! BOT_ID: {self.bot_id}")
        return True

    def handle_command(self, dialog_id: str, command: str):
        """Обработка команд"""
        command = command.lower().strip()

        if command == '/start':
            self.api.send_message(dialog_id,
                "🚀 Привет! Я готов к работе.\n\n"
                "Отправь мне ZIP архив со структурой:\n"
                "📂 archive.zip\n"
                "   ├── data.xlsx (обязательно)\n"
                "   └── img/ (опционально)\n"
                "       ├── logos/\n"
                "       ├── certificates/\n"
                "       └── mark_images/\n\n"
                "Я создам этикетки и марки в PDF."
            )

        elif command == '/help':
            self.api.send_message(dialog_id,
                "📋 Инструкция:\n\n"
                "1. Создай ZIP:\n"
                "   - data.xlsx\n"
                "   - img/logos/ (логотипы)\n"
                "   - img/certificates/ (еас.png, рст.png)\n"
                "   - img/mark_images/ (mark_images.png)\n\n"
                "2. Отправь архив сюда\n\n"
                "3. Получишь result.zip с:\n"
                "   - labels/ (этикетки PDF)\n"
                "   - marks/ (марки PDF)\n\n"
                "Команды:\n"
                "/start - начать\n"
                "/help - справка"
            )

        else:
            self.api.send_message(dialog_id, "❓ Неизвестная команда. Используй /help")

    def handle_message(self, message: Dict, dialog_id: str):
        """Обработка входящего сообщения"""
        message_text = message.get("text", "").strip()
        message_id = message.get("id")
        author_id = message.get("author_id")

        # Игнорируем свои сообщения
        if str(author_id) == str(self.bot_id):
            return

        logger.info(f"💬 Сообщение от {author_id} в {dialog_id}: {message_text[:50]}")

        # Команды
        if message_text.startswith('/'):
            self.handle_command(dialog_id, message_text)

        # Файлы
        elif "files" in message and message["files"]:
            self.handle_files(dialog_id, message["files"])

        # Обычное сообщение
        else:
            self.api.send_message(dialog_id,
                "👋 Привет! Отправь мне ZIP архив с данными.\n"
                "Команды: /start, /help"
            )

    def handle_files(self, dialog_id: str, files: List):
        """Обработка файлов"""
        try:
            # Ищем ZIP
            zip_file = None
            for file in files:
                name = file.get("name", "")
                if name.lower().endswith('.zip'):
                    zip_file = file
                    break

            if not zip_file:
                self.api.send_message(dialog_id, "❌ Нужен ZIP архив. Используй /help")
                return

            self.api.send_message(dialog_id, "📦 Получен архив, обрабатываю...")

            # Скачиваем
            file_id = zip_file.get("id")
            temp_archive = os.path.join(tempfile.gettempdir(), f"input_{file_id}.zip")

            if not self.api.download_file(file_id, temp_archive):
                self.api.send_message(dialog_id, "❌ Ошибка скачивания архива")
                return

            self.api.send_message(dialog_id, "⚙️ Генерирую этикетки и марки...")

            # Обрабатываем
            result_archive = self.processor.process_archive(temp_archive)

            if result_archive:
                # TODO: Загрузить результат обратно в Bitrix24
                # Пока просто сообщаем о готовности
                self.api.send_message(dialog_id,
                    "✅ Обработка завершена!\n\n"
                    f"📦 Создан архив: {os.path.basename(result_archive)}\n"
                    f"📊 Размер: {os.path.getsize(result_archive) / 1024:.1f} KB\n\n"
                    "⚠️ Загрузка файла обратно в Bitrix24 в разработке.\n"
                    f"Результат сохранен локально: {result_archive}"
                )
            else:
                self.api.send_message(dialog_id,
                    "❌ Ошибка обработки.\n"
                    "Проверь структуру архива и наличие Excel файла."
                )

            # Очистка
            self.processor.cleanup()
            if os.path.exists(temp_archive):
                os.remove(temp_archive)

        except Exception as e:
            logger.error(f"❌ Ошибка обработки файлов: {e}", exc_info=True)
            self.api.send_message(dialog_id, f"❌ Ошибка: {str(e)}")

    def poll(self):
        """Основной цикл опроса"""
        logger.info("🔄 Запуск polling...")
        self.running = True

        while self.running:
            try:
                # Получаем диалоги
                dialogs = self.api.get_dialogs()

                for dialog in dialogs:
                    dialog_id = dialog.get("id")

                    # Получаем последний ID сообщения для этого диалога
                    last_id = LAST_MESSAGE_ID.get(dialog_id, 0)

                    # Получаем новые сообщения
                    messages = self.api.get_messages(dialog_id, last_id)

                    for message in messages:
                        message_id = message.get("id")

                        # Обрабатываем только новые
                        if message_id > last_id:
                            self.handle_message(message, dialog_id)
                            LAST_MESSAGE_ID[dialog_id] = message_id

                # Ждем перед следующим опросом
                time.sleep(POLL_INTERVAL)

            except KeyboardInterrupt:
                logger.info("\n⛔ Остановка бота...")
                self.running = False
                break

            except Exception as e:
                logger.error(f"❌ Ошибка в цикле polling: {e}")
                time.sleep(POLL_INTERVAL)

    def start(self):
        """Запуск бота"""
        if not self.setup():
            logger.error("❌ Не удалось настроить бота")
            return

        logger.info("=" * 60)
        logger.info("✅ Бот запущен и работает!")
        logger.info(f"🔄 Проверка новых сообщений каждые {POLL_INTERVAL} сек")
        logger.info("📝 Напиши боту в Bitrix24: /start")
        logger.info("⛔ Для остановки нажми Ctrl+C")
        logger.info("=" * 60)

        self.poll()


def main():
    """Точка входа"""
    logger.info("=" * 60)
    logger.info("🤖 Bitrix24 Bot - Генератор Этикеток (Polling Mode)")
    logger.info("=" * 60)
    logger.info(f"📡 Webhook: {BITRIX_WEBHOOK[:50]}...")
    logger.info(f"🔄 Интервал опроса: {POLL_INTERVAL} сек")
    logger.info("=" * 60)
    logger.info("")

    # Создаем API и бота
    api = BitrixAPI(BITRIX_WEBHOOK)
    bot = BitrixBot(api)

    # Запускаем
    bot.start()

    logger.info("\n👋 Бот остановлен")


if __name__ == '__main__':
    main()
