"""
Простой эхо-бот для Bitrix24 (polling режим)
Отвечает на все сообщения и поддерживает команды

Как использовать:
1. Найди бота "Генератор Этикеток (RPA)" в Bitrix24
2. Напиши ему любое сообщение
3. Бот ответит!

Команды:
/start - приветствие
/help - справка
/echo [текст] - повторит текст
"""

import os
import time
import requests
import json
from datetime import datetime

# Конфигурация из .env
WEBHOOK = os.getenv("BITRIX_WEBHOOK", "https://poryadok.bitrix24.ru/rest/159096/v8v0zdx3ekts7l8b/")
BOT_ID = int(os.getenv("BOT_ID", "197562"))
CLIENT_ID = os.getenv("CLIENT_ID", "7cqq1c42v4qv8s72iq3eswrqc8mp24vg")

POLL_INTERVAL = 2  # Проверять каждые 2 секунды
LAST_MESSAGE_ID = {}  # Храним последний ID для каждого диалога

print("=" * 60)
print("🤖 Bitrix24 Эхо-Бот")
print("=" * 60)
print(f"🆔 BOT_ID: {BOT_ID}")
print(f"🔄 Интервал опроса: {POLL_INTERVAL} сек")
print("=" * 60)
print()
print("📝 Инструкция:")
print("1. Открой Bitrix24")
print("2. Найди бота 'Генератор Этикеток (RPA)' в чате")
print("3. Напиши ему: /start")
print()
print("🚀 Запуск polling...")
print("⛔ Для остановки: Ctrl+C")
print("=" * 60)
print()


def call_api(method, params=None):
    """Вызов Bitrix24 API"""
    url = f"{WEBHOOK}{method}"
    try:
        response = requests.post(url, json=params or {}, timeout=30)
        response.raise_for_status()
        return response.json()
    except Exception as e:
        return {"error": str(e)}


def send_message(dialog_id, message):
    """Отправка сообщения от имени бота"""
    params = {
        "BOT_ID": BOT_ID,
        "CLIENT_ID": CLIENT_ID,
        "DIALOG_ID": dialog_id,
        "MESSAGE": message
    }

    result = call_api("imbot.message.add", params)

    if "result" in result:
        print(f"  ✅ Отправлено: {message[:50]}...")
        return True
    else:
        print(f"  ❌ Ошибка отправки: {result.get('error_description', result.get('error'))}")
        return False


def handle_command(dialog_id, command):
    """Обработка команд"""
    command = command.strip().lower()

    if command == '/start':
        send_message(dialog_id,
            "🚀 Привет! Я эхо-бот!\n\n"
            "Я умею:\n"
            "• /start - приветствие\n"
            "• /help - справка\n"
            "• /echo [текст] - повторю текст\n"
            "• Просто напиши мне что угодно - я отвечу!\n\n"
            "Попробуй написать: Привет!"
        )

    elif command == '/help':
        send_message(dialog_id,
            "📋 Справка:\n\n"
            "Команды:\n"
            "• /start - начать работу\n"
            "• /help - эта справка\n"
            "• /echo [текст] - повторить текст\n\n"
            "Или просто напиши любой текст - я его повторю!"
        )

    elif command.startswith('/echo '):
        text = command[6:]  # Убираем '/echo '
        send_message(dialog_id, f"🔊 Эхо: {text}")

    else:
        send_message(dialog_id, f"❓ Неизвестная команда: {command}\n\nИспользуй /help")


def handle_message(message, dialog_id):
    """Обработка входящего сообщения"""
    message_text = message.get("text", "").strip()
    message_id = int(message.get("id", 0))
    author_id = int(message.get("author_id", 0))

    # Игнорируем свои сообщения
    if author_id == BOT_ID:
        return

    print(f"💬 [{datetime.now().strftime('%H:%M:%S')}] От {author_id}: {message_text[:50]}")

    # Обработка команд
    if message_text.startswith('/'):
        handle_command(dialog_id, message_text)

    # Эхо любого текста
    else:
        send_message(dialog_id, f"📢 Ты написал: {message_text}")


def get_dialogs():
    """Получить список диалогов"""
    result = call_api("im.dialog.get")

    if "result" in result:
        return result["result"]

    return []


def get_messages(dialog_id, last_id=0):
    """Получить сообщения из диалога"""
    params = {
        "DIALOG_ID": dialog_id,
        "LIMIT": 20
    }

    if last_id > 0:
        params["LAST_ID"] = last_id

    result = call_api("im.dialog.messages.get", params)

    if "result" in result and "messages" in result["result"]:
        return result["result"]["messages"]

    return []


def poll():
    """Основной цикл polling"""
    message_count = 0

    try:
        while True:
            # Получаем диалоги
            dialogs = get_dialogs()

            for dialog in dialogs:
                dialog_id = dialog.get("id")

                # Получаем последний ID для этого диалога
                last_id = LAST_MESSAGE_ID.get(dialog_id, 0)

                # Получаем новые сообщения
                messages = get_messages(dialog_id, last_id)

                # Обрабатываем в обратном порядке (от старых к новым)
                for message in reversed(messages):
                    message_id = int(message.get("id", 0))

                    # Только новые сообщения
                    if message_id > last_id:
                        handle_message(message, dialog_id)
                        LAST_MESSAGE_ID[dialog_id] = message_id
                        message_count += 1

            # Ждем перед следующим опросом
            time.sleep(POLL_INTERVAL)

    except KeyboardInterrupt:
        print()
        print("=" * 60)
        print("⛔ Бот остановлен")
        print(f"📊 Обработано сообщений: {message_count}")
        print("=" * 60)


if __name__ == "__main__":
    # Тестируем подключение
    print("🔍 Проверка подключения к Bitrix24...")
    result = call_api("methods")

    if "result" in result:
        print(f"✅ Подключение OK (доступно методов: {len(result['result'])})")
        print()
    else:
        print(f"❌ Ошибка подключения: {result}")
        exit(1)

    # Запускаем polling
    poll()
