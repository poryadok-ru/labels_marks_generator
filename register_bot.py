"""
Скрипт для регистрации бота в Bitrix24
Запуск: python register_bot.py
"""

import os
import requests
import json

# Конфигурация
BITRIX_WEBHOOK = os.getenv("BITRIX_WEBHOOK", "https://poryadok.bitrix24.ru/rest/159096/1g7d9dxu9rd1kpxc/")
BOT_CODE = "labels_generator_bot"
BOT_NAME = "Генератор Этикеток"

def call_api(method: str, params: dict = None):
    """Вызов Bitrix24 REST API"""
    url = f"{BITRIX_WEBHOOK}{method}"
    try:
        response = requests.post(url, json=params or {}, timeout=30)
        response.raise_for_status()
        return response.json()
    except Exception as e:
        return {"error": str(e)}

def register_bot():
    """Регистрация бота"""
    print("=" * 60)
    print("🤖 Регистрация бота в Bitrix24")
    print("=" * 60)
    print(f"📡 Webhook: {BITRIX_WEBHOOK[:50]}...")
    print(f"🔖 Код бота: {BOT_CODE}")
    print(f"📝 Имя бота: {BOT_NAME}")
    print()

    params = {
        "CODE": BOT_CODE,
        "TYPE": "B",  # B = обычный бот
        "PROPERTIES": {
            "NAME": BOT_NAME,
            "COLOR": "AZURE",
            "WORK_POSITION": "Генератор этикеток и марок из Excel"
        }
    }

    print("📤 Отправка запроса imbot.register...")
    result = call_api("imbot.register", params)

    print()
    print("📥 Ответ от Bitrix24:")
    print(json.dumps(result, indent=2, ensure_ascii=False))
    print()

    if "result" in result:
        bot_id = result["result"]
        print("=" * 60)
        print(f"✅ УСПЕХ! Бот зарегистрирован!")
        print(f"🆔 BOT_ID: {bot_id}")
        print("=" * 60)
        print()
        print("📝 Сохрани BOT_ID в файл .env:")
        print(f"BOT_ID={bot_id}")
        print()
        print("▶️ Теперь можно запустить:")
        print("   python bitrix_bot_polling.py")
        print()
        return bot_id

    elif "error_description" in result:
        error_desc = result.get("error_description", "")

        if "already exists" in error_desc.lower():
            print("⚠️ Бот уже зарегистрирован!")
            print()
            print("📋 Получаю список ботов...")

            # Получаем список ботов
            bots_result = call_api("imbot.bot.list")

            if "result" in bots_result:
                print()
                print("🤖 Найденные боты:")
                for bot in bots_result["result"]:
                    print(f"   - {bot.get('NAME')} (ID: {bot.get('ID')}, CODE: {bot.get('CODE')})")

                    if bot.get("CODE") == BOT_CODE:
                        bot_id = bot.get("ID")
                        print()
                        print("=" * 60)
                        print(f"✅ Найден твой бот!")
                        print(f"🆔 BOT_ID: {bot_id}")
                        print("=" * 60)
                        print()
                        print("▶️ Можно запускать:")
                        print("   python bitrix_bot_polling.py")
                        print()
                        return bot_id
            else:
                print("❌ Не удалось получить список ботов")
                print(json.dumps(bots_result, indent=2, ensure_ascii=False))
        else:
            print("❌ Ошибка регистрации:")
            print(f"   {error_desc}")

        return None

    else:
        print("❌ Неожиданный ответ от API")
        return None

def check_api():
    """Проверка доступности API"""
    print("🔍 Проверка доступа к Bitrix24 API...")
    result = call_api("methods")

    if "result" in result:
        methods = result["result"]
        print(f"✅ API доступен! Доступно методов: {len(methods)}")

        # Проверяем наличие методов бота
        imbot_methods = [m for m in methods if m.startswith("imbot.")]
        print(f"🤖 Методов для ботов: {len(imbot_methods)}")

        if imbot_methods:
            print("   Примеры:", ", ".join(imbot_methods[:5]))

        return True
    else:
        print("❌ Ошибка доступа к API:")
        print(json.dumps(result, indent=2, ensure_ascii=False))
        return False

if __name__ == "__main__":
    # Проверяем API
    if not check_api():
        print()
        print("⚠️ Проверь BITRIX_WEBHOOK в .env файле")
        exit(1)

    print()

    # Регистрируем бота
    bot_id = register_bot()

    if bot_id:
        print("=" * 60)
        print("🎉 Всё готово!")
        print("=" * 60)
    else:
        print()
        print("=" * 60)
        print("❌ Регистрация не удалась")
        print("=" * 60)
        print()
        print("💡 Возможные причины:")
        print("   1. Бот с таким CODE уже существует")
        print("   2. Недостаточно прав")
        print("   3. Проблема с webhook URL")
        print()
        print("🔧 Что попробовать:")
        print("   1. Измени BOT_CODE на другой")
        print("   2. Проверь права webhook")
        print("   3. Попробуй удалить старого бота через Bitrix24")
