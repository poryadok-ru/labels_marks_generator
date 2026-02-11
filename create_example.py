"""
Скрипт для создания примера архива example_archive.zip
Запуск: python create_example.py
"""

import pandas as pd
import os
from PIL import Image, ImageDraw
import zipfile
import tempfile
import shutil

def create_example_archive():
    """Создает пример архива с тестовыми данными"""

    # Создаем временную директорию
    temp_dir = tempfile.mkdtemp(prefix='example_')
    print(f'📁 Создана временная директория: {temp_dir}')

    try:
        # Создаем Excel с тестовыми данными
        data = {
            'Наименование': [
                'Футболка хлопковая мужская',
                'Штаны спортивные женские',
                'Кроссовки детские'
            ],
            'Артикул': ['TSH-001', 'PNT-002', 'SHO-003'],
            'Код': ['1234567', '2345678', '3456789'],
            'Штрихкод': ['4600000000011', '4600000000028', '4600000000035'],
            'Лого': ['company', 'company', 'company'],
            'Назначение': ['Повседневная одежда', 'Спортивная одежда', 'Детская обувь'],
            'Материал': ['100% хлопок', 'Полиэстер 80%, хлопок 20%', 'Кожа, резина'],
            'Производитель': ['ООО "Текстиль"', 'ООО "Спорт"', 'ООО "Обувь"'],
            'Импортер': ['ООО "Импорт"', 'ООО "Импорт"', 'ООО "Импорт"'],
            'Страна происхождения': ['Россия', 'Китай', 'Вьетнам'],
            'Дата изготовления': ['01.2026', '02.2026', '03.2026'],
            'Сертификация': ['РСТ 123456', 'EAC 654321', 'РСТ 111222'],
            'Тип сертификации': ['РСТ', 'EAC', 'РСТ']
        }

        df = pd.DataFrame(data)
        excel_path = os.path.join(temp_dir, 'data.xlsx')
        df.to_excel(excel_path, index=False)
        print('✅ Excel создан: data.xlsx')

        # Создаем структуру img/
        img_dir = os.path.join(temp_dir, 'img')
        os.makedirs(img_dir)

        # Создаем папки
        logos_dir = os.path.join(img_dir, 'logos')
        certs_dir = os.path.join(img_dir, 'certificates')
        marks_dir = os.path.join(img_dir, 'mark_images')

        os.makedirs(logos_dir)
        os.makedirs(certs_dir)
        os.makedirs(marks_dir)

        # Логотип company.png
        logo = Image.new('RGB', (200, 60), color='white')
        draw = ImageDraw.Draw(logo)
        draw.rectangle([10, 10, 190, 50], outline='blue', width=3)
        draw.text((100, 30), 'COMPANY', fill='blue', anchor='mm')
        logo.save(os.path.join(logos_dir, 'company.png'))
        print('✅ Создан логотип: company.png')

        # Сертификат РСТ
        rst = Image.new('RGB', (100, 100), color='white')
        draw = ImageDraw.Draw(rst)
        draw.ellipse([10, 10, 90, 90], outline='green', width=3)
        draw.text((50, 50), 'РСТ', fill='green', anchor='mm')
        rst.save(os.path.join(certs_dir, 'рст.png'))
        print('✅ Создан сертификат: рст.png')

        # Сертификат EAC
        eac = Image.new('RGB', (100, 100), color='white')
        draw = ImageDraw.Draw(eac)
        draw.ellipse([10, 10, 90, 90], outline='red', width=3)
        draw.text((50, 50), 'EAC', fill='red', anchor='mm')
        eac.save(os.path.join(certs_dir, 'eac.png'))
        print('✅ Создан сертификат: eac.png')

        # Изображение марки
        mark = Image.new('RGB', (200, 200), color='white')
        draw = ImageDraw.Draw(mark)
        draw.rectangle([20, 20, 180, 180], outline='black', width=2)
        draw.text((100, 100), 'MARK', fill='black', anchor='mm')
        mark.save(os.path.join(marks_dir, 'mark_images.png'))
        print('✅ Создана марка: mark_images.png')

        # Создаем ZIP архив
        archive_path = 'example_archive.zip'
        with zipfile.ZipFile(archive_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            # Добавляем Excel
            zipf.write(excel_path, 'data.xlsx')

            # Добавляем папку img со всеми файлами
            for root, dirs, files in os.walk(img_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.join('img', os.path.relpath(file_path, img_dir))
                    zipf.write(file_path, arcname)

        print('✅ Архив создан: example_archive.zip')

        # Проверяем содержимое архива
        print('\n📦 Содержимое архива:')
        with zipfile.ZipFile(archive_path, 'r') as zipf:
            for name in sorted(zipf.namelist()):
                info = zipf.getinfo(name)
                print(f'   {name} ({info.file_size} bytes)')

        print(f'\n✅ Пример архива готов: example_archive.zip')

        return True

    except Exception as e:
        print(f'❌ Ошибка: {e}')
        return False

    finally:
        # Удаляем временную директорию
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)
            print(f'🗑️  Временные файлы удалены')


if __name__ == '__main__':
    print('=' * 60)
    print('Создание примера архива example_archive.zip')
    print('=' * 60)
    print()

    success = create_example_archive()

    print()
    print('=' * 60)
    if success:
        print('✅ Успешно завершено!')
        print('\nТеперь можно:')
        print('1. Отправить example_archive.zip боту в Bitrix24')
        print('2. Использовать как шаблон для своих данных')
    else:
        print('❌ Произошла ошибка')
    print('=' * 60)