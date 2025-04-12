#!/usr/bin/env python
"""
Демонстрационный скрипт для использования ExcelManager для вставки изображений в Excel
"""
import os
import sys
import logging
import argparse
import subprocess
import importlib.util
from pathlib import Path

# Функция для проверки и установки необходимых библиотек
def check_and_install_dependencies():
    """
    Проверяет наличие необходимых библиотек и устанавливает их при необходимости
    """
    # Проверяем, был ли передан флаг, сигнализирующий о том, что мы уже пытались установить зависимости
    if "--deps_installed" in sys.argv:
        # Если флаг есть, просто выходим из функции, чтобы избежать зацикливания
        return True
    
    # Список необходимых библиотек
    required_packages = [
        "openpyxl",
        "Pillow",
        "streamlit",
        "numpy",
        "pandas",
        "watchdog",
        "python-dotenv"
    ]
    
    packages_to_install = []
    
    # Проверяем наличие каждой библиотеки
    for package in required_packages:
        try:
            # Пытаемся импортировать пакет напрямую, это более надежный способ
            if package == "Pillow":
                __import__("PIL")
            else:
                package_name = package.split(">=")[0]
                __import__(package_name)
        except ImportError:
            packages_to_install.append(package)
    
    # Если есть библиотеки для установки
    if packages_to_install:
        print("Обнаружены отсутствующие библиотеки. Начинаю установку...")
        
        for package in packages_to_install:
            print(f"Устанавливаю {package}...")
            try:
                subprocess.check_call([sys.executable, "-m", "pip", "install", package])
                print(f"Библиотека {package} успешно установлена")
            except subprocess.CalledProcessError as e:
                print(f"Ошибка при установке библиотеки {package}: {e}")
                return False
        
        print("Все необходимые библиотеки установлены")
        
        # Перезапускаем скрипт после установки библиотек с флагом, указывающим, что библиотеки уже были установлены
        print("Перезапуск скрипта...")
        new_args = sys.argv.copy()
        new_args.append("--deps_installed")
        os.execv(sys.executable, [sys.executable] + new_args)
    
    return True

# Проверяем и устанавливаем необходимые библиотеки до импорта модулей
if not check_and_install_dependencies():
    print("Возникли проблемы при установке необходимых библиотек. Программа будет завершена.")
    sys.exit(1)

# Удаляем флаг --deps_installed из аргументов, чтобы он не мешал обработке аргументов командной строки
if "--deps_installed" in sys.argv:
    sys.argv.remove("--deps_installed")

# Настройка логирования
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler(os.path.join(os.path.dirname(__file__), 'logs', 'excel_with_images.log'))
    ]
)
logger = logging.getLogger(__name__)

# Добавляем корневую папку проекта в PYTHONPATH
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

# Импортируем после проверки зависимостей
from excel_manager import ExcelManager

def ensure_dir(directory):
    """Создает директорию, если она не существует"""
    if not os.path.exists(directory):
        os.makedirs(directory)

def demo_create_new_with_images(images_dir, output_file):
    """
    Демонстрирует создание нового Excel-файла с изображениями
    
    Args:
        images_dir (str): Путь к директории с изображениями
        output_file (str): Путь к выходному файлу Excel
    """
    logger.info(f"Создание нового файла Excel с изображениями из: {images_dir}")
    
    # Создаем экземпляр ExcelManager
    manager = ExcelManager()
    
    # Задаем заголовки таблицы
    headers = ["Изображение", "Артикул", "Наименование", "Цена"]
    
    # Создаем каталог с изображениями
    success = manager.create_catalog_with_images(
        image_dir=images_dir,
        output_file=output_file,
        headers=headers,
        image_width=300,
        image_height=200
    )
    
    if success:
        logger.info(f"Файл успешно создан: {output_file}")
    else:
        logger.error(f"Ошибка при создании файла: {output_file}")

def demo_add_images_to_existing(excel_file, images_dir, output_file):
    """
    Демонстрирует добавление изображений в существующий Excel-файл
    
    Args:
        excel_file (str): Путь к существующему файлу Excel
        images_dir (str): Путь к директории с изображениями
        output_file (str): Путь к выходному файлу Excel
    """
    logger.info(f"Добавление изображений в существующий файл: {excel_file}")
    
    # Создаем экземпляр ExcelManager с существующим файлом
    manager = ExcelManager(excel_file)
    
    # Добавляем изображения по артикулам
    count = manager.add_images_by_article(
        image_dir=images_dir,
        article_column="B",  # столбец с артикулами
        image_column="A",    # столбец для вставки изображений
        start_row=2,         # начиная со второй строки (после заголовков)
        max_width=300,
        max_height=200
    )
    
    logger.info(f"Добавлено {count} изображений")
    
    # Сохраняем результат
    if manager.save(output_file):
        logger.info(f"Файл успешно сохранен: {output_file}")
    else:
        logger.error(f"Ошибка при сохранении файла: {output_file}")

def demo_create_image_grid(images_dir, output_file):
    """
    Демонстрирует создание сетки изображений
    
    Args:
        images_dir (str): Путь к директории с изображениями
        output_file (str): Путь к выходному файлу Excel
    """
    logger.info(f"Создание сетки изображений из: {images_dir}")
    
    # Создаем экземпляр ExcelManager
    manager = ExcelManager()
    
    # Получаем список всех изображений в директории
    images = []
    for filename in os.listdir(images_dir):
        if filename.lower().endswith(('.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.webp')):
            images.append(os.path.join(images_dir, filename))
    
    # Создаем сетку изображений
    if images:
        success = manager.create_image_grid(
            images=images,
            output_file=output_file,
            cols=3,               # количество столбцов в сетке
            image_width=250,
            image_height=200,
            include_names=True    # добавлять имена файлов
        )
        
        if success:
            logger.info(f"Сетка изображений успешно создана: {output_file}")
        else:
            logger.error(f"Ошибка при создании сетки изображений: {output_file}")
    else:
        logger.error(f"В директории {images_dir} не найдено изображений")

def main():
    """Основная функция"""
    # Создаем парсер аргументов командной строки
    parser = argparse.ArgumentParser(description='Демонстрация работы с ExcelManager для вставки изображений в Excel')
    parser.add_argument('--mode', choices=['new', 'add', 'grid'], default='new',
                        help='Режим работы: new - создать новый файл, add - добавить в существующий, grid - создать сетку')
    parser.add_argument('--excel', help='Путь к существующему файлу Excel (для режима add)')
    parser.add_argument('--images', required=True, help='Путь к директории с изображениями')
    parser.add_argument('--output', required=True, help='Путь для сохранения результата')
    parser.add_argument('--deps_installed', action='store_true', help=argparse.SUPPRESS)
    
    args = parser.parse_args()
    
    # Создаем директории для логов, если они не существуют
    ensure_dir(os.path.join(os.path.dirname(__file__), 'logs'))
    
    # Проверяем, что директория с изображениями существует
    if not os.path.exists(args.images):
        logger.error(f"Директория с изображениями не существует: {args.images}")
        return
    
    # Выполняем соответствующую операцию
    if args.mode == 'new':
        demo_create_new_with_images(args.images, args.output)
    elif args.mode == 'add':
        if not args.excel or not os.path.exists(args.excel):
            logger.error("Для режима add необходимо указать существующий файл Excel (--excel)")
            return
        demo_add_images_to_existing(args.excel, args.images, args.output)
    elif args.mode == 'grid':
        demo_create_image_grid(args.images, args.output)

if __name__ == '__main__':
    main() 