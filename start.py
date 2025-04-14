#!/usr/bin/env python
"""
Скрипт запуска ExcelWithImages с выбором режима работы
"""
import os
import sys
import subprocess
import platform
import argparse
import shutil
from pathlib import Path

# Добавляем текущую директорию в PYTHONPATH
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

# Функция для проверки и создания структуры проекта
def ensure_project_structure():
    """
    Проверяет и создает необходимую структуру проекта
    """
    # Корневая директория проекта
    root_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Обязательные директории
    required_dirs = [
        "app",
        "examples",
        "examples/sample_data",
        "examples/output",
        "logs",
        "settings_presets",
        "utils",
        "utils/config_manager",
        "temp"
    ]
    
    # Создаем все необходимые директории
    for d in required_dirs:
        dir_path = os.path.join(root_dir, d)
        os.makedirs(dir_path, exist_ok=True)
        print(f"Директория {d} проверена")
    
    # Проверяем наличие важных файлов
    if not os.path.exists(os.path.join(root_dir, "utils/config_manager/__init__.py")):
        print("ВНИМАНИЕ: Отсутствует файл utils/config_manager/__init__.py")
    
    if not os.path.exists(os.path.join(root_dir, "utils/config_manager/config_manager.py")):
        print("ВНИМАНИЕ: Отсутствует файл utils/config_manager/config_manager.py")
    
    if not os.path.exists(os.path.join(root_dir, "app/app.py")):
        print("ВНИМАНИЕ: Отсутствует файл app/app.py")

# Функция для очистки временных файлов
def clean_temp_directory():
    """
    Удаляет все файлы в папке temp при запуске приложения
    """
    # Путь к директории temp
    temp_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "temp")
    
    # Проверяем существование директории
    if os.path.exists(temp_dir) and os.path.isdir(temp_dir):
        print("Очистка временных файлов...")
        # Перебираем все файлы в директории temp
        for filename in os.listdir(temp_dir):
            file_path = os.path.join(temp_dir, filename)
            try:
                # Если это файл, удаляем его
                if os.path.isfile(file_path):
                    os.unlink(file_path)
                    print(f"Удален временный файл: {filename}")
                # Если это папка, удаляем её со всем содержимым
                elif os.path.isdir(file_path):
                    shutil.rmtree(file_path)
                    print(f"Удалена временная папка: {filename}")
            except Exception as e:
                print(f"Ошибка при удалении {file_path}: {e}")
        print("Очистка временных файлов завершена")
    else:
        print("Директория temp не существует или не является директорией")

# Функция для очистки консоли
def clear_screen():
    """Очищает экран консоли в зависимости от операционной системы"""
    if platform.system() == "Windows":
        os.system('cls')
    else:
        os.system('clear')

# Функция для запуска веб-интерфейса
def start_web_app():
    """Запускает веб-интерфейс на Streamlit"""
    # Очищаем временные файлы перед запуском
    clean_temp_directory()
    
    app_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "app", "app.py")
    
    if not os.path.exists(app_path):
        print(f"Ошибка: Файл приложения не найден: {app_path}")
        input("Нажмите Enter для продолжения...")
        return
        
    print("Запуск веб-интерфейса ExcelWithImages...")
    try:
        subprocess.run(["streamlit", "run", app_path])
    except Exception as e:
        print(f"Ошибка при запуске веб-интерфейса: {e}")
        input("Нажмите Enter для продолжения...")

# Функция для запуска консольного режима
def start_console_mode():
    """Запускает консольный режим с выбором примеров"""
    while True:
        clear_screen()
        print("=== ExcelWithImages: Консольный режим ===")
        print("Выберите пример для запуска:")
        print("1. Создать новый Excel с изображениями")
        print("2. Добавить изображения в существующий Excel")
        print("3. Создать сетку изображений")
        print("4. Добавить одно изображение в ячейку")
        print("0. Вернуться в главное меню")
        
        choice = input("\nВаш выбор: ")
        
        if choice == "1":
            path_to_images = input("Введите путь к директории с изображениями: ")
            output_file = input("Введите путь для сохранения файла Excel: ")
            
            # Создаем директории, если они не существуют
            if output_file:
                os.makedirs(os.path.dirname(output_file), exist_ok=True)
            
            # Проверяем наличие директории с изображениями
            if not os.path.exists(path_to_images):
                print(f"Ошибка: Директория {path_to_images} не существует")
                input("\nНажмите Enter для продолжения...")
                continue
            
            # Запускаем main.py в режиме создания нового файла
            subprocess.run([
                sys.executable, 
                os.path.join(os.path.dirname(os.path.abspath(__file__)), "main.py"),
                "--mode", "new",
                "--images", path_to_images,
                "--output", output_file
            ])
            
            input("\nНажмите Enter для продолжения...")
            
        elif choice == "2":
            excel_file = input("Введите путь к существующему файлу Excel: ")
            
            # Проверяем наличие файла Excel
            if not os.path.exists(excel_file):
                print(f"Ошибка: Файл {excel_file} не существует")
                input("\nНажмите Enter для продолжения...")
                continue
                
            path_to_images = input("Введите путь к директории с изображениями: ")
            
            # Проверяем наличие директории с изображениями
            if not os.path.exists(path_to_images):
                print(f"Ошибка: Директория {path_to_images} не существует")
                input("\nНажмите Enter для продолжения...")
                continue
                
            output_file = input("Введите путь для сохранения результата: ")
            
            # Создаем директории, если они не существуют
            if output_file:
                os.makedirs(os.path.dirname(output_file), exist_ok=True)
            
            # Запускаем main.py в режиме добавления в существующий файл
            subprocess.run([
                sys.executable, 
                os.path.join(os.path.dirname(os.path.abspath(__file__)), "main.py"),
                "--mode", "add",
                "--excel", excel_file,
                "--images", path_to_images,
                "--output", output_file
            ])
            
            input("\nНажмите Enter для продолжения...")
            
        elif choice == "3":
            path_to_images = input("Введите путь к директории с изображениями: ")
            
            # Проверяем наличие директории с изображениями
            if not os.path.exists(path_to_images):
                print(f"Ошибка: Директория {path_to_images} не существует")
                input("\nНажмите Enter для продолжения...")
                continue
                
            output_file = input("Введите путь для сохранения файла Excel с сеткой: ")
            
            # Создаем директории, если они не существуют
            if output_file:
                os.makedirs(os.path.dirname(output_file), exist_ok=True)
            
            # Запускаем main.py в режиме создания сетки
            subprocess.run([
                sys.executable, 
                os.path.join(os.path.dirname(os.path.abspath(__file__)), "main.py"),
                "--mode", "grid",
                "--images", path_to_images,
                "--output", output_file
            ])
            
            input("\nНажмите Enter для продолжения...")
            
        elif choice == "4":
            # Проверяем наличие примера
            example_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "examples", "add_single_image.py")
            if not os.path.exists(example_path):
                print(f"Ошибка: Пример не найден: {example_path}")
                input("\nНажмите Enter для продолжения...")
                continue
                
            # Запускаем пример добавления одного изображения
            subprocess.run([
                sys.executable, 
                example_path
            ])
            
            input("\nНажмите Enter для продолжения...")
            
        elif choice == "0":
            return
        else:
            print("Неверный выбор. Пожалуйста, попробуйте снова.")
            input("\nНажмите Enter для продолжения...")

# Главное меню
def main_menu():
    """Отображает главное меню программы"""
    while True:
        clear_screen()
        print("=== ExcelWithImages ===")
        print("Выберите режим работы:")
        print("1. Веб-интерфейс (Streamlit)")
        print("2. Консольный режим")
        print("0. Выход")
        
        choice = input("\nВаш выбор: ")
        
        if choice == "1":
            start_web_app()
        elif choice == "2":
            start_console_mode()
        elif choice == "0":
            print("Выход из программы. До свидания!")
            sys.exit(0)
        else:
            print("Неверный выбор. Пожалуйста, попробуйте снова.")
            input("\nНажмите Enter для продолжения...")

# Обработка аргументов командной строки
def parse_args():
    """Обрабатывает аргументы командной строки"""
    parser = argparse.ArgumentParser(description="ExcelWithImages - Инструмент для работы с Excel и изображениями")
    parser.add_argument("--web", action="store_true", help="Запустить веб-интерфейс")
    parser.add_argument("--console", action="store_true", help="Запустить консольный режим")
    
    return parser.parse_args()

if __name__ == "__main__":
    # Проверяем структуру проекта
    ensure_project_structure()
    
    # Очищаем временные файлы
    clean_temp_directory()
    
    args = parse_args()
    
    # Проверка наличия модуля streamlit
    try:
        import streamlit
    except ImportError:
        print("Установка необходимых зависимостей...")
        requirements_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "requirements.txt")
        if os.path.exists(requirements_file):
            subprocess.run([sys.executable, "-m", "pip", "install", "-r", requirements_file])
        else:
            print("Файл requirements.txt не найден, устанавливаем основные зависимости...")
            subprocess.run([sys.executable, "-m", "pip", "install", "openpyxl", "Pillow", "streamlit"])
    
    # Проверка каталогов
    examples_data_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "examples", "sample_data")
    examples_output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "examples", "output")
    logs_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")
    
    # Создаем необходимые директории
    os.makedirs(examples_data_dir, exist_ok=True)
    os.makedirs(examples_output_dir, exist_ok=True)
    os.makedirs(logs_dir, exist_ok=True)
    
    # Если образца изображения нет, создаем его
    sample_image_path = os.path.join(examples_data_dir, "sample_image.jpg")
    if not os.path.exists(sample_image_path):
        try:
            from PIL import Image, ImageDraw
            img = Image.new('RGB', (200, 150), color=(73, 109, 137))
            d = ImageDraw.Draw(img)
            d.text((10, 10), 'Sample Image', fill=(255, 255, 0))
            img.save(sample_image_path)
            print(f"Создано тестовое изображение: {sample_image_path}")
        except Exception as e:
            print(f"Не удалось создать тестовое изображение: {e}")
    
    # Автоматически запускаем веб-интерфейс без отображения меню
    print("Запуск веб-интерфейса ExcelWithImages...")
    start_web_app()