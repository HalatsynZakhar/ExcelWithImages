import os
import sys
import logging
import pandas as pd
from datetime import datetime
import tempfile
from pathlib import Path
import json
import time
from typing import Dict, List, Any, Optional

# Add parent directory to path
current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.dirname(current_dir)
if parent_dir not in sys.path:
    sys.path.append(parent_dir)

# Import utils modules directly
from utils import config_manager
from utils import excel_utils
from utils import image_utils

# Import get_downloads_folder from config_manager
from utils.config_manager import get_downloads_folder

# Setup logging
log = logging.getLogger(__name__)

# Helper function for temp directories
def ensure_temp_dir(prefix: str = "") -> str:
    """
    Создает и возвращает путь к временной директории.
    
    Args:
        prefix (str): Префикс для имени временной директории
    
    Returns:
        Путь к временной директории
    """
    temp_dir = os.path.join(tempfile.gettempdir(), f"{prefix}excelwithimages")
    os.makedirs(temp_dir, exist_ok=True)
    return temp_dir

def process_excel_file(excel_file_path, article_column, image_column, image_folder, sheet_name=0):
    """
    Обрабатывает Excel файл и добавляет изображения к соответствующим артикулам.
    
    Args:
        excel_file_path (str): Путь к Excel файлу для обработки
        article_column (str): Имя колонки с артикулами
        image_column (str): Имя колонки для вставки изображений
        image_folder (str): Путь к папке с изображениями
        sheet_name (str, int): Имя или индекс листа для обработки
        
    Returns:
        str: Путь к обработанному файлу
    """
    # Логируем параметры обработки
    log.info(f"Начало обработки файла: {excel_file_path}")
    log.info(f"Колонка с артикулами: {article_column}")
    log.info(f"Колонка для изображений: {image_column}")
    log.info(f"Папка с изображениями: {image_folder}")
    log.info(f"Выбранный лист: {sheet_name}")
    
    try:
        # Логируем параметры для отладки
        log.info(f"Обработка Excel-файла: {excel_file_path}")
        log.info(f"Колонка с артикулами: {article_column}, Колонка для изображений: {image_column}")
        log.info(f"Папка с изображениями: {image_folder}, Имя листа: {sheet_name}")
        
        # Чтение Excel-файла
        log.info(f"Чтение Excel-файла: {excel_file_path}")
        if not os.path.exists(excel_file_path):
            raise FileNotFoundError(f"Файл не найден: {excel_file_path}")
            
        # Загружаем файл с указанным листом
        if sheet_name:
            df = pd.read_excel(excel_file_path, sheet_name=sheet_name)
        else:
            df = pd.read_excel(excel_file_path)
            
        log.info(f"Файл успешно загружен. Размерность данных: {df.shape}")
        
        # Преобразуем буквенные обозначения столбцов в индексы/имена столбцов
        try:
            # Получаем индекс колонки для article_column (буква -> индекс)
            article_idx = excel_utils.column_letter_to_index(article_column)
            article_name = df.columns[article_idx]
            log.info(f"Колонка с артикулами преобразована: {article_column} -> {article_name}")
            
            # Получаем индекс колонки для image_column (буква -> индекс)
            if image_column:
                image_idx = excel_utils.column_letter_to_index(image_column)
                image_name = df.columns[image_idx]
                log.info(f"Колонка для изображений преобразована: {image_column} -> {image_name}")
            else:
                image_name = None
        except Exception as e:
            log.error(f"Ошибка при преобразовании буквенного обозначения колонки: {e}")
            raise ValueError(f"Неверное буквенное обозначение колонки: {article_column} или {image_column}. Ошибка: {e}")
        
        # Создаем временную директорию для обработанных изображений
        temp_dir = ensure_temp_dir(prefix="excel_images_")
        log.info(f"Создана временная директория для обработанных изображений: {temp_dir}")
        
        # Путь к результирующему файлу
        output_folder = get_downloads_folder()
        os.makedirs(output_folder, exist_ok=True)
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        result_filename = f"processed_{timestamp}_{os.path.basename(excel_file_path)}"
        result_file_path = os.path.join(output_folder, result_filename)
        
        # Массив для хранения путей ко всем обработанным изображениям
        processed_image_paths = []
        total_rows = len(df)
        
        # Обработка каждой строки DataFrame
        for idx, row in df.iterrows():
            try:
                # Получение и нормализация артикула
                article = row[article_name]
                if pd.isna(article):
                    log.warning(f"Строка {idx+1}: артикул отсутствует, пропускаем")
                    continue
                    
                normalized_article = image_utils.normalize_article_number(str(article))
                
                # Поиск изображений для артикула
                image_paths = image_utils.find_images_by_article(image_folder, normalized_article)
                
                if not image_paths:
                    log.warning(f"Для артикула {article} (нормализовано: {normalized_article}) изображения не найдены")
                    continue
                    
                log.info(f"Для артикула {article} найдено {len(image_paths)} изображений")
                
                # Оптимизация изображений для Excel
                processed_images = []
                for img_path in image_paths:
                    try:
                        # Создание имени для обработанного изображения
                        img_filename = f"{normalized_article}_{len(processed_images)}_{os.path.basename(img_path)}"
                        processed_img_path = os.path.join(temp_dir, img_filename)
                        
                        # Оптимизация изображения
                        image_utils.optimize_image_for_excel(img_path, processed_img_path)
                        processed_images.append(processed_img_path)
                        processed_image_paths.append(processed_img_path)
                        
                        log.debug(f"Изображение {img_path} обработано и сохранено как {processed_img_path}")
                    except Exception as img_err:
                        log.error(f"Ошибка при обработке изображения {img_path}: {img_err}")
                
                # Обновление DataFrame с путями к обработанным изображениям
                if processed_images and image_name:
                    df.at[idx, image_name] = ",".join(processed_images)
            except Exception as row_err:
                log.error(f"Ошибка при обработке строки {idx+1}: {row_err}")
        
        # Сохранение обработанного DataFrame в новый Excel файл
        log.info(f"Сохранение результата в файл: {result_file_path}")
        try:
            # Используем ExcelWriter для сохранения с изображениями
            with pd.ExcelWriter(result_file_path, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name=sheet_name if isinstance(sheet_name, str) else 'Sheet1')
                
                # Вставка изображений в файл Excel
                excel_utils.insert_images_to_excel(writer, df, image_name)
                
            # Убеждаемся, что файл сохранен до удаления временных файлов
            if os.path.exists(result_file_path):
                log.info(f"Файл {result_file_path} успешно создан")
            else:
                raise FileNotFoundError(f"Не удалось создать файл {result_file_path}")
                
        except Exception as excel_err:
            log.error(f"Ошибка при сохранении Excel файла: {excel_err}")
            raise
            
        log.info("Обработка файла завершена успешно")
        return result_file_path
        
    except Exception as e:
        log.error(f"Ошибка при обработке файла: {e}")
        raise 