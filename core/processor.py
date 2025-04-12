import os
import sys
import logging
import pandas as pd
from datetime import datetime
import tempfile
from pathlib import Path
import json
import time
from typing import Dict, List, Any, Optional, Tuple

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
logger = logging.getLogger(__name__)

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

def process_excel_file(file_path, article_column, image_folder, image_output_folder=None, 
                       output_path=None, output_folder=None, image_column_name='Image_Path',
                       max_width=None, max_height=None, quality=85, **kwargs) -> Tuple[str, pd.DataFrame, int]:
    """
    Обрабатывает Excel-файл, добавляя изображения для каждой строки.
    
    Args:
        file_path (str): Путь к Excel-файлу
        article_column (str): Буквенное обозначение столбца с артикулами
        image_folder (str): Папка с изображениями
        image_output_folder (str, optional): Папка для сохранения обработанных изображений
        output_path (str, optional): Путь для сохранения результата
        output_folder (str, optional): Папка для сохранения результата
        image_column_name (str, optional): Название столбца для путей к изображениям
        max_width (int, optional): Максимальная ширина изображения
        max_height (int, optional): Максимальная высота изображения
        quality (int, optional): Качество сжатия изображений (1-100)
        **kwargs: Дополнительные параметры

    Returns:
        tuple: (output_path, processed_df, images_inserted)
        
    Raises:
        FileNotFoundError: Если файл Excel или папка с изображениями не найдены
        ValueError: Если Excel-файл пуст или указан неверный столбец с артикулами
        Exception: Другие ошибки обработки
    """
    logger.info(f"Обработка Excel-файла: {file_path}")
    logger.info(f"Параметры: article_column={article_column}, image_folder={image_folder}, "
                f"image_output_folder={image_output_folder}, output_path={output_path}, "
                f"output_folder={output_folder}, image_column_name={image_column_name}, "
                f"max_width={max_width}, max_height={max_height}, quality={quality}")
    
    # Проверка наличия файла
    if not os.path.exists(file_path):
        logger.error(f"Файл не найден: {file_path}")
        raise FileNotFoundError(f"Файл не найден: {file_path}")
    
    # Проверка наличия папки с изображениями
    if not os.path.exists(image_folder):
        logger.error(f"Папка с изображениями не найдена: {image_folder}")
        raise FileNotFoundError(f"Папка с изображениями не найдена: {image_folder}")
    
    # Чтение Excel-файла
    try:
        logger.debug(f"Чтение Excel-файла: {file_path}")
        # Извлекаем параметр sheet_name из kwargs, если он есть
        sheet_name = kwargs.get('sheet_name', 0)  # По умолчанию берём первый лист
        logger.debug(f"Используем лист: {sheet_name}")
        
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        logger.debug(f"Excel-файл успешно прочитан, форма DataFrame: {df.shape}")
    except Exception as e:
        logger.error(f"Ошибка при чтении Excel-файла: {e}")
        raise
    
    # Проверка наличия данных в DataFrame
    if df.empty:
        logger.error("Excel-файл не содержит данных")
        raise ValueError("Excel-файл не содержит данных")
    
    # Проверка наличия достаточного количества столбцов
    try:
        # Преобразование буквенного обозначения столбца в индекс
        article_idx = excel_utils.column_letter_to_index(article_column)
        
        # Проверка, что индекс валидный
        if article_idx < 0:
            logger.error(f"Неверное буквенное обозначение столбца: {article_column}")
            raise ValueError(f"Неверное буквенное обозначение столбца: {article_column}")
        
        # Проверка, существует ли указанный столбец в DataFrame
        if article_idx >= len(df.columns):
            logger.error(f"Столбец с индексом {article_idx} (обозначение {article_column}) не существует в файле")
            raise ValueError(f"Столбец с обозначением {article_column} не существует в файле. "
                           f"Доступны столбцы от A до {excel_utils.convert_column_index_to_letter(len(df.columns) - 1)}")
    except ValueError as e:
        logger.error(f"Ошибка при определении индекса столбца: {e}")
        raise
        
    # Создание временной директории для обработанных изображений, если не указана
    if not image_output_folder:
        image_output_folder = ensure_temp_dir('processed_images')
        logger.info(f"Создана временная директория для обработанных изображений: {image_output_folder}")
    
    # Добавление столбца для путей к изображениям, если его еще нет
    if image_column_name not in df.columns:
        df[image_column_name] = None
        logger.debug(f"Добавлен столбец {image_column_name} для путей к изображениям")
    
    # Обработка каждой строки DataFrame
    processed_images_count = 0
    errors_count = 0
    
    logger.info(f"Начинаем обработку {len(df)} строк")
    for idx, row in df.iterrows():
        try:
            # Получение артикула из строки
            article_value = row.iloc[article_idx]
            
            # Пропускаем пустые значения
            if pd.isna(article_value) or article_value == "":
                logger.warning(f"Пропуск строки {idx+1}: пустое значение артикула")
                continue
            
            # Преобразуем в строку
            article = str(article_value)
            
            logger.debug(f"Обработка строки {idx+1}, артикул: {article}")
            
            # Нормализация артикула
            normalized_article = image_utils.normalize_article(article)
            logger.debug(f"Нормализованный артикул: {normalized_article}")
            
            # Поиск изображения для артикула
            image_path = image_utils.find_image_by_article(normalized_article, image_folder)
            
            if image_path:
                logger.info(f"Найдено изображение для артикула {article}: {image_path}")
                
                # Оптимизация изображения, если указаны параметры
                if max_width or max_height:
                    processed_image_path = os.path.join(
                        image_output_folder, 
                        f"{normalized_article}_{int(time.time())}.jpg"
                    )
                    
                    # Обработка изображения
                    img_buffer, dimensions = image_utils.process_image(
                        image_path=image_path,
                        width=max_width,
                        height=max_height,
                        max_size_kb=quality * 10  # Примерная зависимость размера от качества
                    )
                    
                    # Сохранение в файл
                    with open(processed_image_path, 'wb') as f:
                        f.write(img_buffer.getvalue())
                    
                    logger.debug(f"Изображение оптимизировано: {processed_image_path}, размеры: {dimensions}")
                    
                    # Обновление пути к изображению
                    df.at[idx, image_column_name] = processed_image_path
                else:
                    # Использование оригинального изображения
                    df.at[idx, image_column_name] = image_path
                    logger.debug(f"Использовано оригинальное изображение: {image_path}")
                
                processed_images_count += 1
            else:
                logger.warning(f"Изображение для артикула '{article}' не найдено")
                
        except Exception as e:
            errors_count += 1
            logger.error(f"Ошибка при обработке строки {idx+1}: {e}")
            # Пропускаем строку с ошибкой и продолжаем
            continue
    
    logger.info(f"Обработка завершена: обработано {processed_images_count} изображений из {len(df)} строк, ошибок: {errors_count}")
    
    # Определение пути для сохранения результата
    if not output_path:
        if output_folder:
            base_name = os.path.basename(file_path)
            output_path = os.path.join(output_folder, f"processed_{base_name}")
        else:
            # Добавление суффикса к имени файла
            name, ext = os.path.splitext(file_path)
            output_path = f"{name}_processed{ext}"
    
    # Создание директории для выходного файла, если она не существует
    output_dir = os.path.dirname(output_path)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
        logger.info(f"Создана директория для выходного файла: {output_dir}")
    
    # Сохранение DataFrame в Excel
    try:
        logger.info(f"Сохранение результата в файл: {output_path}")
        
        # Сначала сохраняем DataFrame в Excel без изображений
        df.to_excel(output_path, index=False, sheet_name='Sheet1', header=False)
        
        # Теперь открываем файл через openpyxl и добавляем изображения
        from openpyxl import load_workbook
        from openpyxl.drawing.image import Image
        
        # Открываем созданный файл
        logger.info(f"Открываем файл {output_path} для вставки изображений")
        wb = load_workbook(output_path)
        ws = wb.active
        
        # Инициализируем счетчик вставленных изображений
        images_inserted = 0
        
        # Вставляем изображения
        logger.info(f"Вставляем изображения в Excel из столбца {image_column_name}")
        img_col = df.columns.get_loc(image_column_name)
        
        for idx, row in df.iterrows():
            if pd.notna(row[image_column_name]) and os.path.exists(row[image_column_name]):
                img_path = row[image_column_name]
                cell_address = f"{excel_utils.convert_column_index_to_letter(img_col + 1)}{idx + 1}"
                logger.debug(f"Вставка изображения {img_path} в ячейку {cell_address}")
                
                try:
                    # Создаем объект изображения
                    img = Image(img_path)
                    
                    # Установим разумные размеры изображения, если они не указаны
                    if not max_width and not max_height:
                        img.width = 150
                        img.height = 100
                    else:
                        if max_width:
                            img.width = max_width
                        if max_height:
                            img.height = max_height
                    
                    # Добавляем изображение на лист с привязкой к ячейке
                    ws.add_image(img, cell_address)
                    images_inserted += 1
                    logger.debug(f"Изображение успешно вставлено в ячейку {cell_address}")
                    
                except Exception as img_err:
                    logger.error(f"Ошибка при вставке изображения в ячейку {cell_address}: {str(img_err)}")
        
        # Устанавливаем оптимальные размеры ячеек для отображения изображений
        if images_inserted > 0:
            # Устанавливаем ширину столбца с изображениями
            ws.column_dimensions[excel_utils.convert_column_index_to_letter(img_col + 1)].width = 25
            
            # Устанавливаем высоту строк
            for idx in range(len(df)):
                ws.row_dimensions[idx + 1].height = 80
        
        # Сохраняем файл с изображениями
        logger.info(f"Сохраняем файл с {images_inserted} изображениями")
        wb.save(output_path)
        
        logger.info(f"Вставлено изображений в Excel: {images_inserted}")
    
    except Exception as e:
        logger.error(f"Ошибка при сохранении результатов: {e}")
        raise
    
    logger.info(f"Результат сохранен: {output_path}")
    return output_path, df, images_inserted 