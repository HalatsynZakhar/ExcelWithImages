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
import openpyxl
import shutil
from PIL import Image as PILImage

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

def process_excel_file(
    file_path: str,
    article_column: str,
    image_folder: str,
    image_output_folder: Optional[str] = None,
    output_path: Optional[str] = None,
    output_folder: Optional[str] = None,
    image_column_name: str = 'Image',
    adjust_cell_size: bool = False,
    column_width: int = 150,
    row_height: int = 120,
    max_file_size_mb: int = 50
) -> Tuple[str, pd.DataFrame, int]:
    """
    Обрабатывает Excel-файл, добавляя к нему изображения

    Args:
        file_path (str): Путь к Excel-файлу
        article_column (str): Буквенное обозначение колонки с артикулами
        image_folder (str): Папка, где хранятся изображения
        image_output_folder (Optional[str]): Папка для сохранения обработанных изображений (если None, используется временная)
        output_path (Optional[str]): Полный путь для сохранения результата. Если None, генерируется автоматически.
        output_folder (Optional[str]): Папка для сохранения результата (если output_path не указан). Если None, используется папка исходного файла.
        image_column_name (str): Название колонки для вставки изображений (будет создана или перезаписана).
        adjust_cell_size (bool): Изменять размер ячейки для лучшего отображения изображений.
        column_width (int): Желаемая ширина колонки с изображениями в пикселях (при adjust_cell_size=True).
        row_height (int): Желаемая высота строки с изображениями в пикселях (при adjust_cell_size=True).
        max_file_size_mb (int): Максимальный ОБЩИЙ размер выходного файла Excel в мегабайтах (для расчета лимита на изображение).

    Returns:
        Tuple[str, pd.DataFrame, int]: Путь к выходному файлу, обработанный DataFrame и количество вставленных изображений
    """
    logger.info(f"Начало обработки Excel-файла: {file_path}")
    logger.info(f"Параметры: article_column={article_column}, image_folder={image_folder}, "
               f"image_column_name={image_column_name}, adjust_cell_size={adjust_cell_size}, max_file_size_mb={max_file_size_mb}")

    # --- Валидация входных данных ---
    if not os.path.exists(file_path):
        err_msg = f"Файл не найден: {file_path}"
        logger.error(err_msg)
        raise FileNotFoundError(err_msg)
    
    if not os.path.exists(image_folder):
        err_msg = f"Папка с изображениями не найдена: {image_folder}"
        logger.error(err_msg)
        raise FileNotFoundError(err_msg)

    # --- Чтение Excel ---
    try:
        # Читаем без заголовка, чтобы не потерять первую строку, если она нужна
        # Заголовки получим позже из openpyxl
        df = pd.read_excel(file_path, header=None) 
        logger.info(f"Excel-файл прочитан в DataFrame. Строк: {len(df)}")
        # Загружаем книгу openpyxl для работы с ячейками и вставки изображений
        wb = openpyxl.load_workbook(file_path)
        ws = wb.active # Предполагаем работу с активным листом
        logger.info(f"Загружена рабочая книга, активный лист: {ws.title}")
        # Получаем заголовки из openpyxl (первая строка)
        headers = [cell.value for cell in ws[1]] 
        df.columns = headers # Устанавливаем заголовки в DataFrame
        logger.info(f"Установлены заголовки: {headers}")
        
    except Exception as e:
        err_msg = f"Ошибка при чтении Excel-файла: {e}"
        logger.error(err_msg, exc_info=True)
        raise RuntimeError(err_msg) from e

    if df.empty:
        err_msg = "Excel-файл не содержит данных"
        logger.error(err_msg)
        raise ValueError(err_msg)

    # --- Определение колонки артикулов ---
    try:
        article_col_idx_excel = excel_utils.column_letter_to_index(article_column) # 1-based index for openpyxl
        article_col_name_df = df.columns[article_col_idx_excel - 1] # 0-based index for pandas
        logger.info(f"Колонка с артикулами: '{article_col_name_df}' (Excel: {article_column}, Index: {article_col_idx_excel})")
    except Exception as e:
        err_msg = f"Ошибка при определении колонки с артикулами ('{article_column}'): {e}"
        logger.error(err_msg, exc_info=True)
        raise ValueError(err_msg) from e
        
    if article_col_name_df not in df.columns:
        err_msg = f"Колонка с артикулами '{article_col_name_df}' не найдена в заголовках файла: {list(df.columns)}"
        logger.error(err_msg)
        raise ValueError(err_msg)

    # --- Определение КОЛИЧЕСТВА артикулов для расчета лимита на изображение ---
    article_count = df[article_col_name_df].nunique() # Считаем уникальные артикулы, чтобы не занижать лимит
    if article_count == 0:
        article_count = 1 # Избегаем деления на ноль
        logger.warning("Не найдено уникальных артикулов для расчета лимита размера изображения. Используется значение по умолчанию.")
    else:
        logger.info(f"Найдено {article_count} уникальных артикулов.")
        
    # --- Расчет лимита размера на одно изображение ---
    # Оставляем небольшой запас (например, 10%) для самого Excel файла
    image_size_budget_mb = max_file_size_mb * 0.90 
    target_kb_per_image = (image_size_budget_mb * 1024) / article_count
    # Устанавливаем разумный минимум (например, 10KB) и максимум (например, 5MB)
    target_kb_per_image = max(10, min(target_kb_per_image, 5 * 1024)) 
    logger.info(f"Расчетный лимит размера на изображение: {target_kb_per_image:.2f} КБ")

    # --- Подготовка папки для обработанных изображений ---
    temp_image_dir_created = False
    if not image_output_folder:
        image_output_folder = ensure_temp_dir("processed_images_")
        temp_image_dir_created = True
        logger.info(f"Создана временная директория для обработанных изображений: {image_output_folder}")
    elif not os.path.exists(image_output_folder):
         os.makedirs(image_output_folder)
         logger.info(f"Создана папка для обработанных изображений: {image_output_folder}")


    # --- Подготовка к вставке изображений ---
    # Определяем колонку для вставки в Excel (может отличаться от image_column_name в df)
    try:
        # Если image_column_name уже существует как колонка Excel, используем её
        image_col_idx_excel = excel_utils.find_column_by_header(ws, image_column_name, header_row=1)
        if image_col_idx_excel is None:
            # Если нет, создаем новую колонку в конце
            image_col_idx_excel = ws.max_column + 1
            ws.cell(row=1, column=image_col_idx_excel).value = image_column_name # Добавляем заголовок
            logger.info(f"Колонка '{image_column_name}' не найдена, будет создана новая в Excel (столбец {excel_utils.get_column_letter(image_col_idx_excel)})")
        else:
             logger.info(f"Изображения будут вставляться в существующую колонку '{image_column_name}' (столбец {excel_utils.get_column_letter(image_col_idx_excel)})")
        image_col_letter_excel = excel_utils.get_column_letter(image_col_idx_excel)
        
        # Добавляем/обновляем колонку в DataFrame для хранения путей (если нужно, но вроде не используется)
        # df[image_column_name] = '' # Не будем хранить пути в df, вставляем сразу

    except Exception as e:
         err_msg = f"Ошибка при подготовке колонки для изображений ('{image_column_name}'): {e}"
         logger.error(err_msg, exc_info=True)
         raise RuntimeError(err_msg) from e

    # --- Настройка размеров ячеек (если нужно) ---
    if adjust_cell_size:
        try:
            excel_utils.set_column_width(ws, image_col_letter_excel, column_width / 7) # Приблизительная конвертация
            logger.info(f"Установлена ширина столбца {image_col_letter_excel} на {column_width} пикс.")
        except Exception as e:
            logger.warning(f"Не удалось установить ширину столбца {image_col_letter_excel}: {e}")


    # --- Обработка строк и вставка изображений ---
    images_inserted = 0
    rows_processed = 0
    total_rows_df = len(df) # Используем длину df, т.к. читаем без header
    
    # Итерация по DataFrame (индексы pandas + 2 = строка Excel, т.к. header=None и Excel 1-based)
    for df_index, row_data in df.iterrows():
        excel_row_index = df_index + 2 # +1 for 1-based index, +1 because header was row 1
        rows_processed += 1
        if rows_processed % 50 == 0 or rows_processed == total_rows_df:
             logger.info(f"Обработано {rows_processed} из {total_rows_df} строк ({rows_processed/total_rows_df*100:.1f}%)")

        article = row_data.get(article_col_name_df)
        if pd.isna(article): # Пропускаем строки без артикула
            logger.debug(f"Пропуск строки {excel_row_index}: нет артикула.")
            continue
        
        article_str = str(article).strip()
        if not article_str:
            logger.debug(f"Пропуск строки {excel_row_index}: пустой артикул.")
            continue

        logger.debug(f"Обработка строки {excel_row_index}, артикул: {article_str}")

        try:
            image_paths = image_utils.find_images_by_article(article_str, image_folder)
            if not image_paths:
                logger.warning(f"Для артикула '{article_str}' (строка {excel_row_index}) не найдено изображений в {image_folder}")
                continue

            image_path = image_paths[0] # Берем первое найденное
            logger.debug(f"Найдено изображение: {image_path}")

            # Оптимизация изображения
            base_name = os.path.basename(image_path)
            processed_image_path = os.path.join(image_output_folder, base_name)
            
            try:
                # Используем функцию optimize_image_for_excel С РАССЧИТАННЫМ ЛИМИТОМ
                optimized_buffer = image_utils.optimize_image_for_excel(
                    image_path, 
                    max_size_kb=target_kb_per_image, # <--- Используем рассчитанный лимит
                    quality=90, # <--- Стартуем с высоким качеством
                    min_quality=30 
                )
                
                with open(processed_image_path, 'wb') as f_out:
                    f_out.write(optimized_buffer.getvalue())
                logger.info(f"Изображение оптимизировано до {optimized_buffer.tell() / 1024:.1f}КБ: {processed_image_path}")
                
            except Exception as opt_e:
                 logger.warning(f"Не удалось оптимизировать изображение {image_path}: {opt_e}. Копируем оригинал.")
                 shutil.copy2(image_path, processed_image_path)


            # Вставка изображения в Excel
            if os.path.exists(processed_image_path):
                try:
                    # --- Aspect Ratio Calculation --- 
                    try:
                        with PILImage.open(processed_image_path) as pil_img:
                             original_width, original_height = pil_img.size
                    except Exception as pil_e:
                        logger.error(f"Не удалось прочитать размеры изображения {processed_image_path} с помощью PIL: {pil_e}")
                        original_width, original_height = column_width, row_height # Fallback

                    target_cell_width = column_width
                    target_cell_height = row_height
                    
                    img_aspect = original_width / original_height
                    cell_aspect = target_cell_width / target_cell_height

                    if img_aspect > cell_aspect: 
                        # Image is wider than cell -> scale by width
                        target_img_width = target_cell_width
                        target_img_height = target_img_width / img_aspect
                    else: 
                        # Image is taller than cell (or same aspect) -> scale by height
                        target_img_height = target_cell_height
                        target_img_width = target_img_height * img_aspect
                    # --- End Aspect Ratio Calculation ---

                    # Настройка высоты строки (если нужно)
                    if adjust_cell_size:
                         try:
                            # Устанавливаем высоту строки на основе РАССЧИТАННОЙ высоты изображения
                            # Добавляем небольшой отступ
                            effective_row_height = target_img_height + 10 # Add some padding
                            excel_utils.set_row_height(ws, excel_row_index, effective_row_height)
                            logger.debug(f"Установлена высота строки {excel_row_index} на {effective_row_height} пикс.")
                         except Exception as e:
                             logger.warning(f"Не удалось установить высоту строки {excel_row_index}: {e}")

                    # Вставка изображения
                    img = openpyxl.drawing.image.Image(processed_image_path)
                    
                    # Устанавливаем РАССЧИТАННЫЕ РАЗМЕРЫ изображения
                    img.width = target_img_width
                    img.height = target_img_height
                    logger.debug(f"Установка РАССЧИТАННЫХ размеров изображения: {img.width:.0f}x{img.height:.0f}")
                    
                    anchor_cell = f"{image_col_letter_excel}{excel_row_index}"
                    ws.add_image(img, anchor_cell)
                    images_inserted += 1
                    logger.info(f"Изображение {base_name} вставлено в ячейку {anchor_cell} с размерами {img.width:.0f}x{img.height:.0f}")

                except Exception as insert_e:
                     logger.error(f"Ошибка при вставке изображения {processed_image_path} в ячейку {anchor_cell}: {insert_e}", exc_info=True)
            else:
                 logger.error(f"Обработанное изображение не найдено: {processed_image_path}")


        except Exception as row_e:
            logger.error(f"Ошибка при обработке строки {excel_row_index} (артикул: {article_str}): {row_e}", exc_info=True)

    logger.info(f"Обработано строк: {rows_processed}, вставлено изображений: {images_inserted}")

    # --- Формирование имени выходного файла ---
    if output_path:
        final_output_file = output_path
        # Убедимся, что директория существует
        final_output_dir = os.path.dirname(final_output_file)
        if final_output_dir and not os.path.exists(final_output_dir):
            os.makedirs(final_output_dir)
    else:
        file_name = os.path.basename(file_path)
        name, ext = os.path.splitext(file_name)
        output_file_name = f"{name}_processed_{datetime.now().strftime('%Y%m%d%H%M%S')}{ext}"
        
        if output_folder:
             if not os.path.exists(output_folder):
                 os.makedirs(output_folder)
             final_output_file = os.path.join(output_folder, output_file_name)
        else:
            # Сохраняем в ту же папку, что и исходный файл
            final_output_file = os.path.join(os.path.dirname(file_path), output_file_name)

    logger.info(f"Сохранение результата в файл: {final_output_file}")

    # --- Сохранение результата ---
    try:
        wb.save(final_output_file)
        logger.info(f"Результат успешно сохранен в файл: {final_output_file}")
    except Exception as e:
        err_msg = f"Ошибка при сохранении результата в {final_output_file}: {e}"
        logger.error(err_msg, exc_info=True)
        # Попытка сохранить с другим именем во временную папку
        try:
            alt_output_folder = ensure_temp_dir("save_failed_")
            alt_output_file = os.path.join(alt_output_folder, os.path.basename(final_output_file))
            wb.save(alt_output_file)
            logger.warning(f"Не удалось сохранить в {final_output_file}. Результат сохранен в {alt_output_file}")
            final_output_file = alt_output_file # Возвращаем альтернативный путь
        except Exception as alt_e:
             err_msg = f"Ошибка при сохранении результата (попытка 2) в {alt_output_file}: {alt_e}"
             logger.error(err_msg, exc_info=True)
             raise RuntimeError(err_msg) from alt_e


    # --- Очистка временных изображений (если создавались) ---
    if temp_image_dir_created:
        try:
            shutil.rmtree(image_output_folder)
            logger.info(f"Удалена временная папка с изображениями: {image_output_folder}")
        except Exception as e:
            logger.warning(f"Не удалось удалить временную папку {image_output_folder}: {e}")

    # Возвращаем путь к файлу, DataFrame (он может быть изменен, хотя мы работали с wb) и кол-во картинок
    return final_output_file, df, images_inserted 