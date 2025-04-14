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
import re
import io

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
# <<< Тест 1: Проверяем, работает ли логгер этого модуля >>>
logger.critical("--- Logger for core.processor initialized ---") 

# <<< Constants for image fitting >>>
DEFAULT_CELL_WIDTH_PX = 150 
DEFAULT_CELL_HEIGHT_PX = 120
DEFAULT_IMG_QUALITY = 90
MIN_IMG_QUALITY = 1  # Снижено с 5% до 1% для еще большего сжатия
MIN_KB_PER_IMAGE = 10
MAX_KB_PER_IMAGE = 2048 # 2MB max per image, prevents extreme cases
SIZE_BUDGET_FACTOR = 0.85 # Use 85% of total size budget for images
ROW_HEIGHT_PADDING = 1 # Минимальный отступ для высоты строки
MIN_ASPECT_RATIO = 0.5 # Минимальное соотношение сторон (высота/ширина)
MAX_ASPECT_RATIO = 2.0 # Максимальное соотношение сторон (высота/ширина)
EXCEL_PX_TO_PT_RATIO = 0.75 # Коэффициент преобразования пикселей в пункты Excel

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
    article_col_name: str,
    image_folder: str,
    image_col_name: str = None,
    output_folder: str = None,
    max_total_file_size_mb: int = 100,
    progress_callback: callable = None,
    config: dict = None,
    header_row: int = 0,
    sheet_name: str = None,  # Добавляем параметр для имени листа
) -> Tuple[str, Optional[pd.DataFrame], int, Dict[str, List[str]], List[str]]:
    """
    Processes an Excel file by finding and inserting images based on article numbers.
    
    Args:
        file_path: Path to the Excel file
        article_col_name: Column name containing article numbers
        image_folder: Folder with images to search
        image_col_name: Column name where to insert images
        output_folder: Folder to save the result
        max_total_file_size_mb: Maximum file size in MB
        progress_callback: Function to call with progress updates
        config: Configuration parameters for image processing
        header_row: Row index containing header (0-based)
        sheet_name: Name of the sheet to process (if None, uses active sheet)
    
    Returns:
        Tuple with processing results
    """
    # <<< Используем print в stderr вместо logger >>>
    print(">>> ENTERING process_excel_file <<<\n", file=sys.stderr)
    sys.stderr.flush()
    
    print(f"[PROCESSOR] Начало обработки: {file_path}", file=sys.stderr)
    print(f"[PROCESSOR] Параметры: article_col={article_col_name}, img_folder={image_folder}, img_col={image_col_name}, max_total_mb={max_total_file_size_mb}, sheet_name={sheet_name}", file=sys.stderr)

    # --- Валидация входных данных ---
    # Проверка валидности обозначений колонок
    try:
        # Принимаем только буквенные обозначения колонок (A, B, C...)
        if not (article_col_name.isalpha() and image_col_name.isalpha()):
            err_msg = f"Неверное обозначение колонки: '{article_col_name}' или '{image_col_name}'. Используйте только буквенные обозначения (A, B, C...)"
            print(f"[PROCESSOR ERROR] {err_msg}", file=sys.stderr)
            raise ValueError(err_msg)
            
        article_col_idx = excel_utils.column_letter_to_index(article_col_name)
        image_col_idx = excel_utils.column_letter_to_index(image_col_name)
    except Exception as e:
        err_msg = f"Неверное обозначение колонки: '{article_col_name}' или '{image_col_name}'. Ошибка: {str(e)}"
        print(f"[PROCESSOR ERROR] {err_msg}", file=sys.stderr)
        raise ValueError(err_msg)
        
    if not os.path.exists(file_path):
        err_msg = f"Файл не найден: {file_path}"
        print(f"[PROCESSOR ERROR] {err_msg}", file=sys.stderr)
        raise FileNotFoundError(err_msg)
    
    if not os.path.exists(image_folder):
        err_msg = f"Папка с изображениями не найдена: {image_folder}"
        print(f"[PROCESSOR ERROR] {err_msg}", file=sys.stderr)
        raise FileNotFoundError(err_msg)

    # --- Чтение Excel ---
    try:
        # Если указан конкретный лист, читаем его
        if sheet_name:
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=0, engine='openpyxl')
            print(f"[PROCESSOR] Excel-файл прочитан в DataFrame (sheet={sheet_name}, header=0). Строк данных: {len(df)}", file=sys.stderr)
        else:
            df = pd.read_excel(file_path, header=0, engine='openpyxl') 
            print(f"[PROCESSOR] Excel-файл прочитан в DataFrame (header=0). Строк данных: {len(df)}", file=sys.stderr)
        
        # --- Загрузка книги openpyxl ---
        wb = openpyxl.load_workbook(file_path, read_only=False, keep_vba=False)
        try:
            # Проверяем наличие листов в книге
            if not wb.sheetnames:
                print("[PROCESSOR ERROR] В файле нет листов для обработки.", file=sys.stderr)
                raise ValueError("Excel-файл не содержит листов. Пожалуйста, выберите файл с данными.")
                
            # Фильтруем листы, исключая листы с макросами
            valid_sheets = [sheet_name for sheet_name in wb.sheetnames if not sheet_name.startswith('xl/macrosheets/')]
            if not valid_sheets:
                print("[PROCESSOR ERROR] В файле нет обычных листов, только макросы.", file=sys.stderr)
                raise ValueError("Внимание! Этот файл Excel содержит только макросы, а не обычные таблицы данных. Пожалуйста, выберите файл Excel с обычными листами, содержащими таблицы с артикулами и данными для обработки.")
            
            # Если указан лист, выбираем его, иначе используем активный
            if sheet_name:
                if sheet_name in wb.sheetnames:
                    ws = wb[sheet_name]
                    print(f"[PROCESSOR] Работаем с указанным листом: {sheet_name}", file=sys.stderr)
                else:
                    print(f"[PROCESSOR ERROR] Указанный лист {sheet_name} не найден в файле. Доступные листы: {wb.sheetnames}", file=sys.stderr)
                    raise ValueError(f"Лист '{sheet_name}' не найден в файле. Доступные листы: {wb.sheetnames}")
            else:
                # Используем первый лист
                ws = wb.active
                print(f"[PROCESSOR] Загружена рабочая книга, работаем с активным листом: {ws.title}", file=sys.stderr)
        except Exception as e:
            print(f"[PROCESSOR ERROR] Ошибка при выборе листа: {e}", file=sys.stderr)
            # Делаем сообщение об ошибке более понятным для пользователя
            if "'dict' object has no attribute 'shape'" in str(e):
                raise ValueError("Выбранный лист не содержит табличных данных. Пожалуйста, выберите лист с необходимыми данными.")
            else:
                raise ValueError(f"Ошибка при выборе листа: {e}")
        
    except Exception as e:
        err_msg = f"Ошибка при чтении Excel-файла: {e}"
        print(f"[PROCESSOR ERROR] {err_msg}", file=sys.stderr)
        # Выводим traceback в консоль
        import traceback
        traceback.print_exc(file=sys.stderr)
        
        # Делаем сообщение об ошибке более понятным для пользователя
        user_friendly_msg = err_msg
        if "'dict' object has no attribute 'shape'" in str(e):
            user_friendly_msg = "Выбранный лист не содержит табличных данных. Пожалуйста, выберите лист с необходимыми данными."
        elif "No sheet" in str(e) or "not found" in str(e):
            user_friendly_msg = "Указанный лист не найден в файле. Пожалуйста, выберите существующий лист."
        elif "Empty" in str(e) or "no data" in str(e):
            user_friendly_msg = "Выбранный лист не содержит данных. Пожалуйста, выберите лист с данными."
            
        raise RuntimeError(user_friendly_msg) from e

    if df.empty:
        err_msg = "Excel-файл не содержит данных"
        print(f"[PROCESSOR ERROR] {err_msg}", file=sys.stderr)
        raise ValueError(err_msg)

    # --- Проверка существования колонки артикулов ---
    # Преобразуем букву колонки в индекс
    article_col_idx = excel_utils.column_letter_to_index(article_col_name)
    article_col_name = df.columns[article_col_idx]
    articles = df[article_col_name].astype(str).tolist()
    print(f"[PROCESSOR] Получено {len(articles)} артикулов из колонки {article_col_name}", file=sys.stderr)
    
    if article_col_name not in df.columns:
        err_msg = f"Колонка с артикулами '{article_col_name}' не найдена в файле. Доступные колонки: {list(df.columns)}"
        print(f"[PROCESSOR ERROR] {err_msg}", file=sys.stderr)
        raise ValueError(err_msg)
    print(f"[PROCESSOR] Колонка с артикулами: '{article_col_name}'", file=sys.stderr)

    # --- Определение КОЛИЧЕСТВА строк с НЕНУЛЕВЫМИ артикулами для расчета лимита ---
    # Считаем строки, где артикул не пустой
    non_empty_article_rows = df[article_col_name].notna() & (df[article_col_name].astype(str).str.strip() != '')
    article_count = non_empty_article_rows.sum()
    
    if article_count == 0:
        article_count = 1 # Избегаем деления на ноль
        print("[PROCESSOR WARNING] Не найдено строк с непустыми артикулами для расчета лимита размера изображения. Используется значение по умолчанию.", file=sys.stderr)
    else:
        print(f"[PROCESSOR] Найдено {article_count} строк с непустыми артикулами.", file=sys.stderr)
        
    # --- Расчет лимита размера на одно изображение ---
    image_size_budget_mb = max_total_file_size_mb * SIZE_BUDGET_FACTOR
    target_kb_per_image = (image_size_budget_mb * 1024) / article_count if article_count > 0 else MAX_KB_PER_IMAGE
    target_kb_per_image = max(MIN_KB_PER_IMAGE, min(target_kb_per_image, MAX_KB_PER_IMAGE)) 
    print(f"[PROCESSOR] Расчетный лимит размера на изображение: {target_kb_per_image:.1f} КБ", file=sys.stderr)

    # --- Подготовка папки для обработанных изображений ---
    temp_image_dir_created = False
    if not image_folder:
        image_folder = ensure_temp_dir("processed_images_")
        temp_image_dir_created = True
        print(f"[PROCESSOR] Создана временная директория для обработанных изображений: {image_folder}", file=sys.stderr)
    elif not os.path.exists(image_folder):
         os.makedirs(image_folder)
         print(f"[PROCESSOR] Создана папка для обработанных изображений: {image_folder}", file=sys.stderr)


    # --- Подготовка к вставке изображений ---
    try:
        # НАПРЯМУЮ ИСПОЛЬЗУЕМ УКАЗАННУЮ БУКВУ КОЛОНКИ
        image_col_letter_excel = image_col_name
        print(f"[PROCESSOR] Изображения будут вставляться в колонку: '{image_col_letter_excel}'", file=sys.stderr)
    except Exception as e:
         err_msg = f"Ошибка при подготовке колонки для изображений ('{article_col_name}'): {e}"
         print(f"[PROCESSOR ERROR] {err_msg}", file=sys.stderr)
         import traceback
         traceback.print_exc(file=sys.stderr)
         raise RuntimeError(err_msg) from e

    # --- Настройка ШИРИНЫ КОЛОНКИ (используем константу) ---
    try:
        excel_utils.set_column_width(ws, image_col_letter_excel, DEFAULT_CELL_WIDTH_PX / 7) # Approx conversion
        print(f"[PROCESSOR] Установлена ширина столбца {image_col_letter_excel} на {DEFAULT_CELL_WIDTH_PX} пикс.", file=sys.stderr)
    except Exception as e:
        print(f"[PROCESSOR WARNING] Не удалось установить ширину столбца {image_col_letter_excel}: {e}", file=sys.stderr)


    # --- Обработка строк и вставка изображений ---
    images_inserted = 0
    rows_processed = 0
    total_processed_image_size_kb = 0
    
    print("[PROCESSOR] --- Начало итерации по строкам DataFrame ---", file=sys.stderr)
    
    # Initialize collections for reporting
    not_found_articles = []
    multiple_images_found = {}
    
    # Переменная для сохранения найденного качества сжатия первого изображения
    successful_quality = DEFAULT_IMG_QUALITY
    # Флаг для определения, было ли обработано первое изображение
    first_image_processed = False
    
    # --- Итерация по строкам ---
    # Начинаем с 1-й строки (после заголовка), но номер строки будет в формате Excel (от 1)
    for df_index, row in df.iterrows():
        excel_row_index = df_index + 1 + 1  # +1 потому что в Excel нумерация с 1, а header_row смещение заголовка
        article_str = ""  # Инициализируем article_str пустой строкой
        
        try:
            # Получаем значение артикула
            article_str = str(row[article_col_name]) if row[article_col_name] is not None else ""
            
            print(f"[PROCESSOR] Обработка строки {excel_row_index}, артикул: '{article_str}'", file=sys.stderr)
            
            if pd.isna(row[article_col_name]) or article_str.strip() == "":
                print(f"[PROCESSOR]   Пустой артикул в строке {excel_row_index}, пропускаем", file=sys.stderr)
                continue
            
            # Find all images for this article
            # Определяем поддерживаемые расширения
            supported_extensions = ('.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.webp')
            all_image_paths = image_utils.find_images_by_article_name(
                article_str, 
                image_folder,
                supported_extensions,
                search_recursively=True  # Enable recursive search for images
            )
            
            # If no images found, record and continue
            if not all_image_paths:
                print(f"[PROCESSOR WARNING]   Для артикула '{article_str}' (строка {excel_row_index}) не найдено изображений. Пропускаем.", file=sys.stderr)
                # Добавляем артикул в список не найденных
                not_found_articles.append(article_str)
                continue
            
            # If multiple images found, record for report
            if len(all_image_paths) > 1:
                print(f"[PROCESSOR INFO]   Найдено несколько изображений для артикула '{article_str}': {len(all_image_paths)}", file=sys.stderr)
                multiple_images_found[article_str] = all_image_paths
                # Still proceed with the first image
            
            image_path = all_image_paths[0]
            print(f"[PROCESSOR]   Выбрано первое найденное изображение: {image_path}", file=sys.stderr)

            # 1. ОПТИМИЗАЦИЯ ИЗОБРАЖЕНИЯ
            optimized_buffer = None
            print(f"[PROCESSOR]   Вызов optimize_image_for_excel для {image_path} с лимитом {target_kb_per_image:.1f} КБ", file=sys.stderr)
            
            try:
                # Для САМОГО ПЕРВОГО изображения во ВСЕМ файле выполняем детальную оптимизацию
                if not first_image_processed:
                    # Для первого изображения выполняем полный поиск оптимального качества
                    print(f"[PROCESSOR]   ПЕРВОЕ ИЗОБРАЖЕНИЕ В ФАЙЛЕ: поиск оптимального качества от {DEFAULT_IMG_QUALITY}% до {MIN_IMG_QUALITY}%", file=sys.stderr)
                    optimized_buffer = image_utils.optimize_image_for_excel(
                        image_path, 
                        target_size_kb=target_kb_per_image,
                        quality=DEFAULT_IMG_QUALITY,
                        min_quality=MIN_IMG_QUALITY    
                    )
                    # После первого изображения больше не выполняем детальную оптимизацию
                    first_image_processed = True
                    
                    # После успешной оптимизации первого изображения определяем качество,
                    # которое лучше всего подошло
                    if optimized_buffer and optimized_buffer.getbuffer().nbytes > 0:
                        # Мы можем полагаться только на лог-сообщения для определения качества
                        # Находим в логах последнюю строку с "Итоговое качество сжатия"
                        try:
                            # Сначала попытаемся прочитать текущий stderr буфер, если это возможно
                            quality_found = False
                            
                            # Попытка найти в логах
                            with open("logs/app_latest.log", "r", encoding="utf-8") as log_file:
                                log_lines = log_file.readlines()
                                # Ищем два типа строк с информацией о качестве
                                quality_pattern1 = re.compile(r'\[optimize_excel\] Итоговое качество сжатия: (\d+)%')
                                quality_pattern2 = re.compile(r'-> Успех! Размер .* c качеством (\d+)')
                                # Добавляем поиск специального маркера
                                quality_marker = re.compile(r'\[QUALITY_MARKER\] НАЙДЕНО_КАЧЕСТВО_ДЛЯ_ИЗОБРАЖЕНИЯ: (\d+)')
                                
                                # Сначала ищем в последних 100 строках (самая свежая информация)
                                recent_lines = log_lines[-100:] if len(log_lines) > 100 else log_lines
                                
                                for line in reversed(recent_lines):
                                    # Приоритетно ищем специальный маркер
                                    match = quality_marker.search(line)
                                    if match:
                                        found_quality = int(match.group(1))
                                        successful_quality = found_quality
                                        print(f"[PROCESSOR]   НАЙДЕН СПЕЦИАЛЬНЫЙ МАРКЕР КАЧЕСТВА: {successful_quality}%", file=sys.stderr)
                                        quality_found = True
                                        break
                                        
                                    match = quality_pattern1.search(line)
                                    if match:
                                        found_quality = int(match.group(1))
                                        successful_quality = found_quality
                                        print(f"[PROCESSOR]   Определено оптимальное качество из последних логов (pattern 1): {successful_quality}%", file=sys.stderr)
                                        quality_found = True
                                        break
                                        
                                    match = quality_pattern2.search(line)
                                    if match:
                                        found_quality = int(match.group(1))
                                        successful_quality = found_quality
                                        print(f"[PROCESSOR]   Определено оптимальное качество из последних логов (pattern 2): {successful_quality}%", file=sys.stderr)
                                        quality_found = True
                                        break
                                
                                # Если не нашли в последних строках, ищем во всем файле
                                if not quality_found:
                                    print(f"[PROCESSOR]   Качество не найдено в последних строках логов, ищем во всем файле...", file=sys.stderr)
                                    for line in reversed(log_lines):
                                        match = quality_pattern1.search(line)
                                        if match:
                                            found_quality = int(match.group(1))
                                            successful_quality = found_quality
                                            print(f"[PROCESSOR]   Определено оптимальное качество из всех логов: {successful_quality}%", file=sys.stderr)
                                            quality_found = True
                                            break
                                
                                # Если все еще не нашли, попробуем прочитать из временного файла
                                if not quality_found:
                                    try:
                                        temp_quality_file = os.path.join(tempfile.gettempdir(), "last_image_quality.txt")
                                        if os.path.exists(temp_quality_file):
                                            with open(temp_quality_file, "r") as qf:
                                                file_quality = int(qf.read().strip())
                                                successful_quality = file_quality
                                                print(f"[PROCESSOR]   Определено качество из временного файла: {successful_quality}%", file=sys.stderr)
                                                quality_found = True
                                    except Exception as file_err:
                                        print(f"[PROCESSOR]   Ошибка чтения временного файла с качеством: {file_err}", file=sys.stderr)
                                
                                # Если все еще не нашли, используем мин. значение
                                if not quality_found:
                                    successful_quality = MIN_IMG_QUALITY  # Прямое использование мин. качества
                                    print(f"[PROCESSOR]   Качество не найдено в логах. Используем минимальное значение: {successful_quality}%", file=sys.stderr)
                        except Exception as log_e:
                            print(f"[PROCESSOR WARNING]   Ошибка при чтении лог-файла: {log_e}. Используем минимальное качество.", file=sys.stderr)
                            successful_quality = MIN_IMG_QUALITY  # Минимальное качество без добавления
                    
                    # Сообщаем о выбранном качестве для всех последующих изображений
                    print(f"[PROCESSOR]   ВАЖНО: Для всех последующих изображений будет использовано качество {successful_quality}%", file=sys.stderr)
                else:
                    # Для всех последующих изображений используем найденное качество первого изображения
                    print(f"[PROCESSOR]   Используем найденное качество {successful_quality}% для изображения {image_path}", file=sys.stderr)
                    optimized_buffer = image_utils.optimize_image_for_excel(
                        image_path, 
                        target_size_kb=target_kb_per_image,
                        quality=successful_quality,  # Начальное = найденное качество 
                        min_quality=successful_quality  # Мин. качество = найденное качество (без итераций)
                    )
                
                if optimized_buffer and optimized_buffer.getbuffer().nbytes > 0:
                    buffer_size_kb = optimized_buffer.tell() / 1024
                    print(f"[PROCESSOR]   Оптимизация успешна. Размер буфера: {buffer_size_kb:.1f} КБ", file=sys.stderr)
                    current_image_size_kb = buffer_size_kb
                    total_processed_image_size_kb += current_image_size_kb
                    
                    # Дополнительная проверка буфера - убеждаемся, что это действительно изображение
                    optimized_buffer.seek(0)
                    try:
                        verification_img = PILImage.open(optimized_buffer)
                        img_format = verification_img.format
                        img_width_px, img_height_px = verification_img.size
                        print(f"[PROCESSOR]   ВЕРИФИКАЦИЯ: буфер содержит изображение формата {img_format}, {img_width_px}x{img_height_px}", file=sys.stderr)
                        
                        # Создаем временную копию буфера для сохранения в файл (для отладки)
                        try:
                            debug_copy = io.BytesIO(optimized_buffer.getvalue())
                            temp_debug_path = os.path.join(tempfile.gettempdir(), f"debug_image_{time.time()}.jpg")
                            with open(temp_debug_path, "wb") as debug_file:
                                debug_file.write(debug_copy.getvalue())
                            print(f"[PROCESSOR]   Создана отладочная копия изображения: {temp_debug_path}", file=sys.stderr)
                        except Exception as debug_e:
                            print(f"[PROCESSOR]   Примечание: не удалось создать отладочную копию: {debug_e}", file=sys.stderr)
                        
                        # Сбрасываем указатель в начало буфера после верификации
                        optimized_buffer.seek(0)
                    except Exception as verify_e:
                        print(f"[PROCESSOR] ОШИБКА ВЕРИФИКАЦИИ: Буфер не содержит корректного изображения: {verify_e}", file=sys.stderr)
                        # Пробуем сохранить проблемный буфер для анализа
                        try:
                            error_path = os.path.join(tempfile.gettempdir(), f"error_buffer_{time.time()}.bin")
                            with open(error_path, "wb") as error_file:
                                error_file.write(optimized_buffer.getvalue())
                            print(f"[PROCESSOR]   Сохранён проблемный буфер для анализа: {error_path}", file=sys.stderr)
                        except Exception as err_save_e:
                            print(f"[PROCESSOR]   Не удалось сохранить проблемный буфер: {err_save_e}", file=sys.stderr)
                            
                        # Если буфер некорректен, пробуем загрузить оригинальное изображение
                        try:
                            print(f"[PROCESSOR]   Пробуем загрузить оригинальное изображение как резервный вариант", file=sys.stderr)
                            with open(image_path, "rb") as original_file:
                                optimized_buffer = io.BytesIO(original_file.read())
                            print(f"[PROCESSOR]   Загружено оригинальное изображение размером {optimized_buffer.getbuffer().nbytes / 1024:.1f} КБ", file=sys.stderr)
                            optimized_buffer.seek(0)
                            verification_img = PILImage.open(optimized_buffer)
                            img_width_px, img_height_px = verification_img.size
                        except Exception as orig_load_e:
                            print(f"[PROCESSOR] КРИТИЧЕСКАЯ ОШИБКА: Не удалось загрузить даже оригинальное изображение: {orig_load_e}", file=sys.stderr)
                            continue  # Пропускаем эту итерацию
                    
                    # Получаем размеры изображения напрямую из буфера
                    try:
                        optimized_buffer.seek(0)
                        img = PILImage.open(optimized_buffer)
                        img_width_px, img_height_px = img.size
                        print(f"[PROCESSOR]     Получены размеры из буфера: {img_width_px}x{img_height_px}", file=sys.stderr)
                    except Exception as dim_e:
                        print(f"[PROCESSOR] WARNING: Не удалось получить размеры изображения из буфера: {dim_e}", file=sys.stderr)
                        img_width_px, img_height_px = DEFAULT_CELL_WIDTH_PX, DEFAULT_CELL_HEIGHT_PX
                    
                    # Расчет пропорций изображения для правильной высоты строки и сохранения пропорций
                    original_aspect_ratio = img_height_px / img_width_px if img_width_px > 0 else 1.0
                    
                    # Ограничиваем соотношение сторон, чтобы избежать слишком вытянутых изображений
                    aspect_ratio = max(MIN_ASPECT_RATIO, min(original_aspect_ratio, MAX_ASPECT_RATIO))
                    print(f"[PROCESSOR]     Соотношение сторон: оригинальное {original_aspect_ratio:.2f}, используемое {aspect_ratio:.2f}", file=sys.stderr)
                    
                    # Рассчитываем новую высоту изображения, сохраняя пропорции
                    scaled_img_height = int(DEFAULT_CELL_WIDTH_PX * aspect_ratio)
                    
                    # Рассчитываем высоту строки с учетом отступа (3-5 пикселей)
                    target_row_height_pt = (scaled_img_height + ROW_HEIGHT_PADDING) * EXCEL_PX_TO_PT_RATIO
                    
                    print(f"[PROCESSOR]     Вызов set_row_height для строки {excel_row_index} на {target_row_height_pt:.1f} pt (ширина: {DEFAULT_CELL_WIDTH_PX}px, высота: {scaled_img_height:.1f}px)", file=sys.stderr)
                    excel_utils.set_row_height(ws, excel_row_index, target_row_height_pt)
                    
                    cell_address = image_col_letter_excel + str(excel_row_index)
                    print(f"[PROCESSOR]     Вызов insert_image_from_buffer в ячейку {cell_address} (width: {DEFAULT_CELL_WIDTH_PX}, height: {scaled_img_height})", file=sys.stderr)
                    
                    # Сбросить указатель в начало буфера
                    optimized_buffer.seek(0)
                    
                    # Вставляем изображение прямо из буфера
                    excel_utils.insert_image_from_buffer(
                        ws,
                        optimized_buffer,
                        anchor_cell=cell_address,
                        width=DEFAULT_CELL_WIDTH_PX,
                        height=scaled_img_height,
                        preserve_aspect_ratio=True
                    )
                    
                    # Добавляем дополнительную проверку для подтверждения успешной вставки
                    print(f"[PROCESSOR]     Изображение вставлено в Excel с размерами {DEFAULT_CELL_WIDTH_PX}x{scaled_img_height} пикселей", file=sys.stderr)
                    
                    images_inserted += 1
                    print(f"[PROCESSOR]   Изображение успешно вставлено в ячейку {cell_address}", file=sys.stderr)
                else:
                    print(f"[PROCESSOR WARNING]   Оптимизация {image_path} вернула пустой или нулевой буфер. Пропускаем вставку.", file=sys.stderr)
                    continue

            except Exception as opt_e:
                 print(f"[PROCESSOR WARNING]   Ошибка при оптимизации {image_path}: {opt_e}. Пропускаем вставку.", file=sys.stderr)
                 import traceback
                 traceback.print_exc(file=sys.stderr)
                 continue

        except Exception as row_e:
            print(f"[PROCESSOR ERROR]   Непредвиденная ошибка при обработке строки {excel_row_index} (артикул: {article_str}): {row_e}", file=sys.stderr)
            import traceback
            traceback.print_exc(file=sys.stderr)

    print("[PROCESSOR] --- Завершение итерации по строкам --- ", file=sys.stderr)
    # --- Предупреждение о размере ---
    total_processed_image_size_mb = total_processed_image_size_kb / 1024
    print(f"[PROCESSOR] Суммарный размер оптимизированных изображений: {total_processed_image_size_mb:.2f} МБ.", file=sys.stderr)
    if total_processed_image_size_mb > image_size_budget_mb:
        print(f"[PROCESSOR WARNING] Суммарный размер оптимизированных изображений ({total_processed_image_size_mb:.2f} МБ) превышает расчетный бюджет ({image_size_budget_mb:.2f} МБ).", file=sys.stderr)

    # --- Формирование имени выходного файла ---
    output_file_name = f"processed_{datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx"
    
    # Проверяем и создаем выходную папку
    try:
        if output_folder:
            # Проверяем существование и доступность для записи
            if not os.path.exists(output_folder):
                try:
                    os.makedirs(output_folder, exist_ok=True)
                    print(f"[PROCESSOR] Создана папка для результатов: {output_folder}", file=sys.stderr)
                except Exception as mkdir_e:
                    print(f"[PROCESSOR WARNING] Не удалось создать папку {output_folder}: {mkdir_e}", file=sys.stderr)
                    output_folder = None  # Сбрасываем, чтобы использовать альтернативный путь
            elif not os.access(output_folder, os.W_OK):
                print(f"[PROCESSOR WARNING] Нет прав на запись в папку {output_folder}. Будет использован альтернативный путь.", file=sys.stderr)
                output_folder = None
        
        # Если папка не указана или недоступна, используем рабочую папку или временную
        if not output_folder:
            # Сначала пробуем папку с исходным файлом
            src_dir = os.path.dirname(file_path)
            if os.path.exists(src_dir) and os.access(src_dir, os.W_OK):
                output_folder = src_dir
                print(f"[PROCESSOR] Используем папку исходного файла для результатов: {output_folder}", file=sys.stderr)
            else:
                # Иначе используем временную директорию
                output_folder = ensure_temp_dir("excel_results_")
                print(f"[PROCESSOR] Используем временную папку для результатов: {output_folder}", file=sys.stderr)
        
        # Формируем полный путь к выходному файлу
        output_file_path = os.path.join(output_folder, output_file_name)
        print(f"[PROCESSOR] Полный путь к выходному файлу: {output_file_path}", file=sys.stderr)
    except Exception as path_e:
        print(f"[PROCESSOR WARNING] Ошибка при подготовке пути для сохранения: {path_e}", file=sys.stderr)
        # В случае ошибки используем запасной вариант во временной директории
        output_folder = ensure_temp_dir("excel_results_emergency_")
        output_file_path = os.path.join(output_folder, output_file_name)
        print(f"[PROCESSOR] Аварийный путь к выходному файлу: {output_file_path}", file=sys.stderr)

    print(f"[PROCESSOR] Сохранение результата в файл: {output_file_path}", file=sys.stderr)

    # --- Подготовка к сохранению ---
    # Преобразуем путь в абсолютный, если это еще не сделано
    if not os.path.isabs(output_file_path):
        output_file_path = os.path.abspath(output_file_path)
        print(f"[PROCESSOR] Преобразован в абсолютный путь: {output_file_path}", file=sys.stderr)
    
    # Печатаем информацию о директории и правах доступа
    output_dir = os.path.dirname(output_file_path)
    dir_exists = os.path.exists(output_dir)
    dir_writable = os.access(output_dir, os.W_OK) if dir_exists else False
    
    print(f"[PROCESSOR] Информация о директории: {output_dir}", file=sys.stderr)
    print(f"[PROCESSOR]   - Существует: {dir_exists}", file=sys.stderr)
    print(f"[PROCESSOR]   - Доступна для записи: {dir_writable}", file=sys.stderr)
    
    # Если директория не существует или нет прав на запись, попробуем создать другую
    if not dir_exists or not dir_writable:
        fallback_dir = ensure_temp_dir("excel_fallback_")
        fallback_path = os.path.join(fallback_dir, os.path.basename(output_file_path))
        print(f"[PROCESSOR] Используем запасной путь из-за проблем с основным: {fallback_path}", file=sys.stderr)
        output_file_path = fallback_path
    
    # --- Сохранение результата ---
    try:
        wb.save(output_file_path)
        print(f"[PROCESSOR] Результат успешно сохранен в файл: {output_file_path}", file=sys.stderr)
        
        # Очистка временных файлов изображений после успешного сохранения
        temp_files_cleaned = 0
        for sheet in wb.worksheets:
            if hasattr(sheet, '_temp_image_files'):
                for temp_path in sheet._temp_image_files:
                    try:
                        if os.path.exists(temp_path):
                            os.unlink(temp_path)
                            temp_files_cleaned += 1
                    except Exception as e:
                        print(f"[PROCESSOR WARNING] Не удалось удалить временный файл {temp_path}: {e}", file=sys.stderr)
                sheet._temp_image_files = []
        print(f"[PROCESSOR] Очищено {temp_files_cleaned} временных файлов изображений.", file=sys.stderr)
        
        # Добавляем проверку существования файла
        if os.path.exists(output_file_path):
            file_size = os.path.getsize(output_file_path)
            print(f"[PROCESSOR] Проверка файла: {output_file_path} существует, размер: {file_size/1024:.2f} КБ", file=sys.stderr)
        else:
            print(f"[PROCESSOR ERROR] КРИТИЧЕСКАЯ ОШИБКА: Файл {output_file_path} не был создан после вызова save()", file=sys.stderr)
            # Попробуем использовать абсолютный путь
            abs_path = os.path.abspath(output_file_path)
            print(f"[PROCESSOR] Попытка использовать абсолютный путь: {abs_path}", file=sys.stderr)
            wb.save(abs_path)
            output_file_path = abs_path
            
            if os.path.exists(output_file_path):
                file_size = os.path.getsize(output_file_path)
                print(f"[PROCESSOR] Успех при использовании абсолютного пути! Файл {output_file_path} создан, размер: {file_size/1024:.2f} КБ", file=sys.stderr)
            else:
                print(f"[PROCESSOR ERROR] КРИТИЧЕСКАЯ ОШИБКА: Файл {output_file_path} не был создан даже при использовании абсолютного пути", file=sys.stderr)
    except Exception as e:
        err_msg = f"Ошибка при сохранении результата в {output_file_path}: {e}"
        print(f"[PROCESSOR ERROR] {err_msg}", file=sys.stderr)
        # Попытка сохранить с другим именем во временную папку
        try:
            alt_output_folder = ensure_temp_dir("save_failed_")
            alt_output_file = os.path.join(alt_output_folder, os.path.basename(output_file_path))
            wb.save(alt_output_file)
            print(f"[PROCESSOR WARNING] Не удалось сохранить в {output_file_path}. Результат сохранен в {alt_output_file}", file=sys.stderr)
            
            # Очистка временных файлов изображений после успешного сохранения
            temp_files_cleaned = 0
            for sheet in wb.worksheets:
                if hasattr(sheet, '_temp_image_files'):
                    for temp_path in sheet._temp_image_files:
                        try:
                            if os.path.exists(temp_path):
                                os.unlink(temp_path)
                                temp_files_cleaned += 1
                        except Exception as e:
                            print(f"[PROCESSOR WARNING] Не удалось удалить временный файл {temp_path}: {e}", file=sys.stderr)
                    sheet._temp_image_files = []
            print(f"[PROCESSOR] Очищено {temp_files_cleaned} временных файлов изображений.", file=sys.stderr)
            
            output_file_path = alt_output_file # Возвращаем альтернативный путь
            
            # Проверяем существование альтернативного файла
            if os.path.exists(output_file_path):
                file_size = os.path.getsize(output_file_path)
                print(f"[PROCESSOR] Проверка альтернативного файла: {output_file_path} существует, размер: {file_size/1024:.2f} КБ", file=sys.stderr)
            else:
                print(f"[PROCESSOR ERROR] КРИТИЧЕСКАЯ ОШИБКА: Альтернативный файл {output_file_path} не был создан после вызова save()", file=sys.stderr)
        except Exception as alt_e:
             err_msg = f"Ошибка при сохранении результата (попытка 2) в {alt_output_file}: {alt_e}"
             print(f"[PROCESSOR ERROR] {err_msg}", file=sys.stderr)
             import traceback
             traceback.print_exc(file=sys.stderr)
             raise RuntimeError(err_msg) from alt_e


    # --- Очистка временных изображений (если создавались) ---
    if temp_image_dir_created:
        try:
            shutil.rmtree(image_folder)
            print(f"[PROCESSOR] Удалена временная папка с изображениями: {image_folder}", file=sys.stderr)
        except Exception as e:
            print(f"[PROCESSOR WARNING] Не удалось удалить временную папку {image_folder}: {e}", file=sys.stderr)

    print(f"[PROCESSOR] Обработка завершена. Вставлено изображений: {images_inserted}", file=sys.stderr)
    
    # Печатаем статистику о не найденных артикулах и артикулах с множественными изображениями
    if not_found_articles:
        print(f"[PROCESSOR REPORT] Не найдены изображения для {len(not_found_articles)} артикулов", file=sys.stderr)
    
    if multiple_images_found:
        print(f"[PROCESSOR REPORT] Найдено несколько вариантов изображений для {len(multiple_images_found)} артикулов", file=sys.stderr)
    
    # Возвращаем кортеж с результатами в порядке, ожидаемом в app.py:
    # output_file_path, result_df, images_inserted, multiple_images_found, not_found_articles
    return output_file_path, None, images_inserted, multiple_images_found, not_found_articles