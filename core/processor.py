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
MIN_IMG_QUALITY = 5  # Снижено с 30% до 5% для большего сжатия
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
    article_col_letter: str,  # Буква колонки с артикулами ('A')
    image_col_letter: str,  # Буква колонки для вставки изображений ('B')
    image_folder: str,
    image_output_folder: Optional[str] = None,
    output_path: Optional[str] = None,
    output_folder: Optional[str] = None,
    max_total_file_size_mb: int = 20,
    supported_formats: Tuple[str, ...] = ('.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff', '.webp')
) -> Tuple[str, pd.DataFrame, int]:
    # <<< Используем print в stderr вместо logger >>>
    print(">>> ENTERING process_excel_file <<<\n", file=sys.stderr)
    sys.stderr.flush()
    
    print(f"[PROCESSOR] Начало обработки: {file_path}", file=sys.stderr)
    print(f"[PROCESSOR] Параметры: article_col={article_col_letter}, img_folder={image_folder}, img_col={image_col_letter}, max_total_mb={max_total_file_size_mb}", file=sys.stderr)

    # --- Валидация входных данных ---
    # Проверка валидности обозначений колонок
    try:
        # Принимаем только буквенные обозначения колонок (A, B, C...)
        if not (article_col_letter.isalpha() and image_col_letter.isalpha()):
            err_msg = f"Неверное обозначение колонки: '{article_col_letter}' или '{image_col_letter}'. Используйте только буквенные обозначения (A, B, C...)"
            print(f"[PROCESSOR ERROR] {err_msg}", file=sys.stderr)
            raise ValueError(err_msg)
            
        article_col_idx = excel_utils.column_letter_to_index(article_col_letter)
        image_col_idx = excel_utils.column_letter_to_index(image_col_letter)
    except Exception as e:
        err_msg = f"Неверное обозначение колонки: '{article_col_letter}' или '{image_col_letter}'. Ошибка: {str(e)}"
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
                
            # Используем первый лист
            ws = wb.active
            print(f"[PROCESSOR] Загружена рабочая книга, работаем с листом: {ws.title}", file=sys.stderr)
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
    article_col_idx = excel_utils.column_letter_to_index(article_col_letter)
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
    if not image_output_folder:
        image_output_folder = ensure_temp_dir("processed_images_")
        temp_image_dir_created = True
        print(f"[PROCESSOR] Создана временная директория для обработанных изображений: {image_output_folder}", file=sys.stderr)
    elif not os.path.exists(image_output_folder):
         os.makedirs(image_output_folder)
         print(f"[PROCESSOR] Создана папка для обработанных изображений: {image_output_folder}", file=sys.stderr)


    # --- Подготовка к вставке изображений ---
    try:
        # НАПРЯМУЮ ИСПОЛЬЗУЕМ УКАЗАННУЮ БУКВУ КОЛОНКИ
        image_col_letter_excel = image_col_letter
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
    
    # Переменная для сохранения найденного качества сжатия первого изображения
    successful_quality = DEFAULT_IMG_QUALITY
    # Флаг для определения, было ли обработано первое изображение
    first_image_processed = False
    
    # --- Итерация по строкам ---
    # Начинаем с 1-й строки (после заголовка), но номер строки будет в формате Excel (от 1)
    for df_index, row in df.iterrows():
        excel_row_index = df_index + 1 + 1  # +1 потому что в Excel нумерация с 1, а header_row смещение заголовка
        
        try:
            # Получаем значение артикула
            article_str = str(row[article_col_name]) if row[article_col_name] is not None else ""
            
            print(f"[PROCESSOR] Обработка строки {excel_row_index}, артикул: '{article_str}'", file=sys.stderr)
            
            if pd.isna(row[article_col_name]) or article_str.strip() == "":
                print(f"[PROCESSOR]   Пустой артикул в строке {excel_row_index}, пропускаем", file=sys.stderr)
                continue
            
            # <<< ПОИСК ИЗОБРАЖЕНИЙ >>>
            # Вместо вызова find_images_by_article встраиваем его код непосредственно:
            # --- Начало кода из find_images_by_article ---
            found_image_paths_debug = []
            print(f"[PROCESSOR]   Нормализация артикула '{article_str}'", file=sys.stderr)
            normalized_article_to_find_debug = image_utils.normalize_article(article_str)
            print(f"[PROCESSOR]   Нормализованный артикул: '{normalized_article_to_find_debug}'", file=sys.stderr)
            if os.path.isdir(image_folder) and normalized_article_to_find_debug:
                try:
                    all_files_in_dir_debug = os.listdir(image_folder)
                    print(f"[PROCESSOR]   Найдено файлов в папке: {len(all_files_in_dir_debug)}", file=sys.stderr)
                    normalized_name_to_original_path_debug: Dict[str, str] = {}
                    for filename_debug in all_files_in_dir_debug:
                        full_path_debug = os.path.join(image_folder, filename_debug)
                        if os.path.isfile(full_path_debug):
                            file_ext_lower_debug = os.path.splitext(filename_debug)[1].lower()
                            if file_ext_lower_debug in supported_formats:
                                name_without_ext_debug = os.path.splitext(filename_debug)[0]
                                normalized_name_debug = image_utils.normalize_article(name_without_ext_debug)
                                if normalized_name_debug:
                                     normalized_name_to_original_path_debug[normalized_name_debug] = full_path_debug
                    # Точное совпадение
                    if normalized_article_to_find_debug in normalized_name_to_original_path_debug:
                        exact_match_path_debug = normalized_name_to_original_path_debug[normalized_article_to_find_debug]
                        print(f"[PROCESSOR]   Найдено ТОЧНОЕ совпадение: {exact_match_path_debug}", file=sys.stderr)
                        if os.access(exact_match_path_debug, os.R_OK):
                            found_image_paths_debug.append(exact_match_path_debug)
                    # Частичное (если точного нет)
                    elif not found_image_paths_debug:
                         for norm_name_debug, original_path_debug in normalized_name_to_original_path_debug.items():
                            if normalized_article_to_find_debug in norm_name_debug or norm_name_debug in normalized_article_to_find_debug:
                                print(f"[PROCESSOR]   Найдено ЧАСТИЧНОЕ совпадение: {original_path_debug}", file=sys.stderr)
                                if os.access(original_path_debug, os.R_OK):
                                    found_image_paths_debug.append(original_path_debug)
                                    break # Берем первое частичное
                except Exception as find_e:
                    print(f"[PROCESSOR]   Ошибка поиска: {find_e}", file=sys.stderr)
            # --- Конец кода из find_images_by_article --- 
            image_paths = found_image_paths_debug
            print(f"[PROCESSOR]   Результат поиска изображений: {image_paths}", file=sys.stderr)
            
            if not image_paths:
                print(f"[PROCESSOR WARNING]   Для артикула '{article_str}' (строка {excel_row_index}) не найдено изображений. Пропускаем.", file=sys.stderr)
                continue

            image_path = image_paths[0]
            print(f"[PROCESSOR]   Выбрано первое найденное изображение: {image_path}", file=sys.stderr)

            # 1. ОПТИМИЗАЦИЯ ИЗОБРАЖЕНИЯ
            optimized_buffer = None
            optimized_image_path_for_excel = None
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
                    base_name, _ = os.path.splitext(os.path.basename(image_path))
                    temp_img_filename = f"optimized_{excel_row_index}_{base_name}.jpg"
                    optimized_image_path_for_excel = os.path.join(image_output_folder, temp_img_filename)
                    with open(optimized_image_path_for_excel, 'wb') as f_out:
                        f_out.write(optimized_buffer.getvalue())
                    current_image_size_kb = buffer_size_kb
                    total_processed_image_size_kb += current_image_size_kb
                    print(f"[PROCESSOR]   Изображение оптимизировано до {current_image_size_kb:.1f}КБ и сохранено: {optimized_image_path_for_excel}", file=sys.stderr)
                else:
                    print(f"[PROCESSOR WARNING]   Оптимизация {image_path} вернула пустой или нулевой буфер. Пропускаем вставку.", file=sys.stderr)
                    continue

            except Exception as opt_e:
                 print(f"[PROCESSOR WARNING]   Ошибка при оптимизации {image_path}: {opt_e}. Пропускаем вставку.", file=sys.stderr)
                 import traceback
                 traceback.print_exc(file=sys.stderr)
                 continue

            # 2. ПОЛУЧЕНИЕ РАЗМЕРОВ и ВСТАВКА
            if optimized_image_path_for_excel and os.path.exists(optimized_image_path_for_excel):
                 print(f"[PROCESSOR]   Подготовка к вставке файла: {optimized_image_path_for_excel}", file=sys.stderr)
                 try:
                    print(f"[PROCESSOR]     Вызов get_image_dimensions для {optimized_image_path_for_excel}", file=sys.stderr)
                    # Добавляем проверку на случай, если get_image_dimensions вернет None
                    dimensions = image_utils.get_image_dimensions(optimized_image_path_for_excel)
                    if dimensions:
                        img_width_px, img_height_px = dimensions
                    else:
                        print(f"[PROCESSOR] WARNING: Не удалось получить размеры изображения, используем значения по умолчанию", file=sys.stderr)
                        img_width_px, img_height_px = DEFAULT_CELL_WIDTH_PX, DEFAULT_CELL_HEIGHT_PX
                    print(f"[PROCESSOR]     Получены размеры: {img_width_px}x{img_height_px}", file=sys.stderr)
                    
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
                    print(f"[PROCESSOR]     Вызов insert_image в ячейку {cell_address} (width: {DEFAULT_CELL_WIDTH_PX}, height: {scaled_img_height})", file=sys.stderr)
                    excel_utils.insert_image(
                        ws, 
                        optimized_image_path_for_excel,
                        anchor_cell=cell_address,
                        width=DEFAULT_CELL_WIDTH_PX,
                        height=scaled_img_height,  # Используем масштабированную высоту вместо оригинальной
                        preserve_aspect_ratio=True  # Явно указываем сохранение пропорций
                    )
                    images_inserted += 1
                    print(f"[PROCESSOR]   Изображение успешно вставлено в ячейку {cell_address}", file=sys.stderr)

                 except Exception as insert_e:
                    print(f"[PROCESSOR ERROR]   Ошибка вставки изображения {optimized_image_path_for_excel} в строку {excel_row_index}: {insert_e}", file=sys.stderr)
                    import traceback
                    traceback.print_exc(file=sys.stderr)
            else:
                 print(f"[PROCESSOR WARNING]   Пропуск вставки: путь к оптимизированному файлу отсутствует.", file=sys.stderr)
            
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

    print(f"[PROCESSOR] Сохранение результата в файл: {final_output_file}", file=sys.stderr)

    # --- Сохранение результата ---
    try:
        wb.save(final_output_file)
        print(f"[PROCESSOR] Результат успешно сохранен в файл: {final_output_file}", file=sys.stderr)
    except Exception as e:
        err_msg = f"Ошибка при сохранении результата в {final_output_file}: {e}"
        print(f"[PROCESSOR ERROR] {err_msg}", file=sys.stderr)
        # Попытка сохранить с другим именем во временную папку
        try:
            alt_output_folder = ensure_temp_dir("save_failed_")
            alt_output_file = os.path.join(alt_output_folder, os.path.basename(final_output_file))
            wb.save(alt_output_file)
            print(f"[PROCESSOR WARNING] Не удалось сохранить в {final_output_file}. Результат сохранен в {alt_output_file}", file=sys.stderr)
            final_output_file = alt_output_file # Возвращаем альтернативный путь
        except Exception as alt_e:
             err_msg = f"Ошибка при сохранении результата (попытка 2) в {alt_output_file}: {alt_e}"
             print(f"[PROCESSOR ERROR] {err_msg}", file=sys.stderr)
             import traceback
             traceback.print_exc(file=sys.stderr)
             raise RuntimeError(err_msg) from alt_e


    # --- Очистка временных изображений (если создавались) ---
    if temp_image_dir_created:
        try:
            shutil.rmtree(image_output_folder)
            print(f"[PROCESSOR] Удалена временная папка с изображениями: {image_output_folder}", file=sys.stderr)
        except Exception as e:
            print(f"[PROCESSOR WARNING] Не удалось удалить временную папку {image_output_folder}: {e}", file=sys.stderr)

    print(f"[PROCESSOR] Обработка завершена. Вставлено изображений: {images_inserted}", file=sys.stderr)
    return final_output_file, df, images_inserted