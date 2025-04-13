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
# <<< Тест 1: Проверяем, работает ли логгер этого модуля >>>
logger.critical("--- Logger for core.processor initialized ---") 

# <<< Constants for image fitting >>>
DEFAULT_CELL_WIDTH_PX = 150 
DEFAULT_CELL_HEIGHT_PX = 120
DEFAULT_IMG_QUALITY = 90
MIN_IMG_QUALITY = 30
MIN_KB_PER_IMAGE = 10
MAX_KB_PER_IMAGE = 2048 # 2MB max per image, prevents extreme cases
SIZE_BUDGET_FACTOR = 0.85 # Use 85% of total size budget for images
ROW_HEIGHT_PADDING = 15 # Pixels to add to image height for row height
MIN_ASPECT_RATIO = 0.5 # Минимальное соотношение сторон (высота/ширина)
MAX_ASPECT_RATIO = 2.0 # Максимальное соотношение сторон (высота/ширина)
EXCEL_PX_TO_PT_RATIO = 1.33 # Коэффициент преобразования пикселей в пункты Excel

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
    header_row: int = 0
) -> Tuple[str, pd.DataFrame, int]:
    # <<< Используем print в stderr вместо logger >>>
    print(">>> ENTERING process_excel_file <<<\n", file=sys.stderr)
    sys.stderr.flush()
    
    print(f"[PROCESSOR] Начало обработки: {file_path}", file=sys.stderr)
    print(f"[PROCESSOR] Параметры: article_col={article_col_letter}, img_folder={image_folder}, img_col={image_col_letter}, max_total_mb={max_total_file_size_mb}, header_row={header_row}", file=sys.stderr)

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
        wb = openpyxl.load_workbook(file_path)
        try:
            ws = wb.active
            if ws is None:
                ws = wb.worksheets[0]
                print("[PROCESSOR WARNING] Активный лист не определен, используем первый лист.", file=sys.stderr)
        except IndexError:
             print("[PROCESSOR ERROR] В файле нет листов для обработки.", file=sys.stderr)
             raise ValueError("Excel-файл не содержит листов.")
             
        print(f"[PROCESSOR] Загружена рабочая книга, работаем с листом: {ws.title}", file=sys.stderr)
        
    except Exception as e:
        err_msg = f"Ошибка при чтении Excel-файла: {e}"
        print(f"[PROCESSOR ERROR] {err_msg}", file=sys.stderr)
        # Выводим traceback в консоль
        import traceback
        traceback.print_exc(file=sys.stderr)
        raise RuntimeError(err_msg) from e

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
         err_msg = f"Ошибка при подготовке колонки для изображений ('{image_col_name}'): {e}"
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
    
    # Обрабатываем все строки, начиная с первой (без игнорирования)
    for df_index, row_data in df.iterrows():
        excel_row_index = df_index + 2
        rows_processed += 1
        print(f"[PROCESSOR] -- Обработка строки Excel #{excel_row_index} (DataFrame index: {df_index}) --", file=sys.stderr)
        
        if rows_processed % 50 == 0 or rows_processed == len(df):
             print(f"[PROCESSOR] Обработано {rows_processed} из {len(df)} строк ({rows_processed/len(df)*100:.1f}%)", file=sys.stderr)

        article = row_data.get(article_col_name)
        print(f"[PROCESSOR]   Исходное значение артикула: {repr(article)} (тип: {type(article)})", file=sys.stderr)
        
        if pd.isna(article) or not str(article).strip():
            print("[PROCESSOR]   Пропуск строки: артикул отсутствует или пуст после strip().", file=sys.stderr)
            continue

        article_str = str(article).strip()
        print(f"[PROCESSOR]   Артикул для поиска (строка, strip): '{article_str}'", file=sys.stderr)

        try:
            print(f"[PROCESSOR]   Вызов find_images_by_article для '{article_str}' в '{image_folder}'", file=sys.stderr)
            # Заменяем вызов image_utils на прямой вызов с print внутри
            # --- Начало кода из find_images_by_article --- 
            print(f"  [find_images] Поиск для '{article_str}' в '{image_folder}'", file=sys.stderr)
            found_image_paths_debug = []
            normalized_article_to_find_debug = image_utils.normalize_article(article_str)
            print(f"  [find_images] Нормализованный артикул: '{normalized_article_to_find_debug}'", file=sys.stderr)
            if os.path.isdir(image_folder) and normalized_article_to_find_debug:
                try:
                    all_files_in_dir_debug = os.listdir(image_folder)
                    print(f"  [find_images] Файлы в папке: {all_files_in_dir_debug}", file=sys.stderr)
                    normalized_name_to_original_path_debug: Dict[str, str] = {}
                    for filename_debug in all_files_in_dir_debug:
                        full_path_debug = os.path.join(image_folder, filename_debug)
                        if os.path.isfile(full_path_debug):
                            file_ext_lower_debug = os.path.splitext(filename_debug)[1].lower()
                            supported_extensions_debug = ('.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.webp')
                            if file_ext_lower_debug in supported_extensions_debug:
                                name_without_ext_debug = os.path.splitext(filename_debug)[0]
                                normalized_name_debug = image_utils.normalize_article(name_without_ext_debug)
                                print(f"    [find_images] Файл: '{filename_debug}', Норм: '{normalized_name_debug}'", file=sys.stderr)
                                if normalized_name_debug:
                                     normalized_name_to_original_path_debug[normalized_name_debug] = full_path_debug
                    print(f"  [find_images] Нормализованный словарь: {normalized_name_to_original_path_debug}", file=sys.stderr)
                    # Точное совпадение
                    if normalized_article_to_find_debug in normalized_name_to_original_path_debug:
                        exact_match_path_debug = normalized_name_to_original_path_debug[normalized_article_to_find_debug]
                        print(f"  [find_images] Найдено ТОЧНОЕ совпадение: {exact_match_path_debug}", file=sys.stderr)
                        if os.access(exact_match_path_debug, os.R_OK):
                            found_image_paths_debug.append(exact_match_path_debug)
                    # Частичное (если точного нет)
                    elif not found_image_paths_debug:
                         for norm_name_debug, original_path_debug in normalized_name_to_original_path_debug.items():
                            print(f"    [find_images] Проверка частичного: '{normalized_article_to_find_debug}' vs '{norm_name_debug}'", file=sys.stderr)
                            if normalized_article_to_find_debug in norm_name_debug or norm_name_debug in normalized_article_to_find_debug:
                                print(f"  [find_images] Найдено ЧАСТИЧНОЕ совпадение: {original_path_debug}", file=sys.stderr)
                                if os.access(original_path_debug, os.R_OK):
                                    found_image_paths_debug.append(original_path_debug)
                                    break # Берем первое частичное
                except Exception as find_e:
                    print(f"  [find_images] Ошибка поиска: {find_e}", file=sys.stderr)
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
                # <<< Вызываем реальную функцию оптимизации, но логируем результат >>>
                optimized_buffer = image_utils.optimize_image_for_excel(
                    image_path, 
                    max_size_kb=target_kb_per_image,
                    quality=DEFAULT_IMG_QUALITY,
                    min_quality=MIN_IMG_QUALITY    
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
                    
                    # Рассчитываем высоту строки с учетом отступа
                    calculated_height = scaled_img_height + ROW_HEIGHT_PADDING
                    
                    # Ensure the height is within the allowed range
                    min_height = DEFAULT_CELL_HEIGHT_PX
                    max_height = DEFAULT_CELL_HEIGHT_PX * 3  # Увеличиваем максимальную высоту
                    target_height = min(max(calculated_height, min_height), max_height)
                    
                    # Преобразуем пиксели в единицы Excel с использованием константы
                    target_row_height_pt = target_height / EXCEL_PX_TO_PT_RATIO
                    
                    print(f"[PROCESSOR]     Вызов set_row_height для строки {excel_row_index} на {target_row_height_pt:.1f} pt (расчет из пропорций {aspect_ratio:.2f}, высота {target_height} px)", file=sys.stderr)
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