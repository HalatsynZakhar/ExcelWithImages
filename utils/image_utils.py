"""
Утилиты для работы с изображениями
"""
import os
import re
import io
import logging
import math
from pathlib import Path
from typing import List, Dict, Optional, Tuple, Any, Union, Set

from PIL import Image as PILImage

logger = logging.getLogger(__name__)

def normalize_article(article: Any) -> str:
    """
    Нормализует артикул для поиска
    
    Args:
        article (Any): Артикул в любом формате
        
    Returns:
        str: Нормализованный артикул
    """
    if article is None:
        return ""
        
    # Преобразуем в строку
    article_str = str(article)
    
    # Удаляем все нецифровые и небуквенные символы, приводим к нижнему регистру
    normalized = re.sub(r'[^a-zA-Z0-9а-яА-Я]', '', article_str).lower()
    
    return normalized

def find_image_by_article(article: Any, images_folder: str, 
                         supported_extensions: Tuple[str, ...] = ('.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.webp')) -> Optional[str]:
    """
    Находит изображение по артикулу в указанной папке
    
    Args:
        article (Any): Артикул для поиска
        images_folder (str): Путь к папке с изображениями
        supported_extensions (Tuple[str, ...]): Поддерживаемые расширения файлов
        
    Returns:
        Optional[str]: Путь к найденному изображению или None, если не найдено
    """
    try:
        if not article:
            logger.warning("Пустой артикул")
            return None
            
        if not os.path.exists(images_folder):
            logger.error(f"Папка не найдена: {images_folder}")
            return None
            
        normalized_article = normalize_article(article)
        if not normalized_article:
            logger.warning(f"Артикул после нормализации пуст: {article}")
            return None
            
        logger.debug(f"Ищем изображение для артикула '{article}' (нормализованный: '{normalized_article}')")
        
        # Проверяем, существуют ли файлы в папке
        files = os.listdir(images_folder)
        if not files:
            logger.warning(f"Папка пуста: {images_folder}")
            return None
        
        # Создаем словарь нормализованных имен файлов
        file_dict = {}
        img_count = 0
        
        for filename in files:
            if any(filename.lower().endswith(ext) for ext in supported_extensions):
                img_count += 1
                name_without_ext = os.path.splitext(filename)[0]
                normalized_name = normalize_article(name_without_ext)
                file_dict[normalized_name] = filename
                logger.debug(f"Найдено изображение: {filename} (нормализованное имя: '{normalized_name}')")
                
        logger.debug(f"Всего найдено {img_count} изображений с поддерживаемыми расширениями")
                
        # Проверяем точное совпадение
        if normalized_article in file_dict:
            image_path = os.path.join(images_folder, file_dict[normalized_article])
            logger.debug(f"Найдено точное совпадение для артикула '{article}': {image_path}")
            
            # Дополнительная проверка, что файл существует и доступен
            if os.path.isfile(image_path) and os.access(image_path, os.R_OK):
                return image_path
            else:
                logger.warning(f"Найденный файл не существует или недоступен: {image_path}")
                return None
            
        # Проверяем частичное совпадение
        for norm_name, filename in file_dict.items():
            if normalized_article in norm_name or norm_name in normalized_article:
                image_path = os.path.join(images_folder, filename)
                logger.info(f"Найдено частичное совпадение для артикула '{article}': {image_path}")
                
                # Дополнительная проверка, что файл существует и доступен
                if os.path.isfile(image_path) and os.access(image_path, os.R_OK):
                    return image_path
                else:
                    logger.warning(f"Найденный файл не существует или недоступен: {image_path}")
                    continue
                
        logger.warning(f"Изображение для артикула '{article}' не найдено")
        return None
    except Exception as e:
        logger.error(f"Ошибка при поиске изображения по артикулу '{article}': {e}")
        return None

def optimize_image(image_path: str, max_size_kb: int = 200, 
                  quality_step: int = 5, min_quality: int = 30, 
                  target_width: int = 500, target_height: int = 500) -> io.BytesIO:
    """
    Оптимизирует изображение для вставки в Excel
    
    Args:
        image_path (str): Путь к изображению
        max_size_kb (int): Максимальный размер файла в КБ
        quality_step (int): Шаг снижения качества JPEG
        min_quality (int): Минимальное допустимое качество
        target_width (int): Целевая ширина
        target_height (int): Целевая высота
        
    Returns:
        io.BytesIO: Буфер с оптимизированным изображением
    """
    try:
        if not os.path.exists(image_path):
            logger.error(f"Изображение не найдено: {image_path}")
            raise FileNotFoundError(f"Изображение не найдено: {image_path}")
            
        # Открываем изображение
        img = PILImage.open(image_path)
        
        # Получаем исходный формат и размер
        original_format = img.format
        original_size_kb = os.path.getsize(image_path) / 1024
        logger.debug(f"Исходное изображение: формат {original_format}, размер {original_size_kb:.2f} КБ")
        
        # Изменяем размер, сохраняя пропорции
        original_width, original_height = img.size
        ratio = min(target_width / original_width, target_height / original_height)
        new_width = int(original_width * ratio)
        new_height = int(original_height * ratio)
        
        img = img.resize((new_width, new_height), PILImage.Resampling.LANCZOS)
        logger.debug(f"Изменен размер до {new_width}x{new_height}")
        
        # Вначале пробуем сохранить в исходном формате
        output = io.BytesIO()
        formats_to_try = []
        
        # Пробуем сначала исходный формат, затем другие в порядке уменьшения качества/размера
        if original_format in ['JPEG', 'JPG']:
            formats_to_try = ['JPEG', 'PNG', 'WEBP']
        elif original_format == 'PNG':
            formats_to_try = ['PNG', 'JPEG', 'WEBP']
        else:
            formats_to_try = ['JPEG', 'PNG', 'WEBP']
        
        logger.debug(f"Порядок форматов для оптимизации: {formats_to_try}")
        
        # Конвертируем в RGB, если необходимо (для форматов, не поддерживающих прозрачность)
        has_alpha = img.mode == 'RGBA'
        
        # Пробуем разные форматы и находим оптимальный по размеру
        best_format = None
        best_quality = None
        best_size = float('inf')
        best_buffer = None
        
        for img_format in formats_to_try:
            logger.debug(f"Пробуем формат: {img_format}")
            
            # Подготавливаем изображение в зависимости от формата
            if img_format == 'JPEG' and has_alpha:
                # Конвертируем в RGB для JPEG (убираем прозрачность)
                rgb_img = img.convert('RGB')
            else:
                rgb_img = img
            
            # Если это JPEG или WEBP, пробуем разное качество
            if img_format in ['JPEG', 'WEBP']:
                # Начинаем с высокого качества и постепенно снижаем
                quality = 95
                
                while quality >= min_quality:
                    # Очищаем буфер
                    temp_output = io.BytesIO()
                    
                    # Сохраняем изображение с текущим качеством
                    rgb_img.save(temp_output, format=img_format, quality=quality, optimize=True)
                    
                    # Проверяем размер
                    size_kb = temp_output.tell() / 1024
                    logger.debug(f"Формат {img_format}, качество {quality}: размер {size_kb:.2f} КБ")
                    
                    if size_kb <= max_size_kb and size_kb < best_size:
                        best_size = size_kb
                        best_format = img_format
                        best_quality = quality
                        # Сохраняем копию буфера
                        temp_output.seek(0)
                        best_buffer = io.BytesIO(temp_output.getvalue())
                        logger.debug(f"Найден новый лучший вариант: {img_format}, качество {quality}, размер {size_kb:.2f} КБ")
                    
                    # Если размер уже приемлемый, можно выходить
                    if size_kb <= max_size_kb:
                        break
                        
                    # Уменьшаем качество
                    quality -= quality_step
            else:
                # Для форматов без параметра качества (например, PNG)
                temp_output = io.BytesIO()
                rgb_img.save(temp_output, format=img_format, optimize=True)
                
                size_kb = temp_output.tell() / 1024
                logger.debug(f"Формат {img_format}: размер {size_kb:.2f} КБ")
                
                if size_kb <= max_size_kb and size_kb < best_size:
                    best_size = size_kb
                    best_format = img_format
                    best_quality = None
                    # Сохраняем копию буфера
                    temp_output.seek(0)
                    best_buffer = io.BytesIO(temp_output.getvalue())
                    logger.debug(f"Найден новый лучший вариант: {img_format}, размер {size_kb:.2f} КБ")
        
        # Если даже после всех попыток не удалось достичь требуемого размера
        if best_buffer is None or best_size > max_size_kb:
            logger.warning(f"Не удалось достичь требуемого размера {max_size_kb} КБ. Уменьшаем изображение.")
            
            # Пробуем уменьшать изображение до тех пор, пока не достигнем требуемого размера
            scale_factor = 0.9  # Уменьшаем на 10%
            current_img = img
            
            while scale_factor > 0.3:  # Ограничиваем минимальное уменьшение до 30% от исходного размера
                # Уменьшаем размер изображения
                new_width = int(new_width * scale_factor)
                new_height = int(new_height * scale_factor)
                
                if new_width < 50 or new_height < 50:
                    logger.warning("Изображение стало слишком маленьким. Прекращаем уменьшение.")
                    break
                    
                smaller_img = img.resize((new_width, new_height), PILImage.Resampling.LANCZOS)
                
                # Пробуем сохранить в формате JPEG с низким качеством
                temp_output = io.BytesIO()
                rgb_img = smaller_img.convert('RGB') if has_alpha else smaller_img
                rgb_img.save(temp_output, format='JPEG', quality=min_quality, optimize=True)
                
                size_kb = temp_output.tell() / 1024
                logger.debug(f"Уменьшенное до {new_width}x{new_height}, качество {min_quality}: размер {size_kb:.2f} КБ")
                
                if size_kb <= max_size_kb:
                    best_size = size_kb
                    best_format = 'JPEG'
                    best_quality = min_quality
                    temp_output.seek(0)
                    best_buffer = io.BytesIO(temp_output.getvalue())
                    logger.info(f"После уменьшения размера найден вариант: JPEG, размер {size_kb:.2f} КБ, {new_width}x{new_height}")
                    break
                
                scale_factor -= 0.1
        
        # Если все равно не удалось, возвращаем JPEG с минимальным качеством и размером
        if best_buffer is None:
            logger.warning("Не удалось оптимизировать изображение до требуемого размера. Возвращаем минимальный вариант.")
            smaller_img = img.resize((100, 100), PILImage.Resampling.LANCZOS)
            output = io.BytesIO()
            smaller_img.convert('RGB').save(output, format='JPEG', quality=min_quality, optimize=True)
            output.seek(0)
            
            best_size = output.tell() / 1024
            best_format = 'JPEG'
            best_quality = min_quality
            best_buffer = output
            
        # Логируем результат оптимизации
        logger.info(f"Изображение оптимизировано: формат {best_format}, " +
                   (f"качество {best_quality}, " if best_quality else "") +
                   f"размер {best_size:.2f} КБ")
        
        # Возвращаем оптимизированное изображение
        best_buffer.seek(0)
        return best_buffer
        
    except Exception as e:
        logger.error(f"Ошибка при оптимизации изображения {image_path}: {e}")
        raise

def optimize_image_for_excel(image_path: str, max_size_kb: int = 100, 
                          quality: int = 90, min_quality: int = 30) -> io.BytesIO:
    """
    Оптимизирует изображение для вставки в Excel без изменения размеров.
    Уменьшает качество изображения до достижения целевого размера файла.
    
    Args:
        image_path (str): Путь к изображению
        max_size_kb (int): Максимальный размер файла в КБ
        quality (int): Начальное качество (1-100)
        min_quality (int): Минимальное допустимое качество (1-100)
    
    Returns:
        io.BytesIO: Буфер с оптимизированным изображением
    """
    try:
        # Открываем изображение
        img = PILImage.open(image_path)
        
        # Обработка прозрачности (конвертирование в RGB)
        if img.mode in ('RGBA', 'LA') or (img.mode == 'P' and 'transparency' in img.info):
            # Создаем новый RGB-образ с белым фоном
            background = PILImage.new('RGB', img.size, (255, 255, 255))
            # Объединяем с исходным изображением
            if img.mode == 'RGBA':
                background.paste(img, mask=img.split()[3])  # 3 - альфа-канал
            else:
                background.paste(img, mask=img)
            img = background
        
        # Сохраняем изображение в буфер с оптимизацией по качеству
        result_buffer = io.BytesIO()
        current_quality = quality
        while current_quality >= min_quality:
            # Сбрасываем позицию буфера
            result_buffer.seek(0)
            result_buffer.truncate(0)
            
            # Сохраняем изображение с текущим качеством
            img.save(result_buffer, 'JPEG', quality=current_quality, optimize=True)
            
            # Проверяем размер
            file_size_kb = result_buffer.tell() / 1024
            
            logger.debug(f"Качество: {current_quality}, размер: {file_size_kb:.2f} КБ")
            
            if file_size_kb <= max_size_kb:
                break
            
            # Если размер файла все еще слишком большой, снижаем качество
            current_quality -= 5
        
        # Если даже на минимальном качестве файл слишком большой,
        # нужно будет действительно уменьшить размеры
        if current_quality < min_quality:
            logger.warning(f"Не удалось достичь целевого размера файла {max_size_kb} КБ. "
                          f"Текущий размер: {file_size_kb:.2f} КБ.")
            
            # Сбрасываем позицию буфера
            result_buffer.seek(0)
            result_buffer.truncate(0)
            
            # Находим пропорции для изменения размера изображения
            target_size_kb = max_size_kb * 0.9  # Берем с небольшим запасом
            size_ratio = math.sqrt(target_size_kb / file_size_kb)
            
            # Новые размеры
            new_width = int(img.width * size_ratio)
            new_height = int(img.height * size_ratio)
            
            # Изменяем размер и сохраняем
            resized_img = img.resize((new_width, new_height), PILImage.Resampling.LANCZOS)
            resized_img.save(result_buffer, 'JPEG', quality=min_quality, optimize=True)
            logger.info(f"Изображение уменьшено до {new_width}x{new_height} для достижения целевого размера")
        
        # Возвращаем в начало буфера
        result_buffer.seek(0)
        
        return result_buffer
    except Exception as e:
        logger.error(f"Ошибка при оптимизации изображения {image_path}: {e}")
        raise

def process_image(image_path: str, width: Optional[int] = None, height: Optional[int] = None,
                 max_size_kb: int = 200) -> Tuple[io.BytesIO, Tuple[int, int]]:
    """
    Обрабатывает изображение для Excel: изменяет размер и оптимизирует
    
    Args:
        image_path (str): Путь к изображению
        width (Optional[int]): Целевая ширина изображения
        height (Optional[int]): Целевая высота изображения
        max_size_kb (int): Максимальный размер файла в КБ
        
    Returns:
        Tuple[io.BytesIO, Tuple[int, int]]: Буфер с изображением и его размеры (ширина, высота)
    """
    try:
        logger.debug(f"Начинаем обработку изображения: {image_path}")
        
        if not os.path.exists(image_path):
            logger.error(f"Изображение не найдено: {image_path}")
            raise FileNotFoundError(f"Изображение не найдено: {image_path}")
        
        # Проверяем размер файла
        file_size_kb = os.path.getsize(image_path) / 1024
        logger.debug(f"Исходный размер файла: {file_size_kb:.2f} КБ")
        
        # Открываем изображение
        try:
            img = PILImage.open(image_path)
            logger.debug(f"Изображение открыто: {img.format}, размер: {img.size}, режим: {img.mode}")
        except Exception as e:
            logger.error(f"Не удалось открыть изображение {image_path}: {e}")
            raise
        
        # Получаем исходные размеры
        original_width, original_height = img.size
        logger.debug(f"Исходные размеры: {original_width}x{original_height}")
        
        # Определяем целевые размеры с сохранением пропорций
        if width is not None and height is not None:
            # Используем указанные размеры
            target_width, target_height = width, height
            logger.debug(f"Используем указанные размеры: {target_width}x{target_height}")
        elif width is not None:
            # Сохраняем соотношение сторон на основе ширины
            ratio = width / original_width
            target_width = width
            target_height = int(original_height * ratio)
            logger.debug(f"Масштабирование по ширине ({width}): новые размеры {target_width}x{target_height}")
        elif height is not None:
            # Сохраняем соотношение сторон на основе высоты
            ratio = height / original_height
            target_height = height
            target_width = int(original_width * ratio)
            logger.debug(f"Масштабирование по высоте ({height}): новые размеры {target_width}x{target_height}")
        else:
            # Если размеры не указаны, используем оригинальные
            target_width, target_height = original_width, original_height
            logger.debug(f"Используем оригинальные размеры: {target_width}x{target_height}")
        
        # Оптимизируем изображение
        try:
            logger.debug(f"Начинаем оптимизацию изображения с параметрами: max_size_kb={max_size_kb}, " +
                        f"target_width={target_width}, target_height={target_height}")
            img_buffer = optimize_image(
                image_path=image_path,
                max_size_kb=max_size_kb,
                target_width=target_width,
                target_height=target_height
            )
            logger.debug(f"Оптимизация завершена, размер буфера: {img_buffer.tell() / 1024:.2f} КБ")
        except Exception as e:
            logger.error(f"Ошибка при оптимизации изображения: {e}")
            raise
        
        # Определяем фактические размеры оптимизированного изображения
        try:
            with PILImage.open(img_buffer) as optimized_img:
                actual_width, actual_height = optimized_img.size
                logger.debug(f"Фактические размеры после оптимизации: {actual_width}x{actual_height}")
        except Exception as e:
            logger.error(f"Ошибка при получении размеров оптимизированного изображения: {e}")
            # Если не удалось получить фактические размеры, используем целевые
            actual_width, actual_height = target_width, target_height
            logger.warning(f"Используем целевые размеры вместо фактических: {actual_width}x{actual_height}")
        
        # Сбрасываем указатель буфера в начало
        img_buffer.seek(0)
        
        logger.info(f"Изображение успешно обработано: {image_path}, размеры: {actual_width}x{actual_height}, " + 
                   f"размер: {img_buffer.tell() / 1024:.2f} КБ")
        
        return img_buffer, (actual_width, actual_height)
    except Exception as e:
        logger.exception(f"Ошибка при обработке изображения {image_path}: {e}")
        raise

def get_image_dimensions(image_path: str) -> Optional[Tuple[int, int]]:
    """
    Получает размеры изображения
    
    Args:
        image_path (str): Путь к изображению
        
    Returns:
        Optional[Tuple[int, int]]: Кортеж (ширина, высота) или None в случае ошибки
    """
    try:
        if not os.path.exists(image_path):
            logger.error(f"Изображение не найдено: {image_path}")
            return None
        
        with PILImage.open(image_path) as img:
            width, height = img.size
            return width, height
    except Exception as e:
        logger.error(f"Ошибка при получении размеров изображения {image_path}: {e}")
        return None

def get_images_in_folder(folder_path: str, 
                       supported_extensions: Tuple[str, ...] = ('.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.webp')) -> List[str]:
    """
    Получает список путей к изображениям в указанной папке
    
    Args:
        folder_path (str): Путь к папке с изображениями
        supported_extensions (Tuple[str, ...]): Поддерживаемые расширения файлов
        
    Returns:
        List[str]: Список путей к изображениям
    """
    try:
        if not os.path.exists(folder_path):
            logger.error(f"Папка не найдена: {folder_path}")
            return []
        
        image_paths = []
        
        for filename in os.listdir(folder_path):
            if any(filename.lower().endswith(ext) for ext in supported_extensions):
                image_path = os.path.join(folder_path, filename)
                image_paths.append(image_path)
        
        logger.info(f"Найдено {len(image_paths)} изображений в папке {folder_path}")
        return image_paths
    except Exception as e:
        logger.error(f"Ошибка при получении списка изображений из папки {folder_path}: {e}")
        return []

def create_thumbnail(image_path: str, max_size: int = 100, quality: int = 85) -> Optional[io.BytesIO]:
    """
    Создает миниатюру изображения
    
    Args:
        image_path (str): Путь к изображению
        max_size (int): Максимальный размер (ширина или высота) в пикселях
        quality (int): Качество JPEG (1-100)
        
    Returns:
        Optional[io.BytesIO]: Буфер с миниатюрой или None в случае ошибки
    """
    try:
        if not os.path.exists(image_path):
            logger.error(f"Изображение не найдено: {image_path}")
            return None
        
        # Открываем изображение
        img = PILImage.open(image_path)
        
        # Получаем размеры
        width, height = img.size
        
        # Определяем новый размер, сохраняя пропорции
        if width > height:
            new_width = max_size
            new_height = int(height * (max_size / width))
        else:
            new_height = max_size
            new_width = int(width * (max_size / height))
        
        # Создаем миниатюру
        img = img.resize((new_width, new_height), PILImage.Resampling.LANCZOS)
        
        # Сохраняем в буфер
        thumb_buffer = io.BytesIO()
        img.save(thumb_buffer, format='JPEG', quality=quality)
        
        # Перемещаем указатель в начало буфера
        thumb_buffer.seek(0)
        
        return thumb_buffer
    except Exception as e:
        logger.error(f"Ошибка при создании миниатюры для {image_path}: {e}")
        return None

def save_buffer_to_file(buffer: io.BytesIO, output_path: str) -> bool:
    """
    Сохраняет содержимое буфера в файл
    
    Args:
        buffer (io.BytesIO): Буфер с данными
        output_path (str): Путь для сохранения файла
        
    Returns:
        bool: True, если сохранение успешно
    """
    try:
        # Создаем директорию, если она не существует
        output_dir = os.path.dirname(output_path)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        # Сохраняем данные из буфера в файл
        with open(output_path, 'wb') as f:
            f.write(buffer.getvalue())
        
        logger.debug(f"Файл сохранен: {output_path}")
        return True
    except Exception as e:
        logger.error(f"Ошибка при сохранении файла {output_path}: {e}")
        return False

def convert_image_format(image_path: str, output_format: str = 'JPEG', 
                       quality: int = 90) -> Optional[io.BytesIO]:
    """
    Конвертирует изображение в другой формат
    
    Args:
        image_path (str): Путь к изображению
        output_format (str): Формат вывода ('JPEG', 'PNG', 'BMP', и т.д.)
        quality (int): Качество (для форматов с потерями)
        
    Returns:
        Optional[io.BytesIO]: Буфер с конвертированным изображением или None в случае ошибки
    """
    try:
        if not os.path.exists(image_path):
            logger.error(f"Изображение не найдено: {image_path}")
            return None
        
        # Открываем изображение
        img = PILImage.open(image_path)
        
        # Если формат требует RGB, преобразуем
        if output_format in ('JPEG', 'JPG') and img.mode != 'RGB':
            img = img.convert('RGB')
        
        # Сохраняем в буфер
        output_buffer = io.BytesIO()
        
        # Для JPEG указываем качество
        if output_format in ('JPEG', 'JPG'):
            img.save(output_buffer, format=output_format, quality=quality)
        else:
            img.save(output_buffer, format=output_format)
        
        # Перемещаем указатель в начало буфера
        output_buffer.seek(0)
        
        return output_buffer
    except Exception as e:
        logger.error(f"Ошибка при конвертации изображения {image_path} в формат {output_format}: {e}")
        return None

def extract_articles_from_image_names(folder_path: str, 
                                    supported_extensions: Tuple[str, ...] = ('.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.webp')) -> Dict[str, str]:
    """
    Извлекает артикулы из имен файлов изображений
    
    Args:
        folder_path (str): Путь к папке с изображениями
        supported_extensions (Tuple[str, ...]): Поддерживаемые расширения файлов
        
    Returns:
        Dict[str, str]: Словарь {нормализованный артикул: путь к изображению}
    """
    try:
        if not os.path.exists(folder_path):
            logger.error(f"Папка не найдена: {folder_path}")
            return {}
        
        article_to_image = {}
        
        for filename in os.listdir(folder_path):
            if any(filename.lower().endswith(ext) for ext in supported_extensions):
                # Извлекаем имя файла без расширения
                name_without_ext = os.path.splitext(filename)[0]
                
                # Нормализуем для получения артикула
                normalized_article = normalize_article(name_without_ext)
                
                # Добавляем в словарь
                image_path = os.path.join(folder_path, filename)
                article_to_image[normalized_article] = image_path
        
        logger.info(f"Извлечено {len(article_to_image)} артикулов из изображений в папке {folder_path}")
        return article_to_image
    except Exception as e:
        logger.error(f"Ошибка при извлечении артикулов из имен файлов в папке {folder_path}: {e}")
        return {} 