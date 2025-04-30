import streamlit as st
import os
import sys
import logging
import io
import time
import tempfile
import shutil
from datetime import datetime
from pathlib import Path
from typing import Dict, Any, List, Optional, Tuple
import pandas as pd
from openpyxl.utils import get_column_letter
from openpyxl import load_workbook
from PIL import Image as PILImage
import json
import platform
import plotly.express as px

# Добавляем корневую папку проекта в PYTHONPATH
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

# Используем относительные импорты вместо абсолютных
from utils import config_manager
from utils import excel_utils
from utils import image_utils
from utils.config_manager import get_downloads_folder, ConfigManager
# <<< ДОБАВЛЯЕМ ГЛОБАЛЬНЫЙ ИМПОРТ >>>
from core.processor import process_excel_file 

# Настройка логирования
log_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'logs')
os.makedirs(log_dir, exist_ok=True)

# Ограничиваем количество файлов логов до 5 последних
log_files = sorted([f for f in os.listdir(log_dir) if f.startswith('app_')])
if len(log_files) > 5:
    for old_log in log_files[:-5]:
        try:
            os.remove(os.path.join(log_dir, old_log))
        except:
            pass

# Переименовываем текущий лог-файл, если он существует и создаем новый с правильной кодировкой
log_file = os.path.join(log_dir, 'app_latest.log')
# Всегда создаем новый лог-файл при запуске приложения
try:
    # Создаем новый файл с правильной кодировкой
    with open(log_file, 'w', encoding='utf-8') as f:
        f.write(f'{datetime.now().strftime("%Y-%m-%d %H:%M:%S")} - INFO - app - New log file created with UTF-8 encoding\n')
except Exception as e:
    print(f"Error creating log file: {e}")

log_stream = io.StringIO()
log_handler = logging.StreamHandler(log_stream)
log_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(name)s - %(message)s'))
log_handler.setLevel(logging.INFO)

# Используем один файл лога для всего приложения
file_handler = logging.FileHandler(log_file, encoding='utf-8')
file_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(name)s - %(message)s'))
file_handler.setLevel(logging.DEBUG)

root_logger = logging.getLogger()
root_logger.setLevel(logging.DEBUG)
# Удаляем существующие обработчики, если они есть
for handler in root_logger.handlers[:]:
    root_logger.removeHandler(handler)
root_logger.addHandler(log_handler)
root_logger.addHandler(file_handler)

# Устанавливаем кодировку для логирования
import sys
sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')

log = logging.getLogger(__name__)

# Определяем настройки по умолчанию
default_settings = {
    "paths": {
        "images_folder_path": get_downloads_folder()
    },
    "excel_settings": {
        "article_column": "A",
        "image_column": "B",
        "start_row": 1,
        "adjust_cell_size": False,
        "column_width": 150,
        "row_height": 120,
        "max_file_size_mb": 50,  # Максимальный размер результирующего файла в МБ
        "max_total_file_size_mb": 20  # New setting name
    },
    "check_images_on_startup": False
}

# Инициализация менеджера конфигурации с созданием настроек по умолчанию
def init_config_manager():
    """Инициализировать менеджер конфигурации и установить значения по умолчанию"""
    if 'config_manager' not in st.session_state:
        # Определяем путь к папке с пресетами
        presets_folder = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'settings_presets')
        
        # Инициализируем config manager с указанием папки пресетов
        config_manager_instance = config_manager.ConfigManager(presets_folder)
        
        # Получаем путь к папке загрузок пользователя
        downloads_folder = get_downloads_folder()
        
        # Устанавливаем значения по умолчанию, если они отсутствуют в конфиге
        if not config_manager_instance.get_setting('paths.images_folder_path'):
            config_manager_instance.set_setting('paths.images_folder_path', downloads_folder)
            # Логируем установку пути по умолчанию
            log.info(f"Установлен путь к папке с изображениями по умолчанию: {downloads_folder}")
            
        if not config_manager_instance.get_setting('excel_settings.max_file_size_mb'):
            config_manager_instance.set_setting('excel_settings.max_file_size_mb', 20)
            
        if not config_manager_instance.get_setting('image_settings.target_width'):
            config_manager_instance.set_setting('image_settings.target_width', 800)
            
        if not config_manager_instance.get_setting('image_settings.target_height'):
            config_manager_instance.set_setting('image_settings.target_height', 600)
        
        # Сохраняем конфигурацию
        config_manager_instance.save_settings("Default")
        
        # Сохраняем менеджер в session_state
        st.session_state.config_manager = config_manager_instance
        
        log.info("Менеджер конфигурации инициализирован с настройками по умолчанию")
    
    return st.session_state.config_manager

def get_downloads_folder():
    """Получить путь к папке с изображениями по умолчанию"""
    # Возвращаем сетевой путь вместо папки загрузок
    return r"\\10.10.100.2\Foto"
    
    # Закомментированный код ниже - оригинальная функция для получения папки загрузок
    # if platform.system() == "Windows":
    #     import winreg
    #     sub_key = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders'
    #     downloads_guid = '{374DE290-123F-4565-9164-39C4925E467B}'
    #     with winreg.OpenKey(winreg.HKEY_CURRENT_USER, sub_key) as key:
    #         downloads_folder = winreg.QueryValueEx(key, downloads_guid)[0]
    #         return downloads_folder
    # elif platform.system() == "Darwin":  # macOS
    #     return os.path.join(os.path.expanduser('~'), 'Downloads')
    # else:  # Linux и другие системы
    #     return os.path.join(os.path.expanduser('~'), 'Downloads')

# Обновляем код инициализации для использования нашей функции
config_folder = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'settings_presets')
# Инициализируем глобальный config_manager в модуле config_manager перед инициализацией нашего
config_manager.init_config_manager(config_folder)
init_config_manager()

# Настройка параметров приложения
st.set_page_config(
    page_title="Excel Image Processor",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Функция для создания временных директорий
def ensure_temp_dir(prefix: str = "") -> str:
    """
    Создает и возвращает путь к временной директории.
    
    Args:
        prefix (str): Префикс для имени временной директории
    
    Returns:
        Путь к временной директории
    """
    # Создаем временную директорию в папке проекта для лучшего доступа
    project_dir = os.path.dirname(os.path.dirname(__file__))
    temp_dir = os.path.join(project_dir, "temp")
    
    # Создаем директорию, если она не существует
    try:
        os.makedirs(temp_dir, exist_ok=True)
        log.info(f"Создана/проверена временная директория: {temp_dir}")
    except Exception as e:
        log.error(f"Ошибка при создании временной директории {temp_dir}: {e}")
        # Если не удалось создать в проекте, используем системную временную директорию
        temp_dir = os.path.join(tempfile.gettempdir(), f"{prefix}excelwithimages")
        try:
            os.makedirs(temp_dir, exist_ok=True)
            log.info(f"Использована системная временная директория: {temp_dir}")
        except Exception as e:
            log.error(f"Ошибка при создании системной временной директории {temp_dir}: {e}")
            # Если и системная не удалась, выбрасываем исключение
            raise RuntimeError("Не удалось создать временную директорию") from e
    
    return temp_dir

# Функция для очистки временных файлов
def cleanup_temp_files():
    """
    Очищает временные файлы, сохраняя только файлы текущей сессии.
    """
    try:
        # Определяем путь к временной директории
        temp_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'temp')
        if not os.path.exists(temp_dir):
            os.makedirs(temp_dir, exist_ok=True)
            log.info(f"Создана временная директория: {temp_dir}")
            return
        
        # Получаем время начала текущей сессии (приложение запущено)
        session_start_time = datetime.now()
        
        # Максимальный возраст файлов, которые мы хотим сохранить (в минутах)
        # Сохраняем только файлы, созданные в течение последнего часа
        max_age_minutes = 60
        
        # Файлы для сохранения (используемые в текущей сессии)
        files_to_keep = [
            st.session_state.get('temp_file_path', ''),
            st.session_state.get('output_file_path', '')
        ]
        
        # Получаем список всех файлов в временной директории
        all_files = os.listdir(temp_dir)
        log.info(f"Найдено {len(all_files)} файлов в директории {temp_dir}")
        
        # Удаляем старые файлы, которые не используются в текущей сессии
        removed_count = 0
        for filename in all_files:
            file_path = os.path.join(temp_dir, filename)
            
            # Пропускаем, если это не файл
            if not os.path.isfile(file_path):
                continue
                
            # Проверяем, используется ли файл в текущей сессии
            if file_path in files_to_keep:
                log.info(f"Сохраняем файл текущей сессии: {file_path}")
                continue
                
            # Получаем время последней модификации файла
            try:
                file_mod_time = datetime.fromtimestamp(os.path.getmtime(file_path))
                file_age = session_start_time - file_mod_time
                
                # Если файл старше максимального возраста или не из текущей сессии
                if file_age.total_seconds() > (max_age_minutes * 60):
                    try:
                        os.remove(file_path)
                        removed_count += 1
                        log.info(f"Удален старый временный файл: {file_path} (возраст: {file_age})")
                    except Exception as e:
                        log.error(f"Ошибка при удалении файла {file_path}: {e}")
            except Exception as e:
                log.error(f"Ошибка при проверке времени файла {file_path}: {e}")
                    
        log.info(f"Очистка временных файлов завершена. Удалено {removed_count} файлов.")
    except Exception as e:
        log.error(f"Ошибка при очистке временных файлов: {e}")

# Вызываем очистку временных файлов при запуске приложения
cleanup_temp_files()

# Функция для добавления сообщения в лог сессии
def add_log_message(message, level="INFO"):
    """
    Добавляет сообщение в лог сессии с временной меткой.
    
    Args:
        message (str): Сообщение для добавления
        level (str): Уровень сообщения (INFO, WARNING, ERROR, SUCCESS)
    """
    if 'log_messages' not in st.session_state:
        st.session_state.log_messages = []
    
    timestamp = datetime.now().strftime("%H:%M:%S")
    st.session_state.log_messages.append(f"[{timestamp}] [{level}] {message}")
    
    # Ограничиваем размер лога
    if len(st.session_state.log_messages) > 100:
        st.session_state.log_messages = st.session_state.log_messages[-100:]
    
    # Также добавляем в обычный лог
    if level == "ERROR":
        log.error(message)
    elif level == "WARNING":
        log.warning(message)
    else:
        log.info(message)

# Функция для обновления кнопок в сайдбаре была удалена
# и заменена функциональностью в settings_tab

# Функция для отображения настроек
def show_settings():
    config_manager = st.session_state.config_manager

    # --- Настройки путей к папкам ---
    with st.sidebar.expander("Настройки путей", expanded=True):
        # Получаем текущие значения из конфига
        config_manager = st.session_state.config_manager
        current_image_folder = config_manager.get_setting('paths.images_folder_path')
        current_secondary_folder = config_manager.get_setting('paths.secondary_images_folder_path')
        current_tertiary_folder = config_manager.get_setting('paths.tertiary_images_folder_path')
        
        # Добавляем пояснение
        st.markdown("### Пути к папкам с изображениями")
        st.markdown("Укажите пути к папкам с изображениями, которые нужно вставить в Excel файл. Поиск будет производиться в порядке приоритета.")
        
        # Отображаем текстовое поле для пути к основной папке с изображениями
        st.markdown("#### Основная папка с изображениями (приоритет 1)")
        image_folder = st.text_input(
            "Путь к основной папке с изображениями",
            value=st.session_state.get('images_folder_path', current_image_folder),
            help="Укажите полный путь к основной папке, где хранятся изображения товаров. Например: C:/Users/User/Downloads/images",
            key="image_folder_input"
        )
        
        # Создаем переменную для путей по умолчанию
        downloads_folder = get_downloads_folder()
        default_secondary_folder = r"\\10.10.100.2\Excel_reserve_photo"

        # Если вторая папка не указана, устанавливаем ее по умолчанию
        if not current_secondary_folder:
            config_manager.set_setting('paths.secondary_images_folder_path', default_secondary_folder)
            current_secondary_folder = default_secondary_folder
        
        # Отображаем текстовое поле для пути к второй папке с изображениями
        st.markdown("#### Запасная папка с изображениями (приоритет 2)")
        secondary_folder = st.text_input(
            "Путь к запасной папке с изображениями",
            value=st.session_state.get('secondary_images_folder_path', current_secondary_folder),
            help="Укажите полный путь к запасной папке с изображениями. Будет использована, если изображение не найдено в основной папке.",
            key="secondary_folder_input"
        )
        
        # Отображаем текстовое поле для пути к третьей папке с изображениями
        st.markdown("#### Дополнительная запасная папка (приоритет 3)")
        tertiary_folder = st.text_input(
            "Путь к дополнительной запасной папке",
            value=st.session_state.get('tertiary_images_folder_path', current_tertiary_folder),
            help="Укажите полный путь к дополнительной папке с изображениями. Будет использована, если изображение не найдено в основной и запасной папках.",
            key="tertiary_folder_input"
        )
        
        # Проверяем, является ли основной путь сетевым
        is_network_path = image_folder.startswith('\\\\') or image_folder.startswith('//')
        if is_network_path:
            st.warning("⚠️ Основной путь указывает на сетевой диск. Убедитесь, что у вас есть доступ к этой папке через проводник Windows.")
        
        # Проверяем доступность основного пути
        path_exists = os.path.exists(image_folder) if image_folder else False
        if path_exists:
            st.success("✅ Основной путь доступен")
        else:
            st.error("❌ Основной путь недоступен")
        
        # Проверяем доступность второго пути
        secondary_path_exists = os.path.exists(secondary_folder) if secondary_folder else False
        if secondary_path_exists:
            st.success("✅ Запасной путь доступен")
        else:
            st.warning("⚠️ Запасной путь недоступен")
            
        # Проверяем доступность третьего пути, если он указан
        if tertiary_folder:
            tertiary_path_exists = os.path.exists(tertiary_folder)
            if tertiary_path_exists:
                st.success("✅ Дополнительный запасной путь доступен")
            else:
                st.warning("⚠️ Дополнительный запасной путь недоступен")
        
        # Если основной путь изменился, сохраняем его в конфиг и session_state
        if image_folder != current_image_folder:
            config_manager.set_setting('paths.images_folder_path', image_folder)
            config_manager.save_settings("Default")
            # Сохраняем в session_state для сохранения между перезагрузками
            st.session_state.images_folder_path = image_folder
            log.info(f"Сохранен новый путь к основной папке с изображениями: {image_folder}")
            
        # Если второй путь изменился, сохраняем его в конфиг и session_state
        if secondary_folder != current_secondary_folder:
            config_manager.set_setting('paths.secondary_images_folder_path', secondary_folder)
            config_manager.save_settings("Default")
            # Сохраняем в session_state для сохранения между перезагрузками
            st.session_state.secondary_images_folder_path = secondary_folder
            log.info(f"Сохранен новый путь к запасной папке с изображениями: {secondary_folder}")
            
        # Если третий путь изменился, сохраняем его в конфиг и session_state
        if tertiary_folder != current_tertiary_folder:
            config_manager.set_setting('paths.tertiary_images_folder_path', tertiary_folder)
            config_manager.save_settings("Default")
            # Сохраняем в session_state для сохранения между перезагрузками
            st.session_state.tertiary_images_folder_path = tertiary_folder
            log.info(f"Сохранен новый путь к дополнительной запасной папке с изображениями: {tertiary_folder}")
            
        # Добавляем кнопку сброса путей к значениям по умолчанию
        if st.button("Сбросить пути к папкам изображений", 
                    help="Сбросить пути к папкам с изображениями на значения по умолчанию",
                    type="secondary"):
            config_manager.set_setting('paths.images_folder_path', downloads_folder)
            config_manager.set_setting('paths.secondary_images_folder_path', r"\\10.10.100.2\Excel_reserve_photo")
            config_manager.set_setting('paths.tertiary_images_folder_path', "")
            
            # Сохраняем в session_state для сохранения между перезагрузками
            st.session_state.images_folder_path = downloads_folder
            st.session_state.secondary_images_folder_path = r"\\10.10.100.2\Excel_reserve_photo"
            st.session_state.tertiary_images_folder_path = ""
            
            config_manager.save_settings("Default")
            st.success(f"Пути сброшены на значения по умолчанию")
            log.info(f"Пути сброшены на значения по умолчанию")
            st.rerun()
    
    # --- Настройки размера файла ---
    with st.sidebar.expander("Размер файла", expanded=True):
        st.subheader("Ограничение размера")
        
        max_total_file_size_mb = st.number_input(
            "Максимальный размер файла Excel (МБ)",
            min_value=1, # Minimum 1MB
            max_value=100, # Maximum 100MB
            value=int(config_manager.get_setting('excel_settings.max_total_file_size_mb', 20)), 
            step=1, # Step 1MB
            help="Приблизительный максимальный размер итогового Excel-файла. Изображения будут сжаты для достижения этого лимита.",
            key="max_total_file_size_mb_input"
        )
        if max_total_file_size_mb != config_manager.get_setting('excel_settings.max_total_file_size_mb', 20):
            config_manager.set_setting('excel_settings.max_total_file_size_mb', max_total_file_size_mb)
            config_manager.save_settings("Default") # Assuming 'Default' preset
            log.info(f"Настройка max_total_file_size_mb изменена на: {max_total_file_size_mb}")

# Функция для отображения предпросмотра таблицы
def show_table_preview(df):
    """
    Отображает предпросмотр таблицы с данными
    """
    if df is not None and not df.empty:
        try:
            # Преобразуем все столбцы с объектами в строки для предотвращения ошибок с pyarrow
            for col in df.select_dtypes(include=['object']).columns:
                df[col] = df[col].astype(str)
                
            # Выводим предпросмотр таблицы
            st.write("### Предпросмотр таблицы")
            
            # Отображаем только первые 10 строк для предпросмотра
            st.dataframe(df.head(10), use_container_width=True)
            
            # Отображаем информацию о количестве строк
            st.write(f"Всего строк в таблице: **{len(df)}**")
        except Exception as e:
            st.error(f"Ошибка при отображении предпросмотра таблицы: {e}")
    else:
        st.warning("Таблица пуста или не загружена.")

# Функция для загрузки Excel файла
def load_excel_file(uploaded_file_arg=None):
    # Используем файл из session_state, если аргумент не передан (для on_change)
    uploaded_file = uploaded_file_arg if uploaded_file_arg else st.session_state.get('file_uploader')
    if not uploaded_file:
        # Если файл удален из загрузчика
        log.warning("Файл был удален из загрузчика.")
        st.session_state.available_sheets = []
        st.session_state.selected_sheet = None
        st.session_state.df = None
        st.session_state.temp_file_path = None
        st.session_state.processing_error = None
        return

    # Используем временный путь из session_state
    temp_file_path = st.session_state.get('temp_file_path')
    if not temp_file_path or not os.path.exists(temp_file_path):
        log.error("Временный путь к файлу отсутствует или файл не найден.")
        st.session_state.processing_error = "Ошибка: временный файл не найден. Попробуйте загрузить файл заново."
        st.session_state.available_sheets = []
        st.session_state.selected_sheet = None
        st.session_state.df = None
        return
        
    try:
        log.info(f"Загрузка листов из файла: {temp_file_path}")
        excel_file = pd.ExcelFile(temp_file_path, engine='openpyxl')
        all_sheets = excel_file.sheet_names
        
        # Фильтруем листы, исключая листы с макросами
        filtered_sheets = [sheet for sheet in all_sheets if not sheet.startswith('xl/macrosheets/')]
        st.session_state.available_sheets = filtered_sheets
        log.info(f"Все листы: {all_sheets}")
        log.info(f"Доступные листы (без макросов): {st.session_state.available_sheets}")
        
        # Проверяем, были ли отфильтрованы листы с макросами
        if len(all_sheets) > len(filtered_sheets):
            log.warning(f"Обнаружены и отфильтрованы листы с макросами: {set(all_sheets) - set(filtered_sheets)}")
            # Если все листы были с макросами и отфильтрованы
            if not filtered_sheets:
                st.session_state.processing_error = "Внимание! Этот файл Excel содержит только макросы, а не обычные таблицы данных. Пожалуйста, выберите файл Excel с обычными листами, содержащими таблицы с артикулами и данными для обработки."
                return
        
        # --- Выбор листа по умолчанию --- 
        current_selection = st.session_state.get('selected_sheet')
        default_sheet = None
        if st.session_state.available_sheets:
            # Пытаемся найти первый "обычный" лист (не пустой, не скрытый - openpyxl может понадобиться для скрытых)
            # Простой вариант: просто берем первый
            default_sheet = st.session_state.available_sheets[0]
            log.info(f"Лист по умолчанию выбран: {default_sheet}")

        # Устанавливаем лист по умолчанию, если он еще не выбран или текущий выбор невалиден
        if default_sheet and (not current_selection or current_selection not in st.session_state.available_sheets):
             st.session_state.selected_sheet = default_sheet
             # Устанавливаем sheet_selector для корректной работы handle_sheet_change
             st.session_state.sheet_selector = default_sheet
             log.info(f"Установлен активный лист: {st.session_state.selected_sheet}")
             # Сбрасываем DataFrame, т.к. лист изменился (или был установлен впервые)
             st.session_state.df = None 
             st.session_state.processing_error = None

        # --- Загрузка данных с выбранного листа (если он есть) ---
        # Вызываем handle_sheet_change, чтобы загрузить данные для ВЫБРАННОГО листа
        # (это также обработает случай, когда лист был только что установлен по умолчанию)
        if st.session_state.selected_sheet:
            handle_sheet_change()  # Эта функция загрузит df и обработает ошибки
        else:
             # Если листов нет или выбрать по умолчанию не удалось
             st.session_state.df = None
             st.session_state.processing_error = "В файле не найдено листов для обработки."
             log.warning("Не удалось выбрать лист по умолчанию или листы отсутствуют.")

    except Exception as e:
        error_msg = f"Ошибка при чтении листов из Excel-файла: {e}"
        log.error(error_msg, exc_info=True)
        st.session_state.processing_error = error_msg
        st.session_state.available_sheets = []
        st.session_state.selected_sheet = None
        st.session_state.df = None

# Проверка валидности всех входных данных перед обработкой
def all_inputs_valid():
    """
    Проверяет, что все необходимые данные для обработки заполнены и валидны.
    
    Returns:
        bool: True, если все входные данные валидны, иначе False
    """
    # Подробная проверка с логированием
    valid = True
    log_msgs = []
    
    # 1. Проверяем наличие DataFrame
    if st.session_state.get('df') is None:
        log_msgs.append("DataFrame не загружен")
        valid = False
    else:
        log_msgs.append(f"DataFrame загружен, размер: {st.session_state.df.shape}")
        
    # 2. Проверяем, выбран ли лист в Excel
    if st.session_state.get('selected_sheet') is None:
        log_msgs.append("Лист Excel не выбран")
        valid = False
    else:
        log_msgs.append(f"Выбран лист: {st.session_state.selected_sheet}")

    # 3. Проверяем, выбрана ли колонка с артикулами (по НАЗВАНИЮ)
    if not st.session_state.get('article_column'): # Проверяем наличие и непустое значение
        log_msgs.append("Колонка с артикулами не выбрана")
        valid = False
    else:
        log_msgs.append(f"Выбрана колонка артикулов: {st.session_state.article_column}")

    # 4. Проверяем, выбрана ли колонка для изображений (по НАЗВАНИЮ)
    if not st.session_state.get('image_column'): # Проверяем наличие и непустое значение
        log_msgs.append("Колонка для изображений не выбрана")
        valid = False
    else:
        log_msgs.append(f"Выбрана колонка для изображений: {st.session_state.image_column}")

    # 5. Проверяем папку с изображениями
    images_folder = config_manager.get_setting("paths.images_folder_path", "")
    if not images_folder:
        log_msgs.append("Папка с изображениями не указана в настройках")
        valid = False
    elif not os.path.exists(images_folder):
        log_msgs.append(f"Папка с изображениями не найдена: {images_folder}")
        valid = False
    else:
        log_msgs.append(f"Папка с изображениями найдена: {images_folder}")

    # Логируем результат проверки
    final_msg = "Проверка валидности завершена. Результат: " + ("Успешно" if valid else "Неуспешно")
    log.info(final_msg)
    for msg in log_msgs:
        log.info(f"- {msg}")
    # add_log_message(final_msg, "INFO" if valid else "WARNING") # Можно добавить в лог сессии
        
    return valid

# Функция для обработки изменения выбранного листа
def handle_sheet_change():
    """
    Обрабатывает изменение выбранного листа Excel и перезагружает данные.
    """
    # Обновляем выбранный лист из селектора, если он был изменен
    if 'sheet_selector' in st.session_state and st.session_state.get("sheet_selector") != st.session_state.selected_sheet:
        st.session_state.selected_sheet = st.session_state.get("sheet_selector")
        log.info(f"Выбран новый лист из селектора: {st.session_state.selected_sheet}")
    
    # Проверяем, что у нас есть выбранный лист
    selected_sheet = st.session_state.get('selected_sheet')
    if not selected_sheet:
        log.warning("Не выбран лист для загрузки данных")
        st.session_state.df = None
        st.session_state.processing_error = "Не выбран лист для загрузки данных"
        return
        
    # Перезагружаем данные с выбранного листа
    if st.session_state.temp_file_path and os.path.exists(st.session_state.temp_file_path):
        try:
            log.info(f"Загрузка данных с листа: {selected_sheet}")
            
            # Всегда используем фиксированные значения: без пропуска строк и заголовок в первой строке
            df = pd.read_excel(
                st.session_state.temp_file_path, 
                sheet_name=selected_sheet, 
                engine='openpyxl',
                skiprows=0,
                header=None
            )
            
            # Преобразуем все столбцы с объектами в строки для предотвращения ошибок с pyarrow
            for col in df.select_dtypes(include=['object']).columns:
                df[col] = df[col].astype(str)
            
            # Проверка на пустой DataFrame
            log.info(f"Размер данных при смене листа: строк={df.shape[0]}, колонок={df.shape[1]}; пустой={df.empty}")
            
            if df.empty:
                error_msg = f"Лист '{selected_sheet}' не содержит данных"
                log.warning(error_msg)
                st.session_state.processing_error = error_msg
                st.session_state.df = None
                return
            
            if df.shape[0] == 0:
                error_msg = f"Лист '{selected_sheet}' не содержит строк с данными"
                log.warning(error_msg)
                st.session_state.processing_error = error_msg
                st.session_state.df = None
                return
                
            if df.shape[1] == 0:
                error_msg = f"Лист '{selected_sheet}' не содержит колонок с данными"
                log.warning(error_msg)
                st.session_state.processing_error = error_msg
                st.session_state.df = None
                return
            
            # Проверка на файл, который имеет колонки, но все значения в них NaN
            if df.notna().sum().sum() == 0:
                error_msg = f"Лист '{selected_sheet}' содержит только пустые ячейки"
                log.warning(error_msg)
                st.session_state.processing_error = error_msg
                st.session_state.df = None
                return
            
            # Все хорошо, сохраняем DataFrame
            st.session_state.df = df
            st.session_state.processing_error = None
            log.info(f"Лист '{selected_sheet}' успешно загружен. Найдено {len(df)} строк и {len(df.columns)} колонок")
            
        except Exception as e:
            error_msg = f"Ошибка при загрузке листа '{selected_sheet}': {str(e)}"
            log.error(error_msg)
            
            # Делаем сообщение об ошибке более понятным для пользователя
            user_friendly_msg = error_msg
            if "'dict' object has no attribute 'shape'" in str(e):
                user_friendly_msg = f"Лист '{selected_sheet}' не содержит табличных данных. Пожалуйста, выберите лист с необходимыми данными."
            elif "No sheet" in str(e) or "not found" in str(e):
                user_friendly_msg = f"Лист '{selected_sheet}' не найден в файле. Пожалуйста, выберите существующий лист."
            elif "Empty" in str(e) or "no data" in str(e):
                user_friendly_msg = f"Лист '{selected_sheet}' не содержит данных. Пожалуйста, выберите лист с данными."
            elif "ArrowTypeError" in str(e) or "Expected bytes" in str(e):
                user_friendly_msg = f"Ошибка преобразования типов данных. Попробуйте выбрать другой лист или перезагрузить файл."
                
            st.session_state.processing_error = user_friendly_msg
            st.session_state.df = None

# Функция для загрузки файла Excel
def file_uploader_section():
    """
    Отображает секцию для загрузки файла Excel.
    """
    with st.container():
        st.write("## Загрузка файла Excel")
        
        # CSS стили для кнопок и сообщений
        st.markdown("""
        <style>
        /* Стили для большой зеленой кнопки */
        .big-button-container {
            display: flex;
            justify-content: center;
            margin: 20px 0;
        }
        
        /* Увеличиваем высоту кнопок */
        .stButton > button:not([kind="secondary"]) {
            height: 80px !important;
            font-size: 20px !important;
            padding: 20px !important;
            width: 100% !important;
        }
        
        /* Специфичные стили для кнопки скачивания */
        div[data-testid="stDownloadButton"] button {
            height: 100px !important;
            font-size: 24px !important;
            padding: 25px !important;
            width: 100% !important;
            background-color: #4CAF50 !important;
            color: white !important;
            border: none !important;
            border-radius: 8px !important;
            transition: all 0.3s ease !important;
        }
        
        div[data-testid="stDownloadButton"] button:hover {
            background-color: #45a049 !important;
            transform: scale(1.02) !important;
        }
        
        /* Стиль для сообщений об ошибках */
        .error-message {
            color: #cc0000;
            background-color: #ffeeee;
            padding: 10px;
            border-radius: 5px;
            border-left: 5px solid #cc0000;
            margin: 10px 0;
        }
        
        /* Стиль для индикатора количества строк */
        .row-count {
            font-weight: bold;
            color: #1f77b4;
        }
        
        /* Стили для улучшения внешнего вида загрузчика файлов */
        div[data-testid="stFileUploader"] {
            border: 1px dashed #cccccc;
            padding: 10px;
            border-radius: 5px;
            background-color: #f8f9fa;
        }
        
        div[data-testid="stFileUploader"]:hover {
            border-color: #4CAF50;
            background-color: #f0f9f0;
        }
        
        /* Стили для лога */
        .log-container {
            max-height: 300px;
            overflow-y: auto;
            font-family: monospace;
            background-color: #f5f5f5;
            padding: 10px;
            border-radius: 5px;
            border: 1px solid #ddd;
            margin-top: 15px;
        }
        
        .log-entry {
            margin: 2px 0;
            font-size: 12px;
        }
        
        .log-info {
            color: #0366d6;
        }
        
        .log-warning {
            color: #e36209;
        }
        
        .log-error {
            color: #d73a49;
        }
        
        .log-success {
            color: #22863a;
        }
        </style>
        """, unsafe_allow_html=True)
        
        # Стили для кнопок
        big_green_button_style = """
            background-color: #4CAF50;
            color: white;
            padding: 15px 32px;
            text-align: center;
            text-decoration: none;
            display: inline-block;
            font-size: 16px;
            margin: 4px 2px;
            cursor: pointer;
            border-radius: 8px;
            border: none;
            box-shadow: 0 4px 8px 0 rgba(0,0,0,0.2);
            transition: 0.3s;
        """
        
        inactive_button_style = """
            background-color: #cccccc;
            color: #666666;
            padding: 15px 32px;
            text-align: center;
            text-decoration: none;
            display: inline-block;
            font-size: 16px;
            margin: 4px 2px;
            cursor: not-allowed;
            border-radius: 8px;
            border: none;
            box-shadow: none;
        """
        
        # Инициализируем переменные в session state, если их нет
        if 'df' not in st.session_state:
            st.session_state.df = None
        if 'temp_file_path' not in st.session_state:
            st.session_state.temp_file_path = None
        if 'processing_result' not in st.session_state:
            st.session_state.processing_result = None
        if 'processing_error' not in st.session_state:
            st.session_state.processing_error = None
        if 'is_processing' not in st.session_state:
            st.session_state.is_processing = False
        if 'output_file_path' not in st.session_state:
            st.session_state.output_file_path = None
        if 'selected_sheet' not in st.session_state:
            st.session_state.selected_sheet = None
        if 'available_sheets' not in st.session_state:
            st.session_state.available_sheets = []
        if 'log_messages' not in st.session_state:
            st.session_state.log_messages = []
            
        # Загрузчик файлов Excel
        uploaded_file = st.file_uploader("Выберите Excel файл для обработки", type=["xlsx", "xls"], key="file_uploader",
                                     on_change=load_excel_file)

        # Отображение информации о загруженном файле
        if uploaded_file is not None:
            st.write(f"**Загружен файл:** {uploaded_file.name}")
            
            current_temp_path = st.session_state.get('temp_file_path', '')
            if not current_temp_path or os.path.basename(current_temp_path) != uploaded_file.name:
                temp_dir = ensure_temp_dir()
                temp_file_path = os.path.join(temp_dir, uploaded_file.name)
                with open(temp_file_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())
                st.session_state.temp_file_path = temp_file_path
                add_log_message(f"Файл сохранен: {temp_file_path}", "INFO")
                load_excel_file()
            
            # Удалены настройки пропуска начальных строк и строки с заголовками
            # Инициализация переменных с фиксированными значениями
            st.session_state.skiprows = 0
            st.session_state.header_row = 0
                    
            # Отображение ошибки обработки, если есть
            if st.session_state.processing_error:
                st.markdown(f"""
                <div class="error-message">
                    <strong>Ошибка:</strong> {st.session_state.processing_error}
                </div>
                """, unsafe_allow_html=True)
                
                # Добавляем подсказку для решения проблемы с пустыми данными
                if "не содержит данных" in st.session_state.processing_error or "содержит только пустые ячейки" in st.session_state.processing_error:
                    st.info("""
                    **Рекомендации по решению проблемы:**
                    
                    1. Убедитесь, что файл Excel содержит данные в выбранном листе
                    2. Проверьте наличие невидимых форматирований или скрытых строк
                    3. Попробуйте открыть файл в Excel и пересохранить его
                    4. Убедитесь, что данные начинаются с первой строки и колонки
                    """)
                
            # Если есть доступные листы, показываем селектор листов
            if st.session_state.available_sheets and len(st.session_state.available_sheets) > 0:
                selected_sheet = st.selectbox(
                    "Выберите лист для обработки:",
                    st.session_state.available_sheets,
                    index=st.session_state.available_sheets.index(st.session_state.selected_sheet) if st.session_state.selected_sheet in st.session_state.available_sheets else 0,
                    key="sheet_selector",
                    on_change=handle_sheet_change
                )
                
            # Если данные успешно загружены, показываем предпросмотр и селекторы колонок
            if st.session_state.df is not None:
                # Отображение размерности данных
                col1, col2 = st.columns(2)
                with col1:
                    st.markdown(f"""
                    <div class="row-count">
                        Количество строк: {st.session_state.df.shape[0]}
                    </div>
                    """, unsafe_allow_html=True)
                with col2:
                    st.write(f"**Количество колонок:** {st.session_state.df.shape[1]}")
                
                # Добавляем предпросмотр данных
                with st.expander("Предпросмотр данных", expanded=False):
                    st.dataframe(st.session_state.df.head(10), use_container_width=True)
                    
                    # Добавляем статистику по колонкам
                    col_stats = pd.DataFrame({
                        'Колонка': st.session_state.df.columns,
                        'Тип данных': [str(dtype) for dtype in st.session_state.df.dtypes.values],
                        'Непустых значений': st.session_state.df.count().values,
                        'Процент заполнения': (st.session_state.df.count() / len(st.session_state.df) * 100).round(2).values
                    })
                    st.write("### Статистика по колонкам")
                    st.dataframe(col_stats, use_container_width=True)
                
                # Получение списка колонок
                column_options = list(st.session_state.df.columns)
                
                # Если колонки есть, показываем селекторы
                if column_options:
                    # Определяем индексы по умолчанию (если колонки A/B существуют)
                    default_article_index = column_options.index("A") if "A" in column_options else 0
                    default_image_index = column_options.index("B") if "B" in column_options else min(1, len(column_options)-1)
                    
                    # Позволяем пользователю ввести буквенные обозначения колонок
                    col1, col2 = st.columns(2)
                    with col1:
                        selected_article_col = st.text_input(
                            "Буква колонки с артикулами (A, B, C...)", 
                            value="A",
                            key="article_column_input",
                            help="Введите букву колонки, содержащей артикулы товаров (например: A, B, C)"
                        )
                        st.session_state.article_column = selected_article_col # Сохраняем выбранную букву
                    with col2:
                        selected_image_col = st.text_input(
                            "Буква колонки для изображений (A, B, C...)", 
                            value="B",
                            key="image_column_input",
                            help="Введите букву колонки для вставки изображений (например: B, C, D)"
                        )
                        st.session_state.image_column = selected_image_col # Сохраняем выбранное НАЗВАНИЕ
                    
                    # Проверка всех необходимых полей перед обработкой
                    process_button_disabled = not all_inputs_valid()
                    
                    # Кнопка для запуска обработки
                    st.button("Обработать файл", 
                              disabled=process_button_disabled, 
                              type="primary", 
                              key="process_button",
                              on_click=trigger_processing,
                              use_container_width=True)  # Добавляем параметр для растягивания на всю ширину
                    
                    # Запускаем обработку, если установлен флаг
                    if st.session_state.get('start_processing', False):
                        st.info("Идет обработка файла. Не закрывайте страницу и не взаимодействуйте с интерфейсом до завершения.")
                        st.write("Это может занять некоторое время в зависимости от количества строк и изображений.")
                        
                        # Блокируем интерфейс на время обработки
                        with st.spinner("Обработка файла..."):
                            # Очищаем предыдущие результаты и ошибки
                            st.session_state.processing_result = None
                            st.session_state.processing_error = None
                            
                            # Выполняем обработку
                            success = process_files()
                            
                            # Записываем результат в session_state для отображения после перезагрузки
                            if success:
                                st.session_state.processing_result = "Файл успешно обработан! Вы можете скачать его ниже."
                                # Устанавливаем флаг для автоматического скролла к секции скачивания после перезагрузки
                                st.session_state.scroll_to_download = True
                            else:
                                st.session_state.processing_error_message = st.session_state.processing_error
                        
                        # Сбрасываем флаг обработки
                        st.session_state.start_processing = False
                        
                        # Форсируем перезагрузку страницы для обновления UI
                        st.rerun()
                    
            else:
                st.warning("Файл не содержит колонок для выбора. Проверьте структуру Excel-файла.")
                
        else:
                    st.warning("Файл не содержит колонок для выбора. Проверьте структуру Excel-файла.")
                    
        # Отображение результатов обработки и ошибок после обработки файла
        # Показываем только если обработка не выполняется сейчас
        if not st.session_state.get('start_processing', False):
            # Успешное завершение обработки
            if st.session_state.get('processing_result'):
                st.success(st.session_state.processing_result)
                # Если нужно автоматически прокрутить к секции скачивания
                if st.session_state.get('scroll_to_download', False):
                    st.markdown('<script>setTimeout(function() { window.scrollTo(0, document.body.scrollHeight); }, 500);</script>', 
                                unsafe_allow_html=True)
                    # Сбрасываем флаг скролла
                    st.session_state.scroll_to_download = False
    
            # Ошибка обработки
            if st.session_state.get('processing_error_message'):
                st.error(f"Ошибка при обработке файла: {st.session_state.processing_error_message}")
                # Очищаем сообщение об ошибке после отображения
                st.session_state.processing_error_message = None
            
            # Добавление кнопки скачивания, если файл был обработан
            if st.session_state.output_file_path and os.path.exists(st.session_state.output_file_path):
                # Создаем колонку для центрирования кнопки (опционально, для лучшего вида)
                col1, col2, col3 = st.columns([1,2,1])
                with col2:
                    with open(st.session_state.output_file_path, "rb") as file:
                        st.download_button(
                            label="СКАЧАТЬ ОБРАБОТАННЫЙ ФАЙЛ",
                            data=file,
                            file_name=os.path.basename(st.session_state.output_file_path),
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True,
                            type="primary",
                            key="download_button"
                        )
        
        # Проверяем, нужно ли отобразить отчет о результатах обработки
        if st.session_state.get('show_processing_report', False):
            show_processing_report()
            # Сбрасываем флаг, чтобы не показывать отчет повторно
            st.session_state.show_processing_report = False
        
        # Добавляем отображение логов вместо отладочной информации
        with st.expander("Журнал событий", expanded=False):
            # Отображаем сообщения из st.session_state.log_messages
            if 'log_messages' in st.session_state and st.session_state.log_messages:
                st.markdown('<div class="log-container">', unsafe_allow_html=True)
                for log_msg in st.session_state.log_messages:
                    # Определяем класс для стилизации
                    log_class = "log-info"
                    if "ERROR" in log_msg:
                        log_class = "log-error"
                    elif "WARNING" in log_msg:
                        log_class = "log-warning"
                    elif "SUCCESS" in log_msg:
                        log_class = "log-success"
                        
                    st.markdown(f'<div class="log-entry {log_class}">{log_msg}</div>', unsafe_allow_html=True)
                st.markdown('</div>', unsafe_allow_html=True)
            else:
                st.info("Журнал пуст")

# Функция для обработки файла
def process_files():
    """
    Основная функция для обработки файлов.
    Обрабатывает Excel файл и добавляет изображения к соответствующим артикулам.
    """
    try:
        log.info("===================== НАЧАЛО ОБРАБОТКИ ФАЙЛА =====================")
        add_log_message("Начало обработки файла", "INFO")
        st.session_state.is_processing = True
        st.session_state.processing_result = None
        st.session_state.processing_error = None
        
        # Отключаем взаимодействие с интерфейсом во время обработки
        # Показываем крутящийся индикатор загрузки
        with st.spinner("Идет обработка файла. Пожалуйста, подождите..."):
            # Получаем необходимые настройки
            images_folder = st.session_state.get('images_folder_path', 
                                               config_manager.get_setting("paths.images_folder_path", ""))
    
            # Создаем директорию для хранения результатов
            temp_dir = ensure_temp_dir()
            output_folder = temp_dir
            
            add_log_message(f"Папка изображений: {images_folder}", "INFO")
            
            # Детальная проверка всех условий
            conditions = {
                "DataFrame загружен": st.session_state.df is not None,
                "Временный файл существует": st.session_state.temp_file_path is not None,
                "Файл доступен": (os.path.exists(st.session_state.temp_file_path) and 
                                  os.access(st.session_state.temp_file_path, os.R_OK)) if st.session_state.temp_file_path else False,
                "Выбран лист": st.session_state.selected_sheet is not None,
                "Указана колонка с артикулами": st.session_state.get('article_column') is not None,
                "Указана колонка с изображениями": st.session_state.get('image_column') is not None,
                "Папка изображений указана": images_folder != "",
                "Папка изображений существует": os.path.exists(images_folder) if images_folder else False,
                "Папка изображений доступна": (os.path.exists(images_folder) and 
                                           os.access(images_folder, os.R_OK | os.X_OK)) if images_folder else False
            }
            
            # Логируем все условия
            for condition, result in conditions.items():
                log.info(f"Проверка: {condition} = {result}")
                add_log_message(f"Проверка: {condition} = {result}", "INFO" if result else "WARNING")
            
            # Проверяем все условия
            all_conditions_met = all(conditions.values())
            if not all_conditions_met:
                failed_conditions = [cond for cond, result in conditions.items() if not result]
                error_msg = f"Не выполнены следующие условия: {', '.join(failed_conditions)}"
                log.error(error_msg)
                add_log_message(error_msg, "ERROR")
                st.session_state.processing_error = error_msg
                st.session_state.is_processing = False
                return False
                
            # Получаем данные из session_state
            excel_file_path = st.session_state.temp_file_path
            # <<< Используем ИМЕНА колонок из session_state >>>
            article_col_name = st.session_state.get('article_column') 
            image_col_name = st.session_state.get('image_column') 
            selected_sheet = st.session_state.selected_sheet
            
            log.info(f"Параметры обработки:")
            log.info(f"- Файл: {excel_file_path}")
            log.info(f"- Лист: {selected_sheet}")
            log.info(f"- Колонка с артикулами: {article_col_name}") # Log name
            log.info(f"- Колонка с изображениями: {image_col_name}") # Log name
            log.info(f"- Папка с изображениями: {images_folder}")
            
            add_log_message(f"Обработка файла: {os.path.basename(excel_file_path)}, лист: {selected_sheet}", "INFO")
            add_log_message(f"Колонки: артикулы - {article_col_name}, изображения - {image_col_name}", "INFO") # Log names
            
            # Создаем timestamp для временных файлов
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            
            try:
                # Сначала копируем весь оригинальный файл вместо сохранения только выбранного листа
                log.info(f"Создание копии оригинального Excel-файла")
                temp_file_with_full_copy = os.path.join(output_folder, f"temp_full_{timestamp}.xlsx")
                shutil.copy2(excel_file_path, temp_file_with_full_copy)
                log.info(f"Создана полная копия исходного файла: {temp_file_with_full_copy}")
                
                # Сохраняем оригинальное имя файла перед заменой пути
                original_filename = os.path.basename(excel_file_path)
                
                # Создаем имя выходного файла, используя оригинальное имя
                output_filename = f"{os.path.splitext(original_filename)[0]}_with Images.xlsx"
                output_file_path = os.path.join(output_folder, output_filename)
                log.info(f"Выходной файл будет: {output_file_path}")
                add_log_message(f"Подготовка выходного файла: {output_filename}", "INFO")
                
                # Заменяем путь к файлу на временный
                excel_file_path = temp_file_with_full_copy

                # <<< ЛОГ ПЕРЕД ВЫЗОВОМ >>>
                log.info("--- Готовимся к вызову process_excel_file ---")
                log.info(f"  file_path: {excel_file_path}")
                log.info(f"  article_col_name: {article_col_name}")
                log.info(f"  image_col_name: {image_col_name}")
                log.info(f"  image_folder: {images_folder}")
                
                # Получаем пути к вторичной и третичной папкам
                secondary_folder = st.session_state.get('secondary_images_folder_path', 
                                                 config_manager.get_setting("paths.secondary_images_folder_path", ""))
                tertiary_folder = st.session_state.get('tertiary_images_folder_path', 
                                                config_manager.get_setting("paths.tertiary_images_folder_path", ""))
                
                log.info(f"  secondary_folder: {secondary_folder}")
                log.info(f"  tertiary_folder: {tertiary_folder}")
                log.info(f"  output_folder: {output_folder}")
                log.info(f"  selected_sheet: {selected_sheet}")  # Добавляем выбранный лист в лог
                current_max_mb = config_manager.get_setting('excel_settings.max_total_file_size_mb', 20)
                log.info(f"  max_total_file_size_mb: {current_max_mb}")
                add_log_message("Запуск основной обработки файла...", "INFO") # Лог для UI

                # Вызываем функцию обработки, передавая имя листа
                result_file_path, result_df, images_inserted, multiple_images_found, not_found_articles, image_search_results = process_excel_file(
                    file_path=excel_file_path,
                    article_col_name=article_col_name if article_col_name.isalpha() else df.columns.get_loc(article_col_name) + 1,
                    image_col_name=image_col_name if image_col_name.isalpha() else df.columns.get_loc(image_col_name) + 1,
                    image_folder=images_folder,
                    secondary_image_folder=secondary_folder,  # Передаем путь к вторичной папке
                    tertiary_image_folder=tertiary_folder,    # Передаем путь к третичной папке
                    output_folder=output_folder,
                    max_total_file_size_mb=current_max_mb,
                    header_row=st.session_state.get('header_row', 0),
                    sheet_name=selected_sheet,  # Добавляем передачу имени листа
                    output_filename=output_filename  # Передаем готовое имя выходного файла
                )

                # <<< ЛОГ ПОСЛЕ УСПЕШНОГО ВЫЗОВА >>>
                log.info("--- process_excel_file завершился успешно ---")
                log.info(f"  result_file_path: {result_file_path}")
                log.info(f"  images_inserted: {images_inserted}")
                add_log_message(f"Обработка завершена. Вставлено изображений: {images_inserted}", "SUCCESS") # Лог для UI
                
                # Отображаем отчет о ненайденных артикулах и артикулах с несколькими изображениями
                if not_found_articles:
                    log.info(f"  Не найдены изображения для {len(not_found_articles)} артикулов")
                    add_log_message(f"Не найдены изображения для {len(not_found_articles)} артикулов", "WARNING")
                    # Создаем список для отображения
                    st.session_state.not_found_articles = not_found_articles
                
                if multiple_images_found:
                    log.info(f"  Найдено несколько вариантов изображений для {len(multiple_images_found)} артикулов")
                    add_log_message(f"Найдено несколько вариантов изображений для {len(multiple_images_found)} артикулов", "INFO")
                    # Сохраняем данные для отображения
                    st.session_state.multiple_images_found = multiple_images_found

                # Проверяем, что результирующий файл создан
                if not os.path.exists(result_file_path):
                    error_msg = "Выходной файл не был создан, хотя ошибок не возникло"
                    log.error(error_msg)
                    add_log_message(error_msg, "ERROR")
                    st.session_state.processing_error = error_msg
                    return False

                # Сохраняем путь к выходному файлу
                st.session_state.output_file_path = result_file_path

                # Формируем сообщение об успешной обработке
                success_msg = f"Обработка успешно завершена. Файл готов к скачиванию."
                log.info(success_msg)
                add_log_message(success_msg, "SUCCESS")
                
                # Сохраняем результаты для отображения отчета
                st.session_state.processing_result = {
                    'output_path': result_file_path,
                    'df': result_df,
                    'images_inserted': images_inserted,
                    'multiple_images': multiple_images_found,
                    'not_found': not_found_articles,
                    'image_search_results': image_search_results
                }

                # Отображаем отчет о результатах обработки сразу после успешной обработки
                # Вызов будет обработан после перезагрузки страницы
                st.session_state.show_processing_report = True
                
                log.info("===================== ОБРАБОТКА ЗАВЕРШЕНА УСПЕШНО =====================")
                return True

            except Exception as e:
                # <<< ЛОГ ПРИ ОШИБКЕ ВЫЗОВА/ВЫПОЛНЕНИЯ >>>
                error_msg = f"Ошибка при вызове/выполнении process_excel_file: {str(e)}"
                log.error(error_msg, exc_info=True) # Добавляем traceback в основной лог
                add_log_message(error_msg, "ERROR") # Сообщение для UI
                st.session_state.processing_error = error_msg
                log.info("===================== ОБРАБОТКА ЗАВЕРШЕНА С ОШИБКОЙ =====================")
                return False
    finally:
        # Сбрасываем флаг обработки в любом случае
        st.session_state.is_processing = False

# Функция для отображения результатов обработки
def show_results(stats: Dict[str, Any]):
    if not stats:
        return
    
    st.success("Обработка завершена успешно!")
    
    # Отображаем основную информацию
    st.subheader("Результаты обработки")
    
    col1, col2, col3 = st.columns(3)
    col1.metric("Всего артикулов", stats["total_articles"])
    col2.metric("Найдено изображений", stats["images_found"])
    col3.metric("Вставлено в Excel", stats["images_inserted"])
    
    # Отображаем путь к выходному файлу
    st.info(f"Файл сохранен: {stats['output_file']}")
    
    # Кнопка для открытия папки с результатом
    if st.button(
        "Открыть папку с результатом", 
        key="open_result_folder_button", 
        help="Открыть папку, содержащую обработанный файл Excel"
    ):
        output_folder = os.path.dirname(stats["output_file"])
        # Используем команду в зависимости от ОС
        if os.name == 'nt':  # Windows
            os.startfile(output_folder)
        elif os.name == 'posix':  # macOS и Linux
            os.system(f"open {output_folder}")  # macOS
            # os.system(f"xdg-open {output_folder}")  # Linux

# Функция для инициализации переменных сессии
def initialize_session_state():
    """
    Инициализирует переменные состояния сессии, если они не существуют.
    """
    # НЕ инициализируем uploaded_file, так как Streamlit управляет им сам
    if 'df' not in st.session_state:
        st.session_state.df = None
    if 'temp_file_path' not in st.session_state:
        st.session_state.temp_file_path = None
    if 'processing_result' not in st.session_state:
        st.session_state.processing_result = None
    if 'processing_error' not in st.session_state:
        st.session_state.processing_error = None
    if 'is_processing' not in st.session_state:
        st.session_state.is_processing = False
    if 'output_file_path' not in st.session_state:
        st.session_state.output_file_path = None
    if 'current_settings' not in st.session_state:
        st.session_state.current_settings = config_manager.get_config_manager().current_settings
    if 'selected_sheet' not in st.session_state:
        st.session_state.selected_sheet = None
    if 'article_column' not in st.session_state:
        st.session_state.article_column = config_manager.get_setting('excel_settings.article_column', "A")
    if 'image_column' not in st.session_state:
        st.session_state.image_column = config_manager.get_setting('excel_settings.image_column', "B")
    if 'images_folder_path' not in st.session_state:
        st.session_state.images_folder_path = config_manager.get_setting('paths.images_folder_path', "")
    if 'secondary_images_folder_path' not in st.session_state:
        st.session_state.secondary_images_folder_path = config_manager.get_setting('paths.secondary_images_folder_path', "")
    if 'tertiary_images_folder_path' not in st.session_state:
        st.session_state.tertiary_images_folder_path = config_manager.get_setting('paths.tertiary_images_folder_path', "")
    if 'header_row' not in st.session_state:
        st.session_state.header_row = config_manager.get_setting('excel_settings.start_row', 0)
    if 'sheet_names' not in st.session_state:
        st.session_state.sheet_names = []
    if 'show_processing_report' not in st.session_state:
        st.session_state.show_processing_report = False
    if 'needs_rerun' not in st.session_state:
        st.session_state.needs_rerun = False

# Функция для отображения вкладки настроек в боковой панели
def settings_tab():
    """
    Отображает вкладку настроек в боковой панели.
    """
    st.sidebar.title("Настройки")
    
    # --- Настройки путей к папкам ---
    with st.sidebar.expander("Настройки путей", expanded=True):
        # Получаем текущие значения из конфига
        config_manager = st.session_state.config_manager
        current_image_folder = config_manager.get_setting('paths.images_folder_path')
        current_secondary_folder = config_manager.get_setting('paths.secondary_images_folder_path')
        current_tertiary_folder = config_manager.get_setting('paths.tertiary_images_folder_path')
        
        # Добавляем пояснение
        st.markdown("### Пути к папкам с изображениями")
        st.markdown("Укажите пути к папкам с изображениями, которые нужно вставить в Excel файл. Поиск будет производиться в порядке приоритета.")
        
        # Отображаем текстовое поле для пути к основной папке с изображениями
        st.markdown("#### Основная папка с изображениями (приоритет 1)")
        image_folder = st.text_input(
            "Путь к основной папке с изображениями",
            value=st.session_state.get('images_folder_path', current_image_folder),
            help="Укажите полный путь к основной папке, где хранятся изображения товаров. Например: C:/Users/User/Downloads/images",
            key="image_folder_input"
        )
        
        # Создаем переменную для путей по умолчанию
        downloads_folder = get_downloads_folder()
        default_secondary_folder = r"\\10.10.100.2\Excel_reserve_photo"

        # Если вторая папка не указана, устанавливаем ее по умолчанию
        if not current_secondary_folder:
            config_manager.set_setting('paths.secondary_images_folder_path', default_secondary_folder)
            current_secondary_folder = default_secondary_folder
        
        # Отображаем текстовое поле для пути к второй папке с изображениями
        st.markdown("#### Запасная папка с изображениями (приоритет 2)")
        secondary_folder = st.text_input(
            "Путь к запасной папке с изображениями",
            value=st.session_state.get('secondary_images_folder_path', current_secondary_folder),
            help="Укажите полный путь к запасной папке с изображениями. Будет использована, если изображение не найдено в основной папке.",
            key="secondary_folder_input"
        )
        
        # Отображаем текстовое поле для пути к третьей папке с изображениями
        st.markdown("#### Дополнительная запасная папка (приоритет 3)")
        tertiary_folder = st.text_input(
            "Путь к дополнительной запасной папке",
            value=st.session_state.get('tertiary_images_folder_path', current_tertiary_folder),
            help="Укажите полный путь к дополнительной папке с изображениями. Будет использована, если изображение не найдено в основной и запасной папках.",
            key="tertiary_folder_input"
        )
        
        # Проверяем, является ли основной путь сетевым
        is_network_path = image_folder.startswith('\\\\') or image_folder.startswith('//')
        if is_network_path:
            st.warning("⚠️ Основной путь указывает на сетевой диск. Убедитесь, что у вас есть доступ к этой папке через проводник Windows.")
        
        # Проверяем доступность основного пути
        path_exists = os.path.exists(image_folder) if image_folder else False
        if path_exists:
            st.success("✅ Основной путь доступен")
        else:
            st.error("❌ Основной путь недоступен")
        
        # Проверяем доступность второго пути
        secondary_path_exists = os.path.exists(secondary_folder) if secondary_folder else False
        if secondary_path_exists:
            st.success("✅ Запасной путь доступен")
        else:
            st.warning("⚠️ Запасной путь недоступен")
            
        # Проверяем доступность третьего пути, если он указан
        if tertiary_folder:
            tertiary_path_exists = os.path.exists(tertiary_folder)
            if tertiary_path_exists:
                st.success("✅ Дополнительный запасной путь доступен")
            else:
                st.warning("⚠️ Дополнительный запасной путь недоступен")
        
        # Если основной путь изменился, сохраняем его в конфиг и session_state
        if image_folder != current_image_folder:
            config_manager.set_setting('paths.images_folder_path', image_folder)
            config_manager.save_settings("Default")
            # Сохраняем в session_state для сохранения между перезагрузками
            st.session_state.images_folder_path = image_folder
            log.info(f"Сохранен новый путь к основной папке с изображениями: {image_folder}")
            
        # Если второй путь изменился, сохраняем его в конфиг и session_state
        if secondary_folder != current_secondary_folder:
            config_manager.set_setting('paths.secondary_images_folder_path', secondary_folder)
            config_manager.save_settings("Default")
            # Сохраняем в session_state для сохранения между перезагрузками
            st.session_state.secondary_images_folder_path = secondary_folder
            log.info(f"Сохранен новый путь к запасной папке с изображениями: {secondary_folder}")
            
        # Если третий путь изменился, сохраняем его в конфиг и session_state
        if tertiary_folder != current_tertiary_folder:
            config_manager.set_setting('paths.tertiary_images_folder_path', tertiary_folder)
            config_manager.save_settings("Default")
            # Сохраняем в session_state для сохранения между перезагрузками
            st.session_state.tertiary_images_folder_path = tertiary_folder
            log.info(f"Сохранен новый путь к дополнительной запасной папке с изображениями: {tertiary_folder}")
            
        # Добавляем кнопку сброса путей к значениям по умолчанию
        if st.button("Сбросить пути к папкам изображений", 
                    help="Сбросить пути к папкам с изображениями на значения по умолчанию",
                    type="secondary"):
            config_manager.set_setting('paths.images_folder_path', downloads_folder)
            config_manager.set_setting('paths.secondary_images_folder_path', r"\\10.10.100.2\Excel_reserve_photo")
            config_manager.set_setting('paths.tertiary_images_folder_path', "")
            
            # Сохраняем в session_state для сохранения между перезагрузками
            st.session_state.images_folder_path = downloads_folder
            st.session_state.secondary_images_folder_path = r"\\10.10.100.2\Excel_reserve_photo"
            st.session_state.tertiary_images_folder_path = ""
            
            config_manager.save_settings("Default")
            st.success(f"Пути сброшены на значения по умолчанию")
            log.info(f"Пути сброшены на значения по умолчанию")
            st.rerun()
    
    # --- Настройки размера файла ---
    with st.sidebar.expander("Размер файла", expanded=True):
        st.subheader("Ограничение размера")
        
        max_total_file_size_mb = st.number_input(
            "Максимальный размер файла Excel (МБ)",
            min_value=1, # Minimum 1MB
            max_value=100, # Maximum 100MB
            value=int(config_manager.get_setting('excel_settings.max_total_file_size_mb', 20)), 
            step=1, # Step 1MB
            help="Приблизительный максимальный размер итогового Excel-файла. Изображения будут сжаты для достижения этого лимита.",
            key="max_total_file_size_mb_input"
        )
        if max_total_file_size_mb != config_manager.get_setting('excel_settings.max_total_file_size_mb', 20):
            config_manager.set_setting('excel_settings.max_total_file_size_mb', max_total_file_size_mb)
            config_manager.save_settings("Default")
            log.info(f"Настройка max_total_file_size_mb изменена на: {max_total_file_size_mb}")
    
    # Добавляем кнопку для полного сброса настроек
    st.sidebar.markdown("""
    <style>
    div[data-testid="stButton"] button[kind="secondary"] {
        background-color: #FF5555;
        color: white;
    }
    </style>
    """, unsafe_allow_html=True)
    
    if st.sidebar.button(
        'Сбросить все настройки', 
        key='sidebar_reset_all_button', 
        help="Сбрасывает все настройки к значениям по умолчанию",
        type="secondary"  # Используем secondary для применения стиля
    ):
        config_manager.reset_settings()
        #st.session_state['current_settings'] = config_manager.get_settings().current_settings
        
        # Устанавливаем путь к изображениям
        downloads_folder = get_downloads_folder()
        images_path = os.path.join(downloads_folder, "images")
        config_manager.set_setting("paths.images_folder_path", images_path)
        st.session_state.images_folder_path = images_path
        
        # Устанавливаем стандартные буквы колонок
        config_manager.set_setting("excel_settings.article_column", "A")
        config_manager.set_setting("excel_settings.image_column", "B")
        
        # Перезагружаем страницу для применения настроек
        st.rerun()

# Функция для проверки новых изображений в папке
def check_new_images_in_folder():
    """
    Проверяет и обновляет изображения в указанной папке.
    """
    images_folder = config_manager.get_setting("paths.images_folder_path", "")
    if not images_folder or not os.path.exists(images_folder):
        return
    
    st.info(f"Проверка изображений в папке: {images_folder}")
    # Здесь можно добавить код для обновления/сканирования изображений

# Добавляем новую функцию для показа настроек с пользовательским префиксом
def show_custom_settings(key_prefix="", use_expanders=True):
    """
    Отображает настройки с заданным префиксом ключей.
    
    Args:
        key_prefix (str): Префикс для ключей элементов
        use_expanders (bool): Использовать ли expanders для группировки настроек
    """
    # Функция для отображения настроек путей
    def show_paths_settings():
        """Отображение настроек путей"""
        st.subheader("Настройки путей")
        
        # Получаем текущие значения из конфига
        config_manager = st.session_state.config_manager
        current_image_folder = config_manager.get_setting('paths.images_folder_path')
        current_secondary_folder = config_manager.get_setting('paths.secondary_images_folder_path', '')
        current_tertiary_folder = config_manager.get_setting('paths.tertiary_images_folder_path', '')
        
        # Добавляем пояснение
        st.markdown("### Путь к папке с изображениями")
        st.markdown("Укажите пути к папкам с изображениями, которые нужно вставить в Excel файл.")
        
        # Основная папка с изображениями
        st.markdown("#### Основная папка")
        image_folder = st.text_input(
            "Путь к основной папке с изображениями",
            value=st.session_state.get('images_folder_path', current_image_folder),
            key=f"{key_prefix}image_folder_path",
            help="Укажите путь к папке, где хранятся изображения товаров (поиск будет выполнен сначала здесь)"
        )
        
        # Вторичная папка с изображениями
        st.markdown("#### Вторичная папка (резервная)")
        secondary_folder = st.text_input(
            "Путь к вторичной папке с изображениями",
            value=current_secondary_folder,
            key=f"{key_prefix}secondary_folder_path",
            help="Укажите путь к резервной папке с изображениями (поиск будет выполнен здесь, если изображение не найдено в основной папке)"
        )
        
        # Третичная папка с изображениями
        st.markdown("#### Третичная папка (дополнительная)")
        tertiary_folder = st.text_input(
            "Путь к третичной папке с изображениями",
            value=current_tertiary_folder,
            key=f"{key_prefix}tertiary_folder_path",
            help="Укажите путь к дополнительной папке с изображениями (поиск будет выполнен здесь в последнюю очередь)"
        )
        
        # Если путь изменился, сохраняем его в конфиг и session_state
        if image_folder != current_image_folder:
            config_manager.set_setting('paths.images_folder_path', image_folder)
            st.session_state.images_folder_path = image_folder
            config_manager.save_settings("Default")
            log.info(f"Сохранен новый путь к основной папке с изображениями: {image_folder}")
            
        # Сохраняем пути к вторичной и третичной папкам
        if secondary_folder != current_secondary_folder:
            config_manager.set_setting('paths.secondary_images_folder_path', secondary_folder)
            st.session_state.secondary_images_folder_path = secondary_folder
            config_manager.save_settings("Default")
            log.info(f"Сохранен новый путь к вторичной папке с изображениями: {secondary_folder}")
            
        if tertiary_folder != current_tertiary_folder:
            config_manager.set_setting('paths.tertiary_images_folder_path', tertiary_folder)
            st.session_state.tertiary_images_folder_path = tertiary_folder
            config_manager.save_settings("Default")
            log.info(f"Сохранен новый путь к третичной папке с изображениями: {tertiary_folder}")
        
        # Добавляем кнопку сброса пути к значениям по умолчанию
        if st.button("Сбросить пути к папкам", key=f"{key_prefix}reset_path_button"):
            downloads_folder = get_downloads_folder()
            config_manager.set_setting('paths.images_folder_path', downloads_folder)
            config_manager.set_setting('paths.secondary_images_folder_path', r"\\10.10.100.2\Excel_reserve_photo")
            config_manager.set_setting('paths.tertiary_images_folder_path', '')
            st.session_state.images_folder_path = downloads_folder
            st.session_state.secondary_images_folder_path = r"\\10.10.100.2\Excel_reserve_photo"
            st.session_state.tertiary_images_folder_path = ''
            config_manager.save_settings("Default")
            st.success(f"Пути сброшены. Основная папка: {downloads_folder}")
            log.info(f"Пути сброшены. Основная папка: {downloads_folder}")
    
    # Функция для отображения настроек изображений
    def show_image_settings():
        """Отображение настроек изображений"""
        st.subheader("Настройки изображений")
        
        # Получаем текущие значения из конфига
        config_manager = st.session_state.config_manager
        
        # Максимальный размер Excel-файла
        st.markdown("### Ограничение размера файла")
        st.markdown("""
        Укажите максимальный размер результирующего Excel-файла в мегабайтах.
        Изображения будут автоматически оптимизированы, чтобы общий размер файла не превышал указанное значение.
        """)
        
        max_file_size_mb = st.number_input(
            "Максимальный размер Excel-файла (МБ)",
            min_value=1,
            max_value=100,
            value=int(config_manager.get_setting('excel_settings.max_total_file_size_mb', 20)),
            key=f"{key_prefix}max_file_size_mb",
            help="Максимальный размер результирующего Excel-файла в мегабайтах"
        )
        
        if max_file_size_mb != config_manager.get_setting('excel_settings.max_total_file_size_mb'):
            config_manager.set_setting('excel_settings.max_total_file_size_mb', max_file_size_mb)
            config_manager.save_settings("Default")
            log.info(f"Установлен максимальный размер Excel-файла: {max_file_size_mb} МБ")
        
        # Качество изображения
        st.markdown("### Качество изображений")
        st.markdown("""
        Укажите качество сжатия изображений (от 1 до 100).
        Большее значение даёт лучшее качество, но увеличивает размер файла.
        """)
        
        quality = st.slider(
            "Качество изображений",
            min_value=1,
            max_value=100,
            value=int(config_manager.get_setting('image_settings.quality', 80)),
            key=f"{key_prefix}quality",
            help="Качество сжатия изображений (от 1 до 100)"
        )
        
        if quality != config_manager.get_setting('image_settings.quality'):
            config_manager.set_setting('image_settings.quality', quality)
            config_manager.save_settings("Default")
            log.info(f"Установлено качество изображений: {quality}")
        
        # Добавляем кнопку сброса настроек к значениям по умолчанию
        if st.button("Сбросить настройки изображений", key=f"{key_prefix}reset_image_settings"):
            config_manager.set_setting('excel_settings.max_total_file_size_mb', 20)
            config_manager.set_setting('image_settings.quality', 80)
            config_manager.save_settings("Default")
            st.success("Настройки изображений сброшены к значениям по умолчанию")
            log.info("Настройки изображений сброшены к значениям по умолчанию")
    
    # Используем expanders если разрешено, иначе просто показываем заголовки
    if use_expanders:
        # Добавляем настройки путей
        with st.expander("Настройки путей", expanded=True):
            show_paths_settings()
        
        # Настройки изображений
        with st.expander("Настройки изображений", expanded=True):
            show_image_settings()
    else:
        # Просто показываем заголовки и содержимое без expanders
        st.subheader("Настройки путей")
        show_paths_settings()
        
        st.subheader("Настройки изображений")
        show_image_settings()

# Функция для запуска процесса обработки через состояние сессии
def trigger_processing():
    """
    Устанавливает флаг для запуска обработки через session_state
    """
    log.info("Установка флага запуска обработки")
    st.session_state.start_processing = True

# Главная функция приложения
def main():
    """
    Главная функция приложения.
    """
    # Проверяем наличие необходимых модулей
    check_required_modules()
    
    # Инициализируем session_state
    initialize_session_state()
    
    # Проверяем, нужно ли отобразить отчет о результатах обработки
    if st.session_state.get('show_processing_report', False):
        show_processing_report()
        # Сбрасываем флаг, чтобы не показывать отчет повторно
        st.session_state.show_processing_report = False
    
    # Проверяем, нужен ли перезапуск страницы
    if st.session_state.get('needs_rerun', False):
        st.session_state['needs_rerun'] = False
        st.rerun()
    
    # Добавляем CSS для скрытия меню и футера
    st.markdown("""
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    .main .block-container {padding-top: 0.5rem;}
    .stDataFrame {overflow-x: auto;}
    </style>
    """, unsafe_allow_html=True)
    
    # --- Боковая панель ТОЛЬКО с настройками ---
    with st.sidebar:
        st.header("Настройки")
        settings_tab() # Показываем настройки путей, размера файла и т.д.
    # --- Конец Боковой панели ---

    # --- Главный раздел приложения ---
    st.title("📊 Excel Image Processor") # Добавим заголовок на главную
    file_uploader_section() # << Вызываем здесь, вне сайдбара
        
    # Проверка и обновление изображений в папке (если необходимо)
    settings = config_manager.get_config_manager().current_settings
    if settings and settings.get("check_images_on_startup", False):
        check_new_images_in_folder()

# Функция для проверки наличия требуемых модулей
def check_required_modules():
    """
    Проверяет наличие всех необходимых модулей для работы приложения.
    """
    # Список модулей для проверки
    required_modules = [
        ("core.processor", "Основной обработчик Excel файлов"),
        ("utils.image_utils", "Утилиты для работы с изображениями"),
        ("utils.excel_utils", "Утилиты для работы с Excel"),
        ("utils.config_manager", "Менеджер конфигурации")
    ]
    
    # Проверяем каждый модуль
    missing_modules = []
    for module_name, description in required_modules:
        try:
            __import__(module_name)
            log.info(f"Модуль {module_name} успешно импортирован")
        except ImportError as e:
            log.error(f"Ошибка импорта модуля {module_name}: {str(e)}")
            missing_modules.append((module_name, description))
    
    # Если есть отсутствующие модули, показываем ошибку
    if missing_modules:
        error_msg = "Не удалось импортировать необходимые модули:\n"
        for module, desc in missing_modules:
            error_msg += f"- {module} ({desc})\n"
        st.error(error_msg)
        st.warning("Приложение не может работать корректно без этих модулей. Проверьте установку.")
        
        # Выводим подробную информацию по решению проблемы
        with st.expander("Варианты решения проблемы"):
            st.markdown("""
            ### Проверьте структуру проекта
            Убедитесь, что структура проекта соответствует ожидаемой:
            ```
            ExcelWithImages/
            ├── core/
            │   └── processor.py
            ├── utils/
            │   ├── __init__.py
            │   ├── image_utils.py
            │   ├── excel_utils.py
            │   └── config_manager.py
            └── app/
                └── app.py
            ```
            
            ### Проверьте PYTHONPATH
            Модули должны быть доступны из корня проекта. Проверьте, что добавлено:
            ```python
            sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
            ```
            
            ### Логи приложения
            Проверьте логи для получения дополнительной информации:
            ```
            ExcelWithImages/logs/app_latest.log
            ```
            """)
        
        # Выводим текущий PYTHONPATH
        st.write("**Текущий PYTHONPATH:**")
        st.code("\n".join(sys.path))
        
        # Выводим текущую директорию
        st.write(f"**Текущая директория:** {os.getcwd()}")
        st.write(f"**Директория приложения:** {os.path.dirname(__file__)}")
        
        # Выводим содержимое папок с модулями
        try:
            core_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'core')
            utils_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'utils')
            
            st.write(f"**Содержимое директории core ({core_dir}):**")
            if os.path.exists(core_dir):
                st.code("\n".join(os.listdir(core_dir)))
            else:
                st.warning("Директория не существует")
                
            st.write(f"**Содержимое директории utils ({utils_dir}):**")
            if os.path.exists(utils_dir):
                st.code("\n".join(os.listdir(utils_dir)))
            else:
                st.warning("Директория не существует")
        except Exception as e:
            st.error(f"Ошибка при проверке директорий: {str(e)}")
            
        return False
    
    return True

def show_processing_report():
    """Отображает отчет о результатах обработки файла"""
    try:
        if 'processing_result' not in st.session_state or not st.session_state.processing_result:
            log.warning("Попытка показать отчет обработки, но результатов обработки нет")
            return
        
        result = st.session_state.processing_result
        
        # Распаковываем результат
        output_path = result.get('output_path', '')
        df = result.get('df', None)
        images_inserted = result.get('images_inserted', 0)
        multiple_images = result.get('multiple_images', {})
        not_found = result.get('not_found', [])
        image_search_results = result.get('image_search_results', [])
        
        # Создаем вкладки для удобного отображения информации
        report_tab, details_tab = st.tabs(["Общий отчет", "Подробности"])
        
        with report_tab:
            st.subheader("Отчет об обработке файла")
            
            # Информация о загруженном файле и процессе
            st.markdown("### Информация о процессе")
            
            # Проверяем, что temp_file_path существует и является строкой
            temp_file_path = st.session_state.get('temp_file_path', "Неизвестный файл")
            file_name = os.path.basename(temp_file_path) if isinstance(temp_file_path, str) else "Неизвестный файл"
            
            st.markdown(f"""
            **Исходный файл:** {file_name}  
            **Время обработки:** {datetime.now().strftime('%d.%m.%Y %H:%M:%S')}
            """)
            
            # Информация о папках с изображениями
            with st.expander("Информация о папках с изображениями", expanded=True):
                # Получаем информацию о папках из первого результата поиска
                folder_info = {}
                if image_search_results and len(image_search_results) > 0:
                    folder_info = image_search_results[0].get("image_folders", {})
                
                st.markdown("""
                #### Пути к папкам с изображениями
                Поиск изображений выполнялся в следующих папках (в порядке приоритета):
                """)
                
                # Основная папка
                primary_folder = folder_info.get("primary", 
                                               config_manager.get_setting("paths.images_folder_path", "Не задана"))
                st.markdown(f"1. **Основная папка:** `{primary_folder}`")
                
                # Вторичная папка
                secondary_folder = folder_info.get("secondary", 
                                                config_manager.get_setting("paths.secondary_images_folder_path", ""))
                if secondary_folder:
                    st.markdown(f"2. **Вторичная папка:** `{secondary_folder}`")
                else:
                    st.markdown("2. **Вторичная папка:** Не задана")
                
                # Третичная папка
                tertiary_folder = folder_info.get("tertiary", 
                                               config_manager.get_setting("paths.tertiary_images_folder_path", ""))
                if tertiary_folder:
                    st.markdown(f"3. **Третичная папка:** `{tertiary_folder}`")
                else:
                    st.markdown("3. **Третичная папка:** Не задана")
            
            # Базовая информация
            col1, col2 = st.columns(2)
            with col1:
                st.metric("Обработано строк", len(df) if df is not None else "N/A")
                st.metric("Вставлено изображений", images_inserted)
            with col2:
                st.metric("Изображения не найдены", len(not_found))
                st.metric("Найдено несколько вариантов", len(multiple_images))
            
            # Информация о выходном файле
            if output_path and os.path.exists(output_path):
                file_size_mb = os.path.getsize(output_path) / (1024 * 1024)
                st.success(f"Результат сохранен в файл: {output_path} (размер: {file_size_mb:.2f} МБ)")
                # Кнопка для скачивания файла
                with open(output_path, "rb") as file:
                    file_name = os.path.basename(output_path)
                    st.download_button(
                        label="Скачать обработанный файл",
                        data=file,
                        file_name=file_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            else:
                st.error(f"Выходной файл не найден или не был создан: {output_path}")

            # Сводка по источникам изображений (если есть данные о поиске)
            if image_search_results:
                st.subheader("Источники изображений")
                
                sources = {"primary": 0, "secondary": 0, "tertiary": 0, "not_found": 0}
                
                for result in image_search_results:
                    if not result.get("found", False):
                        sources["not_found"] += 1
                    else:
                        folder_priority = result.get("source_folder", "primary")
                        sources[folder_priority] += 1
                
                # Отображаем статистику по использованным источникам
                source_data = {
                    "Источник": ["Основная папка", "Вторичная папка", "Третичная папка", "Не найдено"],
                    "Количество": [sources["primary"], sources["secondary"], sources["tertiary"], sources["not_found"]]
                }
                
                source_df = pd.DataFrame(source_data)
                
                # Создаем столбчатую диаграмму
                fig = px.bar(
                    source_df, 
                    x='Источник', 
                    y='Количество',
                    color='Источник',
                    color_discrete_sequence=['#4CAF50', '#2196F3', '#FF9800', '#F44336'],
                    title="Распределение источников изображений"
                )
                st.plotly_chart(fig, use_container_width=True)
        
        with details_tab:
            # Поиск изображений в нескольких папках
            if image_search_results:
                st.subheader("Подробности поиска изображений")
                
                # Преобразуем результаты поиска в DataFrame для удобного отображения
                search_details = []
                for result in image_search_results:
                    row_index = result.get("row_index", "N/A")
                    article = result.get("original_article", result.get("article", ""))
                    found = result.get("found", False)
                    source = result.get("source_folder", "не найдено") if found else "не найдено"
                    source_text = {
                        1: "Основная папка",
                        2: "Вторичная папка",
                        3: "Третичная папка",
                        "не найдено": "Не найдено"
                    }.get(source, "Неизвестно")
                    
                    # Получаем информацию о папках поиска
                    image_folders = result.get("image_folders", {})
                    primary_folder = image_folders.get("primary", "")
                    secondary_folder = image_folders.get("secondary", "")
                    tertiary_folder = image_folders.get("tertiary", "")
                    
                    # Все пути поиска на одной строке для компактности
                    search_paths = []
                    if primary_folder:
                        search_paths.append(f"1) {primary_folder}")
                    if secondary_folder:
                        search_paths.append(f"2) {secondary_folder}")
                    if tertiary_folder:
                        search_paths.append(f"3) {tertiary_folder}")
                    
                    search_paths_text = "\n".join(search_paths)
                    
                    images_count = len(result.get("images", [])) if found else 0
                    found_paths = ", ".join(result.get("images", [])[:2])
                    if images_count > 2:
                        found_paths += f"... (и еще {images_count-2})"
                    
                    search_details.append({
                        "Строка": row_index + 1,  # +1 для человекочитаемой нумерации
                        "Артикул": article,
                        "Статус": "Найдено" if found else "Не найдено",
                        "Источник": source_text,
                        "Кол-во найдено": images_count,
                        "Путь к файлу": found_paths,
                        "Пути поиска": search_paths_text
                    })
                
                search_df = pd.DataFrame(search_details)
                # Сортируем по статусу (сначала не найденные), затем по строке
                search_df = search_df.sort_values(by=["Статус", "Строка"], ascending=[False, True])
                
                # Определяем цвета для разных статусов
                def highlight_status(val):
                    if val == "Не найдено":
                        return 'background-color: #FFCDD2'
                    else:
                        return ''
                
                def highlight_source(val):
                    if val == "Основная папка":
                        return 'background-color: #C8E6C9'
                    elif val == "Вторичная папка":
                        return 'background-color: #BBDEFB'
                    elif val == "Третичная папка":
                        return 'background-color: #FFE0B2'
                    elif val == "Не найдено":
                        return 'background-color: #FFCDD2'
                    else:
                        return ''
                
                # Выводим с подсветкой
                st.dataframe(
                    search_df.style
                    .applymap(highlight_status, subset=["Статус"])
                    .applymap(highlight_source, subset=["Источник"]),
                    use_container_width=True,
                    height=400  # Добавляем фиксированную высоту для лучшей читаемости
                )
                
                # Добавляем кнопку для скачивания отчета в CSV с более заметным оформлением
                st.markdown("### Скачать полный отчет о поиске изображений")
                st.write("Вы можете скачать полный отчет о поиске изображений в формате CSV для дальнейшего анализа в Excel или других программах.")
                
                csv = search_df.to_csv(index=False).encode('utf-8')
                col1, col2, col3 = st.columns([1, 2, 1])
                with col2:
                    st.download_button(
                        label="📥 Скачать полный отчет (CSV)",
                        data=csv,
                        file_name=f"image_search_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                        mime="text/csv",
                        help="Скачать полный отчет о поиске изображений в CSV формате",
                        key="download_report_button",
                        type="primary",  # Делаем кнопку primary для лучшей заметности
                    )
            
            # Информация о строках без изображений
            if not_found:
                st.subheader("Артикулы без изображений")
                st.info(f"Для {len(not_found)} артикулов не найдены изображения:")
                # Выводим список артикулов без найденных изображений
                st.write(", ".join(not_found) if len(not_found) < 20 else 
                         f"{', '.join(not_found[:20])}... и еще {len(not_found) - 20}")
                
                with st.expander("Показать полный список артикулов без изображений"):
                    for i, article in enumerate(not_found):
                        st.write(f"{i+1}. {article}")
                
                # Добавляем кнопку для скачивания списка ненайденных артикулов
                if len(not_found) > 0:
                    # Подготавливаем данные для скачивания
                    not_found_txt = "\n".join(not_found)
                    not_found_bytes = not_found_txt.encode('utf-8')
                    
                    st.download_button(
                        label="📥 Скачать список артикулов без изображений (TXT)",
                        data=not_found_bytes,
                        file_name=f"not_found_articles_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
                        mime="text/plain",
                        help="Скачать список артикулов, для которых не были найдены изображения",
                        key="download_not_found_button"
                    )
            
            # Информация о строках с несколькими вариантами изображений
            if multiple_images:
                st.subheader("Артикулы с несколькими вариантами изображений")
                st.info(f"Для {len(multiple_images)} артикулов найдено несколько вариантов изображений:")
                
                for article, paths in multiple_images.items():
                    with st.expander(f"{article} ({len(paths)} изображений)"):
                        for i, path in enumerate(paths):
                            status = "Выбрано" if i == 0 else "Альтернатива"
                            st.write(f"{status}: {path}")
    except Exception as e:
        log.error(f"Ошибка при отображении отчета: {e}")
        st.error(f"Произошла ошибка при отображении отчета: {e}")
        import traceback
        log.error(traceback.format_exc())

if __name__ == "__main__":
    main()
