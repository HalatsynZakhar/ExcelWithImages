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

# Добавляем корневую папку проекта в PYTHONPATH
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

# Используем относительные импорты вместо абсолютных
from utils import config_manager
from utils import excel_utils
from utils import image_utils
from utils.config_manager import get_downloads_folder

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
        "images_folder_path": os.path.join(get_downloads_folder(), "images")
    },
    "excel_settings": {
        "article_column": "A",
        "image_column": "B",
        "start_row": 1,
        "adjust_cell_size": False,
        "column_width": 150,
        "row_height": 120
    },
    "image_settings": {
        "max_file_size_mb": 20
    },
    "check_images_on_startup": False
}

# Инициализация менеджера конфигурации с созданием настроек по умолчанию
def init_config_manager(config_folder):
    """
    Инициализирует менеджер конфигурации и создает настройки по умолчанию, если их нет.
    
    Args:
        config_folder (str): Путь к папке для хранения настроек
    """
    # Проверяем наличие папки настроек
    os.makedirs(config_folder, exist_ok=True)
    
    # Определяем путь к файлу настроек по умолчанию
    default_settings_path = os.path.join(config_folder, 'default_settings.json')
    
    # Если файл настроек не существует, создаем его
    if not os.path.exists(default_settings_path):
        try:
            with open(default_settings_path, 'w', encoding='utf-8') as f:
                json.dump(default_settings, f, ensure_ascii=False, indent=4)
            log.info(f"Created default settings file: {default_settings_path}")
        except Exception as e:
            log.error(f"Error creating default settings file: {str(e)}")
    
    # Инициализируем менеджер конфигурации
    config_manager.init_config_manager(config_folder)
    
    # Применяем настройки по умолчанию, если текущих настроек нет
    if not config_manager.get_config_manager().current_settings:
        config_manager.get_config_manager().current_settings = default_settings
        log.info("Applied default settings")

# Обновляем код инициализации для использования нашей функции
config_folder = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'settings_presets')
init_config_manager(config_folder)

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
        os.makedirs(temp_dir, exist_ok=True)
        log.info(f"Использована системная временная директория: {temp_dir}")
    
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

# Функция для обновления кнопок в сайдбаре
def update_sidebar_buttons():
    # Кнопка сброса настроек - делаем её красной
    st.sidebar.markdown("""
    <style>
    div[data-testid="stButton"] button[kind="secondary"] {
        background-color: #FF5555;
        color: white;
    }
    </style>
    """, unsafe_allow_html=True)
    
    if st.sidebar.button(
        'Сбросить настройки', 
        key='reset_button', 
        help="Сбрасывает все настройки к значениям по умолчанию",
        type="secondary"  # Используем secondary для применения стиля
    ):
        config_manager.reset_settings()
        st.session_state['current_settings'] = config_manager.get_config_manager().current_settings
        # Сбрасываем пути к папкам по умолчанию
        config_manager.set_setting("paths.output_folder_path", get_downloads_folder())
        config_manager.set_setting("paths.images_folder_path", os.path.join(get_downloads_folder(), "images"))
        # Устанавливаем значения по умолчанию для колонок
        config_manager.set_setting("excel_settings.article_column", "A")
        config_manager.set_setting("excel_settings.image_column", "B")
        
        # Вместо прямого вызова st.rerun() устанавливаем флаг
        st.session_state['needs_rerun'] = True

# Функция для отображения настроек
def show_settings():
    # Убираем дублирующий подзаголовок
    # st.sidebar.subheader("Настройки")
    
    # Добавляем настройки путей в боковую панель
    with st.sidebar.expander("Настройки путей", expanded=True):
        # Папка с изображениями
        images_folder = st.sidebar.text_input(
            "Папка с изображениями",
            value=config_manager.get_setting("paths.images_folder_path", ""),
            help="Укажите путь к папке, где находятся изображения для вставки",
            key="images_folder_input"
        )
        
        if images_folder:
            # Проверяем, существует ли папка
            if not os.path.exists(images_folder):
                st.sidebar.warning(f"Папка {images_folder} не существует!")
            else:
                st.sidebar.success(f"Папка с изображениями: {images_folder}")
                # Подсчитываем количество изображений в папке
                image_files = [f for f in os.listdir(images_folder) 
                              if os.path.isfile(os.path.join(images_folder, f)) and 
                              f.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp'))]
                st.sidebar.info(f"Найдено {len(image_files)} изображений в папке")
            
            # Сохраняем путь в настройках
            config_manager.set_setting("paths.images_folder_path", images_folder)
    
    # Настройки изображений
    with st.sidebar.expander("Настройки изображений", expanded=True):
        # Максимальный размер файла изображения
        max_file_size_mb = st.sidebar.number_input(
            "Максимальный размер файла (МБ)",
            min_value=1,
            max_value=50,
            value=config_manager.get_setting("image_settings.max_file_size_mb", 20),
            help="Максимально допустимый размер файла изображения в мегабайтах",
            key="max_file_size_input",
            label_visibility="visible"
        )
        config_manager.set_setting("image_settings.max_file_size_mb", int(max_file_size_mb))
        
        # Настройка изменения размера ячейки
        adjust_cell_size = st.sidebar.checkbox(
            "Изменить размер ячейки",
            value=config_manager.get_setting("excel_settings.adjust_cell_size", False),
            help="Если включено, размеры ячеек будут изменены для лучшего отображения изображений",
            key="adjust_cell_size_input",
            label_visibility="visible"
        )
        config_manager.set_setting("excel_settings.adjust_cell_size", adjust_cell_size)
        
        # Показываем настройки размеров только если включено изменение размера ячейки
        if adjust_cell_size:
            col1, col2 = st.sidebar.columns(2)
            
            # Ширина колонки
            column_width = col1.number_input(
                "Ширина колонки (пикс.)",
                min_value=50,
                max_value=500,
                value=config_manager.get_setting("excel_settings.column_width", 150),
                help="Желаемая ширина колонки с изображениями в пикселях",
                key="column_width_input",
                label_visibility="visible"
            )
            config_manager.set_setting("excel_settings.column_width", int(column_width))
            
            # Высота строки
            row_height = col2.number_input(
                "Высота строки (пикс.)",
                min_value=30,
                max_value=500,
                value=config_manager.get_setting("excel_settings.row_height", 120),
                help="Желаемая высота строки с изображениями в пикселях",
                key="row_height_input",
                label_visibility="visible"
            )
            config_manager.set_setting("excel_settings.row_height", int(row_height))
            
        # Начальная строка
        start_row = st.sidebar.number_input(
            "Начальная строка",
            min_value=1,
            value=config_manager.get_setting("excel_settings.start_row", 1),
            help="Номер строки, с которой начнется обработка (по умолчанию с первой строки)",
            key="start_row_input",
            label_visibility="visible"
        )
        config_manager.set_setting("excel_settings.start_row", int(start_row))

# Функция для отображения предпросмотра таблицы
def show_table_preview(df):
    """Показывает превью загруженной таблицы.
    
    Args:
        df (pd.DataFrame): DataFrame для отображения
    """
    with st.expander("Предпросмотр загруженного файла", expanded=True):
        # Общее количество строк
        st.write(f"**Общее количество строк:** {len(df)}")
        
        # Получаем количество непустых значений в каждой колонке
        column_stats = {}
        for col in df.columns:
            if col not in ['image', 'image_path', 'image_width', 'image_height']:
                non_empty_count = df[col].notna().sum()
                column_stats[col] = non_empty_count
        
        # Выводим статистику по колонкам
        st.write("**Количество непустых значений в колонках:**")
        stats_df = pd.DataFrame({
            'Колонка': column_stats.keys(),
            'Непустых значений': column_stats.values()
        })
        st.dataframe(stats_df, use_container_width=True)
        
        # Показываем первые 5 строк таблицы
        st.write("**Первые 5 строк:**")
        st.dataframe(df.head(), use_container_width=True)

# Функция для загрузки Excel файла
def load_excel_file(uploaded_file=None):
    """
    Загружает Excel файл, определяет доступные листы и загружает выбранный лист.
    
    Args:
        uploaded_file: Загруженный пользователем файл из file_uploader
    """
    try:
        # Если файл не передан, пробуем взять его из file_uploader
        if uploaded_file is None:
            uploaded_file = st.session_state.get('file_uploader')
            
        if uploaded_file is None:
            st.session_state.df = None
            st.session_state.processing_error = None
            st.session_state.temp_file_path = None
            st.session_state.available_sheets = []
            st.session_state.selected_sheet = None
            return

        # Сохраняем загруженный файл во временную директорию
        temp_dir = ensure_temp_dir()
        temp_file_path = os.path.join(temp_dir, uploaded_file.name)
        
        with open(temp_file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        
        log.info(f"Файл {uploaded_file.name} успешно сохранен: {temp_file_path}")
        st.session_state.temp_file_path = temp_file_path
        
        # Получаем опции загрузки из session_state
        skiprows = st.session_state.get('skiprows', 0)
        header_row = st.session_state.get('header_row', 0)
        log.info(f"Загрузка с параметрами: skiprows={skiprows}, header_row={header_row}")
        
        # Получаем список доступных листов
        try:
            with pd.ExcelFile(temp_file_path) as xls:
                available_sheets = xls.sheet_names
                log.info(f"Доступные листы в файле: {available_sheets}")
                
                if not available_sheets:
                    error_msg = "Excel-файл не содержит листов"
                    log.warning(error_msg)
                    add_log_message(error_msg, "ERROR")
                    st.session_state.processing_error = error_msg
                    st.session_state.df = None
                    st.session_state.available_sheets = []
                    return
                
                st.session_state.available_sheets = available_sheets
                
                # Если лист не выбран или выбранный лист отсутствует в списке, выбираем первый
                if (st.session_state.selected_sheet is None or 
                    st.session_state.selected_sheet not in available_sheets):
                    st.session_state.selected_sheet = available_sheets[0]
                    
                # Загружаем выбранный лист
                selected_sheet = st.session_state.selected_sheet
                log.info(f"Загрузка листа: {selected_sheet}")
                
                try:
                    # Пробуем прочитать с явным указанием движка и параметров
                    df = pd.read_excel(
                        temp_file_path, 
                        sheet_name=selected_sheet, 
                        engine='openpyxl',
                        skiprows=skiprows,
                        header=header_row
                    )
                    
                    # Проверка на пустой DataFrame
                    log.info(f"Размер данных: строк={df.shape[0]}, колонок={df.shape[1]}; пустой={df.empty}")
                    
                    if df.empty:
                        error_msg = f"Лист '{selected_sheet}' не содержит данных"
                        log.warning(error_msg)
                        add_log_message(error_msg, "ERROR")
                        st.session_state.processing_error = error_msg
                        st.session_state.df = None
                        return
                    
                    if df.shape[0] == 0:
                        error_msg = f"Лист '{selected_sheet}' не содержит строк с данными"
                        log.warning(error_msg)
                        add_log_message(error_msg, "ERROR")
                        st.session_state.processing_error = error_msg
                        st.session_state.df = None
                        return
                        
                    if df.shape[1] == 0:
                        error_msg = f"Лист '{selected_sheet}' не содержит колонок с данными"
                        log.warning(error_msg)
                        add_log_message(error_msg, "ERROR")
                        st.session_state.processing_error = error_msg
                        st.session_state.df = None
                        return
                    
                    # Проверка на файл, который имеет колонки, но все значения в них NaN
                    if df.notna().sum().sum() == 0:
                        error_msg = f"Лист '{selected_sheet}' содержит только пустые ячейки"
                        log.warning(error_msg)
                        add_log_message(error_msg, "ERROR")
                        st.session_state.processing_error = error_msg
                        st.session_state.df = None
                        return
                    
                    # Все хорошо, сохраняем DataFrame
                    st.session_state.df = df
                    st.session_state.processing_error = None
                    log.info(f"Файл успешно загружен. Найдено {len(df)} строк и {len(df.columns)} колонок")
                    add_log_message(f"Файл успешно загружен: {len(df)} строк, {len(df.columns)} колонок", "SUCCESS")
                
                except Exception as sheet_error:
                    error_msg = f"Ошибка при чтении листа '{selected_sheet}': {str(sheet_error)}"
                    log.error(error_msg)
                    add_log_message(error_msg, "ERROR")
                    st.session_state.processing_error = error_msg
                    st.session_state.df = None
                
        except Exception as e:
            error_msg = f"Ошибка при чтении файла Excel: {str(e)}"
            log.error(error_msg)
            add_log_message(error_msg, "ERROR")
            st.session_state.processing_error = error_msg
            st.session_state.df = None
            
    except Exception as e:
        error_msg = f"Ошибка при загрузке файла: {str(e)}"
        log.error(error_msg)
        add_log_message(error_msg, "ERROR")
        st.session_state.processing_error = error_msg
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
    
    # Проверяем наличие DataFrame
    if st.session_state.df is None:
        log_msgs.append("DataFrame не загружен")
        valid = False
    else:
        log_msgs.append(f"DataFrame загружен, размер: {st.session_state.df.shape}")
        
    # Проверяем, выбран ли лист в Excel
    if st.session_state.selected_sheet is None:
        log_msgs.append("Лист Excel не выбран")
        valid = False
    else:
        log_msgs.append(f"Выбран лист: {st.session_state.selected_sheet}")
        
    # Проверяем, указана ли папка с изображениями
    images_folder = config_manager.get_setting("paths.images_folder_path", "")
    if not images_folder:
        log_msgs.append("Не указана папка с изображениями")
        valid = False
    elif not os.path.isdir(images_folder):
        log_msgs.append(f"Указанная папка с изображениями не существует: {images_folder}")
        valid = False
    else:
        # Получаем количество изображений в папке
        image_files = [f for f in os.listdir(images_folder) 
                      if os.path.isfile(os.path.join(images_folder, f)) and 
                      f.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp'))]
        log_msgs.append(f"Папка с изображениями: {images_folder}, найдено {len(image_files)} изображений")
        
    # Проверяем колонки
    article_column = st.session_state.get("article_column")
    image_column = st.session_state.get("image_column")
    
    if not article_column:
        log_msgs.append("Не выбрана колонка с артикулами")
        valid = False
    else:
        log_msgs.append(f"Выбрана колонка с артикулами: {article_column}")
        
    if not image_column:
        log_msgs.append("Не выбрана колонка с изображениями")
        valid = False
    else:
        log_msgs.append(f"Выбрана колонка с изображениями: {image_column}")
        
    if article_column and image_column and article_column == image_column:
        log_msgs.append("Колонки с артикулами и изображениями не могут быть одинаковыми")
        valid = False
        
    # Проверяем временный файл
    if not st.session_state.temp_file_path:
        log_msgs.append("Временный файл не создан")
        valid = False
    elif not os.path.exists(st.session_state.temp_file_path):
        log_msgs.append(f"Временный файл не существует: {st.session_state.temp_file_path}")
        valid = False
    else:
        log_msgs.append(f"Временный файл доступен: {st.session_state.temp_file_path}")
    
    # Логируем результаты проверки
    for msg in log_msgs:
        log.info(f"Проверка: {msg}")
    
    log.info(f"Все условия выполнены: {valid}")
    return valid

# Функция для обработки изменения выбранного листа
def handle_sheet_change():
    """
    Обрабатывает изменение выбранного листа Excel и перезагружает данные.
    """
    if st.session_state.get("sheet_selector") != st.session_state.selected_sheet:
        st.session_state.selected_sheet = st.session_state.get("sheet_selector")
        log.info(f"Выбран новый лист: {st.session_state.selected_sheet}")
        
        # Перезагружаем данные с нового листа
        if st.session_state.temp_file_path and os.path.exists(st.session_state.temp_file_path):
            try:
                selected_sheet = st.session_state.selected_sheet
                skiprows = st.session_state.get('skiprows', 0)
                header_row = st.session_state.get('header_row', 0)
                
                df = pd.read_excel(
                    st.session_state.temp_file_path, 
                    sheet_name=selected_sheet, 
                    engine='openpyxl',
                    skiprows=skiprows,
                    header=header_row
                )
                
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
                error_msg = f"Ошибка при загрузке листа '{st.session_state.selected_sheet}': {str(e)}"
                log.error(error_msg)
                st.session_state.processing_error = error_msg
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
            
            # Сохраняем файл во временную директорию если это новый файл
            current_temp_path = st.session_state.get('temp_file_path', '')
            temp_filename = uploaded_file.name if current_temp_path else ''
            
            if not current_temp_path or os.path.basename(current_temp_path) != uploaded_file.name:
                # Новый файл - нужно обработать
                temp_dir = ensure_temp_dir()
                temp_file_path = os.path.join(temp_dir, uploaded_file.name)
                
                with open(temp_file_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())
                
                st.session_state.temp_file_path = temp_file_path
                add_log_message(f"Файл сохранен: {temp_file_path}", "INFO")
                # Загружаем файл
                load_excel_file()
            
            # Настройки загрузки Excel
            with st.expander("Настройки загрузки Excel", expanded=False):
                st.markdown("#### Параметры загрузки Excel")
                st.write("Настройте параметры чтения файла, если в нем есть метаданные или особый формат.")
                
                # Инициализация переменных
                if 'skiprows' not in st.session_state:
                    st.session_state.skiprows = 0
                if 'header_row' not in st.session_state:
                    st.session_state.header_row = 0
                    
                # Колонки для настроек
                col1, col2 = st.columns(2)
                with col1:
                    skiprows = st.number_input(
                        "Пропустить начальные строки", 
                        min_value=0, 
                        max_value=50, 
                        value=st.session_state.skiprows,
                        help="Укажите количество строк для пропуска в начале файла",
                        key="excel_skiprows"
                    )
                with col2:
                    header_row = st.number_input(
                        "Строка с заголовками", 
                        min_value=0, 
                        max_value=50, 
                        value=st.session_state.header_row,
                        help="Укажите номер строки с заголовками колонок (0 = первая непропущенная строка)",
                        key="excel_header_row"
                    )
                    
                # Если значения изменились, перезагружаем файл
                if (skiprows != st.session_state.skiprows or 
                    header_row != st.session_state.header_row):
                    st.session_state.skiprows = skiprows
                    st.session_state.header_row = header_row
                    st.button("Применить настройки", on_click=load_excel_file)
                    
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
                    # Позволяем пользователю выбрать колонки для обработки
                    col1, col2 = st.columns(2)
                    with col1:
                        st.selectbox("Колонка с артикулами", 
                                    options=["A", "B", "C", "D", "E", "F", "G", "H", "I", "J"],
                                    index=0,
                                    key="article_column")
                    with col2:
                        st.selectbox("Колонка с изображениями", 
                                    options=["A", "B", "C", "D", "E", "F", "G", "H", "I", "J"],
                                    index=1,
                                    key="image_column")
                                    
                    # Проверка всех необходимых полей перед обработкой
                    process_button_disabled = not all_inputs_valid()
                    
                    # Заменяем JavaScript-кнопку на стандартную кнопку Streamlit
                    if st.button("Обработать файл", disabled=process_button_disabled, type="primary"):
                        with st.spinner("Обработка файла..."):
                            st.session_state.is_processing = True
                            success = process_files()
                            if success:
                                st.success("Файл успешно обработан!")
                            else:
                                st.error(f"Ошибка при обработке файла: {st.session_state.processing_error}")
                    
                else:
                    st.warning("Файл не содержит колонок для выбора. Проверьте структуру Excel-файла.")
                    
        # Отображение результатов обработки
        if st.session_state.processing_result:
            st.success(st.session_state.processing_result)
            
        # Добавление кнопки скачивания, если файл был обработан
        if st.session_state.output_file_path and os.path.exists(st.session_state.output_file_path):
            with open(st.session_state.output_file_path, "rb") as file:
                st.download_button(
                    label="Скачать обработанный файл",
                    data=file,
                    file_name=os.path.basename(st.session_state.output_file_path),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
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
        
        # Получаем необходимые настройки
        images_folder = config_manager.get_setting("paths.images_folder_path", "")
        
        # Создаем директорию для хранения результатов
        temp_dir = ensure_temp_dir()
        output_folder = temp_dir
        
        add_log_message(f"Папка изображений: {images_folder}", "INFO")
        
        # Детальная проверка всех условий
        conditions = {
            "DataFrame загружен": st.session_state.df is not None,
            "Временный файл существует": st.session_state.temp_file_path is not None,
            "Файл доступен": os.path.exists(st.session_state.temp_file_path) if st.session_state.temp_file_path else False,
            "Выбран лист": st.session_state.selected_sheet is not None,
            "Указана колонка с артикулами": st.session_state.get('article_column') is not None,
            "Указана колонка с изображениями": st.session_state.get('image_column') is not None,
            "Папка изображений указана": images_folder != "",
            "Папка изображений существует": os.path.exists(images_folder) if images_folder else False
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
        article_column = st.session_state.get('article_column')
        image_column = st.session_state.get('image_column')
        selected_sheet = st.session_state.selected_sheet
        
        log.info(f"Параметры обработки:")
        log.info(f"- Файл: {excel_file_path}")
        log.info(f"- Лист: {selected_sheet}")
        log.info(f"- Колонка с артикулами: {article_column}")
        log.info(f"- Колонка с изображениями: {image_column}")
        log.info(f"- Папка с изображениями: {images_folder}")
        
        add_log_message(f"Обработка файла: {os.path.basename(excel_file_path)}, лист: {selected_sheet}", "INFO")
        add_log_message(f"Колонки: артикулы - {article_column}, изображения - {image_column}", "INFO")
        
        # Создаем имя выходного файла
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"processed_{timestamp}_{os.path.basename(excel_file_path)}"
        output_file_path = os.path.join(output_folder, output_filename)
        log.info(f"Выходной файл будет: {output_file_path}")
        add_log_message(f"Подготовка выходного файла: {output_filename}", "INFO")
        
        # Проверяем наличие модуля обработки
        try:
            from core.processor import process_excel_file
            log.info("Модуль process_excel_file успешно импортирован")
        except ImportError as e:
            error_msg = f"Ошибка импорта модуля обработки: {str(e)}"
            log.error(error_msg)
            add_log_message(error_msg, "ERROR")
            st.session_state.processing_error = error_msg
            return False
        
        try:
            log.info("Вызываем process_excel_file с аргументами:")
            log.info(f"- excel_file_path={excel_file_path}")
            log.info(f"- article_column={article_column}")
            log.info(f"- image_column={image_column}")
            log.info(f"- image_folder={images_folder}")
            log.info(f"- sheet_name={selected_sheet}")
            
            add_log_message("Запуск обработки файла...", "INFO")
            
            # Распаковываем результат функции process_excel_file
            result_file_path, result_df, images_inserted = process_excel_file(
                file_path=excel_file_path,
                article_column=article_column,
                image_folder=images_folder,
                image_column_name=image_column,
                sheet_name=selected_sheet,
                output_folder=output_folder
            )
            
            log.info(f"Файл успешно обработан: {result_file_path}")
            log.info(f"Вставлено изображений: {images_inserted}")
            
            add_log_message(f"Обработка завершена. Вставлено изображений: {images_inserted}", "SUCCESS")
            
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
            st.session_state.processing_result = success_msg
            
            log.info("===================== ОБРАБОТКА ЗАВЕРШЕНА УСПЕШНО =====================")
            return True
            
        except Exception as e:
            error_msg = f"Ошибка при обработке файла: {str(e)}"
            log.error(error_msg)
            log.exception("Детали ошибки:")
            add_log_message(error_msg, "ERROR")
            st.session_state.processing_error = error_msg
            log.info("===================== ОБРАБОТКА ЗАВЕРШЕНА С ОШИБКОЙ =====================")
            return False
            
    except Exception as e:
        error_msg = f"Ошибка в процессе обработки: {str(e)}"
        log.error(error_msg)
        log.exception("Детали ошибки:")
        add_log_message(error_msg, "ERROR")
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
    if 'available_sheets' not in st.session_state:
        st.session_state.available_sheets = []
    if 'force_rerun' not in st.session_state:
        st.session_state.force_rerun = False
    if 'start_processing' not in st.session_state:
        st.session_state.start_processing = False
    if 'needs_rerun' not in st.session_state:
        st.session_state.needs_rerun = False
    if 'log_messages' not in st.session_state:
        st.session_state.log_messages = []
    # Добавляем параметры загрузки Excel
    if 'skiprows' not in st.session_state:
        st.session_state.skiprows = 0
    if 'header_row' not in st.session_state:
        st.session_state.header_row = 0

# Функция для отображения вкладки настроек
def settings_tab():
    """
    Отображает вкладку настроек в боковой панели.
    """
    # Показываем текущие настройки с префиксом sidebar_
    show_custom_settings("sidebar_")
    
    # Добавляем кнопки для управления настройками
    update_sidebar_buttons()

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
        # Папка с изображениями
        images_folder = st.text_input(
            "Папка с изображениями",
            value=config_manager.get_setting("paths.images_folder_path", ""),
            help="Укажите путь к папке, где находятся изображения для вставки",
            key=f"{key_prefix}images_folder_input"
        )
        
        if images_folder:
            # Проверяем, существует ли папка
            if not os.path.exists(images_folder):
                st.warning(f"Папка {images_folder} не существует!")
            else:
                st.success(f"Папка с изображениями: {images_folder}")
                # Подсчитываем количество изображений в папке
                image_files = [f for f in os.listdir(images_folder) 
                              if os.path.isfile(os.path.join(images_folder, f)) and 
                              f.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp'))]
                st.info(f"Найдено {len(image_files)} изображений в папке")
            
            # Сохраняем путь в настройках
            config_manager.set_setting("paths.images_folder_path", images_folder)
    
    # Функция для отображения настроек изображений
    def show_image_settings():
        # Максимальный размер файла изображения
        max_file_size_mb = st.number_input(
            "Максимальный размер файла (МБ)",
            min_value=1,
            max_value=50,
            value=config_manager.get_setting("image_settings.max_file_size_mb", 20),
            help="Максимально допустимый размер файла изображения в мегабайтах",
            key=f"{key_prefix}max_file_size_input",
            label_visibility="visible"
        )
        config_manager.set_setting("image_settings.max_file_size_mb", int(max_file_size_mb))
        
        # Настройка изменения размера ячейки
        adjust_cell_size = st.checkbox(
            "Изменить размер ячейки",
            value=config_manager.get_setting("excel_settings.adjust_cell_size", False),
            help="Если включено, размеры ячеек будут изменены для лучшего отображения изображений",
            key=f"{key_prefix}adjust_cell_size_input",
            label_visibility="visible"
        )
        config_manager.set_setting("excel_settings.adjust_cell_size", adjust_cell_size)
        
        # Показываем настройки размеров только если включено изменение размера ячейки
        if adjust_cell_size:
            col1, col2 = st.columns(2)
            
            # Ширина колонки
            column_width = col1.number_input(
                "Ширина колонки (пикс.)",
                min_value=50,
                max_value=500,
                value=config_manager.get_setting("excel_settings.column_width", 150),
                help="Желаемая ширина колонки с изображениями в пикселях",
                key=f"{key_prefix}column_width_input",
                label_visibility="visible"
            )
            config_manager.set_setting("excel_settings.column_width", int(column_width))
            
            # Высота строки
            row_height = col2.number_input(
                "Высота строки (пикс.)",
                min_value=30,
                max_value=500,
                value=config_manager.get_setting("excel_settings.row_height", 120),
                help="Желаемая высота строки с изображениями в пикселях",
                key=f"{key_prefix}row_height_input",
                label_visibility="visible"
            )
            config_manager.set_setting("excel_settings.row_height", int(row_height))
            
        # Начальная строка
        start_row = st.number_input(
            "Начальная строка",
            min_value=1,
            value=config_manager.get_setting("excel_settings.start_row", 1),
            help="Номер строки, с которой начнется обработка (по умолчанию с первой строки)",
            key=f"{key_prefix}start_row_input",
            label_visibility="visible"
        )
        config_manager.set_setting("excel_settings.start_row", int(start_row))
    
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

# Главная функция приложения
def main():
    """
    Главная функция приложения.
    """
    # Проверяем наличие необходимых модулей
    check_required_modules()
    
    # Инициализируем session_state
    initialize_session_state()
    
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
    </style>
    """, unsafe_allow_html=True)
    
    # Боковая панель с настройками
    with st.sidebar:
        st.header("Настройки")
        settings_tab()
    
    # Главный раздел приложения
    file_uploader_section()
    
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

if __name__ == "__main__":
    main() 