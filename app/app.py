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
        "output_folder_path": get_downloads_folder(),
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
    temp_dir = os.path.join(tempfile.gettempdir(), f"{prefix}excelwithimages")
    os.makedirs(temp_dir, exist_ok=True)
    return temp_dir

# Функция для очистки временных файлов
def cleanup_temp_files():
    """
    Очищает временные файлы, сохраняя только текущий файл в session_state (если есть).
    """
    try:
        # Получаем путь к текущему временному файлу (если есть)
        current_temp_file = st.session_state.get('temp_file_path', None)
        
        # Определяем путь к временной директории
        temp_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'temp')
        if not os.path.exists(temp_dir):
            return
            
        # Удаляем все файлы, кроме текущего
        for filename in os.listdir(temp_dir):
            file_path = os.path.join(temp_dir, filename)
            if os.path.isfile(file_path) and file_path != current_temp_file:
                try:
                    os.remove(file_path)
                    log.info(f"Удален временный файл: {file_path}")
                except Exception as e:
                    log.error(f"Ошибка при удалении временного файла {file_path}: {e}")
                    
        log.info("Очистка временных файлов завершена")
    except Exception as e:
        log.error(f"Ошибка при очистке временных файлов: {e}")

# Вызываем очистку временных файлов при запуске приложения
cleanup_temp_files()

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
        # Получаем путь к папке загрузок пользователя
        default_downloads = get_downloads_folder()
        
        # Выходная папка
        output_folder = st.sidebar.text_input(
            "Папка для сохранения результата",
            value=config_manager.get_setting("paths.output_folder_path", default_downloads),
            help="Укажите путь к папке, где будет сохранен обработанный файл Excel",
            key="output_folder_input"
        )
        
        if output_folder:
            # Проверяем, существует ли папка
            if not os.path.exists(output_folder):
                st.sidebar.warning(f"Папка {output_folder} не существует. Она будет создана при обработке файла.")
            else:
                st.sidebar.success(f"Папка для сохранения: {output_folder}")
            
            # Сохраняем путь в настройках
            config_manager.set_setting("paths.output_folder_path", output_folder)
        
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

# Функция для загрузки файла Excel
def file_uploader_section():
    """
    Раздел для загрузки файлов.
    """
    # Проверяем, нужна ли перезагрузка после смены листа
    if st.session_state.get('force_rerun', False):
        st.session_state.force_rerun = False
        
    # CSS стили для кнопок и элементов интерфейса
    st.markdown("""
    <style>
    /* Удаляем красную рамку вокруг file uploader */
    .stFileUploader {
        border: none !important;
    }
    
    /* Большая зеленая кнопка */
    .big-green-button {
        background-color: #4CAF50;
        color: white;
        padding: 15px 25px;
        font-size: 18px;
        font-weight: bold;
        border-radius: 8px;
        border: none;
        cursor: pointer;
        width: 100%;
        text-align: center;
        margin: 20px 0;
        transition: all 0.3s;
    }
    
    .big-green-button:hover {
        background-color: #3e8e41;
    }
    
    /* Неактивная кнопка */
    .inactive-button {
        background-color: white;
        color: #4CAF50;
        padding: 12px 20px;
        font-size: 16px;
        border-radius: 5px;
        border: 2px solid #4CAF50;
        cursor: not-allowed;
        width: 100%;
        text-align: center;
        margin: 10px 0;
    }
    </style>
    """, unsafe_allow_html=True)
    
    # Инициализация переменных session_state если они не существуют
    # Не инициализируем uploaded_file, так как Streamlit управляет им сам
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
    
    def load_excel_file():
        """Загрузка Excel файла и создание DataFrame"""
        # Используем uploaded_file непосредственно из функции, а не из session_state
        uploaded_file = st.session_state.uploaded_file
        if uploaded_file is not None:
            # Создаем временную директорию, если она не существует
            temp_dir = os.path.join(os.getcwd(), "temp")
            os.makedirs(temp_dir, exist_ok=True)
            
            # Сохраняем файл во временной директории
            file_extension = os.path.splitext(uploaded_file.name)[1]
            temp_file_path = os.path.join(temp_dir, f"temp_file_{int(time.time())}{file_extension}")
            
            with open(temp_file_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
            
            # Получаем список листов в Excel файле
            try:
                excel = pd.ExcelFile(temp_file_path)
                st.session_state.available_sheets = excel.sheet_names
                
                # Если есть только один лист, выбираем его автоматически
                if len(st.session_state.available_sheets) == 1:
                    st.session_state.selected_sheet = st.session_state.available_sheets[0]
                else:
                    # Если лист не выбран, выберем первый
                    if 'selected_sheet' not in st.session_state or not st.session_state.selected_sheet:
                        st.session_state.selected_sheet = st.session_state.available_sheets[0]
                
                # Читаем Excel файл с выбранным листом
                df = pd.read_excel(temp_file_path, sheet_name=st.session_state.selected_sheet)
                
                # Проверяем, есть ли у столбцов имена и если нет - создаем имена по буквам (A, B, C, ...)
                # Преобразуем все колонки к строковому типу для безопасного использования str.contains
                renamed_columns = False
                
                # Получаем список столбцов, которые нужно переименовать
                unnamed_cols = []
                for i, col in enumerate(df.columns):
                    col_str = str(col)
                    if 'Unnamed:' in col_str or pd.isna(col):
                        unnamed_cols.append(i)
                
                # Если есть столбцы без имен, переименовываем их
                if unnamed_cols:
                    # Создаем новые имена для столбцов
                    new_columns = list(df.columns)
                    for i in unnamed_cols:
                        new_columns[i] = get_column_letter(i+1)  # A, B, C, ...
                    
                    df.columns = new_columns
                    renamed_columns = True
                    log.info(f"Renamed columns without headers: {[get_column_letter(i+1) for i in unnamed_cols]}")
                
                st.session_state.df = df
                st.session_state.temp_file_path = temp_file_path
                st.session_state.processing_error = None
                
                # Сбрасываем выбранные колонки при загрузке нового файла
                if 'article_column' in st.session_state:
                    del st.session_state.article_column
                if 'image_column' in st.session_state:
                    del st.session_state.image_column
                
                log.info(f"File {uploaded_file.name} loaded successfully. Sheets found: {len(st.session_state.available_sheets)}")
                return True
            except Exception as e:
                st.session_state.processing_error = f"Error reading file: {str(e)}"
                log.error(f"Error reading file: {str(e)}")
                return False
        return False
    
    def select_sheet():
        """Обработчик выбора листа Excel"""
        if st.session_state.temp_file_path and st.session_state.selected_sheet:
            try:
                # Используем тот же код, что и в load_excel_file, чтобы правильно обработать столбцы
                df = pd.read_excel(st.session_state.temp_file_path, sheet_name=st.session_state.selected_sheet)
                
                # Преобразуем все колонки к строковому типу для безопасного использования str.contains
                renamed_columns = False
                
                # Получаем список столбцов, которые нужно переименовать
                unnamed_cols = []
                for i, col in enumerate(df.columns):
                    col_str = str(col)
                    if 'Unnamed:' in col_str or pd.isna(col):
                        unnamed_cols.append(i)
                
                # Если есть столбцы без имен, переименовываем их
                if unnamed_cols:
                    # Создаем новые имена для столбцов
                    new_columns = list(df.columns)
                    for i in unnamed_cols:
                        new_columns[i] = get_column_letter(i+1)  # A, B, C, ...
                    
                    df.columns = new_columns
                    renamed_columns = True
                    log.info(f"Renamed columns without headers: {[get_column_letter(i+1) for i in unnamed_cols]}")
                
                st.session_state.df = df
                
                # Не сбрасываем выбранные колонки при смене листа
                # Оставляем текущие настройки как есть
                
                log.info(f"Selected sheet: {st.session_state.selected_sheet}")
                
                # Принудительно делаем rerun для обновления интерфейса
                # Устанавливаем флаг, что нужно перезагрузить страницу
                st.session_state.force_rerun = True
                st.session_state.needs_rerun = True
            except Exception as e:
                st.session_state.processing_error = f"Error reading sheet: {str(e)}"
                log.error(f"Error reading sheet: {str(e)}")
    
    def all_inputs_valid():
        """Проверка валидности всех входных данных"""
        # Проверяем наличие DataFrame
        if st.session_state.df is None:
            return False
            
        # Проверяем выбранные колонки
        if 'article_column_letter' not in st.session_state or not st.session_state.article_column_letter:
            return False
            
        # Проверяем папку с изображениями
        image_folder = config_manager.get_setting("paths.images_folder_path", "")
        if image_folder == "" or not os.path.exists(image_folder):
            return False
            
        return True
    
    # Загрузка файла - НЕ устанавливаем uploaded_file через session_state
    uploaded_file = st.file_uploader("Загрузите Excel файл", type=["xlsx", "xls"], key="uploaded_file", on_change=load_excel_file)
    
    # Если файл загружен, показываем выбор листа
    if uploaded_file is not None and st.session_state.available_sheets:
        # Отображаем информацию о загруженном файле
        st.success(f"Файл загружен: {uploaded_file.name}")
        
        # Выбор листа (если файл содержит несколько листов)
        if len(st.session_state.available_sheets) > 1:
            st.selectbox(
                "Выберите лист Excel",
                options=st.session_state.available_sheets,
                key="selected_sheet",
                on_change=select_sheet
            )
        else:
            st.info(f"Выбран лист: {st.session_state.selected_sheet}")
    
    # Если DataFrame загружен, показываем форму для настройки обработки
    if st.session_state.df is not None:
        df = st.session_state.df
        
        # Отображаем информацию о данных
        st.write(f"Количество строк: {len(df)} | Количество колонок: {len(df.columns)}")
        
        col1, col2 = st.columns(2)
        
        with col1:
            # Поле ввода для колонки с артикулами
            st.text_input(
                "Введите букву колонки с артикулами (A, B, C, ...)", 
                value="A",
                key="article_column_letter", 
                help="Укажите букву колонки, в которой находятся артикулы"
            )
            
            # Устанавливаем имя колонки в session_state
            if 'article_column_letter' in st.session_state and st.session_state.article_column_letter:
                col_idx = excel_utils.column_letter_to_index(st.session_state.article_column_letter)
                if col_idx < len(df.columns):
                    st.session_state.article_column = df.columns[col_idx]
        
        with col2:
            # Поле ввода для колонки с изображениями
            st.text_input(
                "Введите букву колонки для вставки изображений", 
                value="B",
                key="image_column_letter",
                help="Укажите букву колонки, в которую будут вставлены изображения"
            )
            
            # Устанавливаем имя колонки в session_state
            if 'image_column_letter' in st.session_state and st.session_state.image_column_letter:
                col_idx = excel_utils.column_letter_to_index(st.session_state.image_column_letter)
                if col_idx < len(df.columns):
                    st.session_state.image_column = df.columns[col_idx]
            else:
                st.session_state.image_column = ""
        
        # Отображаем текущие настройки, включая путь к изображениям
        with st.expander("Текущие настройки обработки", expanded=True):
            st.write("**Настройки данных:**")
            st.write(f"- Выбранный лист: **{st.session_state.selected_sheet}**")
            
            # Отображаем буквенные обозначения колонок
            article_column_letter = st.session_state.get('article_column_letter', 'A')
            image_column_letter = st.session_state.get('image_column_letter', 'B')
            
            st.write(f"- Колонка с артикулами: **{article_column_letter}**")
            st.write(f"- Колонка для вставки изображений: **{image_column_letter}**")
            
            st.write("**Настройки путей:**")
            images_folder = config_manager.get_setting("paths.images_folder_path", "")
            output_folder = config_manager.get_setting("paths.output_folder_path", "")
            
            if images_folder and os.path.exists(images_folder):
                st.write(f"- Папка с изображениями: **{images_folder}** ✅")
            else:
                st.write("- Папка с изображениями: **Не указана или не существует** ❌")
            
            if output_folder:
                if not os.path.exists(output_folder):
                    st.write(f"- Папка для сохранения результата: **{output_folder}** ⚠️ (будет создана)")
                else:
                    st.write(f"- Папка для сохранения результата: **{output_folder}** ✅")
        
        # Отображение предпросмотра таблицы
        show_table_preview(df)
        
        # Контейнер для кнопки и сообщений об ошибках
        st.subheader("Запуск обработки")
        
        # Проверяем валидность данных для активации кнопки
        button_enabled = all_inputs_valid()
        
        if not button_enabled:
            # Выводим подсказку, что нужно сделать для активации кнопки
            missing = []
            if st.session_state.df is None:
                missing.append("- Загрузите Excel файл")
            if 'article_column_letter' not in st.session_state or not st.session_state.article_column_letter:
                missing.append("- Введите букву колонки с артикулами")
            if not config_manager.get_setting("paths.images_folder_path", "") or not os.path.exists(config_manager.get_setting("paths.images_folder_path", "")):
                missing.append("- Укажите корректный путь к папке с изображениями в настройках")
            
            if missing:
                st.warning("Для запуска обработки необходимо:\n" + "\n".join(missing))
        
        # Создаем кнопку с соответствующим классом
        start_processing = st.button(
            "ЗАПУСТИТЬ ОБРАБОТКУ ФАЙЛА", 
            disabled=not button_enabled or st.session_state.is_processing,
            key="start_processing_button", 
            use_container_width=True,
            type="primary" if button_enabled else "secondary"
        )
        
        # Отображение сообщений об ошибках
        if st.session_state.processing_error:
            st.error(st.session_state.processing_error)
        
        # Отображение результатов обработки
        if st.session_state.processing_result:
            st.success(st.session_state.processing_result)
            
            if st.session_state.output_file_path and os.path.exists(st.session_state.output_file_path):
                # Используем безопасный способ открытия файла
                try:
                    with open(st.session_state.output_file_path, "rb") as f:
                        file_data = f.read()
                    
                    st.download_button(
                        label="Скачать обработанный файл",
                        data=file_data,
                        file_name=os.path.basename(st.session_state.output_file_path),
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                except Exception as e:
                    st.error(f"Ошибка при подготовке файла для скачивания: {str(e)}")
        
        # Обработка клика по кнопке
        if start_processing and not st.session_state.is_processing:
            st.session_state.is_processing = True
            st.session_state.processing_result = None
            st.session_state.processing_error = None
            
            try:
                # Вызов функции обработки из основного модуля
                # Используем корректные импорты
                import sys
                current_dir = os.path.dirname(os.path.abspath(__file__))
                parent_dir = os.path.dirname(current_dir)
                if parent_dir not in sys.path:
                    sys.path.append(parent_dir)
                
                from core.processor import process_excel_file
                
                # Получаем фактические имена колонок из DataFrame
                article_column = st.session_state.article_column_letter
                image_column = st.session_state.image_column_letter
                
                output_file = process_excel_file(
                    excel_file_path=st.session_state.temp_file_path,
                    article_column=article_column,
                    image_column=image_column,
                    image_folder=config_manager.get_setting("paths.images_folder_path", ""),
                    sheet_name=st.session_state.selected_sheet
                )
                
                st.session_state.output_file_path = output_file
                st.session_state.processing_result = f"Файл успешно обработан: {output_file}"
                log.info(f"Обработка файла завершена успешно. Результат сохранен в: {output_file}")
            except Exception as e:
                error_msg = f"Ошибка при обработке файла: {str(e)}"
                st.session_state.processing_error = error_msg
                log.error(error_msg)
            finally:
                st.session_state.is_processing = False
                # Устанавливаем флаг для перезагрузки страницы
                st.session_state.needs_rerun = True
    
    # Добавляем окно логов
    with st.expander("Logs", expanded=False):
        # Получаем содержимое лог-файла
        try:
            if os.path.exists(log_file):
                with open(log_file, 'r', encoding='utf-8') as f:
                    logs = f.readlines()
                    # Выводим только последние 50 строк логов
                    log_text = "".join(logs[-50:])
                    st.code(log_text, language="text")
            else:
                st.warning("Log file not found")
        except Exception as e:
            # Если возникает ошибка, пытаемся прочитать в двоичном режиме и декодировать с игнорированием ошибок
            try:
                with open(log_file, 'rb') as f:
                    binary_logs = f.read()
                    # Декодируем с игнорированием ошибок
                    log_text = binary_logs.decode('utf-8', errors='replace')
                    st.code(log_text[-5000:], language="text")
                    st.warning("Log file is displayed with replacement characters for characters that could not be decoded")
            except Exception as e2:
                st.error(f"Error reading log file: {str(e2)}")

# Функция для проверки готовности к обработке (все ли поля заполнены)
def is_ready_for_processing():
    # Проверяем, есть ли все необходимые настройки
    input_file = config_manager.get_setting("paths.input_file_path", "")
    output_folder = config_manager.get_setting("paths.output_folder_path", "")
    images_folder = config_manager.get_setting("paths.images_folder_path", "")
    
    # Проверяем, загружен ли файл
    if not input_file or not os.path.isfile(input_file):
        return False
    
    # Проверяем, указана ли папка с изображениями
    if not images_folder or not os.path.isdir(images_folder):
        return False
    
    # Проверяем, указана ли выходная папка
    if not output_folder:
        return False
    
    return True

# Функция для обработки файла
def process_files():
    """
    Основная функция для обработки файлов.
    Обрабатывает Excel файл и добавляет изображения к соответствующим артикулам.
    """
    try:
        log.info("Начало обработки файлов")
        st.session_state['is_processing'] = True
        
        # Проверка наличия всех необходимых данных
        if not all([
            st.session_state.get('df') is not None,
            st.session_state.get('excel_file_path'),
            st.session_state.get('article_column'),
            st.session_state.get('image_column'),
            st.session_state.get('image_folder_path')
        ]):
            raise ValueError("Не все необходимые данные указаны для обработки")
            
        # Получение параметров из состояния сессии
        df = st.session_state['df'].copy()
        excel_file_path = st.session_state['excel_file_path']
        article_column = st.session_state['article_column']
        image_column = st.session_state['image_column']
        image_folder_path = st.session_state['image_folder_path']
        
        # Создание временной директории для обработанных изображений
        temp_dir = ensure_temp_dir(prefix="excel_images_")
        log.info(f"Создана временная директория для обработанных изображений: {temp_dir}")
        
        # Путь к результирующему файлу
        result_filename = f"processed_{os.path.basename(excel_file_path)}"
        result_file_path = os.path.join(temp_dir, result_filename)
        
        # Массив для хранения путей ко всем обработанным изображениям
        processed_image_paths = []
        total_rows = len(df)
        
        # Обработка каждой строки DataFrame
        for idx, row in df.iterrows():
            # Обновление прогресса
            progress = (idx + 1) / total_rows
            log.debug(f"Обработка строки {idx+1}/{total_rows} ({progress:.1%})")
            
            try:
                # Получение и нормализация артикула
                article = row[article_column]
                if pd.isna(article):
                    log.warning(f"Строка {idx+1}: артикул отсутствует, пропускаем")
                    continue
                    
                normalized_article = image_utils.normalize_article_number(str(article))
                
                # Поиск изображений для артикула
                image_paths = image_utils.find_images_by_article(image_folder_path, normalized_article)
                
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
                if processed_images:
                    df.at[idx, image_column] = ",".join(processed_images)
            except Exception as row_err:
                log.error(f"Ошибка при обработке строки {idx+1}: {row_err}")
        
        # Сохранение обработанного DataFrame в новый Excel файл
        log.info(f"Сохранение результата в файл: {result_file_path}")
        try:
            # Используем ExcelWriter для сохранения с изображениями
            with pd.ExcelWriter(result_file_path, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
                
                # Вставка изображений в файл Excel
                excel_utils.insert_images_to_excel(writer, df, image_column)
                
            # Убеждаемся, что файл сохранен до удаления временных файлов
            if os.path.exists(result_file_path):
                log.info(f"Файл {result_file_path} успешно создан")
                st.session_state['result_file_path'] = result_file_path
                st.session_state['process_complete'] = True
            else:
                raise FileNotFoundError(f"Не удалось создать файл {result_file_path}")
                
        except Exception as excel_err:
            log.error(f"Ошибка при сохранении Excel файла: {excel_err}")
            raise
            
        log.info("Обработка файлов завершена успешно")
        return True
    except Exception as e:
        log.error(f"Ошибка при обработке файлов: {e}")
        st.error(f"Ошибка при обработке файлов: {str(e)}")
        return False
    finally:
        # Сбрасываем флаг обработки в любом случае
        st.session_state['is_processing'] = False

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

# Функция для отображения вкладки настроек
def settings_tab():
    """
    Отображает вкладку настроек в боковой панели.
    """
    # Показываем текущие настройки
    show_settings()
    
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

# Главная функция приложения
def main():
    """
    Главная функция приложения.
    """
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

if __name__ == "__main__":
    main() 