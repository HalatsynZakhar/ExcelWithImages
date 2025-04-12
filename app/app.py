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

# Добавляем корневую папку проекта в PYTHONPATH
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

# Используем относительные импорты вместо абсолютных
from utils import config_manager
from utils import excel_utils
from openpyxl import load_workbook
from PIL import Image as PILImage

# Определяем папку загрузок пользователя по умолчанию
def get_downloads_folder():
    """Возвращает путь к папке загрузок пользователя"""
    if os.name == 'nt':  # Windows
        import winreg
        sub_key = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders'
        downloads_guid = '{374DE290-123F-4565-9164-39C4925E467B}'
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, sub_key) as key:
            try:
                return winreg.QueryValueEx(key, downloads_guid)[0]
            except:
                return os.path.join(os.path.expanduser('~'), 'Downloads')
    else:  # Linux, macOS и другие
        return os.path.join(os.path.expanduser('~'), 'Downloads')

# Настройка логирования
log_stream = io.StringIO()
log_handler = logging.StreamHandler(log_stream)
log_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(name)s - %(message)s'))
log_handler.setLevel(logging.INFO)

file_handler = logging.FileHandler(os.path.join(os.path.dirname(os.path.dirname(__file__)), 'logs', f'app_{datetime.now().strftime("%Y%m%d_%H%M%S")}.log'))
file_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(name)s - %(message)s'))
file_handler.setLevel(logging.DEBUG)

root_logger = logging.getLogger()
root_logger.setLevel(logging.DEBUG)
root_logger.addHandler(log_handler)
root_logger.addHandler(file_handler)

log = logging.getLogger(__name__)

# Инициализация менеджера конфигурации
config_folder = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'settings_presets')
config_manager.init_config_manager(config_folder)

# Настройка параметров приложения
st.set_page_config(
    page_title="ExcelWithImages - Вставка изображений в Excel",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Функция для создания временных директорий
def ensure_temp_dir() -> str:
    """
    Создает и возвращает путь к временной директории.
    
    Returns:
        Путь к временной директории
    """
    temp_dir = os.path.join(tempfile.gettempdir(), 'excelwithimages')
    os.makedirs(temp_dir, exist_ok=True)
    return temp_dir

# Функция для обновления кнопок в сайдбаре
def update_sidebar_buttons():
    # Кнопка сброса настроек
    if st.sidebar.button('Сбросить настройки', key='reset_button'):
        config_manager.reset_settings()
        st.session_state['current_settings'] = config_manager.get_config_manager().current_settings
        st.rerun()
    
    # Работа с пресетами
    st.sidebar.subheader("Управление пресетами")
    
    # Получаем список пресетов
    presets = config_manager.get_presets_list()
    
    # Ввод имени пресета
    preset_name = st.sidebar.text_input("Имя пресета", 
                                        value=config_manager.get_config_manager().current_preset_name,
                                        key='preset_name')
    
    # Кнопки для работы с пресетами
    col1, col2 = st.sidebar.columns(2)
    
    if col1.button('Сохранить пресет', key='save_preset'):
        if preset_name:
            config_manager.save_settings(preset_name)
            st.sidebar.success(f"Пресет '{preset_name}' сохранен!")
            st.rerun()
        else:
            st.sidebar.error("Укажите имя пресета!")
    
    if col2.button('Удалить пресет', key='delete_preset'):
        if preset_name in presets:
            if config_manager.delete_preset(preset_name):
                st.sidebar.success(f"Пресет '{preset_name}' удален!")
                st.rerun()
        else:
            st.sidebar.error("Выберите существующий пресет для удаления!")
    
    # Выбор пресета
    if presets:
        selected_preset = st.sidebar.selectbox(
            "Выберите пресет", 
            options=presets,
            key='selected_preset'
        )
        
        if st.sidebar.button('Загрузить пресет', key='load_preset'):
            if config_manager.load_settings(selected_preset):
                st.session_state['current_settings'] = config_manager.get_config_manager().current_settings
                st.sidebar.success(f"Пресет '{selected_preset}' загружен!")
                st.rerun()
            else:
                st.sidebar.error(f"Не удалось загрузить пресет '{selected_preset}'!")

# Функция для отображения настроек
def show_settings():
    st.sidebar.subheader("Настройки")
    
    # Настройки Excel
    with st.sidebar.expander("Настройки Excel", expanded=True):
        col1, col2 = st.sidebar.columns(2)
        
        # Столбец с артикулами
        article_column = col1.text_input(
            "Столбец с артикулами",
            value=config_manager.get_setting("excel_settings.article_column", "C"),
            help="Буква столбца, содержащего артикулы товаров"
        )
        config_manager.set_setting("excel_settings.article_column", article_column.upper())
        
        # Столбец для вставки изображений
        image_column = col2.text_input(
            "Столбец для вставки изображений",
            value=config_manager.get_setting("excel_settings.image_column", "A"),
            help="Буква столбца, куда будут вставлены изображения"
        )
        config_manager.set_setting("excel_settings.image_column", image_column.upper())
        
        # Начальная строка
        start_row = col1.number_input(
            "Начальная строка",
            min_value=1,
            value=config_manager.get_setting("excel_settings.start_row", 2),
            help="Номер строки, с которой начнется обработка (обычно 2, т.к. в первой строке заголовки)"
        )
        config_manager.set_setting("excel_settings.start_row", int(start_row))
        
        # Добавляем выбор номера листа
        sheet_index = col2.number_input(
            "Номер листа",
            min_value=1,
            value=config_manager.get_setting("excel_settings.sheet_index", 1),
            help="Номер листа в Excel файле (начиная с 1)"
        )
        config_manager.set_setting("excel_settings.sheet_index", int(sheet_index))
        
        # Автоматическая настройка размеров
        adjust_dimensions = col1.checkbox(
            "Настраивать размеры строк и столбцов",
            value=config_manager.get_setting("excel_settings.adjust_dimensions", True),
            help="Автоматически изменять высоту строк и ширину столбцов под размер изображений"
        )
        config_manager.set_setting("excel_settings.adjust_dimensions", adjust_dimensions)
    
    # Настройки изображений
    with st.sidebar.expander("Настройки изображений", expanded=True):
        col1, col2 = st.sidebar.columns(2)
        
        # Максимальный размер изображения
        max_size_kb = col1.number_input(
            "Максимальный размер (КБ)",
            min_value=10,
            max_value=1000,
            value=config_manager.get_setting("image_settings.max_size_kb", 100),
            help="Максимальный размер изображения в килобайтах. Большие изображения будут сжаты."
        )
        config_manager.set_setting("image_settings.max_size_kb", int(max_size_kb))
        
        # Качество JPEG
        quality = col2.slider(
            "Качество JPEG",
            min_value=30,
            max_value=100,
            value=config_manager.get_setting("image_settings.quality", 90),
            help="Начальное качество JPEG при сжатии изображений"
        )
        config_manager.set_setting("image_settings.quality", int(quality))
        
        # Целевые размеры изображения
        target_width = col1.number_input(
            "Целевая ширина (пикс.)",
            min_value=50,
            max_value=1000,
            value=config_manager.get_setting("image_settings.target_width", 300),
            help="Целевая ширина изображения в пикселях"
        )
        config_manager.set_setting("image_settings.target_width", int(target_width))
        
        target_height = col2.number_input(
            "Целевая высота (пикс.)",
            min_value=50,
            max_value=1000,
            value=config_manager.get_setting("image_settings.target_height", 300),
            help="Целевая высота изображения в пикселях"
        )
        config_manager.set_setting("image_settings.target_height", int(target_height))
        
        # Поддерживаемые расширения файлов
        supported_extensions = st.sidebar.multiselect(
            "Поддерживаемые типы файлов",
            options=[".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tiff", ".webp"],
            default=config_manager.get_setting("image_settings.supported_extensions", [".jpg", ".jpeg", ".png", ".gif", ".bmp"]),
            help="Типы файлов изображений, которые будут обрабатываться"
        )
        config_manager.set_setting("image_settings.supported_extensions", supported_extensions)
    
    # UI настройки
    with st.sidebar.expander("Настройки интерфейса", expanded=False):
        col1, col2 = st.sidebar.columns(2)
        
        # Показывать предпросмотр
        show_preview = col1.checkbox(
            "Показывать предпросмотр",
            value=config_manager.get_setting("ui_settings.show_preview", True),
            help="Показывать предпросмотр найденных изображений"
        )
        config_manager.set_setting("ui_settings.show_preview", show_preview)
        
        # Показывать статистику
        show_stats = col2.checkbox(
            "Показывать статистику",
            value=config_manager.get_setting("ui_settings.show_stats", True),
            help="Показывать статистику обработки"
        )
        config_manager.set_setting("ui_settings.show_stats", show_stats)
        
        # Тема
        theme = col1.selectbox(
            "Тема оформления",
            options=["light", "dark"],
            index=0 if config_manager.get_setting("ui_settings.theme", "light") == "light" else 1,
            help="Тема оформления интерфейса"
        )
        config_manager.set_setting("ui_settings.theme", theme)

# Функция для загрузки файла Excel
def file_uploader_section():
    st.subheader("Загрузка файла Excel")
    
    # Загрузка Excel файла
    uploaded_file = st.file_uploader(
        "Выберите файл Excel",
        type=["xlsx", "xls"],
        help="Загрузите файл Excel, в который нужно вставить изображения"
    )
    
    if uploaded_file is not None:
        # Сохраняем файл во временную директорию
        temp_dir = ensure_temp_dir()
        temp_file = os.path.join(temp_dir, uploaded_file.name)
        
        with open(temp_file, "wb") as f:
            f.write(uploaded_file.getbuffer())
        
        # Сохраняем путь к файлу в настройках
        config_manager.set_setting("paths.input_file_path", temp_file)
        
        # Сохраняем информацию о загрузке файла в session_state
        st.session_state['excel_uploaded'] = True
        st.session_state['excel_file_path'] = temp_file
        st.session_state['excel_file_name'] = uploaded_file.name
        
        # Показываем информацию о файле
        st.success(f"Файл {uploaded_file.name} успешно загружен!")
        
        try:
            # Загружаем файл для отображения информации
            wb = load_workbook(temp_file, read_only=True)
            
            # Получаем список листов
            sheet_names = wb.sheetnames
            
            # Сохраняем список листов в session_state
            st.session_state['excel_sheets'] = sheet_names
            
            # Выбор листа для анализа
            selected_sheet_index = st.selectbox(
                "Выберите лист для анализа",
                options=list(range(1, len(sheet_names) + 1)),
                format_func=lambda x: f"{x}. {sheet_names[x-1]}",
                index=config_manager.get_setting("excel_settings.sheet_index", 1) - 1,
                key="preview_sheet_index"
            )
            
            # Устанавливаем выбранный лист как активный для последующей обработки
            config_manager.set_setting("excel_settings.sheet_index", selected_sheet_index)
            
            # Получаем выбранный лист
            active_sheet = wb[sheet_names[selected_sheet_index - 1]]
            
            # Получаем информацию о листе
            col1, col2 = st.columns(2)
            col1.write(f"**Имя листа:** {active_sheet.title}")
            col1.write(f"**Размер:** {active_sheet.max_row} строк × {active_sheet.max_column} столбцов")
            
            # Получаем данные для таблицы
            article_column = config_manager.get_setting("excel_settings.article_column", "C")
            image_column = config_manager.get_setting("excel_settings.image_column", "A")
            start_row = config_manager.get_setting("excel_settings.start_row", 2)
            
            col2.write(f"**Столбец с артикулами:** {article_column}")
            col2.write(f"**Столбец для изображений:** {image_column}")
            col2.write(f"**Начальная строка:** {start_row}")
            
            # Отображаем предупреждение, если максимальная колонка меньше, чем выбранный столбец
            max_col_letter = get_column_letter(active_sheet.max_column) if hasattr(active_sheet, 'max_column') else 'A'
            if ord(article_column.upper()) - ord('A') + 1 > active_sheet.max_column:
                st.warning(f"⚠️ Столбец артикулов '{article_column}' находится за пределами данных (максимум: {max_col_letter})")
            
            if ord(image_column.upper()) - ord('A') + 1 > active_sheet.max_column:
                st.warning(f"⚠️ Столбец изображений '{image_column}' находится за пределами данных (максимум: {max_col_letter})")
            
            # Отображаем первые 5 строк в таблице
            rows_data = []
            
            # Получаем заголовки столбцов (первая строка)
            headers = []
            for cell in active_sheet[1]:
                headers.append(cell.value if cell.value is not None else "")
            
            # Получаем данные (строки 2-6)
            for row in active_sheet.iter_rows(min_row=2, max_row=6):
                row_data = []
                for cell in row:
                    row_data.append(cell.value if cell.value is not None else "")
                rows_data.append(row_data)
            
            # Отображаем данные в таблице
            st.write("Предпросмотр данных:")
            df_preview = pd.DataFrame(rows_data, columns=headers)
            st.dataframe(df_preview)
            
            # Отображаем данные из столбца с артикулами
            if len(rows_data) > 0:
                article_col_index = ord(article_column.upper()) - ord('A')
                if 0 <= article_col_index < len(headers):
                    article_values = []
                    for row in rows_data:
                        # Проверяем, что индекс находится в пределах размера строки
                        if article_col_index < len(row):
                            article_values.append(row[article_col_index])
                        else:
                            article_values.append(None)
                    
                    articles_found = sum(1 for a in article_values if a)
                    
                    if articles_found > 0:
                        st.success(f"Найдено {articles_found} артикулов в столбце {article_column} в предпросмотре.")
                    else:
                        st.warning(f"В столбце {article_column} не найдено артикулов. Проверьте правильность выбора столбца с артикулами.")
                        
                        # Предлагаем другие листы, если они есть
                        if len(sheet_names) > 1:
                            other_sheets = [name for i, name in enumerate(sheet_names) if i+1 != selected_sheet_index]
                            st.info(f"💡 Попробуйте выбрать другой лист: {', '.join(other_sheets)}")
                else:
                    st.warning(f"Столбец {article_column} находится за пределами таблицы. В файле только {len(headers)} столбцов.")
            
            # Закрываем файл
            wb.close()
        except Exception as e:
            st.error(f"Ошибка при чтении файла: {e}")
            log.error(f"Ошибка при чтении файла: {e}")
    else:
        # Если файл не загружен, сбрасываем соответствующие флаги
        st.session_state['excel_uploaded'] = False
        st.session_state.pop('excel_file_path', None)
        st.session_state.pop('excel_file_name', None)
        st.session_state.pop('excel_sheets', None)

# Функция для выбора выходной папки
def output_folder_section():
    st.subheader("Папка для сохранения результата")
    
    # Получаем путь к папке загрузок пользователя
    default_downloads = get_downloads_folder()
    
    # Получаем текущую выходную папку
    current_output_folder = config_manager.get_setting("paths.output_folder_path", default_downloads)
    
    # Поле для ввода пути
    output_folder = st.text_input(
        "Путь к папке для сохранения результата",
        value=current_output_folder,
        help="Укажите путь к папке, где будет сохранен обработанный файл Excel"
    )
    
    if output_folder:
        # Проверяем, существует ли папка
        if not os.path.exists(output_folder):
            st.warning(f"Папка {output_folder} не существует. Она будет создана при обработке файла.")
        
        # Сохраняем путь в настройках
        config_manager.set_setting("paths.output_folder_path", output_folder)
    
    # Помощь по выбору папки
    with st.expander("Как указать путь к папке"):
        st.write("""
        Вы можете ввести полный путь к папке, например:
        - C:\\Users\\Username\\Documents\\Excel_Results
        - /home/username/documents/excel_results
        
        Или относительный путь от текущей папки:
        - ./results
        - ../output
        """)

# Функция для выбора папки с изображениями
def images_folder_section():
    st.subheader("Папка с изображениями")
    
    # Получаем текущую папку с изображениями
    current_images_folder = config_manager.get_setting("paths.images_folder_path", "")
    
    # Поле для ввода пути
    images_folder = st.text_input(
        "Путь к папке с изображениями",
        value=current_images_folder,
        help="Укажите путь к папке, где находятся изображения для вставки"
    )
    
    if images_folder:
        # Проверяем, существует ли папка
        if os.path.exists(images_folder):
            # Сохраняем путь в настройках
            config_manager.set_setting("paths.images_folder_path", images_folder)
            
            # Если включен предпросмотр, показываем несколько изображений
            if config_manager.get_setting("ui_settings.show_preview", True):
                # Получаем список файлов изображений
                supported_extensions = tuple(config_manager.get_setting(
                    "image_settings.supported_extensions", 
                    [".jpg", ".jpeg", ".png", ".gif", ".bmp"]
                ))
                
                image_files = []
                for filename in os.listdir(images_folder):
                    if any(filename.lower().endswith(ext) for ext in supported_extensions):
                        image_files.append(os.path.join(images_folder, filename))
                
                # Показываем до 5 изображений
                if image_files:
                    st.write(f"Найдено {len(image_files)} изображений. Примеры:")
                    
                    # Отображаем изображения в несколько колонок
                    columns = st.columns(min(5, len(image_files)))
                    for i, img_path in enumerate(image_files[:5]):
                        try:
                            img = PILImage.open(img_path)
                            columns[i].image(img, caption=os.path.basename(img_path), width=150)
                        except Exception as e:
                            columns[i].error(f"Ошибка: {e}")
                else:
                    st.warning(f"В папке не найдено изображений с поддерживаемыми расширениями: {', '.join(supported_extensions)}")
        else:
            st.error(f"Папка {images_folder} не существует!")

# Функция для обработки файла
def process_excel_file():
    # Проверяем, есть ли все необходимые настройки
    input_file = config_manager.get_setting("paths.input_file_path", "")
    output_folder = config_manager.get_setting("paths.output_folder_path", "")
    images_folder = config_manager.get_setting("paths.images_folder_path", "")
    
    log.info(f"Начинаем обработку файла: {input_file}")
    log.info(f"Выходная папка: {output_folder}")
    log.info(f"Папка с изображениями: {images_folder}")
    
    if not input_file or not output_folder or not images_folder:
        error_msg = "Не указаны все необходимые пути!"
        log.error(error_msg)
        st.error(error_msg)
        return
    
    if not os.path.isfile(input_file):
        error_msg = f"Файл {input_file} не найден!"
        log.error(error_msg)
        st.error(error_msg)
        return
    
    if not os.path.isdir(images_folder):
        error_msg = f"Папка с изображениями {images_folder} не найдена!"
        log.error(error_msg)
        st.error(error_msg)
        return
    
    # Создаем выходную папку, если она не существует
    try:
        os.makedirs(output_folder, exist_ok=True)
        log.info(f"Выходная папка создана или уже существует: {output_folder}")
    except Exception as e:
        error_msg = f"Ошибка при создании выходной папки: {e}"
        log.error(error_msg)
        st.error(error_msg)
        return
    
    # Получаем настройки
    article_column = config_manager.get_setting("excel_settings.article_column", "C")
    image_column = config_manager.get_setting("excel_settings.image_column", "A")
    start_row = config_manager.get_setting("excel_settings.start_row", 2)
    sheet_index = config_manager.get_setting("excel_settings.sheet_index", 1)
    max_size_kb = config_manager.get_setting("image_settings.max_size_kb", 100)
    quality = config_manager.get_setting("image_settings.quality", 90)
    adjust_dimensions = config_manager.get_setting("excel_settings.adjust_dimensions", True)
    
    log.info(f"Настройки: article_column={article_column}, image_column={image_column}, start_row={start_row}, sheet_index={sheet_index}")
    log.info(f"Настройки изображений: max_size_kb={max_size_kb}, quality={quality}, adjust_dimensions={adjust_dimensions}")
    
    # Создаем копию исходного файла
    temp_dir = ensure_temp_dir()
    try:
        # Создаем копию файла во временную директорию
        log.info(f"Создаем копию файла {input_file} во временную директорию {temp_dir}")
        temp_file = excel_utils.create_excel_copy(input_file, temp_dir)
        log.info(f"Создана копия файла: {temp_file}")
        
        # Отображаем прогресс бар
        progress_bar = st.progress(0, text="Подготовка к обработке...")
        
        # Запускаем обработку
        log.info("Начинаем обработку файла...")
        with st.spinner("Обработка файла..."):
            try:
                stats = excel_utils.process_excel_file(
                    excel_file=temp_file,
                    article_column=article_column,
                    image_column=image_column,
                    images_folder=images_folder,
                    start_row=start_row,
                    sheet_index=sheet_index,
                    max_size_kb=max_size_kb,
                    quality=quality,
                    adjust_dimensions=adjust_dimensions
                )
                
                log.info(f"Обработка завершена. Статистика: {stats}")
                
                # Проверяем, что статистика содержит ожидаемые ключи
                if not stats or not isinstance(stats, dict) or "output_file" not in stats:
                    error_msg = "Неверный формат результата обработки. Отсутствуют необходимые данные."
                    log.error(error_msg)
                    st.error(error_msg)
                    return
                
                # Обновляем прогресс
                progress_bar.progress(100, text="Обработка завершена!")
                
                # Проверяем существование выходного файла
                if not os.path.exists(stats["output_file"]):
                    error_msg = f"Файл результата не найден: {stats['output_file']}"
                    log.error(error_msg)
                    st.error(error_msg)
                    return
                
                # Копируем результат в выходную папку
                output_filename = f"{os.path.splitext(os.path.basename(input_file))[0]}_with_images.xlsx"
                output_file = os.path.join(output_folder, output_filename)
                log.info(f"Копируем результат из {stats['output_file']} в {output_file}")
                
                # Если файл уже существует, удаляем его
                if os.path.exists(output_file):
                    os.remove(output_file)
                    log.info(f"Удален существующий файл: {output_file}")
                
                shutil.copy2(stats["output_file"], output_file)
                log.info(f"Файл скопирован в: {output_file}")
                
                # Удаляем временные файлы
                try:
                    os.remove(stats["output_file"])
                    os.remove(temp_file)
                    log.info("Временные файлы удалены")
                except Exception as e:
                    log.warning(f"Не удалось удалить временные файлы: {e}")
                
                # Обновляем статистику с новым путем
                stats["output_file"] = output_file
                
                # Сохраняем статистику в session_state
                st.session_state["processing_stats"] = stats
                st.session_state["processing_complete"] = True
                
                # Отображаем результаты
                log.info("Отображаем результаты обработки")
                show_results(stats)
                
            except Exception as e:
                error_msg = f"Ошибка при обработке Excel файла: {e}"
                log.exception(error_msg)  # Логируем с полным стеком вызовов
                st.error(error_msg)
    except Exception as e:
        error_msg = f"Ошибка при подготовке файла к обработке: {e}"
        log.exception(error_msg)  # Логируем с полным стеком вызовов
        st.error(error_msg)

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
    if st.button("Открыть папку с результатом"):
        output_folder = os.path.dirname(stats["output_file"])
        # Используем команду в зависимости от ОС
        if os.name == 'nt':  # Windows
            os.startfile(output_folder)
        elif os.name == 'posix':  # macOS и Linux
            os.system(f"open {output_folder}")  # macOS
            # os.system(f"xdg-open {output_folder}")  # Linux

# Главная функция приложения
def main():
    # Заголовок
    st.title("ExcelWithImages")
    st.write("Вставка изображений в Excel по артикулам")
    
    # Инициализируем session_state для хранения состояния
    if 'excel_uploaded' not in st.session_state:
        st.session_state['excel_uploaded'] = False
    
    if 'start_processing' not in st.session_state:
        st.session_state['start_processing'] = False
    
    if 'processing_complete' not in st.session_state:
        st.session_state['processing_complete'] = False
    
    if 'processing_stats' not in st.session_state:
        st.session_state['processing_stats'] = None
    
    if 'current_settings' not in st.session_state:
        st.session_state['current_settings'] = config_manager.get_config_manager().current_settings
    
    # Добавляем боковую панель с настройками
    update_sidebar_buttons()
    show_settings()
    
    # Если обработка завершена, показываем результаты и кнопку для новой обработки
    if st.session_state.get('processing_complete', False):
        show_results(st.session_state.get('processing_stats', {}))
        
        if st.button("Начать новую обработку"):
            st.session_state['processing_complete'] = False
            st.session_state['processing_stats'] = None
            st.session_state['start_processing'] = False
            st.rerun()
    else:
        # Основные секции приложения
        
        # Добавляем большую кнопку "Начать обработку" сверху
        if st.session_state.get('excel_uploaded', False):
            st.markdown("<style>div.stButton > button {background-color: #4CAF50; color: white; font-size: 20px; height: 60px; width: 100%;}</style>", unsafe_allow_html=True)
            if st.button("НАЧАТЬ ОБРАБОТКУ", key="big_process_button"):
                st.session_state['start_processing'] = True
        
        # 1. Загрузка файла Excel
        file_uploader_section()
        
        # 2. Выбор папки с изображениями
        images_folder_section()
        
        # 3. Выбор выходной папки
        output_folder_section()
        
        # Если нажата кнопка "Обработать Excel файл", запускаем обработку
        if st.session_state.get('start_processing', False):
            process_excel_file()
            # Сбрасываем флаг, чтобы не запускать обработку повторно
            st.session_state['start_processing'] = False
    
    # Отображаем логи
    with st.expander("Логи", expanded=False):
        st.text_area("", value=log_stream.getvalue(), height=300)

if __name__ == "__main__":
    main() 