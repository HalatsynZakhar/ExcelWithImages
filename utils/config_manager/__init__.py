from .config_manager import ConfigManager

# Хранит глобальный экземпляр ConfigManager
_config_manager_instance = None

def init_config_manager(presets_folder: str) -> ConfigManager:
    """
    Инициализирует менеджер конфигурации и возвращает его
    
    Args:
        presets_folder: Путь к папке с пресетами настроек
        
    Returns:
        Экземпляр ConfigManager
    """
    global _config_manager_instance
    _config_manager_instance = ConfigManager(presets_folder)
    return _config_manager_instance

def get_config_manager() -> ConfigManager:
    """
    Возвращает экземпляр ConfigManager
    
    Returns:
        Экземпляр ConfigManager
    """
    global _config_manager_instance
    if _config_manager_instance is None:
        raise RuntimeError("ConfigManager не инициализирован. Сначала вызовите init_config_manager()")
    return _config_manager_instance

def get_setting(path: str, default=None):
    """
    Получает значение настройки по указанному пути
    
    Args:
        path: Путь к настройке в формате dot notation (например, "paths.input_folder")
        default: Значение по умолчанию, если настройка не найдена
        
    Returns:
        Значение настройки или default, если настройка не найдена
    """
    return get_config_manager().get_setting(path, default)

def set_setting(path: str, value):
    """
    Устанавливает значение настройки по указанному пути
    
    Args:
        path: Путь к настройке в формате dot notation (например, "paths.input_folder")
        value: Новое значение настройки
    """
    get_config_manager().set_setting(path, value)

def save_settings(preset_name: str = None) -> bool:
    """
    Сохраняет текущие настройки в файл
    
    Args:
        preset_name: Имя пресета для сохранения
        
    Returns:
        True, если настройки успешно сохранены, иначе False
    """
    return get_config_manager().save_settings(preset_name)

def load_settings(preset_name: str) -> bool:
    """
    Загружает настройки из файла
    
    Args:
        preset_name: Имя пресета для загрузки
        
    Returns:
        True, если настройки успешно загружены, иначе False
    """
    return get_config_manager().load_settings(preset_name)

def reset_settings():
    """
    Сбрасывает настройки к значениям по умолчанию
    """
    get_config_manager().reset_settings()

def get_presets_list() -> list:
    """
    Возвращает список доступных пресетов настроек
    
    Returns:
        Список имен пресетов
    """
    return get_config_manager().get_presets_list()

def delete_preset(preset_name: str) -> bool:
    """
    Удаляет пресет с указанным именем
    
    Args:
        preset_name: Имя пресета для удаления
        
    Returns:
        True, если пресет успешно удален, иначе False
    """
    return get_config_manager().delete_preset(preset_name) 