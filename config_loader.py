import json
import os

CONFIG_FILE = 'config.json'

def load_config() -> dict:
    """
    Завантажує конфігураційний файл config.json.

    Returns:
        Словник з налаштуваннями.
    
    Raises:
        FileNotFoundError: Якщо файл config.json не знайдено.
        ValueError: Якщо виникає помилка при розборі JSON.
    """
    if not os.path.exists(CONFIG_FILE):
        raise FileNotFoundError(f"Конфігураційний файл не знайдено: '{CONFIG_FILE}'. "
                                f"Створіть його та вкажіть шлях до Excel-файлу.")
    
    try:
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            config = json.load(f)
        return config
    except json.JSONDecodeError as e:
        raise ValueError(f"Помилка синтаксису у файлі '{CONFIG_FILE}': {e}")