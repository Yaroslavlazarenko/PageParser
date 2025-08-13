    
# excel_reader.py
import pandas as pd
import os
from typing import List, Dict, Optional

def read_students_from_excel(file_path: str) -> Optional[List[Dict]]:
    """
    Читає дані студентів з Excel-файлу та форматує їх для подальшого пошуку.
    Також зчитує вже існуючі результати перевірки, щоб не виконувати роботу повторно.
    """
    if not os.path.exists(file_path):
        print(f"❌ Помилка: Файл не знайдено за шляхом: {file_path}")
        return None

    try:
        print(f"📖 Читаю дані з файлу '{os.path.basename(file_path)}'...")
        df = pd.read_excel(file_path, engine='openpyxl')

        # Перевіряємо наявність обов'язкових стовпців
        required_columns = ['Прізвище', "Ім'я", 'По батькові', 'Конк. бал']
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            print(f"❌ Помилка: У файлі Excel відсутні обов'язкові стовпці: {missing_columns}")
            return None

        students_list = []
        for index, row in df.iterrows():
            # Пропускаємо порожні рядки, орієнтуючись на прізвище
            if pd.isna(row['Прізвище']):
                continue

            last_name = str(row['Прізвище']).strip()
            first_name = str(row["Ім'я"]).strip()
            patronymic = str(row['По батькові']).strip()
            
            # Формуємо ім'я для пошуку у форматі "Прізвище І. П."
            first_name_initial = f"{first_name[0]}." if first_name else ''
            patronymic_initial = f" {patronymic[0]}." if patronymic else ''
            search_name = f"{last_name} {first_name_initial}{patronymic_initial}".strip()
            
            # --- ВИПРАВЛЕННЯ ТУТ ---
            # Безпечна обробка коду спеціальності
            raw_spec_code = row.get('Код спец')
            if pd.notna(raw_spec_code):
                # Просто перетворюємо на рядок і прибираємо '.0' в кінці, якщо це було число
                specialty_code = str(raw_spec_code).replace('.0', '').strip()
            else:
                specialty_code = None
            # --- КІНЕЦЬ ВИПРАВЛЕННЯ ---

            # Обробка вже існуючих результатів
            raw_result = row.get('Результат перевірки')
            current_result = str(raw_result).strip() if pd.notna(raw_result) else None

            student_data = {
                'index': index,
                'search_name': search_name,
                'last_name': last_name,
                'first_name': first_name,
                'patronymic': patronymic,
                'score': row['Конк. бал'],
                'specialty_code': specialty_code,
                'Результат перевірки': current_result
            }
            students_list.append(student_data)
        
        print(f"✅ Успішно оброблено {len(students_list)} записів з файлу.")
        return students_list
        
    except Exception as e:
        # Додаємо виведення типу помилки для кращої діагностики
        print(f"❌ Виникла неочікувана помилка під час читання файлу Excel: {type(e).__name__}: {e}")
        return None

  