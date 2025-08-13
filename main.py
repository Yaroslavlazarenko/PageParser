# main.py
import os
import shutil
import math
import asyncio
from typing import Dict, List

from config_loader import load_config
from excel_reader import read_students_from_excel
from api_parser import fetch_applications_html, parse_applications, close_session
from excel_writer import save_results_to_excel

MAX_CONCURRENT_REQUESTS = 50

async def process_student_async(student: Dict) -> None:
    """
    Асинхронна функція, яка обробляє одного студента.
    Результат записується напряму в словник студента.
    """
    search_query = student['search_name']
    score_to_find = student['score']
    specialty_to_find = student['specialty_code']

    try:
        html_content = await fetch_applications_html(search_query)
        if not html_content:
            student['final_result'] = "Не знайдено жодної заяви"
            return

        all_apps_by_name = parse_applications(html_content)
        if not all_apps_by_name:
            student['final_result'] = "Не знайдено. Треба дзвонити"
            return
        
        reference_app = None
        for app in all_apps_by_name:
            # Перевірка типів перед порівнянням для більшої надійності
            try:
                # Переконуємося, що score_to_find - це число
                student_score = float(score_to_find)
                app_score = float(app.get('total_score', 0))
                score_matches = math.isclose(app_score, student_score, rel_tol=1e-5)
            except (ValueError, TypeError):
                score_matches = False # Якщо не можемо перетворити на число, вважаємо, що не збігається

            specialty_matches = (not specialty_to_find) or (app.get('specialty_code') == specialty_to_find)
            
            if score_matches and specialty_matches:
                reference_app = app
                break
        
        if not reference_app:
            student['final_result'] = "Знайдено, але не ідентифіковано"
            return
        
        score_fingerprint = reference_app.get('score_components')
        if not score_fingerprint:
            student['final_result'] = "Не вдалось отримати бали НМТ"
            return
            
        identified_student_apps = [
            app for app in all_apps_by_name if app.get('score_components') == score_fingerprint
        ]
        
        has_originals = any(app['originals_submitted'] for app in identified_student_apps)
        student['final_result'] = "Вже визначився" if has_originals else "Потрібно дзвонити"
    
    except Exception as e:
        # --- ЗМІНА ТУТ ---
        # Записуємо не тільки тип помилки, а й її текст
        student['final_result'] = f"Помилка обробки: {type(e).__name__}: {e}"


async def main_async_logic() -> List[Dict]:
    """Основна асинхронна логіка, повертає список студентів з результатами."""
    config = load_config()
    original_excel_path = config['excel_file_path']
    output_excel_filename = "students_with_results.xlsx"
    
    if not os.path.exists(output_excel_filename):
        print(f"📄 Створюю копію файлу для результатів: '{output_excel_filename}'")
        shutil.copy(original_excel_path, output_excel_filename)
    else:
        print(f"📂 Використовую існуючий файл з результатами: '{output_excel_filename}'")
    
    all_students = read_students_from_excel(output_excel_filename)
    if not all_students:
        print("Не вдалося прочитати дані з Excel або файл порожній.")
        return []

    students_to_process = [s for s in all_students if not s.get('Результат перевірки')]
    
    if not students_to_process:
        print("\n✅ Всі студенти вже оброблені.")
        return []

    print(f"\nВсього студентів у файлі: {len(all_students)}")
    print(f"Потребують перевірки: {len(students_to_process)}. Запускаю до {MAX_CONCURRENT_REQUESTS} асинхронних завдань...")
    print("Натисніть Ctrl+C для безпечної зупинки та збереження прогресу.")

    tasks = [asyncio.create_task(process_student_async(student)) for student in students_to_process]
    
    for i, task in enumerate(asyncio.as_completed(tasks), 1):
        await task
        print(f"[{i}/{len(students_to_process)}] Оброблено...", end='\r')

    print(f"\n🎉 Всі {len(students_to_process)} студентів успішно оброблені.")
    return all_students

def main():
    """Точка входу, яка запускає асинхронний цикл і обробляє виключення."""
    all_students: List[Dict] = []
    output_filename = "students_with_results.xlsx"
    config = {} # Оголошуємо тут, щоб було доступно в finally

    try:
        config = load_config() # Зчитуємо конфігурацію на початку
        all_students = asyncio.run(main_async_logic())
            
    except KeyboardInterrupt:
        print("\n\n🛑 Переривання роботи за командою користувача (Ctrl+C).")
    except SystemExit as e:
         print(f"\n❌ Критична помилка: {e}")
    except Exception as e:
        print(f"\n❌ Сталася неочікувана помилка: {type(e).__name__}: {e}")
    finally:
        print("\nЗавершую роботу...")
        
        asyncio.run(close_session())

        if all_students:
            students_with_new_results = [s for s in all_students if 'final_result' in s]
            if students_with_new_results:
                print("Зберігаю поточний прогрес...")
                # --- ЗМІНА ТУТ: Беремо назву стовпця з конфігу ---
                output_col = config.get("output_column_name", "Результат перевірки")
                save_results_to_excel(output_filename, students_with_new_results, output_col)
            else:
                 print("Немає нових результатів для збереження.")
        else:
            print("Немає даних для збереження.")
        
        print("Роботу безпечно зупинено.")

if __name__ == "__main__":
    main()