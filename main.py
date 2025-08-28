# main.py (ВЕРСИЯ, СОВМЕСТИМАЯ С ТВОИМ EXCEL_WRITER)
import os
import shutil
import math
import asyncio
from typing import Dict, List
import sys

# Предполагается, что эти модули существуют в проекте
from config_loader import load_config
from excel_reader import read_students_from_excel
from excel_writer import save_results_to_excel

# Импортируем все необходимое из нашего модуля api_parser
from api_parser import (
    fetch_applications_html, 
    parse_applications, 
    APIError
)

# --- 1. Логика обработки одного студента (без изменений) ---
async def process_student_async(student: Dict, semaphore: asyncio.Semaphore) -> None:
    search_query = student['search_name']
    score_to_find = student['score']
    specialty_to_find = student['specialty_code']

    try:
        html_content = await fetch_applications_html(search_query, semaphore)
        if not html_content:
            student['final_result'] = "Не знайдено жодної заяви"
            return

        all_apps_by_name = parse_applications(html_content)
        if not all_apps_by_name:
            student['final_result'] = "Не знайдено. Треба дзвонити"
            return
        
        reference_app = None
        for app in all_apps_by_name:
            try:
                student_score = float(score_to_find)
                app_score = float(app.get('total_score', 0))
                score_matches = math.isclose(app_score, student_score, rel_tol=1e-5)
            except (ValueError, TypeError):
                score_matches = False

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
    
    except APIError as e:
        student['final_result'] = f"Помилка API: {e}"
    except asyncio.TimeoutError:
        student['final_result'] = "Помилка: Час очікування запиту вичерпано"
    except Exception as e:
        student['final_result'] = f"Критична помилка: {type(e).__name__}: {e}"


# --- 2. Основная асинхронная логика ---
async def main_async_logic(config: Dict) -> List[Dict]:
    original_excel_path = config['excel_file_path']
    output_excel_filename = "students_with_results.xlsx"
    
    if not os.path.exists(original_excel_path):
        print(f"❌ Помилка: Вхідний файл Excel не знайдено за шляхом: {original_excel_path}")
        return []
    
    if not os.path.exists(output_excel_filename):
        print(f"📄 Створюю копію файлу для результатів: '{output_excel_filename}'")
        shutil.copy(original_excel_path, output_excel_filename)
    else:
        print(f"📂 Використовую існуючий файл з результатами: '{output_excel_filename}'")
    
    all_students = read_students_from_excel(output_excel_filename)
    if not all_students:
        print("Не вдалося прочитати дані з Excel або файл порожній.")
        return []
    
    output_col_name = config.get("output_column_name", "Результат перевірки")
    students_to_process = [s for s in all_students if not s.get(output_col_name)]
    
    if not students_to_process:
        print("\n✅ Всі студенти вже оброблені.")
        return all_students

    CONCURRENT_LIMIT = config.get("concurrent_requests", 5)
    semaphore = asyncio.Semaphore(CONCURRENT_LIMIT)

    print(f"\nВсього студентів у файлі: {len(all_students)}")
    print(f"Потребують перевірки: {len(students_to_process)}. Запускаю до {CONCURRENT_LIMIT} одночасних запитів...")
    print("Натисніть Ctrl+C для безпечної зупинки та збереження прогресу.")

    tasks = [asyncio.create_task(process_student_async(student, semaphore)) for student in students_to_process]
    
    processed_count = 0
    total_to_process = len(students_to_process)
    for task in asyncio.as_completed(tasks):
        await task
        processed_count += 1
        print(f"[{processed_count}/{total_to_process}] Оброблено...", end='\r')

    print(f"\n🎉 Всі {total_to_process} студентів успішно оброблені.")
    return all_students

# --- 3. Точка входа в программу ---
def main():
    """Точка входа, которая запускает асинхронный цикл и обрабатывает исключения."""
    all_students: List[Dict] = []
    output_filename = "students_with_results.xlsx"
    config = {}

    try:
        config = load_config()
        # Проверим только самый важный ключ
        if 'excel_file_path' not in config:
            print("❌ СТОП! У файлі config.yaml відсутній параметр 'excel_file_path'!")
            sys.exit(1)
            
        all_students = asyncio.run(main_async_logic(config))
            
    except (KeyboardInterrupt, asyncio.CancelledError):
        print("\n\n🛑 Переривання роботи за командою користувача (Ctrl+C).")
    except Exception as e:
        print(f"\n❌ Сталася неочікувана помилка на верхньому рівні: {type(e).__name__}: {e}")
    finally:
        print("\nЗавершую роботу...")
        
        # --- ИЗМЕНЕНИЕ! ВОЗВРАЩАЕМСЯ К ПРОСТОЙ И НАДЕЖНОЙ СИСТЕМЕ! ---
        if all_students and config:
            # Просто, блядь, фильтруем тех, кого обработали
            students_with_new_results = [s for s in all_students if 'final_result' in s]

            if students_with_new_results:
                print("Зберігаю поточний прогрес...")
                # Имя колонки для вывода
                output_col = config.get("output_column_name", "Результат перевірки")
                
                # Вызываем твой охуенный writer, который сам разберется по 'index'
                save_results_to_excel(
                    file_path=output_filename, 
                    students_data=students_with_new_results, 
                    col_name=output_col
                )
            else:
                 print("Немає нових результатів для збереження.")
        else:
            print("Немає даних для збереження (можливо, сталася помилка на старті).")
        
        print("Роботу безпечно зупинено.")

if __name__ == "__main__":
    main()