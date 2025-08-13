# api_parser.py
import asyncio
import certifi
import re
from bs4 import BeautifulSoup
import aiohttp
from typing import List, Dict, Optional
import json

# Створюємо одну сесію для всіх запитів, це більш ефективно
AIOHTTP_SESSION = None

async def get_session():
    """Створює та повертає глобальну сесію aiohttp."""
    global AIOHTTP_SESSION
    if AIOHTTP_SESSION is None or AIOHTTP_SESSION.closed:
        # Для уникнення помилки SSL на деяких системах
        ssl_context = certifi.where()
        connector = aiohttp.TCPConnector(ssl=False) # Часто вирішує проблеми з SSL
        AIOHTTP_SESSION = aiohttp.ClientSession(connector=connector)
    return AIOHTTP_SESSION

async def close_session():
    """Закриває глобальну сесію aiohttp."""
    global AIOHTTP_SESSION
    if AIOHTTP_SESSION and not AIOHTTP_SESSION.closed:
        await AIOHTTP_SESSION.close()
        AIOHTTP_SESSION = None

async def fetch_applications_html(search_query: str) -> Optional[str]:
    """
    Асинхронно отримує HTML з усіма заявами для даного ПІБ,
    з механізмом повторних спроб у разі невдачі.
    """
    api_url = "http://abit-poisk.org.ua/api/statements/"
    headers = {'User-Agent':'Mozilla/5.0','Accept':'application/json, text/javascript, */*; q=0.01','Content-Type':'application/x-www-form-urlencoded','X-Requested-With':'XMLHttpRequest','Origin':'http://abit-poisk.org.ua','Referer':'http://abit-poisk.org.ua/',}
    payload = {'search': search_query, 'offset': 0}
    all_html_parts = []
    
    session = await get_session()

    # --- ЛОГІКА ПОВТОРНИХ СПРОБ ---
    MAX_RETRIES = 3
    RETRY_DELAY = 5 # секунд

    while True: # Цей цикл для пагінації (завантаження всіх сторінок)
        
        # Внутрішній цикл для повторних спроб конкретного запиту
        for attempt in range(MAX_RETRIES):
            try:
                params = {'nocache': int(asyncio.get_event_loop().time() * 1000)}
                async with session.post(api_url, headers=headers, params=params, data=payload) as response:
                    
                    # Якщо це помилка сервера (5xx), варто спробувати ще раз
                    if response.status >= 500:
                        # Спровокуємо повторну спробу, піднявши виключення
                        response.raise_for_status() 

                    # Якщо помилка клієнта (4xx) - повторювати немає сенсу
                    if response.status >= 400:
                        return "".join(all_html_parts) if all_html_parts else None

                    text_response = await response.text()
                    if not text_response: # Якщо відповідь порожня, пробуємо ще раз
                        raise ValueError("Empty response from server")

                    # Парсимо JSON. Якщо тут помилка, блок except її зловить
                    data = json.loads(text_response)
                    
                    # Якщо все добре, виходимо з циклу повторних спроб
                    break

            except (aiohttp.ClientError, json.JSONDecodeError, ValueError) as e:
                if attempt < MAX_RETRIES - 1: # Якщо це не остання спроба
                    # print(f"Помилка '{type(e).__name__}' для '{search_query}'. Спроба {attempt + 2}/{MAX_RETRIES} через {RETRY_DELAY} сек...")
                    await asyncio.sleep(RETRY_DELAY)
                    continue # Переходимо до наступної спроби
                else:
                    # Всі спроби вичерпано
                    # print(f"Не вдалося отримати дані для '{search_query}' після {MAX_RETRIES} спроб.")
                    return "".join(all_html_parts) if all_html_parts else None
        # --- КІНЕЦЬ ЛОГІКИ ПОВТОРНИХ СПРОБ ---

        # Обробка успішної відповіді
        if not (data.get('success') and 'html' in data):
            return "".join(all_html_parts) if all_html_parts else None
        
        all_html_parts.append(data['html'])
        total_count = data.get('count', 0)
        
        current_count = len(parse_applications("".join(all_html_parts)))
        if current_count >= total_count:
            break # Успішно завантажили всі сторінки, виходимо з циклу пагінації
        
        payload['offset'] = current_count
        await asyncio.sleep(0.3)
            
    return "".join(all_html_parts)


def parse_applications(html_content: str) -> List[Dict]:
    """Синхронна функція парсингу (вона не робить I/O, тому не потребує async)."""
    # Цей код залишається без змін
    soup = BeautifulSoup(html_content, 'lxml'); applications = []; base_url = "https://abit-poisk.org.ua"
    table_bodies = soup.find_all('tbody')
    if not table_bodies: return []
    for table_body in table_bodies:
        for row in table_body.find_all('tr'):
            cells = row.find_all('td');
            if len(cells) < 14: continue
            scores = {}; score_components_cell = cells[8]; subjects = score_components_cell.find_all('dt'); points = score_components_cell.find_all('dd')
            for subject, point in zip(subjects, points):
                subject_name = subject.get_text(strip=True); score_value = point.get_text(strip=True)
                scores[subject_name] = int(score_value) if score_value.isdigit() else score_value
            coefficients = {}; coeffs_list = score_components_cell.find('ul', class_='list-unstyled')
            if coeffs_list:
                for li in coeffs_list.find_all('li'):
                    coeff_text = li.get_text(strip=True)
                    if ':' in coeff_text: key, value = coeff_text.split(':', 1); coefficients[key.strip()] = value.strip()
            places_info = {}; places_cell = cells[5]
            for span in places_cell.find_all('span'):
                text = span.get_text(strip=True); tooltip = span.get('data-stooltip', '').lower()
                numbers = re.findall(r'\d+', text)
                if not numbers: continue
                number = int(numbers[0])
                if 'вм' in text.lower() or 'загальна кількість' in tooltip: places_info['total'] = number
                elif 'бм' in text.lower() or 'бюджетних' in tooltip: places_info['budget_max'] = number
                elif 'к' in text.lower() or 'контракт' in tooltip: places_info['contract'] = number
            specialty_divs = cells[11].find_all('div'); specialty_full = specialty_divs[0].get_text(strip=True) if specialty_divs else ''
            specialty_code = ''; specialty_name = ''
            if specialty_full:
                match = re.match(r'^([A-Z0-9]+)(.*)', specialty_full)
                if match: specialty_code = match.group(1); specialty_name = match.group(2).strip()
                else: specialty_name = specialty_full
            specialization = specialty_divs[1].get_text(strip=True) if len(specialty_divs) > 1 else ''
            application_data = {'degree_level_short': cells[0].find('div').get_text(strip=True), 'degree_level_full': cells[0].get('title', '').strip(),'applicant_name': cells[1].get_text(strip=True), 'status': cells[2].get_text(strip=True),'rank_position': int(cells[3].get_text(strip=True)), 'rank_url': base_url + cells[3].find('a').get('href', ''),'priority': ' '.join(cells[4].get_text(strip=True).split()), 'places': places_info, 'total_score': float(cells[6].get_text(strip=True)),'avg_document_score': cells[7].get_text(strip=True), 'score_components': scores, 'coefficients': coefficients,'university_name': cells[9].get_text(strip=True), 'university_url': base_url + cells[9].find('a').get('href', ''),'faculty_short': cells[10].get_text(strip=True), 'faculty_full': cells[10].get('title', '').strip(),'specialty_code': specialty_code, 'specialty_name': specialty_name, 'specialization': specialization,'quota': cells[12].get_text(strip=True), 'originals_submitted': cells[13].get_text(strip=True) == '+'}
            applications.append(application_data)
    return applications