# api_parser.py
import asyncio
import re
from bs4 import BeautifulSoup
import aiohttp
from typing import List, Dict, Optional
import json
from datetime import datetime
import aiofiles

# --- Классы исключений и вспомогательные функции (без изменений) ---
class APIError(Exception):
    """Базовый класс для всех ошибок API abit-poisk."""
    pass
# ... (остальные классы ошибок без изменений) ...
class APIRateLimitError(APIError):
    def __init__(self, message="Превышен лимит запросов или ошибка 'max_user_connections'"):
        self.message = message
        super().__init__(self.message)
class APIUnavailableError(APIError):
    def __init__(self, status_code: int, reason: str):
        self.message = f"API временно недоступен. Статус: {status_code} {reason}"
        super().__init__(self.message)
class APIInvalidResponseError(APIError):
    def __init__(self, message="Некорректный или неконсистентный ответ от API"):
        self.message = message
        super().__init__(self.message)
def _parse_delay_from_message(message: str) -> Optional[int]:
    match = re.search(r'через (\d+)\s+секунд', message)
    if match:
        try: return int(match.group(1))
        except (ValueError, IndexError): return None
    return None

LOG_FILE = 'debug_log.txt'
log_lock = asyncio.Lock()

async def log_to_file(content: str):
    async with log_lock:
        try:
            async with aiofiles.open(LOG_FILE, mode='a', encoding='utf-8') as f:
                await f.write(content)
        except Exception as e:
            print(f"!!! Критическая ошибка записи в лог-файл: {e}")

# --- ИЗМЕНЕНИЕ 1: НАХУЙ ГЛОБАЛЬНУЮ СЕССИЮ! УНИЧТОЖАЕМ УЛИКИ! ---
# AIOHTTP_SESSION = None  <--- УБРАНО НАХУЙ
# async def get_session(): ... <--- УБРАНО НАХУЙ
# async def close_session(): ... <--- УБРАНО НАХУЙ

async def fetch_applications_html(search_query: str, semaphore: asyncio.Semaphore) -> Optional[str]:
    api_url = "http://abit-poisk.org.ua/api/statements/"
    headers = {
        'Host': 'abit-poisk.org.ua',
        'Sec-Ch-Ua-Platform': '"macOS"',
        'X-Requested-With': 'XMLHttpRequest',
        'Accept-Language': 'ru-RU,ru;q=0.9',
        'Sec-Ch-Ua': '"Chromium";v="139", "Not;A=Brand";v="99"',
        'Content-Type': 'application/x-www-form-urlencoded',
        'Sec-Ch-Ua-Mobile': '?0',
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/139.0.0.0 Safari/537.36',
        'Accept': '*/*',
        'Origin': 'http://abit-poisk.org.ua',
        'Sec-Fetch-Site': 'same-origin',
        'Sec-Fetch-Mode': 'cors',
        'Sec-Fetch-Dest': 'empty',
        'Referer': 'http://abit-poisk.org.ua/',
        'Accept-Encoding': 'gzip, deflate, br',
        'Priority': 'u=1, i',
        'Connection': 'keep-alive'
    }
    payload = {'search': search_query, 'offset': 0}
    all_html_parts = []
    
    # session = await get_session() # <--- ЭТОЙ ХУЙНИ БОЛЬШЕ НЕТ

    RETRY_DELAY = 20
    MAX_RETRIES = 5
    previous_count = -1

    while True:
        data = None
        is_successful_request = False
        last_exception = None
        
        attempt = 0
        while attempt < MAX_RETRIES:
            last_exception = None
            log_timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')
            log_parts = [f"\n{'='*120}\n", f"TRANSACTION AT: {log_timestamp} | Search: '{search_query}' | Offset: {payload.get('offset')} | Attempt: {attempt + 1}/{MAX_RETRIES}\n"]
            
            try:
                text_response, status_code, reason, response_headers = "", 0, "", {}
                async with semaphore:
                    log_parts.extend([f"{'-'*55} REQUEST {'-'*56}\n", f"URL: {api_url}\n", f"Headers: {json.dumps(headers, indent=2, ensure_ascii=False)}\n", f"Payload Body: {json.dumps(payload, indent=2, ensure_ascii=False)}\n"])
                    params = {'nocache': int(asyncio.get_event_loop().time() * 1000)}

                    # --- ИЗМЕНЕНИЕ 2: СОЗДАЁМ, БЛЯДЬ, НОВУЮ СЕССИЮ ДЛЯ КАЖДОГО ЗАПРОСА! ---
                    # Каждый запрос - это новый агент-смертник! 💥
                    connector = aiohttp.TCPConnector(ssl=False)
                    async with aiohttp.ClientSession(connector=connector) as session:
                        async with session.post(api_url, headers=headers, params=params, data=payload, timeout=30) as response:
                            text_response = await response.text()
                            status_code = response.status
                            reason = str(response.reason)
                            response_headers = dict(response.headers)
                            log_parts.extend([f"{'-'*54} RESPONSE {'-'*55}\n", f"Status: {status_code} {reason}\n", f"Headers: {json.dumps(response_headers, indent=2, ensure_ascii=False)}\n", f"Body:\n{text_response}\n"])
                            response.raise_for_status()

                # ... (вся остальная логика обработки ответа без изменений) ...
                if 'max_user_connections' in text_response: raise APIRateLimitError("Обнаружена ошибка 'max_user_connections' в теле ответа.")
                try: data = json.loads(text_response)
                except json.JSONDecodeError: raise APIInvalidResponseError(f"Ожидался JSON, но получен невалидный ответ. Начало: {text_response[:150]}")
                if not data.get('success', False):
                    error_message = data.get('message', '') or data.get('error', '')
                    if 'Частота запитів' in error_message: raise APIRateLimitError(f"API сообщил об ограничении: {error_message}")
                    else: raise APIInvalidResponseError(f"API вернул success=false: {error_message}")
                if data.get('count', 0) > 0 and not data.get('html'): raise APIInvalidResponseError("API сообщил о наличии результатов, но вернул пустой HTML.")
                
                is_successful_request = True
                await log_to_file("".join(log_parts) + f"{'='*120}\n")
                break
            
            except (aiohttp.ClientError, asyncio.TimeoutError, APIError) as e:
                last_exception = e

            # ... (вся остальная логика обработки ошибок без изменений) ...
            if isinstance(last_exception, APIRateLimitError):
                log_parts.extend([f"{'-'*54} EXCEPTION CAUGHT {'-'*48}\n", f"Error Details: {last_exception}\n"])
                current_delay = RETRY_DELAY
                parsed_delay = _parse_delay_from_message(last_exception.message)
                if parsed_delay is not None: current_delay = parsed_delay + 1
                log_parts.append(f"INFO: Ошибка лимита запросов. Ожидание {current_delay} сек. Попытка не засчитана.\n")
                await log_to_file("".join(log_parts) + f"{'='*120}\n")
                await asyncio.sleep(current_delay)
            else:
                attempt += 1
                log_parts.extend([f"{'-'*54} EXCEPTION CAUGHT {'-'*48}\n", f"Error Details: {last_exception}\n"])
                if attempt < MAX_RETRIES:
                    log_parts.append(f"INFO: Произошла ошибка. Повтор через {RETRY_DELAY} сек.\n")
                    await log_to_file("".join(log_parts) + f"{'='*120}\n")
                    await asyncio.sleep(RETRY_DELAY)
                else:
                    await log_to_file("".join(log_parts) + f"{'='*120}\n")
        
        if not is_successful_request:
            raise last_exception

        if not (data and data.get('success') and data.get('html')):
            break
        
        all_html_parts.append(data['html'])
        total_count = data.get('count', 0)
        
        current_count = len(parse_applications("".join(all_html_parts)))
        if current_count >= total_count: break
        
        if current_count == previous_count and current_count > 0:
            await log_to_file(f"WARNING: Пагинация застряла для '{search_query}' на смещении {current_count}. Прерываю цикл.\n")
            break

        previous_count = current_count
        payload['offset'] = current_count
        await asyncio.sleep(0.5)
            
    return "".join(all_html_parts) if all_html_parts else None


def parse_applications(html_content: str) -> List[Dict]:
    # ... (код парсинга без изменений, он охуенен как есть) ...
    soup = BeautifulSoup(html_content, 'lxml'); applications = []; base_url = "https://abit-poisk.org.ua"
    table_bodies = soup.find_all('tbody')
    if not table_bodies: return []
    for table_body in table_bodies:
        for row in table_body.find_all('tr'):
            cells = row.find_all('td');
            if len(cells) < 14: continue
            try:
                rank_position_text = cells[3].get_text(strip=True); rank_position = int(rank_position_text) if rank_position_text.isdigit() else 0
                total_score_text = cells[6].get_text(strip=True).replace(',', '.'); total_score = float(total_score_text) if total_score_text else 0.0
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
                rank_url_tag = cells[3].find('a'); rank_url = base_url + rank_url_tag.get('href', '') if rank_url_tag else ''
                university_url_tag = cells[9].find('a'); university_url = base_url + university_url_tag.get('href', '') if university_url_tag else ''
                application_data = {'degree_level_short': cells[0].find('div').get_text(strip=True), 'degree_level_full': cells[0].get('title', '').strip(),'applicant_name': cells[1].get_text(strip=True), 'status': cells[2].get_text(strip=True),'rank_position': rank_position, 'rank_url': rank_url,'priority': ' '.join(cells[4].get_text(strip=True).split()), 'places': places_info, 'total_score': total_score,'avg_document_score': cells[7].get_text(strip=True), 'score_components': scores, 'coefficients': coefficients,'university_name': cells[9].get_text(strip=True), 'university_url': university_url,'faculty_short': cells[10].get_text(strip=True), 'faculty_full': cells[10].get('title', '').strip(),'specialty_code': specialty_code, 'specialty_name': specialty_name, 'specialization': specialization,'quota': cells[12].get_text(strip=True), 'originals_submitted': cells[13].get_text(strip=True) == '+'}
                applications.append(application_data)
            except Exception: continue
    return applications