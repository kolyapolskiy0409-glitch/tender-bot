import os
import sys
import json
import re
import tempfile
import shutil
import requests
from dotenv import load_dotenv
from flask import Flask, request, jsonify

# Загружаем переменные из .env для локального запуска
load_dotenv()

api = Flask(__name__)

# --- 1. НАСТРОЙКИ (Читаем из переменных окружения Render) ---
BITRIX24_WEBHOOK = os.getenv("BITRIX24_WEBHOOK")
DEEPSEEK_API_KEY = os.getenv("DEEPSEEK_TOKEN")
DEAL_CATEGORY_ID = int(os.getenv("DEAL_CATEGORY_ID", 42))
DEAL_STAGE_ID = os.getenv("DEAL_STAGE_ID", "8704")
FIELD_LINK_CODE = os.getenv("FIELD_LINK_CODE", "UF_CRM_1774428455758")
FIELD_COMPANY_DIRECTION = os.getenv("FIELD_COMPANY_DIRECTION", "UF_CRM_1774954195201")
# ID корневой папки на Google Drive (ОБЯЗАТЕЛЬНО УКАЗАТЬ)
GOOGLE_DRIVE_ROOT_FOLDER_ID = os.getenv("GOOGLE_DRIVE_ROOT_FOLDER_ID")

# --- 2. КОНСТАНТЫ ДЛЯ DEEPSEEK API ---
DEEPSEEK_API_URL = "https://api.deepseek.com/v1/chat/completions"

# --- 3. ПРОМПТ ДЛЯ DEEPSEEK (скопируйте ваш полный текст) ---
PROMPT = """
(скопируйте сюда ваш полный промпт)
"""

# --- 4. ФУНКЦИИ ДЛЯ РАБОТЫ С БИТРИКС24 ---
def call_bitrix24(method, params):
    """Универсальный вызов REST API Битрикс24."""
    url = f"{BITRIX24_WEBHOOK}{method}"
    try:
        resp = requests.post(url, json=params, timeout=60)
        resp.raise_for_status()
        return resp.json()
    except Exception as e:
        return {"error": str(e)}

def get_enum_value_id(field_name, target_value):
    """Получает ID значения 'target_value' для пользовательского поля-списка."""
    fields_response = call_bitrix24("crm.company.fields", {})
    if "error" in fields_response:
        raise Exception(f"Не удалось получить поля компаний: {fields_response['error']}")
    field_info = fields_response["result"].get(field_name)
    if not field_info:
        raise Exception(f"Поле {field_name} не найдено")
    for item in field_info.get("items", []):
        if item.get("VALUE") == target_value:
            return item.get("ID")
    raise Exception(f"Значение '{target_value}' не найдено для поля {field_name}")

def find_company_by_inn(inn):
    """Поиск компании по ИНН через реквизиты."""
    result = call_bitrix24("crm.requisite.list", {
        "filter": {"RQ_INN": inn},
        "select": ["ID", "ENTITY_TYPE_ID", "ENTITY_ID"]
    })
    if result.get("error"):
        print(f"Ошибка поиска реквизитов: {result['error']}")
        return None
    if result.get("result"):
        for req in result["result"]:
            if req.get("ENTITY_TYPE_ID") == 4:
                return req.get("ENTITY_ID")
    return None

def create_company(company_name, inn):
    """Создание компании + реквизиты + поле 'Направление компании'."""
    direction_id = get_enum_value_id(FIELD_COMPANY_DIRECTION, "НВ")
    company_res = call_bitrix24("crm.company.add", {
        "fields": {
            "TITLE": company_name,
            FIELD_COMPANY_DIRECTION: direction_id
        }
    })
    company_id = company_res.get("result")
    if not company_id:
        raise Exception(f"Ошибка создания компании: {company_res}")

    # Получаем шаблон реквизитов
    presets = call_bitrix24("crm.requisite.preset.list", {})
    preset_id = 1
    if presets.get("result"):
        for p in presets["result"]:
            if "юр" in p.get("NAME", "").lower() or "organiz" in p.get("NAME", "").lower():
                preset_id = p["ID"]
                break

    # Добавляем реквизиты с ИНН
    call_bitrix24("crm.requisite.add", {
        "fields": {
            "ENTITY_TYPE_ID": 4,
            "ENTITY_ID": company_id,
            "PRESET_ID": preset_id,
            "NAME": company_name,
            "RQ_INN": inn
        }
    })
    return company_id

def create_contact(name, phone, email, company_id):
    """Создание контакта и привязка к компании."""
    if not any([name, phone, email]):
        return None
    contact_res = call_bitrix24("crm.contact.add", {
        "fields": {
            "NAME": name or "",
            "PHONE": [{"VALUE": phone, "VALUE_TYPE": "WORK"}] if phone else [],
            "EMAIL": [{"VALUE": email, "VALUE_TYPE": "WORK"}] if email else []
        }
    })
    contact_id = contact_res.get("result")
    if contact_id and company_id:
        call_bitrix24("crm.company.contact.add", {
            "id": company_id,
            "fields": {"CONTACT_ID": contact_id}
        })
    return contact_id

def create_deal(company_id, deal_name, purchase_link):
    """Создание сделки в нужной воронке и статусе."""
    deal_res = call_bitrix24("crm.deal.add", {
        "fields": {
            "TITLE": deal_name,
            "COMPANY_ID": company_id,
            "CATEGORY_ID": DEAL_CATEGORY_ID,
            "STAGE_ID": DEAL_STAGE_ID,
            FIELD_LINK_CODE: purchase_link
        }
    })
    return deal_res.get("result")

# --- 5. ФУНКЦИИ ДЛЯ ФОРМАТИРОВАНИЯ ТЕКСТА (MARKDOWN -> BBCODE) ---
def markdown_table_to_bbcode(table_lines):
    """Преобразует Markdown-таблицу в BBCode [table]."""
    if len(table_lines) < 2:
        return '\n'.join(table_lines)
    # Парсим заголовки
    header_line = table_lines[0].strip('|').split('|')
    headers = [h.strip() for h in header_line]
    data_rows = []
    for line in table_lines[2:]:
        cells = line.strip('|').split('|')
        row = [c.strip() for c in cells]
        data_rows.append(row)
    # Формируем BBCode
    bbcode = '[table]\n'
    bbcode += '[tr]' + ''.join(f'[th]{h}[/th]' for h in headers) + '[/tr]\n'
    for row in data_rows:
        bbcode += '[tr]' + ''.join(f'[td]{cell}[/td]' for cell in row) + '[/tr]\n'
    bbcode += '[/table]'
    return bbcode

def markdown_to_bbcode(text):
    """Преобразует Markdown в BBCode для Битрикс24 с поддержкой таблиц."""
    lines = text.split('\n')
    result = []
    i = 0
    while i < len(lines):
        line = lines[i]
        if line.strip().startswith('|') and i+1 < len(lines) and re.match(r'^[\s]*\|[\s\-:]+[\s]*\|', lines[i+1]):
            table_lines = []
            while i < len(lines) and lines[i].strip().startswith('|'):
                table_lines.append(lines[i].strip())
                i += 1
            result.append(markdown_table_to_bbcode(table_lines))
            continue
        else:
            # Жирный, курсив, заголовки, списки
            line = re.sub(r'\*\*(.*?)\*\*', r'[b]\1[/b]', line)
            line = re.sub(r'__(.*?)__', r'[b]\1[/b]', line)
            line = re.sub(r'\*(.*?)\*', r'[i]\1[/i]', line)
            line = re.sub(r'_(.*?)_', r'[i]\1[/i]', line)
            line = re.sub(r'^# (.*?)$', r'[size=18][b]\1[/b][/size]', line, flags=re.MULTILINE)
            line = re.sub(r'^## (.*?)$', r'[size=16][b]\1[/b][/size]', line, flags=re.MULTILINE)
            line = re.sub(r'^### (.*?)$', r'[size=14][b]\1[/b][/size]', line, flags=re.MULTILINE)
            if re.match(r'^[\*\-\+]\s+', line):
                line = re.sub(r'^[\*\-\+]\s+', '[*]', line)
            result.append(line)
            i += 1
    full_text = '\n'.join(result)
    full_text = re.sub(r'([\n]*)(\[\*].*?)(?=\n\[\*]|\Z)', r'\n[list]\2[/list]\n', full_text, flags=re.DOTALL)
    full_text = re.sub(r'\n{3,}', '\n\n', full_text)
    return full_text

def add_comment_to_deal(deal_id, comment_text):
    """Добавляет отформатированный комментарий в таймлайн сделки."""
    formatted_comment = markdown_to_bbcode(comment_text)
    params = {"fields": {"ENTITY_ID": deal_id, "ENTITY_TYPE": "deal", "COMMENT": formatted_comment}}
    return call_bitrix24("crm.timeline.comment.add", params)

# --- 6. ФУНКЦИИ ДЛЯ РАБОТЫ С GOOGLE DRIVE ---
def find_subfolder_id(parent_folder_id, subfolder_name):
    """Находит ID подпапки в публичной папке Google Drive."""
    api_key = "AIzaSyC4qZpJZ6KzqBQyQ6xq9Fp0lX0sx8N0wA"  # Публичный ключ, работает только для публичных данных
    query = f"'{parent_folder_id}' in parents and name='{subfolder_name}' and mimeType='application/vnd.google-apps.folder'"
    url = f"https://www.googleapis.com/drive/v3/files?q={requests.utils.quote(query)}&key={api_key}&fields=files(id,name)"
    try:
        response = requests.get(url, timeout=30)
        response.raise_for_status()
        data = response.json()
        files = data.get('files', [])
        return files[0]['id'] if files else None
    except Exception as e:
        print(f"Ошибка поиска подпапки: {e}")
        return None

def download_public_folder_files(folder_id, destination_dir):
    """Скачивает все файлы из публичной папки Google Drive во временную директорию."""
    import gdown
    try:
        # gdown.download_folder рекурсивно скачивает все файлы из публичной папки
        # use_cookies=False необходимо для обхода предупреждений Google
        gdown.download_folder(id=folder_id, output=destination_dir, use_cookies=False)
        # Возвращаем список путей к скачанным файлам
        downloaded_files = []
        for root, dirs, files in os.walk(destination_dir):
            for file in files:
                if not file.endswith('.html'):  # Иногда скачиваются лишние html-файлы
                    downloaded_files.append(os.path.join(root, file))
        return downloaded_files
    except Exception as e:
        print(f"Ошибка скачивания папки {folder_id}: {e}")
        return []

# --- 7. ФУНКЦИИ ДЛЯ РАБОТЫ С DEEPSEEK ---
def extract_text_from_file(file_path):
    """Извлекает текст из файлов разных форматов."""
    ext = os.path.splitext(file_path)[1].lower()
    try:
        if ext == '.pdf':
            import PyPDF2
            text = ""
            with open(file_path, 'rb') as f:
                reader = PyPDF2.PdfReader(f)
                for page in reader.pages:
                    page_text = page.extract_text()
                    if page_text:
                        text += page_text + "\n"
            return text
        elif ext == '.docx':
            import docx
            doc = docx.Document(file_path)
            return "\n".join([p.text for p in doc.paragraphs])
        elif ext == '.txt':
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                return f.read()
        elif ext in ('.xls', '.xlsx'):
            import openpyxl
            wb = openpyxl.load_workbook(file_path, data_only=True)
            text = ""
            for sheet in wb.worksheets:
                for row in sheet.iter_rows(values_only=True):
                    text += " ".join(str(cell) for cell in row if cell) + "\n"
            return text
        else:
            return ""
    except Exception as e:
        print(f"Ошибка извлечения текста из {file_path}: {e}")
        return ""

def analyze_with_deepseek(file_paths):
    """Отправляет текст извлеченных файлов в DeepSeek API и возвращает ответ."""
    full_text = ""
    for path in file_paths:
        text = extract_text_from_file(path)
        if text:
            full_text += f"\n\n--- Файл: {os.path.basename(path)} ---\n{text}\n"
    if not full_text.strip():
        raise Exception("Не удалось извлечь текст из файлов")
    if len(full_text) > 300000:
        full_text = full_text[:300000] + "...\n[Текст документа обрезан из-за ограничения длины]"
    
    # Формируем полный промпт
    user_prompt = f"{PROMPT}\n\nТекст документов:\n{full_text}"
    headers = {
        "Authorization": f"Bearer {DEEPSEEK_API_KEY}",
        "Content-Type": "application/json"
    }
    payload = {
        "model": "deepseek-chat",
        "messages": [{"role": "user", "content": user_prompt}],
        "temperature": 0.3,
        "max_tokens": 4000
    }
    try:
        response = requests.post(DEEPSEEK_API_URL, headers=headers, json=payload, timeout=120)
        response.raise_for_status()
        result = response.json()
        if "choices" in result and len(result["choices"]) > 0:
            return result["choices"][0]["message"]["content"]
        else:
            raise Exception("Неожиданный формат ответа от DeepSeek API")
    except Exception as e:
        print(f"Ошибка при запросе к DeepSeek API: {e}")
        raise

# --- 8. ОСНОВНАЯ ЛОГИКА ОБРАБОТКИ ---
def process_purchase(data):
    """Обрабатывает закупку: создаёт сущности в Битрикс24, скачивает файлы из Google Drive и отправляет в DeepSeek."""
    # 1. Поиск/создание компании
    company_id = find_company_by_inn(data["inn"])
    if not company_id:
        company_id = create_company(data["company_name"], data["inn"])
        print(f"Создана новая компания ID {company_id}")
    else:
        print(f"Компания уже существует, ID {company_id}")

    # 2. Создание контакта и сделки
    create_contact(data.get("contact_name"), data.get("phone"), data.get("email"), company_id)
    deal_id = create_deal(company_id, data["company_name"], data.get("purchase_link", ""))
    print(f"Создана сделка ID {deal_id}")

    # 3. Работа с файлами из Google Drive
    purchase_number = data["purchase_number"]
    print(f"Поиск папки с номером {purchase_number} в Google Drive")
    
    # Находим ID папки закупки
    target_folder_id = find_subfolder_id(GOOGLE_DRIVE_ROOT_FOLDER_ID, purchase_number)
    if not target_folder_id:
        return {"status": "error", "message": f"В корневой папке Google Drive не найдена подпапка с именем {purchase_number}"}
    
    # Создаём временную папку для скачивания
    temp_dir = tempfile.mkdtemp()
    try:
        print(f"Скачивание файлов из папки {purchase_number}...")
        downloaded_files = download_public_folder_files(target_folder_id, temp_dir)
        if not downloaded_files:
            return {"status": "error", "message": f"Не удалось скачать файлы из папки {purchase_number}"}
        
        print(f"Успешно скачано файлов: {len(downloaded_files)}")
        for f in downloaded_files:
            print(f"  - {os.path.basename(f)}")
        
        # 4. Отправка в DeepSeek
        print("Отправка файлов в DeepSeek...")
        analysis = analyze_with_deepseek(downloaded_files)
        
        # 5. Добавление комментария в сделку
        comment_text = f"🤖 Анализ от DeepSeek:\n\n{analysis}"
        add_comment_to_deal(deal_id, comment_text)
        print("Комментарий добавлен в сделку")
        
        return {"status": "success", "deal_id": deal_id, "analysis_preview": analysis[:300]}
    finally:
        # Очистка временной папки
        shutil.rmtree(temp_dir, ignore_errors=True)

# --- 9. ТОЧКА ВХОДА ДЛЯ FLASK-СЕРВЕРА ---
@api.route('/process', methods=['POST'])
def process_webhook():
    """Обрабатывает POST-запросы от Google Apps Script."""
    try:
        data = request.get_json()
        if not data:
            return jsonify({"status": "error", "message": "Нет данных"}), 400
        
        # Валидация обязательных полей
        required = ["inn", "company_name", "purchase_number"]
        for field in required:
            if field not in data:
                return jsonify({"status": "error", "message": f"В запросе отсутствует поле '{field}'"}), 400
        
        result = process_purchase(data)
        return jsonify(result)
    except Exception as e:
        print(f"Ошибка в процессе обработки: {e}")
        return jsonify({"status": "error", "message": str(e)}), 500

# --- 10. ЗАПУСК ---
if __name__ == "__main__":
    # Для локальных тестов
    port = int(os.environ.get("PORT", 5000))
    api.run(host='0.0.0.0', port=port)
