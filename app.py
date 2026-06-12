import os
import sys
import json
import re
import tempfile
import shutil
import requests
from urllib.parse import quote, urljoin
from dotenv import load_dotenv
from flask import Flask, request, jsonify
from bs4 import BeautifulSoup
import gdown

load_dotenv()

api = Flask(__name__)

# --- 1. НАСТРОЙКИ ---
BITRIX24_WEBHOOK = os.getenv("BITRIX24_WEBHOOK")
DEEPSEEK_API_KEY = os.getenv("DEEPSEEK_TOKEN")
DEAL_CATEGORY_ID = int(os.getenv("DEAL_CATEGORY_ID", 42))
DEAL_STAGE_ID = os.getenv("DEAL_STAGE_ID", "8704")
FIELD_LINK_CODE = os.getenv("FIELD_LINK_CODE", "UF_CRM_1774428455758")
FIELD_COMPANY_DIRECTION = os.getenv("FIELD_COMPANY_DIRECTION", "UF_CRM_1774954195201")
GOOGLE_DRIVE_ROOT_FOLDER_ID = os.getenv("GOOGLE_DRIVE_ROOT_FOLDER_ID")

DEEPSEEK_API_URL = "https://api.deepseek.com/v1/chat/completions"

# --- 2. ПРОМПТ (замените на свой) ---
PROMPT = """
Во вложении техническая документация по закупке...
(весь ваш промпт)
"""

# --- 3. ФУНКЦИИ БИТРИКС24 (без изменений) ---
def call_bitrix24(method, params):
    url = f"{BITRIX24_WEBHOOK}{method}"
    try:
        resp = requests.post(url, json=params, timeout=60)
        resp.raise_for_status()
        return resp.json()
    except Exception as e:
        return {"error": str(e)}

def get_enum_value_id(field_name, target_value):
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
    presets = call_bitrix24("crm.requisite.preset.list", {})
    preset_id = 1
    if presets.get("result"):
        for p in presets["result"]:
            if "юр" in p.get("NAME", "").lower() or "organiz" in p.get("NAME", "").lower():
                preset_id = p["ID"]
                break
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

def markdown_table_to_bbcode(table_lines):
    if len(table_lines) < 2:
        return '\n'.join(table_lines)
    header_line = table_lines[0].strip('|').split('|')
    headers = [h.strip() for h in header_line]
    data_rows = []
    for line in table_lines[2:]:
        cells = line.strip('|').split('|')
        row = [c.strip() for c in cells]
        data_rows.append(row)
    bbcode = '[table]\n'
    bbcode += '[tr]' + ''.join(f'[th]{h}[/th]' for h in headers) + '[/tr]\n'
    for row in data_rows:
        bbcode += '[tr]' + ''.join(f'[td]{cell}[/td]' for cell in row) + '[/tr]\n'
    bbcode += '[/table]'
    return bbcode

def markdown_to_bbcode(text):
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
    formatted_comment = markdown_to_bbcode(comment_text)
    params = {"fields": {"ENTITY_ID": deal_id, "ENTITY_TYPE": "deal", "COMMENT": formatted_comment}}
    return call_bitrix24("crm.timeline.comment.add", params)

# --- 4. НОВАЯ ФУНКЦИЯ ПОИСКА ПОДПАПКИ ЧЕРЕЗ ПАРСИНГ HTML ---
def find_subfolder_id_by_parsing(parent_folder_id, target_name):
    """
    Парсит HTML страницу публичной папки Google Drive,
    находит ссылку на подпапку с именем target_name и возвращает её ID.
    """
    url = f"https://drive.google.com/drive/folders/{parent_folder_id}"
    print(f"Открываем страницу корневой папки: {url}")
    try:
        resp = requests.get(url, timeout=30)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, 'html.parser')
        # Ищем все ссылки, которые содержат '/drive/folders/'
        links = soup.find_all('a', href=re.compile(r'/drive/folders/'))
        for link in links:
            href = link.get('href')
            # Извлекаем название папки из текста ссылки или из атрибута aria-label
            folder_name = None
            # Попробуем найти элемент с классом, содержащим название
            # Часто название папки находится в <div class="WpHeLc VY20Jd"> или в атрибуте aria-label
            parent = link.find_parent()
            if parent:
                name_div = parent.find('div', class_=re.compile(r'WpHeLc|VY20Jd|T5pZmf'))
                if name_div:
                    folder_name = name_div.get_text(strip=True)
            if not folder_name:
                # Если не нашли, пробуем взять текст самой ссылки
                folder_name = link.get_text(strip=True)
            if folder_name == target_name:
                # Извлекаем ID из href
                match = re.search(r'/drive/folders/([a-zA-Z0-9_-]+)', href)
                if match:
                    folder_id = match.group(1)
                    print(f"Найдена подпапка '{target_name}' с ID {folder_id}")
                    return folder_id
        print(f"Подпапка с именем '{target_name}' не найдена на странице")
        return None
    except Exception as e:
        print(f"Ошибка парсинга страницы {url}: {e}")
        return None

# --- 5. ФУНКЦИИ ДЛЯ СКАЧИВАНИЯ И АНАЛИЗА ---
def download_folder_by_id(folder_id, destination_dir):
    try:
        print(f"Скачивание папки с ID {folder_id}")
        gdown.download_folder(id=folder_id, output=destination_dir, use_cookies=False, quiet=False)
        downloaded_files = []
        for root, dirs, files in os.walk(destination_dir):
            for file in files:
                if not file.endswith('.html') and not file.startswith('.'):
                    downloaded_files.append(os.path.join(root, file))
        print(f"Скачано файлов: {len(downloaded_files)}")
        return downloaded_files
    except Exception as e:
        print(f"Ошибка скачивания: {e}")
        return []

def extract_text_from_file(file_path):
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
    full_text = ""
    for path in file_paths:
        text = extract_text_from_file(path)
        if text:
            full_text += f"\n\n--- Файл: {os.path.basename(path)} ---\n{text}\n"
    if not full_text.strip():
        raise Exception("Не удалось извлечь текст из файлов")
    if len(full_text) > 300000:
        full_text = full_text[:300000] + "...\n[Обрезано]"
    user_prompt = f"{PROMPT}\n\nТекст документов:\n{full_text}"
    headers = {"Authorization": f"Bearer {DEEPSEEK_API_KEY}", "Content-Type": "application/json"}
    payload = {
        "model": "deepseek-chat",
        "messages": [{"role": "user", "content": user_prompt}],
        "temperature": 0.3,
        "max_tokens": 4000
    }
    response = requests.post(DEEPSEEK_API_URL, headers=headers, json=payload, timeout=120)
    response.raise_for_status()
    result = response.json()
    return result["choices"][0]["message"]["content"]

# --- 6. ОСНОВНАЯ ЛОГИКА ---
def process_purchase(data):
    # CRM часть
    company_id = find_company_by_inn(data["inn"])
    if not company_id:
        company_id = create_company(data["company_name"], data["inn"])
        print(f"Создана новая компания ID {company_id}")
    else:
        print(f"Компания уже существует, ID {company_id}")

    create_contact(data.get("contact_name"), data.get("phone"), data.get("email"), company_id)
    deal_id = create_deal(company_id, data["company_name"], data.get("purchase_link", ""))
    print(f"Создана сделка ID {deal_id}")

    # Поиск подпапки по номеру закупки (автоматически)
    purchase_number = data["purchase_number"]
    print(f"Поиск папки для закупки {purchase_number} в корневой папке Google Drive...")
    subfolder_id = find_subfolder_id_by_parsing(GOOGLE_DRIVE_ROOT_FOLDER_ID, purchase_number)
    if not subfolder_id:
        return {"status": "error", "message": f"Не найдена подпапка с именем {purchase_number} в корневой папке. Проверьте, что папка существует и корневая папка открыта для общего доступа."}

    temp_dir = tempfile.mkdtemp()
    try:
        downloaded_files = download_folder_by_id(subfolder_id, temp_dir)
        if not downloaded_files:
            return {"status": "error", "message": f"Не удалось скачать файлы из папки {purchase_number}"}
        print("Отправка в DeepSeek...")
        analysis = analyze_with_deepseek(downloaded_files)
        comment_text = f"🤖 Анализ от DeepSeek:\n\n{analysis}"
        add_comment_to_deal(deal_id, comment_text)
        print("Комментарий добавлен")
        return {"status": "success", "deal_id": deal_id, "analysis_preview": analysis[:300]}
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)

@api.route('/process', methods=['POST'])
def process_webhook():
    try:
        data = request.get_json()
        if not data:
            return jsonify({"status": "error", "message": "Нет данных"}), 400
        required = ["inn", "company_name", "purchase_number"]
        for field in required:
            if field not in data:
                return jsonify({"status": "error", "message": f"В запросе отсутствует поле '{field}'"}), 400
        result = process_purchase(data)
        return jsonify(result)
    except Exception as e:
        print(f"Ошибка: {e}")
        return jsonify({"status": "error", "message": str(e)}), 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    api.run(host='0.0.0.0', port=port)
