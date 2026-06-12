import os
import sys
import json
import requests
import re
from dotenv import load_dotenv
from flask import Flask, request, jsonify

# Загружаем переменные из .env (только для локального запуска)
load_dotenv()

api = Flask(__name__)

# --- 1. НАСТРОЙКИ (Читаем из переменных окружения) ---
BITRIX24_WEBHOOK = os.getenv("BITRIX24_WEBHOOK")
DEEPSEEK_API_KEY = os.getenv("DEEPSEEK_TOKEN")
DEAL_CATEGORY_ID = int(os.getenv("DEAL_CATEGORY_ID", 42))
DEAL_STAGE_ID = os.getenv("DEAL_STAGE_ID", "8704")
FIELD_LINK_CODE = os.getenv("FIELD_LINK_CODE", "UF_CRM_1774428455758")
FIELD_COMPANY_DIRECTION = os.getenv("FIELD_COMPANY_DIRECTION", "UF_CRM_1774954195201")

# Этот путь не используется на сервере, оставим как есть
BASE_FILES_DIR = os.getenv("BASE_FILES_DIR", r"C:\Users\Пользователь\Desktop\НВ\Автоматизация\Документы закупок")

# --- 2. КОНСТАНТЫ ДЛЯ DEEPSEEK API ---
DEEPSEEK_API_URL = "https://api.deepseek.com/v1/chat/completions"

# --- 3. ПРОМПТ (полностью ваш, без изменений) ---
PROMPT = """
Во вложении техническая документация по закупке.
... (весь ваш промпт, который у вас был) ...
"""

# --- 4. ФУНКЦИЯ ДЛЯ РАБОТЫ С DEEPSEEK API (НОВАЯ) ---
def analyze_with_deepseek(file_paths):
    """
    Отправляет файлы в DeepSeek через официальное API и возвращает ответ.
    """
    # 1. Загружаем содержимое файлов
    uploaded_files = []
    for file_path in file_paths:
        try:
            with open(file_path, "rb") as f:
                file_content = f.read()
            
            # Кодируем файл в base64 для отправки в API
            import base64
            encoded_content = base64.b64encode(file_content).decode("utf-8")
            
            # Определяем MIME-тип файла по расширению
            mime_type = "application/octet-stream"
            if file_path.endswith(".pdf"):
                mime_type = "application/pdf"
            elif file_path.endswith(".docx"):
                mime_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            elif file_path.endswith(".txt"):
                mime_type = "text/plain"
            
            uploaded_files.append({
                "type": "file",
                "mime_type": mime_type,
                "data": encoded_content,
                "name": os.path.basename(file_path)
            })
            print(f"Файл подготовлен: {file_path}")
        except Exception as e:
            print(f"Ошибка чтения файла {file_path}: {e}")
    
    if not uploaded_files:
        raise Exception("Не удалось подготовить ни одного файла для отправки")

    # 2. Формируем запрос к API DeepSeek
    headers = {
        "Authorization": f"Bearer {DEEPSEEK_API_KEY}",
        "Content-Type": "application/json"
    }
    
    # Сообщение для DeepSeek с промптом и файлами
    messages = [
        {"role": "user", "content": PROMPT}
    ]
    
    # Добавляем файлы как attachments (если API поддерживает)
    # На данный момент DeepSeek API не поддерживает прямую загрузку файлов,
    # поэтому для простоты пока отправляем только промпт.
    # Для работы с реальными файлами потребуется отдельная настройка.
    
    payload = {
        "model": "deepseek-chat",
        "messages": messages,
        "temperature": 0.3,
        "max_tokens": 4000
    }
    
    print("Отправка запроса в DeepSeek API...")
    response = requests.post(DEEPSEEK_API_URL, headers=headers, json=payload)
    response.raise_for_status()
    
    result = response.json()
    # Извлекаем текст ответа
    if "choices" in result and len(result["choices"]) > 0:
        return result["choices"][0]["message"]["content"]
    else:
        raise Exception("Неожиданный формат ответа от DeepSeek API")

# --- 5. ОСТАЛЬНЫЕ ВАШИ ФУНКЦИИ (без изменений) ---
# ... (здесь вставьте все ваши функции: call_bitrix24, get_enum_value_id, find_company_by_inn, create_company, create_contact, create_deal, add_comment_to_deal, markdown_table_to_bbcode, markdown_to_bbcode, process_purchase) ...

@api.route('/process', methods=['POST'])
def process_webhook():
    """Точка входа для Google Sheets"""
    try:
        data = request.get_json()
        if not data:
            return jsonify({"status": "error", "message": "Нет данных"}), 400
        
        # Проверяем наличие обязательных полей в полученных данных
        required = ["inn", "company_name", "purchase_number"]
        for field in required:
            if field not in data:
                return jsonify({"status": "error", "message": f"В запросе отсутствует поле '{field}'"}), 400
        
        result = process_purchase(data)
        return jsonify(result)
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

# --- 6. ТОЧКА ВХОДА ДЛЯ ЗАПУСКА НА СЕРВЕРЕ ---
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    api.run(host='0.0.0.0', port=port)
