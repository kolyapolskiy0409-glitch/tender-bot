import os
import tempfile
import shutil
import requests
import base64
from flask import Flask, request, jsonify
from openai import OpenAI
from docx import Document
import pandas as pd
import fitz  # PyMuPDF для работы с PDF
import pytesseract
from PIL import Image
from io import BytesIO

app = Flask(__name__)

# === Конфигурация KodikRouter ===
KODIK_API_KEY = os.environ.get("KODIK_API_KEY")
if not KODIK_API_KEY:
    raise RuntimeError("Переменная окружения KODIK_API_KEY не установлена")

client = OpenAI(
    api_key=KODIK_API_KEY,
    base_url="https://api.kodikrouter.ru/v1",
)
MODEL_NAME = "deepseek/deepseek-v4-pro"

# === Функции извлечения текста ===
def extract_text_from_docx(file_path: str) -> str:
    doc = Document(file_path)
    paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
    return "\n".join(paragraphs)

def extract_text_from_excel(file_path: str) -> str:
    """Читает все листы Excel и возвращает текст в виде таблиц"""
    xls = pd.ExcelFile(file_path)
    all_text = []
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(file_path, sheet_name=sheet_name, dtype=str)
        df = df.fillna("")
        all_text.append(f"=== Лист: {sheet_name} ===\n{df.to_string(index=False)}")
    return "\n\n".join(all_text)

def extract_text_from_pdf(file_path: str) -> str:
    """
    Извлекает текст из PDF-файла с помощью PyMuPDF.
    Возвращает весь текст, объединенный из всех страниц.
    """
    doc = fitz.open(file_path)
    text = ""
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        text += page.get_text()
    doc.close()
    return text.strip() if text else "[Не удалось извлечь текст из PDF]"

def extract_text_from_image(file_path: str) -> str:
    """
    Распознает текст на изображении с помощью pytesseract (OCR).
    Поддерживает форматы: PNG, JPEG, BMP, TIFF, WebP.
    """
    try:
        image = Image.open(file_path)
        # Конвертируем в RGB, если изображение в другом режиме (например, RGBA)
        if image.mode not in ("L", "RGB"):
            image = image.convert("RGB")
        # Распознаем текст на русском и английском языках
        text = pytesseract.image_to_string(image, lang="rus+eng")
        return text.strip() if text else "[На изображении не найден текст]"
    except Exception as e:
        return f"[Ошибка OCR: {str(e)}]"

# === Скачивание с Google Drive ===
def download_file_from_google_drive(file_id: str, destination: str) -> None:
    URL = f"https://drive.google.com/uc?export=download&id={file_id}"
    session = requests.Session()
    response = session.get(URL, stream=True)
    # Обход предупреждения о вирусах
    for key, value in response.cookies.items():
        if key.startswith('download_warning'):
            confirm_url = URL + "&confirm=" + value
            response = session.get(confirm_url, stream=True)
            break
    with open(destination, 'wb') as f:
        for chunk in response.iter_content(chunk_size=8192):
            if chunk:
                f.write(chunk)

# === Анализ документов через DeepSeek ===
def analyze_documents(files_info: list) -> str:
    """
    files_info: список словарей [{"path": str, "name": str}, ...]
    """
    all_texts = []
    for fi in files_info:
        path = fi["path"]
        name = fi["name"]
        if name.endswith('.docx'):
            text = extract_text_from_docx(path)
            all_texts.append(f"--- Документ: {name} ---\n{text}")
        elif name.endswith('.xlsx'):
            text = extract_text_from_excel(path)
            all_texts.append(f"--- Документ (реестр): {name} ---\n{text}")
        elif name.endswith('.pdf'):
            text = extract_text_from_pdf(path)
            all_texts.append(f"--- Документ: {name} ---\n{text}")
        elif name.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.tiff', '.webp')):
            text = extract_text_from_image(path)
            all_texts.append(f"--- Изображение: {name} ---\n{text}")
        else:
            all_texts.append(f"--- Документ: {name} ---\n[Неподдерживаемый формат]")
    
    combined_content = "\n\n".join(all_texts)
    
    # Расширенный промт (полностью по заданию пользователя)
    system_prompt = "Ты – эксперт по анализу тендерной документации."
    
    user_prompt = f"""
Во вложении техническая документация по закупке.
Проанализируй документы и предоставь подробную информацию в отчёт, удобный для копирования в документ WORD:

1) **Название кампании Заказчика, его ИНН и контакты сотрудников** (если такие представлены с указанием номера телефона, почты, должности и ФИО). Отдельно выпиши ФИО представителя Заказчика, подготовившего ТЗ. Кратко укажи, чем занимается организация, её отрасль.

2) **Адреса проведения работ**.

3) **Название закупки и обоснование для проведения работ** (если есть).
   3.1 Есть ли общий выделенный бюджет на закупку? (в смете или НМЦ; если в НМЦ – возьми самую минимальную цену из КП)

4) **Перечень оборудования** (пункт 4.1 в виде таблицы):
   - Наименование оборудования (котёл/теплообменник/калорифер/реактор и т.д.)
   - Модель (например, Visman vitomasx 200-WS или Alfa Laval A15BW)
   - Количество
   - Вид работ (химическая/механическая/разборная/безразборная промывка и т.д.)

   **Важно!** Всегда сначала проверь, есть ли информация в документации. Если нет – ищи в открытых источниках (с пометкой "из открытых источников"). Если не найдёшь – напиши "уточнить у Заказчика".

   **Для котлов/реакторов/сосудов**: укажи объём водяной рубашки.
   **Для теплообменников**: сначала определи тип (пластинчатый, паянный, кожухотрубный). Затем:
     - Для пластинчатых: разборная/безразборная/механическая. При разборной – количество пластин, их размеры (высота/ширина), DN.
     - Для паянных: только безразборная химическая (если не указано иное). Указать DN и размеры пластин.
     - Для кожухотрубных: механическая/безразборная. Указать объём, количество трубок, их DN.
   **Если в документах нет размеров/объёмов**, сверься с приложенным файлом "Реестр оборудования и его размеров.xlsx". Данные из ТЗ/паспортов приоритетнее. Если данные из реестра – в таблице поставь пометку.

   Результат по п.4.1 выведи в формате **таблицы** (можно текстовой, с разделителями или маркировкой).

4.2) **Перечень дополнительных требований к выполнению работ** (ультразвук, гидроимпульсы, баражирование и т.п.)

4.3) **Требования к результатам работ** и способ фиксации.

5) **Требования к моющему средству** (биоорганическое, кислотное, без соляной кислоты и т.д.)  
   Обязательно укажи требования по **утилизации отработанного раствора**.

6) **Требования к персоналу** (допуски, количество, квалификация).

7) **Требования к опыту компании** и как его подтвердить (например, наличие договора на сумму не менее 3 млн руб. за последний год).

8) **Сроки проведения работ**.

9) **Условия оплаты**.

10) **Порядок определения победителя**.

11) **Штрафы, санкции, неустойки** (цена, за что).

12) **Ключевые вопросы, которые необходимо уточнить у Заказчика**.

В ответе пиши только сам отчёт, без лишних комментариев. Используй русский язык.

Вот содержимое документов:
{combined_content}
"""
    try:
        response = client.chat.completions.create(
            model=MODEL_NAME,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ],
            temperature=0.3,
            timeout=180
        )
        return response.choices[0].message.content
    except Exception as e:
        raise RuntimeError(f"Ошибка при запросе к KodikRouter: {str(e)}")

# === Flask endpoint ===
@app.route('/process', methods=['POST'])
def process_webhook():
    """
    Ожидает JSON:
    {
      "files": [
        {"id": "google_drive_file_id_1", "name": "document1.docx"},
        {"id": "google_drive_file_id_2", "name": "реестр.xlsx"},
        {"id": "google_drive_file_id_3", "name": "specification.pdf"},
        {"id": "google_drive_file_id_4", "name": "photo.jpg"}
      ]
    }
    """
    data = request.get_json()
    if not data or 'files' not in data:
        return jsonify({"error": "Не передан массив 'files'"}), 400
    
    files_list = data['files']
    temp_dir = tempfile.mkdtemp()
    downloaded_files = []
    
    try:
        for file_info in files_list:
            file_id = file_info.get('id')
            file_name = file_info.get('name')
            if not file_id or not file_name:
                continue
            local_path = os.path.join(temp_dir, file_name)
            download_file_from_google_drive(file_id, local_path)
            downloaded_files.append({"path": local_path, "name": file_name})
        
        if not downloaded_files:
            return jsonify({"error": "Не удалось скачать ни одного файла"}), 400
        
        # Анализируем все файлы
        analysis_result = analyze_documents(downloaded_files)
        
        return jsonify({"analysis": analysis_result}), 200
        
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        # Удаляем временную папку со всем содержимым
        shutil.rmtree(temp_dir, ignore_errors=True)

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 10000))
    app.run(host='0.0.0.0', port=port)
