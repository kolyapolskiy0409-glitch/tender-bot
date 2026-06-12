import os
import sys
import json
import requests
import re
from deepseek import DeepSeekClient
from dotenv import load_dotenv

load_dotenv()

BITRIX24_WEBHOOK = os.getenv("BITRIX24_WEBHOOK")
DEEPSEEK_TOKEN = os.getenv("DEEPSEEK_TOKEN")
DEAL_CATEGORY_ID = int(os.getenv("DEAL_CATEGORY_ID", 42))
DEAL_STAGE_ID = os.getenv("DEAL_STAGE_ID", "8704")
FIELD_LINK_CODE = os.getenv("FIELD_LINK_CODE", "UF_CRM_1774428455758")
FIELD_COMPANY_DIRECTION = os.getenv("FIELD_COMPANY_DIRECTION", "UF_CRM_1774954195201")
BASE_FILES_DIR = os.getenv("BASE_FILES_DIR", r"C:\Users\Пользователь\Desktop\НВ\Автоматизация\Документы закупок")

PROMPT = """
Во вложении техническая документация по закупке.
Проанализируй документы и предоставь подробную информацию в отчет удобный для копирования в документ WORD:
1) Название кампании Заказчика, его ИНН и контакты сотрудников если такие представлены (с указанием номера телефона, почты, должности и ФИО). Отдельно выпиши ФИО представителя Заказчика подготовившего ТЗ. Так же кратко укажи чем занимается организация ее отрасль;
2) Адреса проведения работ;
3) Название закупки и обоснование для проведения работ, если такая информация есть;
3.1 Есть ли общий выделенный бюджет на закупку? в смете или в НМЦ (Если в НМЦ возьми самую минимальную цену предложенную в КП)
4) Перечисли указанные в документации: 
4.1) Перечень оборудования (Например: котел/теплообменник/калорифер/реактор и т.д.), его наименование, его модель (Например: Visman vitomasx 200-WS или Alfa Laval A15BW), количество оборудования, вид работ с этим оборудованием (Например: химическая промывка/механическая промывка/разборная промывка/Без разборная промывка),

Важно! Всегда сначала проверь есть ли во вложенной документации необходимая информация, в случае если информация присутствует, напиши ее, если информации нет, то проверить ее наличие в открытом доступе интернет, если найти ее удалось со 100% точностью, то напиши результат с пометкой \"из открытых источников\", если найти в открытых источниках не удалось, то напиши \"уточнить у Заказчика\".

Важно! Если работы предусматривают промывку котла/реактора/емкости или др. сосудов: необходимо указать объем водяной рубашки для этого оборудования.

Важно! Если работы предусматривают промывку теплообменников, то прежде всего необходимо определить и указать его тип: пластинчатый, паянный кожухотрубный т.д., 
Далее необходимо определить и указать вид промывки: для пластинчатых теплообменников это может быть разборная, без разборная, механическая промывка и другая указанная в документации. Для Паянных аппаратов, только без разборная, химическая если в Документации не указано иное. Для кожухотрубных может быть механическая, без разборная промывка. 
Если промывка Разборная, то обязательно указать количество пластин в каждом теплообменнике их размеры пластин (высота и ширина), размер Ду(DN). Если работы предусматривают промывку теплообменников пластинчатых или паянных безразборно, то обязательно указать размер ДУ(Dn) присоединений и размеры пластин(высота и ширина). Если работы предусматривают промывку теплообменника кожухотрубного, то необходимо выяснить и указать объем этого теплообменника, количество трубок и их диаметр Ду(DN).

Важно!! Если в документах Заказчика нет информации о размерах, объемах оборудования, то всегда сверься с приложенным файлом \"Реестр оборудования и его размеров.xlsx\" приложенным к запросу, информация из документов (ТЗ, паспорта) является приоритетной, обязательно напиши в таблице если взял данные из нашего реестра.
Результат по пункту 4.1 выведи в формате таблицы.

4.2 Укажи Перечень дополнительных требований к выполнению работ(Например: ультразвук, гидроипульсы, баражирование и т.д).
4.3 Укажи требования к результатам работ и способ которым это будет фиксироваться;
5) Укажи есть ли требования к подбору моющего средства (Например: требуется ли определенное средство, биоорганическое, кислотное, либо наоборот без содержания соляной кислоты и т.д.)
Обязательно укажи требования по утилизации объема отработанного раствора/моющего средства.

6) Укажи есть ли требования к персоналу? (Например: наличие определенных допусков к работе, определенное количество сотрудников и требования к их квалификации)

7) Укажи есть ли требования к опыту кампании, нужно ли его подтвердить. Как? Например: наличие договор на оказание услуг суммой 3 000 000 рублей миниму за последний год.
7) Укажи сроки проведения работ

8) Укажи условия оплаты

9) Укажи порядок определения победителя

10) Предусмотренные штрафы, санкции, неустойки. Цена? за что?

11) Ниже выпиши ключевые вопросы которые необходимо уточнить у Заказчика
"""

def markdown_table_to_bbcode(table_lines):
    """
    Преобразует markdown-таблицу в BBCode [table]
    table_lines: список строк таблицы (включая строку с разделителями)
    """
    if len(table_lines) < 2:
        return '\n'.join(table_lines)
    
    # Парсим заголовки
    header_line = table_lines[0].strip('|').split('|')
    headers = [h.strip() for h in header_line]
    
    # Определяем выравнивание из второй строки (не используется, но нужно пропустить)
    align_line = table_lines[1].strip('|').split('|')
    aligns = [a.strip() for a in align_line]
    
    # Собираем строки данных
    data_rows = []
    for line in table_lines[2:]:
        cells = line.strip('|').split('|')
        row = [c.strip() for c in cells]
        data_rows.append(row)
    
    # Формируем BBCode
    bbcode = '[table]\n'
    # Заголовок (можно сделать [th])
    bbcode += '[tr]'
    for h in headers:
        bbcode += f'[th]{h}[/th]'
    bbcode += '[/tr]\n'
    # Данные
    for row in data_rows:
        bbcode += '[tr]'
        for cell in row:
            bbcode += f'[td]{cell}[/td]'
        bbcode += '[/tr]\n'
    bbcode += '[/table]'
    return bbcode

def markdown_to_bbcode(text):
    """
    Преобразует Markdown в BBCode для Битрикс24 с поддержкой таблиц
    """
    # Обрабатываем таблицы: ищем блоки, окруженные символами |
    lines = text.split('\n')
    result = []
    i = 0
    while i < len(lines):
        line = lines[i]
        # Проверяем, начинается ли строка с | и следующая строка содержит разделители :---:
        if line.strip().startswith('|') and i+1 < len(lines) and re.match(r'^[\s]*\|[\s\-:]+[\s]*\|', lines[i+1]):
            table_lines = []
            # Собираем все строки таблицы, пока они начинаются с |
            while i < len(lines) and lines[i].strip().startswith('|'):
                table_lines.append(lines[i].strip())
                i += 1
            # Преобразуем таблицу в BBCode
            if table_lines:
                bbcode_table = markdown_table_to_bbcode(table_lines)
                result.append(bbcode_table)
            continue
        else:
            # Обычная обработка остальных элементов Markdown
            # Жирный
            line = re.sub(r'\*\*(.*?)\*\*', r'[b]\1[/b]', line)
            line = re.sub(r'__(.*?)__', r'[b]\1[/b]', line)
            # Курсив
            line = re.sub(r'\*(.*?)\*', r'[i]\1[/i]', line)
            line = re.sub(r'_(.*?)_', r'[i]\1[/i]', line)
            # Заголовки
            line = re.sub(r'^# (.*?)$', r'[size=18][b]\1[/b][/size]', line, flags=re.MULTILINE)
            line = re.sub(r'^## (.*?)$', r'[size=16][b]\1[/b][/size]', line, flags=re.MULTILINE)
            line = re.sub(r'^### (.*?)$', r'[size=14][b]\1[/b][/size]', line, flags=re.MULTILINE)
            # Маркированные списки (простой вариант)
            if re.match(r'^[\*\-\+]\s+', line):
                line = re.sub(r'^[\*\-\+]\s+', '[*]', line)
                # Временная метка, позже обернем в [list]
                result.append(line)
            else:
                result.append(line)
            i += 1
    
    # Объединяем результат и обрабатываем списки
    full_text = '\n'.join(result)
    # Заменяем последовательные строки с [*] на [list]...[/list]
    full_text = re.sub(r'([\n]*)(\[\*].*?)(?=\n\[\*]|\Z)', r'\n[list]\2[/list]\n', full_text, flags=re.DOTALL)
    # Убираем лишние переносы
    full_text = re.sub(r'\n{3,}', '\n\n', full_text)
    return full_text

def call_bitrix24(method, params):
    url = f"{BITRIX24_WEBHOOK}{method}"
    try:
        resp = requests.post(url, json=params, timeout=30)
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
    """
    Поиск компании по ИНН через реквизиты.
    Используем прямой поиск по полю RQ_INN в таблице реквизитов.
    """
    # Получаем список реквизитов, отфильтрованных по ИНН
    result = call_bitrix24("crm.requisite.list", {
        "filter": {"RQ_INN": inn},
        "select": ["ID", "ENTITY_TYPE_ID", "ENTITY_ID"]
    })
    if result.get("error"):
        print(f"Ошибка поиска реквизитов: {result['error']}")
        return None
    if result.get("result") and len(result["result"]) > 0:
        # Берем первый попавшийся реквизит, у которого ENTITY_TYPE_ID=4 (компания)
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

def add_comment_to_deal(deal_id, comment_text):
    formatted_comment = markdown_to_bbcode(comment_text)
    params = {
        "fields": {
            "ENTITY_ID": deal_id,
            "ENTITY_TYPE": "deal",
            "COMMENT": formatted_comment
        }
    }
    return call_bitrix24("crm.timeline.comment.add", params)

def analyze_with_deepseek(file_paths):
    client = DeepSeekClient(api_key=DEEPSEEK_TOKEN)
    response = client.chat(PROMPT, files=file_paths)
    return response.response

def process_purchase(data):
    company_id = find_company_by_inn(data["inn"])
    if not company_id:
        company_id = create_company(data["company_name"], data["inn"])
        print(f"Создана новая компания ID {company_id}")
    else:
        print(f"Компания уже существует, ID {company_id}")

    create_contact(data.get("contact_name"), data.get("phone"), data.get("email"), company_id)

    deal_id = create_deal(company_id, data["company_name"], data.get("purchase_link", ""))
    print(f"Создана сделка ID {deal_id}")

    purchase_number = data["purchase_number"]
    folder_path = os.path.join(BASE_FILES_DIR, data["company_name"], purchase_number)
    if not os.path.isdir(folder_path):
        return {"status": "error", "message": f"Папка не найдена: {folder_path}"}

    files = []
    for root, dirs, filenames in os.walk(folder_path):
        for f in filenames:
            if f.lower().endswith(('.pdf', '.docx', '.doc', '.txt', '.xls', '.xlsx')):
                files.append(os.path.join(root, f))
    if not files:
        return {"status": "error", "message": "Нет поддерживаемых файлов в папке"}

    print("Отправка файлов в DeepSeek...")
    analysis = analyze_with_deepseek(files)

    comment_text = f"🤖 Анализ от DeepSeek:\n\n{analysis}"
    add_comment_to_deal(deal_id, comment_text)
    print("Комментарий добавлен в сделку")

    return {"status": "success", "deal_id": deal_id, "analysis_preview": analysis[:200]}

def main():
    if len(sys.argv) < 2:
        print("Ошибка: не передан JSON с данными")
        print('Пример: python main.py "{\\"inn\\": \\"123\\", \\"company_name\\": \\"ООО Тест\\", \\"purchase_number\\": \\"123\\"}"')
        sys.exit(1)
    json_input = sys.argv[1]
    try:
        data = json.loads(json_input)
    except json.JSONDecodeError:
        print("Ошибка: аргумент не является валидным JSON")
        sys.exit(1)
    required = ["inn", "company_name", "purchase_number"]
    for field in required:
        if field not in data:
            print(f"Ошибка: в JSON отсутствует поле '{field}'")
            sys.exit(1)
    result = process_purchase(data)
    print(json.dumps(result, ensure_ascii=False))

from flask import Flask, request, jsonify
import threading

# Создаём самого 'диспетчера' (Flask-приложение)
api = Flask(__name__)

@api.route('/process', methods=['POST'])
def process_webhook():
    """Это точка входа, куда будут стучаться Google Таблицы."""
    try:
        # Получаем данные от Google Sheets
        data = request.get_json()
        if not data:
            return jsonify({"status": "error", "message": "Нет данных"}), 400
        
        # Здесь вызывается ваша основная функция обработки (process_purchase)
        # Обратите внимание: я передаю нужные параметры вручную
        result = process_purchase(data)
        return jsonify(result)
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

def run_flask():
    """Запускает веб-сервер, который слушает команды."""
    # Render сам назначит порт через переменную PORT
    port = int(os.environ.get("PORT", 5000))
    api.run(host='0.0.0.0', port=port)

if __name__ == "__main__":
    # Запускаем ваш основной обработчик в фоне? Нет, теперь он будет запускаться
    # по запросу от Google Sheets. Просто запускаем Flask-сервер.
    run_flask()