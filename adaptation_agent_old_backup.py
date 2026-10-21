pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
import os
import json
import re
import fitz  # PyMuPDF (лучше работает со сканами, чем pdf2image)
import pytesseract
from docx import Document
from docx.shared import Pt
from langchain_ollama import ChatOllama
from langchain.agents import AgentExecutor, create_react_agent
from langchain.tools import DuckDuckGoSearchResults
from langchain.prompts import PromptTemplate

# --- НАСТРОЙКИ ПУТЕЙ И КОНСТАНТ ---
SOURCE_FOLDER = r"C:\Users\pzbayur\Desktop\Проработка Адаптация\Сотрудники"
OUTPUT_FOLDER = r"C:\Users\pzbayur\Desktop\Проработка Адаптация\План адаптации"
TEMPLATE_PATH = r"2_РИ_План адаптации_шаблон.docx" # Убедитесь, что шаблон лежит рядом со скриптом или укажите полный путь

# Системный промпт (см. выше)
SYSTEM_PROMPT = """Ты — эксперт по адаптации персонала... [ВСТАВЬТЕ СЮДА ТЕКСТ ПРОМПТА ИЗ БЛОКА 1]"""

# --- 1. ФУНКЦИЯ ИЗВЛЕЧЕНИЯ ТЕКСТА ИЗ PDF (OCR) ---
def extract_text_from_pdf(pdf_path):
    """Извлекает текст из сканированного PDF с помощью OCR."""
    text = ""
    try:
        # Открываем PDF
        doc = fitz.open(pdf_path)
        for page_num in range(len(doc)):
            page = doc[page_num]
            # Рендерим страницу в изображение
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2)) # Увеличиваем для лучшего OCR
            img_data = pix.tobytes("png")
            
            # Распознаем текст (Tesseract)
            # Если текст распознается плохо, можно поэкспериментировать с параметрами pytesseract
            page_text = pytesseract.image_to_string(img_data, lang='rus+eng')
            text += page_text + "\n"
        doc.close()
        return text
    except Exception as e:
        print(f"Ошибка при чтении PDF {pdf_path}: {e}")
        return ""

# --- 2. НАСТРОЙКА LLM И АГЕНТА ---
def setup_agent():
    """Инициализирует локальную LLM и инструменты поиска."""
    # Инициализация Ollama (убедитесь, что модель скачана, например: ollama pull qwen2.5)
    llm = ChatOllama(
        model="qwen2.5:14b", # Рекомендуется qwen2.5, она отлично понимает русский и JSON
        temperature=0.1,
        format="json" # Принудительный вывод в JSON
    )
    
    # Инструмент для поиска профстандартов в интернете
    search_tool = DuckDuckGoSearchResults(name="SearchProfstandarts", num_results=3)
    
    # Шаблон для ReAct агента (чтобы он мог использовать поиск)
    tool_prompt = PromptTemplate.from_template(
        SYSTEM_PROMPT + 
        "\n\nУ тебя есть доступ к инструменту поиска в интернете. Если в тексте профиля не хватает информации о знаниях и умениях, найди Профстандарт или ЕТКС для этой должности.\n" +
        "Input: {input}\n" +
        "{agent_scratchpad}"
    )
    
    agent = create_react_agent(llm, [search_tool], tool_prompt)
    agent_executor = AgentExecutor(agent=agent, tools=[search_tool], verbose=True, max_iterations=5)
    
    return agent_executor

# --- 3. ГЕНЕРАЦИЯ DOCX НА ОСНОВЕ ШАБЛОНА ---
def generate_docx(data, output_path):
    """Заполняет шаблон Word данными из JSON."""
    doc = Document(TEMPLATE_PATH)
    
    # Заполняем шапку (обычно это первая таблица или абзацы)
    # Примечание: индексы таблиц зависят от верстки вашего шаблона. 
    # Здесь мы ищем таблицы по наличию специфического текста.
    
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                cell_text = cell.text.strip()
                
                # Заполнение шапки
                if "Ф . И . О . адаптируемого" in cell_text or "Ф.И.О. адаптируемого" in cell_text:
                    cell.text = data.get("employee_name", "")
                elif "Подразделение" in cell_text and len(cell.text) < 20:
                    cell.text = data.get("department", "")
                elif "Должность" in cell_text and len(cell.text) < 20:
                    cell.text = data.get("position", "")
                elif "руководителя структурного" in cell_text:
                    cell.text = data.get("manager", "")
                elif "наставника" in cell_text:
                    cell.text = data.get("mentor", "")
                
                # Заполнение Карты адаптационных мероприятий
                elif "ОБЩИЙ БЛОК" in cell_text:
                    # Находим строку с общим блоком и добавляем мероприятия ниже
                    add_events_to_table(table, row, data.get("general_block", []))
                    
                elif "ИНДИВИДУАЛЬНЫЙ БЛОК" in cell_text:
                    add_events_to_table(table, row, data.get("individual_block", []))
                    
                elif "Итоговое индивидуальное практическое задание" in cell_text:
                    # Заполняем итоговое задание (обычно в соседней ячейке)
                    final_task = data.get("final_assignment", {})
                    task_text = f"Задание: {final_task.get('task', '')}\nКритерии: {final_task.get('check_criteria', '')}"
                    # Находим следующую ячейку в строке или объединенную
                    next_cell = row.cells[-1] 
                    next_cell.text = task_text

    doc.save(output_path)
    print(f"✅ План успешно сохранен: {output_path}")

def add_events_to_table(table, anchor_row, events):
    """Вспомогательная функция для добавления строк с мероприятиями в таблицу."""
    if not events:
        return
        
    # Находим индекс строки, после которой нужно вставлять
    anchor_idx = table.rows._tr.index(anchor_row._tr)
    
    # Вставляем строки в обратном порядке, чтобы индексы не сбились
    for event in reversed(events):
        # Создаем новую строку
        new_row = table.add_row()
        # Заполняем ячейки (Наименование, Участники, Срок, и т.д.)
        # Примечание: количество ячеек должно совпадать с шаблоном
        cells = new_row.cells
        if len(cells) >= 4:
            cells[1].text = event.get("event", "")
            cells[2].text = event.get("participants", "")
            cells[3].text = event.get("deadline", "")
            
        # Перемещаем строку на нужную позицию (сразу после anchor_row)
        table._tbl.insert(anchor_idx + 1, new_row._tr)

# --- 4. ГЛАВНЫЙ ЦИКЛ ---
def main():
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    agent = setup_agent()
    
    pdf_files = [f for f in os.listdir(SOURCE_FOLDER) if f.lower().endswith('.pdf')]
    
    if not pdf_files:
        print("⚠️ В папке с сотрудниками не найдено PDF файлов.")
        return

    for pdf_file in pdf_files:
        print(f"\n🔄 Обработка файла: {pdf_file}")
        pdf_path = os.path.join(SOURCE_FOLDER, pdf_file)
        
        # 1. Извлечение текста
        raw_text = extract_text_from_pdf(pdf_path)
        if not raw_text.strip():
            print("❌ Не удалось извлечь текст из PDF. Проверьте качество скана.")
            continue
            
        # 2. Запрос к Агенту
        try:
            response = agent.invoke({"input": f"Профиль сотрудника из файла {pdf_file}:\n\n{raw_text}"})
            result_text = response.get("output", "{}")
            
            # Очистка ответа на случай, если LLM добавила markdown
            result_text = re.sub(r'^```json\n|```$', '', result_text.strip(), flags=re.MULTILINE)
            data = json.loads(result_text)
            
        except Exception as e:
            print(f"❌ Ошибка при анализе профиля агентом: {e}")
            continue
            
        # 3. Проверка на исключение
        if data.get("skip", False):
            print(f"⏭️ Должность '{data.get('position')}' в списке исключений. Пропускаем.")
            continue
            
        # 4. Генерация DOCX
        output_filename = f"План_адаптации_{data.get('employee_name', 'Сотрудник').replace(' ', '_')}.docx"
        output_path = os.path.join(OUTPUT_FOLDER, output_filename)
        
        try:
            generate_docx(data, output_path)
        except Exception as e:
            print(f"❌ Ошибка при генерации DOCX: {e}")

if __name__ == "__main__":
    main()