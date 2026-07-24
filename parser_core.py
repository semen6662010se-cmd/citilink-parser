import logging
import sys
import openpyxl
from openpyxl.styles import Font, Alignment
import winsound
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError, Error as PlaywrightError
from tenacity import retry, stop_after_attempt, wait_fixed

############################################################## МОДУЛЬ ЖУРНАЛИРОВАНИЯ
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%H:%M:%S',
    handlers=[
        logging.FileHandler("parser_logs.txt", encoding="utf-8", mode="w"),
        logging.StreamHandler(sys.stdout)
    ]
)

############################################################## ОБРАБОТЧИК ДАННЫХ
@retry(stop=stop_after_attempt(3), wait=wait_fixed(2))
def parse_single_card(page, url, active_filters):
    ############################################################## БЛОКИРОВКА ИНТЕРАКТИВНОСТИ
    page.add_init_script("document.addEventListener('DOMContentLoaded', () => document.body.style.pointerEvents = 'none');")
    page.goto(url, wait_until="domcontentloaded", timeout=45000)
    
    page.wait_for_selector('h1', timeout=15000)
    page.wait_for_timeout(1000) 
    
    raw_title = page.locator('h1').first.inner_text(timeout=5000).strip()
    title = raw_title.replace("Характеристики ", "") if raw_title.startswith("Характеристики ") else raw_title
    
    ############################################################## НОРМАЛИЗАЦИЯ ДАННЫХ ЦЕНЫ
    price = "нет цены"
    price_locators = page.locator('span:has-text("₽"), div:has-text("₽")').all_inner_texts()
    for p in price_locators:
        clean_p = p.replace(" ", "").replace("₽", "").replace(" ", "").replace("\xa0", "").strip()
        if clean_p.isdigit() and len(clean_p) > 2:
            price = p.strip()
            break
    
    ############################################################## ИЗВЛЕЧЕНИЕ ХАРАКТЕРИСТИК
    specs_dict = {}
    try:
        page.wait_for_selector('div[class*="PropertiesItem"]', timeout=10000)
        items = page.locator('div[class*="PropertiesItem"]').all()
        for item in items:
            key_loc = item.locator('[class*="PropertiesItemTitle"]')
            val_loc = item.locator('[class*="PropertiesValue"]')
            if key_loc.count() > 0 and val_loc.count() > 0:
                k = key_loc.first.inner_text().strip().lower()
                v = val_loc.first.inner_text().strip()
                specs_dict[k] = v
    except PlaywrightTimeoutError:
        logging.warning(f"[{url}] Блок характеристик не обнаружен.")

    row_data = [title, price]
    for f in active_filters:
        found_val = "нет"
        
        for dict_key, dict_val in specs_dict.items():
            for search_word in f["keys"]:
                if search_word in dict_key:
                    found_val = dict_val
                    break
            if found_val != "нет":
                break
        
        final_value = found_val.title() if found_val != "нет" else "нет"
        row_data.append(final_value)
        
    return row_data

def main_parser():
    ############################################################## ИНИЦИАЛИЗАЦИЯ ПАРАМЕТРОВ
    print("=== КОНФИГУРАЦИЯ ПАРАМЕТРОВ ИЗВЛЕЧЕНИЯ ДАННЫХ ===")
    print("1 - Форм-фактор\n2 - Скорость чтения\n3 - Скорость записи\n4 - Тип памяти")
    print("5 - Ресурс TBW\n6 - Интерфейс\n7 - Контроллер\n8 - Гарантия")
    
    user_choice = input("Введите конфигурационную маску (например, 1234): ")
    if user_choice == "":
        user_choice = "12345678"

    test_limit_input = input("Укажите лимит итераций (Enter - без ограничений): ")
    if test_limit_input == "":
        test_limit = 999999
    else:
        test_limit = int(test_limit_input)

    filter_map = {
        "1": {"name": "Форм-фактор", "keys": ["форм", "m.2", "2.5"]},
        "2": {"name": "Чтение (МБ/с)", "keys": ["чтени", "мб/с"]},
        "3": {"name": "Запись (МБ/с)", "keys": ["запис"]},
        "4": {"name": "Тип памяти", "keys": ["памяти", "nand", "3d", "tlc", "qlc"]},
        "5": {"name": "Ресурс TBW", "keys": ["tbw", "ресурс"]},
        "6": {"name": "Интерфейс", "keys": ["интерфейс", "pci", "sata"]},
        "7": {"name": "Контроллер", "keys": ["контроллер"]},
        "8": {"name": "Гарантия", "keys": ["гарантия"]}
    }

    active_filters = []
    for char in user_choice:
        if char in filter_map:
            active_filters.append(filter_map[char])
            
    headers = ["Название", "Цена"]
    for f in active_filters:
        headers.append(f["name"])
    
    all_data = []

    ############################################################## ИНИЦИАЛИЗАЦИЯ ДРАЙВЕРА
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        catalog_page = browser.new_page()
        catalog_page.add_init_script("document.addEventListener('DOMContentLoaded', () => document.body.style.pointerEvents = 'none');")

        product_links = []
        page_num = 1
        
        try:
            logging.info("--- ФАЗА 1: Индексация ссылок каталога ---")
            ############################################################## МАРШРУТИЗАЦИЯ
            while True:
                if page_num == 1:
                    cat_url = "https://www.citilink.ru/catalog/ssd-nakopiteli/?ref=mainmenu_plate"
                else:
                    cat_url = f"https://www.citilink.ru/catalog/ssd-nakopiteli/?p={page_num}&ref=mainmenu_plate"
                
                try:
                    catalog_page.goto(cat_url, wait_until="domcontentloaded", timeout=45000)
                    catalog_page.wait_for_selector('[data-meta-name="Snippet__title"]', state='visible', timeout=20000)
                except PlaywrightTimeoutError:
                    if page_num == 1:
                        logging.error("Тайм-аут загрузки корневой страницы каталога.")
                    break

                ############################################################## ЭМУЛЯЦИЯ ПРОКРУТКИ
                for _ in range(8):
                    try:
                        c = catalog_page.locator('[data-meta-name="Snippet__title"]').count()
                        catalog_page.evaluate("window.scrollBy(0, 1000)")
                        catalog_page.wait_for_function(f"document.querySelectorAll('[data-meta-name=\"Snippet__title\"]').length > {c}", timeout=2500)
                    except Exception: break
                
                ############################################################## ИЗВЛЕЧЕНИЕ УЗЛОВ
                hrefs = catalog_page.evaluate("() => Array.from(document.querySelectorAll('[data-meta-name=\"Snippet__title\"]')).map(el => { const a = el.closest('a'); return a ? a.href : null; }).filter(h => h)")
                
                new_count = 0
                ############################################################## НОРМАЛИЗАЦИЯ ССЫЛОК
                for h in set(hrefs):
                    link = h
                    if not link.endswith('/'):
                        link += '/'
                    if not link.endswith('properties/'):
                        link += 'properties/'
                        
                    if link not in product_links:
                        product_links.append(link)
                        new_count += 1
                
                logging.info(f"Итерация {page_num}: Индексировано {new_count} уникальных URL.")
                if new_count == 0 or len(product_links) >= test_limit: 
                    break
                page_num += 1

            product_links = product_links[:test_limit]
            
            logging.info(f"--- ФАЗА 2: Извлечение данных (Всего объектов: {len(product_links)}) ---")
            try:
                ############################################################## ЦИКЛИЧЕСКАЯ ОБРАБОТКА
                for i, link in enumerate(product_links):
                    logging.info(f"Статус выполнения: [{i+1}/{len(product_links)}]")
                    product_page = browser.new_page() 
                    try:
                        row_data = parse_single_card(product_page, link, active_filters)
                        if row_data:
                            all_data.append(row_data)
                    except Exception as e:
                        logging.error(f"[{link}] Ошибка обработки объекта. Системный вывод: {e}")
                    finally:
                        product_page.close()
            ############################################################## ОБРАБОТКА ПРЕРЫВАНИЙ
            except KeyboardInterrupt:
                logging.warning("Зафиксировано ручное прерывание. Инициализация безопасного завершения...")
                    
        finally:
            browser.close()

    ############################################################## ЭКСПОРТ ДАННЫХ
    logging.info("--- ФАЗА 3: Экспорт и форматирование файла XLSX ---")
    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet.title = "custom_data"
    
    sheet.append(headers)
    for col in range(1, len(headers) + 1):
        cell = sheet.cell(row=1, column=col)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")
    
    for row in all_data:
        sheet.append(row)
        
    wb.save("citilink_deep_filtered.xlsx")
    
    if len(all_data) == 0:
        logging.warning("Исполнение завершено. Результирующий массив данных пуст.")
    else:
        logging.info("Экспорт успешно завершен. Файл доступен для чтения.")
        
    ############################################################## СИСТЕМНОЕ ОПОВЕЩЕНИЕ
    try:
        winsound.Beep(1000, 500) 
    except Exception: 
        pass

if __name__ == "__main__":
    main_parser()
