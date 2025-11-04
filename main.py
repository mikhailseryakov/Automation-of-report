import os
import glob
import requests
from datetime import datetime
from pdf_parser import PDFParser
from excel_handler import ExcelHandler

def get_eur_rate_from_cb():
    """
    Получает актуальный курс евро от ЦБ РФ
    Returns:
        float: Курс евро или None при ошибке
    """
    try:
        url = "https://www.cbr-xml-daily.ru/daily_json.js"
        response = requests.get(url, timeout=10)
        response.raise_for_status()  # Проверяет HTTP-статус
        data = response.json()
        return float(data["Valute"]["EUR"]["Value"])
    except requests.exceptions.RequestException as e:
        print(f"Ошибка сети при запросе курса ЦБ: {e}")
        return None
    except (KeyError, ValueError, TypeError) as e:
        print(f"Ошибка обработки данных от ЦБ: {e}")
        return None
    except Exception as e:
        print(f"Неизвестная ошибка при получении курса: {e}")
        return None

def main():
    print("=== ZEGNA Catalog Processor ===")
    
    # 1. Проверяем наличие PDF файла
    pdf_files = glob.glob("*.pdf")
    if not pdf_files:
        print("Ошибка: В текущей директории не найден PDF файл")
        input("Нажмите Enter для выхода...")
        return
    
    pdf_path = pdf_files[0]
    print(f"Найден PDF файл: {pdf_path}")
    
    # 2. Проверяем наличие settings.xlsx
    if not os.path.exists("settings.xlsx"):
        print("Ошибка: Файл settings.xlsx не найден в текущей директории")
        input("Нажмите Enter для выхода...")
        return
    
    # 3. Парсим PDF
    print("Чтение данных из PDF...")
    pdf_parser = PDFParser()
    products_data = pdf_parser.parse_pdf(pdf_path)

    if not products_data:
        print("Ошибка: Не удалось извлечь данные из PDF файла")
        input("Нажмите Enter для выхода...")
        return

    print(f"Извлечено записей: {len(products_data)}")
    
    # Проверка первых 5 записей
    print("\n--- ПРОВЕРКА ПАРСИНГА PDF (первые 5 записей) ---")
    for i, product in enumerate(products_data[:5]):
        print(f"{i+1}. ARTICLE: {product['article']}, FABRIC: {product['fabric_code']}, "
              f"TIPOLOGY: {product['product_tipology']}, PRICE: {product['price_eur']} €")


    # 4. Работаем с Excel
    print("\nОбработка Excel файла...")
    excel_handler = ExcelHandler("settings.xlsx")

    if not excel_handler.load_workbook():
        print("Ошибка при загрузке Excel файла")
        input("Нажмите Enter для выхода...")
        return

    settings = excel_handler.get_settings()
    translations = excel_handler.get_translations()

    # Получаем курс евро: сначала из ЦБ, затем из Excel (если ЦБ недоступен)
    eur_rate = get_eur_rate_from_cb()
    if eur_rate is not None:
        print(f"\nКурс евро (ЦБ РФ): {eur_rate} руб.")
        exchange_rate = eur_rate
        source = "ЦБ РФ"
    else:
        # Резерв: берём из Excel
        exchange_rate = settings['exchange_rate']
        source = f"Excel (B1: {exchange_rate})"
        print(f"!Курс евро (резерв из Excel): {exchange_rate} руб.")

    print("\n--- НАСТРОЙКИ ---")
    print(f"Источник курса: {source}")
    print(f"Наценка: 3% (фиксировано)")
    print(f"Начальная позиция: {settings['start_position']}")
    print(f"Единица измерения: {settings['measure_unit']}")
    print(f"Записей в словаре переводов: {len(translations)}")


    # Очищаем лист "Рабочий"
    excel_handler.clear_work_sheet()

    # Записываем заголовки в первую строку колонки A
    # ВНИМАНИЕ: PositionId убран из заголовков!
    header_line = "ItemCode,Price,Measure,Name"
    excel_handler.work_sheet.cell(row=1, column=1, value=header_line)


    # Заполняем данные (начиная со 2‑й строки)
    current_position = settings['start_position']
    for row_idx, product in enumerate(products_data, 2):
        if not product:
            continue

        # Формируем ItemCode: добавляем PositionId в начало
        base_item_code = f"{product['article']}{product['fabric_code']}".replace(' ', '')
        item_code = f"{current_position}{base_item_code}"  # Например, 1E7T207300M


        # Цена: евро × курс × 1,03 (3% наценки) → в копейки
        price_eur = product['price_eur']
        price_rub_kopecks = int(price_eur * exchange_rate * 1.03 * 100)

        # Название (с переводом, если есть)
        product_name = product['product_tipology']
        translated_name = translations.get(product_name, product_name)

        # Формируем CSV‑строку для колонки A
        # ВНИМАНИЕ: PositionId НЕ выводим отдельно!
        csv_line = (
            f"{item_code},"
            f"{price_rub_kopecks},"
            f"{settings['measure_unit']},"
            f"{translated_name}"
        )

        # Пишем в колонку A текущей строки
        excel_handler.work_sheet.cell(row=row_idx, column=1, value=csv_line)

        current_position += 1  # Увеличиваем PositionId для следующего товара

    # Сохраняем
    if excel_handler.save_workbook():
        print("\n=== РЕЗУЛЬТАТ ===")
        print("Данные успешно записаны в settings.xlsx")
        print(f"Обработано товаров: {len(products_data)}")
        print("Формат: CSV в колонке A (заголовки в строке 1, данные — со строки 2)")
        print("Поля: ItemCode,Price,Measure,Name")
        print("ItemCode включает PositionId в начале (например, 1E7T207300M)")
    else:
        print("Ошибка при сохранении Excel файла")
    print("\nОбработка завершена!")
    input("Нажмите Enter для выхода...")

if __name__ == "__main__":
    main()
