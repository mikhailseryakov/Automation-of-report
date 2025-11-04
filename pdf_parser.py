import re
import PyPDF2
from typing import List, Dict

class PDFParser:
    """
    Класс для парсинга данных из PDF файла каталога ZEGNA
    """
    
    def __init__(self):
        self.data = []
    
    def parse_pdf(self, pdf_path: str) -> List[Dict]:
        """
        Парсит PDF файл и извлекает данные о товарах
        
        Args:
            pdf_path (str): Путь к PDF файлу
            
        Returns:
            List[Dict]: Список словарей с данными о товарах
        """
        try:
            with open(pdf_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                
                for page_num in range(len(pdf_reader.pages)):
                    page = pdf_reader.pages[page_num]
                    text = page.extract_text()
                    
                    # Парсим данные со страницы
                    page_data = self._parse_page(text)
                    self.data.extend(page_data)
                
                return self.data
                
        except Exception as e:
            print(f"Ошибка при чтении PDF файла: {e}")
            return []
    
    def _parse_page(self, text: str) -> List[Dict]:
        """
        Парсит данные с одной страницы PDF
        
        Args:
            text (str): Текст страницы
            
        Returns:
            List[Dict]: Данные товаров со страницы
        """
        page_data = []
        lines = text.split('\n')
        
        i = 0
        while i < len(lines):
            line = lines[i].strip()
            
            # Пропускаем строки с заголовками и не содержащие данных о товарах
            if (not line or 
                'COLLECTION:' in line or 
                'ARTICLE' in line and 'PRODUCT' in line and 'PRICE' in line or
                'PRICE' in line and 'CURRENCY' in line and 'DATE' in line):
                i += 1
                continue
            
            # Парсим строку с данными о товаре
            product_data = self._parse_product_line(line)
            if product_data:
                page_data.append(product_data)
            
            i += 1
        
        return page_data
    
    def _parse_product_line(self, line: str) -> Dict:
        """
        Парсит строку с данными о товаре
        
        Args:
            line (str): Строка с данными
            
        Returns:
            Dict: Словарь с распарсенными данными или None если строка невалидна
        """
        # Улучшенное регулярное выражение для парсинга строки товара
        # Формат: ARTICLE FABRIC_CODE PRODUCT_TIPOLOGY PRICE € CURRENCY DATE
        pattern = r'^([A-Z0-9]+)\s+([A-Z0-9]+)\s+([A-Za-z\s\-]+)\s+([\d\s]+)\s*€\s*([A-Z]+)\s+([\d.]+)$'
        
        match = re.match(pattern, line.strip())
        if match:
            article, fabric_code, product_tipology, price, currency, date = match.groups()
            
            # Очищаем цену от пробелов и преобразуем в число
            price_clean = int(price.replace(' ', ''))
            
            return {
                'article': article,
                'fabric_code': fabric_code,
                'product_tipology': product_tipology.strip(),
                'price_eur': price_clean,
                'currency': currency,
                'date': date
            }
        
        # Альтернативный метод парсинга для сложных случаев
        parts = line.split()
        if len(parts) >= 6:
            try:
                # Ищем позицию символа €
                euro_pos = -1
                for i, part in enumerate(parts):
                    if '€' in part:
                        euro_pos = i
                        break
                
                if euro_pos >= 3:  # Должны быть как минимум ARTICLE, FABRIC_CODE, PRODUCT_TIPOLOGY и PRICE
                    article = parts[0]
                    fabric_code = parts[1]
                    
                    # PRODUCT_TIPOLOGY - все между fabric_code и ценой
                    product_tipology_parts = parts[2:euro_pos-1] if euro_pos > 3 else [parts[2]]
                    product_tipology = ' '.join(product_tipology_parts)
                    
                    # PRICE - часть перед €
                    price_str = parts[euro_pos-1].replace(' ', '')
                    price_clean = int(price_str)
                    
                    # Остальные поля
                    currency = parts[euro_pos+1] if euro_pos+1 < len(parts) else 'EUR'
                    date = parts[euro_pos+2] if euro_pos+2 < len(parts) else '01.08.25'
                    
                    return {
                        'article': article,
                        'fabric_code': fabric_code,
                        'product_tipology': product_tipology,
                        'price_eur': price_clean,
                        'currency': currency,
                        'date': date
                    }
            except (ValueError, IndexError) as e:
                print(f"Ошибка парсинга строки альтернативным методом: {line} - {e}")
        
        print(f"Не удалось распарсить строку: {line}")
        return None