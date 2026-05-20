import requests
import pandas as pd
import xml.etree.ElementTree as ET
from xml.dom import minidom
from datetime import datetime

# =========================================================
# НАСТРОЙКИ
# =========================================================

SOURCE_URL = "https://barbaris66.ru/export/products/yandex"

EXCEL_FILE = "products.xlsx"

OUTPUT_FILE = "filtered_yml.xml"

SHOP_NAME = "Barbaris"
SHOP_COMPANY = "Barbaris"
SHOP_URL = "https://barbaris66.ru"

# =========================================================
# ЧИТАЕМ EXCEL
# =========================================================

print("Читаем Excel...")

df = pd.read_excel(EXCEL_FILE)

needed_products = (
    df["name"]
    .dropna()
    .astype(str)
    .str.strip()
    .tolist()
)

print(f"Товаров в Excel: {len(needed_products)}")

# =========================================================
# СКАЧИВАЕМ JSON
# =========================================================

print("Скачиваем товары...")

response = requests.get(
    SOURCE_URL,
    timeout=60
)

response.raise_for_status()

products = response.json()

print(f"Товаров на сайте: {len(products)}")

# =========================================================
# СОЗДАЕМ XML
# =========================================================

current_date = datetime.now().strftime("%Y-%m-%d %H:%M")

root = ET.Element(
    "yml_catalog",
    date=current_date
)

shop = ET.SubElement(root, "shop")

# =========================================================
# ОБЩАЯ ИНФОРМАЦИЯ
# =========================================================

ET.SubElement(shop, "name").text = SHOP_NAME
ET.SubElement(shop, "company").text = SHOP_COMPANY
ET.SubElement(shop, "url").text = SHOP_URL

# =========================================================
# ВАЛЮТЫ
# =========================================================

currencies = ET.SubElement(shop, "currencies")

currency = ET.SubElement(
    currencies,
    "currency"
)

currency.set("id", "RUR")
currency.set("rate", "1")

# =========================================================
# КАТЕГОРИИ
# =========================================================

print("Обработка категорий...")
categories_element = ET.SubElement(shop, "categories")

category_map = {}
category_id_counter = 1

# Сначала соберем все уникальные категории из JSON
for product in products:
    cat_name = str(product.get("category", "Товары")).strip()
    if cat_name not in category_map:
        category_map[cat_name] = category_id_counter
        cat_el = ET.SubElement(categories_element, "category")
        cat_el.set("id", str(category_id_counter))
        cat_el.text = cat_name
        category_id_counter += 1

# =========================================================
# ТОВАРЫ
# =========================================================

offers = ET.SubElement(shop, "offers")

added = 0
not_found = 0

# Индексируем товары для быстрого поиска (O(1) вместо O(N))
products_by_name = {}
for p in products:
    name = str(p.get("name", "")).strip().lower()
    if name:
        products_by_name[name] = p

print("Сопоставление товаров...")

for needed_name in needed_products:
    needed_name_clean = str(needed_name).strip().lower()

    # 1. Пробуем СТРОГОЕ совпадение (без учета регистра)
    product = products_by_name.get(needed_name_clean)
    
    # 2. Если строгого нет, ищем самое длинное совпадение или первое вхождение
    if not product:
        for p_name, p_data in products_by_name.items():
            # Проверяем, что названия идентичны после очистки лишних пробелов
            if p_name.strip() == needed_name_clean:
                product = p_data
                break

    # 3. И только если совсем ничего не нашли, пробуем частичное "вхождение"
    if not product:
        for p_name, p_data in products_by_name.items():
            if needed_name_clean in p_name:
                product = p_data
                break


    if product:
        product_name = str(product.get("name", "")).strip()
        
        # Проверка наличия
        available_val = product.get("available", False)
        stock_val = int(product.get("stock", 0))
        
        # Яндекс Маркет понимает available="true/false"
        is_available = "true" if (available_val and stock_val > 0) else "false"

        offer = ET.SubElement(offers, "offer")
        offer.set("id", str(product.get("sku", "0")))
        offer.set("available", is_available)

        # Основные теги
        ET.SubElement(offer, "name").text = product_name
        
        brand = product.get("brand")
        if brand:
            ET.SubElement(offer, "vendor").text = str(brand)

        ET.SubElement(offer, "price").text = str(product.get("price", 0))
        ET.SubElement(offer, "currencyId").text = "RUR"

        cat_name = str(product.get("category", "Товары")).strip()
        ET.SubElement(offer, "categoryId").text = str(category_map.get(cat_name, 1))

        image_url = str(product.get("image", "")).strip()
        if image_url:
            # Убираем экранирование слешей
            image_url = image_url.replace("\\/", "/")

            # Сайт склеивает ссылки. Мы разделяем их по расширению файла.
            import re
            if ".jpg" in image_url.lower() or ".png" in image_url.lower():
                # Разрезаем строку везде, где есть расширение
                parts = re.split(r'(\.jpg|\.png|\.jpeg)', image_url, flags=re.IGNORECASE)
                # Собираем ссылки обратно (часть с текстом + расширение)
                found_links = []
                for i in range(0, len(parts)-1, 2):
                    link = parts[i] + parts[i+1]
                    # Очищаем от мусора: если внутри ссылки остался старый домен, берем только то, что после последнего http
                    if "http" in link:
                        link = "http" + link.split("http")[-1]
                    found_links.append(link)
                
                if found_links:
                    # Берем САМУЮ ПОСЛЕДНЮЮ ссылку из склеенных
                    image_url = found_links[-1].strip()

            # Проверяем, есть ли уже домен в ссылке
            if "barbaris66.ru" in image_url:
                if not image_url.startswith("http"):
                    image_url = "https://" + image_url.lstrip("/")
            elif image_url.startswith('/'):
                image_url = f"https://barbaris66.ru{image_url}"
            
            # Исправляем возможные ошибки протокола
            image_url = image_url.replace("http://", "https://")
            # Убираем двойные слеши, кроме https://
            image_url = re.sub(r'(?<!:)/+', '/', image_url)
            image_url = image_url.replace("https:/barbaris66.ru", "https://barbaris66.ru")
            
            ET.SubElement(offer, "picture").text = image_url


        ET.SubElement(offer, "url").text = str(product.get("url", ""))

        # Описание (берем из JSON или создаем стандартное)
        description = product.get("description") or f"{product_name}. Товар в наличии."
        ET.SubElement(offer, "description").text = str(description)

        # Дополнительно: артикул
        ET.SubElement(offer, "vendorCode").text = str(product.get("sku", ""))

        added += 1
        # print(f"✓ {product_name}") # Раскомментируйте, если нужен подробный лог
    else:
        not_found += 1
        print(f"✗ Не найден в JSON: {needed_name}")


# =========================================================
# СОХРАНЯЕМ XML
# =========================================================

xml_string = ET.tostring(
    root,
    encoding="utf-8"
)

pretty_xml = minidom.parseString(
    xml_string
).toprettyxml(
    indent="  ",
    encoding="utf-8"
)

with open(
    OUTPUT_FILE,
    "wb"
) as f:

    f.write(pretty_xml)

# =========================================================
# ГОТОВО
# =========================================================

print()
print("===================================")
print("ГОТОВО ✅")
print("===================================")

print(f"Добавлено товаров: {added}")
print(f"Не найдено: {not_found}")

print()
print(f"Файл создан: {OUTPUT_FILE}")