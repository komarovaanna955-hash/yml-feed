import json
import re
import subprocess
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

curl_command = [
    "curl",

    # Завершить работу с ошибкой при HTTP-кодах 400/500
    "--fail",

    # Переходить по перенаправлениям
    "--location",

    # Используем HTTP/1.0, чтобы избежать проблем
    # с некорректной chunked-передачей сервера
    "--http1.0",

    # Повторные попытки
    "--retry", "5",
    "--retry-delay", "5",
    "--retry-max-time", "180",
    "--retry-all-errors",

    # Ограничения по времени
    "--connect-timeout", "20",
    "--max-time", "120",

    # Показывать текст ошибки, но не прогресс загрузки
    "--silent",
    "--show-error",

    # Заголовки запроса
    "--header", "Accept: application/json, text/plain, */*",
    "--header", "Accept-Encoding: identity",
    "--header", "Connection: close",
    "--header", (
        "User-Agent: Mozilla/5.0 "
        "(Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 "
        "(KHTML, like Gecko) "
        "Chrome/130.0 Safari/537.36"
    ),

    SOURCE_URL,
]

try:
    result = subprocess.run(
        curl_command,
        capture_output=True,
        check=True,
        timeout=240
    )

except FileNotFoundError as error:
    raise RuntimeError(
        "Команда curl не найдена на сервере GitHub Actions."
    ) from error

except subprocess.TimeoutExpired as error:
    raise RuntimeError(
        "Превышено время ожидания при скачивании товаров."
    ) from error

except subprocess.CalledProcessError as error:
    error_text = error.stderr.decode(
        "utf-8",
        errors="replace"
    ).strip()

    raise RuntimeError(
        "Не удалось скачать товары через curl. "
        f"Ошибка curl: {error_text}"
    ) from error

response_bytes = result.stdout

if not response_bytes:
    raise RuntimeError(
        "Сервер вернул пустой ответ."
    )

print(
    f"Получено данных: {len(response_bytes)} байт"
)

try:
    response_text = response_bytes.decode(
        "utf-8-sig"
    )

except UnicodeDecodeError as error:
    raise RuntimeError(
        "Ответ сервера имеет неизвестную кодировку."
    ) from error

try:
    products = json.loads(response_text)

except json.JSONDecodeError as error:
    preview_start = response_text[:300]
    preview_end = response_text[-300:]

    print("Начало ответа сервера:")
    print(preview_start)

    print("Конец ответа сервера:")
    print(preview_end)

    raise RuntimeError(
        "Сервер вернул некорректный или неполный JSON. "
        f"Ошибка около символа {error.pos}: {error.msg}"
    ) from error

if not isinstance(products, list):
    raise RuntimeError(
        "Ожидался список товаров, но сервер вернул другой формат JSON."
    )

print("Данные успешно скачаны.")
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

print("Обработка валют...")

currencies = ET.SubElement(shop, "currencies")

currency = ET.SubElement(
    currencies,
    "currency"
)

currency.set("id", "RUB")
currency.set("rate", "1")

# =========================================================
# КАТЕГОРИИ
# =========================================================

print("Обработка категорий...")

categories_element = ET.SubElement(
    shop,
    "categories"
)

category_map = {}
category_id_counter = 1

# Сначала собираем все уникальные категории из JSON
for product in products:
    cat_name = str(
        product.get("category", "Товары")
    ).strip()

    if cat_name not in category_map:
        category_map[cat_name] = category_id_counter

        cat_el = ET.SubElement(
            categories_element,
            "category"
        )

        cat_el.set(
            "id",
            str(category_id_counter)
        )

        cat_el.text = cat_name

        category_id_counter += 1

# =========================================================
# ТОВАРЫ
# =========================================================

offers = ET.SubElement(
    shop,
    "offers"
)

added = 0
not_found = 0

# Индексируем товары для быстрого поиска
products_by_name = {}

for product in products:
    name = str(
        product.get("name", "")
    ).strip().lower()

    if name:
        products_by_name[name] = product

print("Сопоставление товаров...")

for needed_name in needed_products:
    needed_name_clean = str(
        needed_name
    ).strip().lower()

    # 1. Пробуем строгое совпадение без учёта регистра
    product = products_by_name.get(
        needed_name_clean
    )

    # 2. Если строгого совпадения нет,
    # проверяем названия после удаления лишних пробелов
    if not product:
        for product_name, product_data in products_by_name.items():
            if product_name.strip() == needed_name_clean:
                product = product_data
                break

    # 3. Если ничего не нашли,
    # пробуем частичное вхождение
    if not product:
        for product_name, product_data in products_by_name.items():
            if needed_name_clean in product_name:
                product = product_data
                break

    if product:
        product_name = str(
            product.get("name", "")
        ).strip()

        # Проверка наличия
        available_val = product.get(
            "available",
            False
        )

        try:
            stock_val = int(
                product.get("stock", 0)
            )

        except (ValueError, TypeError):
            stock_val = 0

        # Яндекс Маркет понимает available="true/false"
        is_available = (
            "true"
            if available_val and stock_val > 0
            else "false"
        )

        offer = ET.SubElement(
            offers,
            "offer"
        )

        offer.set(
            "id",
            str(product.get("sku", "0"))
        )

        offer.set(
            "available",
            is_available
        )

        # Основные теги
        ET.SubElement(
            offer,
            "name"
        ).text = product_name

        brand = product.get("brand")

        if brand:
            ET.SubElement(
                offer,
                "vendor"
            ).text = str(brand)

        price = product.get(
            "price",
            0
        )

        old_price = product.get(
            "old_price"
        )

        ET.SubElement(
            offer,
            "price"
        ).text = str(price)

        ET.SubElement(
            offer,
            "currencyId"
        ).text = "RUB"

        # Если есть старая цена и она больше текущей,
        # добавляем теги для Яндекса и VK
        if old_price:
            try:
                if float(old_price) > float(price):
                    # Тег для Яндекса
                    ET.SubElement(
                        offer,
                        "oldprice"
                    ).text = str(old_price)

                    # Тег для VK
                    ET.SubElement(
                        offer,
                        "old_price"
                    ).text = str(old_price)

            except (ValueError, TypeError):
                pass

        cat_name = str(
            product.get("category", "Товары")
        ).strip()

        ET.SubElement(
            offer,
            "categoryId"
        ).text = str(
            category_map.get(cat_name, 1)
        )

        image_url = str(
            product.get("image", "")
        ).strip()

        if image_url:
            # Убираем экранирование слешей
            image_url = image_url.replace(
                "\\/",
                "/"
            )

            # Сайт иногда склеивает ссылки.
            # Разделяем их по расширению файла.
            if (
                ".jpg" in image_url.lower()
                or ".png" in image_url.lower()
                or ".jpeg" in image_url.lower()
            ):
                parts = re.split(
                    r"(\.jpg|\.png|\.jpeg)",
                    image_url,
                    flags=re.IGNORECASE
                )

                found_links = []

                for index in range(
                    0,
                    len(parts) - 1,
                    2
                ):
                    link = (
                        parts[index]
                        + parts[index + 1]
                    )

                    # Если внутри ссылки остался старый домен,
                    # берём значение после последнего http
                    if "http" in link:
                        link = (
                            "http"
                            + link.split("http")[-1]
                        )

                    found_links.append(link)

                if found_links:
                    # Берём последнюю ссылку
                    image_url = (
                        found_links[-1].strip()
                    )

            # Проверяем, есть ли уже домен в ссылке
            if "barbaris66.ru" in image_url:
                if not image_url.startswith("http"):
                    image_url = (
                        "https://"
                        + image_url.lstrip("/")
                    )

            elif image_url.startswith("/"):
                image_url = (
                    f"https://barbaris66.ru"
                    f"{image_url}"
                )

            # Исправляем возможные ошибки протокола
            image_url = image_url.replace(
                "http://",
                "https://"
            )

            # Убираем двойные слеши, кроме https://
            image_url = re.sub(
                r"(?<!:)/+",
                "/",
                image_url
            )

            image_url = image_url.replace(
                "https:/barbaris66.ru",
                "https://barbaris66.ru"
            )

            ET.SubElement(
                offer,
                "picture"
            ).text = image_url

        ET.SubElement(
            offer,
            "url"
        ).text = str(
            product.get("url", "")
        )

        # Описание берём из JSON.
        # Если его нет, создаём стандартное.
        description = (
            product.get("description")
            or f"{product_name}. Товар в наличии."
        )

        ET.SubElement(
            offer,
            "description"
        ).text = str(description)

        # Артикул
        ET.SubElement(
            offer,
            "vendorCode"
        ).text = str(
            product.get("sku", "")
        )

        added += 1

    else:
        not_found += 1

        print(
            f"✗ Не найден в JSON: {needed_name}"
        )

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
) as file:
    file.write(pretty_xml)

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
