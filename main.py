import os
import time
import requests
from playwright.sync_api import sync_playwright
from openpyxl import Workbook
from urllib.parse import urljoin
from PIL import Image
from io import BytesIO

BASE_URL = "https://dveri.com/catalog/dveri-mezhkomnatnyye"
DOWNLOAD_FOLDER = "images"
os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)

EXCLUDED_CATEGORIES = [
    "Арки и порталы",
    "Плинтус",
    "Деко Рейка",
    "Фурнитура и прочее",
    "Монтаж и реставрация",
    "В помощь продавцам"
]


def download_and_convert_image(url, name):
    try:
        response = requests.get(url, timeout=10)
        if response.status_code == 200:
            image = Image.open(BytesIO(response.content)).convert("RGB")
            file_path = os.path.join(DOWNLOAD_FOLDER, name + ".webp")
            image.save(file_path, "webp")
            return file_path
        else:
            print(f"  ❌ Не удалось загрузить изображение: {url} (код {response.status_code})")
    except Exception as e:
        print(f"  ❌ Ошибка при загрузке изображения: {e}")
    return ""


def parse_price(price_text):
    """Очистка цены от пробелов и знака рубля, добавление 50%."""
    price = price_text.replace(" ", "").replace("₽", "").strip()
    try:
        price = float(price)
        price_with_discount = price * 1.5  # добавление 50%
        return round(price_with_discount, 2)
    except ValueError:
        return None


def run_parser():
    print("🚀 Запуск парсера...")
    wb = Workbook()
    ws = wb.active
    ws.title = "Товары"
    ws.append(["Категория", "Название", "Цвет", "Цена", "Ссылка", "Картинка (.webp)"])

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        page = browser.new_page()
        page.goto("https://dveri.com/")

        # Ожидание 10 секунд для регистрации
        print("⏳ Ожидаем 10 секунд для регистрации...")
        time.sleep(10)

        page.wait_for_timeout(3000)

        print("📥 Получаем список категорий...")
        category_links = page.query_selector_all("ul.sidebar__list a")
        categories = []
        for item in category_links:
            href = item.get_attribute("href")
            name = item.inner_text().strip()
            if href and name and "/catalog/" in href and name.lower() not in [excl.lower() for excl in
                                                                              EXCLUDED_CATEGORIES]:
                categories.append((name, urljoin(BASE_URL, href)))

        print(f"🔍 Найдено категорий: {len(categories)}")

        for category_name, category_url in categories:
            print(f"\n📂 Обрабатываем категорию: {category_name}")
            page.goto(category_url)
            page.wait_for_timeout(3000)
            page.wait_for_load_state("domcontentloaded")

            while True:
                product_cards = page.query_selector_all(".card")

                if not product_cards:
                    print("⚠️ Нет товаров на странице.")
                    break

                print(f"🔄 Найдено товаров: {len(product_cards)}")
                for idx, card in enumerate(product_cards):
                    try:
                        title = card.query_selector(".card__title")
                        title = title.inner_text().strip() if title else "Без названия"

                        color = card.query_selector(".card__color")
                        color = color.inner_text().strip() if color else "Не указан"

                        price = card.query_selector(".card__price")
                        price = price.inner_text().strip() if price else "-"

                        # Проверка на наличие метки "на заказ" или "sale"
                        badge = card.query_selector(".badge--card")
                        if badge:
                            badge_text = badge.inner_text().strip().lower()
                            if "на заказ" in badge_text or "sale" in badge_text:
                                print(f"  🔴 {title} — Помечен как 'На заказ' или 'Sale'")
                                card.set_style("background-color: black; color: white;")  # помечаем чёрным

                        # Парсим цену
                        parsed_price = parse_price(price)

                        href = card.query_selector("a")
                        href = href.get_attribute("href") if href else ""
                        full_link = urljoin(BASE_URL, href)

                        img_tag = card.query_selector(".card__img-wrapper img")
                        img_src = img_tag.get_attribute("src") if img_tag else None
                        full_img_url = urljoin(BASE_URL, img_src) if img_src else ""

                        safe_name = title.replace(" ", "_").replace("/", "_").replace("\\", "_")[:50]
                        img_path = download_and_convert_image(full_img_url, safe_name) if full_img_url else ""

                        ws.append([category_name, title, color, parsed_price, full_link, img_path])
                        print(f"  ✅ [{idx + 1}] {title} — {color} — {parsed_price}")
                    except Exception as e:
                        print(f"  ❌ Ошибка в карточке: {e}")

                next_btn = page.query_selector(".pagination__arrow--right:not(.disabled)")
                if next_btn:
                    print("➡️ Переход на следующую страницу...")
                    next_btn.click()
                    page.wait_for_timeout(3000)
                else:
                    print("⛔ Последняя страница.")
                    break

        browser.close()
    wb.save("dveri_products.xlsx")
    print("\n✅ Парсинг завершён. Файл: dveri_products.xlsx")


if __name__ == "__main__":
    run_parser()
