"""
Парсер каталога интернет-магазина → Excel / CSV.

Демо работает на books.toscrape.com — сайте, созданном специально для
практики парсинга. Тот же подход применим к любому каталогу: категории,
постраничный обход, карточки товаров, выгрузка в таблицу.

Примеры запуска:
    python parser.py                                  весь каталог → books.xlsx
    python parser.py --limit 50                       первые 50 товаров (быстрая проверка)
    python parser.py --categories "Travel, Mystery"   только выбранные категории
    python parser.py --output books.csv               выгрузка в CSV
"""
from __future__ import annotations

import argparse
import csv
import random
import re
import sys
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import astuple, dataclass
from datetime import datetime
from pathlib import Path
from statistics import mean
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

BASE_URL = "https://books.toscrape.com/"
USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) catalog-parser/1.0"
RATING_WORDS = {"One": 1, "Two": 2, "Three": 3, "Four": 4, "Five": 5}


@dataclass
class Product:
    title: str
    category: str
    price: float
    rating: int
    in_stock: int
    upc: str
    url: str
    image_url: str
    description: str


# Заголовки колонок в том же порядке, что и поля Product
COLUMNS = [
    ("Название", 45),
    ("Категория", 20),
    ("Цена, £", 10),
    ("Рейтинг", 9),
    ("В наличии, шт.", 14),
    ("Артикул (UPC)", 20),
    ("Ссылка", 14),
    ("Изображение", 14),
    ("Описание", 80),
]


# ---------------------------------------------------------------------------
# Сеть: повторы при сбоях и вежливые паузы между запросами
# ---------------------------------------------------------------------------

class Fetcher:
    """HTTP-клиент с повторами при сбоях и паузами, чтобы не перегружать сайт."""

    def __init__(self, delay: float):
        self.delay = delay
        self.session = requests.Session()
        self.session.headers["User-Agent"] = USER_AGENT
        retry = Retry(
            total=4,
            backoff_factor=1,  # паузы 1, 2, 4, 8 секунд между повторами
            status_forcelist=[429, 500, 502, 503, 504],
            allowed_methods=["GET"],
        )
        adapter = HTTPAdapter(max_retries=retry, pool_maxsize=16)
        self.session.mount("http://", adapter)
        self.session.mount("https://", adapter)

    def soup(self, url: str) -> BeautifulSoup:
        if self.delay > 0:
            time.sleep(random.uniform(self.delay * 0.5, self.delay * 1.5))
        resp = self.session.get(url, timeout=20)
        resp.raise_for_status()
        resp.encoding = "utf-8"
        return BeautifulSoup(resp.text, "html.parser")


# ---------------------------------------------------------------------------
# Разбор страниц
# ---------------------------------------------------------------------------

def get_categories(fetcher: Fetcher) -> dict[str, str]:
    """Название категории → ссылка на её первую страницу."""
    soup = fetcher.soup(BASE_URL)
    links = soup.select("div.side_categories ul li ul li a")
    return {a.get_text(strip=True): urljoin(BASE_URL, a["href"]) for a in links}


def get_product_urls(fetcher: Fetcher, category_url: str, limit: int | None) -> list[str]:
    """Обходит все страницы категории (кнопка «next») и собирает ссылки на товары."""
    urls: list[str] = []
    page_url: str | None = category_url
    while page_url:
        soup = fetcher.soup(page_url)
        for a in soup.select("article.product_pod h3 a"):
            urls.append(urljoin(page_url, a["href"]))
            if limit is not None and len(urls) >= limit:
                return urls
        next_link = soup.select_one("li.next a")
        page_url = urljoin(page_url, next_link["href"]) if next_link else None
    return urls


def table_value(soup: BeautifulSoup, header: str) -> str:
    """Значение из таблицы характеристик товара по названию строки."""
    for row in soup.select("table.table tr"):
        th, td = row.find("th"), row.find("td")
        if th and td and th.get_text(strip=True) == header:
            return td.get_text(strip=True)
    return ""


def parse_price(text: str) -> float:
    """"£51.77" → 51.77"""
    cleaned = re.sub(r"[^\d.,]", "", text).replace(",", ".")
    return float(cleaned) if cleaned else 0.0


def parse_product(fetcher: Fetcher, url: str) -> Product:
    soup = fetcher.soup(url)

    rating_tag = soup.select_one("p.star-rating")
    rating_word = next((c for c in (rating_tag.get("class") or []) if c in RATING_WORDS), "") if rating_tag else ""

    stock_match = re.search(r"\d+", table_value(soup, "Availability"))
    crumbs = soup.select("ul.breadcrumb li a")
    description_tag = soup.select_one("#product_description + p")
    image_tag = soup.select_one("div.item.active img") or soup.select_one("#product_gallery img")

    return Product(
        title=soup.select_one("div.product_main h1").get_text(strip=True),
        category=crumbs[-1].get_text(strip=True) if crumbs else "",
        price=parse_price(soup.select_one("div.product_main p.price_color").get_text()),
        rating=RATING_WORDS.get(rating_word, 0),
        in_stock=int(stock_match.group()) if stock_match else 0,
        upc=table_value(soup, "UPC"),
        url=url,
        image_url=urljoin(url, image_tag["src"]) if image_tag else "",
        description=description_tag.get_text(strip=True) if description_tag else "",
    )


# ---------------------------------------------------------------------------
# Сохранение
# ---------------------------------------------------------------------------

HEADER_FONT = Font(bold=True, color="FFFFFF")
HEADER_FILL = PatternFill("solid", fgColor="2F5597")


def style_header(ws, widths: list[int]) -> None:
    for i, width in enumerate(widths, start=1):
        cell = ws.cell(row=1, column=i)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(vertical="center")
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.freeze_panes = "A2"


def save_excel(products: list[Product], path: Path) -> None:
    wb = Workbook()

    # Лист 1: все товары
    ws = wb.active
    ws.title = "Товары"
    ws.append([name for name, _ in COLUMNS])
    for p in products:
        ws.append(list(astuple(p)))
        row = ws.max_row
        ws.cell(row=row, column=3).number_format = "0.00"
        # Ссылки делаем кликабельными
        for col, text in ((7, "Открыть"), (8, "Картинка")):
            cell = ws.cell(row=row, column=col)
            if cell.value:
                cell.hyperlink = cell.value
                cell.value = text
                cell.style = "Hyperlink"
    style_header(ws, [w for _, w in COLUMNS])
    ws.auto_filter.ref = ws.dimensions

    # Лист 2: сводка по категориям
    summary = wb.create_sheet("Сводка")
    summary.append(["Категория", "Товаров", "Средняя цена", "Мин. цена", "Макс. цена", "Средний рейтинг", "Всего на складе"])
    by_category: dict[str, list[Product]] = {}
    for p in products:
        by_category.setdefault(p.category, []).append(p)
    for category, items in sorted(by_category.items(), key=lambda kv: -len(kv[1])):
        prices = [i.price for i in items]
        summary.append([
            category,
            len(items),
            round(mean(prices), 2),
            min(prices),
            max(prices),
            round(mean(i.rating for i in items), 1),
            sum(i.in_stock for i in items),
        ])
    total_row = summary.max_row + 1
    summary.append([
        "ИТОГО",
        len(products),
        round(mean(p.price for p in products), 2),
        min(p.price for p in products),
        max(p.price for p in products),
        round(mean(p.rating for p in products), 1),
        sum(p.in_stock for p in products),
    ])
    for cell in summary[total_row]:
        cell.font = Font(bold=True)
    for row in summary.iter_rows(min_row=2, min_col=3, max_col=5):
        for cell in row:
            cell.number_format = "0.00"
    style_header(summary, [22, 10, 14, 11, 11, 16, 17])
    summary.append([])
    summary.append([f"Собрано: {datetime.now():%d.%m.%Y %H:%M}, источник: {BASE_URL}"])

    wb.save(path)


def save_csv(products: list[Product], path: Path, delimiter: str) -> None:
    # utf-8-sig — чтобы Excel открыл файл без кракозябр
    with path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f, delimiter=delimiter)
        writer.writerow([name for name, _ in COLUMNS])
        for p in products:
            writer.writerow(astuple(p))


# ---------------------------------------------------------------------------
# Запуск
# ---------------------------------------------------------------------------

def parse_args() -> argparse.Namespace:
    ap = argparse.ArgumentParser(description="Парсер каталога интернет-магазина в Excel/CSV")
    ap.add_argument("--categories", help='Категории через запятую, например "Travel, Mystery". По умолчанию — все')
    ap.add_argument("--limit", type=int, help="Максимум товаров (для быстрой проверки)")
    ap.add_argument("--output", default="books.xlsx", help="Файл результата: .xlsx или .csv (по умолчанию books.xlsx)")
    ap.add_argument("--delimiter", default=";", help='Разделитель для CSV (по умолчанию ";" — для русского Excel)')
    ap.add_argument("--workers", type=int, default=4, help="Сколько карточек загружать параллельно (по умолчанию 4)")
    ap.add_argument("--delay", type=float, default=0.3, help="Средняя пауза между запросами, сек (по умолчанию 0.3)")
    return ap.parse_args()


def main() -> int:
    args = parse_args()
    output = Path(args.output)
    if output.suffix.lower() not in (".xlsx", ".csv"):
        print("Файл результата должен быть .xlsx или .csv")
        return 1

    fetcher = Fetcher(args.delay)
    started = time.time()

    print("Загружаю список категорий...")
    categories = get_categories(fetcher)
    if args.categories:
        wanted = {c.strip().lower() for c in args.categories.split(",") if c.strip()}
        categories = {name: url for name, url in categories.items() if name.lower() in wanted}
        missing = wanted - {name.lower() for name in categories}
        if missing:
            print(f"Не найдены категории: {', '.join(sorted(missing))}")
        if not categories:
            return 1
    print(f"Категорий: {len(categories)}")

    print("Собираю ссылки на товары...")
    urls: list[str] = []
    for name, url in categories.items():
        remaining = None if args.limit is None else args.limit - len(urls)
        if remaining is not None and remaining <= 0:
            break
        found = get_product_urls(fetcher, url, remaining)
        urls.extend(found)
        print(f"  {name}: {len(found)}")
    urls = list(dict.fromkeys(urls))  # без повторов, порядок сохраняется
    print(f"Товаров к загрузке: {len(urls)}")

    products: list[Product] = []
    errors: list[str] = []
    lock = threading.Lock()
    done = 0
    with ThreadPoolExecutor(max_workers=max(1, args.workers)) as pool:
        futures = {pool.submit(parse_product, fetcher, url): url for url in urls}
        for future in as_completed(futures):
            url = futures[future]
            try:
                product = future.result()
                with lock:
                    products.append(product)
            except Exception as exc:  # одна битая карточка не должна ронять весь сбор
                errors.append(f"{url}: {exc}")
            done += 1
            print(f"\r  Загружено {done}/{len(urls)}", end="", flush=True)
    print()

    if not products:
        print("Не удалось собрать ни одного товара.")
        return 1

    # Порядок как на сайте, а не в порядке завершения потоков
    order = {url: i for i, url in enumerate(urls)}
    products.sort(key=lambda p: order.get(p.url, 0))

    if output.suffix.lower() == ".xlsx":
        save_excel(products, output)
    else:
        save_csv(products, output, args.delimiter)

    print(f"Готово: {len(products)} товаров → {output.resolve()}")
    print(f"Время: {time.time() - started:.0f} сек")
    if errors:
        print(f"Не удалось загрузить {len(errors)} карточек:")
        for line in errors[:10]:
            print(f"  {line}")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\nОстановлено.")
        sys.exit(130)
