"""Парсер каталога товаров в Excel (демо-проект).

Примеры:
    python catalog_parser.py --file sample/catalog.html --out catalog.xlsx
    python catalog_parser.py --url https://example.com/catalog --item ".card" --name ".card-title" \
        --price ".card-price" --next "a.next" --out catalog.xlsx
"""
from __future__ import annotations

import argparse
import re
import sys
import time
from pathlib import Path
from urllib.parse import urljoin, urlparse
from urllib.robotparser import RobotFileParser

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font

USER_AGENT = "Mozilla/5.0 (compatible; price-parser-demo/1.0)"
HEADERS = ["Название", "Цена, ₸", "Ссылка"]

PRICE_RE = re.compile(r"\d[\d\s ]*(?:[.,]\d+)?")


def parse_price(text: str) -> float | int | None:
    """'от 12 500 ₸' -> 12500, 'по запросу' -> None."""
    match = PRICE_RE.search(text or "")
    if not match:
        return None
    number = re.sub(r"[\s ]", "", match.group()).replace(",", ".")
    value = float(number)
    return int(value) if value.is_integer() else value


def parse_catalog(
    html: str,
    base_url: str = "",
    item: str = ".product",
    name: str = ".product__title",
    price: str = ".product__price",
    link: str = "a",
) -> list[dict]:
    """Достаёт из HTML список товаров по CSS-селекторам."""
    soup = BeautifulSoup(html, "lxml")
    rows = []
    for card in soup.select(item):
        name_tag = card.select_one(name)
        if name_tag is None:
            continue  # карточка без названия: баннер или реклама
        price_tag = card.select_one(price)
        link_tag = card.select_one(link)
        href = link_tag.get("href") if link_tag else ""
        rows.append(
            {
                "name": name_tag.get_text(strip=True),
                "price": parse_price(price_tag.get_text(" ", strip=True)) if price_tag else None,
                "link": urljoin(base_url, href) if href else "",
            }
        )
    return rows


def find_next_url(html: str, base_url: str, next_selector: str | None) -> str | None:
    if not next_selector:
        return None
    tag = BeautifulSoup(html, "lxml").select_one(next_selector)
    href = tag.get("href") if tag else None
    return urljoin(base_url, href) if href else None


def allowed_by_robots(url: str) -> bool:
    """Проверяет robots.txt: если сайт запрещает сбор, парсер не продолжает."""
    parts = urlparse(url)
    robots = RobotFileParser()
    robots.set_url(f"{parts.scheme}://{parts.netloc}/robots.txt")
    try:
        robots.read()
    except OSError:
        return True  # robots.txt недоступен: считаем, что ограничений нет
    return robots.can_fetch(USER_AGENT, url)


def fetch(url: str, timeout: int = 20) -> str:
    response = requests.get(url, headers={"User-Agent": USER_AGENT}, timeout=timeout)
    response.raise_for_status()
    response.encoding = response.apparent_encoding or response.encoding
    return response.text


def crawl(start_url: str, max_pages: int, delay: float, **selectors) -> list[dict]:
    """Проходит по страницам каталога, соблюдая паузу между запросами."""
    rows, url, seen = [], start_url, set()
    while url and url not in seen and len(seen) < max_pages:
        if not allowed_by_robots(url):
            raise PermissionError(f"Сайт запрещает автоматический сбор этой страницы: {url}")
        seen.add(url)
        html = fetch(url)
        rows += parse_catalog(html, base_url=url, **{k: v for k, v in selectors.items() if k != "next"})
        url = find_next_url(html, url, selectors.get("next"))
        if url:
            time.sleep(delay)
    return rows


def save_xlsx(rows: list[dict], path: str | Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Каталог"
    sheet.append(HEADERS)
    for cell in sheet[1]:
        cell.font = Font(bold=True)
    for row in rows:
        sheet.append([row["name"], row["price"], row["link"]])
    for cell in sheet["B"][1:]:
        cell.number_format = "#,##0"
    sheet.column_dimensions["A"].width = 46
    sheet.column_dimensions["B"].width = 14
    sheet.column_dimensions["C"].width = 60
    sheet.auto_filter.ref = sheet.dimensions
    sheet.freeze_panes = "A2"
    workbook.save(path)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Парсер каталога в Excel")
    source = parser.add_mutually_exclusive_group(required=True)
    source.add_argument("--url", help="адрес страницы каталога")
    source.add_argument("--file", help="локальный HTML-файл (для проверки селекторов)")
    parser.add_argument("--out", default="catalog.xlsx")
    parser.add_argument("--item", default=".product", help="CSS-селектор карточки товара")
    parser.add_argument("--name", default=".product__title", help="селектор названия внутри карточки")
    parser.add_argument("--price", default=".product__price", help="селектор цены внутри карточки")
    parser.add_argument("--link", default="a", help="селектор ссылки внутри карточки")
    parser.add_argument("--next", default=None, help="селектор ссылки на следующую страницу")
    parser.add_argument("--max-pages", type=int, default=5)
    parser.add_argument("--delay", type=float, default=1.5, help="пауза между страницами, сек")
    args = parser.parse_args(argv)

    selectors = dict(item=args.item, name=args.name, price=args.price, link=args.link, next=args.next)
    try:
        if args.file:
            html = Path(args.file).read_text(encoding="utf-8")
            rows = parse_catalog(html, **{k: v for k, v in selectors.items() if k != "next"})
        else:
            rows = crawl(args.url, args.max_pages, args.delay, **selectors)
    except (PermissionError, requests.RequestException) as error:
        print(f"Ошибка: {error}", file=sys.stderr)
        return 1

    if not rows:
        print("Ничего не найдено: проверьте CSS-селекторы.", file=sys.stderr)
        return 1
    save_xlsx(rows, args.out)
    print(f"Сохранено товаров: {len(rows)} -> {args.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
