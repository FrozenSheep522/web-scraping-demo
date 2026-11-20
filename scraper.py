#!/usr/bin/env python3
"""
books.toscrape.com -> clean Excel.

A tiny, polite scraper: walks the public catalogue, pulls the fields a
client actually asks for, cleans them into usable types, and writes a
formatted .xlsx.

Usage:
    python scraper.py                 # all pages
    python scraper.py --pages 3       # first 3 pages only
    python scraper.py --out demo.xlsx
"""

from __future__ import annotations

import argparse
import logging
import re
import sys
import time
from pathlib import Path
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

BASE_URL = "https://books.toscrape.com/"
CATALOGUE_URL = urljoin(BASE_URL, "catalogue/page-1.html")
USER_AGENT = "books-demo-scraper/1.0 (+portfolio demo; contact: example@example.com)"
DELAY_SECONDS = 0.25

RATING_WORDS = {"One": 1, "Two": 2, "Three": 3, "Four": 4, "Five": 5}
PRICE_RE = re.compile(r"[\d.]+")


logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-7s %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("scraper")


def fetch(session: requests.Session, url: str, attempts: int = 3) -> str:
    """GET a page with small retries. Returns HTML as text."""
    last_error: Exception | None = None
    for attempt in range(1, attempts + 1):
        try:
            response = session.get(url, timeout=20)
            response.raise_for_status()
            return response.text
        except requests.RequestException as exc:  # network hiccup, 5xx, timeout
            last_error = exc
            log.warning("attempt %d/%d failed on %s (%s)", attempt, attempts, url, exc)
            time.sleep(attempt * 1.5)
    raise RuntimeError(f"could not fetch {url}") from last_error


def clean_price(raw: str) -> float | None:
    """'拢51.77' -> 51.77"""
    match = PRICE_RE.search(raw or "")
    return float(match.group()) if match else None


def clean_rating(classes: list[str]) -> int | None:
    """['star-rating', 'Three'] -> 3"""
    for name in classes:
        if name in RATING_WORDS:
            return RATING_WORDS[name]
    return None


def clean_stock(raw: str) -> bool:
    """'In stock' / 'Out of stock' -> True / False"""
    return (raw or "").strip().lower().startswith("in stock")


def parse_page(html: str, page_url: str) -> list[dict]:
    soup = BeautifulSoup(html, "html.parser")
    rows: list[dict] = []

    for card in soup.select("article.product_pod"):
        link = card.select_one("h3 a")
        price = card.select_one("p.price_color")
        rating = card.select_one("p.star-rating")
        stock = card.select_one("p.instock.availability")
        if not link or not price:
            continue

        rows.append(
            {
                "title": link.get("title", "").strip(),
                "price_gbp": clean_price(price.get_text()),
                "rating": clean_rating(rating.get("class", []) if rating else []),
                "in_stock": clean_stock(stock.get_text() if stock else ""),
                "url": urljoin(page_url, link.get("href", "")),
            }
        )
    return rows


def next_page_url(html: str, page_url: str) -> str | None:
    soup = BeautifulSoup(html, "html.parser")
    link = soup.select_one("li.next a")
    return urljoin(page_url, link["href"]) if link else None


def scrape(max_pages: int | None) -> list[dict]:
    session = requests.Session()
    session.headers["User-Agent"] = USER_AGENT

    rows: list[dict] = []
    url, page = CATALOGUE_URL, 1

    while url and (max_pages is None or page <= max_pages):
        log.info("page %d  %s", page, url)
        html = fetch(session, url)
        page_rows = parse_page(html, url)
        rows.extend(page_rows)
        log.info("  %d books (total %d)", len(page_rows), len(rows))

        url = next_page_url(html, url)
        page += 1
        time.sleep(DELAY_SECONDS)

    return rows


def write_excel(rows: list[dict], out_path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Books"

    headers = ["Title", "Price (GBP)", "Rating", "In stock", "URL"]
    keys = ["title", "price_gbp", "rating", "in_stock", "url"]
    widths = [50, 12, 9, 10, 62]

    ws.append(headers)
    header_fill = PatternFill("solid", fgColor="1F3864")
    for col, width in enumerate(widths, start=1):
        cell = ws.cell(row=1, column=col)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions[get_column_letter(col)].width = width

    for row in rows:
        ws.append([row[k] for k in keys])

    for r in range(2, ws.max_row + 1):
        ws.cell(row=r, column=2).number_format = "0.00"
        link_cell = ws.cell(row=r, column=5)
        if link_cell.value:
            link_cell.hyperlink = link_cell.value
            link_cell.font = Font(color="0563C1", underline="single")

    ws.row_dimensions[1].height = 26
    for r in range(2, ws.max_row + 1):
        ws.row_dimensions[r].height = 18
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:E{ws.max_row}"
    wb.save(out_path)


def main() -> int:
    parser = argparse.ArgumentParser(description="Scrape books.toscrape.com into Excel.")
    parser.add_argument("--pages", type=int, default=None, help="how many catalogue pages (default: all)")
    parser.add_argument("--out", type=Path, default=Path("books.xlsx"), help="output .xlsx file")
    args = parser.parse_args()

    started = time.time()
    rows = scrape(args.pages)
    if not rows:
        log.error("no rows scraped - the page layout may have changed")
        return 1

    write_excel(rows, args.out)
    log.info("done: %d books -> %s (%.1fs)", len(rows), args.out, time.time() - started)
    return 0


if __name__ == "__main__":
    sys.exit(main())


