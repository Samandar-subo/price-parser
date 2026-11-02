import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from openpyxl import load_workbook  # noqa: E402

import catalog_parser  # noqa: E402

SAMPLE = Path(__file__).resolve().parent.parent / "sample" / "catalog.html"


def test_parse_price_formats():
    assert catalog_parser.parse_price("12 500 ₸") == 12500
    assert catalog_parser.parse_price("от 18 900 ₸") == 18900
    assert catalog_parser.parse_price("21 300 ₸") == 21300
    assert catalog_parser.parse_price("9 750,50 ₸") == 9750.5
    assert catalog_parser.parse_price("по запросу") is None
    assert catalog_parser.parse_price("") is None


def test_parse_catalog_skips_ads_and_resolves_links():
    rows = catalog_parser.parse_catalog(SAMPLE.read_text(encoding="utf-8"), base_url="https://shop.test/catalog")
    assert len(rows) == 5  # реклама без названия пропущена
    assert rows[0] == {
        "name": "Фасад МДФ «Филёнка» 600×400",
        "price": 12500,
        "link": "https://shop.test/item/101",
    }
    assert rows[3]["price"] is None


def test_next_page_url():
    html = SAMPLE.read_text(encoding="utf-8")
    assert catalog_parser.find_next_url(html, "https://shop.test/catalog", "a.next") == "https://shop.test/catalog?page=2"
    assert catalog_parser.find_next_url(html, "https://shop.test/catalog", None) is None


def test_cli_writes_excel(tmp_path):
    out = tmp_path / "catalog.xlsx"
    code = catalog_parser.main(["--file", str(SAMPLE), "--out", str(out)])
    assert code == 0
    sheet = load_workbook(out).active
    assert sheet.max_row == 6
    assert sheet["B2"].value == 12500
    assert sheet["A1"].value == "Название"


def test_cli_reports_wrong_selectors(tmp_path):
    code = catalog_parser.main(["--file", str(SAMPLE), "--out", str(tmp_path / "x.xlsx"), "--item", ".nothing"])
    assert code == 1
