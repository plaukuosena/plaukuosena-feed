#!/usr/bin/env python3
"""
Plaukuosena.lt -> Kaina24 XML feed generator.

What it does:
1. Reads sitemap.xml
2. Finds likely product URLs
3. Reads product pages
4. Extracts JSON-LD Product data when available
5. Falls back to OpenGraph/meta/HTML text
6. Generates:
   - kaina24.xml
   - products_snapshot.csv
   - products_snapshot.xlsx

Run:
    pip install -r requirements.txt
    python generate_feed.py
"""
from __future__ import annotations

import csv
import hashlib
import html
import json
import re
import sys
import time
import unicodedata
import xml.etree.ElementTree as ET
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Any, Iterable
from urllib.parse import urlparse

import requests
requests.packages.urllib3.disable_warnings()

from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

ROOT = Path(__file__).resolve().parent
CONFIG_PATH = ROOT / "config.json"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
    "Accept-Language": "lt-LT,lt;q=0.9,en-US;q=0.8,en;q=0.7",
    "Cache-Control": "no-cache",
    "Pragma": "no-cache",
}


@dataclass
class Product:
    sku: str
    title: str
    brand: str
    price: str
    old_price: str
    currency: str
    url: str
    image: str
    category: str
    availability: str
    size: str
    description: str
    source_status: str


def load_config() -> dict[str, Any]:
    with CONFIG_PATH.open("r", encoding="utf-8") as f:
        return json.load(f)


def fetch_text(url: str, timeout: int = 25) -> str:
    r = requests.get(
        url,
        headers=HEADERS,
        timeout=timeout,
        verify=False
    )

    r.raise_for_status()
    r.encoding = "utf-8"

    return r.text


def normalize_text(value: str | None) -> str:
    if not value:
        return ""
    value = html.unescape(value)
    value = re.sub(r"\s+", " ", value).strip()
    return value


def strip_accents(value: str) -> str:
    return "".join(
        ch for ch in unicodedata.normalize("NFKD", value)
        if not unicodedata.combining(ch)
    )


def codeify(value: str, max_len: int = 12) -> str:
    value = strip_accents(value).upper()
    value = re.sub(r"[^A-Z0-9]+", "-", value).strip("-")
    value = re.sub(r"-+", "-", value)
    return value[:max_len].strip("-") or "PROD"


def sitemap_urls(xml_text: str) -> list[str]:
    root = ET.fromstring(xml_text.encode("utf-8"))
    urls: list[str] = []
    for loc in root.iter():
        if loc.tag.endswith("loc") and loc.text:
            urls.append(loc.text.strip())
    return urls


def likely_product_url(url: str, cfg: dict[str, Any]) -> bool:
    filters = cfg.get("filters", {})
    low = url.lower()
    for bad in filters.get("exclude_url_contains", []):
        if bad.lower() in low:
            return False
    includes = filters.get("include_url_contains", [])
    if includes and not any(x.lower() in low for x in includes):
        return False
    # If URL looks like category only, keep it out unless product hints exist.
    hints = filters.get("product_hint_words", [])
    return any(h.lower() in low for h in hints)


def jsonld_objects(soup: BeautifulSoup) -> Iterable[Any]:
    for script in soup.find_all("script", type="application/ld+json"):
        raw = script.string or script.get_text(" ", strip=True)
        if not raw:
            continue
        try:
            data = json.loads(raw)
        except Exception:
            continue
        if isinstance(data, list):
            for item in data:
                yield item
        else:
            yield data


def iter_graph_nodes(obj: Any) -> Iterable[dict[str, Any]]:
    if isinstance(obj, dict):
        if "@graph" in obj and isinstance(obj["@graph"], list):
            for node in obj["@graph"]:
                if isinstance(node, dict):
                    yield from iter_graph_nodes(node)
        else:
            yield obj
    elif isinstance(obj, list):
        for item in obj:
            yield from iter_graph_nodes(item)


def find_product_jsonld(soup: BeautifulSoup) -> dict[str, Any] | None:
    for obj in jsonld_objects(soup):
        for node in iter_graph_nodes(obj):
            typ = node.get("@type")
            if isinstance(typ, list):
                is_product = "Product" in typ
            else:
                is_product = typ == "Product"
            if is_product:
                return node
    return None


def get_meta(soup: BeautifulSoup, *names: str) -> str:
    for name in names:
        tag = soup.find("meta", attrs={"property": name}) or soup.find("meta", attrs={"name": name})
        if tag and tag.get("content"):
            return normalize_text(tag.get("content"))
    return ""


def first_image_from_json(value: Any) -> str:
    if isinstance(value, str):
        return value
    if isinstance(value, list) and value:
        return first_image_from_json(value[0])
    if isinstance(value, dict):
        return value.get("url", "") or value.get("contentUrl", "")
    return ""


def offer_data(product_ld: dict[str, Any]) -> dict[str, Any]:
    offers = product_ld.get("offers") or {}
    if isinstance(offers, list) and offers:
        offers = offers[0]
    if not isinstance(offers, dict):
        return {}
    return offers


def infer_brand(title: str, cfg: dict[str, Any], ld_brand: Any = None) -> str:
    if isinstance(ld_brand, dict):
        name = ld_brand.get("name")
        if name:
            return normalize_text(str(name))
    if isinstance(ld_brand, str) and ld_brand:
        return normalize_text(ld_brand)

    low_title = title.lower()
    for brand in cfg.get("brand_codes", {}).keys():
        if brand.lower() in low_title:
            return brand
    # URL/name fallback for common variants
    aliases = {
        "milk_shake": "Milk Shake",
        "milk shake": "Milk Shake",
        "milkshake": "Milk Shake",
        "american crew": "American Crew",
        "wella": "Wella Professionals",
        "olaplex": "Olaplex",
        "davines": "Davines",
        "kadus": "Kadus Professionals",
        "qiqi": "Qiqi",
    }
    for key, val in aliases.items():
        if key in low_title:
            return val
    return ""


def infer_size(text: str) -> str:
    # Handles 250ml, 250 ml, 85g, 1l, 12x10ml
    m = re.search(r"(\d+(?:[,.]\d+)?)\s?(ml|g|l|kg|vnt|pcs)\b", text, flags=re.I)
    if not m:
        return ""
    number = m.group(1).replace(",", ".")
    unit = m.group(2).lower()
    return f"{number}{unit}"


def infer_type_code(title: str, cfg: dict[str, Any]) -> str:
    low = title.lower()
    for key, code in cfg.get("type_codes", {}).items():
        if key.lower() in low:
            return code
    return "PRD"


def model_code(title: str, brand: str, type_code: str) -> str:
    # Remove brand and common product words; take 1-2 meaningful tokens.
    clean = title
    if brand:
        clean = re.sub(re.escape(brand), " ", clean, flags=re.I)
    common = [
        "shampoo", "conditioner", "mask", "oil", "spray", "serum", "cream", "gel",
        "šampūnas", "sampunas", "kondicionierius", "kaukė", "kauke", "aliejus",
        "purškiklis", "purskiklis", "serumas", "kremas", "gelis",
        "plaukams", "plauku", "plaukų", "prieziuros", "priežiūros"
    ]
    for word in common:
        clean = re.sub(rf"\b{re.escape(word)}\b", " ", clean, flags=re.I)
    clean = re.sub(r"\d+(?:[,.]\d+)?\s?(ml|g|l|kg|vnt|pcs)\b", " ", clean, flags=re.I)
    tokens = [codeify(t, 8) for t in re.split(r"\s+", clean) if codeify(t, 8)]
    tokens = [t for t in tokens if t not in {type_code, "PROD", "IR", "SU", "BE", "THE"}]
    if not tokens:
        return "GEN"
    return "-".join(tokens[:2])


def generate_sku(title: str, brand: str, size: str, cfg: dict[str, Any]) -> str:
    brand_code = cfg.get("brand_codes", {}).get(brand) or codeify(brand or "PLK", 4)
    type_code = infer_type_code(title, cfg)
    model = model_code(title, brand, type_code)
    size_code = codeify(size.replace(".", ""), 8) if size else "NA"
    base = f"{brand_code}-{model}-{type_code}-{size_code}"
    return re.sub(r"-+", "-", base).strip("-")


def availability_from_offer(offer: dict[str, Any], cfg: dict[str, Any]) -> str:
    av = str(offer.get("availability", "")).lower()
    if "instock" in av or "in_stock" in av or "in stock" in av:
        return "in_stock"
    if "outofstock" in av or "out_of_stock" in av or "out of stock" in av:
        return "out_of_stock"
    if "preorder" in av:
        return "preorder"
    return cfg.get("store", {}).get("default_availability", "in_stock")


def clean_price(value: Any) -> str:
    if value is None:
        return ""
    s = str(value)
    s = s.replace("€", "").replace("EUR", "")
    s = re.sub(r"[^0-9,.]", "", s).replace(",", ".")
    # If multiple numbers accidentally joined, use first decimal-like number
    m = re.search(r"\d+(?:\.\d+)?", s)
    return m.group(0) if m else ""


def parse_product(url: str, cfg: dict[str, Any]) -> Product | None:
    try:
        page = fetch_text(url)
    except Exception as e:
        print(f"WARN fetch failed: {url} ({e})", file=sys.stderr)
        return None

    soup = BeautifulSoup(page, "lxml")
    ld = find_product_jsonld(soup) or {}
    offer = offer_data(ld)

    title = normalize_text(ld.get("name") if isinstance(ld, dict) else "")
    if not title:
        title = get_meta(soup, "og:title", "twitter:title")
    if not title and soup.title:
        title = normalize_text(soup.title.get_text(" ", strip=True))
    title = re.sub(r"\s*[|–-]\s*Plaukuosena.*$", "", title, flags=re.I).strip()

    if not title:
        return None

    brand = infer_brand(title, cfg, ld.get("brand") if isinstance(ld, dict) else None)
    desc = normalize_text(ld.get("description") if isinstance(ld, dict) else "")
    if not desc:
        desc = get_meta(soup, "og:description", "description")
    image = first_image_from_json(ld.get("image") if isinstance(ld, dict) else None)
    if not image:
        image = get_meta(soup, "og:image", "twitter:image")

    price = clean_price(offer.get("price") or get_meta(soup, "product:price:amount"))
    currency = normalize_text(offer.get("priceCurrency") or get_meta(soup, "product:price:currency") or cfg["store"].get("currency", "EUR"))
    availability = availability_from_offer(offer, cfg)
    size = infer_size(title + " " + desc)
    category = ""
    crumbs = soup.select('[itemtype*="BreadcrumbList"] [itemprop="name"], nav a, .breadcrumb a')
    if crumbs:
        category = normalize_text(crumbs[-2].get_text(" ", strip=True) if len(crumbs) > 1 else crumbs[-1].get_text(" ", strip=True))

    sku = normalize_text(ld.get("sku") if isinstance(ld, dict) else "")
    if not sku:
        sku = generate_sku(title, brand, size, cfg)

    return Product(
        sku=sku,
        title=title,
        brand=brand,
        price=price,
        old_price="",
        currency=currency,
        url=url,
        image=image,
        category=category,
        availability=availability,
        size=size,
        description=desc[:500],
        source_status="OK" if price else "CHECK_PRICE"
    )


def dedupe_skus(products: list[Product]) -> list[Product]:
    seen: dict[str, int] = {}
    for p in products:
        base = p.sku
        if base not in seen:
            seen[base] = 1
            continue
        seen[base] += 1
        suffix = hashlib.md5(p.url.encode("utf-8")).hexdigest()[:4].upper()
        p.sku = f"{base}-{suffix}"
    return products


def write_csv(products: list[Product], path: Path) -> None:
    fields = list(asdict(products[0]).keys()) if products else [
        "sku", "title", "brand", "price", "old_price", "currency", "url", "image", "category", "availability", "size", "description", "source_status"
    ]
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        for p in products:
            writer.writerow(asdict(p))


def write_xlsx(products: list[Product], path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Products"
    headers = ["sku", "title", "brand", "price", "currency", "url", "image", "category", "availability", "size", "source_status"]
    ws.append(headers)
    for p in products:
        ws.append([getattr(p, h) for h in headers])
    header_fill = PatternFill("solid", fgColor="0F766E")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")
    widths = {"A": 22, "B": 48, "C": 22, "D": 10, "E": 10, "F": 55, "G": 45, "H": 22, "I": 14, "J": 12, "K": 16}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    ws.freeze_panes = "A2"
    wb.save(path)


def xml_text_el(parent: ET.Element, name: str, value: str) -> ET.Element:
    el = ET.SubElement(parent, name)
    el.text = value or ""
    return el


def cdata(text: str) -> str:
    text = text or ""
    return f"<![CDATA[{text}]]>"


def write_kaina24_xml(products: list[Product], path: Path, cfg: dict[str, Any]) -> None:
    lines = ['<?xml version="1.0" encoding="UTF-8"?>', '<products>']

    for p in products:
        lines.append(f'  <product id="{p.sku}">')

        lines.append(f'    <title>{cdata(p.title)}</title>')
        lines.append(f'    <description>{cdata(p.description)}</description>')

        lines.append(f'    <price>{p.price}</price>')

        lines.append('    <condition>new</condition>')

        lines.append('    <stock>5</stock>')

        lines.append(f'    <manufacturer>{cdata(p.brand)}</manufacturer>')

        lines.append(f'    <image_url>{cdata(p.image)}</image_url>')

        lines.append(f'    <product_url>{cdata(p.url)}</product_url>')

        lines.append(f'    <category_name>{cdata(p.category)}</category_name>')

        lines.append('    <delivery>')
        lines.append('      <home_delivery>')
        lines.append('        <working_days><![CDATA[2]]></working_days>')
        lines.append('        <price><![CDATA[3.99]]></price>')
        lines.append('      </home_delivery>')
        lines.append('    </delivery>')

        lines.append('  </product>')

    lines.append('</products>')

    xml_content = "\n".join(lines)

    with open(path, "w", encoding="utf-8") as f:
        f.write(xml_content)

    path.write_text("\n".join(lines), encoding="utf-8")


def main() -> int:
    cfg = load_config()
    sitemap = fetch_text(cfg["store"]["sitemap_url"])
    all_urls = sitemap_urls(sitemap)
    product_urls = [u for u in all_urls if likely_product_url(u, cfg)]
    print(f"Found {len(all_urls)} sitemap URLs; {len(product_urls)} likely product URLs")

    products: list[Product] = []
    for i, url in enumerate(product_urls, 1):
        print(f"[{i}/{len(product_urls)}] {url}")
        p = parse_product(url, cfg)
        if p and p.price:
            products.append(p)
        elif p:
            products.append(p)
        time.sleep(0.25)

    products = dedupe_skus(products)
    products.sort(key=lambda x: (x.brand, x.title))

    out_cfg = cfg.get("output", {})
    write_kaina24_xml(products, ROOT / out_cfg.get("xml_file", "kaina24.xml"), cfg)
    write_csv(products, ROOT / out_cfg.get("csv_file", "products_snapshot.csv"))
    write_xlsx(products, ROOT / out_cfg.get("xlsx_file", "products_snapshot.xlsx"))

    print(f"Generated {len(products)} products")
    print(f"XML: {ROOT / out_cfg.get('xml_file', 'kaina24.xml')}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
