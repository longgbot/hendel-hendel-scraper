import argparse
import csv
import html
import json
import os
import re
import sys
import time
from datetime import datetime
from pathlib import Path
from urllib.parse import urlparse

import requests
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

BASE_URL = "https://hendelandhendel.co.uk"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    "Accept": "application/json,text/html,*/*",
    "Accept-Language": "en-GB,en;q=0.9",
}
DELAY = 0.8
VAT_RATE = 20.0
PAGE_SIZE = 250
MAX_PAGES = 200


def exc_vat(price):
    try:
        return round(float(price) / (1 + VAT_RATE / 100), 2)
    except (TypeError, ValueError):
        return ""


def log(msg):
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}", flush=True)


def safe_name(text, max_len=80):
    text = html.unescape(text or "").replace("&", "and")
    text = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "", text)
    text = re.sub(r"\s+", " ", text).strip().rstrip(".")
    return text[:max_len] or "untitled"


def html_to_text(raw):
    if not raw:
        return ""
    t = re.sub(r"(?i)<br\s*/?>", "\n", raw)
    t = re.sub(r"(?i)</(p|div|li|h[1-6]|tr)>", "\n", t)
    t = re.sub(r"(?i)<li[^>]*>", "- ", t)
    t = re.sub(r"<[^>]+>", "", t)
    t = html.unescape(t)
    t = re.sub(r"[ \t]+", " ", t)
    t = re.sub(r"\n\s*\n+", "\n\n", t)
    return t.strip()


def extract_sections(text):
    sections = {}
    current = "Description"
    buffer = []
    for line in text.splitlines():
        stripped = line.strip()
        if re.fullmatch(r"[A-Za-z &]{3,40}:", stripped):
            if buffer:
                sections[current] = "\n".join(buffer).strip()
            current = stripped.rstrip(":")
            buffer = []
        else:
            buffer.append(line)
    if buffer:
        sections[current] = "\n".join(buffer).strip()
    return sections


def find_dimensions(desc_text):
    m = re.search(r"Dimensions?:\s*([^\n]+)", desc_text, re.I)
    return m.group(1).strip() if m else ""


def find_material(desc_text):
    m = re.search(r"Crafted from (?:high-quality )?([A-Za-z ]+?)(?:,|\.|\n)", desc_text, re.I)
    return m.group(1).strip() if m else ""


def find_finishes(desc_text):
    m = re.search(r"Available in (?:multiple finishes:|the following finishes:)?\s*([^\n]+)", desc_text, re.I)
    return m.group(1).strip().rstrip(".") if m else ""


def collection_name(title):
    m = re.match(r"^\s*([A-Z][A-Za-z'’]+)\b", title or "")
    return m.group(1) if m else ""


def image_filename(src, position):
    path = urlparse(src).path
    name = os.path.basename(path)
    if not name or "." not in name:
        name = f"image_{position}.jpg"
    return name


def full_size_url(src):
    url = src.split("?")[0]
    if url.startswith("//"):
        url = "https:" + url
    return url


def fetch_all_products(session, base_url, stop_after=0):
    products = []
    seen = set()
    for page in range(1, MAX_PAGES + 1):
        url = f"{base_url}/products.json?limit={PAGE_SIZE}&page={page}"
        log(f"Fetching product feed page {page}")
        r = session.get(url, headers=HEADERS, timeout=30)
        if r.status_code != 200:
            log(f"Feed returned HTTP {r.status_code} on page {page}, stopping")
            break
        try:
            batch = r.json().get("products", [])
        except ValueError:
            log(f"Feed page {page} was not JSON (Content-Type: {r.headers.get('Content-Type', '')}), stopping")
            break
        new = [p for p in batch if p.get("id") not in seen]
        if not new:
            if batch:
                log(f"Feed page {page} only repeated products already fetched, stopping")
            break
        seen.update(p.get("id") for p in new)
        products.extend(new)
        if len(batch) < PAGE_SIZE:
            break
        if stop_after and len(products) >= stop_after:
            log(f"Have {len(products)} products, enough for --limit {stop_after}, stopping")
            break
        time.sleep(DELAY)
    else:
        log(f"Reached the {MAX_PAGES} page cap, stopping")
    log(f"Feed complete: {len(products)} products")
    return products


def fetch_product_page(session, base_url, handle):
    url = f"{base_url}/products/{handle}"
    try:
        r = session.get(url, headers=HEADERS, timeout=30)
        if r.status_code == 200:
            return r.text
    except requests.RequestException as e:
        log(f"Page fetch failed for {handle}: {e}")
    return ""


def extract_dimensions_tab(page_html):
    if not page_html:
        return ""
    m = re.search(
        r"(?is)<(?:summary|h[2-6]|button|div)[^>]*>\s*Dimensions\s*</(?:summary|h[2-6]|button|div)>(.*?)(?:</details>|<(?:summary|h[2-6])[^>]*>|</section>)",
        page_html,
    )
    if not m:
        return ""
    text = html_to_text(m.group(1))
    return text[:2000]


def download_image(session, url, dest):
    if dest.exists() and dest.stat().st_size > 0:
        return "existing"
    try:
        r = session.get(url, headers=HEADERS, timeout=60, stream=True)
        if r.status_code != 200:
            return f"http {r.status_code}"
        with open(dest, "wb") as f:
            for chunk in r.iter_content(65536):
                f.write(chunk)
        time.sleep(DELAY / 2)
        return "downloaded"
    except requests.RequestException as e:
        return f"error {e}"


def price_range(variants):
    prices = []
    for v in variants:
        try:
            prices.append(float(v.get("price") or 0))
        except (TypeError, ValueError):
            pass
    if not prices:
        return "", ""
    return f"{min(prices):.2f}", f"{max(prices):.2f}"


def process_product(session, base_url, product, out_root, fetch_pages, download_images):
    title = html.unescape(product.get("title", ""))
    handle = product.get("handle", "")
    collection = collection_name(title)
    folder_name = safe_name(title)
    folder = out_root / "Products" / folder_name
    images_dir = folder / "images"
    folder.mkdir(parents=True, exist_ok=True)
    images_dir.mkdir(exist_ok=True)

    desc_html = product.get("body_html") or ""
    desc_text = html_to_text(desc_html)
    sections = extract_sections(desc_text)
    option_names = [o.get("name") for o in product.get("options", [])]
    variants = product.get("variants", [])
    images = product.get("images", [])
    url = f"{base_url}/products/{handle}"

    page_html = fetch_product_page(session, base_url, handle) if fetch_pages else ""
    dimensions_tab = extract_dimensions_tab(page_html)
    if page_html:
        (folder / "page.html").write_text(page_html, encoding="utf-8")
        time.sleep(DELAY)

    variant_image = {}
    for img in images:
        for vid in img.get("variant_ids", []):
            variant_image[vid] = image_filename(img.get("src", ""), img.get("position", 0))

    image_rows = []
    for img in images:
        src = full_size_url(img.get("src", ""))
        fname = image_filename(src, img.get("position", 0))
        status = "skipped"
        if download_images:
            status = download_image(session, src, images_dir / fname)
        image_rows.append({
            "filename": fname,
            "position": img.get("position"),
            "width": img.get("width"),
            "height": img.get("height"),
            "variant_ids": img.get("variant_ids", []),
            "source_url": src,
            "status": status,
        })

    variant_rows = []
    for v in variants:
        opts = {}
        for i, name in enumerate(option_names, start=1):
            opts[name] = v.get(f"option{i}")
        variant_rows.append({
            "product_title": title,
            "product_handle": handle,
            "collection": collection,
            "sku": v.get("sku") or "",
            "variant_title": v.get("title") or "",
            **{f"opt_{safe_name(n, 30).lower().replace(' ', '_')}": val for n, val in opts.items()},
            "price_inc_vat_gbp": v.get("price") or "",
            "price_exc_vat_gbp": exc_vat(v.get("price")),
            "compare_at_inc_vat_gbp": v.get("compare_at_price") or "",
            "compare_at_exc_vat_gbp": exc_vat(v.get("compare_at_price")) if v.get("compare_at_price") else "",
            "available": v.get("available"),
            "weight_grams": v.get("grams"),
            "variant_id": v.get("id"),
            "image_file": variant_image.get(v.get("id"), ""),
        })

    low, high = price_range(variants)
    product_row = {
        "folder": folder_name,
        "title": title,
        "handle": handle,
        "collection": collection,
        "product_type": product.get("product_type") or "",
        "vendor": product.get("vendor") or "",
        "tags": ", ".join(product.get("tags", [])) if isinstance(product.get("tags"), list) else (product.get("tags") or ""),
        "options": ", ".join(option_names),
        "variant_count": len(variants),
        "skus": ", ".join(v.get("sku") for v in variants if v.get("sku")),
        "price_min_inc_vat_gbp": low,
        "price_max_inc_vat_gbp": high,
        "price_min_exc_vat_gbp": exc_vat(low) if low else "",
        "price_max_exc_vat_gbp": exc_vat(high) if high else "",
        "any_available": any(v.get("available") for v in variants),
        "material": find_material(desc_text),
        "finishes": find_finishes(desc_text),
        "dimensions": find_dimensions(desc_text) or dimensions_tab[:200],
        "image_count": len(images),
        "images_downloaded": sum(1 for r in image_rows if r["status"] in ("downloaded", "existing")),
        "description_text": desc_text,
        "url": url,
        "published_at": product.get("published_at") or "",
        "updated_at": product.get("updated_at") or "",
        "product_id": product.get("id"),
    }

    (folder / "description.txt").write_text(desc_text, encoding="utf-8")
    (folder / "description.html").write_text(desc_html, encoding="utf-8")
    specs = {
        "title": title,
        "collection": collection,
        "product_type": product_row["product_type"],
        "vendor": product_row["vendor"],
        "tags": product.get("tags", []),
        "url": url,
        "options": product.get("options", []),
        "material": product_row["material"],
        "finishes": product_row["finishes"],
        "dimensions": product_row["dimensions"],
        "dimensions_tab_text": dimensions_tab,
        "description_sections": sections,
        "variants": variant_rows,
        "images": image_rows,
    }
    (folder / "specs.json").write_text(json.dumps(specs, indent=2, ensure_ascii=False), encoding="utf-8")
    with open(folder / "variants.csv", "w", newline="", encoding="utf-8") as f:
        if variant_rows:
            keys = list(variant_rows[0].keys())
            w = csv.DictWriter(f, fieldnames=keys)
            w.writeheader()
            w.writerows(variant_rows)

    return product_row, variant_rows, image_rows, folder_name


def write_ledger(out_root, product_rows, variant_rows, image_rows):
    wb = Workbook()
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="1F3A5F")

    def add_sheet(ws, rows, title):
        ws.title = title
        if not rows:
            return
        keys = []
        for r in rows:
            for k in r.keys():
                if k not in keys:
                    keys.append(k)
        ws.append(keys)
        for c in range(1, len(keys) + 1):
            cell = ws.cell(row=1, column=c)
            cell.font = header_font
            cell.fill = header_fill
        for r in rows:
            ws.append([_cell(r.get(k)) for k in keys])
        for c, k in enumerate(keys, start=1):
            width = min(60, max(12, max(len(str(r.get(k, ""))) for r in rows[:200]) + 2))
            ws.column_dimensions[get_column_letter(c)].width = width
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

    add_sheet(wb.active, product_rows, "Products")
    add_sheet(wb.create_sheet(), variant_rows, "Variants")
    add_sheet(wb.create_sheet(), image_rows, "Images")

    summary = wb.create_sheet("Summary")
    summary.append(["Generated", datetime.now().strftime("%Y-%m-%d %H:%M")])
    summary.append(["Source", BASE_URL])
    summary.append(["VAT rate used for exc VAT columns", f"{VAT_RATE:g}%"])
    summary.append(["Products", len(product_rows)])
    summary.append(["Variants", len(variant_rows)])
    summary.append(["Images", len(image_rows)])
    summary.append(["Images downloaded", sum(1 for r in image_rows if r["status"] in ("downloaded", "existing"))])
    summary.column_dimensions["A"].width = 22
    summary.column_dimensions["B"].width = 40

    xlsx_path = out_root / "Hendel and Hendel - Master Ledger.xlsx"
    wb.save(xlsx_path)

    for name, rows in (("products", product_rows), ("variants", variant_rows), ("images", image_rows)):
        if not rows:
            continue
        keys = []
        for r in rows:
            for k in r.keys():
                if k not in keys:
                    keys.append(k)
        with open(out_root / f"ledger_{name}.csv", "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=keys)
            w.writeheader()
            for r in rows:
                w.writerow({k: _cell(r.get(k)) for k in keys})
    return xlsx_path


def _cell(v):
    if isinstance(v, (list, dict)):
        return json.dumps(v, ensure_ascii=False)
    if v is None:
        return ""
    return v


def main():
    ap = argparse.ArgumentParser(description="Pull Hendel and Hendel products into folders plus a master ledger")
    ap.add_argument("--out", required=True, help="Output root folder, e.g. your Google Drive sync folder")
    ap.add_argument("--base-url", default=BASE_URL)
    ap.add_argument("--types", default="", help="Comma separated product_type filter, case insensitive substring match")
    ap.add_argument("--include", default="", help="Comma separated words, title must contain one")
    ap.add_argument("--exclude", default="", help="Comma separated words, skip titles containing any")
    ap.add_argument("--limit", type=int, default=0, help="Only process the first N products (for testing)")
    ap.add_argument("--no-images", action="store_true", help="Skip image downloads")
    ap.add_argument("--no-pages", action="store_true", help="Skip fetching product HTML pages (faster, loses Dimensions tab)")
    ap.add_argument("--from-json", default="", help="Use a saved products.json file instead of fetching")
    ap.add_argument("--vat-rate", type=float, default=20.0, help="VAT percentage used to calculate exc VAT prices")
    args = ap.parse_args()

    global VAT_RATE
    VAT_RATE = args.vat_rate

    out_root = Path(args.out).expanduser()
    out_root.mkdir(parents=True, exist_ok=True)
    session = requests.Session()

    types = [t.strip().lower() for t in args.types.split(",") if t.strip()]
    include = [t.strip().lower() for t in args.include.split(",") if t.strip()]
    exclude = [t.strip().lower() for t in args.exclude.split(",") if t.strip()]

    if args.from_json:
        products = json.loads(Path(args.from_json).read_text(encoding="utf-8")).get("products", [])
    else:
        stop_after = args.limit if not (types or include or exclude) else 0
        products = fetch_all_products(session, args.base_url, stop_after=stop_after)
        (out_root / "raw_products.json").write_text(json.dumps({"products": products}, indent=1), encoding="utf-8")

    if not products:
        log("No products returned. Check the URL in a browser: " + args.base_url + "/products.json")
        sys.exit(1)

    selected = []
    for p in products:
        t = (p.get("title") or "").lower()
        pt = (p.get("product_type") or "").lower()
        if types and not any(x in pt for x in types):
            continue
        if include and not any(x in t for x in include):
            continue
        if exclude and any(x in t for x in exclude):
            continue
        selected.append(p)
    if args.limit:
        selected = selected[: args.limit]
    log(f"Processing {len(selected)} of {len(products)} products")

    product_rows, variant_rows, image_rows = [], [], []
    for i, p in enumerate(selected, start=1):
        log(f"[{i}/{len(selected)}] {p.get('title')}")
        try:
            pr, vr, ir, folder_name = process_product(
                session, args.base_url, p, out_root,
                fetch_pages=not args.no_pages, download_images=not args.no_images,
            )
        except Exception as e:
            log(f"  FAILED: {e}")
            continue
        product_rows.append(pr)
        variant_rows.extend(vr)
        for r in ir:
            image_rows.append({"product_folder": folder_name, **r})

    xlsx = write_ledger(out_root, product_rows, variant_rows, image_rows)
    log(f"Done. Ledger: {xlsx}")
    log(f"Products: {len(product_rows)}  Variants: {len(variant_rows)}  Images: {len(image_rows)}")


if __name__ == "__main__":
    main()
