import csv
import shutil
from datetime import datetime
from decimal import ROUND_HALF_UP, Decimal
from urllib.parse import quote

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

PARENT_SKU_PREFIX = "HH-"
UPLOAD_FOLDER = "WooCommerce upload"
OTHER = "Other"
IRONMONGERY = "Ironmongery"
MONEY = '"£"#,##0.00'

HEADER_FONT = Font(bold=True, color="FFFFFF")
HEADER_FILL = PatternFill("solid", fgColor="1F3A5F")
PARENT_FONT = Font(bold=True)
PARENT_FILL = PatternFill("solid", fgColor="DCE6F1")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")


def exc_vat_exact(price, vat_rate):
    """Exc VAT price that WooCommerce turns back into the exact inc VAT price.

    Rounding to pennies can make the site show a different inc VAT price
    (17.25 -> 14.38 -> 17.26), so keep up to 4 decimals when needed.
    """
    try:
        inc = Decimal(str(price))
    except ArithmeticError:
        return ""
    factor = 1 + Decimal(str(vat_rate)) / 100
    exc = inc / factor
    pennies = exc.quantize(Decimal("0.01"), ROUND_HALF_UP)
    if (pennies * factor).quantize(Decimal("0.01"), ROUND_HALF_UP) == inc:
        return f"{pennies:.2f}"
    return format(exc.quantize(Decimal("0.0001"), ROUND_HALF_UP).normalize(), "f")


def collection_for(title_word, collection_tags):
    return title_word if title_word and title_word in collection_tags else OTHER


def image_url(base_url, folder, filename):
    path = f"Products/{folder}/images/{filename}"
    if not base_url:
        return path
    return base_url.rstrip("/") + "/" + quote(path)


def woo_columns(max_attributes):
    cols = ["Type", "SKU", "Parent", "Name"]
    for n in range(1, max_attributes + 1):
        cols += [f"Attribute {n} name", f"Attribute {n} value(s)", f"Attribute {n} visible", f"Attribute {n} global"]
    cols += [
        "Regular price", "In stock?", "Categories", "Tags", "Images",
        "Published", "Visibility in catalog", "Tax status", "Tax class", "Short description", "Description",
    ]
    return cols


def build_product_rows(item, vat_rate, image_base_url):
    """WooCommerce import rows for one product: a variable parent plus one row per variation."""
    variants = item["variants"]
    option_names = item["option_names"]
    images = item["images"]
    folder = item["folder"]

    urls = []
    for img in images:
        u = image_url(image_base_url, folder, img["filename"])
        if u not in urls:
            urls.append(u)
    variant_url = {}
    for img in images:
        for vid in img["variant_ids"]:
            variant_url.setdefault(vid, image_url(image_base_url, folder, img["filename"]))

    tags = [item["collection"]] if item["collection"] != OTHER else []
    tags += [t for t in item["tags"] if t not in tags]
    common = {
        "Published": 1,
        "Visibility in catalog": "visible",
        "Tax status": "taxable",
    }
    info = {"Collection": item["collection"], "Source page": item["url"]}

    is_simple = len(variants) == 1 and option_names in ([], ["Title"])
    if is_simple:
        v = variants[0]
        return [{
            **info, **common,
            "Type": "simple",
            "SKU": v.get("sku") or "",
            "Name": item["title"],
            "Regular price": exc_vat_exact(v.get("price"), vat_rate),
            "Price inc VAT": v.get("price"),
            "In stock?": 1 if v.get("available") else 0,
            "Categories": item["product_type"] or OTHER,
            "Tags": ", ".join(tags),
            "Images": ", ".join(urls),
            "Description": item["description_html"],
        }]

    parent_sku = PARENT_SKU_PREFIX + item["handle"]
    parent = {
        **info, **common,
        "Type": "variable",
        "SKU": parent_sku,
        "Name": item["title"],
        "In stock?": 1 if any(v.get("available") for v in variants) else 0,
        "Categories": item["product_type"] or OTHER,
        "Tags": ", ".join(tags),
        "Images": ", ".join(urls),
        "Description": item["description_html"],
    }
    for n, name in enumerate(option_names, start=1):
        values = []
        for v in variants:
            val = v.get(f"option{n}")
            if val and val not in values:
                values.append(val)
        parent[f"Attribute {n} name"] = name
        parent[f"Attribute {n} value(s)"] = ", ".join(values)
        parent[f"Attribute {n} visible"] = 1
        parent[f"Attribute {n} global"] = 1

    rows = [parent]
    for v in variants:
        row = {
            **info, **common,
            "Type": "variation",
            "SKU": v.get("sku") or "",
            "Parent": parent_sku,
            "Regular price": exc_vat_exact(v.get("price"), vat_rate),
            "Price inc VAT": v.get("price"),
            "In stock?": 1 if v.get("available") else 0,
            "Images": variant_url.get(v.get("id"), ""),
            "Tax class": "parent",
        }
        for n, name in enumerate(option_names, start=1):
            row[f"Attribute {n} name"] = name
            row[f"Attribute {n} value(s)"] = v.get(f"option{n}") or ""
            row[f"Attribute {n} global"] = 1
        rows.append(row)
    return rows


def collection_order(items):
    present = {i["collection"] for i in items}
    names = sorted(present - {IRONMONGERY, OTHER})
    return names + [c for c in (IRONMONGERY, OTHER) if c in present]


def write_outputs(out_root, items, vat_rate, image_base_url, images_total, images_downloaded, source):
    """Write the WooCommerce upload CSVs and the master ledger workbook. Returns the workbook path."""
    max_attributes = max([len(i["option_names"]) for i in items] + [1])
    columns = woo_columns(max_attributes)
    collections = collection_order(items)

    by_collection = {c: [] for c in collections}
    for item in sorted(items, key=lambda i: i["title"].lower()):
        by_collection[item["collection"]].append(item)

    rows_by_collection = {
        c: [r for item in by_collection[c] for r in build_product_rows(item, vat_rate, image_base_url)]
        for c in collections
    }

    upload_dir = out_root / UPLOAD_FOLDER
    if upload_dir.exists():
        shutil.rmtree(upload_dir)
    upload_dir.mkdir(parents=True)

    def write_csv(path, rows):
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=columns, extrasaction="ignore")
            w.writeheader()
            for r in rows:
                w.writerow({k: r.get(k, "") for k in columns})

    upload_files = {}
    for n, c in enumerate(collections, start=1):
        name = f"{n:02d} {c}.csv"
        write_csv(upload_dir / name, rows_by_collection[c])
        upload_files[c] = name
    all_rows = [r for c in collections for r in rows_by_collection[c]]
    write_csv(upload_dir / "All products.csv", all_rows)

    wb = Workbook()
    _collections_sheet(wb.active, collections, by_collection, rows_by_collection, upload_files, vat_rate)
    _all_products_sheet(wb.create_sheet("All Products"), columns, all_rows)
    _summary_sheet(
        wb.create_sheet("Summary"), items, all_rows, collections, vat_rate, image_base_url,
        images_total, images_downloaded, source,
    )
    path = out_root / "Hendel and Hendel - Master Ledger.xlsx"
    wb.save(path)
    return path


def _style_header(ws, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 32


def _money(value):
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _collections_sheet(ws, collections, by_collection, rows_by_collection, upload_files, vat_rate):
    ws.title = "Collections"
    headers = [
        "#", "Collection", "Products", "Variants",
        "Lowest price exc VAT", "Highest price exc VAT", "Lowest price inc VAT", "Highest price inc VAT",
        "Upload file", "Uploaded on", "Notes",
    ]
    ws.append(headers)
    _style_header(ws, len(headers))
    total_products = total_variants = 0
    for n, c in enumerate(collections, start=1):
        sellable = [r for r in rows_by_collection[c] if r["Type"] in ("variation", "simple")]
        inc = [p for p in (_money(r.get("Price inc VAT")) for r in sellable) if p is not None]
        exc = [p for p in (_money(r.get("Regular price")) for r in sellable) if p is not None]
        ws.append([
            n, c, len(by_collection[c]), len(sellable),
            min(exc) if exc else None, max(exc) if exc else None,
            min(inc) if inc else None, max(inc) if inc else None,
            upload_files[c], None, None,
        ])
        total_products += len(by_collection[c])
        total_variants += len(sellable)
        row = ws.max_row
        for col in (5, 6, 7, 8):
            ws.cell(row=row, column=col).number_format = MONEY
        for col in (10, 11):
            ws.cell(row=row, column=col).fill = INPUT_FILL
        ws.cell(row=row, column=10).number_format = "dd/mm/yyyy"
    ws.append([None, "Total", total_products, total_variants, None, None, None, None, "All products.csv"])
    for cell in ws[ws.max_row]:
        cell.font = PARENT_FONT
    widths = [5, 22, 10, 10, 13, 13, 13, 13, 26, 14, 30]
    for c, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{ws.max_row - 1}"


def _all_products_sheet(ws, columns, rows):
    sheet_columns = ["Collection"] + columns + ["Source page"]
    price_at = sheet_columns.index("Regular price")
    sheet_columns.insert(price_at + 1, "Price inc VAT")
    labels = {"Regular price": "Regular price (exc VAT)", "Price inc VAT": "Price inc VAT (info only)"}
    ws.append([labels.get(c, c) for c in sheet_columns])
    _style_header(ws, len(sheet_columns))

    money_cols = {sheet_columns.index("Regular price") + 1, sheet_columns.index("Price inc VAT") + 1}
    for r in rows:
        values = []
        for c in sheet_columns:
            v = r.get(c, "")
            if c in ("Regular price", "Price inc VAT"):
                v = _money(v)
            values.append(v if v != "" else None)
        ws.append(values)
        row = ws.max_row
        for col in money_cols:
            ws.cell(row=row, column=col).number_format = MONEY
        if r["Type"] in ("variable", "simple"):
            for cell in ws[row]:
                cell.font = PARENT_FONT
                cell.fill = PARENT_FILL

    widths = {
        "Collection": 14, "Type": 10, "SKU": 22, "Parent": 26, "Name": 34,
        "Regular price": 13, "Price inc VAT": 13, "In stock?": 9, "Categories": 22, "Tags": 30,
        "Images": 40, "Description": 50, "Short description": 16, "Source page": 40,
    }
    for c, name in enumerate(sheet_columns, start=1):
        if name.endswith(" name"):
            w = 14
        elif name.endswith(" value(s)"):
            w = 26
        elif name.endswith(" visible") or name.endswith(" global"):
            w = 9
        else:
            w = widths.get(name, 12)
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = "F2"
    ws.auto_filter.ref = ws.dimensions


def _summary_sheet(ws, items, all_rows, collections, vat_rate, image_base_url, images_total, images_downloaded, source):
    sellable = sum(1 for r in all_rows if r["Type"] in ("variation", "simple"))
    image_note = image_base_url or (
        "NOT SET: the Images column holds file paths, not web addresses. "
        "Upload the Products folder to the website, then set HENDEL_IMAGE_BASE_URL and re-run before importing."
    )
    lines = [
        ("Generated", datetime.now().strftime("%d/%m/%Y %H:%M")),
        ("Source", source),
        ("Collections", len(collections)),
        ("Products", len(items)),
        ("Variants (items you can buy)", sellable),
        ("Images", images_total),
        ("Images downloaded", images_downloaded),
        ("VAT rate", f"{vat_rate:g}%"),
        ("Prices in the upload files", "Exc VAT. Kept to up to 4 decimals (e.g. 14.375) so the site shows the exact inc VAT price."),
        ("Image web address", image_note),
        ("", ""),
        ("How to upload a collection", ""),
        ("1", "Open the 'WooCommerce upload' folder and download the collection's file (see the Collections tab)."),
        ("2", "In WordPress: Products > Import > choose the file > Continue."),
        ("3", "Leave the column matching as it is and click Run the importer."),
        ("4", "Fill in 'Uploaded on' for that collection on the Collections tab."),
        ("", ""),
        ("Reading the All Products tab", "Blue bold rows are products. The rows under each one are its variants (one per finish/size)."),
    ]
    for a, b in lines:
        ws.append([a, b])
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
        row[0].font = PARENT_FONT
        row[1].alignment = Alignment(wrap_text=True, vertical="top")
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 100
