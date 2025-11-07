import pandas as pd
import re
import json
import os

# === Paths ===
SOURCE_FILE = "./read-from/products.xlsx"
MATRIXIFY_TEMPLATE = "./read-from/test_products_excel_matrixify.xlsx"
OUTPUT_FILE = "./write-to/matrixify_ready.xlsx"

# === Utility functions ===
def clean_html(text: str) -> str:
    """Remove HTML tags and non-ASCII characters from text."""
    if pd.isna(text):
        return ""
    s = str(text)
    s = re.sub(r"<(?!br\s*\/?)[^>]+>", "", s)
    s = re.sub(r"[^\x00-\x7F\n\s\.,!?;:\(\)\-/<>\[\]'\"@#%&*+=]", "", s)
    s = re.sub(r"\s+", " ", s)
    return s.strip()

def generate_handle(title: str, variant_suffix: str = "") -> str:
    """Generate a clean Shopify handle from the product title and variant name."""
    if not title or pd.isna(title):
        return ""

    # Convert to strings (to handle numeric input)
    title = str(title)
    variant_suffix = "" if pd.isna(variant_suffix) else str(variant_suffix)

    # Base handle
    h = title.lower().strip()
    h = re.sub(r"\s+", "-", h)
    h = re.sub(r"[^a-z0-9\-]", "", h)
    h = re.sub(r"-+", "-", h).strip("-")

    # Add variant suffix if exists
    if variant_suffix:
        vs = variant_suffix.lower().strip().replace(" ", "-")
        vs = re.sub(r"[^a-z0-9\-]", "", vs)
        vs = re.sub(r"-+", "-", vs).strip("-")
        h += "-" + vs

    return h

def to_rich_text(value: str) -> str:
    """Convert plain text to Shopify-compatible rich text JSON."""
    value = "" if pd.isna(value) else str(value)
    value = clean_html(value)
    payload = {
        "type": "root",
        "children": [
            {"type": "paragraph", "children": [{"type": "text", "value": value}]}
        ],
    }
    return json.dumps(payload, ensure_ascii=False)

def price_or(value):
    """Return None if the price cell is empty, otherwise return its value."""
    return None if (pd.isna(value) or str(value).strip() == "") else value


# === File validation ===
if not os.path.exists(SOURCE_FILE):
    raise SystemExit(f"❌ SOURCE_FILE not found: {SOURCE_FILE}")
if not os.path.exists(MATRIXIFY_TEMPLATE):
    raise SystemExit(f"❌ MATRIXIFY_TEMPLATE not found: {MATRIXIFY_TEMPLATE}")

src = pd.read_excel(SOURCE_FILE, engine="openpyxl")
template = pd.read_excel(MATRIXIFY_TEMPLATE, engine="openpyxl")
matrix_cols = list(template.columns)

variant_invalid_pattern = re.compile(r"\bvar(?:iant)?s?\.?\b", re.IGNORECASE)

# Check required columns
need_src_cols = [
    "title_us (NEW)", "SKU", "Variant",
    "normalPrice (GBP)", "discountPrice (GBP)",
    "description_us (NEW)", "shortDescription_us (NEW)",
    "closingSummaryTitle_us", "closingSummaryMainText_us",
]
for c in need_src_cols:
    if c not in src.columns:
        raise SystemExit(f"❌ Missing required column in source file: '{c}'")

# === Add required Matrixify columns ===
extra_cols = [
    "Variant Inventory Tracker", "Variant Inventory Qty", "Variant Inventory Policy",
    "Variant Fulfillment Service",
    "Inventory Available: Shop location", "Inventory Available Adjust: Shop location",
    "Inventory On Hand: Shop location", "Inventory On Hand Adjust: Shop location",
    "Inventory Committed: Shop location", "Inventory Reserved: Shop location",
    "Inventory Damaged: Shop location", "Inventory Damaged Adjust: Shop location",
    "Inventory Safety Stock: Shop location", "Inventory Safety Stock Adjust: Shop location",
    "Inventory Quality Control: Shop location", "Inventory Quality Control Adjust: Shop location",
    "Inventory Incoming: Shop location",
]
for col in extra_cols:
    if col not in matrix_cols:
        matrix_cols.append(col)

out = pd.DataFrame(columns=matrix_cols)

# === Product row builder ===
def add_product_row(title, sku, normal_price, discount_price, body_html, short_desc, closing_title, closing_body, variant_suffix="", is_variant=False):
    """Build one complete product row for Matrixify import."""
    row = {col: "" for col in matrix_cols}

    # Title logic
    if is_variant and variant_suffix:
        row["Title"] = clean_html(f"{title} - {variant_suffix}")
    else:
        row["Title"] = clean_html(title)

    # Handle + SKU
    row["Handle"] = generate_handle(title, variant_suffix if is_variant else "")
    row["Variant SKU"] = "" if pd.isna(sku) else str(sku).strip()

    # Pricing
    discount = price_or(discount_price)
    normal = price_or(normal_price)
    row["Variant Price"] = discount if discount else (normal if normal else "")
    row["Variant Compare At Price"] = normal if normal else ""

    # Core data
    row["Body HTML"] = clean_html(body_html)
    row["Variant Requires Shipping"] = True
    row["Variant Taxable"] = True
    row["Option1 Name"] = "Title"
    row["Option1 Value"] = "Default Title"
    row["Status"] = "active"

    # Metafields
    row["Metafield: custom.short_description [rich_text_field]"] = to_rich_text(short_desc)
    row["Metafield: custom.closing_summary_title [rich_text_field]"] = to_rich_text(closing_title)
    row["Metafield: custom.closing_summary_body [rich_text_field]"] = to_rich_text(closing_body)

    # Inventory setup
    row["Variant Inventory Tracker"] = "shopify"
    row["Variant Inventory Qty"] = 0
    row["Variant Inventory Policy"] = "deny"
    row["Variant Fulfillment Service"] = "manual"

    # Fake warehouse (Shopify requires at least one location)
    for col in extra_cols[4:]:
        row[col] = 0

    return row

# === Main processing loop ===
added_single = 0
added_groups = 0
skipped = 0
rows_to_append = []
seen_skus = set()

i = len(src) - 1
while i >= 0:
    r = src.iloc[i]
    title = r.get("title_us (NEW)")
    variant_text = str(r.get("Variant")).strip() if not pd.isna(r.get("Variant")) else ""
    sku = str(r.get("SKU")).strip() if not pd.isna(r.get("SKU")) else ""

    # Skip rows with no title
    if pd.isna(title) or str(title).strip() == "":
        i -= 1
        continue

    # === Single product (has SKU) ===
    if sku:
        row = add_product_row(
            title, sku,
            r.get("normalPrice (GBP)"), r.get("discountPrice (GBP)"),
            r.get("description_us (NEW)"), r.get("shortDescription_us (NEW)"),
            r.get("closingSummaryTitle_us"), r.get("closingSummaryMainText_us")
        )
        sku_out = row.get("Variant SKU", "")
        if sku_out and sku_out in seen_skus:
            print(f"⚠️  Skipped duplicate SKU: {sku_out}")
        else:
            rows_to_append.append(row)
            if sku_out:
                seen_skus.add(sku_out)
            print(f"🟩 Added product: {title}")
            added_single += 1
        i -= 1
        continue

    # === Parent product with variants ===
    if variant_invalid_pattern.search(variant_text):
        parent = {
            "title": str(title).strip(),
            "body": r.get("description_us (NEW)"),
            "short": r.get("shortDescription_us (NEW)"),
            "closing_title": r.get("closingSummaryTitle_us"),
            "closing_body": r.get("closingSummaryMainText_us"),
        }

        valid_variants = []
        j = i - 1
        while j >= 0:
            rr = src.iloc[j]
            if pd.notna(rr.get("title_us (NEW)")) and str(rr.get("title_us (NEW)")).strip() != "":
                break
            variant_val = rr.get("Variant")
            variant_sku = rr.get("SKU")

            # Skip invalid rows
            if pd.isna(variant_sku) or str(variant_sku).strip() == "":
                j -= 1
                continue
            if pd.isna(variant_val) or str(variant_val).strip() == "":
                j -= 1
                continue
            if variant_invalid_pattern.search(str(variant_val)):
                j -= 1
                continue

            valid_variants.append(rr)
            j -= 1

        if not valid_variants:
            skipped += 1
            i -= 1
            continue

        print(f"🟩 Group parent (no SKU) -> only variants: {parent['title']}")

        # Add variants
        for v in reversed(valid_variants):
            variant_title = f"{parent['title']} - {v.get('Variant')}"
            row = add_product_row(
                parent["title"],
                v.get("SKU"),
                v.get("normalPrice (GBP)"),
                v.get("discountPrice (GBP)"),
                parent["body"], parent["short"],
                parent["closing_title"], parent["closing_body"],
                variant_suffix=v.get("Variant"),
                is_variant=True
            )
            sku_out = row.get("Variant SKU", "")
            if sku_out and sku_out in seen_skus:
                print(f"⚠️  Skipped duplicate SKU: {sku_out}")
            else:
                rows_to_append.append(row)
                if sku_out:
                    seen_skus.add(sku_out)
                print(f"   🟦 Added variant: {variant_title}")

        added_groups += 1

    i -= 1

# === Save output ===
out = pd.concat([out, pd.DataFrame(rows_to_append)], ignore_index=True)
os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
out.to_excel(OUTPUT_FILE, index=False)

print("\n✅ Done:", OUTPUT_FILE)
print(f"  ➕ Single products: {added_single}")
print(f"  ➕ Product groups: {added_groups}")
print(f"  🚫 Skipped (no variants): {skipped}")
