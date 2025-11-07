# 🧾 Matrixify Product Formatter

This Python script converts a raw Excel product export into a **Matrixify-compatible import file** for Shopify.

It automatically detects single products and grouped products with variants, cleans HTML descriptions, and builds the correct column structure for Matrixify.

---

## 📦 Features

✅ Automatically separates **single products** and **variant groups**  
✅ Cleans and normalizes text and HTML  
✅ Generates **Shopify-friendly handles**  
✅ Builds metafields (`short_description`, `closing_summary_title`, `closing_summary_body`)  
✅ Adds **all required inventory columns** for Matrixify import  
✅ Prevents duplicate SKUs  
✅ Supports both **normal** and **discounted** prices

---

## 🧰 Requirements

- Python **3.9+**
- Installed libraries:
  ```bash
  pip install pandas openpyxl
  ```

## 📁 Folder Structure

excel-products-python/
│
├── read-from/
│ ├── products.xlsx # Source data from your supplier
│ └── test_products_excel_matrixify.xlsx # Template file from Matrixify
│
├── write-to/
│ └── matrixify_ready.xlsx # Final formatted file (auto-created)
│
├── formatter.py # Main script
└── README.md

## ⚙️ How It Works

## 1️⃣ Input Files

You should place two files inside the read-from folder:

products.xlsx — your raw product data file;

test_products_excel_matrixify.xlsx — a template from Matrixify (used to keep the column structure).

## 2️⃣ Running the Script

In terminal (from project root):

```bash
python formatter.py
```

If everything is set up correctly, you’ll see output logs in the terminal, like:

```bash
🟩 Added product: Example Watch
🟩 Group parent (no SKU) -> only variants: Referee's Watch
   🟦 Added variant: Referee's Watch - Blue
   🟦 Added variant: Referee's Watch - Grey
✅ Done: ./write-to/matrixify_ready.xlsx
```

The processed file will appear in:

```bash
./write-to/matrixify_ready.xlsx
```

## 🧠 Notes & Best Practices

Handles are generated automatically in Shopify-friendly format
(lowercase, hyphens, no special characters)

If duplicate SKUs are found, they are skipped automatically

The output file is always rewritten (matrixify_ready.xlsx)

Only rows with valid title_us (NEW) will be processed

### To use need to have a plan of Matrixify (on depending of your purposes)

### 🧑‍💻 Author

Created by Serrakura
GitHub: github.com/Serrakura1
