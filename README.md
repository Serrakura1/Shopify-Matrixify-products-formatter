# 🧾 Matrixify Product Formatter

A Python script that converts a raw Excel product list into a Matrixify-compatible Shopify import file.

## ⚙️ Features
- Automatically detects grouped products with variants  
- Cleans HTML from descriptions  
- Generates Shopify handles  
- Creates metafields (short description, closing summary)  
- Adds inventory columns for Matrixify import  

## 🧰 Requirements
- Python 3.9+
- pandas
- openpyxl

Install dependencies:
```bash
pip install pandas openpyxl
