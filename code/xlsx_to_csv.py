"""Convert output/oil_products_wide.xlsx -> output/oil_products_wide.csv

Produces a CSV in the exact wide format of the Excel file:
  - first column date_monthly (ISO YYYY-MM-01)
  - one column per product, using the full Portuguese product labels as headers
  - UTF-8 with BOM, so Excel opens the accented names correctly

Run from anywhere:  python code/xlsx_to_csv.py
"""
from pathlib import Path
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent   # project root (contains /code, /output)
SRC = ROOT / "output" / "oil_products_wide.xlsx"
DST = ROOT / "output" / "oil_products_wide.csv"

df = pd.read_excel(SRC)
df["date_monthly"] = pd.to_datetime(df["date_monthly"]).dt.strftime("%Y-%m-01")
df.to_csv(DST, index=False, encoding="utf-8-sig")
print(f"wrote {DST}  ({df.shape[0]} rows x {df.shape[1]} cols)")
