"""Build the output datasets from data/processed/dataset_fuel_sales.xlsx.

The workbook has one sheet per year: month names in column A, one column per
product. Some aggregate columns are Excel formulas (e.g. =SUM(X2:Z2)); they are
evaluated here, so the build never depends on values cached by Excel.

Writes to output/:
  oil_products_panel.csv              long: one row per product x month
  oil_products_wide.csv / .xlsx       one column per product
  oil_products_timeseries.csv / .xlsx same, gap-free monthly index

Run:  python code/build.py [--also-to DIR ...]
      --also-to copies the output files and the source workbook to other folders.
"""
from pathlib import Path
import argparse
import re
import shutil

import numpy as np
import openpyxl
import pandas as pd
from openpyxl.utils import range_boundaries

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "data" / "processed" / "dataset_fuel_sales.xlsx"
OUT = ROOT / "output"

MONTHS = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho",
          "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
TERM = re.compile(r"([+-]?)(SUM\(([A-Z]+\d+:[A-Z]+\d+)\)|[A-Z]+\d+)")


def to_number(v):
    """Cell value -> float; blanks, '.', and footnote markers become NaN."""
    if v is None:
        return np.nan
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace("*", "").replace(" ", "").strip()
    try:
        return float(s) if s not in ("", ".") else np.nan
    except ValueError:
        return np.nan


def cell_value(ws, coord):
    v = ws[coord].value
    if isinstance(v, str) and v.startswith("="):
        return evaluate(ws, v)
    return to_number(v)


def evaluate(ws, formula):
    """Evaluate sums and differences of SUM(range) and single-cell terms."""
    body = formula[1:].replace(" ", "")
    terms = TERM.findall(body)
    if "".join(sign + term for sign, term, _ in terms) != body:
        raise ValueError(f"unsupported formula in sheet {ws.title}: {formula}")
    total = 0.0
    for sign, term, rng in terms:
        if rng:
            c1, r1, c2, r2 = range_boundaries(rng)
            vals = [cell_value(ws, ws.cell(r, c).coordinate)
                    for r in range(r1, r2 + 1) for c in range(c1, c2 + 1)]
            x = float(np.nansum(vals))
        else:
            x = cell_value(ws, term)
            x = 0.0 if np.isnan(x) else x
        total += -x if sign == "-" else x
    return total


def read_panel():
    """Long table: year, month, product, consumption (NaN where not reported)."""
    wb = openpyxl.load_workbook(SRC)
    rows = []
    for ws in wb.worksheets:
        year = int(ws.title)
        products = [(c, str(ws.cell(1, c).value).strip())
                    for c in range(2, ws.max_column + 1) if ws.cell(1, c).value]
        for r in range(2, ws.max_row + 1):
            name = str(ws.cell(r, 1).value or "").strip()
            if name not in MONTHS:
                continue  # TOTAL row, footnotes, blanks
            for c, product in products:
                rows.append((year, MONTHS.index(name) + 1, product,
                             cell_value(ws, ws.cell(r, c).coordinate)))
    p = pd.DataFrame(rows, columns=["year", "month", "product", "consumption"])
    p = p.drop_duplicates(["product", "year", "month"], keep="first")
    # drop months not yet published (every product blank)
    published = p.groupby(["year", "month"])["consumption"].transform(lambda s: s.notna().any())
    return p[published].sort_values(["product", "year", "month"], kind="stable").reset_index(drop=True)


OUTPUTS = ["oil_products_panel.csv", "oil_products_wide.csv", "oil_products_wide.xlsx",
           "oil_products_timeseries.csv", "oil_products_timeseries.xlsx"]


def main():
    parser = argparse.ArgumentParser(description="Build output/ from the workbook.")
    parser.add_argument("--also-to", nargs="+", default=[], metavar="DIR",
                        help="also copy the output files and the source workbook to these folders")
    args = parser.parse_args()

    p = read_panel()
    p["period"] = p["year"].astype(str) + "m" + p["month"].astype(str)
    p["date_monthly"] = pd.to_datetime(dict(year=p["year"], month=p["month"], day=1))
    ids = {name: i + 1 for i, name in enumerate(sorted(p["product"].unique()))}
    p["product_id"] = p["product"].map(ids)

    OUT.mkdir(exist_ok=True)
    panel = p.assign(date_monthly=p["date_monthly"].dt.strftime("%Y-%m-01"))
    panel[["period", "date_monthly", "year", "month", "product_id", "product", "consumption"]].to_csv(
        OUT / "oil_products_panel.csv", index=False, encoding="utf-8-sig")

    wide = p.pivot(index="date_monthly", columns="product", values="consumption")
    wide = wide.reindex(pd.date_range(wide.index.min(), wide.index.max(), freq="MS"))
    wide.index.name = "date_monthly"
    period = [f"{d.year}m{d.month}" for d in wide.index]

    w = wide.reset_index()
    w.insert(0, "period", period)
    w.to_excel(OUT / "oil_products_wide.xlsx", index=False)
    w.assign(date_monthly=w["date_monthly"].dt.strftime("%Y-%m-01")).to_csv(
        OUT / "oil_products_wide.csv", index=False, encoding="utf-8-sig")

    ts = wide.reset_index(drop=True)
    ts.insert(0, "period", period)
    ts.to_excel(OUT / "oil_products_timeseries.xlsx", index=False)
    ts.to_csv(OUT / "oil_products_timeseries.csv", index=False, encoding="utf-8-sig")

    print(f"panel: {len(panel)} rows, {len(ids)} products | "
          f"wide/time series: {period[0]} -> {period[-1]} ({len(period)} months)")

    for folder in map(Path, args.also_to):
        for name in OUTPUTS:
            shutil.copy2(OUT / name, folder / name)
        shutil.copy2(SRC, folder / SRC.name)
        print("copied outputs and workbook to", folder)


if __name__ == "__main__":
    main()
