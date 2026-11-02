"""Add newly published months from a DGEG monthly sales file to the workbook.

DGEG publishes one file per year, "Vendas Mensais de Produtos de Petróleo em
Portugal" (dgeg-omn-YYYY-MM.xlsx), on:
  https://www.dgeg.gov.pt/pt/estatistica/energia/petroleo-e-derivados/vendas-mensais/

Usage:
  python code/update_dgeg.py --download            # fetch latest file into data/raw/dgeg/
  python code/update_dgeg.py path/to/dgeg-omn.xlsx # use a file already downloaded

Writes every published month of that year into its sheet of
data/processed/dataset_fuel_sales.xlsx (revised months are overwritten and
reported). Then run code/build.py.
"""
from pathlib import Path
import re
import sys
import time
import urllib.request

import openpyxl
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
WORKBOOK = ROOT / "data" / "processed" / "dataset_fuel_sales.xlsx"
RAW = ROOT / "data" / "raw" / "dgeg"
PAGE = "https://www.dgeg.gov.pt/pt/estatistica/energia/petroleo-e-derivados/vendas-mensais/"

MONTHS = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho",
          "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]

# DGEG row label -> workbook column. Rows under "Mercado de Bancas Marítimas"
# reuse generic names, so they are mapped separately.
COLUMNS = {
    "Butano": "Butano",
    "Propano": "Propano",
    "Gás Auto": "Gás Auto",
    "Propileno": "Propileno",
    "Gasolina s/ Chumbo 98": "Gasolina s/Chumbo 98",
    "Gasolina s/ Chumbo 95": "Gasolina s/Chumbo 95",
    "Nafta Química e Aromáticos": "Nafta Química e Aromáticos",
    "Petróleos (Iluminante + carburante)": "Petróleos (iluminante + carburante)",
    "Gasóleo Rodoviário": "Gasóleo rodoviário",
    "Gasóleo colorido e marcado destinado a aquecimento": "Gasóleo aquecimento",
    "Gasóleo colorido e marcado, exceto o destinado a aquecimento": "Gasóleo colorido e marcado",
    "Fuelóleo nº3 BTE": "Fuelóleo nº3 BTE",
    "Fuelóleo nº4 BTE": "Fuelóleo nº4 BTE",
    "Coque de petróleo": "Coque de petróleo",
    "Lubrificantes": "Lubrificantes",
    "Asfaltos": "Asfaltos",
    "Parafinas": "Parafinas",
    "Solventes": "Solventes",
    "Fuelóleo para a produção de energia elétrica": "Fuelóleo para a produção de energia elétrica",
    "Fuelóleo para os restantes setores": "Fuelóleo para os restantes setores",
    "Mercado de Bancas Marítimas": "Mercado de Bancas Marítimas",
    "Mercado de Aviação": "Mercado de Aviação",
    "Gasolina de aviação": "Gasolina de aviação",
    "Jet": "Jet",
    "Biocombustíveis": "Biocombustíveis",
    "Biodiesel incorporado no gasóleo": "Biodiesel incorporado no gasóleo",
    "Biodiesel fornecido diretamente ao mercado": "Biodiesel fornecido diretamente ao mercado",
    "Biogasolina incorporada": "Biogasolina incorporada",
}
BUNKERS = {
    "Gasóleo": "Gasóleo marítimo",
    "Gasóleo colorido e marcado": "Gasóleo colorido e marcado marítimo",
    "Fuelóleo Marítimo": "Fuelóleo nº4 BTE marítimo",
    "Lubrificantes": "Lubrificantes marítimo",
}


def download():
    """Find the newest dgeg-omn file on the DGEG page and save it to data/raw/dgeg/."""
    for attempt in range(3):  # the page intermittently returns a server error
        try:
            html = urllib.request.urlopen(PAGE, timeout=30).read().decode("utf-8", "replace")
            break
        except OSError:
            if attempt == 2:
                raise
            time.sleep(5)
    links = re.findall(r'href="(/media/[^"]+/dgeg-omn-(\d{4})-(\d{2})\.xlsx)"', html)
    if not links:
        sys.exit("no dgeg-omn-YYYY-MM.xlsx link found on " + PAGE)
    path, _, _ = max(links, key=lambda m: (m[1], m[2]))
    dest = RAW / path.rsplit("/", 1)[1]
    RAW.mkdir(parents=True, exist_ok=True)
    urllib.request.urlretrieve("https://www.dgeg.gov.pt" + path, dest)
    print("downloaded", dest.relative_to(ROOT))
    return dest


def read_dgeg(path):
    """Return (year, {month: {column: value}}) for the published months."""
    raw = pd.read_excel(path, header=None)
    labels = raw[0].astype(str).str.strip()
    year = int(re.search(r"(\d{4})", labels[labels.str.startswith("Vendas Mensais")].iloc[0]).group(1))
    header = labels[labels.eq("Mercado Interno")].index[0]
    assert [str(raw.iat[header, m]).strip() for m in range(1, 13)] == MONTHS, "unexpected month header"

    values, section = {}, "interno"
    for r in range(header + 1, len(raw)):
        label = re.sub(r"\s*\([a-z]\)$", "", labels[r])
        if label.startswith("Mercado de Bancas"):
            section = "bancas"
        elif label.startswith(("Mercado de Avia", "Biocombust", "Memo Fuel")):
            section = "outros"
        if label in ("nan", "", "Memo Fuel") or label.startswith("Do qual"):
            continue
        if label.startswith("("):
            break  # footnotes
        column = BUNKERS.get(label) if section == "bancas" and label in BUNKERS else COLUMNS.get(label)
        if column is None:
            sys.exit(f"unmapped DGEG row {r}: {label!r} - add it to COLUMNS")
        values[column] = [raw.iat[r, m] for m in range(1, 13)]

    assert len(values) == 32, f"expected 32 products, got {len(values)}"
    months = {}
    for m in range(12):
        col = {c: float(v[m]) for c, v in values.items() if pd.notna(v[m])}
        if any(col.values()):  # unpublished months are all zero
            months[m + 1] = col
    return year, months


def main():
    if sys.argv[1:] == ["--download"]:
        path = download()
    elif len(sys.argv) == 2:
        path = Path(sys.argv[1])
    else:
        sys.exit(__doc__)

    year, months = read_dgeg(path)
    wb = openpyxl.load_workbook(WORKBOOK)
    if str(year) not in wb.sheetnames:
        prev = wb[wb.sheetnames[-1]]
        ws = wb.create_sheet(str(year))
        for c in range(1, prev.max_column + 1):
            ws.cell(1, c, prev.cell(1, c).value)
        for i, name in enumerate(MONTHS):
            ws.cell(i + 2, 1, name)
        print("created sheet", year)
    ws = wb[str(year)]
    cols = {str(ws.cell(1, c).value).strip(): c for c in range(2, ws.max_column + 1) if ws.cell(1, c).value}
    rows = {str(ws.cell(r, 1).value).strip(): r for r in range(2, ws.max_row + 1) if ws.cell(r, 1).value}
    missing = set(months[min(months)]) - set(cols)
    if missing:
        sys.exit(f"columns missing from sheet {year}: {sorted(missing)}")

    added, revised = [], []
    for m, col in months.items():
        r = rows[MONTHS[m - 1]]
        was_empty = all(ws.cell(r, c).value in (None, "") for c in cols.values())
        for name, v in col.items():
            cell = ws.cell(r, cols[name])
            old = cell.value
            if isinstance(old, (int, float)) and abs(old - v) < 1e-9:
                continue
            cell.value = int(v) if v.is_integer() else v
            if not was_empty:
                revised.append(f"{MONTHS[m - 1]} {name}: {old} -> {v:g}")
        if was_empty:
            added.append(MONTHS[m - 1])

    print(f"{path.name}: {year}, published through {MONTHS[max(months) - 1]}")
    print("added months:", ", ".join(added) or "none")
    print(f"revised values: {len(revised)}")
    for line in revised:
        print("  ", line)
    if added or revised:
        # openpyxl drops Excel's stored formula results; have Excel recompute on open
        wb.calculation.fullCalcOnLoad = True
        wb.save(WORKBOOK)
        print("saved", WORKBOOK.relative_to(ROOT), "- now run: python code/build.py")


if __name__ == "__main__":
    main()
