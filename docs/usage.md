# Usage

All paths default to the repository layout, so run commands from the
repository root.

## Monthly update

DGEG publishes one file per year, `dgeg-omn-YYYY-MM.xlsx`, and replaces it as
each month comes out. To bring the dataset up to date:

```bash
uv run dgeg-fuelsales refresh
```

`refresh` downloads the newest file into `data/raw/dgeg/`, writes every
published month of that year into `data/processed/dataset_fuel_sales.xlsx`,
reports any values DGEG revised, and rebuilds `outputs/`. The same steps one at
a time:

```bash
uv run dgeg-fuelsales download
uv run dgeg-fuelsales update --file data/raw/dgeg/dgeg-omn-2026-09.xlsx
uv run dgeg-fuelsales build
```

`update` without `--file` downloads the newest file first. `build` and
`refresh` take `--also-to DIR` (or `--also-to '[DIR1,DIR2]'`) to copy the
outputs and the workbook to other folders as well.

Global flags, accepted with any command: `--log-level INFO` (or `DEBUG`) prints
progress on stderr, and `--log-file run.log` appends a full log.

## From Python

```python
from dgeg_fuelsales import (
    build,
    download_latest,
    read_dgeg_file,
    read_panel,
    seasonally_adjust,
    update_workbook,
)

path = download_latest()  # data/raw/dgeg/dgeg-omn-YYYY-MM.xlsx
dgeg = read_dgeg_file(path)  # DgegFile(year, months)
result = update_workbook(dgeg)  # UpdateResult(added, revisions, saved)
for r in result.revisions:
    print(r.month, r.product, r.old, "->", r.new)

panel = read_panel()  # long DataFrame: year, month, product, consumption
import pandas as pd

ts = pd.read_csv("outputs/oil_products_timeseries.csv", encoding="utf-8-sig")
ts.index = pd.date_range("1995-01-01", periods=len(ts), freq="MS")
sa = seasonally_adjust(ts.drop(columns="period"))  # same shape; blank where not adjusted
build(also_to=["../my_folder"])  # write outputs/ and copy them
```

## Seasonal adjustment

`build` also writes `outputs/oil_products_timeseries_sa.csv` / `.xlsx`, each
product seasonally adjusted with STL ({func}`~dgeg_fuelsales.seasonally_adjust`).
Products that are too short, have gaps or zero months, or fail the
annual-totals check are left blank; `--log-level INFO` names the products that
fail the check. The thresholds are {data}`~dgeg_fuelsales.SA_MIN_MONTHS` and
{data}`~dgeg_fuelsales.SA_MAX_ANNUAL_GAP`. The method is described in the
[data dictionary](data.md).

## If DGEG changes its file

{func}`~dgeg_fuelsales.read_dgeg_file` maps each report row to a workbook
column by name and stops with an error on a row it does not know, or when the
number of products is not 32. If DGEG adds or renames a product, add it to
`_COLUMNS` in `src/dgeg_fuelsales/core.py` (and a column to the workbook) and
run the update again.

## The workbook's formulas

Some aggregate columns in older years of the workbook are Excel formulas
(e.g. `=SUM(X2:Z2)`). {func}`~dgeg_fuelsales.read_panel` evaluates them itself,
so outputs never depend on values cached by Excel. Saving the workbook from
Python drops those cached values, so tools that read cached values directly
(such as `pandas.read_excel`) see blanks in those cells until Excel recalculates
the file; the package flags the workbook to recalculate when Excel opens it.
