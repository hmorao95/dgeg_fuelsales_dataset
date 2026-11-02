# dgeg_fuelsales_dataset

[![CI](https://github.com/hmorao95/dgeg_fuelsales_dataset/actions/workflows/ci.yml/badge.svg)](https://github.com/hmorao95/dgeg_fuelsales_dataset/actions/workflows/ci.yml)
[![Docs](https://github.com/hmorao95/dgeg_fuelsales_dataset/actions/workflows/docs.yml/badge.svg)](https://hmorao95.github.io/dgeg_fuelsales_dataset/)
[![Python](https://img.shields.io/badge/python-3.10%2B-blue.svg)](https://www.python.org/downloads/)
[![uv](https://img.shields.io/endpoint?url=https://raw.githubusercontent.com/astral-sh/uv/main/assets/badge/v0.json)](https://github.com/astral-sh/uv)
[![Ruff](https://img.shields.io/endpoint?url=https://raw.githubusercontent.com/astral-sh/ruff/main/assets/badge/v2.json)](https://github.com/astral-sh/ruff)
[![Checked with mypy](https://img.shields.io/badge/mypy-checked-2a6db2.svg)](https://mypy-lang.org/)
[![pre-commit](https://img.shields.io/badge/pre--commit-enabled-brightgreen?logo=pre-commit)](https://github.com/pre-commit/pre-commit)
[![Keep a Changelog](https://img.shields.io/badge/changelog-Keep%20a%20Changelog-orange.svg)](CHANGELOG.md)
[![License: CC BY 4.0](https://img.shields.io/badge/license-CC%20BY%204.0-lightgrey.svg)](LICENSE)

Monthly sales of **50 oil products in Portugal**, in tonnes, **1995m1–2026m8**,
compiled from the statistics published by DGEG (Direção-Geral de Energia e
Geologia), plus a small, typed Python package that keeps the dataset current.

## Data files

| File | Shape | Contents |
|---|---|---|
| [`oil_products_timeseries.csv`](outputs/oil_products_timeseries.csv) · [`.xlsx`](outputs/oil_products_timeseries.xlsx) | 380 × 51 | Full monthly time series: `period` + one column per product, gap-free from 1995m1 |
| [`oil_products_wide.csv`](outputs/oil_products_wide.csv) · [`.xlsx`](outputs/oil_products_wide.xlsx) | 380 × 52 | Same, with an ISO `date_monthly` column as well |
| [`oil_products_panel.csv`](outputs/oil_products_panel.csv) | 13,120 × 7 | Long panel: one row per product × month |

Every column, unit and aggregate is described in the
[data dictionary](DATA_DICTIONARY.md). Some columns are totals of others
(e.g. *Mercado de Aviação* = *Gasolina de aviação* + *Jet*); don't add them
together with their parts.

Source workbook: [`data/processed/dataset_fuel_sales.xlsx`](data/processed/dataset_fuel_sales.xlsx)
(one sheet per year). DGEG files as downloaded: [`data/raw/dgeg/`](data/raw/dgeg).

## Quickstart

To update the dataset when DGEG publishes a new month you need
[uv](https://docs.astral.sh/uv/):

```bash
# macOS / Linux
curl -LsSf https://astral.sh/uv/install.sh | sh
```

```powershell
# Windows (PowerShell)
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
```

Then:

```bash
# 1. Get the code and data
git clone https://github.com/hmorao95/dgeg_fuelsales_dataset.git
cd dgeg_fuelsales_dataset

# 2. Install (uv creates a .venv and pulls dependencies)
uv sync

# 3. Download DGEG's newest file, add its months, rebuild outputs/
uv run dgeg-fuelsales refresh
```

`refresh` prints the months it added and any values DGEG revised. No API key is
needed.

## Usage from Python

```python
from dgeg_fuelsales import (
    build,
    download_latest,
    read_dgeg_file,
    read_panel,
    update_workbook,
)

# Download and parse DGEG's newest yearly file
path = download_latest()  # data/raw/dgeg/dgeg-omn-YYYY-MM.xlsx
dgeg = read_dgeg_file(path)  # DgegFile(year=2026, months={1: {...}, ...})

# Add its months to the workbook; DGEG's revisions overwrite old values
result = update_workbook(dgeg)
print(result.added, len(result.revisions))

# Long table straight from the workbook (formulas evaluated)
panel = read_panel()

# Rebuild outputs/, optionally copying them elsewhere
build(also_to=["../my_folder"])
```

## Command line

```bash
uv run dgeg-fuelsales refresh                      # download + update + build
uv run dgeg-fuelsales download                     # newest DGEG file -> data/raw/dgeg/
uv run dgeg-fuelsales update                       # download, then add months to the workbook
uv run dgeg-fuelsales update --file data/raw/dgeg/dgeg-omn-2026-08.xlsx
uv run dgeg-fuelsales build                        # rebuild outputs/ from the workbook
uv run dgeg-fuelsales build --also-to ../dgeg_fuel # also copy outputs + workbook there
```

### Options

| Flag | Commands | Meaning |
|---|---|---|
| `--file PATH` | `update` | Use a downloaded DGEG file instead of fetching the newest |
| `--workbook PATH` | `update`, `build`, `refresh` | Source workbook (default `data/processed/dataset_fuel_sales.xlsx`) |
| `--outputs DIR` | `build`, `refresh` | Output folder (default `outputs`) |
| `--dest DIR` | `download`, `update`, `refresh` | Download folder (default `data/raw/dgeg`) |
| `--also-to DIR` | `build`, `refresh` | Copy outputs and workbook to `DIR` too (`'[A,B]'` for several) |
| `--log-level LEVEL` | all | `DEBUG`, `INFO`, `WARNING` (default) or `ERROR`, on stderr |
| `--log-file PATH` | all | Append a full DEBUG log |

## How it works

- DGEG publishes one file per year, *Vendas Mensais de Produtos de Petróleo em
  Portugal* (`dgeg-omn-YYYY-MM.xlsx`), on its
  [Vendas mensais](https://www.dgeg.gov.pt/pt/estatistica/energia/petroleo-e-derivados/vendas-mensais/)
  page, and replaces it as months come out. Unpublished months are zeros.
- `read_dgeg_file` maps each of its 32 product rows to a workbook column by
  name (rows in the marine-bunker section map to the marine columns) and stops
  with an error if DGEG adds or renames a product.
- `update_workbook` writes the published months into the year's sheet and
  reports revisions. The figures are provisional, so DGEG may revise them.
- `read_panel` reads all sheets from 1995, evaluating the workbook's formula
  cells itself, and `build` writes the outputs. A regression test checks that
  the committed outputs match the workbook.

## Repository layout

```
├── outputs/                  the published dataset (CSV / XLSX)
├── data/
│   ├── raw/dgeg/             DGEG files as downloaded (yearly, 1995-)
│   ├── raw/eurostat/         auxiliary Eurostat nrg_* extracts
│   └── processed/            source workbook + naming notes (notas.txt)
├── src/dgeg_fuelsales/       the package (core.py) and CLI
├── tests/                    network-free pytest suite
├── docs/                     Sphinx documentation site
├── DATA_DICTIONARY.md        columns, units, aggregates
└── CHANGELOG.md, CITATION.cff, LICENSE, AGENTS.md
```

## Development

```bash
uv sync
uv run ruff check .
uv run ruff format --check .
uv run codespell src tests README.md CHANGELOG.md DATA_DICTIONARY.md
uv run deptry src
uv run interrogate -c pyproject.toml src tests
uv run mypy
uv run pytest
```

CI (`.github/workflows/ci.yml`) runs the same gate on every push and pull
request; `.github/workflows/docs.yml` builds the documentation and, once the
repository is public, publishes it to GitHub Pages (Pages needs a paid plan for
private repositories). Hooks: `uvx pre-commit install`. See [AGENTS.md](AGENTS.md)
for conventions and the release process.

## Changelog

See [CHANGELOG.md](CHANGELOG.md) (Keep a Changelog format).

## Citation

If you use this dataset, please cite it (see [CITATION.cff](CITATION.cff);
GitHub renders a "Cite this repository" button):

> Morão, H. (2026). *DGEG fuel sales dataset: monthly oil-product sales in
> Portugal, 1995m1–2026m8* (v1.1.0) [Data set].
> https://github.com/hmorao95/dgeg_fuelsales_dataset

Please also cite the original source, DGEG.

## Data source & license

Underlying statistics © DGEG — Direção-Geral de Energia e Geologia (Portugal),
[monthly oil-product sales](https://www.dgeg.gov.pt/pt/estatistica/energia/petroleo-e-derivados/vendas-mensais/);
2026 figures are provisional. Eurostat extracts under `data/raw/eurostat/` are
auxiliary and not part of the outputs.

This dataset and code are released under
[CC BY 4.0](LICENSE). This is an independent project, not affiliated with or
endorsed by DGEG.
