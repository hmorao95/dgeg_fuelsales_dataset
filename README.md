# dgeg_fuel — Portuguese fuel/oil products data

Reorganized copy of the original `dgeg_fuel` project. The original folder was left untouched.

## Structure

```
dgeg_fuel_organized/
├── code/
│   ├── build.py       # builds all output/ files from the workbook
│   ├── update_dgeg.py # adds newly published DGEG months to the workbook
│   └── notas.txt      # Notes on product-name assumptions / imputations
├── data/
│   ├── raw/
│   │   ├── dgeg/       # DGEG sources: yearly fuel_sales_YYYY.*, pgp, pma
│   │   └── eurostat/   # Eurostat nrg_* extracts
│   └── processed/
│       ├── dataset_fuel_sales.xlsx   # source workbook (one sheet per year)
│       └── complete_fuel_sales.xlsx
├── output/                           # all cover 1995m1–2026m8
│   ├── oil_products_panel.csv              # long panel (product × month)
│   ├── oil_products_wide.xlsx / .csv       # wide: one column per product
│   └── oil_products_timeseries.xlsx / .csv # wide, gap-free monthly index
└── _archive/          # duplicates, stale Excel lock file, loose originals
```

## Running

Python 3.10+ with the packages in `requirements.txt`:

```bash
pip install -r requirements.txt
```

Rebuild every file in `output/` from the workbook:

```bash
python code/build.py
```

Add `--also-to DIR` to copy the output files and the source workbook to another folder as well.

Monthly update, when DGEG publishes a new month:

```bash
python code/update_dgeg.py --download
python code/build.py
```

`update_dgeg.py --download` fetches the latest *Vendas mensais* file from DGEG into
`data/raw/dgeg/`, writes every published month of that year into the workbook, and
reports any values DGEG revised. You can also pass a file you downloaded yourself:
`python code/update_dgeg.py path/to/dgeg-omn-YYYY-MM.xlsx`.

Pipeline: DGEG `dgeg-omn-YYYY-MM.xlsx` → `data/processed/dataset_fuel_sales.xlsx` → `output/`.

## Documentation & license

- [DATA_DICTIONARY.md](DATA_DICTIONARY.md) — description of every column, units, aggregates, and naming assumptions.
- [LICENSE](LICENSE) — CC BY 4.0. Underlying statistics are published by DGEG (Portugal) and Eurostat; cite the original sources too.

## Notes on cleanup (vs. original)

- Duplicate `(1)`/`(2)` files were resolved to a single canonical copy; extras kept in `_archive/`.
- Stale `~$dataset_fuel_sales.xlsx` Excel lock file moved to `_archive/`.
- Loose `fuel_sales.xls` (a duplicate of a `raw/` yearly file) moved to `_archive/`.
