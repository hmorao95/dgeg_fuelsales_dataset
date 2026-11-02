# dgeg_fuel — Portuguese fuel/oil products data

Reorganized copy of the original `dgeg_fuel` project. The original folder was left untouched.

## Structure

```
dgeg_fuel_organized/
├── code/
│   ├── main.do        # Stata pipeline: builds panel + wide dataset
│   └── notas.txt      # Notes on product-name assumptions / imputations
├── data/
│   ├── raw/
│   │   ├── dgeg/       # DGEG sources: yearly fuel_sales_YYYY.*, pgp, pma
│   │   └── eurostat/   # Eurostat nrg_* extracts
│   └── processed/
│       ├── dataset_fuel_sales.xlsx   # INPUT to main.do (one sheet per year)
│       └── complete_fuel_sales.xlsx
├── output/                           # all cover 1995m1–2026m8
│   ├── oil_products_panel.dta / .csv       # long panel (product × month)
│   ├── oil_products_wide.xlsx / .csv       # wide: one column per product
│   └── oil_products_timeseries.xlsx / .csv # wide, gap-free monthly index
└── _archive/          # duplicates, stale Excel lock file, loose originals
```

## Running

`main.do` sets `$root` to this folder and derives all paths from it.
If you move the folder, update the single `global root` line at the top.

Pipeline: `data/processed/dataset_fuel_sales.xlsx` → `output/oil_products_panel.dta` → `output/oil_products_wide.xlsx`.

## Documentation & license

- [DATA_DICTIONARY.md](DATA_DICTIONARY.md) — description of every column, units, aggregates, and naming assumptions.
- [LICENSE](LICENSE) — CC BY 4.0. Underlying statistics are published by DGEG (Portugal) and Eurostat; cite the original sources too.

## Notes on cleanup (vs. original)

- Duplicate `(1)`/`(2)` files were resolved to a single canonical copy; extras kept in `_archive/`.
- Stale `~$dataset_fuel_sales.xlsx` Excel lock file moved to `_archive/`.
- Loose `fuel_sales.xls` (a duplicate of a `raw/` yearly file) moved to `_archive/`.
