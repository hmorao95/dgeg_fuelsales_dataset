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
├── output/
│   ├── oil_products_panel.dta        # long panel (product × year × month)
│   └── oil_products_wide.xlsx        # wide export
└── _archive/          # duplicates, stale Excel lock file, loose originals
```

## Running

`main.do` sets `$root` to this folder and derives all paths from it.
If you move the folder, update the single `global root` line at the top.

Pipeline: `data/processed/dataset_fuel_sales.xlsx` → `output/oil_products_panel.dta` → `output/oil_products_wide.xlsx`.

## Notes on cleanup (vs. original)

- Duplicate `(1)`/`(2)` files were resolved to a single canonical copy; extras kept in `_archive/`.
- Stale `~$dataset_fuel_sales.xlsx` Excel lock file moved to `_archive/`.
- Loose `fuel_sales.xls` (a duplicate of a `raw/` yearly file) moved to `_archive/`.
