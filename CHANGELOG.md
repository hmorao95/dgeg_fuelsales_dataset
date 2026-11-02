# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com),
and this project adheres to [Semantic Versioning](https://semver.org). Data
updates (new months from DGEG) are minor versions.

## [Unreleased]

### Changed

- The repository follows the layout of the author's data-API wrappers (e.g.
  [eia-api](https://github.com/hmorao95/eia-api)): a typed `dgeg_fuelsales`
  package in `src/`, managed with uv, with a network-free pytest suite, a Sphinx
  documentation site, and CI running ruff, codespell, deptry, interrogate, mypy
  (strict) and pytest (coverage >= 92%).
- The scripts `code/update_dgeg.py` and `code/build.py` are replaced by the
  `dgeg-fuelsales` command (`download`, `update`, `build`, `refresh`) and the
  functions `download_latest`, `read_dgeg_file`, `update_workbook`,
  `read_panel` and `build`. The outputs are unchanged.
- `output/` is renamed `outputs/`; `code/notas.txt` moves to
  `data/processed/notas.txt`. `requirements.txt` is replaced by
  `pyproject.toml` and `uv.lock`.
- The README links each data file directly.

### Added

- A regression test that rebuilds the workbook and compares the result with
  the committed outputs, so a workbook change without a rebuild fails CI.
- `refresh` downloads, updates and rebuilds in one command, and reports
  revised values.

### Removed (since 1.1.0, before this restructure)

- The Stata pipeline (`main.do`, `oil_products_panel.dta`): the panel is
  published as `oil_products_panel.csv`, built in Python.

### Fixed (since 1.1.0, before this restructure)

- 276 cells in five aggregate columns (Mercado de Aviação, Mercado de Bancas,
  Mercado de Bancas Marítimas, Biocombustíveis, Fuelóleo para os restantes
  setores; 1995-2000 and 2020-2021) were blank in the wide and time-series
  files, because saving the workbook with openpyxl had dropped Excel's cached
  formula values. The build now evaluates the formulas itself.

## [1.1.0] - 2026-11-02

### Added

- June-August 2026 from DGEG's `dgeg-omn-2026-08.xlsx` (kept in
  `data/raw/dgeg/`); coverage is now 1995m1-2026m8.
- `period` column (`YYYYmM`) in the wide files, and the gap-free
  `oil_products_timeseries` files.

## [1.0.0] - 2026-11-02

### Added

- The dataset: monthly sales of 50 oil products, 1995m1-2026m5, as a long
  panel and a wide table, with the raw DGEG and Eurostat files, a data
  dictionary, a CC BY 4.0 license and `CITATION.cff`.

[Unreleased]: https://github.com/hmorao95/dgeg_fuelsales_dataset/compare/c4cd6a6...HEAD
[1.1.0]: https://github.com/hmorao95/dgeg_fuelsales_dataset/compare/6c49f5f...c4cd6a6
[1.0.0]: https://github.com/hmorao95/dgeg_fuelsales_dataset/compare/a48e7fa...6c49f5f
