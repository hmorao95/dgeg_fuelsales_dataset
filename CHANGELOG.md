# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com),
and this project adheres to [Semantic Versioning](https://semver.org). Data
updates (new months from DGEG) are minor versions.

## [Unreleased]

### Changed

- wrapkit is now pinned at
  [v0.2.10](https://github.com/hmorao95/wrapkit/releases/tag/v0.2.10) (was
  v0.2.9). Its one fix, a cached answer dated ahead of the clock now counts as
  expired, is in wrapkit's `DiskCache`, which the package does not use (it
  takes only wrapkit's logger, logging setup and Fire runner), so the dataset
  and the command's output are unchanged.

## [1.2.1] - 2026-11-18

### Changed

- wrapkit is now pinned at
  [v0.2.9](https://github.com/hmorao95/wrapkit/releases/tag/v0.2.9) (was
  v0.2.1). The package uses only wrapkit's logger, logging setup and Fire
  runner, so the dataset and the command's output are unchanged:
  - The CLI's one-line `Error:` message is now masked by any `RedactFilter`
    on the package logger. DGEG needs no key and the package installs no
    filter, so error messages print exactly as before. The other
    `RedactFilter` fixes (more HTTP loggers, tracebacks) do not apply for
    the same reason.
  - `snake` now drops accents (`Região` becomes `regiao`). The package does
    not use `snake`: column names such as `Mercado de Aviação` keep their
    spelling, and no column in `outputs/` is renamed.
  - The fixes to incremental CSV updates, period parsing, the disk cache,
    HTTP retries and output paths do not apply: the package does not use
    those helpers.

## [1.2.0] - 2026-11-06

### Added

- September 2026 from DGEG's `dgeg-omn-2026-09.xlsx` (kept in
  `data/raw/dgeg/`); coverage is now 1995m1-2026m9. DGEG revised no earlier
  values.

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

### Removed

- The Stata pipeline (`main.do`, `oil_products_panel.dta`): the panel is
  published as `oil_products_panel.csv`, built in Python.

### Fixed

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

[Unreleased]: https://github.com/hmorao95/dgeg_fuelsales_dataset/compare/v1.2.1...HEAD
[1.2.1]: https://github.com/hmorao95/dgeg_fuelsales_dataset/compare/v1.2.0...v1.2.1
[1.2.0]: https://github.com/hmorao95/dgeg_fuelsales_dataset/compare/c4cd6a6...v1.2.0
[1.1.0]: https://github.com/hmorao95/dgeg_fuelsales_dataset/compare/6c49f5f...c4cd6a6
[1.0.0]: https://github.com/hmorao95/dgeg_fuelsales_dataset/compare/a48e7fa...6c49f5f
