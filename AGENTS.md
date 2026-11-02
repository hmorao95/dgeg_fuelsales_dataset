# Agent guide

Instructions for AI coding agents working in this repository. Human contributors
may find it useful too.

## What this is

`dgeg_fuelsales_dataset` publishes Portugal's monthly oil-product sales (50
products, tonnes, from 1995m1) compiled from DGEG statistics, plus the small,
typed Python package `dgeg_fuelsales` that keeps it current: it downloads DGEG's
newest yearly file (`dgeg-omn-YYYY-MM.xlsx`), writes its published months into
the source workbook, and rebuilds the panel, wide and time-series outputs, with
a python-fire CLI.

Layout:

- `src/dgeg_fuelsales/`: the package (`core.py` is the whole implementation;
  `__init__.py` re-exports; `__main__.py` enables `python -m`).
- `tests/`: network-free pytest suite (synthetic DGEG files and workbooks, plus
  one regression test against the committed outputs).
- `data/raw/`: files as downloaded (DGEG yearly files, Eurostat extracts).
- `data/processed/dataset_fuel_sales.xlsx`: the source workbook, one sheet per
  year; `notas.txt` records naming assumptions.
- `outputs/`: the published dataset. **Committed**, unlike in the API wrappers:
  the data is the product.
- `docs/`: Sphinx site; `DATA_DICTIONARY.md` (root) is included in it.

## Setup

Managed with [uv](https://docs.astral.sh/uv/). One command:

```bash
uv sync
```

Run anything through `uv run`, from the repository root (default paths are
relative to it):

```bash
uv run dgeg-fuelsales refresh
uv run python -m dgeg_fuelsales --help
```

On a machine whose user path has non-ASCII characters, uv's editable install
may not be importable; `pytest` is unaffected (`pythonpath = ["src"]`), and
`PYTHONPATH=src uv run python -m dgeg_fuelsales ...` works.

## Quality gate (must pass before every commit)

CI runs exactly these; run them locally first:

```bash
uv run ruff check .
uv run ruff format --check .
uv run codespell src tests README.md CHANGELOG.md DATA_DICTIONARY.md
uv run deptry src
uv run interrogate -c pyproject.toml src tests         # docstring coverage >= 95%
uv run mypy                                             # strict
uv run pytest                                           # also checks pyproject == CITATION
```

`uv run pytest` also enforces test coverage >= 92% (`fail_under` in
`[tool.coverage.report]`).

Do not weaken the gate to make it pass. Fix the code, or if a rule is genuinely
wrong for a case, add a narrowly-scoped ignore with a comment explaining why.

## Data updates

- Update with `uv run dgeg-fuelsales refresh` (or `update` + `build`). Never
  edit files in `outputs/` by hand; they are always rebuilt from the workbook.
- `test_build_matches_committed_outputs` fails if the workbook changed but
  `outputs/` was not rebuilt: rebuild and commit both together.
- The workbook has formula cells in aggregate columns (1995-2000, 2020-2021).
  openpyxl saves drop Excel's cached values, so never read the workbook with
  `pandas.read_excel` to build outputs; `read_panel` evaluates the formulas.
- A data update bumps the minor version and updates the coverage period in
  `CITATION.cff` (title, abstract), `DATA_DICTIONARY.md`, the README and the
  repository's About description.

## Conventions

- **Docstrings**: verbose Google style: a summary line, an explanatory
  paragraph, then typed `Args:` / `Returns:` / `Raises:`. `interrogate` and the
  ruff `D`/`DOC` rules enforce this. Comment non-obvious lines.
- **Typing**: `mypy --strict` must pass; annotate everything. `from __future__
  import annotations` is on.
- **Exceptions**: assign the message to a `msg` variable, then `raise` (ruff
  `EM`).
- **HTTP**: every request goes through `core._get`, which retries (the DGEG page
  intermittently returns server errors). Tests monkeypatch it or pass a fake
  session; never hit the network in tests.
- **Line length**: formatter targets 88; `E501` allows up to 150.
- **Shared plumbing**: logging and the Fire runner come from
  [wrapkit](https://github.com/hmorao95/wrapkit), pinned by tag in
  `pyproject.toml`.

## Tests

- Never hit the network. See `FakeSession` and the `_get` monkeypatches in
  `tests/test_dgeg_fuelsales.py`.
- Build inputs with `write_dgeg` and `write_workbook`, which reproduce the real
  layouts.
- Every test runs in a fresh temporary working directory (autouse `workdir`
  fixture), so default relative paths never touch the repository's data.
- CLI tests drive Fire directly: `fire.Fire(core._Cli, command=[...])`.

## Git and releases

- Work on a branch (`feature/…`, `fix/…`, `chore/…`, `data/…`), open a PR, and
  squash-merge once CI is green.
- **Commit as the repo owner.** Do not skip hooks or bypass signing.
- **Releasing** (SemVer; data updates are minor versions): after merging,
  1. move the `[Unreleased]` changelog items under a new `## [X.Y.Z] - <date>`
     section and update the compare links at the bottom;
  2. bump the version in **`pyproject.toml`**, **`CITATION.cff`**, the README
     citation line, and **`uv.lock`** (`uv lock`);
  3. commit, push, wait for CI;
  4. `git tag -a vX.Y.Z` and `gh release create vX.Y.Z`.
