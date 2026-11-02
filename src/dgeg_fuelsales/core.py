"""
Update and build the DGEG monthly oil-product sales dataset.

DGEG (Direção-Geral de Energia e Geologia) publishes Portugal's monthly sales of
oil products as one Excel file per year, *Vendas Mensais de Produtos de Petróleo
em Portugal* (``dgeg-omn-YYYY-MM.xlsx``), on its "Vendas mensais" page. Each file
is a formatted report: a title block, a header row of month names, one row per
product grouped into market sections, and footnotes. Months not yet published
are filled with zeros.

This module keeps a long-run copy of those figures in a source workbook,
``data/processed/dataset_fuel_sales.xlsx`` (one sheet per year from 1995, one
column per product), and turns it into the published dataset. Concretely:

  * :func:`download_latest` finds the newest DGEG file on the page and saves it,
  * :func:`read_dgeg_file` parses one DGEG file into its published months,
    mapping each report row to a workbook column by product name,
  * :func:`update_workbook` writes those months into the year's sheet and
    reports any values DGEG revised,
  * :func:`read_panel` reads the workbook into a long table, evaluating the
    workbook's formula cells itself (so it never depends on values cached by
    Excel),
  * :func:`build` writes the panel, wide and time-series files to ``outputs/``.

The module also defines a small :class:`_Cli` facade that `python-fire
<https://github.com/google/python-fire>`_ turns into a command-line interface.

Data source:
    https://www.dgeg.gov.pt/pt/estatistica/energia/petroleo-e-derivados/vendas-mensais/

Example:
    >>> from dgeg_fuelsales import (
    ...     build,
    ...     download_latest,
    ...     read_dgeg_file,
    ...     update_workbook,
    ... )
    >>> path = download_latest()  # newest dgeg-omn-YYYY-MM.xlsx into data/raw/dgeg/
    >>> update_workbook(read_dgeg_file(path))  # add new months to the workbook
    >>> build()  # rebuild everything in outputs/

Command line:
    dgeg-fuelsales refresh
    dgeg-fuelsales update --file data/raw/dgeg/dgeg-omn-2026-08.xlsx
    dgeg-fuelsales build --also-to ../dgeg_fuel
"""

from __future__ import (  # Enables modern (PEP 604 / postponed) annotations.
    annotations,
)

import re  # Parses DGEG links, row labels and workbook formulas.
import shutil  # Copies finished files to extra folders (--also-to).
import time  # Waits between download retries.
from dataclasses import dataclass  # Builds the immutable result types.
from operator import itemgetter  # Sorts DGEG links by (year, month).
from pathlib import Path  # Cross-platform filesystem paths.
from typing import TYPE_CHECKING  # Typing-only imports below.

import numpy as np  # NaN handling and NaN-aware sums.
import openpyxl  # Reads and writes the source workbook, formulas included.
import pandas as pd  # The tabular data type for every output.
import requests  # HTTP client for the DGEG website.
from openpyxl.utils import range_boundaries  # "B2:D2" -> column/row bounds.
from wrapkit.cli import run  # Shared python-fire runner (errors -> exit 1).
from wrapkit.log import configure_logging, package_logger  # Shared logging.

if TYPE_CHECKING:
    # Imported only for type checking to keep the runtime import graph small;
    # these names are used purely in annotations.
    from collections.abc import Iterable

    from openpyxl.worksheet.worksheet import Worksheet

__all__ = [
    "DEFAULT_OUTPUTS",
    "DEFAULT_RAW_DIR",
    "DEFAULT_WORKBOOK",
    "MONTHS",
    "OUTPUT_FILES",
    "BuildResult",
    "DgegFile",
    "Revision",
    "UpdateResult",
    "build",
    "download_latest",
    "latest_file_url",
    "main",
    "read_dgeg_file",
    "read_panel",
    "update_workbook",
]

#: The package logger; silent until a command line attaches handlers.
_log = package_logger("dgeg_fuelsales")

#: DGEG page that lists the monthly sales files, one per year.
PAGE_URL = "https://www.dgeg.gov.pt/pt/estatistica/energia/petroleo-e-derivados/vendas-mensais/"
#: Site root that the page's relative ``/media/...`` links hang off.
SITE_URL = "https://www.dgeg.gov.pt"

#: Source workbook: one sheet per year, month names in column A.
DEFAULT_WORKBOOK = Path("data/processed/dataset_fuel_sales.xlsx")
#: Where downloaded DGEG files are kept.
DEFAULT_RAW_DIR = Path("data/raw/dgeg")
#: Where the published dataset is written.
DEFAULT_OUTPUTS = Path("outputs")

#: Month names as DGEG and the workbook write them, January first.
MONTHS = (
    "Janeiro",
    "Fevereiro",
    "Março",
    "Abril",
    "Maio",
    "Junho",
    "Julho",
    "Agosto",
    "Setembro",
    "Outubro",
    "Novembro",
    "Dezembro",
)

#: Files written by :func:`build`, all under the outputs folder.
OUTPUT_FILES = (
    "oil_products_panel.csv",
    "oil_products_wide.csv",
    "oil_products_wide.xlsx",
    "oil_products_timeseries.csv",
    "oil_products_timeseries.xlsx",
)

#: DGEG row label (footnote marker removed) -> workbook column.
_COLUMNS = {
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

#: Rows under "Mercado de Bancas Marítimas" reuse generic product names, so
#: inside that section they map to the marine columns instead.
_BUNKERS = {
    "Gasóleo": "Gasóleo marítimo",
    "Gasóleo colorido e marcado": "Gasóleo colorido e marcado marítimo",
    "Fuelóleo Marítimo": "Fuelóleo nº4 BTE marítimo",
    "Lubrificantes": "Lubrificantes marítimo",
}

#: Values closer than this count as unchanged.
_TOLERANCE = 1e-9

#: Number of product rows in a DGEG monthly file.
_PRODUCTS_PER_FILE = 32

#: One signed term of a workbook formula: ``SUM(B2:D2)`` or a cell like ``Y2``.
_TERM = re.compile(r"([+-]?)(SUM\(([A-Z]+\d+:[A-Z]+\d+)\)|[A-Z]+\d+)")

#: Link to a yearly DGEG file on the "Vendas mensais" page.
_FILE_LINK = re.compile(r'href="(/media/[^"]+/dgeg-omn-(\d{4})-(\d{2})\.xlsx)"')

#: Browser-like User-Agent; the DGEG site serves the page either way, but this
#: keeps requests looking like an ordinary visit.
_USER_AGENT = "Mozilla/5.0 (compatible; dgeg-fuelsales; +https://github.com/hmorao95/dgeg_fuelsales_dataset)"


# --------------------------------------------------------------------------- #
# Result types
# --------------------------------------------------------------------------- #


@dataclass(frozen=True)
class DgegFile:
    """
    The published months of one DGEG monthly sales file.

    Attributes:
        year (int): The year the file covers (from its title).
        months (dict[int, dict[str, float]]): Month number (1-12) to a mapping
            of workbook column to value in tonnes. Only published months are
            present; DGEG fills later months with zeros.
    """

    year: int
    months: dict[int, dict[str, float]]

    @property
    def last_month(self) -> int:
        """
        The latest published month.

        Returns:
            int: Month number (1-12), or 0 when the file has no published month.
        """
        return max(self.months, default=0)


@dataclass(frozen=True)
class Revision:
    """
    A value already in the workbook that DGEG has since changed.

    Attributes:
        month (int): Month number (1-12).
        product (str): Workbook column name.
        old (float | None): The value the workbook held (``None`` if blank).
        new (float): DGEG's current value.
    """

    month: int
    product: str
    old: float | None
    new: float


@dataclass(frozen=True)
class UpdateResult:
    """
    What :func:`update_workbook` changed.

    Attributes:
        year (int): The year updated.
        added (tuple[str, ...]): Names of months that were empty and are now
            filled.
        revisions (tuple[Revision, ...]): Values changed in months that were
            already filled.
        saved (bool): Whether the workbook was written (only when something
            changed).
    """

    year: int
    added: tuple[str, ...]
    revisions: tuple[Revision, ...]
    saved: bool


@dataclass(frozen=True)
class BuildResult:
    """
    What :func:`build` wrote.

    Attributes:
        panel_rows (int): Rows in the long panel.
        products (int): Number of distinct products.
        first (str): First period, as ``YYYYmM``.
        last (str): Last period, as ``YYYYmM``.
        months (int): Number of monthly periods in the time series.
        files (tuple[Path, ...]): Every file written, extra copies included.
    """

    panel_rows: int
    products: int
    first: str
    last: str
    months: int
    files: tuple[Path, ...]


# --------------------------------------------------------------------------- #
# Download
# --------------------------------------------------------------------------- #


def _get(
    url: str,
    session: requests.Session | None = None,
    retries: int = 3,
    wait: float = 5.0,
) -> bytes:
    """
    Fetch a URL, retrying on errors.

    This is the module's only network call; tests replace it. The DGEG page
    intermittently answers with a server error, so failed attempts (connection
    errors and HTTP error statuses alike) are retried after a pause.

    Args:
        url (str): The address to fetch.
        session (requests.Session | None): Session to use; a fresh one with a
            browser-like User-Agent when ``None``.
        retries (int): Total number of attempts (at least 1).
        wait (float): Seconds to wait between attempts.

    Returns:
        bytes: The response body.

    Raises:
        ValueError: If ``retries`` is less than 1.
    """
    if retries < 1:
        msg = f"retries must be at least 1, got {retries}"
        raise ValueError(msg)
    http = session or requests.Session()
    if session is None:
        http.headers["User-Agent"] = _USER_AGENT
    for attempt in range(1, retries):
        try:
            return _fetch(http, url)
        except requests.RequestException:  # ruff: ignore[try-except-in-loop] - retrying is the point
            _log.debug("GET %s failed (attempt %d of %d)", url, attempt, retries)
            time.sleep(wait)
    # Last attempt: an error here propagates to the caller.
    return _fetch(http, url)


def _fetch(http: requests.Session, url: str) -> bytes:
    """
    Make one GET request.

    Args:
        http (requests.Session): The session.
        url (str): The address to fetch.

    Returns:
        bytes: The response body.
    """
    response = http.get(url, timeout=60)
    response.raise_for_status()  # HTTP error statuses raise requests.HTTPError
    _log.debug("GET %s -> %d bytes", url, len(response.content))
    return response.content


def latest_file_url(html: str) -> str:
    """
    Find the newest yearly DGEG file linked from the "Vendas mensais" page.

    Files are named ``dgeg-omn-YYYY-MM.xlsx``, where ``YYYY-MM`` is the latest
    month covered; the link with the largest year and month wins.

    Args:
        html (str): The page's HTML.

    Returns:
        str: Absolute URL of the newest file.

    Raises:
        ValueError: If the page links no such file.
    """
    links: list[tuple[str, str, str]] = _FILE_LINK.findall(html)  # (path, year, month)
    if not links:
        msg = f"no dgeg-omn-YYYY-MM.xlsx link found on {PAGE_URL}"
        raise ValueError(msg)
    path, _, _ = max(links, key=itemgetter(1, 2))
    return SITE_URL + path


def download_latest(
    dest_dir: str | Path = DEFAULT_RAW_DIR,
    session: requests.Session | None = None,
) -> Path:
    """
    Download the newest DGEG monthly sales file.

    Reads the "Vendas mensais" page, picks the newest ``dgeg-omn-YYYY-MM.xlsx``
    link and saves the file under its own name, replacing any earlier copy.

    Args:
        dest_dir (str | Path): Folder to save into (created if missing).
        session (requests.Session | None): Optional session for the requests.

    Returns:
        Path: Where the file was saved.
    """
    html = _get(PAGE_URL, session).decode("utf-8", "replace")
    url = latest_file_url(html)
    dest = Path(dest_dir) / url.rsplit("/", 1)[1]
    dest.parent.mkdir(parents=True, exist_ok=True)
    dest.write_bytes(_get(url, session))
    _log.info("Downloaded %s", dest)
    return dest


# --------------------------------------------------------------------------- #
# DGEG file -> workbook
# --------------------------------------------------------------------------- #


def _label(value: object) -> str:
    """
    Normalise a DGEG row label.

    Strips whitespace and a trailing footnote marker such as ``(a)``, so
    ``"Gasóleo Rodoviário (c)"`` becomes ``"Gasóleo Rodoviário"``.

    Args:
        value (object): The cell in the label column.

    Returns:
        str: The cleaned label (``""`` for a blank cell).
    """
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return ""
    return re.sub(r"\s*\([a-z]\)$", "", str(value).strip())


def _column_for(label: str, section: str) -> str | None:
    """
    Map a DGEG row label to its workbook column.

    Args:
        label (str): The cleaned row label.
        section (str): ``"bancas"`` inside the marine bunker section, where
            generic names mean the marine products; anything else otherwise.

    Returns:
        str | None: The workbook column, or ``None`` for an unknown label.
    """
    if section == "bancas" and label in _BUNKERS:
        return _BUNKERS[label]
    return _COLUMNS.get(label)


def _locate(grid: np.ndarray, labels: list[str], path: str | Path) -> tuple[int, int]:
    """
    Find a DGEG file's year and the row of its month header.

    Args:
        grid (np.ndarray): The sheet's cells, row by row.
        labels (list[str]): The cleaned first-column labels.
        path (str | Path): The file, for error messages.

    Returns:
        tuple[int, int]: The year (from the title) and the index of the
        ``Mercado Interno`` header row.

    Raises:
        ValueError: If there is no title with a year, no header row, or the
            header's month names are not the expected ones.
    """
    title = next((s for s in labels if s.startswith("Vendas Mensais")), "")
    year_match = re.search(r"\d{4}", title)
    if "Mercado Interno" not in labels or year_match is None:
        msg = f"{path}: not a DGEG monthly sales file (no title or 'Mercado Interno' header)"
        raise ValueError(msg)
    header = labels.index("Mercado Interno")
    found_months = [str(grid[header, m]).strip() for m in range(1, 13)]
    if found_months != list(MONTHS):
        msg = f"{path}: unexpected month header {found_months}"
        raise ValueError(msg)
    return int(year_match.group()), header


def _product_rows(
    grid: np.ndarray, labels: list[str], header: int, path: str | Path
) -> dict[str, list[float]]:
    """
    Read a DGEG file's product rows into workbook columns.

    Reads from just below the month header down to the footnotes, skipping
    blank rows, the ``Memo Fuel`` heading and the ``Do qual ...`` breakdowns.

    Args:
        grid (np.ndarray): The sheet's cells, row by row.
        labels (list[str]): The cleaned first-column labels.
        header (int): Index of the month header row.
        path (str | Path): The file, for error messages.

    Returns:
        dict[str, list[float]]: Workbook column to its 12 monthly values
        (NaN where blank).

    Raises:
        ValueError: For an unknown product row (add it to ``_COLUMNS``) or a
            product count other than 32.
    """
    values: dict[str, list[float]] = {}
    section = "interno"
    for r in range(header + 1, len(labels)):
        label = labels[r]
        if label.startswith("("):
            break  # the footnotes start; no more data rows
        # Section headers change how generic names (Gasóleo, Lubrificantes) map.
        if label.startswith("Mercado de Bancas"):
            section = "bancas"
        elif label.startswith(("Mercado de Avia", "Biocombust", "Memo Fuel")):
            section = "outros"
        if not label or label == "Memo Fuel" or label.startswith("Do qual"):
            continue
        column = _column_for(label, section)
        if column is None:
            msg = f"{path}: unmapped DGEG row {r}: {label!r}"
            raise ValueError(msg)
        values[column] = [_to_number(grid[r, m]) for m in range(1, 13)]
    if len(values) != _PRODUCTS_PER_FILE:
        msg = f"{path}: expected {_PRODUCTS_PER_FILE} products, found {len(values)}"
        raise ValueError(msg)
    return values


def read_dgeg_file(path: str | Path) -> DgegFile:
    """
    Parse a DGEG monthly sales file.

    Locates the title (for the year) and the ``Mercado Interno`` header row
    (for the month columns), then reads every product row down to the
    footnotes, skipping the ``Do qual ...`` breakdown rows and section
    headers. A month counts as published when any product has a non-zero
    value.

    A file whose layout is not the expected one (no title or header,
    unexpected month names, an unknown product row, or a product count other
    than 32) raises :class:`ValueError`; an unknown product means DGEG added
    one, and it must be added to ``_COLUMNS``.

    Args:
        path (str | Path): The ``dgeg-omn-YYYY-MM.xlsx`` file.

    Returns:
        DgegFile: The year and its published months.
    """
    grid = pd.read_excel(path, header=None).to_numpy(dtype=object)
    labels = [_label(v) for v in grid[:, 0]]
    year, header = _locate(grid, labels, path)
    values = _product_rows(grid, labels, header, path)

    months: dict[int, dict[str, float]] = {}
    for m in range(12):
        column_values = {c: v[m] for c, v in values.items() if not np.isnan(v[m])}
        if any(column_values.values()):  # unpublished months are all zero
            months[m + 1] = column_values
    _log.info(
        "%s: %d, published through %s",
        Path(path).name,
        year,
        MONTHS[max(months, default=1) - 1],
    )
    return DgegFile(year=year, months=months)


def _new_year_sheet(wb: openpyxl.Workbook, year: int) -> Worksheet:
    """
    Add a sheet for a new year, copying the latest sheet's product header.

    Args:
        wb (openpyxl.Workbook): The source workbook.
        year (int): The new year.

    Returns:
        Worksheet: The new sheet, with the header row and month names filled.
    """
    previous = wb[wb.sheetnames[-1]]
    ws: Worksheet = wb.create_sheet(str(year))
    for c in range(1, previous.max_column + 1):
        ws.cell(1, c, previous.cell(1, c).value)
    for i, name in enumerate(MONTHS):
        ws.cell(i + 2, 1, name)
    _log.info("Created sheet %d", year)
    return ws


def update_workbook(
    dgeg: DgegFile,
    workbook: str | Path = DEFAULT_WORKBOOK,
) -> UpdateResult:
    """
    Write a DGEG file's published months into the workbook.

    Every published month is written into the year's sheet (created when it is
    a new year). A month that was empty counts as added; in a month that was
    already filled, each changed value is a revision and DGEG's value wins.
    The workbook is only saved when something changed.

    Saving with openpyxl drops the values Excel caches for formula cells, so
    the workbook is flagged to recalculate when Excel next opens it.
    :func:`read_panel` evaluates the formulas itself and is not affected.

    Args:
        dgeg (DgegFile): The parsed DGEG file.
        workbook (str | Path): The source workbook.

    Returns:
        UpdateResult: Months added, values revised, and whether it was saved.

    Raises:
        ValueError: If the year's sheet lacks a column the DGEG file has.
    """
    wb = openpyxl.load_workbook(workbook)
    sheet = str(dgeg.year)
    ws = wb[sheet] if sheet in wb.sheetnames else _new_year_sheet(wb, dgeg.year)
    cols = {
        str(ws.cell(1, c).value).strip(): c
        for c in range(2, ws.max_column + 1)
        if ws.cell(1, c).value
    }
    rows = {
        str(ws.cell(r, 1).value).strip(): r
        for r in range(2, ws.max_row + 1)
        if ws.cell(r, 1).value
    }
    needed = {name for month in dgeg.months.values() for name in month}
    missing = sorted(needed - set(cols))
    if missing:
        msg = f"sheet {sheet} has no column for: {', '.join(missing)}"
        raise ValueError(msg)

    added: list[str] = []
    revisions: list[Revision] = []
    for month, values in dgeg.months.items():
        r = rows[MONTHS[month - 1]]
        was_empty = all(ws.cell(r, c).value in {None, ""} for c in cols.values())
        for name, value in values.items():
            cell = ws.cell(r, cols[name])
            old = cell.value
            if isinstance(old, (int, float)) and abs(old - value) < _TOLERANCE:
                continue
            cell.value = int(value) if value.is_integer() else value
            if not was_empty:
                revisions.append(
                    Revision(
                        month,
                        name,
                        float(old) if isinstance(old, (int, float)) else None,
                        value,
                    )
                )
        if was_empty:
            added.append(MONTHS[month - 1])

    saved = bool(added or revisions)
    if saved:
        wb.calculation.fullCalcOnLoad = True
        wb.save(workbook)
    _log.info(
        "Sheet %s: %d month(s) added, %d value(s) revised",
        sheet,
        len(added),
        len(revisions),
    )
    return UpdateResult(dgeg.year, tuple(added), tuple(revisions), saved)


# --------------------------------------------------------------------------- #
# Workbook -> datasets
# --------------------------------------------------------------------------- #


def _to_number(value: object) -> float:
    """
    Convert a workbook cell to a number.

    Older sheets store some values as text with thousands spaces or footnote
    asterisks (``"12 345*"``); those are cleaned first.

    Args:
        value (object): The cell value.

    Returns:
        float: The number, or NaN for a blank, ``"."`` or unreadable cell.
    """
    if value is None:
        return float("nan")
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).replace("*", "").replace(" ", "").strip()
    try:
        return float(text)
    except ValueError:
        return float("nan")


def _cell_value(ws: Worksheet, coord: str) -> float:
    """
    Read a cell as a number, evaluating it if it holds a formula.

    Args:
        ws (Worksheet): The sheet.
        coord (str): The cell address, e.g. ``"W2"``.

    Returns:
        float: The value (NaN for a blank cell).
    """
    value = ws[coord].value
    if isinstance(value, str) and value.startswith("="):
        return _evaluate(ws, value)
    return _to_number(value)


def _evaluate(ws: Worksheet, formula: str) -> float:
    """
    Evaluate a workbook formula made of sums and differences.

    The workbook's aggregate columns use ``=SUM(range)`` and
    ``=SUM(range)-cell``; any sum or difference of ``SUM(range)`` terms and
    single cells is supported, with blank cells counting as zero (as in
    Excel). Referenced formulas are evaluated recursively.

    Args:
        ws (Worksheet): The sheet holding the formula.
        formula (str): The formula, with its leading ``=``.

    Returns:
        float: The result.

    Raises:
        ValueError: For any other kind of formula.
    """
    body = formula[1:].replace(" ", "")
    terms = _TERM.findall(body)
    # Every character must belong to a recognised term, or the formula has
    # something (another function, a constant, ...) this evaluator ignores.
    if "".join(sign + term for sign, term, _ in terms) != body:
        msg = f"unsupported formula in sheet {ws.title}: {formula}"
        raise ValueError(msg)
    total = 0.0
    for sign, term, cell_range in terms:
        if cell_range:
            c1, r1, c2, r2 = range_boundaries(cell_range)
            values = [
                _cell_value(ws, ws.cell(r, c).coordinate)
                for r in range(r1 or 1, (r2 or 1) + 1)
                for c in range(c1 or 1, (c2 or 1) + 1)
            ]
            x = float(np.nansum(values))
        else:
            x = _cell_value(ws, term)
            x = 0.0 if np.isnan(x) else x
        total += -x if sign == "-" else x
    return total


def read_panel(workbook: str | Path = DEFAULT_WORKBOOK) -> pd.DataFrame:
    """
    Read the workbook into a long product x month table.

    Every year sheet contributes one row per product column and month; the
    ``TOTAL`` row, footnotes and blank rows are skipped. A product listed in a
    year's sheet but blank for a month keeps its row with a missing value.
    Months not yet published (every product blank) are dropped.

    Args:
        workbook (str | Path): The source workbook.

    Returns:
        pd.DataFrame: Columns ``year``, ``month``, ``product``, ``consumption``
        (tonnes), sorted by product, year and month.
    """
    wb = openpyxl.load_workbook(workbook)  # formulas, not cached values
    rows: list[tuple[int, int, str, float]] = []
    for ws in wb.worksheets:
        year = int(ws.title)
        products = [
            (c, str(ws.cell(1, c).value).strip())
            for c in range(2, ws.max_column + 1)
            if ws.cell(1, c).value
        ]
        for r in range(2, ws.max_row + 1):
            name = str(ws.cell(r, 1).value or "").strip()
            if name not in MONTHS:
                continue  # TOTAL row, footnotes, blanks
            month = MONTHS.index(name) + 1
            rows.extend(
                (year, month, product, _cell_value(ws, ws.cell(r, c).coordinate))
                for c, product in products
            )
    panel = pd.DataFrame(rows, columns=["year", "month", "product", "consumption"])
    panel = panel.drop_duplicates(["product", "year", "month"], keep="first")
    published = panel.groupby(["year", "month"])["consumption"].transform(
        lambda s: s.notna().any()
    )
    return (
        panel[published.astype(bool)]
        .sort_values(["product", "year", "month"], kind="stable")
        .reset_index(drop=True)
    )


def build(
    workbook: str | Path = DEFAULT_WORKBOOK,
    outputs: str | Path = DEFAULT_OUTPUTS,
    also_to: Iterable[str | Path] = (),
) -> BuildResult:
    """
    Write the published dataset from the workbook.

    Writes, under ``outputs``:

    * ``oil_products_panel.csv``: long, one row per product x month, with
      ``period`` (``YYYYmM``), ``date_monthly`` (``YYYY-MM-01``), ``year``,
      ``month``, ``product_id`` (1-50, alphabetical), ``product`` and
      ``consumption`` (tonnes);
    * ``oil_products_wide.csv`` / ``.xlsx``: one column per product, led by
      ``period`` and ``date_monthly``;
    * ``oil_products_timeseries.csv`` / ``.xlsx``: the same wide table led by
      ``period`` only.

    The wide tables have a gap-free monthly index from the first to the last
    published month. CSVs are UTF-8 with a byte-order mark so Excel shows the
    Portuguese names correctly.

    Args:
        workbook (str | Path): The source workbook.
        outputs (str | Path): Folder for the files (created if missing).
        also_to (Iterable[str | Path]): Extra folders that receive a copy of
            every output file and of the workbook.

    Returns:
        BuildResult: Sizes, period range and the files written.
    """
    out = Path(outputs)
    out.mkdir(parents=True, exist_ok=True)
    p = read_panel(workbook)
    p["period"] = p["year"].astype(str) + "m" + p["month"].astype(str)
    p["date_monthly"] = pd.to_datetime(
        {"year": p["year"], "month": p["month"], "day": 1}
    )
    ids = {name: i + 1 for i, name in enumerate(sorted(p["product"].unique()))}
    p["product_id"] = p["product"].map(ids)

    panel = p.assign(date_monthly=p["date_monthly"].dt.strftime("%Y-%m-01"))
    panel[
        [
            "period",
            "date_monthly",
            "year",
            "month",
            "product_id",
            "product",
            "consumption",
        ]
    ].to_csv(out / "oil_products_panel.csv", index=False, encoding="utf-8-sig")

    # pivot (not pivot_table) keeps missing cells as NaN; the keys are unique.
    wide = p.pivot(index="date_monthly", columns="product", values="consumption")  # ruff: ignore[pandas-use-of-dot-pivot-or-unstack]
    wide = wide.reindex(pd.date_range(wide.index.min(), wide.index.max(), freq="MS"))
    wide.index.name = "date_monthly"
    wide.columns.name = None
    period = [f"{d.year}m{d.month}" for d in wide.index]

    w = wide.reset_index()
    w.insert(0, "period", period)
    w.to_excel(out / "oil_products_wide.xlsx", index=False)
    w.assign(date_monthly=w["date_monthly"].dt.strftime("%Y-%m-01")).to_csv(
        out / "oil_products_wide.csv", index=False, encoding="utf-8-sig"
    )

    ts = wide.reset_index(drop=True)
    ts.insert(0, "period", period)
    ts.to_excel(out / "oil_products_timeseries.xlsx", index=False)
    ts.to_csv(out / "oil_products_timeseries.csv", index=False, encoding="utf-8-sig")

    files = [out / name for name in OUTPUT_FILES]
    for folder in map(Path, also_to):
        files.extend(
            Path(shutil.copy2(out / name, folder / name)) for name in OUTPUT_FILES
        )
        files.append(Path(shutil.copy2(workbook, folder / Path(workbook).name)))
        _log.info("Copied outputs and workbook to %s", folder)

    result = BuildResult(
        len(panel), len(ids), period[0], period[-1], len(period), tuple(files)
    )
    _log.info(
        "Built %d panel rows, %s -> %s", result.panel_rows, result.first, result.last
    )
    return result


# --------------------------------------------------------------------------- #
# Command line
# --------------------------------------------------------------------------- #


def _as_paths(value: str | Path | Iterable[str | Path] | None) -> tuple[Path, ...]:
    """
    Normalise an ``--also-to`` argument.

    Fire passes one folder as a string and several (``--also-to '[a,b]'``) as
    a list.

    Args:
        value (str | Path | Iterable[str | Path] | None): What Fire passed.

    Returns:
        tuple[Path, ...]: The folders (empty when none were given).
    """
    if value is None:
        return ()
    if isinstance(value, (str, Path)):
        return (Path(value),)
    return tuple(Path(v) for v in value)


def _describe_update(result: UpdateResult) -> str:
    """
    Format an :class:`UpdateResult` for the terminal.

    Args:
        result (UpdateResult): What the update changed.

    Returns:
        str: One status line, plus one line per revised value.
    """
    added = ", ".join(result.added) or "no new months"
    saved = "saved" if result.saved else "unchanged"
    lines = [
        f"{result.year}: added {added}; {len(result.revisions)} revised value(s); workbook {saved}"
    ]
    lines.extend(
        f"  {MONTHS[v.month - 1]} {v.product}: {v.old} -> {v.new:g}"
        for v in result.revisions
    )
    return "\n".join(lines)


def _describe_build(result: BuildResult) -> str:
    """
    Format a :class:`BuildResult` for the terminal.

    Args:
        result (BuildResult): What the build wrote.

    Returns:
        str: One status line.
    """
    return (
        f"panel: {result.panel_rows} rows, {result.products} products | "
        f"time series: {result.first} -> {result.last} ({result.months} months) | "
        f"{len(result.files)} files written"
    )


class _Cli:
    """
    Command-line facade over the module's functions.

    `python-fire <https://github.com/google/python-fire>`_ turns each method
    here into a subcommand and each parameter into a flag. The methods return
    display-ready text rather than result objects, which Fire would otherwise
    print attribute by attribute.

    The constructor's parameters become global flags, accepted with any
    command: ``--log-level`` and ``--log-file``.
    """

    def __init__(self, log_level: str = "WARNING", log_file: str | None = None) -> None:
        """
        Set up CLI logging.

        Args:
            log_level (str): Console log threshold: ``DEBUG``, ``INFO``,
                ``WARNING`` (default, quiet) or ``ERROR``. Log lines go to
                stderr, so command output on stdout stays clean.
            log_file (str | None): Append a full DEBUG log with timestamps to
                this file.
        """
        configure_logging(_log, log_level, log_file)

    def download(self, dest: str = str(DEFAULT_RAW_DIR)) -> str:  # ruff: ignore[no-self-use] - Fire needs methods
        """
        Download the newest DGEG monthly sales file.

        Args:
            dest (str): Folder to save into.

        Returns:
            str: Where the file was saved.
        """
        return str(download_latest(dest))

    def update(  # ruff: ignore[no-self-use] - Fire needs methods
        self,
        file: str | None = None,
        workbook: str = str(DEFAULT_WORKBOOK),
        dest: str = str(DEFAULT_RAW_DIR),
    ) -> str:
        """
        Add a DGEG file's published months to the workbook.

        Args:
            file (str | None): A downloaded ``dgeg-omn-YYYY-MM.xlsx``; when
                omitted, the newest file is downloaded into ``dest`` first.
            workbook (str): The source workbook.
            dest (str): Download folder, used when ``file`` is omitted.

        Returns:
            str: What changed.
        """
        path = Path(file) if file else download_latest(dest)
        return _describe_update(update_workbook(read_dgeg_file(path), workbook))

    def build(  # ruff: ignore[no-self-use] - Fire needs methods
        self,
        workbook: str = str(DEFAULT_WORKBOOK),
        outputs: str = str(DEFAULT_OUTPUTS),
        also_to: str | list[str] | None = None,
    ) -> str:
        """
        Rebuild every output file from the workbook.

        Args:
            workbook (str): The source workbook.
            outputs (str): Folder for the files.
            also_to (str | list[str] | None): Extra folder(s) that receive a
                copy of the outputs and the workbook.

        Returns:
            str: A summary of what was written.
        """
        return _describe_build(build(workbook, outputs, _as_paths(also_to)))

    def refresh(
        self,
        workbook: str = str(DEFAULT_WORKBOOK),
        outputs: str = str(DEFAULT_OUTPUTS),
        dest: str = str(DEFAULT_RAW_DIR),
        also_to: str | list[str] | None = None,
    ) -> str:
        """
        Download the newest DGEG file, update the workbook, and rebuild.

        Args:
            workbook (str): The source workbook.
            outputs (str): Folder for the files.
            dest (str): Download folder.
            also_to (str | list[str] | None): Extra folder(s) that receive a
                copy of the outputs and the workbook.

        Returns:
            str: What the update changed and what the build wrote.
        """
        return (
            self.update(workbook=workbook, dest=dest)
            + "\n"
            + self.build(workbook, outputs, also_to)
        )


def main() -> None:
    """
    Run the Fire-generated command-line interface.

    Every method of :class:`_Cli` becomes a subcommand and its parameters
    become flags. For example::

        dgeg-fuelsales refresh
        dgeg-fuelsales update --file data/raw/dgeg/dgeg-omn-2026-08.xlsx
        dgeg-fuelsales build --also-to ../dgeg_fuel

    An expected error (an unexpected file layout, a network failure, a missing
    file) prints a concise message and exits with code 1 (``SystemExit``);
    Fire also exits for CLI usage errors.
    """
    run(
        _Cli,
        name="dgeg-fuelsales",
        logger=_log,
        expected=(ValueError, OSError, requests.RequestException),
    )


if __name__ == "__main__":
    main()
