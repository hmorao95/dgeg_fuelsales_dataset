"""Unit tests for :mod:`dgeg_fuelsales`.

The tests never touch the network: downloads go through :func:`core._get`,
which is monkeypatched, or through a fake ``requests`` session. DGEG files and
source workbooks are synthetic, written with openpyxl in the same layout as
the real ones. One regression test rebuilds the repository's real workbook and
checks it against the committed outputs.
"""

from __future__ import annotations

import sys
from pathlib import Path
from typing import TYPE_CHECKING

import fire
import openpyxl
import pandas as pd
import pytest
import requests

from dgeg_fuelsales import (
    MONTHS,
    OUTPUT_FILES,
    build,
    core,
    download_latest,
    latest_file_url,
    read_dgeg_file,
    read_panel,
    update_workbook,
)

if TYPE_CHECKING:
    from collections.abc import Iterable

ROOT = Path(__file__).resolve().parents[1]

# --------------------------------------------------------------------------- #
# Synthetic DGEG file and workbook
# --------------------------------------------------------------------------- #

#: Workbook columns, in the order of the real 2025/2026 sheets.
PRODUCTS = [
    "Butano",
    "Propano",
    "Gás Auto",
    "Propileno",
    "Gasolina s/Chumbo 98",
    "Gasolina s/Chumbo 95",
    "Nafta Química e Aromáticos",
    "Petróleos (iluminante + carburante)",
    "Gasóleo rodoviário",
    "Gasóleo aquecimento",
    "Gasóleo colorido e marcado",
    "Fuelóleo nº3 BTE",
    "Fuelóleo nº4 BTE",
    "Coque de petróleo",
    "Lubrificantes",
    "Asfaltos",
    "Parafinas",
    "Solventes",
    "Fuelóleo para a produção de energia elétrica",
    "Fuelóleo para os restantes setores",
    "Mercado de Bancas Marítimas",
    "Gasóleo marítimo",
    "Gasóleo colorido e marcado marítimo",
    "Fuelóleo nº4 BTE marítimo",
    "Lubrificantes marítimo",
    "Mercado de Aviação",
    "Gasolina de aviação",
    "Jet",
    "Biocombustíveis",
    "Biodiesel incorporado no gasóleo",
    "Biodiesel fornecido diretamente ao mercado",
    "Biogasolina incorporada",
]

#: DGEG report rows: (label, workbook column it feeds, or None). ``"values"``
#: marks breakdown rows that carry numbers but feed no column; blank strings
#: are spacer rows.
LAYOUT: list[tuple[str, str | None]] = [
    ("Butano", "Butano"),
    ("Do qual engarrafado", "values"),
    ("Do qual petroquimica", "values"),
    ("Propano", "Propano"),
    ("Do qual engarrafado", "values"),
    ("Do qual petroquimica", "values"),
    ("Gás Auto", "Gás Auto"),
    ("Propileno", "Propileno"),
    ("Gasolina s/ Chumbo 98 (a)", "Gasolina s/Chumbo 98"),
    ("Gasolina s/ Chumbo 95 (b)", "Gasolina s/Chumbo 95"),
    ("Nafta Química e Aromáticos", "Nafta Química e Aromáticos"),
    ("Petróleos (Iluminante + carburante)", "Petróleos (iluminante + carburante)"),
    ("Gasóleo Rodoviário (c)", "Gasóleo rodoviário"),
    ("Gasóleo colorido e marcado destinado a aquecimento", "Gasóleo aquecimento"),
    (
        "Gasóleo colorido e marcado, exceto o destinado a aquecimento",
        "Gasóleo colorido e marcado",
    ),
    ("Fuelóleo nº3 BTE", "Fuelóleo nº3 BTE"),
    ("Fuelóleo nº4 BTE", "Fuelóleo nº4 BTE"),
    ("Coque de petróleo", "Coque de petróleo"),
    ("Lubrificantes", "Lubrificantes"),
    ("Asfaltos", "Asfaltos"),
    ("Parafinas", "Parafinas"),
    ("Solventes", "Solventes"),
    ("", None),
    ("Memo Fuel", None),
    (
        "Fuelóleo para a produção de energia elétrica (d)",
        "Fuelóleo para a produção de energia elétrica",
    ),
    ("Fuelóleo para os restantes setores", "Fuelóleo para os restantes setores"),
    ("", None),
    ("Mercado de Bancas Marítimas (e)", "Mercado de Bancas Marítimas"),
    ("Gasóleo (f)", "Gasóleo marítimo"),
    ("Gasóleo colorido e marcado", "Gasóleo colorido e marcado marítimo"),
    ("Fuelóleo Marítimo", "Fuelóleo nº4 BTE marítimo"),
    ("Lubrificantes", "Lubrificantes marítimo"),
    ("", None),
    ("Mercado de Aviação", "Mercado de Aviação"),
    ("Gasolina de aviação", "Gasolina de aviação"),
    ("Jet", "Jet"),
    ("", None),
    ("Biocombustíveis (g)", "Biocombustíveis"),
    ("Biodiesel incorporado no gasóleo", "Biodiesel incorporado no gasóleo"),
    (
        "Biodiesel fornecido diretamente ao mercado",
        "Biodiesel fornecido diretamente ao mercado",
    ),
    ("Biogasolina incorporada", "Biogasolina incorporada"),
    ("", None),
    ("(a) Inclui a gasolina 98 especial.", None),
]


def value(product: str, month: int, shift: float = 0.0) -> float:
    """Return the synthetic value of a product in a month."""
    return (PRODUCTS.index(product) + 1) * 100 + month + shift


def write_dgeg(
    path: Path,
    year: int,
    published: int,
    rows: Iterable[tuple[str, str | None]] = LAYOUT,
    shift: float = 0.0,
) -> Path:
    """Write a DGEG-style monthly sales file with ``published`` months filled."""
    wb = openpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.append([None])
    ws.append(["Direção Geral de Energia e Geologia"])
    ws.append([f"Vendas Mensais de Produtos de Petróleo em Portugal {year}"])
    ws.append(["(dados provisórios)"])
    ws.append([0, *[None] * 12, "unidade: tonelada"])
    ws.append(["Mercado Interno", *MONTHS, "TOTAL"])
    for label, column in rows:
        if column is None:
            ws.append([label or None])
            continue
        product = column if column in PRODUCTS else "Butano"
        cells = [
            value(product, m, shift) if m <= published else 0 for m in range(1, 13)
        ]
        ws.append([label, *cells, sum(cells)])
    wb.save(path)
    return path


def write_workbook(path: Path, filled: dict[int, int]) -> Path:
    """Write a source workbook with one sheet per year and ``filled`` months."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # type: ignore[arg-type]
    for year, months in filled.items():
        ws = wb.create_sheet(str(year))
        ws.append([None, *PRODUCTS])
        for m, name in enumerate(MONTHS, start=1):
            ws.append([name, *[value(p, m) if m <= months else None for p in PRODUCTS]])
        ws.append(["TOTAL"])
    wb.save(path)
    return path


@pytest.fixture
def workbook(tmp_path: Path) -> Path:
    """Return a workbook with 2025 complete and 2026 filled for Jan-Feb."""
    return write_workbook(tmp_path / "dataset_fuel_sales.xlsx", {2025: 12, 2026: 2})


@pytest.fixture
def formula_workbook(tmp_path: Path) -> Path:
    """Return a small workbook with formulas, text numbers and blanks."""
    wb = openpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "1995"
    ws.append([None, "A", "B", "Total", "Net", "Text"])
    ws.append(["Janeiro", 1, 2, "=SUM(B2:C2)", "=SUM(B2:D2)-C2", "1 234*"])
    ws.append(["Fevereiro", 3, None, "=SUM(B3:C3)", "=D3+B3", "."])
    ws.append(["Março", None, None, None, None, None])  # not published yet
    ws.append(["TOTAL", "=SUM(B2:B4)"])
    other = wb.create_sheet("1996")
    other.append([None, "A"])
    other.append(["Janeiro", 5])
    wb.save(tmp_path / "formulas.xlsx")
    return tmp_path / "formulas.xlsx"


# --------------------------------------------------------------------------- #
# Download
# --------------------------------------------------------------------------- #

PAGE = """
<a href="/media/aaa/dgeg-omn-2025-12.xlsx">2025 (xlsx)</a>
<a href="/media/bbb/dgeg-omn-2026-08.xlsx">2026 (xlsx)</a>
<a href="/media/ccc/dgeg-omn-2024-12.xlsx">2024 (xlsx)</a>
<a href="/media/ddd/i007113.xls">1995 a 2000 (xls)</a>
"""


def test_latest_file_url_picks_newest() -> None:
    """The link with the largest year and month wins."""
    assert (
        latest_file_url(PAGE)
        == "https://www.dgeg.gov.pt/media/bbb/dgeg-omn-2026-08.xlsx"
    )


def test_latest_file_url_without_links() -> None:
    """A page with no yearly file is an error."""
    with pytest.raises(ValueError, match="no dgeg-omn"):
        latest_file_url("<html></html>")


def test_download_latest(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    """The newest file is fetched and saved under its own name."""
    calls: list[str] = []

    def fake_get(url: str, session: requests.Session | None = None) -> bytes:
        calls.append(url)
        return PAGE.encode() if url == core.PAGE_URL else b"xlsx-bytes"

    monkeypatch.setattr(core, "_get", fake_get)
    path = download_latest(tmp_path / "raw")
    assert path == tmp_path / "raw" / "dgeg-omn-2026-08.xlsx"
    assert path.read_bytes() == b"xlsx-bytes"
    assert calls == [
        core.PAGE_URL,
        "https://www.dgeg.gov.pt/media/bbb/dgeg-omn-2026-08.xlsx",
    ]


class FakeResponse:
    """A minimal stand-in for :class:`requests.Response`."""

    def __init__(self, status: int, content: bytes = b"ok") -> None:
        """Store the status and body."""
        self.status_code = status
        self.content = content

    def raise_for_status(self) -> None:
        """Raise like requests does for an error status."""
        if self.status_code >= 400:
            raise requests.HTTPError(str(self.status_code))


class FakeSession(requests.Session):
    """A session that replays canned responses instead of using the network."""

    def __init__(self, responses: list[FakeResponse | Exception] | None = None) -> None:
        """Keep the responses to replay, in order."""
        super().__init__()
        self.responses = list(responses or [FakeResponse(200)])
        self.urls: list[str] = []

    def get(self, url: str | bytes, **kwargs: object) -> FakeResponse:  # type: ignore[override]
        """Return (or raise) the next canned response."""
        self.urls.append(str(url))
        response = self.responses.pop(0)
        if isinstance(response, Exception):
            raise response
        return response


@pytest.fixture
def no_sleep(monkeypatch: pytest.MonkeyPatch) -> list[float]:
    """Replace the retry pause with a recorder."""
    waits: list[float] = []
    monkeypatch.setattr("dgeg_fuelsales.core.time.sleep", waits.append)
    return waits


def test_get_retries_then_succeeds(no_sleep: list[float]) -> None:
    """Server errors and connection errors are retried."""
    session = FakeSession(
        [
            FakeResponse(500),
            requests.ConnectionError("down"),
            FakeResponse(200, b"page"),
        ]
    )
    assert core._get("https://x", session, retries=3, wait=2.0) == b"page"
    assert no_sleep == [2.0, 2.0]
    assert len(session.urls) == 3


def test_get_gives_up(no_sleep: list[float]) -> None:
    """The last error propagates once every attempt failed."""
    session = FakeSession([FakeResponse(500), FakeResponse(503)])
    with pytest.raises(requests.HTTPError, match="503"):
        core._get("https://x", session, retries=2)
    assert len(no_sleep) == 1


def test_get_rejects_zero_retries() -> None:
    """At least one attempt is required."""
    with pytest.raises(ValueError, match="at least 1"):
        core._get("https://x", FakeSession(), retries=0)


def test_get_default_session_sets_user_agent(monkeypatch: pytest.MonkeyPatch) -> None:
    """Without a session, a fresh one with the package's User-Agent is used."""
    made: list[FakeSession] = []

    def factory() -> FakeSession:
        made.append(FakeSession())
        return made[-1]

    monkeypatch.setattr("dgeg_fuelsales.core.requests.Session", factory)
    assert core._get("https://x") == b"ok"
    assert "dgeg-fuelsales" in made[0].headers["User-Agent"]


# --------------------------------------------------------------------------- #
# Reading DGEG files
# --------------------------------------------------------------------------- #


def test_read_dgeg_file(tmp_path: Path) -> None:
    """Published months are read, by product, into workbook columns."""
    dgeg = read_dgeg_file(write_dgeg(tmp_path / "f.xlsx", 2027, published=3))
    assert dgeg.year == 2027
    assert sorted(dgeg.months) == [1, 2, 3]
    assert dgeg.last_month == 3
    assert set(dgeg.months[1]) == set(PRODUCTS)
    # Marine rows reuse generic names but land in the marine columns.
    assert dgeg.months[2]["Gasóleo marítimo"] == value("Gasóleo marítimo", 2)
    assert dgeg.months[2]["Lubrificantes"] == value("Lubrificantes", 2)
    assert dgeg.months[3]["Lubrificantes marítimo"] == value(
        "Lubrificantes marítimo", 3
    )


def test_read_dgeg_file_no_months_published(tmp_path: Path) -> None:
    """A file with nothing published has no months."""
    dgeg = read_dgeg_file(write_dgeg(tmp_path / "f.xlsx", 2027, published=0))
    assert dgeg.months == {}
    assert dgeg.last_month == 0


def test_read_dgeg_file_unknown_product(tmp_path: Path) -> None:
    """A product DGEG adds must be mapped before it can be read."""
    rows = [*LAYOUT[:5], ("Hidrogénio", "values"), *LAYOUT[5:]]
    with pytest.raises(ValueError, match=r"unmapped DGEG row .*Hidrogénio"):
        read_dgeg_file(write_dgeg(tmp_path / "f.xlsx", 2027, 1, rows))


def test_read_dgeg_file_missing_product(tmp_path: Path) -> None:
    """A file with fewer products than expected is rejected."""
    rows = [r for r in LAYOUT if r[0] != "Jet"]
    with pytest.raises(ValueError, match="expected 32 products, found 31"):
        read_dgeg_file(write_dgeg(tmp_path / "f.xlsx", 2027, 1, rows))


def test_read_dgeg_file_not_a_dgeg_file(tmp_path: Path) -> None:
    """A workbook without the title and header is rejected."""
    pd.DataFrame({"a": [1]}).to_excel(tmp_path / "other.xlsx", index=False)
    with pytest.raises(ValueError, match="not a DGEG monthly sales file"):
        read_dgeg_file(tmp_path / "other.xlsx")


def test_read_dgeg_file_bad_month_header(tmp_path: Path) -> None:
    """Month names other than DGEG's are rejected."""
    path = write_dgeg(tmp_path / "f.xlsx", 2027, 1)
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    assert ws is not None
    ws.cell(6, 2, "January")
    wb.save(path)
    with pytest.raises(ValueError, match="unexpected month header"):
        read_dgeg_file(path)


def test_read_real_dgeg_file() -> None:
    """The DGEG file kept in the repository parses to 2026 through August."""
    dgeg = read_dgeg_file(ROOT / "data" / "raw" / "dgeg" / "dgeg-omn-2026-08.xlsx")
    assert (dgeg.year, dgeg.last_month) == (2026, 8)
    assert dgeg.months[8]["Jet"] == 191151
    assert dgeg.months[6]["Propileno"] == 6


# --------------------------------------------------------------------------- #
# Updating the workbook
# --------------------------------------------------------------------------- #


def test_update_adds_months(workbook: Path, tmp_path: Path) -> None:
    """Empty months are filled and reported as added."""
    result = update_workbook(
        read_dgeg_file(write_dgeg(tmp_path / "f.xlsx", 2026, 4)), workbook
    )
    assert result.added == ("Março", "Abril")
    assert result.revisions == ()
    assert result.saved
    ws = openpyxl.load_workbook(workbook)["2026"]
    assert ws.cell(5, PRODUCTS.index("Jet") + 2).value == value("Jet", 4)
    assert openpyxl.load_workbook(workbook).calculation.fullCalcOnLoad


def test_update_reports_revisions(workbook: Path, tmp_path: Path) -> None:
    """Changed values in filled months are written and reported."""
    result = update_workbook(
        read_dgeg_file(write_dgeg(tmp_path / "f.xlsx", 2026, 2, shift=0.5)), workbook
    )
    assert result.added == ()
    assert len(result.revisions) == 2 * len(PRODUCTS)
    first = result.revisions[0]
    assert (first.month, first.product, first.old, first.new) == (
        1,
        "Butano",
        101.0,
        101.5,
    )


def test_update_nothing_new(workbook: Path, tmp_path: Path) -> None:
    """When nothing changed, the workbook is not rewritten."""
    before = workbook.read_bytes()
    result = update_workbook(
        read_dgeg_file(write_dgeg(tmp_path / "f.xlsx", 2026, 2)), workbook
    )
    assert (result.added, result.revisions, result.saved) == ((), (), False)
    assert workbook.read_bytes() == before


def test_update_creates_new_year(workbook: Path, tmp_path: Path) -> None:
    """A new year gets its own sheet with the product header."""
    result = update_workbook(
        read_dgeg_file(write_dgeg(tmp_path / "f.xlsx", 2027, 1)), workbook
    )
    assert result.added == ("Janeiro",)
    ws = openpyxl.load_workbook(workbook)["2027"]
    assert [ws.cell(1, c).value for c in range(2, 34)] == PRODUCTS
    assert ws.cell(13, 1).value == "Dezembro"


def test_update_missing_column(tmp_path: Path) -> None:
    """A sheet without a column the DGEG file has is an error."""
    path = tmp_path / "wb.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "2026"
    ws.append([None, "Butano"])
    ws.append(["Janeiro"])
    wb.save(path)
    with pytest.raises(ValueError, match="no column for"):
        update_workbook(read_dgeg_file(write_dgeg(tmp_path / "f.xlsx", 2026, 1)), path)


# --------------------------------------------------------------------------- #
# Reading the workbook and building outputs
# --------------------------------------------------------------------------- #


def test_read_panel_evaluates_formulas(formula_workbook: Path) -> None:
    """Formulas, text numbers and blanks are read like Excel would show them."""
    panel = read_panel(formula_workbook).set_index(["product", "year", "month"])[
        "consumption"
    ]
    assert panel["Total", 1995, 1] == 3
    assert panel["Net", 1995, 1] == 4  # SUM(B:D) - C = 1 + 2 + 3 - 2
    assert panel["Total", 1995, 2] == 3  # blank counts as zero
    assert panel["Net", 1995, 2] == 6  # a formula that references a formula
    assert panel["Text", 1995, 1] == 1234
    assert pd.isna(panel["Text", 1995, 2])  # "." is missing
    assert pd.isna(panel["B", 1995, 2])  # listed but blank keeps its row
    assert (1995, 3) not in {(y, m) for _, y, m in panel.index}  # unpublished
    assert panel["A", 1996, 1] == 5


def test_read_panel_unsupported_formula(tmp_path: Path) -> None:
    """Formulas other than sums and differences are refused."""
    wb = openpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "1995"
    ws.append([None, "A", "B"])
    ws.append(["Janeiro", 1, "=AVERAGE(B2:B2)"])
    wb.save(tmp_path / "wb.xlsx")
    with pytest.raises(ValueError, match="unsupported formula"):
        read_panel(tmp_path / "wb.xlsx")


def test_build_writes_outputs(workbook: Path, tmp_path: Path) -> None:
    """Build writes the panel, wide and time-series files."""
    result = build(workbook, tmp_path / "out")
    assert (result.first, result.last, result.months) == ("2025m1", "2026m2", 14)
    assert result.products == len(PRODUCTS)
    assert result.panel_rows == 14 * len(PRODUCTS)
    assert [p.name for p in result.files] == list(OUTPUT_FILES)

    panel = pd.read_csv(
        tmp_path / "out" / "oil_products_panel.csv", encoding="utf-8-sig"
    )
    assert list(panel.columns) == [
        "period",
        "date_monthly",
        "year",
        "month",
        "product_id",
        "product",
        "consumption",
    ]
    assert (
        panel.loc[panel["product"] == "Asfaltos", "product_id"].iloc[0] == 1
    )  # alphabetical

    ts = pd.read_csv(
        tmp_path / "out" / "oil_products_timeseries.csv", encoding="utf-8-sig"
    )
    assert ts["period"].tolist()[:2] == ["2025m1", "2025m2"]
    assert ts.loc[13, "Jet"] == value("Jet", 2)

    wide = pd.read_excel(tmp_path / "out" / "oil_products_wide.xlsx")
    assert list(wide.columns[:2]) == ["period", "date_monthly"]
    assert len(wide) == 14


def test_build_fills_gaps_in_index(tmp_path: Path) -> None:
    """A month with no data inside the range still gets a row."""
    path = write_workbook(tmp_path / "wb.xlsx", {2024: 1, 2025: 1})
    result = build(path, tmp_path / "out")
    assert result.months == 13  # 2024m1 .. 2025m1
    ts = pd.read_csv(
        tmp_path / "out" / "oil_products_timeseries.csv", encoding="utf-8-sig"
    )
    assert ts.loc[1, "period"] == "2024m2"
    assert ts[PRODUCTS].iloc[1].isna().all()


def test_build_also_to(workbook: Path, tmp_path: Path) -> None:
    """Extra folders receive the outputs and the workbook."""
    extra = tmp_path / "extra"
    extra.mkdir()
    result = build(workbook, tmp_path / "out", also_to=[extra])
    assert len(result.files) == 2 * len(OUTPUT_FILES) + 1
    for name in OUTPUT_FILES:
        assert (extra / name).read_bytes() == (tmp_path / "out" / name).read_bytes()
    assert (extra / workbook.name).read_bytes() == workbook.read_bytes()


def test_build_matches_committed_outputs(tmp_path: Path) -> None:
    """Rebuilding the repository's workbook reproduces the committed CSVs."""
    build(ROOT / "data" / "processed" / "dataset_fuel_sales.xlsx", tmp_path)
    for name in (
        "oil_products_panel.csv",
        "oil_products_wide.csv",
        "oil_products_timeseries.csv",
    ):
        fresh = pd.read_csv(tmp_path / name, encoding="utf-8-sig")
        committed = pd.read_csv(ROOT / "outputs" / name, encoding="utf-8-sig")
        pd.testing.assert_frame_equal(fresh, committed, check_exact=False, rtol=1e-12)


# --------------------------------------------------------------------------- #
# Command line
# --------------------------------------------------------------------------- #


def run_cli(*args: str) -> str:
    """Run a CLI command through Fire and return its output."""
    result = fire.Fire(core._Cli, command=list(args))
    assert isinstance(result, str)
    return result


def test_cli_build(workbook: Path, tmp_path: Path) -> None:
    """``build`` summarises what it wrote, copying to ``--also-to``."""
    extra = tmp_path / "extra"
    extra.mkdir()
    out = run_cli(
        "build",
        "--workbook",
        str(workbook),
        "--outputs",
        str(tmp_path / "out"),
        "--also-to",
        str(extra),
    )
    assert "2025m1 -> 2026m2 (14 months)" in out
    assert "11 files written" in out
    assert (extra / "oil_products_panel.csv").exists()


def test_cli_update_with_file(workbook: Path, tmp_path: Path) -> None:
    """``update --file`` lists added months and revisions."""
    path = write_dgeg(tmp_path / "f.xlsx", 2026, 3, shift=0.0)
    wb = openpyxl.load_workbook(workbook)
    wb["2026"].cell(2, 2, 1.0)  # change January Butano so DGEG revises it
    wb.save(workbook)
    out = run_cli("update", "--file", str(path), "--workbook", str(workbook))
    assert "2026: added Março; 1 revised value(s); workbook saved" in out
    assert "Janeiro Butano: 1.0 -> 101" in out


def test_cli_refresh_downloads(
    monkeypatch: pytest.MonkeyPatch, workbook: Path, tmp_path: Path
) -> None:
    """``refresh`` downloads the newest file, updates and rebuilds."""
    dgeg = write_dgeg(tmp_path / "dgeg-omn-2026-03.xlsx", 2026, 3)
    seen: list[str] = []

    def fake_download(
        dest_dir: str | Path = core.DEFAULT_RAW_DIR,
        session: requests.Session | None = None,
    ) -> Path:
        seen.append(str(dest_dir))
        return dgeg

    monkeypatch.setattr(core, "download_latest", fake_download)
    out = run_cli(
        "refresh",
        "--workbook",
        str(workbook),
        "--outputs",
        str(tmp_path / "out"),
        "--dest",
        "raw",
    )
    assert seen == ["raw"]
    assert "added Março" in out
    assert "2025m1 -> 2026m3" in out


def test_cli_download(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    """``download`` prints where the file was saved."""
    monkeypatch.setattr(
        core,
        "_get",
        lambda url, _session=None: PAGE.encode() if url == core.PAGE_URL else b"x",
    )
    out = run_cli("download", "--dest", str(tmp_path))
    assert out.endswith("dgeg-omn-2026-08.xlsx")


def test_main_reports_errors(
    monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]
) -> None:
    """An expected error prints one line and exits with code 1."""
    monkeypatch.setattr(
        sys, "argv", ["dgeg-fuelsales", "update", "--file", "missing.xlsx"]
    )
    with pytest.raises(SystemExit) as excinfo:
        core.main()
    assert excinfo.value.code == 1
    assert capsys.readouterr().err.startswith("Error:")


# --------------------------------------------------------------------------- #
# Packaging metadata
# --------------------------------------------------------------------------- #


def _file_version(path: Path, pattern: str) -> str:
    """Return the first version captured by ``pattern`` in ``path``."""
    import re

    match = re.search(pattern, path.read_text(encoding="utf-8"), re.MULTILINE)
    assert match is not None, f"no version found in {path}"
    return match.group(1)


def test_pyproject_and_citation_versions_agree() -> None:
    """The pyproject.toml and CITATION.cff versions must stay in lockstep."""
    pyproject = _file_version(ROOT / "pyproject.toml", r'^version = "([^"]+)"')
    citation = _file_version(ROOT / "CITATION.cff", r'^version:\s*"?([^"\s]+)"?')
    assert pyproject == citation
