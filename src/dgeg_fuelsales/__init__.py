"""
Public package interface for :mod:`dgeg_fuelsales`.

This package updates and builds the DGEG monthly oil-product sales dataset for
Portugal. The implementation lives in :mod:`dgeg_fuelsales.core`; the names
most callers need are re-exported here, so ``from dgeg_fuelsales import build``
is enough to rebuild the dataset.

Exports:
    download_latest: Download the newest DGEG monthly sales file.
    latest_file_url: Find the newest file linked from the DGEG page.
    read_dgeg_file: Parse a DGEG file into its published months.
    update_workbook: Write those months into the source workbook.
    read_panel: Read the workbook into a long product x month table.
    seasonally_adjust: Seasonally adjust each product's series (STL).
    build: Write the panel, wide, time-series and adjusted files.
    DgegFile, Revision, UpdateResult, BuildResult: Result types.
    DEFAULT_WORKBOOK, DEFAULT_RAW_DIR, DEFAULT_OUTPUTS: Default paths, relative
        to the repository root.
    MONTHS, OUTPUT_FILES: Month names and output file names.
    SA_MIN_MONTHS, SA_MAX_ANNUAL_GAP: Which series are seasonally adjusted.
    main: Entry point for the ``python-fire`` command-line interface.
"""

from __future__ import annotations

from dgeg_fuelsales.core import (
    DEFAULT_OUTPUTS,
    DEFAULT_RAW_DIR,
    DEFAULT_WORKBOOK,
    MONTHS,
    OUTPUT_FILES,
    SA_MAX_ANNUAL_GAP,
    SA_MIN_MONTHS,
    BuildResult,
    DgegFile,
    Revision,
    UpdateResult,
    build,
    download_latest,
    latest_file_url,
    main,
    read_dgeg_file,
    read_panel,
    seasonally_adjust,
    update_workbook,
)

__all__ = [
    "DEFAULT_OUTPUTS",
    "DEFAULT_RAW_DIR",
    "DEFAULT_WORKBOOK",
    "MONTHS",
    "OUTPUT_FILES",
    "SA_MAX_ANNUAL_GAP",
    "SA_MIN_MONTHS",
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
    "seasonally_adjust",
    "update_workbook",
]
