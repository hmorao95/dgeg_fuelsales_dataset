"""
Shared pytest fixtures for the whole suite.

The library and the CLI default to paths relative to the working directory
(``data/processed/...``, ``data/raw/dgeg``, ``outputs``). The autouse
:func:`workdir` fixture runs every test in a fresh temporary directory, so no
test that relies on those defaults can read or overwrite the repository's own
data.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

import pytest

if TYPE_CHECKING:
    from pathlib import Path


@pytest.fixture(autouse=True)
def workdir(
    tmp_path_factory: pytest.TempPathFactory, monkeypatch: pytest.MonkeyPatch
) -> Path:
    """
    Change into an empty per-test directory.

    Autouse, so isolation never depends on a test remembering to do it. A test
    that needs the directory can request this fixture by name.

    Args:
        tmp_path_factory (pytest.TempPathFactory): Source of the fresh directory;
            kept separate from ``tmp_path`` so a test's own files never land in
            the working directory by accident.
        monkeypatch (pytest.MonkeyPatch): Restores the working directory after
            the test.

    Returns:
        Path: The working directory (empty at the start of the test).
    """
    cwd = tmp_path_factory.mktemp("cwd")
    monkeypatch.chdir(cwd)
    return cwd
