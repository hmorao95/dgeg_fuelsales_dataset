# CLAUDE.md

This project's agent instructions live in [AGENTS.md](AGENTS.md). Read and follow
them.

@AGENTS.md

Key reminders for this repo:

- Run the full quality gate before committing (see AGENTS.md): ruff,
  ruff-format, codespell, deptry, interrogate, mypy, pytest.
- Branch, open a PR, squash-merge when CI is green.
- Commit as the repository owner.
- `outputs/` is committed and always rebuilt from the workbook, never edited by
  hand; rebuild it in the same commit as any workbook change.
- Releases bump the version in `pyproject.toml`, `CITATION.cff`, the README
  citation and `uv.lock` together.
