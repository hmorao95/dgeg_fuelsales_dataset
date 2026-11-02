# Installation

The data files need no installation: download them from the
[`outputs/`](https://github.com/hmorao95/dgeg_fuelsales_dataset/tree/main/outputs)
folder of the repository.

To update or rebuild the dataset you need Python 3.10 or newer and a clone of
the repository, since the package works on the repository's `data/` and
`outputs/` folders:

```bash
git clone https://github.com/hmorao95/dgeg_fuelsales_dataset.git
cd dgeg_fuelsales_dataset
uv sync
```

[uv](https://docs.astral.sh/uv/) creates a virtual environment with every
dependency, including [wrapkit](https://github.com/hmorao95/wrapkit), which is
installed from GitHub. Run commands through `uv run`:

```bash
uv run dgeg-fuelsales --help
```

No API key or account is needed: DGEG publishes the files openly.
