# dgeg-fuelsales

Monthly sales of 50 oil products in Portugal, in tonnes, from January 1995,
compiled from the statistics published by DGEG (Direção-Geral de Energia e
Geologia). The dataset comes as a long panel and as wide monthly time series
(CSV and Excel), with a small, typed Python package that downloads DGEG's
newest monthly file, adds it to the source workbook and rebuilds every output.

```python
from dgeg_fuelsales import build, download_latest, read_dgeg_file, update_workbook

update_workbook(read_dgeg_file(download_latest()))  # add DGEG's new months
build()  # rebuild outputs/
```

Or, from the repository root:

```bash
uv run dgeg-fuelsales refresh
```

```{toctree}
:maxdepth: 2
:caption: Guide

installation
usage
data
```

```{toctree}
:maxdepth: 2
:caption: Reference

api
changelog
```
