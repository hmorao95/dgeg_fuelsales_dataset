"""
Module entry point enabling ``python -m dgeg_fuelsales``.

Running the package as a module simply delegates to
:func:`dgeg_fuelsales.core.main`, which launches the Fire-generated
command-line interface. This mirrors the ``dgeg-fuelsales`` console script
installed via the project's entry points.
"""

from __future__ import annotations

from dgeg_fuelsales.core import main

if __name__ == "__main__":
    main()
