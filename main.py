#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
PriceBot — uruchamia interfejs graficzny (selektor_csv.py) w tym samym procesie.
"""

from pathlib import Path

def run_gui():
    # Import lokalny (żeby PyInstaller na pewno dołączył moduł)
    import selektor_csv  # noqa
    if hasattr(selektor_csv, "App"):
        selektor_csv.App().mainloop()
    else:
        # Fallback: jeśli moduł jest skryptem
        import runpy
        here = Path(__file__).resolve().parent
        runpy.run_path(str(here / "selektor_csv.py"))

if __name__ == "__main__":
    run_gui()
