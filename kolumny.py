# -*- coding: utf-8 -*-
import sys
import csv
import argparse
from pathlib import Path
from typing import List
import tkinter as tk
from tkinter import filedialog, messagebox
from openpyxl import Workbook, load_workbook

RAPORT_SHEET = "raport"
RAPORT_ODF = "raport_odfiltrowane"

REQ_COLS: List[str] = [
    "Nr KW","Typ Księgi","Stan Księgi","Województwo","Powiat","Gmina",
    "Miejscowość","Dzielnica","Położenie","Nr działek po średniku","Obręb po średniku",
    "Ulica","Sposób korzystania","Obszar","Ulica(dla budynku)",
    "przeznaczenie (dla budynku)","Ulica(dla lokalu)","Nr budynku( dla lokalu)",
    "Przeznaczenie (dla lokalu)","Cały adres (dla lokalu)","Czy udziały?"
]

VALUE_COLS: List[str] = [
    "Średnia cena za m2 ( z bazy)",
    "Średnia skorygowana cena za m2",
    "Statyczna wartość nieruchomości"
]

WYNIKI_HEADER: List[str] = [
    "cena","cena_za_metr","metry","liczba_pokoi","pietro","rynek","rok_budowy","material",
    "wojewodztwo","powiat","gmina","miejscowosc","dzielnica","ulica","link",
]

SUPPORTED = {".xlsx", ".xlsm"}

VOIVODESHIPS_LABEL_SLUG: list[tuple[str, str]] = [
    ("Dolnośląskie", "dolnoslaskie"),
    ("Kujawsko-Pomorskie", "kujawsko-pomorskie"),
    ("Lubelskie", "lubelskie"),
    ("Lubuskie", "lubuskie"),
    ("Łódzkie", "lodzkie"),
    ("Małopolskie", "malopolskie"),
    ("Mazowieckie", "mazowieckie"),
    ("Opolskie", "opolskie"),
    ("Podkarpackie", "podkarpackie"),
    ("Podlaskie", "podlaskie"),
    ("Pomorskie", "pomorskie"),
    ("Śląskie", "slaskie"),
    ("Świętokrzyskie", "swietokrzyskie"),
    ("Warmińsko-Mazurskie", "warminsko-mazurskie"),
    ("Wielkopolskie", "wielkopolskie"),
    ("Zachodniopomorskie", "zachodniopomorskie"),
]

# --------------------- Desktop/Pulpit ---------------------
def _detect_desktop() -> Path:
    home = Path.home()
    for name in ("Desktop", "Pulpit"):
        p = home / name
        if p.exists():
            return p
    return home

def ensure_base_dirs(base_override: Path | None = None) -> Path:
    """
    Zwraca katalog bazowy 'baza danych' i upewnia się, że istnieją:
      • <base>/linki
      • <base>/województwa
      • <base>/timing.csv (z nagłówkiem)
    """
    if base_override:
        base = Path(base_override)
    else:
        base = _detect_desktop() / "baza danych"

    (base / "linki").mkdir(parents=True, exist_ok=True)
    (base / "województwa").mkdir(parents=True, exist_ok=True)

    timing = base / "timing.csv"
    if not timing.exists():
        with timing.open("w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f)
            w.writerow(["Województwo", "Stan pobierania"])
    return base
# ---------------------------------------------------------

# --------------------- CSV helpery -----------------------
def _ensure_csv(path: Path, header: List[str]) -> bool:
    if path.exists():
        return False
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        if header:
            w.writerow(header)
    return True

def create_voivodeship_csvs(base: Path) -> dict:
    created = {"linki": 0, "województwa": 0}
    linki_dir = base / "linki"
    woj_dir = base / "województwa"
    for label, _slug in VOIVODESHIPS_LABEL_SLUG:
        if _ensure_csv(linki_dir / f"{label}.csv", ["link"]):
            created["linki"] += 1
        if _ensure_csv(woj_dir / f"{label}.csv", WYNIKI_HEADER):
            created["województwa"] += 1
    return created
# ---------------------------------------------------------

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Tworzenie struktury plików PriceBot (kolumny.py).")
    parser.add_argument("--base-dir", help="Gdzie utworzyć 'baza danych' (domyślnie: Desktop/Pulpit).")
    args = parser.parse_args()

    base_override = Path(args.base_dir) if args.base_dir else None
    base = ensure_base_dirs(base_override)
    created = create_voivodeship_csvs(base)

    # krótki komunikat w konsoli
    print(f"[kolumny] Baza: {base}")
    print(f"[kolumny] Utworzone: linki={created['linki']}, województwa={created['województwa']}")
