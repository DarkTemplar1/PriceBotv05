#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from __future__ import annotations

import argparse
from pathlib import Path
import re
import sys
import pandas as pd

# ===================== ustawienia =====================
LOC_COLS = ["Województwo", "Powiat", "Gmina", "Miejscowość", "Dzielnica"]
OUT_COLS = [
    "Średnia cena za m² (z bazy)",
    "Średnia skorygowana cena za m² (z bazy)",
    "Statystyczna wartość nieruchomości",
]

# domyślna lokalizacja TERYT
def _default_teryt() -> Path:
    for p in [
        Path.home() / "Pulpit" / "baza danych" / "TERYT.xlsx",
        Path.home() / "Desktop" / "baza danych" / "TERYT.xlsx",
    ]:
        if p.exists():
            return p
    return Path("TERYT.xlsx")


# ===================== normalizacja =====================
_deacc = str.maketrans(
    "ĄĆĘŁŃÓŚŹŻąćęłńóśźż",
    "ACELNOSZZacelnoszz",
)


def _norm(s) -> str:
    # szybka, bezpieczna normalizacja (obsługuje pd.NA)
    if s is pd.NA or s is None:
        return ""
    s = str(s)
    s = s.replace("\u00a0", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def _canon(s) -> str:
    s = _norm(s)
    s = s.translate(_deacc).lower()
    s = re.sub(r"[^a-z0-9\s\-]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


# stare / błędne zapisy województw -> współczesne
OLD_WOJ_MAP = {
    # warianty pisowni / literówki / case
    "malopolskie": "MAŁOPOLSKIE",
    "mazowieckie": "MAZOWIECKIE",
    "podlaskie": "PODLASKIE",
    "pomorskie": "POMORSKIE",
    "zachodniopomorskie": "ZACHODNIOPOMORSKIE",
    "dolnoslaskie": "DOLNOŚLĄSKIE",
    "kujawskopomorskie": "KUJAWSKO-POMORSKIE",
    "lubelskie": "LUBELSKIE",
    "lubuskie": "LUBUSKIE",
    "lodzkie": "ŁÓDZKIE",
    "opolskie": "OPOLSKIE",
    "podkarpackie": "PODKARPACKIE",
    "slaskie": "ŚLĄSKIE",
    "swietokrzyskie": "ŚWIĘTOKRZYSKIE",
    "warminsko-mazurskie": "WARMIŃSKO-MAZURSKIE",
    "wielkopolskie": "WIELKOPOLSKIE",
    # częste myłki dawnymi „puckie” etc. — normalizujemy do powiatów niżej
}

# heurystyka: czy tekst wygląda na adres (numery, „ul”, itp.)
ADDR_PAT = re.compile(
    r"\b(ul\.?|ulica|al\.?|aleja|plac|pl\.|os\.?|osiedle|rynek)\b|\d+\s*[a-zA-Z]?(?:\s*/\s*\d+[a-zA-Z]?)?",
    re.IGNORECASE,
)


def _looks_like_address(val) -> bool:
    s = _norm(val)
    if not s:
        return False
    if ADDR_PAT.search(s):
        return True
    # np. sama "DWORCOWA 49"
    if re.search(r"[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż]\s*\d", s):
        return True
    return False


# ===================== TERYT: szybkie wczytanie =====================
def _read_teryt(teryt_xlsx: Path) -> pd.DataFrame:
    """
    Oczekujemy arkusza SIMC lub przygotowanego zbioru z kolumnami:
    'woj_name','pow_name','gmi_name','msc','dzielnica' (dzielnica opcjonalnie).
    Robimy sobie jedynie wewnętrzne kolumny kanoniczne na czas dopasowań.
    """
    try:
        xl = pd.ExcelFile(teryt_xlsx, engine="openpyxl")
    except Exception as e:
        raise SystemExit(f"[BŁĄD] Nie mogę otworzyć TERYT: {teryt_xlsx} ({e})")

    # wybierz arkusz zawierający SIMC/”miejscowości”
    sheet = None
    for s in xl.sheet_names:
        name = _canon(s)
        if any(k in name for k in ["simc", "miejscowosci", "miejscowości", "msc"]):
            sheet = s
            break
    if sheet is None:
        sheet = xl.sheet_names[0]

    df = xl.parse(sheet)

    # zmapuj nazwy kolumn na oczekiwane
    def _pick(colnames, *cands):
        lc = { _canon(c): c for c in colnames }
        for c in cands:
            if c in lc:
                return lc[c]
        return None

    woj = _pick(df.columns, "woj_name", "wojewodztwo", "województwo", "woj")
    powi = _pick(df.columns, "pow_name", "powiat", "pow")
    gmi = _pick(df.columns, "gmi_name", "gmina", "gm")
    msc = _pick(df.columns, "msc", "miejscowosc", "miejscowość", "miejscowosc_nazwa")
    dzl = _pick(df.columns, "dzielnica", "jednostka_ pomocnicza", "jedn_pom", "czesc")

    keep = [c for c in [woj, powi, gmi, msc, dzl] if c]
    df = df[keep].copy()

    # wewnętrzne kanony (tylko w RAM)
    df["_woj_c"] = df[woj].map(_canon)
    df["_pow_c"] = df[powi].map(_canon) if powi else ""
    df["_gmi_c"] = df[gmi].map(_canon) if gmi else ""
    df["_msc_c"] = df[msc].map(_canon)
    df["_dzl_c"] = df[dzl].map(_canon) if dzl else ""

    df.rename(
        columns={
            woj: "woj_name",
            powi: "pow_name" if powi else "pow_name",
            gmi: "gmi_name" if gmi else "gmi_name",
            msc: "msc",
            dzl: "dzielnica" if dzl else "dzielnica",
        },
        inplace=True,
    )

    # brakujące kolumny (jeśli w źródłowym TERYT ich nie ma)
    for c in ["pow_name", "gmi_name", "dzielnica"]:
        if c not in df.columns:
            df[c] = ""

    return df


# ===================== logika dopasowania =====================
def _fix_woj_name(raw_woj: str) -> str:
    c = _canon(raw_woj)
    fixed = OLD_WOJ_MAP.get(c)
    if fixed:
        return fixed
    # bezpośrednio tylko title-case, jeśli już poprawne
    s = _norm(raw_woj)
    return s.upper() if s else s


def _resolve_row_against_teryt(simc: pd.DataFrame, row: pd.Series):
    """
    Zwraca (status, fill_dict)
      status: 'ok' | 'brak' | 'ambiguous' | 'addr_in_loc'
      fill_dict: słownik wartości do uzupełnienia (Województwo/Powiat/Gmina/Miejscowość/Dzielnica)
    """
    # 1) jeśli w polach lokalizacji jest adres -> brak
    if any(_looks_like_address(row.get(c, "")) for c in LOC_COLS):
        return "addr_in_loc", {}

    woj = _fix_woj_name(row.get("Województwo", ""))
    powiat = _norm(row.get("Powiat", ""))
    gmina = _norm(row.get("Gmina", ""))
    miasto = _norm(row.get("Miejscowość", ""))
    dziel = _norm(row.get("Dzielnica", ""))

    # 2) kanony
    woj_c = _canon(woj)
    pow_c = _canon(powiat)
    gmi_c = _canon(gmina)
    msc_c = _canon(miasto)
    dzl_c = _canon(dziel)

    # bez miejscowości nie potwierdzimy — brak
    if not msc_c:
        return "brak", {}

    mask = (simc["_msc_c"] == msc_c)
    if woj_c:
        mask &= (simc["_woj_c"] == woj_c)
    if pow_c:
        mask &= (simc["_pow_c"] == pow_c)
    if gmi_c:
        mask &= (simc["_gmi_c"] == gmi_c)
    if dzl_c:
        mask &= (simc["_dzl_c"] == dzl_c)

    cand = simc.loc[mask]
    if len(cand) == 1:
        r = cand.iloc[0]
        fill = {
            "Województwo": r["woj_name"],
            "Powiat": r["pow_name"],
            "Gmina": r["gmi_name"],
            "Miejscowość": r["msc"],
        }
        # Dzielnica tylko jeśli SIMC ją ma i w raporcie brak lub pusta
        if "dzielnica" in cand.columns and (not dziel):
            fill["Dzielnica"] = r["dzielnica"]
        return "ok", fill

    if len(cand) == 0:
        # spróbuj tylko po miejscowości (często wystarcza)
        cand2 = simc.loc[simc["_msc_c"] == msc_c]
        if len(cand2) == 1:
            r = cand2.iloc[0]
            fill = {
                "Województwo": r["woj_name"],
                "Powiat": r["pow_name"],
                "Gmina": r["gmi_name"],
                "Miejscowość": r["msc"],
            }
            if "dzielnica" in cand2.columns and (not dziel):
                fill["Dzielnica"] = r["dzielnica"]
            return "ok", fill
        return "brak", {}
    else:
        return "ambiguous", {}


def _set_brak(df: pd.DataFrame, idx):
    for c in OUT_COLS:
        if c in df.columns:
            df.at[idx, c] = "brak adresu"


# ===================== główna funkcja =====================
def clean_file(input_xlsx: Path, teryt_xlsx: Path | None):
    if teryt_xlsx is None:
        teryt_xlsx = _default_teryt()
    simc = _read_teryt(teryt_xlsx)

    # wybór arkusza „raport”
    xl = pd.ExcelFile(input_xlsx, engine="openpyxl")
    sheet = "raport" if "raport" in xl.sheet_names else xl.sheet_names[0]
    df = xl.parse(sheet, dtype=object)

    # upewnij się, że kolumny istnieją
    for c in LOC_COLS:
        if c not in df.columns:
            df[c] = ""
    for c in OUT_COLS:
        if c not in df.columns:
            df[c] = ""

    # przetwarzanie wiersz po wierszu (szybkie – mało logiki na rekord)
    for i in range(len(df)):
        row = df.loc[i]

        status, fill = _resolve_row_against_teryt(simc, row)

        if status == "ok":
            # uzupełnij brakujące (nie nadpisujemy niepustych)
            for k, v in fill.items():
                if not _norm(df.at[i, k]):
                    df.at[i, k] = v
        else:
            # wpisz „brak adresu” w trzy kolumny wynikowe
            _set_brak(df, i)

        # zawsze podnieś województwo do poprawnego zapisu (mapa + upper)
        if _norm(df.at[i, "Województwo"]):
            df.at[i, "Województwo"] = _fix_woj_name(df.at[i, "Województwo"])

    # zapis – bez kolumn pomocniczych
    with pd.ExcelWriter(input_xlsx, engine="openpyxl", mode="w") as xw:
        df.to_excel(xw, sheet_name=sheet, index=False)


# ===================== CLI =====================
def main():
    ap = argparse.ArgumentParser(
        description="Czyszczenie nazw lokalizacji i walidacja w oparciu o TERYT (bez dopisywania kolumn pomocniczych)."
    )
    ap.add_argument("--in", dest="inp", required=True, type=Path, help="Plik raportu (Excel)")
    ap.add_argument("--teryt", dest="teryt", type=Path, default=None, help="Plik TERYT.xlsx (SIMC). Domyślnie: ~/Pulpit/baza danych/TERYT.xlsx")
    args = ap.parse_args()

    if not args.inp.exists():
        print(f"[BŁĄD] Nie znaleziono pliku: {args.inp}")
        sys.exit(1)

    clean_file(args.inp.resolve(), args.teryt.resolve() if args.teryt else None)
    print("OK – zapisano poprawiony plik.")


if __name__ == "__main__":
    main()
