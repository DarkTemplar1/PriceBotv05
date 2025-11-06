#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from __future__ import annotations

import argparse
import math
import re
from pathlib import Path

import numpy as np
import pandas as pd

# ========== helpers: numbers/text ==========

def _coerce_num_series(s: pd.Series) -> pd.Series:
    s = s.astype(str).str.replace("\u00a0", " ", regex=False).str.replace(" ", "", regex=False)
    s = s.str.replace(",", ".", regex=False)
    s = s.str.replace(r"[^0-9\.\-]", "", regex=True)
    return pd.to_numeric(s, errors="coerce")

def _clean_obszar_to_float(val) -> float | None:
    s = str(val or "").strip()
    if not s:
        return None
    s = s.replace("\u00a0", " ")
    s = re.sub(r"[^\d,\.]", "", s).replace(",", ".")
    try:
        return float(s)
    except Exception:
        return None

def _safe_str(x) -> str:
    if x is None:
        return ""
    if isinstance(x, float) and (math.isnan(x) or x != x):
        return ""
    return str(x)

_DEACCENT_TABLE = str.maketrans("ąćęłńóśźżĄĆĘŁŃÓŚŹŻ", "acelnoszzACELNOSZZ")

def _deaccent(s) -> str:
    return _safe_str(s).translate(_DEACCENT_TABLE)

def _norm_txt(v: str) -> str:
    v = _safe_str(v).replace("\u00a0", " ")
    v = re.sub(r"\s+", " ", v).strip().casefold()
    return v

def _street_name_only(s: str) -> str:
    """Usuń numer budynku z ulicy – zostaw samą nazwę ulicy."""
    txt = _safe_str(s).strip()
    m = re.match(r"^\s*([^\d,;]+?)(?:\s+\d.*)?\s*$", txt)
    name = m.group(1) if m else txt
    return re.sub(r"\s+", " ", name).strip()

# ========== TAK/NIE ==========

def _norm_yes_no(v: str) -> str:
    s = ("" if v is None else str(v)).strip().casefold()
    if s in {"", "brak", "false", "0"}: return "NIE"
    if s in {"w", "t", "tak", "yes", "y", "true", "1"}: return "TAK"
    if s in {"n", "nie", "no"}: return "NIE"
    if s[:1] in {"t", "y", "w"}: return "TAK"
    if s[:1] in {"n"}: return "NIE"
    return "NIE"

# ========== czytanie ofert ==========

def _pick_cols(df: pd.DataFrame) -> pd.DataFrame:
    want = {
        "cena": None, "cena_za_metr": None, "metry": None,
        "wojewodztwo": None, "powiat": None, "gmina": None, "miejscowosc": None, "dzielnica": None, "ulica": None,
        "Czy udziały?": None,
    }
    def canon(c: str) -> str:
        c = str(c).strip().lower()
        tr = str.maketrans("łśóżąęńć", "lsozaenc")
        c = c.translate(tr).replace("²","2")
        c = re.sub(r"[\s\-]+","_",c)
        return c
    cand = {canon(c): c for c in df.columns}

    for pat in ["cena_za_metr","cena_za_m2","cena_m2","cena_metr"]:
        if pat in cand: want["cena_za_metr"] = cand[pat]; break
    for pat in ["cena","cena_calkowita","cena_pln"]:
        if pat in cand: want["cena"] = cand[pat]; break
    for pat in ["metry","metraz","powierzchnia","powierzchnia_m2"]:
        if pat in cand: want["metry"] = cand[pat]; break
    for key in ["wojewodztwo","powiat","gmina","miejscowosc","dzielnica","ulica"]:
        if key in cand: want[key] = cand[key]
    for pat in ["czy_udzialy?","czy_udzialy","udzialy"]:
        if pat in cand: want["Czy udziały?"] = cand[pat]; break

    out = pd.DataFrame()
    for k in want:
        src = want[k]
        out[k] = df[src] if (src is not None and src in df.columns) else np.nan
    return out

def read_offers_excel(path: Path) -> pd.DataFrame:
    xl = pd.ExcelFile(path, engine="openpyxl")
    # wybór pierwszej sensownej zakładki
    sheet = None
    for s in xl.sheet_names:
        if len(xl.parse(s, nrows=0).columns) >= 2:
            sheet = s; break
    if sheet is None: sheet = xl.sheet_names[0]

    df = xl.parse(sheet)
    df = _pick_cols(df)

    df["cena"]         = _coerce_num_series(df["cena"])
    df["cena_za_metr"] = _coerce_num_series(df["cena_za_metr"])
    df["metry"]        = _coerce_num_series(df["metry"])
    df["Czy udziały?"] = df["Czy udziały?"].map(_norm_yes_no)

    for a in ["wojewodztwo","powiat","gmina","miejscowosc","dzielnica","ulica"]:
        df[a] = df[a].astype(str).fillna("").map(lambda x: _norm_txt(_deaccent(x)))

    return df

# ========== metryka i filtry ==========

def ensure_per_m2(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    per = df["cena_za_metr"].copy()
    miss = per.isna() | (per <= 0)
    if "cena" in df.columns:
        alt = pd.to_numeric(df["cena"], errors="coerce") / pd.to_numeric(df["metry"], errors="coerce")
        per = np.where(miss, alt, per)
    df["per_m2"] = pd.to_numeric(per, errors="coerce")
    # zdrowy zakres
    df = df.loc[df["per_m2"].between(1000, 50000)]
    return df

def iqr_filter(df: pd.DataFrame, col: str, k: float = 1.5) -> pd.DataFrame:
    x = pd.to_numeric(df[col], errors="coerce").replace([np.inf, -np.inf], np.nan)
    if x.notna().sum() < 4:
        return df.copy()
    q1, q3 = x.quantile(0.25), x.quantile(0.75)
    iqr = q3 - q1
    lo, hi = q1 - k*iqr, q3 + k*iqr
    return df.loc[x.between(lo, hi)].copy()

# dobór tolerancji wg ludności
BIG_200K = {
    "warszawa","krakow","lodz","wroclaw","poznan","gdansk","szczecin","bydgoszcz",
    "lublin","bialystok","katowice","gdynia"
}
MID_50_200K = {
    "czestochowa","radom","torun","kielce","rzeszow","gliwice","zabrze","olsztyn",
    "bielsko-biala","bytom","rybnik","ruda slaska","opole","tychy","plock","elblag",
    "gorzow wielkopolski","dabrowa gornicza","tarnow","chorzow","koszalin","kalisz",
    "legnica","grudziadz","slupsk","jaworzno","jastrzebie-zdroj","nowy sacz","jelenia gora",
    "konin","piotrkow trybunalski","inowroclaw","lubin","ostrowiec swietokrzyski",
    "siemianowice slaskie","ostroleka","kedzierzyn-kozle"
}

def tolerance_for_city(city: str) -> float:
    c = _norm_txt(_deaccent(city))
    if c in BIG_200K:
        return 5.0
    if c in MID_50_200K:
        return 10.0
    return 20.0

def progressive_address_filter(offers: pd.DataFrame, rp_row: pd.Series, base_mask: pd.Series, min_cnt: int = 5):
    ulica = _norm_txt(_deaccent(_street_name_only(rp_row.get("Ulica", ""))))
    dziel = _norm_txt(_deaccent(rp_row.get("Dzielnica", "")))
    miejsc = _norm_txt(_deaccent(rp_row.get("Miejscowość", "")))

    if ulica:
        m = base_mask & (offers["ulica"] == ulica)
        df = offers.loc[m].copy()
        if len(df) >= min_cnt:
            return df, "ulica"
    if dziel:
        m = base_mask & (offers["dzielnica"] == dziel)
        df = offers.loc[m].copy()
        if len(df) >= min_cnt:
            return df, "dzielnica"
    if miejsc:
        m = base_mask & (offers["miejscowosc"] == miejsc)
        df = offers.loc[m].copy()
        if len(df) >= min_cnt:
            return df, "miejscowosc"
    return offers.loc[base_mask].copy(), "brak"

# ========== raport (Excel) ==========

COL_MEAN_M2     = "Średnia cena za m² (z bazy)"
COL_MEAN_M2_ADJ = "Średnia skorygowana cena za m² (z bazy)"
COL_PROP_VALUE  = "Statystyczna wartość nieruchomości"

def pick_report_sheet(xlsx: Path) -> str:
    xl = pd.ExcelFile(xlsx, engine="openpyxl")
    return "raport" if "raport" in xl.sheet_names else xl.sheet_names[0]

def _fmt_pl(x) -> str:
    if x is None or (isinstance(x, float) and math.isnan(x)):
        return "-"
    return f"{x:,.0f}".replace(",", " ").replace(".", ",")

def format_price_per_m2(x) -> str:
    s = _fmt_pl(x)
    return "-" if s == "-" else f"{s} zł/m²"

def format_currency(x) -> str:
    s = _fmt_pl(x)
    return "-" if s == "-" else f"{s} zł"

def _try_save_excel(path: Path, df: pd.DataFrame, sheet: str) -> Path:
    """Zapisz podaną ramkę do istniejącego pliku w tej samej zakładce."""
    try:
        with pd.ExcelWriter(path, engine="openpyxl", mode="w") as xw:
            df.to_excel(xw, sheet_name=sheet, index=False)
        return path
    except PermissionError:
        # otwarty plik – zapisz kopię
        alt = path.with_name(f"{path.stem} (wyniki){path.suffix}")
        with pd.ExcelWriter(alt, engine="openpyxl", mode="w") as xw:
            df.to_excel(xw, sheet_name=sheet, index=False)
        print(f"[UWAGA] Plik był otwarty – zapisano kopię: {alt}")
        return alt

# ========== główna pętla ==========

def run_all(raport_xlsx: Path, offers_xlsx: Path) -> None:
    sheet = pick_report_sheet(raport_xlsx)
    df_rp = pd.read_excel(raport_xlsx, sheet_name=sheet, engine="openpyxl", dtype=str)

    offers = read_offers_excel(offers_xlsx)
    if offers.empty:
        print("[BŁĄD] Plik z ofertami nie zawiera danych.")
        return

    # dopilnuj, by kolumny wynikowe istniały
    for col in (COL_MEAN_M2, COL_MEAN_M2_ADJ, COL_PROP_VALUE):
        if col not in df_rp.columns:
            df_rp[col] = ""

    # przejście przez wiersze od 2. (index 1)
    total_rows = len(df_rp) - 1
    ok_rows = 0

    for idx in range(1, len(df_rp)):
        row = df_rp.iloc[idx]

        # koniec, jeśli całkowicie pusty wiersz (brak obszaru i brak adresu)
        area = _clean_obszar_to_float(row.get("Obszar", ""))
        addr_parts = [_safe_str(row.get(k,"")).strip() for k in ["Województwo","Powiat","Gmina","Miejscowość","Dzielnica","Ulica"]]
        if not any(addr_parts) and area is None:
            print(f"[{idx}] pusty wiersz — zatrzymuję.")
            break
        if area is None:
            print(f"[{idx}] pominięty (brak 'Obszar').")
            continue

        # dobór ±m² wg miasta
        tol = tolerance_for_city(row.get("Miejscowość", ""))

        # filtr bazowy: 'udzialy = NIE' + metraż
        m_area = offers["metry"].between(area - tol, area + tol)
        m_owner = (offers["Czy udziały?"].astype(str).str.upper() == "NIE")
        base_mask = m_owner & m_area

        # progresywny adres: ulica -> dzielnica -> miejscowość
        fdf, level = progressive_address_filter(offers, row, base_mask, min_cnt=5)

        # metryka i IQR
        fdf = ensure_per_m2(fdf)
        fdf_iqr = iqr_filter(fdf, "per_m2", k=1.5)

        mean_m2_base = float(fdf_iqr["per_m2"].mean()) if not fdf_iqr.empty else (float(fdf["per_m2"].mean()) if not fdf.empty else float("nan"))
        mean_m2_adj  = mean_m2_base * 0.85 if mean_m2_base == mean_m2_base else float("nan")
        prop_value   = mean_m2_adj * area if mean_m2_adj == mean_m2_adj else float("nan")

        # zapis do raportu (sformatowane)
        df_rp.loc[idx, COL_MEAN_M2]     = format_price_per_m2(mean_m2_base)
        df_rp.loc[idx, COL_MEAN_M2_ADJ] = format_price_per_m2(mean_m2_adj)
        df_rp.loc[idx, COL_PROP_VALUE]  = format_currency(prop_value)

        def _addr_preview() -> str:
            parts = []
            for k in ["Województwo","Powiat","Gmina","Miejscowość","Dzielnica","Ulica"]:
                v = _safe_str(row.get(k,"")).strip()
                if v and v.lower() not in {"nan","none"}:
                    if k == "Ulica":
                        v = _street_name_only(v)
                    parts.append(v)
            return " / ".join(parts) if parts else "[brak adresu]"

        print(f"[{idx}] {_addr_preview()} | {area:.2f} m² | Ofert po filtrach: {len(fdf)} (adres: {level}, ±{tol})")
        print(f"     Średnia cena za m² (z bazy): {format_price_per_m2(mean_m2_base)}")
        print(f"     Średnia skorygowana cena za m² (0.85x): {format_price_per_m2(mean_m2_adj)}")
        print(f"     Statystyczna wartość nieruchomości: {format_currency(prop_value)}")

        if len(fdf) > 0:
            ok_rows += 1

    # zapis arkusza z wynikami
    saved_path = _try_save_excel(raport_xlsx, df_rp, sheet)
    print(f"Gotowe. Zapisano wyniki do: {saved_path}")
    print(f"Podsumowanie: wiersze z wynikiem {ok_rows} / {total_rows}")

# ========== CLI ==========

def main():
    ap = argparse.ArgumentParser(description="Automatyczne liczenie wartości m² z progresywnym filtrem adresu i zapisem do raportu.")
    ap.add_argument("--raport", required=True, type=Path, help="Plik raportu (Excel) — arkusz 'raport' lub pierwszy.")
    ap.add_argument("--oferty", required=True, type=Path, help="Plik z ofertami (Excel) — jedno źródło.")
    args = ap.parse_args()

    if not args.raport.exists():
        print(f"[BŁĄD] Nie znaleziono pliku raportu: {args.raport}"); return
    if not args.oferty.exists():
        print(f"[BŁĄD] Nie znaleziono pliku z ofertami: {args.oferty}"); return

    run_all(args.raport.resolve(), args.oferty.resolve())

if __name__ == "__main__":
    main()
