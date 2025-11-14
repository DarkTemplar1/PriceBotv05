#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from pathlib import Path
import os
import unicodedata
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import numpy as np
import subprocess, sys
import json  # <--- DODANE do zapisywania ustawień

APP_TITLE = "PriceBot"

# ---------- Helpers nazewnicze ----------
def _norm(s: str) -> str:
    return (s or "").strip().lower().replace(" ", "").replace("\xa0", "").replace("\t", "")

def _plain(s: str) -> str:
    s = (s or "").lower()
    s = "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))
    return s

def _find_col(cols, candidates):
    """Zwróć istniejącą kolumnę dopasowaną do listy kandydatów (po normalizacji / zawieraniu)."""
    norm_map = { _norm(c): c for c in cols }
    for cand in candidates:
        key = _norm(cand)
        if key in norm_map:
            return norm_map[key]
    for c in cols:
        if any(_norm(x) in _norm(c) for x in candidates):
            return c
    return None

def _trim_after_semicolon(val):
    if pd.isna(val): return ""
    s = str(val)
    if ";" in s: s = s.split(";", 1)[0].strip()
    return s

def _to_float_maybe(x):
    if pd.isna(x): return None
    s = str(x).replace(" ", "").replace("\xa0", "").replace(",", ".")
    s = "".join(ch for ch in s if (ch.isdigit() or ch == "." or ch == "-"))
    try:
        return float(s) if s else None
    except Exception:
        return None

# ---------- USTAWIENIA PODGLĄDU ----------
PREVIEW_SPEC = [
    ("Nr KW",        ["Nr KW", "nr_ksiegi", "nrksiegi", "nr księgi", "numer księgi"]),
    ("Województwo",  ["Województwo", "wojewodztwo", "woj"]),
    ("Powiat",       ["Powiat"]),
    ("Gmina",        ["Gmina"]),
    ("Miejscowość",  ["Miejscowość", "Miejscowosc", "Miasto"]),
    ("Dzielnica",    ["Dzielnica", "Osiedle"]),
    ("Ulica",        ["Ulica", "Ulica(dla budynku)", "Ulica(dla lokalu)"]),
    ("Obszar",       ["Obszar", "metry", "powierzchnia", "Nr działek po średniku", "Nr działek", "Obręb po średniku", "Obręb"]),
]
HIDDEN_PREVIEW_COLS = { _norm("Typ Księgi"), _norm("Stan Księgi") }

# ---------- Główna klasa ----------
class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_TITLE)
        self.minsize(900, 420)

        self.df: pd.DataFrame | None = None
        self.input_path: Path | None = None
        self.current_idx: int | None = None

        # Ustawienia
        self.input_file_var = tk.StringVar(value="")
        self.folder_var = tk.StringVar(value=str(Path.home()))      # baza: tu jest Polska.xlsx
        self.output_folder_var = tk.StringVar(value="")             # folder zapisu (Nr KW).xlsx
        self.margin_m2_var = tk.DoubleVar(value=5.0)                # okno +- m2
        self.margin_pct_var = tk.DoubleVar(value=0.0)               # okno +- % (jeśli >0, ma pierwszeństwo)

        # wczytaj zapisane ustawienia (np. ostatni folder bazowy)
        self._load_settings()

        # --- UI ---
        root = ttk.Frame(self, padding=10); root.pack(fill="both", expand=True)

        group_in = ttk.LabelFrame(root, text="Plik raportu (wejście)"); group_in.pack(fill="x")
        row_in = ttk.Frame(group_in); row_in.pack(fill="x", padx=8, pady=6)
        ttk.Entry(row_in, textvariable=self.input_file_var).pack(side="left", fill="x", expand=True)
        ttk.Button(row_in, text="Wybierz...", command=self.choose_input_file).pack(side="left", padx=(8, 0))
        ttk.Button(row_in, text="Czyszczenie Pliku", command=self.clean_input_file).pack(side="left", padx=(8, 0))

        group_base = ttk.LabelFrame(root, text="Miejsce tworzenia plików i folderów"); group_base.pack(fill="x", pady=(8, 0))
        row_base = ttk.Frame(group_base); row_base.pack(fill="x", padx=8, pady=6)
        ttk.Entry(row_base, textvariable=self.folder_var).pack(side="left", fill="x", expand=True)
        ttk.Button(row_base, text="Wybierz folder...", command=self.choose_base_folder).pack(side="left", padx=(8, 0))
        ttk.Button(row_base, text="Przygotowanie Aplikacji", command=self.prepare_app).pack(side="left", padx=(8, 0))

        group_db = ttk.LabelFrame(root, text="Baza danych"); group_db.pack(fill="x", pady=(8, 0))
        row_db = ttk.Frame(group_db); row_db.pack(fill="x", padx=8, pady=6)
        ttk.Button(row_db, text="Baza danych", command=self.run_bazadanych).pack(side="left")

        # ---------- FILTRY RAPORTU (NOWA SEKCJA) ----------
        group_filters = ttk.LabelFrame(root, text="Filtry raportu")
        group_filters.pack(fill="x", pady=(8, 0))

        row_filters = ttk.Frame(group_filters)
        row_filters.pack(fill="x", padx=8, pady=6)

        ttk.Label(row_filters, text="Filtr:").pack(side="left")

        self.filter_var = tk.StringVar(value="— brak —")
        self.filter_box = ttk.Combobox(
            row_filters,
            textvariable=self.filter_var,
            state="readonly",
            width=40,
            values=[
                "— brak —",
                "Jeden właściciel",
                "Lokal mieszkalny",
                "Jeden właściciel i Lokal mieszkalny",
                "Cofnij",
            ]
        )
        self.filter_box.pack(side="left", padx=10)
        self.filter_box.bind("<<ComboboxSelected>>", self.run_filter_script)

        group_out = ttk.LabelFrame(root, text="Folder zapisu wyników"); group_out.pack(fill="x", pady=(8, 0))
        row_out = ttk.Frame(group_out); row_out.pack(fill="x", padx=8, pady=6)
        ttk.Entry(row_out, textvariable=self.output_folder_var).pack(side="left", fill="x", expand=True)
        ttk.Button(row_out, text="Wybierz folder...", command=self.choose_output_folder).pack(side="left", padx=(8, 0))

        group_ctrl = ttk.LabelFrame(root, text="Parametry i sterowanie"); group_ctrl.pack(fill="x", pady=(8, 0))
        row_ctrl1 = ttk.Frame(group_ctrl); row_ctrl1.pack(fill="x", padx=8, pady=6)

        ttk.Label(row_ctrl1, text="Pomiary brzegowe metrażu:").pack(side="left")
        ttk.Label(row_ctrl1, text="± m²:").pack(side="left", padx=(8, 2))
        ttk.Spinbox(row_ctrl1, from_=0.0, to=200.0, increment=0.5, width=6, textvariable=self.margin_m2_var).pack(side="left")
        ttk.Label(row_ctrl1, text="- %:").pack(side="left", padx=(12, 2))
        ttk.Spinbox(row_ctrl1, from_=0.0, to=100.0, increment=0.5, width=6, textvariable=self.margin_pct_var).pack(side="left")

        ttk.Button(row_ctrl1, text="‹ Poprzedni", command=self.prev_row).pack(side="left", padx=(16, 0))
        ttk.Button(row_ctrl1, text="Następny ›", command=self.next_row).pack(side="left", padx=(6, 0))
        ttk.Button(row_ctrl1, text="Oblicz i zapisz ten wiersz", command=self.calc_and_save_row).pack(side="left", padx=(16, 0))
        ttk.Button(row_ctrl1, text="Automat", command=self.automate).pack(side="left", padx=(6, 0))

        group_preview = ttk.LabelFrame(root, text="Bieżący wiersz (podgląd)"); group_preview.pack(fill="both", expand=True, pady=(8, 0))
        self.preview_label = ttk.Label(group_preview, text="{Wybierz plik raportu}", anchor="w", justify="left")
        self.preview_label.pack(fill="both", expand=True, padx=8, pady=6)

    # ---------- zapisywanie / wczytywanie ustawień ----------
    def _settings_path(self) -> Path:
        """Plik z ustawieniami obok selektor_csv.py."""
        return Path(__file__).resolve().with_name("pricebot_settings.json")

    def _load_settings(self):
        try:
            cfg_path = self._settings_path()
            if cfg_path.exists():
                with cfg_path.open("r", encoding="utf-8") as f:
                    data = json.load(f)
                base = data.get("base_folder") or ""
                if base:
                    self.folder_var.set(base)
        except Exception:
            # nie krzycz, jak coś pójdzie nie tak – po prostu użyj domyślnego
            pass

    def _save_settings(self):
        try:
            cfg_path = self._settings_path()
            data = {
                "base_folder": self.folder_var.get() or "",
            }
            with cfg_path.open("w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception:
            # brak zapisu = brak pamiętania, ale aplikacja ma działać dalej
            pass

    # ---------- uruchamianie zewnętrznych skryptów ----------
    def _run_script(self, candidates: list[str], extra_args: list[str] | None = None):
        """
        Szuka skryptu w tym samym folderze co selektor_csv.py i uruchamia go w osobnym procesie.
        candidates: np. ["Automat.py", "automat.py"]
        """
        extra_args = extra_args or []
        here = Path(__file__).resolve().parent
        for name in candidates:
            script = here / name
            if script.exists():
                try:
                    subprocess.Popen(
                        [sys.executable, str(script), *extra_args],
                        cwd=str(here),
                        close_fds=(os.name != "nt"),
                        creationflags=(subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0)
                    )
                    messagebox.showinfo("Uruchomiono", f"Start: {script.name}")
                    return
                except Exception as e:
                    messagebox.showerror("Błąd uruchamiania", f"Nie udało się uruchomić {script.name}:\n{e}")
                    return
        messagebox.showerror("Brak pliku", f"Nie znaleziono żadnego ze skryptów: {', '.join(candidates)}")

    # ---------- GUI actions ----------
    def choose_input_file(self):
        path = filedialog.askopenfilename(
            title="Wybierz plik raportu (CSV/XLSX)",
            filetypes=[("CSV", "*.csv"), ("Excel", "*.xlsx;*.xls"), ("Wszystkie pliki", "*.*")]
        )
        if not path: return
        self.input_file_var.set(path)
        self.input_path = Path(path)
        self.load_dataframe(self.input_path)
        self.goto_row(0)

    def load_dataframe(self, path: Path):
        try:
            if path.suffix.lower() in (".xlsx", ".xls"):
                self.df = pd.read_excel(path)
            else:
                self.df = pd.read_csv(path, sep=None, engine="python")
        except Exception as e:
            messagebox.showerror("Błąd odczytu", f"Nie mogę wczytać pliku:\n{path}\n\n{e}")
            self.df = None; self.current_idx = None
            self.preview_label.config(text="{Błąd odczytu pliku}")

    def clean_input_file(self):
        """Uruchamia skrypt CzyszczenieAdresu.py, przekazując ścieżkę do aktualnie wybranego pliku (jeśli jest)."""
        extra = [self.input_file_var.get()] if self.input_file_var.get() else []
        self._run_script(["CzyszczenieAdresu.py", "czyszczeniadresu.py"], extra_args=extra)

    def choose_base_folder(self):
        d = filedialog.askdirectory(title="Wybierz folder bazowy", initialdir=self.folder_var.get())
        if d:
            self.folder_var.set(d)
            self._save_settings()  # <--- ZAPISUJEMY WYBÓR

    def prepare_app(self):
        base = Path(self.folder_var.get()).resolve()
        for p in ["linki", "województwa", "logs"]:
            (base / p).mkdir(parents=True, exist_ok=True)
        # na wszelki wypadek też zapisz ustawienia po przygotowaniu
        self._save_settings()
        messagebox.showinfo("Przygotowanie Aplikacji", f"Przygotowano strukturę w:\n{base}")

    def choose_output_folder(self):
        d = filedialog.askdirectory(title="Wybierz folder zapisu wyników", initialdir=self.output_folder_var.get() or self.folder_var.get())
        if d: self.output_folder_var.set(d)

    def run_bazadanych(self):
        root_dir = Path(self.folder_var.get()).resolve()
        if not root_dir.exists():
            messagebox.showerror("Baza danych", f"Folder bazowy nie istnieje:\n{root_dir}")
            return
        try:
            from bazadanych import open_ui
        except Exception as e:
            messagebox.showerror("Baza danych", f"Nie mogę zaimportować bazadanych.py:\n{e}")
            return
        open_ui(root_dir, parent=self)

    def run_filter_script(self, event=None):
        """Uruchamia odpowiedni filtr po wyborze z listy, na AKTUALNYM pliku raportu."""
        choice = self.filter_var.get()
        report_path = self.input_file_var.get().strip()

        if not report_path:
            messagebox.showerror("Brak pliku", "Najpierw wybierz plik raportu (wejście).")
            self.filter_var.set("— brak —")
            return

        mapping = {
            "Jeden właściciel": ["jeden_właściciel.py", "jeden_wlasciciel.py"],
            "Lokal mieszkalny": ["LOKAL_MIESZKALNY.py", "lokal_mieszkalny.py"],
            "Jeden właściciel i Lokal mieszkalny": [
                "jeden_właściciel_i_LOKAL_MIESZKALNY.py",
                "jeden_wlasciciel_i_lokal_mieszkalny.py",
            ],
            "Cofnij": ["cofnij.py"],
        }

        if choice == "— brak —":
            return

        scripts = mapping.get(choice)
        if not scripts:
            messagebox.showerror("Błąd", f"Brak obsługi dla filtru: {choice}")
            return

        # każdy filtr działa na wybranym „Plik raportu (wejście)”
        self._run_script(scripts, extra_args=["--in", report_path])

    def automate(self):
        """Uruchamia skrypt Automat.py (lub automat.py) — bez argumentów."""
        self._run_script(["Automat.py", "automat.py"])

    # ---------- Nawigacja ----------
    def prev_row(self):
        if self.df is None: messagebox.showinfo("Nawigacja", "Najpierw wybierz plik raportu."); return
        i = 0 if self.current_idx is None else max(0, self.current_idx - 1)
        self.goto_row(i)

    def next_row(self):
        if self.df is None: messagebox.showinfo("Nawigacja", "Najpierw wybierz plik raportu."); return
        n = len(self.df.index); i = 0 if self.current_idx is None else min(n - 1, self.current_idx + 1)
        self.goto_row(i)

    # ---------- PODGLĄD ----------
    def goto_row(self, i: int):
        if self.df is None or i < 0 or i >= len(self.df.index): return
        self.current_idx = i
        row = self.df.iloc[i]
        lines = [f"Wiersz {i+1}/{len(self.df)}"]
        for label, candidates in PREVIEW_SPEC:
            if any(_norm(c) in HIDDEN_PREVIEW_COLS for c in candidates): pass
            col = _find_col(self.df.columns, candidates)
            val = _trim_after_semicolon(row[col]) if col else ""
            lines.append(f"• {label}: {val}")
        self.preview_label.config(text="\n".join(lines))

    # ---------- (pozostałe: kalkulacja, zapis) ----------
    # (wersja, która odszukuje w Polska.xlsx, filtruje IQR, zapisuje (Nr KW).xlsx
    #  i uzupełnia średnie + statystyczną wartość w raporcie)

    def calc_and_save_row(self):
        if self.df is None or self.current_idx is None:
            messagebox.showinfo("Zapis", "Najpierw wybierz plik raportu i wiersz."); return
        if not self.output_folder_var.get() and not self.folder_var.get():
            messagebox.showerror("Brak folderu", "Wybierz 'Folder zapisu wyników'."); return

        # 1) Pobierz dane z bieżącego wiersza
        row = self.df.iloc[self.current_idx]
        kw_col = _find_col(self.df.columns, ["Nr KW","nr_kw","nrksiegi","nr księgi","nr_ksiegi","numer księgi"])
        kw_value = (str(row[kw_col]).strip() if (kw_col and pd.notna(row[kw_col]) and str(row[kw_col]).strip()) else f"WIERSZ_{self.current_idx+1}")

        area_col = _find_col(self.df.columns, ["Obszar","metry","powierzchnia"])
        area_val = _to_float_maybe(_trim_after_semicolon(row[area_col])) if area_col else None
        if area_val is None:
            messagebox.showerror("Brak obszaru", "Nie znalazłem wartości obszaru/metry."); return

        # lokalizacja (opcjonalna poprawa dopasowania)
        def _get(cands):
            c = _find_col(self.df.columns, cands)
            return _trim_after_semicolon(row[c]) if c else ""
        woj_r = _get(["Województwo","wojewodztwo","woj"])
        pow_r = _get(["Powiat"])
        gmi_r = _get(["Gmina"])
        mia_r = _get(["Miejscowość","Miejscowosc","Miasto"])
        dzl_r = _get(["Dzielnica","Osiedle"])
        uli_r = _get(["Ulica","Ulica(dla budynku)","Ulica(dla lokalu)"])

        # 2) Otwórz Polska.xlsx
        base_dir = Path(self.folder_var.get()).resolve()
        polska_path = base_dir / "Polska.xlsx"
        if not polska_path.exists():
            messagebox.showerror("Brak pliku", f"Nie znaleziono pliku: {polska_path}"); return
        try:
            df_pl = pd.read_excel(polska_path)
        except Exception as e:
            messagebox.showerror("Błąd odczytu", f"Nie mogę wczytać {polska_path}:\n{e}"); return

        # kolumny w Polska.xlsx
        col_area_pl  = _find_col(df_pl.columns, ["metry","powierzchnia","m2","obszar"])
        col_price_pl = _find_col(df_pl.columns, ["cena_za_metr","cena za metr","cena za m²","cena za m2","cena/m2"])
        if col_area_pl is None or col_price_pl is None:
            messagebox.showerror("Kolumny w Polska.xlsx", "Nie znalazłem kolumn metrażu i/lub ceny za m² w Polska.xlsx.")
            return

        # 3) Filtr metrażu + lokalizacja
        margin_pct = float(self.margin_pct_var.get() or 0.0)
        margin_m2  = float(self.margin_m2_var.get() or 0.0)
        delta = (area_val * (margin_pct/100.0)) if margin_pct > 0 else margin_m2
        low, high = max(0.0, area_val - delta), area_val + delta

        m = df_pl[col_area_pl].map(_to_float_maybe)
        mask = (m >= low) & (m <= high)

        def _eq_mask(col_candidates, value):
            col = _find_col(df_pl.columns, col_candidates)
            if col is None or not str(value).strip():
                return pd.Series(True, index=df_pl.index)
            s = df_pl[col].astype(str).str.strip().str.lower()
            v = str(value).strip().lower()
            return s == v

        mask &= _eq_mask(["wojewodztwo","województwo"], woj_r)
        mask &= _eq_mask(["powiat"], pow_r)
        mask &= _eq_mask(["gmina"], gmi_r)
        mask &= _eq_mask(["miejscowosc","miasto","miejscowość"], mia_r)
        if dzl_r: mask &= _eq_mask(["dzielnica","osiedle"], dzl_r)
        if uli_r: mask &= _eq_mask(["ulica"], uli_r)

        df_sel = df_pl[mask].copy()

        if df_sel.empty:
            messagebox.showinfo("Brak dopasowań", f"Nie znaleziono rekordów w zakresie [{low:.2f} ; {high:.2f}] m².")
            return

        # 4) Usuń skrajne dane IQR
        prices = df_sel[col_price_pl].map(_to_float_maybe)
        df_sel = df_sel[prices.notna()].copy()
        prices = df_sel[col_price_pl].map(_to_float_maybe)
        if len(prices) >= 4:
            q1 = np.nanpercentile(prices, 25)
            q3 = np.nanpercentile(prices, 75)
            iqr = q3 - q1
            lo = q1 - 1.5*iqr
            hi = q3 + 1.5*iqr
            df_sel = df_sel[(prices >= lo) & (prices <= hi)].copy()

        # 5) Zapisz wybrane rekordy do (Nr KW).xlsx
        out_dir = Path(self.output_folder_var.get() or self.folder_var.get()).resolve()
        out_dir.mkdir(parents=True, exist_ok=True)
        out_path = out_dir / f"({kw_value}).xlsx"

        prices = df_sel[col_price_pl].map(_to_float_maybe)
        avg = float(np.nanmean(prices)) if not df_sel.empty else None
        summary = {c: "" for c in df_sel.columns}
        summary[col_price_pl] = avg if avg is not None else ""
        df_out = pd.concat([df_sel, pd.DataFrame([summary])], ignore_index=True)
        df_out.loc[len(df_out)-1, "ŚREDNIA_CENA_M2"] = avg if avg is not None else ""

        try:
            df_out.to_excel(out_path, index=False)
        except Exception as e:
            messagebox.showerror("Błąd zapisu", f"Nie udało się zapisać pliku:\n{out_path}\n\n{e}")
            return

        # 6) Wpisz średnie do pliku raportowego i zapisz raport
        corrected = (avg * (1 - margin_pct/100.0)) if (avg is not None and margin_pct > 0) else avg

        col_avg = _find_col(self.df.columns, ["Średnia cena za m² (z bazy)", "Srednia cena za m2 (z bazy)", "srednia cena za metr ( z bazy)"])
        col_avg_corr = _find_col(self.df.columns, ["Średnia skorygowana cena za m² (z bazy)", "Srednia skorygowana cena za m2 (z bazy)"])
        col_stat = _find_col(self.df.columns, ["Statystyczna wartość nieruchomości", "Statystyczna wartosc nieruchomosci"])

        if col_avg is None:
            col_avg = "Średnia cena za m² (z bazy)"
            self.df[col_avg] = ""
        if col_avg_corr is None:
            col_avg_corr = "Średnia skorygowana cena za m² (z bazy)"
            self.df[col_avg_corr] = ""
        if col_stat is None:
            col_stat = "Statystyczna wartość nieruchomości"
            self.df[col_stat] = ""

        self.df.at[self.current_idx, col_avg] = avg if avg is not None else ""
        self.df.at[self.current_idx, col_avg_corr] = corrected if corrected is not None else ""
        stat_val = (area_val * corrected) if (area_val is not None and corrected is not None) else ""
        self.df.at[self.current_idx, col_stat] = stat_val

        try:
            if self.input_path and self.input_path.suffix.lower() in (".xlsx", ".xls"):
                self.df.to_excel(self.input_path, index=False)
            elif self.input_path:
                self.df.to_csv(self.input_path, index=False)
        except Exception as e:
            messagebox.showwarning("Zapis raportu", f"Wyliczono wartości, ale nie udało się zapisać raportu:\n{self.input_path}\n\n{e}")

        msg = [f"Zapisano dobrane rekordy do: {out_path}"]
        if avg is not None:
            msg.append(f"Średnia cena/m²: {avg:,.2f}".replace(",", " ").replace(".", ","))
        if corrected is not None and corrected != avg:
            msg.append(f"Skorygowana (−{self.margin_pct_var.get():.1f}%): {corrected:,.2f}".replace(",", " ").replace(".", ","))
        if isinstance(stat_val, (int, float)):
            msg.append(f"Statystyczna wartość: {stat_val:,.2f}".replace(",", " ").replace(".", ","))
        messagebox.showinfo("Zakończono", "\n".join(msg))

def main():
    App().mainloop()

if __name__ == "__main__":
    main()
