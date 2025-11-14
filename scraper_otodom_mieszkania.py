#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import argparse
import csv
import json
import random
import re
import time
import os
from pathlib import Path

import requests
from bs4 import BeautifulSoup

# ====== KONFIG ======
# Chrome-only UA (bez prefiksu Mozilla/5.0)
UA = "Chrome/127.0.0.0"

FIELDS = [
    "cena", "cena_za_metr", "metry", "liczba_pokoi", "pietro", "rynek", "rok_budowy",
    "material", "wojewodztwo", "powiat", "gmina", "miejscowosc", "dzielnica",
    "ulica", "link"
]

FLOOR_MAP = {
    "ground_floor": "parter",
    "basement": "suterena",
    "loft": "poddasze",
}
MARKET_MAP = {
    "primary": "pierwotny",
    "secondary": "wtórny",
}

# ====== wykrywanie dzielnicy ======
KNOWN_DISTRICTS = [
    "Nowe Miasto", "Staromieście", "Baranówka", "Zalesie", "Drabinianka",
    "Budziwój", "Słocina", "Przybyszówka", "Zwięczyca", "Wilkowyja",
    "Bacieczki", "Bojary", "Dziesięciny", "Piasta",
]
FRAN_ANY = re.compile(r"\b(Frani\w*\s+Kotuli)\b", re.I)

BETWEEN_STREET_CITY = re.compile(
    r"(ul\.|ulica)?\s*([A-ZŁŚŻŹĆŃ][\w\-\s\.']+)\s*,\s*([A-ZŁŚŻŹĆŃ][\w\-\s\.']+)\s*,\s*([A-ZŁŚŻŹĆŃ][\w\-\s\.']+)",
    re.I
)

# -------------------------------------------------
# TLS / CA bundle
# -------------------------------------------------

def _find_ca_bundle() -> str | bool:
    """
    Zwraca ścieżkę do pliku CA bundle, jeśli uda się ją wykryć.
    Priorytety:
      1) REQUESTS_CA_BUNDLE / SSL_CERT_FILE (jeśli wskazują istniejący plik)
      2) certifi.where() (jeśli dostępne)
      3) True -> użyj domyślnego store systemowego (requests verify=True)
    """
    for var in ("REQUESTS_CA_BUNDLE", "SSL_CERT_FILE"):
        p = os.environ.get(var)
        if p and Path(p).is_file():
            return p
    try:
        import certifi  # type: ignore
        p = certifi.where()
        if p and Path(p).is_file():
            return p
    except Exception:
        pass
    return True  # domyślny mechanizm requests (verify=True)


def mk_session() -> requests.Session:
    """
    Buduje sesję HTTP z nagłówkami i automatycznie ustawionym verify.
    """
    s = requests.Session()
    s.headers.update({
        "User-Agent": UA,
        "Accept": "*/*",
        "Accept-Language": "pl-PL,pl;q=0.9,en;q=0.8",
        "Connection": "keep-alive",
        "Pragma": "no-cache",
        "Cache-Control": "no-cache",
        "Referer": "https://www.otodom.pl/",
    })
    s.verify = _find_ca_bundle()
    # diagnostyka (zobaczysz w terminalu, czego używa TLS verify)
    try:
        print(f"[HTTP] TLS verify -> {s.verify}")
    except Exception:
        pass
    return s

# -------------------------------------------------
# NARZĘDZIA HTML / JSON
# -------------------------------------------------

def extract_next_data(html: str):
    """
    Szukamy __NEXT_DATA__ albo podobnego JSON-a z props->pageProps.
    """
    soup = BeautifulSoup(html, "html.parser")
    tag = soup.find("script", id="__NEXT_DATA__", type="application/json")
    if tag and tag.string:
        try:
            return json.loads(tag.string)
        except Exception:
            pass

    # fallback: dowolny <script type="application/json"> z pageProps
    for s in soup.find_all("script", attrs={"type": "application/json"}):
        try:
            obj = json.loads(s.string or "")
            if isinstance(obj, dict) and "props" in obj and "pageProps" in obj["props"]:
                return obj
        except Exception:
            continue
    return None


def deep_iter(obj):
    if isinstance(obj, dict):
        for k, v in obj.items():
            yield k, v
            yield from deep_iter(v)
    elif isinstance(obj, list):
        for v in obj:
            yield from deep_iter(v)


def get_char(characteristics, key, prefer_localized=True):
    """
    Pobierz cechę o danym 'key' z listy characteristics (otodom ad["characteristics"]).
    """
    if not characteristics:
        return ""
    for ch in characteristics:
        if ch.get("key") == key:
            if prefer_localized and ch.get("localizedValue"):
                return str(ch["localizedValue"]).strip()
            return str(ch.get("value") or "").strip()
    return ""


def pick_name(d, key):
    """
    Jeżeli pole jest dictem z 'name'/'label', zwróć to.
    W innym wypadku zwróć surową wartość.
    """
    v = (d or {}).get(key)
    if isinstance(v, dict):
        return v.get("name", "") or v.get("label", "") or ""
    return v or ""


def all_strings(obj, max_len=200):
    """
    Generator wszystkich krótkich stringów ze zagnieżdżonego JSON-a.
    Używane do heurystyk (dzielnica).
    """
    seen = set()
    for _k, v in deep_iter(obj):
        if isinstance(v, str):
            s = v.strip()
            if s and len(s) <= max_len and s not in seen:
                seen.add(s)
                yield s


def detect_dzielnica(next_data, miasto, ulica):
    """
    Heurystyka dzielnicy/osiedla:
    1) Jeżeli w tekście mamy wzorzec "ul. X, COŚ, Miasto" — to "COŚ" traktujemy jako dzielnicę.
    2) Specjalny case os. Franciszka Kotuli itd.
    3) fallback: lista znanych osiedli.
    """
    text = " | ".join(all_strings(next_data, 300))

    # 1) "ulica, [osiedle], Miasto"
    try:
        for m in BETWEEN_STREET_CITY.finditer(text):
            _ul_lab, ul_name, maybe_dist, city = m.groups()
            if miasto and city and city.lower() == str(miasto).lower():
                if ulica and ul_name and ul_name.lower() in str(ulica).lower():
                    if maybe_dist and maybe_dist.lower() != city.lower():
                        return maybe_dist.strip()
    except Exception:
        pass

    # 2) specyficzne osiedla (np. "os. Franciszka Kotuli")
    m = FRAN_ANY.search(text)
    if m:
        return m.group(1)

    # 3) twarda lista znanych dzielnic/osiedli
    for name in KNOWN_DISTRICTS:
        if re.search(rf"\b{name}\b", text, flags=re.I):
            return name

    return ""

# --- znacznik do świadomego pomijania ogłoszeń ---
class SkipAd(Exception):
    def __init__(self, reason: str):
        super().__init__(reason)
        self.reason = reason


def parse_ad(next_data: dict, url: str) -> dict:
    """
    Główny parser pojedynczego ogłoszenia → dict(FIELDS).
    Może rzucić SkipAd('price_unknown' | 'missing_core_fields'), aby POMINĄĆ rekord.
    """
    page_props = (next_data.get("props") or {}).get("pageProps", {})
    ad = page_props.get("ad") or {}

    # fallback: ręczne szukanie węzła z "characteristics" i "location"
    if not ad:
        def walk(d):
            if isinstance(d, dict):
                if "characteristics" in d and "location" in d:
                    return d
                for v in d.values():
                    r = walk(v)
                    if r:
                        return r
            elif isinstance(d, list):
                for v in d:
                    r = walk(v)
                    if r:
                        return r
            return None
        found = walk(page_props)
        if found:
            ad = found

    chars = ad.get("characteristics") or []

    cena = get_char(chars, "price")
    # 1) "Zapytaj o cenę" → pomijamy
    if (cena or "").strip().lower().startswith("zapytaj"):
        raise SkipAd("price_unknown")

    cena_m = get_char(chars, "price_per_m")
    metry = get_char(chars, "m")
    pokoje = get_char(chars, "rooms_num")

    floor_val = get_char(chars, "floor_no", prefer_localized=False)
    pietro = get_char(chars, "floor_no", prefer_localized=True) or FLOOR_MAP.get(
        floor_val, floor_val
    )

    rynek_raw = (get_char(chars, "market", prefer_localized=False) or "").lower()
    rynek = MARKET_MAP.get(rynek_raw, get_char(chars, "market", prefer_localized=True))

    rok = (
        get_char(chars, "build_year", prefer_localized=False)
        or get_char(chars, "build_year")
    )
    material = get_char(chars, "building_material")

    addr = ((ad.get("location") or {}).get("address")) or {}
    woj = pick_name(addr, "province")
    powiat = pick_name(addr, "county")
    gmina = pick_name(addr, "municipality")
    miasto = pick_name(addr, "city")
    dzielnica = pick_name(addr, "district")
    ulica = pick_name(addr, "street")

    # fallback — próbuj zebrać brakujące elementy z innych gałęzi
    if not (woj and gmina and miasto and (dzielnica or ulica)):
        for _k, v in deep_iter(next_data):
            if isinstance(v, dict):
                keys = set(v.keys())
                if {"province", "county", "municipality", "city", "district", "street"} & keys:
                    woj = woj or pick_name(v, "province")
                    powiat = powiat or pick_name(v, "county")
                    gmina = gmina or pick_name(v, "municipality")
                    miasto = miasto or pick_name(v, "city")
                    dzielnica = dzielnica or pick_name(v, "district")
                    ulica = ulica or pick_name(v, "street")

    # heurystyka dzielnicy jeśli wciąż pusto
    if not dzielnica:
        dzielnica = detect_dzielnica(next_data, miasto, ulica)

    link = ad.get("url") or url

    row = {
        "cena": cena or "",
        "cena_za_metr": cena_m or "",
        "metry": metry or "",
        "liczba_pokoi": pokoje or "",
        "pietro": pietro or "",
        "rynek": rynek or "",
        "rok_budowy": (str(rok) if rok is not None else ""),
        "material": material or "",
        "wojewodztwo": woj or "",
        "powiat": powiat or "",
        "gmina": gmina or "",
        "miejscowosc": miasto or "",
        "dzielnica": dzielnica or "",
        "ulica": ulica or "",
        "link": link or "",
    }

    # 2) minimalna walidacja – jeśli kluczowe pola puste → pomijamy
    if not any(row.get(k) for k in ("cena", "metry", "liczba_pokoi")):
        raise SkipAd("missing_core_fields")

    return row


def fetch_one(url: str, session: requests.Session, retries: int = 3, backoff: float = 1.6) -> dict | None:
    """
    Pobiera pojedyncze ogłoszenie (z retry).
    Zwraca dict, a przy świadomym pominięciu/błędzie → None (NIE zapisujemy do CSV).
    """
    last_exc = None
    for attempt in range(1, retries + 1):
        try:
            # verify ustawione na poziomie session.verify (patrz mk_session)
            r = session.get(url, timeout=30)
            r.raise_for_status()
            data = extract_next_data(r.text)
            if not data:
                raise SkipAd("no_next_data")
            row = parse_ad(data, url)
            return row
        except SkipAd as sk:
            if sk.reason == "price_unknown":
                print(f"[skip] cena ukryta (Zapytaj o cenę): {url}")
            elif sk.reason == "missing_core_fields":
                print(f"[skip] brak kluczowych pól: {url}")
            else:
                print(f"[skip] brak danych do parsowania: {url}")
            return None
        except Exception as e:
            last_exc = e
            if attempt < retries:
                time.sleep(backoff ** attempt)
            else:
                print(f"[skip] błąd: {last_exc} -> {url}")
                return None


# -------------------------------------------------
# I/O LINKÓW I CSV
# -------------------------------------------------

def guess_region_name_from_path(path: Path) -> str:
    """
    Zgadujemy nazwę regionu na podstawie nazwy pliku wejściowego
    (np. Podlaskie.csv -> Podlaskie).
    """
    return path.stem


def read_links_any(input_path: Path) -> list[str]:
    """
    Czyta plik linków:
    - CSV (nagłówek link/url albo pierwsza kolumna) lub
    - zwykły txt (jeden URL na linię).
    Zwraca listę unikalnych linków.
    """
    links = []
    text = input_path.read_text(encoding="utf-8", errors="ignore")

    # spróbuj CSV
    try:
        rows = list(csv.reader(text.splitlines()))
        if rows:
            hdr = [h.strip().lower() for h in rows[0]]
            start_idx = 1 if any(h in ("link", "url") for h in hdr) else 0
            for row in rows[start_idx:]:
                for cell in row:
                    if isinstance(cell, str) and cell.startswith("http"):
                        links.append(cell.strip())
                        break
            if links:
                return dedupe_preserve_order(links)
    except Exception:
        pass

    # fallback: linie-URL
    for ln in text.splitlines():
        ln = ln.strip()
        if ln.startswith("http"):
            links.append(ln)

    return dedupe_preserve_order(links)


def dedupe_preserve_order(iterable):
    seen = set()
    out = []
    for x in iterable:
        if x not in seen:
            seen.add(x)
            out.append(x)
    return out


def append_rows_csv(path: Path, rows: list[dict]):
    """
    Dopisuje wiersze do CSV.
    Jeśli plik nie istniał – zapisuje nagłówek FIELDS.
    """
    new_file = not path.exists()
    with path.open("a", encoding="utf-8-sig", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=FIELDS)
        if new_file:
            w.writeheader()
        for r in rows:
            w.writerow({k: r.get(k, "") for k in FIELDS})


def read_processed_links(out_path: Path) -> set[str]:
    """
    Czytamy plik wynikowy (województwa/<region>.csv)
    i zbieramy już zapisane linki, żeby wznawiać pracę.
    """
    if not out_path.exists():
        return set()
    processed = set()
    with out_path.open("r", encoding="utf-8-sig", newline="") as fh:
        rd = csv.DictReader(fh)
        if rd.fieldnames:
            link_col = None
            for name in rd.fieldnames:
                if name.lower() == "link":
                    link_col = name
                    break
            if link_col:
                for row in rd:
                    val = (row.get(link_col) or "").strip()
                    if val:
                        processed.add(val)
    return processed


# -------------------------------------------------
# GŁÓWNY PRZEBIEG
# -------------------------------------------------

def main():
    ap = argparse.ArgumentParser(
        description=(
            "Scraper otodom — tryb B: --input/--output (zgodny z naszym GUI). "
            "Wznawia po linkach już obecnych w pliku wyjściowym."
        )
    )

    # Preferowany tryb
    ap.add_argument("--input", help="pełna ścieżka do pliku z linkami (CSV/txt)", default=None)
    ap.add_argument("--output", help="pełna ścieżka do pliku wynikowego CSV", default=None)

    # Legacy (zostawione dla zgodności, ale w GUI nieużywane)
    ap.add_argument("--region", default=None, help="np. podlaskie, dolnośląskie itd. (legacy)")
    ap.add_argument("--links_dir", default=None, help="katalog z plikami linków (legacy)")
    ap.add_argument("--out_dir", default=None, help="katalog wynikowy CSV (legacy)")

    # Tech parametry
    ap.add_argument("--delay_min", type=float, default=4.0,
                    help="minimalne opóźnienie między ogłoszeniami (sek)")
    ap.add_argument("--delay_max", type=float, default=6.0,
                    help="maksymalne opóźnienie między ogłoszeniami (sek)")
    ap.add_argument("--retries", type=int, default=3,
                    help="ile razy próbować pobrać pojedyncze ogłoszenie")

    args = ap.parse_args()

    # Ustal ścieżki wejście/wyjście
    if args.input and args.output:
        input_path = Path(args.input)
        output_path = Path(args.output)
        region_name = guess_region_name_from_path(input_path)
    else:
        # tryb legacy
        if not (args.region and args.links_dir and args.out_dir):
            ap.error("Podaj --input i --output, albo (legacy) --region, --links_dir i --out_dir.")
        region_file = normalize_region_filename(args.region)
        input_path = Path(args.links_dir) / region_file
        output_path = Path(args.out_dir) / region_file
        region_name = Path(region_file).stem

    if not input_path.exists():
        raise SystemExit(f"[ERR] Brak pliku linków: {input_path}")

    # Sesja HTTP (automatyczne wykrycie CA bundle)
    session = mk_session()

    # Wczytaj listę linków źródłowych
    links = read_links_any(input_path)
    total_links = len(links)

    # Wznawianie: wczytaj już przerobione linki z outputu
    processed = read_processed_links(output_path)
    todo = [u for u in links if u not in processed]

    print(
        f"[start] region='{region_name}' "
        f"links={total_links} input='{input_path}' output='{output_path}'"
    )

    if processed:
        print(f"[resume] wykryto już zapisane rekordy: {len(processed)} — pominę je")
        print(f"[todo] do zrobienia: {len(todo)}")

    if not todo:
        print("[done] Wszystkie linki z pliku wejściowego są już przerobione.")
        return

    # Tworzymy katalog wyjściowy (województwa/) jeśli go nie ma
    output_path.parent.mkdir(parents=True, exist_ok=True)

    # pętla po TODO – zapis po KAŻDYM ogłoszeniu
    done_now = 0
    skipped_now = 0
    for idx, url in enumerate(todo, 1):
        print(f"[{idx}/{len(todo)}] Pobieram: {url}")
        row = fetch_one(url, session, retries=args.retries)
        if row is None:
            skipped_now += 1
        else:
            append_rows_csv(output_path, [row])
            done_now += 1

        # losowa pauza 4.0 - 6.0 s
        if args.delay_max > 0:
            dly = random.uniform(max(0.0, args.delay_min),
                                 max(args.delay_min, args.delay_max))
            print(f"    ↳ pauza {dly:.2f} s…")
            time.sleep(dly)

    print(f"[OK] dopisano {done_now} rekordów (pominięto {skipped_now}) -> {output_path}")


def normalize_region_filename(region: str) -> str:
    """
    Legacy helper:
    np. 'dolnośląskie' -> 'Dolnośląskie.csv'
    Jeśli użytkownik dał już rozszerzenie .csv, nie zmieniamy.
    """
    base = region.strip()
    if not base:
        return "Region.csv"
    if base.lower().endswith(".csv"):
        return base
    return f"{base[0].upper()}{base[1:]}.csv"


if __name__ == "__main__":
    main()
