"""
Aplikasi Merge Excel + Dashboard Penjualan Hotel
-------------------------------------------------
Menggabungkan beberapa file laporan "Sales Summary" (format export sistem
PT. Mitra Tours & Travel) menjadi satu dataset, lalu menampilkannya sebagai
dashboard interaktif.

Cara menjalankan:
    pip install -r requirements.txt
    streamlit run app.py
"""

import hashlib
import html
import io
import json
import os
import re
from difflib import SequenceMatcher
import time
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

try:  # rapidfuzz jauh lebih cepat; jika belum terpasang, pakai difflib bawaan Python
    from rapidfuzz import fuzz, process
    HAS_RAPIDFUZZ = True
except ImportError:
    HAS_RAPIDFUZZ = False

# ---------------------------------------------------------------------------
# Konfigurasi
# ---------------------------------------------------------------------------
st.set_page_config(page_title="Merge Excel & Dashboard", page_icon="📊", layout="wide")
# Angka KPI tidak dipotong ("…") di layar sempit: font menyesuaikan lebar & boleh turun baris
st.markdown("""<style>
[data-testid="stMetricValue"] {font-size: clamp(1.3rem, 2.4vw, 2.25rem);}
[data-testid="stMetricValue"] > div {white-space: normal; overflow: visible; text-overflow: clip;}
/* Tampilan dashboard: bersih & ringkas */
.block-container {padding-top: 2.2rem; padding-bottom: 3rem; max-width: 1400px;}
h1 {font-size: 1.85rem !important; font-weight: 700 !important; letter-spacing: -.01em;}
[class*="st-key-card"] {border-radius: 16px !important; padding: 14px 16px 10px !important;
    box-shadow: 0 1px 2px rgba(16,24,40,.04), 0 4px 14px rgba(16,24,40,.05);}
.stTabs [data-baseweb="tab-list"] {gap: 6px;}
.stTabs [data-baseweb="tab"] {padding: 8px 14px; border-radius: 10px 10px 0 0;}
.sec-h {display:flex; align-items:baseline; gap:10px; margin: 6px 0 2px;}
.sec-h .t {font-size: 1.05rem; font-weight: 700; color: #434343;}
h1 {color: #434343;}
.stTabs [aria-selected="true"] {color: #0D7FCC !important;}
.stTabs [data-baseweb="tab-highlight"] {background-color: #0D7FCC !important;}
.stButton > button[kind="primary"], .stDownloadButton > button[kind="primary"],
[data-testid="stBaseButton-primary"] {background: #0D7FCC !important; border-color: #0D7FCC !important;}
.stButton > button[kind="primary"]:hover, .stDownloadButton > button[kind="primary"]:hover,
[data-testid="stBaseButton-primary"]:hover {background: #1BA0E2 !important; border-color: #1BA0E2 !important;}
[class*="st-key-card"] {border-top: 3px solid #1BA0E2 !important;}
.sec-h .s {font-size: .82rem; opacity: .6;}
</style>""", unsafe_allow_html=True)

# Palet warna tema (brand): #434343 (Pantone 446), #1BA0E2 (2925C),
# #1494C6 (Hexachrome Cyan), #0D7FCC (299 2X) + turunan terangnya
PALETTE = ["#0D7FCC", "#1BA0E2", "#434343", "#8FD0F1", "#1494C6", "#A6A6A6", "#BFE4F7"]


def section(title: str, subtitle: str = "", box=None):
    """Judul bagian kecil yang rapi (judul tebal + keterangan abu-abu)."""
    (box or st).markdown(f'<div class="sec-h"><span class="t">{title}</span>'
                         f'<span class="s">{subtitle}</span></div>', unsafe_allow_html=True)


def clean_fig(fig, height: int = 360, legend_bottom: bool = False):
    """Gaya grafik plotly yang bersih: latar transparan, grid tipis, tanpa judul, legenda rapi."""
    fig.update_layout(
        template="plotly_white", height=height, title=None,
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Source Sans Pro, sans-serif", size=13),
        margin=dict(t=10, b=10 if not legend_bottom else 40, l=10, r=10),
        legend=(dict(orientation="h", y=-0.08, x=0.5, xanchor="center", title=None)
                if legend_bottom else
                dict(orientation="h", y=1.08, x=0, xanchor="left", title=None)),
        hoverlabel=dict(font_size=12))
    fig.update_xaxes(showgrid=False, zeroline=False)
    fig.update_yaxes(gridcolor="rgba(128,128,128,.15)", zeroline=False)
    return fig

EXCEL_EXT = (".xlsx", ".xlsm", ".xls")

DATE_COLS = ["Inv Date", "Due Date", "Issued Date", "Check In", "Check Out",
             "Departure Date", "Arrival Date"]

NUMERIC_COLS = ["Room", "Night", "Base Fare", "Fare Tax", "IWJR", "Add Charge",
                "Insurance", "PSC", "Other Charge", "Incentive", "Agent Comm", "NTA",
                "Travel Services", "Sales Net", "VAT", "Stamp Fee", "MDR", "Sales AR",
                "Extra Disc", "Profit", "Profit %", "Rounding", "Base Sell"]

BULAN = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"]


# ---------------------------------------------------------------------------
# Kategori Supplier
#   Nama supplier dicocokkan tanpa membedakan huruf besar/kecil, spasi & tanda baca
#   (mis. 'TIKET.COM', 'Tiket.com', 'tiket com' -> OTA). Pencocokan berbasis "mengandung",
#   jadi 'EXPEDIA GROUP' atau 'Webbeds FZ LLC' juga ikut terbaca.
#   Supplier yang tidak ada di daftar -> 'Direct To Hotel'.
# ---------------------------------------------------------------------------
SUPPLIER_CATEGORY_RULES = {
    "EXPEDIA": "Wholesaler",
    "FIT RUUMS": "Wholesaler",
    "Destination Of The World": "Wholesaler",
    "MG HOLIDAY": "Wholesaler",
    "Channel Manager": "PERTAMINA Corp Rate",
    "KLIKNBOOK": "Wholesaler",
    "TIKET.COM": "OTA",
    "TRIP.COM": "OTA",
    "Agoda": "OTA",
    "TBO Holidays": "Wholesaler",
    "Traveloka": "OTA",
    "Webbeds": "Wholesaler",
}
SUPPLIER_DEFAULT_CATEGORY = "Direct To Hotel"

# Penulisan baku (Proper) untuk supplier di daftar kategori. Dicocokkan tanpa membedakan
# huruf besar/kecil, spasi & tanda baca (EXPEDIA / Expedia / expedia -> "Expedia").
# Supplier lain otomatis ditulis Proper (mis. "HOTEL MULIA" -> "Hotel Mulia").
# Scorecard Room Night per supplier (tab Filter per Hotel); supplier lain -> "Other"
SUPPLIER_RN_CARDS = ["Expedia", "Fit Ruums", "Destination Of The World", "MG Holiday",
                     "Channel Manager", "Kliknbook", "Tiket.com", "Traveloka", "Webbeds"]

SUPPLIER_STANDARD_NAMES = {
    "EXPEDIA": "Expedia",
    "FIT RUUMS": "Fit Ruums",
    "Destination Of The World": "Destination Of The World",
    "MG HOLIDAY": "MG Holiday",
    "Channel Manager": "Channel Manager",
    "KLIKNBOOK": "Kliknbook",
    "TIKET.COM": "Tiket.com",
    "TRIP.COM": "Trip.com",
    "Agoda": "Agoda",
    "TBO Holidays": "TBO Holidays",
    "Traveloka": "Traveloka",
    "Webbeds": "Webbeds",
}


def _compact(s: str) -> str:
    """Huruf kecil, tanpa spasi & tanda baca: 'TIKET.COM' -> 'tiketcom'."""
    return re.sub(r"[^a-z0-9]+", "", html.unescape(str(s)).lower())


def supplier_category(sr: pd.Series) -> pd.Series:
    """Petakan Supplier Name ke kategori (Wholesaler / OTA / PERTAMINA Corp Rate /
    Direct To Hotel). Tidak peka huruf besar/kecil."""
    key = sr.astype(str).map(_compact)
    out = pd.Series(SUPPLIER_DEFAULT_CATEGORY, index=sr.index, dtype=object)
    for name, cat in SUPPLIER_CATEGORY_RULES.items():
        out[key.str.contains(_compact(name), regex=False)] = cat
    return out


# ---------------------------------------------------------------------------
# Standarisasi nama hotel dari file master (kolom STADARD NAME)
#   File master default: unique_hotels.xlsx di folder yang sama dengan app.py,
#   atau di-upload lewat sidebar. Sheet dipilih otomatis: sheet pertama yang
#   punya kolom 'HOTEL NAME' + 'STADARD NAME' (sheet 'Standarized').
#   Pencocokan: 1) persis (spasi dirapikan, &amp; dsb. di-unescape),
#               2) jika tidak ketemu, tanpa membedakan huruf besar/kecil.
#   Jika satu nama asli punya >1 nama standar di master, baris PALING BAWAH dipakai
#   (blok bawah master = blok terbaru yang berisi CHAIN).
# ---------------------------------------------------------------------------
MASTER_FILE = "unique_hotels.xlsx"
STD_COL_CANDIDATES = ["STADARD NAME", "STANDARD NAME"]


def _clean_hotel(sr: pd.Series) -> pd.Series:
    return (sr.astype(str).map(html.unescape)
            .str.replace("\u00a0", " ").str.replace(r"\s+", " ", regex=True).str.strip())


def read_hotel_master(source) -> pd.DataFrame:
    """Kembalikan master berkolom: _key, _key_lower, Standard Name, Hotel Chain, Hotel Rating."""
    xls = pd.ExcelFile(source)
    pick = None
    for std_col in STD_COL_CANDIDATES:        # prioritas: kolom 'STADARD NAME'
        for sh in xls.sheet_names:
            head = pd.read_excel(xls, sh, nrows=0)
            cols = {str(c).strip().upper(): c for c in head.columns}
            if "HOTEL NAME" in cols and std_col in cols:
                pick = (sh, cols["HOTEL NAME"], cols[std_col], cols.get("CHAIN", cols.get("CHAIN NAME")),
                        cols.get("RATING", cols.get("CATEGORY")))
                break
        if pick:
            break
    if not pick:
        raise ValueError("Tidak ditemukan sheet dengan kolom 'HOTEL NAME' dan 'STADARD NAME'.")
    sh, c_name, c_std, c_chain, c_rating = pick
    m = pd.read_excel(xls, sh, dtype=object)
    out = pd.DataFrame({
        "_key": _clean_hotel(m[c_name]),
        "Standard Name": _clean_hotel(m[c_std]),
        "Hotel Chain": m[c_chain] if c_chain is not None else None,
        "Hotel Rating": pd.to_numeric(m[c_rating], errors="coerce") if c_rating is not None else None,
    })
    out = out[out["_key"].ne("") & out["_key"].ne("nan")
              & out["Standard Name"].ne("") & out["Standard Name"].ne("nan")]
    out["_key_lower"] = out["_key"].str.lower()
    out.attrs["sheet"] = sh
    return out.reset_index(drop=True)


def _lookup(master: pd.DataFrame, key_col: str) -> pd.DataFrame:
    """Satu baris per kunci: nama standar = baris terakhir; chain/rating = nilai terakhir yang terisi."""
    g = master.groupby(key_col, sort=False)
    res = g[["Hotel Chain", "Hotel Rating"]].last()       # last() melewati nilai kosong
    res["Standard Name"] = g["Standard Name"].last()      # selalu terisi -> baris terakhir
    # Lengkapi chain/rating dari baris lain yang nama standarnya sama
    by_std = master.groupby("Standard Name")[["Hotel Chain", "Hotel Rating"]].last()
    for c in ("Hotel Chain", "Hotel Rating"):
        res[c] = res[c].fillna(res["Standard Name"].map(by_std[c]))
    return res


def standardize_hotels(names: pd.Series, master: pd.DataFrame) -> pd.DataFrame:
    """names: nama hotel unik. Return DataFrame index=nama asli dengan kolom
    Standard Name, Hotel Chain, Hotel Rating, Cocok ('Persis'/'Huruf besar-kecil'/None)."""
    names = pd.Series(pd.unique(names.dropna().astype(str)))
    key = _clean_hotel(names)
    exact, lower = _lookup(master, "_key"), _lookup(master, "_key_lower")
    a = exact.reindex(key.values).set_axis(names.values)
    b = lower.reindex(key.str.lower().values).set_axis(names.values)
    res = a.combine_first(b)
    res["Cocok"] = None
    res.loc[b["Standard Name"].notna().values, "Cocok"] = "Huruf besar-kecil"
    res.loc[a["Standard Name"].notna().values, "Cocok"] = "Persis"
    return res


# ---------------------------------------------------------------------------
# Pengelompokan nama hotel yang mirip (fuzzy matching)
#   1. Normalisasi nama: huruf kecil, hapus isi kurung, '&' -> 'and', hapus tanda
#      baca, kata umum (hotel, by IHG, managed by ...) dan nama kota hotel.
#   2. Nama yang identik setelah normalisasi langsung digabung.
#   3. Sisanya dibandingkan dengan skor kemiripan di dalam kota yang sama;
#      skor >= ambang digabung (union-find), dengan pengaman: kata pertama/merek
#      harus mirip dan angka dalam nama harus sama.
#   4. Nama grup = varian dengan transaksi terbanyak.
# ---------------------------------------------------------------------------
DEFAULT_STOPWORDS = [
    "hotel", "hotels", "the", "and", "by ihg", "an ihg hotel", "indonesia",
    "managed by .*", "formerly .*", "syariah",
]


def norm_city(x) -> str:
    s = re.sub(r"\s+", " ", str(x)).strip().lower()
    s = re.sub(r"\b(regency|city|kota|kabupaten|kab)\b", "", s).strip()
    return s if s and s not in ("-", "nan", "none") else "(tanpa kota)"


def normalize_name(name: str, city: str, stop_patterns: list) -> str:
    s = html.unescape(str(name)).lower()
    s = re.sub(r"\(.*?\)|\[.*?\]", " ", s)          # isi dalam kurung
    s = s.replace("&", " and ").replace("'", "").replace("’", "")
    s = re.sub(r"[^a-z0-9 ]+", " ", s)               # tanda baca -> spasi
    s = f" {re.sub(r' +', ' ', s).strip()} "
    for pat in stop_patterns:
        s = re.sub(rf" {pat}(?= )", " ", s)
    if city and city != "(tanpa kota)":
        s = s.replace(f" {city} ", " ")
    s = re.sub(r" +", " ", s).strip()
    return s or str(name).lower().strip()


def _ratio(a: str, b: str) -> float:
    if HAS_RAPIDFUZZ:
        return fuzz.ratio(a, b)
    return SequenceMatcher(None, a, b).ratio() * 100


def _score(a: str, b: str, **_) -> float:
    """Skor kemiripan: maksimum dari perbandingan biasa, urutan kata diabaikan,
    dan perbandingan tanpa spasi (untuk 'de grand city' vs 'degrandcity')."""
    sort_a, sort_b = " ".join(sorted(a.split())), " ".join(sorted(b.split()))
    return max(_ratio(a, b), _ratio(sort_a, sort_b),
               _ratio(a.replace(" ", ""), b.replace(" ", "")))


def _guard(a: str, b: str) -> bool:
    """Pengaman agar hotel berbeda tidak ikut tergabung:
    - kata pertama (biasanya merek) harus mirip, atau awalan tanpa spasi sama
      (contoh ditolak: 'Meru Resort Sanur' vs 'Mercure Resort Sanur')
    - angka di nama harus sama (contoh ditolak: 'Panglima Polim' vs 'Panglima Polim 2')"""
    ta, tb = a.split(), b.split()
    if not ta or not tb:
        return False
    if re.findall(r"\d+", a) != re.findall(r"\d+", b):
        return False
    na, nb = a.replace(" ", ""), b.replace(" ", "")
    return _ratio(ta[0], tb[0]) >= 80 or (len(na) >= 5 and na[:5] == nb[:5])


class _UF:
    def __init__(self, n):
        self.p = list(range(n))

    def find(self, i):
        while self.p[i] != i:
            self.p[i] = self.p[self.p[i]]
            i = self.p[i]
        return i

    def union(self, a, b):
        ra, rb = self.find(a), self.find(b)
        if ra != rb:
            self.p[rb] = ra


def build_hotel_groups(hotels: pd.DataFrame, threshold: int = 90,
                       same_city: bool = True, stopwords=None) -> pd.DataFrame:
    """
    hotels: DataFrame berkolom ['Hotel Name', 'Hotel City', 'Transaksi'] (unik per nama+kota)
    Return: mapping dengan kolom
        Hotel City, Hotel Name, Kota (norm), Nama Normal, Hotel (Grup), Skor, Transaksi, Jumlah Varian
    """
    stop_patterns = [re.escape(w).replace(r"\.\*", ".*").replace(r"\ ", " ")
                     for w in (stopwords if stopwords is not None else DEFAULT_STOPWORDS)
                     if str(w).strip()]
    df = hotels.copy().reset_index(drop=True)
    df["Kota (norm)"] = df["Hotel City"].map(norm_city)
    df["Nama Normal"] = [normalize_name(n, c, stop_patterns)
                         for n, c in zip(df["Hotel Name"], df["Kota (norm)"])]

    uf = _UF(len(df))
    best = [100.0] * len(df)
    if same_city:
        blocks = df.groupby("Kota (norm)").groups
    else:  # lintas kota: kelompokkan berdasarkan kata pertama agar tetap cepat
        blocks = df.groupby(df["Nama Normal"].str.split().str[0].fillna("")).groups

    for _, idx in blocks.items():
        idx = list(idx)
        if len(idx) < 2:
            continue
        names = df.loc[idx, "Nama Normal"].tolist()
        # 1) identik setelah normalisasi
        first = {}
        for i, n in zip(idx, names):
            if n in first:
                uf.union(first[n], i)
            else:
                first[n] = i
        # 2) fuzzy antar nama normal yang unik
        uniq = list(first.keys())
        if len(uniq) < 2:
            continue
        if HAS_RAPIDFUZZ:
            mat = process.cdist(uniq, uniq, scorer=_score, score_cutoff=threshold, workers=-1)
        else:  # fallback lebih lambat tanpa rapidfuzz
            mat = [[_score(x, y) if j > i else 0 for j, y in enumerate(uniq)]
                   for i, x in enumerate(uniq)]
            import numpy as np
            mat = np.array(mat)
        for a in range(len(uniq)):
            for b in range(a + 1, len(uniq)):
                if mat[a, b] >= threshold and _guard(uniq[a], uniq[b]):
                    ia, ib = first[uniq[a]], first[uniq[b]]
                    uf.union(ia, ib)
                    best[ia] = min(best[ia], mat[a, b])
                    best[ib] = min(best[ib], mat[a, b])

    df["_root"] = [uf.find(i) for i in range(len(df))]
    # Nama grup = nama yang sudah standar (bila ada), lalu varian dengan transaksi terbanyak
    if "Standar" not in df.columns:
        df["Standar"] = False
    canon = (df.sort_values(["Standar", "Transaksi"], ascending=False)
             .drop_duplicates("_root").set_index("_root")["Hotel Name"])
    df["Hotel (Grup)"] = df["_root"].map(canon)
    df["Jumlah Varian"] = df.groupby("_root")["Hotel Name"].transform("size")
    df["Skor"] = [round(s, 1) if v > 1 else None for s, v in zip(best, df["Jumlah Varian"])]
    return (df.drop(columns="_root")
            .sort_values(["Kota (norm)", "Hotel (Grup)", "Transaksi"],
                         ascending=[True, True, False])
            .reset_index(drop=True))


# ---------------------------------------------------------------------------
# Fungsi bantu: format angka
# ---------------------------------------------------------------------------
def rupiah(x: float) -> str:
    """Format angka besar jadi ringkas: 1.2 M (miliar), 3.4 Jt (juta)."""
    if pd.isna(x):
        return "-"
    a = abs(x)
    if a >= 1e12:
        return f"Rp {x/1e12:,.2f} T"
    if a >= 1e9:
        return f"Rp {x/1e9:,.2f} M"
    if a >= 1e6:
        return f"Rp {x/1e6:,.2f} Jt"
    return f"Rp {x:,.0f}"


def rupiah_full(x: float) -> str:
    """Format rupiah lengkap tanpa disingkat, pemisah ribuan titik: Rp 492.105.123.456"""
    if pd.isna(x):
        return "-"
    return "Rp " + f"{x:,.0f}".replace(",", ".")


def invoice_key(sr: pd.Series) -> pd.Series:
    """Nomor invoice yang dinormalisasi (tanpa spasi, huruf besar); kosong/'-' diabaikan.
    Dipakai untuk menghitung invoice UNIK, karena satu invoice bisa terdiri dari beberapa baris."""
    k = sr.astype(str).str.strip().str.upper()
    return k.where(~k.isin(["", "-", "NAN", "NONE"]))


def angka(x: float) -> str:
    return f"{x:,.0f}".replace(",", ".")


# ---------------------------------------------------------------------------
# Membaca & membersihkan satu file
# ---------------------------------------------------------------------------
def _read_raw(source) -> pd.DataFrame:
    """Baca sheet pertama tanpa header. Pakai engine 'calamine' (cepat) bila ada."""
    for engine in ("calamine", None):
        try:
            if hasattr(source, "seek"):
                source.seek(0)
            return pd.read_excel(source, header=None, dtype=object, engine=engine)
        except (ImportError, ValueError):
            continue
    raise RuntimeError("Gagal membaca file Excel.")


def _find_header_row(raw: pd.DataFrame, max_scan: int = 30) -> int:
    """Cari baris header: baris yang berisi 'No' + 'Invoice No'/'Inv Date'.
    Jika tidak ketemu, ambil baris dengan sel terisi terbanyak."""
    n = min(max_scan, len(raw))
    for i in range(n):
        vals = {str(v).strip().lower() for v in raw.iloc[i].dropna()}
        if "no" in vals and ({"invoice no", "inv date"} & vals):
            return i
    return int(raw.head(n).notna().sum(axis=1).idxmax())


def _clean_columns(cols) -> list:
    out, seen = [], {}
    for i, c in enumerate(cols):
        name = re.sub(r"\s+", " ", str(c)).strip() if pd.notna(c) else f"Kolom_{i+1}"
        if name in seen:                      # hindari nama kolom ganda
            seen[name] += 1
            name = f"{name}_{seen[name]}"
        else:
            seen[name] = 0
        out.append(name)
    return out


def process_file(source, file_name: str):
    """Kembalikan (DataFrame bersih, info metadata) untuk satu file."""
    raw = _read_raw(source)
    h = _find_header_row(raw)

    # Baris judul di atas header (mis. "Period Invoice 01 Dec 2024 - 30 Jun 2025")
    title_lines = [str(v) for v in raw.iloc[:h, 0].dropna()]
    period = next((t for t in title_lines if "period" in t.lower()), "-")

    df = raw.iloc[h + 1:].copy()
    df.columns = _clean_columns(raw.iloc[h])
    df = df.dropna(how="all")

    # Buang baris total/subtotal di bagian bawah
    first_col = df.columns[0]
    is_total = df[first_col].astype(str).str.contains("total", case=False, na=False)
    df = df[~is_total]
    if "No" in df.columns:
        df = df[pd.to_numeric(df["No"], errors="coerce").notna()]

    # Konversi tipe data
    for c in DATE_COLS:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce")
    for c in NUMERIC_COLS:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    df.insert(0, "Source File", file_name)
    df = df.reset_index(drop=True)

    info = {
        "File": file_name,
        "Periode (dari judul)": period.replace("Period Invoice", "").strip(),
        "Baris Data": len(df),
        "Jumlah Kolom": df.shape[1] - 1,
        "Inv Date Min": df["Inv Date"].min() if "Inv Date" in df else None,
        "Inv Date Max": df["Inv Date"].max() if "Inv Date" in df else None,
    }
    return df, info


@st.cache_data(show_spinner=False, max_entries=20)
def load_uploaded(file_bytes: bytes, file_name: str):
    return process_file(io.BytesIO(file_bytes), file_name)


@st.cache_data(show_spinner=False, max_entries=20)
def load_from_path(path: str, mtime: float):  # mtime ikut jadi kunci cache
    return process_file(path, Path(path).name)


def merge_frames(frames: list, drop_dup: bool) -> pd.DataFrame:
    merged = pd.concat(frames, ignore_index=True, sort=False)
    if drop_dup:
        subset = [c for c in merged.columns if c not in ("Source File", "No")]
        merged = merged.drop_duplicates(subset=subset)
    if "Inv Date" in merged.columns:
        merged = merged.sort_values("Inv Date", kind="stable")
    merged = merged.reset_index(drop=True)
    if {"Room", "Night"} <= set(merged.columns):
        merged["Room Night"] = merged["Room"] * merged["Night"]
    if "Inv Date" in merged.columns:
        merged["Bulan"] = merged["Inv Date"].dt.to_period("M").dt.to_timestamp()
    return merged


# ---------------------------------------------------------------------------
# Ekspor
# ---------------------------------------------------------------------------
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
EXCEL_MAX_ROWS = 1_048_575
# Streamlit versi baru menerima fungsi sebagai 'data': file baru dibuat saat tombol diklik
LAZY_DL = "callable" in (st.download_button.__doc__ or "")


# Kolom nominal hasil ringkasan yang dibulatkan ke rupiah penuh (tanpa desimal)
MONEY_COL_RE = re.compile(r"^(Revenue|Sales AR|Profit|Rata-rata)", re.IGNORECASE)
# Sheet data mentah: angka dibiarkan apa adanya (tidak dibulatkan)
RAW_SHEETS = {"Merged Data", "Detail Transaksi"}


def round_money(x):
    """Bulatkan ke bilangan bulat terdekat (0,5 ke atas, seperti pembulatan rupiah biasa)."""
    import numpy as np
    return np.sign(x) * np.floor(np.abs(x) + 0.5)


def round_money_cols(df: pd.DataFrame) -> pd.DataFrame:
    """Kolom Revenue / Sales AR / Profit / Rata-rata ... -> bilangan bulat (tipe number).
    Kolom persen (mengandung '%') tidak diubah."""
    out = df.copy()
    for c in out.columns:
        if (MONEY_COL_RE.match(str(c)) and "%" not in str(c)
                and pd.api.types.is_numeric_dtype(out[c])):
            out[c] = round_money(out[c].astype(float)).astype("Int64")
    return out


def excel_book(sheets: dict, scorecard: list = None) -> bytes:
    """Buat workbook Excel yang rapi: header tebal, freeze baris pertama, autofilter,
    lebar kolom otomatis, format ribuan untuk angka & format tanggal.
    sheets    : {nama_sheet: DataFrame}
    scorecard : list (label, nilai, format) -> ditulis sebagai sheet 'Scorecard' paling depan;
                format: 'rp', 'num', 'text' atau None untuk baris judul bagian."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter", datetime_format="dd-mmm-yyyy",
                        date_format="dd-mmm-yyyy",
                        engine_kwargs={"options": {"strings_to_urls": False}}) as w:
        wb = w.book
        f_head = wb.add_format({"bold": True, "bg_color": "#0D7FCC", "font_color": "#FFFFFF",
                                "border": 1, "text_wrap": True, "valign": "vcenter"})
        f_num = wb.add_format({"num_format": "#,##0"})
        f_pct = wb.add_format({"num_format": "0.00"})

        if scorecard:
            ws = wb.add_worksheet("Scorecard")
            f_title = wb.add_format({"bold": True, "font_size": 14, "font_color": "#0D7FCC"})
            f_sec = wb.add_format({"bold": True, "bg_color": "#E3F2FB", "border": 1})
            f_lbl = wb.add_format({"border": 1})
            f_rp = wb.add_format({"num_format": '"Rp "#,##0', "border": 1, "bold": True})
            f_n = wb.add_format({"num_format": "#,##0", "border": 1, "bold": True})
            f_txt = wb.add_format({"border": 1, "text_wrap": True})
            ws.set_column(0, 0, 30)
            ws.set_column(1, 1, 70)
            r = 0
            for label, value, fmt in scorecard:
                if fmt == "title":
                    ws.write(r, 0, label, f_title)
                    r += 2
                    continue
                if fmt is None:          # judul bagian
                    if r:
                        r += 1
                    ws.merge_range(r, 0, r, 1, label, f_sec)
                else:
                    ws.write(r, 0, label, f_lbl)
                    if fmt == "rp":
                        ws.write_number(r, 1, int(round_money(float(value or 0))), f_rp)
                    elif fmt == "num":
                        ws.write_number(r, 1, float(value or 0), f_n)
                    else:
                        ws.write(r, 1, "-" if value is None else str(value), f_txt)
                r += 1

        for name, df in sheets.items():
            if df is None or not len(df.columns):
                continue
            sh = name[:31]
            df = df.head(EXCEL_MAX_ROWS)
            if sh not in RAW_SHEETS:
                df = round_money_cols(df)
            df.to_excel(w, sheet_name=sh, index=False)
            ws = w.sheets[sh]
            for j, col in enumerate(df.columns):
                ws.write(0, j, str(col), f_head)
                s = df[col]
                sample = s.head(300).astype(str)
                width = min(max(len(str(col)), sample.str.len().max() if len(sample) else 0) + 2, 50)
                fmt = None
                if pd.api.types.is_numeric_dtype(s) and not pd.api.types.is_bool_dtype(s):
                    fmt = f_pct if "%" in str(col) else (None if "Rating" in str(col) else f_num)
                    width = max(width, 14)
                elif pd.api.types.is_datetime64_any_dtype(s):
                    width = max(len(str(col)) + 2, 13)
                ws.set_column(j, j, width, fmt)
            # Baris "TOTAL" (bila ada) dicetak tebal
            if len(df.columns) and len(df):
                _tot = df.iloc[:, 0].astype(str).str.upper().eq("TOTAL").to_numpy().nonzero()[0]
                if len(_tot):
                    f_tot = wb.add_format({"bold": True, "bg_color": "#F3F8FD", "num_format": "#,##0",
                                           "top": 1})
                    for _i in _tot:
                        ws.set_row(int(_i) + 1, None, f_tot)
            ws.freeze_panes(1, 0)
            if len(df):
                ws.autofilter(0, 0, len(df), len(df.columns) - 1)
    return buf.getvalue()


def to_excel_bytes(df: pd.DataFrame, summary: pd.DataFrame, extra: dict = None) -> bytes:
    sheets = {"Merged Data": df, "Ringkasan File": summary}
    sheets.update({k: v for k, v in (extra or {}).items() if v is not None and len(v)})
    return excel_book(sheets)


def excel_download(box, label: str, make, file_name: str, key: str, **kw):
    """Tombol unduh Excel. 'make' = fungsi tanpa argumen yang menghasilkan bytes.
    Di Streamlit versi baru file baru dibuat saat tombol diklik (halaman tidak melambat)."""
    extra = {"on_click": "ignore"} if LAZY_DL else {}
    box.download_button(label, make if LAZY_DL else make(), file_name=file_name,
                        mime=XLSX_MIME, key=key, type="primary", **extra, **kw)


# ---------------------------------------------------------------------------
# SIDEBAR: sumber data
# ---------------------------------------------------------------------------
st.sidebar.title("📂 Sumber Data")
mode = st.sidebar.radio("Mode input", ["Upload file", "Browse folder"],
                        help="Upload: pilih file dari komputer. "
                             "Browse: baca semua Excel di sebuah folder lokal/server.")

frames, infos, errors = [], [], []

if mode == "Upload file":
    uploads = st.sidebar.file_uploader("Pilih satu atau beberapa file Excel",
                                       type=[e.strip(".") for e in EXCEL_EXT],
                                       accept_multiple_files=True)
    sources = [(u.name, u) for u in (uploads or [])]
else:
    folder = st.sidebar.text_input("Path folder", value=str(Path.cwd()))
    p = Path(folder).expanduser()
    if p.is_dir():
        found = sorted(f for f in p.iterdir() if f.suffix.lower() in EXCEL_EXT
                       and not f.name.startswith("~$")
                       and f.name.lower() != MASTER_FILE)  # file master bukan data penjualan
        if found:
            chosen = st.sidebar.multiselect("File yang ditemukan", [f.name for f in found],
                                            default=[f.name for f in found])
            sources = [(n, p / n) for n in chosen]
        else:
            st.sidebar.warning("Tidak ada file Excel di folder ini.")
            sources = []
    else:
        st.sidebar.error("Folder tidak ditemukan.")
        sources = []

drop_dup = st.sidebar.checkbox("Hapus baris duplikat persis", value=False,
                               help="Berguna jika periode antar-file saling tumpang tindih.")

st.title("📊 Merge Excel & Dashboard Penjualan Hotel")

if not sources:
    st.info("Pilih file lewat **Upload file** atau **Browse folder** di sidebar untuk memulai.")
    st.markdown(
        "- File boleh berbeda periode; aplikasi otomatis mendeteksi baris header "
        "dan membuang baris *Grand Total*.\n"
        "- Setiap baris diberi kolom **Source File** agar asal datanya tetap terlacak."
    )
    st.stop()

# ---------------------------------------------------------------------------
# Proses semua file
# ---------------------------------------------------------------------------
progress = st.progress(0.0, text="Membaca file...")
t0 = time.time()
for i, (name, src) in enumerate(sources, 1):
    progress.progress((i - 1) / len(sources), text=f"Membaca {name} ({i}/{len(sources)})...")
    try:
        if mode == "Upload file":
            df_i, info_i = load_uploaded(src.getvalue(), name)
        else:
            df_i, info_i = load_from_path(str(src), src.stat().st_mtime)
        frames.append(df_i)
        infos.append(info_i)
    except Exception as e:  # noqa: BLE001
        errors.append(f"{name}: {e}")
progress.empty()

for err in errors:
    st.error(f"Gagal membaca {err}")
if not frames:
    st.stop()

merge_key = (tuple((i["File"], i["Baris Data"]) for i in infos), drop_dup)
if st.session_state.get("merged", (None,))[0] != merge_key:
    st.session_state["merged"] = (merge_key, merge_frames(frames, drop_dup))
data = st.session_state["merged"][1]
summary = pd.DataFrame(infos)

# Peringatan struktur kolom yang berbeda antar-file
col_sets = {f["Source File"].iat[0]: set(f.columns) for f in frames if len(f)}
all_cols = set().union(*col_sets.values())
diff = {k: sorted(all_cols - v) for k, v in col_sets.items() if all_cols - v}
if diff:
    with st.expander("⚠️ Struktur kolom antar-file tidak sama", expanded=False):
        for k, v in diff.items():
            st.write(f"**{k}** tidak memiliki kolom: {', '.join(v)}")

# Peringatan periode tumpang tindih
if {"Inv Date Min", "Inv Date Max"} <= set(summary.columns) and len(summary) > 1:
    s = summary.sort_values("Inv Date Min").reset_index(drop=True)
    overlaps = [f"{s.loc[i-1,'File']} ↔ {s.loc[i,'File']}"
                for i in range(1, len(s)) if s.loc[i, "Inv Date Min"] <= s.loc[i-1, "Inv Date Max"]]
    if overlaps:
        st.warning("Periode tanggal invoice tumpang tindih: " + "; ".join(overlaps)
                   + ". Pertimbangkan mengaktifkan 'Hapus baris duplikat'.")

st.caption(f"{len(frames)} file digabung → {angka(len(data))} baris, "
           f"{data.shape[1]} kolom · diproses dalam {time.time()-t0:.1f} detik")

# ---------------------------------------------------------------------------
# SIDEBAR: standarisasi nama hotel (master STADARD NAME)
# ---------------------------------------------------------------------------
@st.cache_data(show_spinner="Membaca master nama hotel...")
def load_master_bytes(b: bytes) -> pd.DataFrame:
    return read_hotel_master(io.BytesIO(b))


@st.cache_data(show_spinner="Membaca master nama hotel...")
def load_master_path(path: str, mtime: float) -> pd.DataFrame:
    return read_hotel_master(path)


@st.cache_data(show_spinner=False)
def cached_standardize(names: tuple, master: pd.DataFrame) -> pd.DataFrame:
    return standardize_hotels(pd.Series(names), master)


st.sidebar.divider()
st.sidebar.title("📘 Standarisasi Nama Hotel")
std_on = "Hotel Name" in data.columns and st.sidebar.toggle(
    "Gunakan nama standar (STADARD NAME)", value=True,
    help="Kolom 'Hotel Name' diganti dengan nama standar dari file master. "
         "Nama asli tetap disimpan di kolom 'Hotel Name (Asli)'.")
std_table = None  # hasil pencocokan per nama asli (untuk tab Grup Hotel & ekspor)

if std_on:
    master_up = st.sidebar.file_uploader(
        "File master hotel (opsional)", type=["xlsx", "xls"], key="master_up",
        help=f"Kosongkan untuk memakai '{MASTER_FILE}' di folder aplikasi. "
             "Wajib berisi kolom HOTEL NAME dan STADARD NAME.")
    master, default_path = None, Path(__file__).parent / MASTER_FILE
    try:
        if master_up is not None:
            master = load_master_bytes(master_up.getvalue())
        elif default_path.exists():
            master = load_master_path(str(default_path), default_path.stat().st_mtime)
        else:
            st.sidebar.warning(f"Upload file master, atau taruh '{MASTER_FILE}' "
                               "di folder yang sama dengan app.py.")
    except Exception as e:  # noqa: BLE001
        st.sidebar.error(f"Gagal membaca master: {e}")

    if master is not None:
        std_table = cached_standardize(tuple(pd.unique(data["Hotel Name"].dropna().astype(str))),
                                       master)
        orig = data["Hotel Name"]
        key_ = orig.astype(str)
        std_name = key_.map(std_table["Standard Name"])
        data = data.assign(**{
            "Hotel Name": std_name.fillna(orig).where(orig.notna(), orig),
            "Hotel Name (Asli)": orig,
            "Hotel Chain": key_.map(std_table["Hotel Chain"]),
            "Hotel Rating": key_.map(std_table["Hotel Rating"]),
        })
        # Taruh kolom tambahan tepat setelah 'Hotel Name'
        cols_ = [c for c in data.columns if c not in ("Hotel Name (Asli)", "Hotel Chain", "Hotel Rating")]
        pos = cols_.index("Hotel Name") + 1
        data = data[cols_[:pos] + ["Hotel Name (Asli)", "Hotel Chain", "Hotel Rating"] + cols_[pos:]]

        n_names = len(std_table)
        n_hit = int(std_table["Standard Name"].notna().sum())
        rows_hit = std_name.notna().mean() * 100 if len(data) else 0
        st.sidebar.caption(
            f"Master: sheet '{master.attrs.get('sheet', '-')}', {angka(len(master))} baris. "
            f"{angka(n_hit)} dari {angka(n_names)} nama hotel cocok "
            f"({rows_hit:.1f}% baris transaksi). Nama yang tidak cocok tetap memakai nama asli — "
            "lihat daftarnya di tab **Grup Hotel**.")

# ---------------------------------------------------------------------------
# Penyeragaman penulisan nama (otomatis) untuk Supplier Name & Hotel Name
#   Pola regex (tidak peka huruf besar/kecil) -> penulisan baku. Tambahkan baris
#   baru di NAME_FIXES bila ada merek lain yang perlu diseragamkan.
# ---------------------------------------------------------------------------
NAME_FIXES = {
    r"[eé]L\s+Hotel": "éL Hotel",       # 'EL HOTEL', 'El Hotel', 'él hotel' -> 'éL Hotel'
    r"[eé]L\s+Royale": "éL Royale",     # 'EL ROYALE' -> 'éL Royale'
    r"Meli[aá]": "Meliá",               # 'MELIA', 'Melia' -> 'Meliá'
}

NAME_FIX_COLS = ["Supplier Name", "Hotel Name"]


def fix_names(sr: pd.Series) -> pd.Series:
    """Terapkan NAME_FIXES (tidak peka huruf besar/kecil) pada sebuah kolom teks."""
    out = sr.astype("string")
    for pat, repl in NAME_FIXES.items():
        out = out.str.replace(rf"\b{pat}\b", repl, regex=True, flags=re.IGNORECASE)
    return out.astype(object).where(sr.notna(), sr)


# Singkatan yang tetap huruf besar setelah dibuat Proper
PROPER_KEEP_UPPER = ["PT", "CV", "CBT", "MG", "TBO", "BUMN", "PLN", "PGN", "KAI", "BRI", "BNI", "BCA", "BTN",
                     "BPJS", "SKK", "LNG", "UD", "TNI", "POLRI", "RI", "II", "III", "IV"]


def proper_case(sr: pd.Series) -> pd.Series:
    """Ubah teks jadi Proper Case ('CBT PERTAMINA' -> 'CBT Pertamina',
    'pt bank mandiri (persero) tbk' -> 'PT Bank Mandiri (Persero) Tbk').
    Singkatan di PROPER_KEEP_UPPER tetap huruf besar; nilai kosong tidak diubah."""
    s_ = sr.astype("string").str.strip().str.replace(r"\s+", " ", regex=True).str.lower()
    s_ = s_.str.replace(r"(?<![\w'’])([a-z])", lambda m: m.group(1).upper(), regex=True)
    keep = "|".join(w.capitalize() for w in PROPER_KEEP_UPPER)
    s_ = s_.str.replace(rf"\b({keep})\b", lambda m: m.group(1).upper(), regex=True)
    return s_.astype(object).where(sr.notna(), sr)


def standard_supplier(sr: pd.Series) -> pd.Series:
    """Supplier Name baku: nama di SUPPLIER_STANDARD_NAMES -> penulisan baku, lainnya Proper.
    Setelah itu NAME_FIXES diterapkan lagi agar 'éL Hotel', 'Meliá', dll. tidak berubah
    oleh Proper Case (mis. 'éL Hotel' -> 'él Hotel')."""
    std = {_compact(k): v for k, v in SUPPLIER_STANDARD_NAMES.items()}
    mapped = sr.astype(str).map(_compact).map(std)
    out = mapped.fillna(proper_case(sr))
    # domain seperti 'Xxx.Com' tetap ditulis '.com'
    out = out.astype("string").str.replace(r"\.(Com|Co\.Id|Id|Net)\b",
                                           lambda m: m.group(0).lower(), regex=True)
    out = fix_names(out.astype(object).where(sr.notna(), sr))
    return out.astype(object).where(sr.notna(), sr)


_fix_cols = [c for c in NAME_FIX_COLS if c in data.columns]
if _fix_cols:
    data = data.assign(**{c: fix_names(data[c]) for c in _fix_cols})

# ---------------------------------------------------------------------------
# SIDEBAR: merge hotel mirip (fuzzy)
# ---------------------------------------------------------------------------
@st.cache_data(show_spinner="Mengelompokkan nama hotel yang mirip...")
def cached_hotel_groups(hotels: pd.DataFrame, threshold: int, same_city: bool,
                        stopwords: tuple) -> pd.DataFrame:
    return build_hotel_groups(hotels, threshold, same_city, list(stopwords))


def hotel_key(city, name) -> pd.Series:
    return city.astype(str) + "||" + name.astype(str)


st.sidebar.divider()
st.sidebar.title("🏨 Merge Hotel Mirip")
has_hotel = {"Hotel Name", "Hotel City"} <= set(data.columns)
hotel_on = has_hotel and st.sidebar.toggle(
    "Gabungkan nama hotel yang mirip", value=True,
    help="Menyatukan variasi penulisan nama hotel yang sama, misalnya "
         "'Crowne Plaza Bandung' dan 'CROWNE PLAZA BANDUNG by IHG'.")
overrides = st.session_state.setdefault("hotel_overrides", {})
mapping = None

if hotel_on:
    threshold = st.sidebar.slider("Ambang kemiripan nama (%)", 80, 100, 90,
                                  help="Semakin tinggi semakin ketat. 100 = hanya yang identik "
                                       "setelah normalisasi (beda huruf besar/kecil, tanda baca, "
                                       "'&'/'and', nama kota, kata 'Hotel', dll.).")
    same_city = st.sidebar.checkbox("Lokasi: hanya gabungkan jika kota sama", value=True,
                                    help="Disarankan aktif. Jika dimatikan, hotel bernama mirip "
                                         "di kota berbeda juga bisa tergabung.")
    with st.sidebar.expander("Kata yang diabaikan saat mencocokkan"):
        sw_text = st.text_area("Satu per baris (boleh pola seperti 'managed by .*')",
                               "\n".join(DEFAULT_STOPWORDS), height=200)
    stopwords = tuple(w.strip().lower() for w in sw_text.splitlines() if w.strip())

    hotels = (data.groupby(["Hotel Name", "Hotel City"], dropna=False).size()
              .reset_index(name="Transaksi"))
    # Nama yang sudah standar diprioritaskan menjadi nama grup
    std_names = set(std_table["Standard Name"].dropna()) if std_table is not None else set()
    hotels["Standar"] = hotels["Hotel Name"].isin(std_names)
    mapping = cached_hotel_groups(hotels, threshold, same_city, stopwords).copy()

    # Koreksi manual dari pengguna (tab Grup Hotel) selalu menang
    mapping["Key"] = hotel_key(mapping["Hotel City"], mapping["Hotel Name"])
    mapping["Diubah Manual"] = mapping["Key"].isin(overrides)
    mapping["Hotel (Grup)"] = (mapping["Key"].map(overrides).fillna(mapping["Hotel (Grup)"])
                               .map(html.unescape))
    mapping["Jumlah Varian"] = mapping.groupby("Hotel (Grup)")["Hotel Name"].transform("size")

    data = data.drop(columns=["Hotel (Grup)"], errors="ignore").merge(
        mapping[["Hotel Name", "Hotel City", "Hotel (Grup)"]],
        on=["Hotel Name", "Hotel City"], how="left")
    data["Hotel (Grup)"] = data["Hotel (Grup)"].fillna(data["Hotel Name"])
    st.sidebar.caption(f"{angka(len(mapping))} nama asli → "
                       f"{angka(mapping['Hotel (Grup)'].nunique())} hotel setelah digabung")
elif has_hotel:
    data = data.assign(**{"Hotel (Grup)": data["Hotel Name"]})

# ---------------------------------------------------------------------------
# SIDEBAR: merge nama customer
# ---------------------------------------------------------------------------
DEFAULT_CUSTOMER_RULES = [
    {"Nama Customer Asli": "CBT PERTAMINA (HOTEL)", "Digabung Menjadi": "CBT PERTAMINA"},
    {"Nama Customer Asli": "PERTAMINA ENERGY TERMINAL (CBT)", "Digabung Menjadi": "CBT PERTAMINA"},
]


def customer_norm(sr: pd.Series) -> pd.Series:
    """Normalisasi untuk pencocokan: huruf besar, spasi rapi, tanpa spasi sebelum '('.
    Jadi 'CBT PERTAMINA (HOTEL)', 'CBT PERTAMINA(HOTEL)' dan 'cbt pertamina (hotel)' sama."""
    s_ = sr.astype(str).str.upper().str.replace(r"\s+", " ", regex=True).str.strip()
    return s_.str.replace(r"\s*\(\s*", "(", regex=True).str.replace(r"\s*\)", ")", regex=True)


# Aturan merge customer diterapkan otomatis (tanpa tampilan di halaman)
if "Customer Name" in data.columns:
    rule_map = {customer_norm(pd.Series([r["Nama Customer Asli"]]))[0]: r["Digabung Menjadi"]
                for r in DEFAULT_CUSTOMER_RULES}
    orig = data["Customer Name"]
    mapped = customer_norm(orig).map(rule_map)
    data = data.assign(**{"Customer Name": proper_case(mapped.fillna(orig)),
                          "Customer Name (Asli)": orig})

# ---------------------------------------------------------------------------
# Merge kolom Invoice To (otomatis, tidak peka huruf besar/kecil & spasi)
#   Nama asli tetap disimpan di kolom 'Invoice To (Asli)'.
# ---------------------------------------------------------------------------
INVOICE_TO_RULES = {
    "CBT PERTAMINA (HOTEL)": "CBT PERTAMINA",
}
inv_to_col = next((c for c in data.columns
                   if re.sub(r"\s+", " ", str(c)).strip().lower() == "invoice to"), None)
if inv_to_col and INVOICE_TO_RULES:
    inv_map = dict(zip(customer_norm(pd.Series(list(INVOICE_TO_RULES))),
                       INVOICE_TO_RULES.values()))
    orig = data[inv_to_col]
    mapped = customer_norm(orig).map(inv_map)
    data = data.drop(columns=[f"{inv_to_col} (Asli)"], errors="ignore").copy()
    data[inv_to_col] = mapped.fillna(orig)
    data.insert(data.columns.get_loc(inv_to_col) + 1, f"{inv_to_col} (Asli)", orig)

# ---------------------------------------------------------------------------
# Kategori Supplier (otomatis dari Supplier Name, tidak peka huruf besar/kecil)
# ---------------------------------------------------------------------------
if "Supplier Name" in data.columns:
    data = data.drop(columns=["Supplier Category", "Supplier Name (Asli)"], errors="ignore")
    _sup_orig = data["Supplier Name"]
    data = data.assign(**{"Supplier Name": standard_supplier(_sup_orig)})
    pos_ = data.columns.get_loc("Supplier Name") + 1
    data.insert(pos_, "Supplier Category", supplier_category(data["Supplier Name"]))
    data.insert(pos_ + 1, "Supplier Name (Asli)", _sup_orig)

# ---------------------------------------------------------------------------
# SIDEBAR: filter
# ---------------------------------------------------------------------------
st.sidebar.divider()
st.sidebar.title("🔎 Filter")
fdata = data

if "Inv Date" in data.columns and data["Inv Date"].notna().any():
    dmin, dmax = data["Inv Date"].min().date(), data["Inv Date"].max().date()
    rng = st.sidebar.date_input("Rentang Inv Date", value=(dmin, dmax),
                                min_value=dmin, max_value=dmax)
    if isinstance(rng, (list, tuple)) and len(rng) == 2:
        fdata = fdata[fdata["Inv Date"].between(pd.Timestamp(rng[0]),
                                                pd.Timestamp(rng[1]) + pd.Timedelta(days=1)
                                                - pd.Timedelta(seconds=1))]


def multi_filter(df: pd.DataFrame, col: str, label: str, box=st.sidebar,
                 placeholder: str = "Semua") -> pd.DataFrame:
    """Filter multiselect. Opsi diambil dari seluruh data agar pilihan tidak hilang
    saat filter lain berubah. Kosong = semua."""
    if col not in df.columns:
        return df
    opts = sorted(data[col].dropna().astype(str).unique())
    sel = box.multiselect(label, opts, placeholder=placeholder)
    return df[df[col].astype(str).isin(sel)] if sel else df


fdata = multi_filter(fdata, "Source File", "Source File")
fdata = multi_filter(fdata, "Branch", "Branch")
fdata = multi_filter(fdata, "Customer Name", "Customer")
fdata = multi_filter(fdata, "Supplier Category", "Kategori Supplier")
fdata = multi_filter(fdata, "Hotel City", "Kota Hotel")
top_n = st.sidebar.slider("Top N pada grafik", 5, 30, 10)

if fdata.empty:
    st.warning("Tidak ada data yang cocok dengan filter.")
    st.stop()

# ---------------------------------------------------------------------------
# Tabs
# ---------------------------------------------------------------------------
tab_dash, tab_find, tab_data, tab_hotel, tab_file = st.tabs(
    ["📈 Dashboard", "🔍 Filter per Hotel", "🗂️ Data Gabungan", "🏨 Grup Hotel",
     "📁 Ringkasan File"])


def market_type(sr: pd.Series) -> pd.Series:
    """Kelompokkan nilai Product Type menjadi Domestic / International / Lainnya
    berdasarkan kata kunci (tidak peka huruf besar/kecil), mis.
    'Hotel Domestic', 'HOTEL DOM' -> Domestic; 'Hotel International', 'Hotel Intl' -> International."""
    s = sr.astype(str).str.lower()
    out = pd.Series("Lainnya", index=sr.index)
    out[s.str.contains(r"\bdom|domestic|domestik|lokal|local", regex=True, na=False)] = "Domestic"
    out[s.str.contains(r"\bint\b|\bintl|inter|internasional|overseas|luar negeri", regex=True,
                       na=False)] = "International"
    return out


def unmatched_hotels() -> pd.DataFrame:
    """Nama hotel asli yang tidak ditemukan di master, beserta jumlah transaksinya."""
    if std_table is None:
        return pd.DataFrame()
    miss = std_table.index[std_table["Standard Name"].isna()]
    d = data[data["Hotel Name (Asli)"].astype(str).isin(miss)]
    return (d.groupby(["Hotel Name (Asli)", "Hotel City"], dropna=False)
            .agg(Transaksi=("Hotel Name (Asli)", "size"), **{"Sales AR": ("Sales AR", "sum")})
            .reset_index().rename(columns={"Hotel Name (Asli)": "HOTEL NAME"})
            .assign(**{"STADARD NAME": ""})
            .sort_values("Transaksi", ascending=False).reset_index(drop=True))


KPI_CSS = """<style>
.kpi-grid {display:grid; grid-template-columns:repeat(3,minmax(0,1fr)); gap:14px; margin:8px 0 6px;}
.kpi-grid.g5 {grid-template-columns:repeat(5,minmax(0,1fr)); gap:12px; margin-bottom:18px;}
.kpi-grid.g5 .val {font-size:clamp(1.15rem,1.6vw,1.55rem);}
@media (max-width: 1200px) {.kpi-grid.g5 {grid-template-columns:repeat(3,minmax(0,1fr));}}
@media (max-width: 900px) {.kpi-grid {grid-template-columns:repeat(2,minmax(0,1fr));}
                           .kpi-grid.g5 {grid-template-columns:repeat(2,minmax(0,1fr));}}
@media (max-width: 560px) {.kpi-grid {grid-template-columns:1fr;}}
.kpi {position:relative; overflow:hidden; border-radius:16px; padding:16px 18px 14px;
      background:var(--kpi-bg); border:1px solid var(--kpi-bd);
      box-shadow:0 1px 2px rgba(16,24,40,.04), 0 4px 14px rgba(16,24,40,.06);
      transition:transform .15s ease, box-shadow .15s ease;}
.kpi:hover {transform:translateY(-2px); box-shadow:0 2px 4px rgba(16,24,40,.06), 0 10px 24px rgba(16,24,40,.10);}
.kpi::before {content:""; position:absolute; inset:0 auto 0 0; width:5px; background:var(--kpi-c);}
.kpi .top {display:flex; align-items:center; gap:10px; margin-bottom:10px;}
.kpi .ic {width:36px; height:36px; border-radius:10px; display:grid; place-items:center;
          font-size:1.15rem; background:var(--kpi-soft);}
.kpi .lbl {font-size:.78rem; font-weight:600; letter-spacing:.05em; text-transform:uppercase;
           color:var(--kpi-c);}
.kpi .val {font-size:clamp(1.35rem,2.1vw,1.95rem); font-weight:700; line-height:1.15;
           color:var(--kpi-ink); word-break:break-word; font-variant-numeric:tabular-nums;}
.kpi .sub {font-size:.82rem; color:var(--kpi-mute); margin-top:6px;}
.kpi .bar {height:6px; border-radius:99px; background:var(--kpi-soft); margin-top:10px; overflow:hidden;}
.kpi .bar > span {display:block; height:100%; border-radius:99px; background:var(--kpi-c);}
.kpi {--kpi-bg:#ffffff; --kpi-bd:#e8ebf0; --kpi-ink:#434343; --kpi-mute:#7A7A7A;}
.kpi.blue   {--kpi-c:#0D7FCC; --kpi-soft:#E3F1FB;}
.kpi.green  {--kpi-c:#1BA0E2; --kpi-soft:#E4F5FC;}
.kpi.orange {--kpi-c:#1494C6; --kpi-soft:#E2F2F8;}
.kpi.purple {--kpi-c:#434343; --kpi-soft:#EDEDED;}
.kpi.teal   {--kpi-c:#0D7FCC; --kpi-soft:#E3F1FB;}
.kpi.rose   {--kpi-c:#1BA0E2; --kpi-soft:#E4F5FC;}
@media (prefers-color-scheme: dark) {
  .kpi {--kpi-bg:#1b1f2a; --kpi-bd:#2b3140; --kpi-ink:#f1f3f7; --kpi-mute:#B5B5B5;}
  .kpi.blue{--kpi-soft:#13324A;} .kpi.green{--kpi-soft:#14384A;} .kpi.orange{--kpi-soft:#133443;}
  .kpi.purple{--kpi-soft:#333333;} .kpi.teal{--kpi-soft:#13324A;} .kpi.rose{--kpi-soft:#14384A;}
}
.kpi-cap {font-size:.8rem; color:#8C8C8C; margin:2px 0 14px;}
</style>"""


def kpi_cards(items: list, caption: str = "", grid_class: str = ""):
    """Scorecard berwarna. items: dict(icon, label, value, sub, tone, bar=None (0-100), tip=None)."""
    cards = []
    for it in items:
        tip = f' title="{html.escape(it["tip"])}"' if it.get("tip") else ""
        bar = (f'<div class="bar"><span style="width:{it["bar"]:.1f}%"></span></div>'
               if it.get("bar") is not None else "")
        cards.append(
            f'<div class="kpi {it.get("tone", "blue")}"{tip}>'
            f'<div class="top"><div class="ic">{it.get("icon", "")}</div>'
            f'<div class="lbl">{html.escape(it["label"])}</div></div>'
            f'<div class="val">{html.escape(str(it["value"]))}</div>'
            f'<div class="sub">{html.escape(it.get("sub", ""))}</div>{bar}</div>')
    cap = f'<div class="kpi-cap">{html.escape(caption)}</div>' if caption else ""
    st.markdown(KPI_CSS + f'<div class="kpi-grid {grid_class}">' + "".join(cards) + "</div>" + cap,
                unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Ekspor halaman Dashboard ke PDF 1 halaman (A4 portrait), dibuat dengan matplotlib
# ---------------------------------------------------------------------------
def _rp_short(x, _pos=None) -> str:
    a = abs(x)
    if a >= 1e12:
        return f"Rp {x / 1e12:,.1f} T".replace(".", ",")
    if a >= 1e9:
        return f"Rp {x / 1e9:,.1f} M".replace(".", ",")
    if a >= 1e6:
        return f"Rp {x / 1e6:,.0f} Jt"
    return f"Rp {x:,.0f}".replace(",", ".")


def _cut(s, n: int = 34) -> str:
    s = str(s)
    return s if len(s) <= n else s[: n - 1] + "…"


def dashboard_pdf(dash: dict) -> bytes:
    """Gambar ulang isi tab Dashboard (scorecard, tren, Top 10 customer, komposisi Room Night,
    Domestic vs International) ke satu halaman PDF A4."""
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib.patches import FancyBboxPatch, Rectangle
    from matplotlib.ticker import FuncFormatter

    INK, MUTE, LINE = "#434343", "#7A7A7A", "#e3e6ec"
    plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 7, "axes.edgecolor": LINE,
                         "axes.labelcolor": MUTE, "xtick.color": MUTE, "ytick.color": MUTE})
    fig = plt.figure(figsize=(8.27, 11.69))
    bg = fig.add_axes([0, 0, 1, 1]); bg.axis("off"); bg.set_xlim(0, 1); bg.set_ylim(0, 1)
    L, R = 0.045, 0.955

    def sec(y, title, sub="", x=L, x_end=None):
        """Judul bagian (kiri) + keterangan kecil rata kanan agar tidak bertumpuk."""
        bg.text(x, y, title, fontsize=9.5, weight="bold", color=INK, va="bottom")
        if sub:
            bg.text(x_end if x_end is not None else R, y + 0.001, sub, fontsize=6.3,
                    color=MUTE, va="bottom", ha="right")

    # Header
    bg.text(L, 0.968, "Dashboard Penjualan Hotel", fontsize=16, weight="bold", color=INK)
    bg.text(L, 0.952, f"Periode Inv Date: {dash.get('periode', '-')}  ·  {dash.get('info', '')}",
            fontsize=7.5, color=MUTE)
    bg.text(R, 0.968, pd.Timestamp.now().strftime("Dicetak %d %b %Y %H:%M"), fontsize=7,
            color=MUTE, ha="right")
    bg.add_line(plt.Line2D([L, R], [0.944, 0.944], color=LINE, lw=1))

    # Scorecard 3 x 2
    kpis = dash.get("kpi", [])
    gap, w, h = 0.014, (R - L - 2 * 0.014) / 3, 0.058
    for i, (label, value, sub, color) in enumerate(kpis[:6]):
        r, c = divmod(i, 3)
        x, y = L + c * (w + gap), 0.874 - r * (h + 0.01)
        bg.add_patch(FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0,rounding_size=0.008",
                                    fc="white", ec=LINE, lw=0.8))
        bg.add_patch(Rectangle((x, y + 0.004), 0.004, h - 0.008, fc=color, ec="none"))
        bg.text(x + 0.016, y + h - 0.014, label.upper(), fontsize=6.3, weight="bold", color=color)
        bg.text(x + 0.016, y + h - 0.036, value, fontsize=12.5, weight="bold", color=INK)
        bg.text(x + 0.016, y + 0.008, _cut(sub, 52), fontsize=6, color=MUTE)

    # Tren bulanan
    m = dash.get("trend")
    sec(0.775, "Tren Bulanan")
    if m is not None and len(m):
        ax = fig.add_axes([0.10, 0.615, 0.79, 0.15])
        ax.bar(m["Label"], m["Sales"], color=PALETTE[0], width=0.6, alpha=.9, label="Sales AR")
        ax.yaxis.set_major_formatter(FuncFormatter(_rp_short))
        ax.grid(axis="y", color=LINE, lw=0.6); ax.set_axisbelow(True)
        for s_ in ("top", "right", "left"):
            ax.spines[s_].set_visible(False)
        ax2 = ax.twinx()
        ax2.plot(m["Label"], m["Profit"], color=PALETTE[2], lw=2, marker="o", ms=3.5, label="Profit")
        ax2.yaxis.set_major_formatter(FuncFormatter(_rp_short))
        for s_ in ("top", "left", "right"):
            ax2.spines[s_].set_visible(False)
        ax.tick_params(length=0, labelsize=6.5); ax2.tick_params(length=0, labelsize=6.5)
        h1, l1 = ax.get_legend_handles_labels(); h2, l2 = ax2.get_legend_handles_labels()
        ax.legend(h1 + h2, l1 + l2, loc="lower right", bbox_to_anchor=(1.075, 1.03), ncol=2,
                  frameon=False, fontsize=6.5)

    # Tabel Top 10 Customer
    def draw_table(rect, df, fmts, col_w, font=6.3):
        ax = fig.add_axes(rect); ax.axis("off")
        cells = [[fmts.get(c, str)(v) for c, v in zip(df.columns, row)] for row in df.itertuples(index=False)]
        tb = ax.table(cellText=cells, colLabels=list(df.columns), colWidths=col_w,
                      loc="upper center", cellLoc="right", colLoc="right")
        tb.auto_set_font_size(False); tb.set_fontsize(font)
        nrows = len(cells) + 1
        for (r_, c_), cell in tb.get_celld().items():
            cell.set_edgecolor(LINE); cell.set_linewidth(0.5)
            cell.set_height(1 / max(nrows, 1))
            if df.columns[c_] in ("Customer", "Kategori"):
                cell._loc = "left"; cell.set_text_props(ha="left")
            if r_ == 0:
                cell.set_facecolor("#f3f5f8"); cell.set_text_props(weight="bold", color=MUTE)
            else:
                cell.set_facecolor("white" if r_ % 2 else "#fafbfc")
        return ax

    pct = lambda v: "-" if pd.isna(v) else f"{v:.2f}%".replace(".", ",")
    num = lambda v: "-" if pd.isna(v) else angka(v)
    dec1 = lambda v: "-" if pd.isna(v) else f"{v:.1f}".replace(".", ",")
    for key, x0 in (("cust_sales", L), ("cust_profit", 0.515)):
        if key not in dash:
            continue
        title, sub, t, cap = dash[key]
        sec(0.565, title, sub.split(" · ")[0], x=x0, x_end=x0 + 0.44)
        t = t.copy()
        t["Customer"] = t["Customer"].map(lambda s: _cut(s, 26 if len(t.columns) > 4 else 30))
        t = t.rename(columns={"Rank": "#", "Avg Room Night / Bulan": "Avg RN/Bln"})
        ncol = len(t.columns)
        cw = {4: [0.06, 0.50, 0.26, 0.18], 5: [0.06, 0.39, 0.15, 0.20, 0.20]}.get(
            ncol, [1 / ncol] * ncol)
        draw_table([x0, 0.375, 0.44, 0.183], t,
                   {"Sales AR": num, "Booking": num, "Share %": pct, "Avg Profit %": pct,
                    "Avg RN/Bln": num}, cw)
        import textwrap
        bg.text(x0, 0.365, "\n".join(textwrap.wrap(cap, 95)), fontsize=5.6, color=MUTE, va="top")

    # Donut komposisi Room Night
    for i, (title, comp, col_) in enumerate(dash.get("pies", [])[:2]):
        x0 = L if i == 0 else 0.515
        sec(0.318, title, x=x0)
        ax = fig.add_axes([x0 - 0.01, 0.135, 0.22, 0.17])
        vals = comp["Room Night"].clip(lower=0)
        wedges, _t, autot = ax.pie(vals, colors=PALETTE[: len(vals)], startangle=90,
                                   counterclock=False, autopct=lambda p: f"{p:.0f}%" if p >= 5 else "",
                                   pctdistance=0.79, wedgeprops=dict(width=0.38, edgecolor="white",
                                                                     linewidth=1.2))
        for a_ in autot:
            a_.set_color("white"); a_.set_fontsize(5.6); a_.set_weight("bold")
        ax.text(0, 0.06, angka(vals.sum()), ha="center", fontsize=9, weight="bold", color=INK)
        ax.text(0, -0.16, "Room Night", ha="center", fontsize=5.8, color=MUTE)
        ax.set_aspect("equal")
        tot = vals.sum() or 1
        for j, (name, v) in enumerate(zip(comp[col_], vals)):
            yy = 0.27 - j * 0.017
            if yy < 0.14:
                break
            bg.add_patch(Rectangle((x0 + 0.225, yy - 0.002), 0.008, 0.008, fc=PALETTE[j % len(PALETTE)]))
            bg.text(x0 + 0.238, yy, _cut(name, 17), fontsize=6.3, color=INK, va="bottom")
            bg.text(x0 + 0.44, yy, f"{angka(v)}  ·  " + pct(v / tot * 100), fontsize=6.3,
                    color=MUTE, va="bottom", ha="right")

    # Domestic vs International
    dvi = dash.get("dvi")
    if dvi is not None and len(dvi):
        sec(0.105, "Domestic vs International", "berdasarkan kolom Product Type")
        dvi = dvi.copy()
        draw_table([L, 0.035, R - L, 0.064], dvi,
                   {"Revenue": lambda v: rupiah_full(v), "Room Night": num, "Booking": num,
                    "Share Revenue %": pct},
                   [0.22, 0.28, 0.16, 0.16, 0.18], font=6.5)

    bg.text(L, 0.014, "Data mengikuti filter yang aktif di aplikasi saat diunduh.", fontsize=5.8,
            color=MUTE)
    buf = io.BytesIO()
    fig.savefig(buf, format="pdf")
    plt.close(fig)
    return buf.getvalue()


def hotel_pdf(fx: dict) -> bytes:
    """PDF 1 halaman A4 untuk tab Filter per Hotel: filter aktif, ringkasan, Room Night per
    supplier, tren Room Night per supplier, dan tabel Ringkasan per Hotel."""
    import textwrap
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib.patches import FancyBboxPatch, Rectangle

    INK, MUTE, LINE = "#434343", "#7A7A7A", "#e3e6ec"
    plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 7, "axes.edgecolor": LINE,
                         "xtick.color": MUTE, "ytick.color": MUTE})
    fig = plt.figure(figsize=(8.27, 11.69))
    bg = fig.add_axes([0, 0, 1, 1]); bg.axis("off"); bg.set_xlim(0, 1); bg.set_ylim(0, 1)
    L, R = 0.045, 0.955

    def sec(y, title, sub=""):
        bg.text(L, y, title, fontsize=9.5, weight="bold", color=INK, va="bottom")
        if sub:
            bg.text(R, y + 0.001, sub, fontsize=6.3, color=MUTE, va="bottom", ha="right")

    # Header + filter aktif
    bg.text(L, 0.968, "Ringkasan Penjualan per Hotel", fontsize=16, weight="bold", color=INK)
    bg.text(R, 0.968, pd.Timestamp.now().strftime("Dicetak %d %b %Y %H:%M"), fontsize=7,
            color=MUTE, ha="right")
    filt = (f"Periode: {fx.get('periode', '-')}   ·   Pasar: {fx.get('pasar', 'Semua')}   ·   "
            f"Supplier: {fx.get('supplier') or 'Semua'}")
    bg.text(L, 0.952, filt, fontsize=7.3, color=MUTE)
    hotels = fx.get("hotels") or []
    htxt = "Hotel: " + ("Semua hotel" if not hotels else "; ".join(hotels))
    bg.text(L, 0.940, textwrap.shorten(htxt, 175, placeholder=" …"), fontsize=7, color=MUTE)
    bg.add_line(plt.Line2D([L, R], [0.932, 0.932], color=LINE, lw=1))

    # Kotak ringkasan (4)
    kp = fx.get("kpi", [])
    gap = 0.012
    w = (R - L - gap * (len(kp) - 1)) / max(len(kp), 1)
    for i, (label, value, color) in enumerate(kp):
        x, y, h = L + i * (w + gap), 0.872, 0.05
        bg.add_patch(FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0,rounding_size=0.008",
                                    fc="white", ec=LINE, lw=0.8))
        bg.add_patch(Rectangle((x, y + 0.004), 0.004, h - 0.008, fc=color, ec="none"))
        bg.text(x + 0.014, y + h - 0.014, label.upper(), fontsize=6.2, weight="bold", color=color)
        bg.text(x + 0.014, y + 0.011, value, fontsize=12, weight="bold", color=INK)

    # Room Night per supplier (2 x 5)
    sec(0.845, "Room Night per Supplier", "Other = supplier di luar daftar")
    sup = fx.get("sup")
    if sup is not None and len(sup):
        tot = sup["rn"].sum() or 1
        tw, th, tg = (R - L - 4 * 0.01) / 5, 0.05, 0.01
        for i, row in enumerate(sup.itertuples(index=False)):
            r, c = divmod(i, 5)
            x, y = L + c * (tw + tg), 0.785 - r * (th + 0.008)
            sel = fx.get("supplier") == row.name
            bg.add_patch(FancyBboxPatch((x, y), tw, th, boxstyle="round,pad=0,rounding_size=0.006",
                                        fc=row.color + "14" if sel else "white",
                                        ec=row.color if sel else LINE, lw=1.4 if sel else 0.7))
            a = 1 if (row.rn and (not fx.get("supplier") or sel)) else 0.45
            bg.add_patch(plt.Circle((x + 0.011, y + th - 0.012), 0.0035, color=row.color, alpha=a))
            bg.text(x + 0.019, y + th - 0.012, _cut(row.name, 22), fontsize=6.2, weight="bold",
                    color=INK, va="center", alpha=a)
            bg.text(x + 0.009, y + 0.017, angka(row.rn), fontsize=10.5, weight="bold",
                    color=INK, alpha=a)
            bg.text(x + 0.009, y + 0.006, f"{row.rn / tot * 100:.1f}% · {angka(row.bk)} booking",
                    fontsize=5.4, color=MUTE, alpha=a)

    # Tren Room Night per supplier
    sec(0.705, "Tren Room Night per Supplier")
    tr = fx.get("trend")
    if tr is not None and len(tr):
        ax = fig.add_axes([0.08, 0.47, 0.87, 0.215])
        for s in [c for c in tr.columns if c != "Label"]:
            ax.plot(tr["Label"], tr[s], lw=1.8, marker="o", ms=2.8, label=s,
                    color=SUPPLIER_COLORS.get(s, "#7A7A7A"), ls=":" if s == "Other" else "-")
        ax.set_ylim(bottom=0)
        ax.grid(axis="y", color=LINE, lw=0.6); ax.set_axisbelow(True)
        for s_ in ("top", "right", "left"):
            ax.spines[s_].set_visible(False)
        ax.tick_params(length=0, labelsize=6.5)
        ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda v, _p: angka(v)))
        ax.legend(loc="lower left", bbox_to_anchor=(0, 1.0), ncol=6, frameon=False, fontsize=6.2)

    # Tabel ringkasan per hotel
    t = fx.get("table")
    if t is not None and len(t):
        n_max = 18
        sec(0.425, "Ringkasan per Hotel",
            f"{angka(len(t))} hotel · urut Revenue terbesar"
            + (f" · ditampilkan {n_max} teratas (lengkap di file Excel)" if len(t) > n_max else ""))
        tt = t.head(n_max).copy()
        keep = [c for c in ["Hotel", "Kota", "Revenue", "Room Night", "Booking",
                            "Rata-rata / Room Night", "Share RN %"] if c in tt.columns]
        tt = tt[keep].rename(columns={"Rata-rata / Room Night": "Avg / RN", "Share RN %": "Share RN"})
        tt["Hotel"] = tt["Hotel"].map(lambda s: _cut(s, 38))
        fm = {"Revenue": angka, "Room Night": angka, "Booking": angka, "Avg / RN": angka,
              "Share RN": lambda v: f"{v:.1f}%".replace(".", ",")}
        cells = [[("-" if pd.isna(v) else fm.get(c, str)(v)) for c, v in zip(tt.columns, row)]
                 for row in tt.itertuples(index=False)]
        cw = {"Hotel": 0.33, "Kota": 0.12, "Revenue": 0.15, "Room Night": 0.1, "Booking": 0.09,
              "Avg / RN": 0.11, "Share RN": 0.1}
        rows = len(cells) + 1
        ax = fig.add_axes([L, 0.415 - 0.0205 * rows, R - L, 0.0205 * rows]); ax.axis("off")
        tb = ax.table(cellText=cells, colLabels=list(tt.columns), loc="upper center",
                      colWidths=[cw.get(c, 0.1) for c in tt.columns], cellLoc="right", colLoc="right")
        tb.auto_set_font_size(False); tb.set_fontsize(6.4)
        for (r_, c_), cell in tb.get_celld().items():
            cell.set_edgecolor(LINE); cell.set_linewidth(0.5); cell.set_height(1 / rows)
            if tt.columns[c_] in ("Hotel", "Kota"):
                cell._loc = "left"; cell.set_text_props(ha="left")
            if r_ == 0:
                cell.set_facecolor("#f3f5f8"); cell.set_text_props(weight="bold", color=MUTE)
            else:
                cell.set_facecolor("white" if r_ % 2 else "#fafbfc")

    bg.text(L, 0.014, "Data mengikuti filter yang aktif di aplikasi saat diunduh.", fontsize=5.8,
            color=MUTE)
    buf = io.BytesIO()
    fig.savefig(buf, format="pdf")
    plt.close(fig)
    return buf.getvalue()


def file_download(box, label: str, make, file_name: str, mime: str, key: str, **kw):
    """Tombol unduh umum; file dibuat saat tombol diklik bila Streamlit mendukung."""
    extra = {"on_click": "ignore"} if LAZY_DL else {}
    box.download_button(label, make if LAZY_DL else make(), file_name=file_name, mime=mime,
                        key=key, type="primary", **extra, **kw)


# ---------------------------------------------------------------------------
# Ringkasan otomatis oleh Claude (Anthropic API)
#   API key dibaca dari Streamlit Secrets / environment: ANTHROPIC_API_KEY
#   Model opsional: ANTHROPIC_MODEL (default di bawah). Hanya angka ringkas halaman
#   (bukan data mentah per transaksi) yang dikirim ke Claude.
# ---------------------------------------------------------------------------
AI_MODEL_DEFAULT = "claude-sonnet-5-5"
def _secret_str(v, name: str = ""):
    """Ambil nilai teks dari Secrets. Mendukung penulisan biasa maupun di dalam [section],
    mis.  ANTHROPIC_API_KEY = "sk-..."   atau   [ANTHROPIC_API_KEY] / [anthropic] api_key = "sk-..." """
    if v is None:
        return None
    if isinstance(v, str):
        return v.strip().strip('"').strip("'").strip() or None
    if hasattr(v, "keys"):  # sebuah [section] -> cari nilai teks di dalamnya
        for k in (name, name.lower(), "api_key", "key", "value", "model"):
            if k and k in v and isinstance(v[k], str) and v[k].strip():
                return v[k].strip()
        for x in v.values():
            if isinstance(x, str) and x.strip():
                return x.strip()
    return None


def _secret(name: str):
    val = None
    try:
        val = _secret_str(st.secrets.get(name), name)
        if not val:  # mungkin ditulis di dalam section lain, mis. [anthropic]
            for sec in st.secrets.values():
                if hasattr(sec, "keys") and (name in sec or name.lower() in sec):
                    val = _secret_str(sec.get(name, sec.get(name.lower())), name)
                    if val:
                        break
    except Exception:  # noqa: BLE001 - secrets.toml tidak ada / format salah
        val = None
    return val or (os.environ.get(name) or "").strip() or None


def _df_txt(df: pd.DataFrame, max_rows: int = 15) -> str:
    """Tabel ringkas -> teks CSV dengan angka dibulatkan, untuk dikirim ke Claude."""
    d = df.head(max_rows).copy()
    for c in d.columns:
        if pd.api.types.is_float_dtype(d[c]):
            d[c] = d[c].round(2)
    return d.to_csv(index=False)


def dash_facts(dash: dict) -> str:
    L_ = [f"Periode Inv Date: {dash.get('periode', '-')}", "", "Scorecard:"]
    L_ += [f"- {lbl}: {val} ({sub})" for lbl, val, sub, _c in dash.get("kpi", [])]
    m = dash.get("trend")
    if m is not None and len(m):
        L_ += ["", "Tren bulanan (Rp):", _df_txt(m[["Label", "Sales", "Profit"]]
                                             .rename(columns={"Label": "Bulan", "Sales": "Sales AR"}), 24)]
    for k in ("cust_sales", "cust_profit"):
        if k in dash:
            title, sub, t, cap = dash[k]
            L_ += ["", f"{title} ({sub}):", _df_txt(t), cap]
    for title, comp, col_ in dash.get("pies", []):
        tot = comp["Room Night"].sum() or 1
        L_ += ["", f"{title}:"] + [f"- {n}: {angka(v)} RN ({v / tot * 100:.1f}%)"
                                   for n, v in zip(comp[col_], comp["Room Night"])]
    if dash.get("dvi") is not None:
        L_ += ["", "Domestic vs International:", _df_txt(dash["dvi"])]
    return "\n".join(L_)


def hotel_facts(fx: dict) -> str:
    hotels = fx.get("hotels") or []
    L_ = [f"Periode Inv Date: {fx.get('periode', '-')}",
          f"Filter pasar: {fx.get('pasar', 'Semua')} · supplier: {fx.get('supplier') or 'Semua'}",
          "Hotel dipilih: " + ("semua hotel" if not hotels else "; ".join(hotels[:15])), "",
          "Ringkasan:"] + [f"- {lbl}: {val}" for lbl, val, _c in fx.get("kpi", [])]
    sup = fx.get("sup")
    if sup is not None and len(sup):
        L_ += ["", "Room Night per supplier (semua supplier, sebelum filter supplier):",
               _df_txt(sup[["name", "rn", "bk"]].rename(
                   columns={"name": "Supplier", "rn": "Room Night", "bk": "Booking"}), 12)]
    tr = fx.get("trend")
    if tr is not None and len(tr):
        L_ += ["", "Tren Room Night bulanan per supplier:", _df_txt(tr, 24)]
    t = fx.get("table")
    if t is not None and len(t):
        L_ += ["", f"Ringkasan per hotel (total {len(t)} hotel, 15 teratas menurut Revenue):",
               _df_txt(t, 15)]
    return "\n".join(L_)


# ---------------------------------------------------------------------------
# "Ask Claude": tanya jawab bebas tentang data. Claude TIDAK menerima data mentah;
# Claude memanggil alat `query_data` (filter + group by + agregasi) yang dijalankan oleh
# aplikasi ini dengan pandas, lalu menjawab berdasarkan hasilnya.
# ---------------------------------------------------------------------------
CHAT_MAX_ROWS = 60
QUERY_TOOL = {
    "name": "query_data",
    "description": (
        "Menjalankan agregasi pada tabel transaksi penjualan hotel (sudah mengikuti filter "
        "aplikasi). Pakai untuk SETIAP angka yang kamu sebutkan. Alur: filters -> group_by -> "
        "metrics -> sort -> limit. Tanpa group_by = satu baris total. "
        "Booking = nunique dari 'Invoice No'. Revenue = sum 'Sales AR'. Room Night = sum "
        "'Room Night'. Margin % = sum Profit / sum Sales AR * 100 (hitung sendiri dari hasil). "
        "Untuk tren bulanan group_by 'Bulan' (tanggal awal bulan). Perbandingan teks tidak "
        "peka huruf besar/kecil; 'contains' cocok untuk sebagian nama."),
    "input_schema": {
        "type": "object",
        "properties": {
            "filters": {"type": "array", "description": "Syarat baris (AND).", "items": {
                "type": "object",
                "properties": {
                    "column": {"type": "string"},
                    "op": {"type": "string", "enum": ["eq", "ne", "in", "not_in", "contains",
                                                      "not_contains", "gt", "gte", "lt", "lte",
                                                      "between"]},
                    "value": {"description": "Nilai; list untuk in/not_in/between. Tanggal: "
                                             "'YYYY-MM-DD'."}},
                "required": ["column", "op", "value"]}},
            "group_by": {"type": "array", "items": {"type": "string"}},
            "metrics": {"type": "array", "items": {
                "type": "object",
                "properties": {
                    "column": {"type": "string"},
                    "agg": {"type": "string", "enum": ["sum", "mean", "median", "min", "max",
                                                       "count", "nunique"]},
                    "as": {"type": "string", "description": "Nama kolom hasil (opsional)."}},
                "required": ["column", "agg"]}},
            "sort_by": {"type": "string", "description": "Nama kolom hasil untuk mengurutkan."},
            "ascending": {"type": "boolean"},
            "limit": {"type": "integer", "description": f"Maks baris (default 20, maks {CHAT_MAX_ROWS})."},
        },
        "required": ["metrics"],
    },
}


def run_query(df: pd.DataFrame, q: dict) -> pd.DataFrame:
    """Eksekusi aman (tanpa eval) dari permintaan query_data."""
    d = df
    for f in q.get("filters") or []:
        col, op, val = f.get("column"), f.get("op"), f.get("value")
        if col not in d.columns:
            raise ValueError(f"Kolom '{col}' tidak ada.")
        s = d[col]
        is_dt = pd.api.types.is_datetime64_any_dtype(s)
        is_num = pd.api.types.is_numeric_dtype(s) and not pd.api.types.is_bool_dtype(s)
        vals = val if isinstance(val, list) else [val]

        def conv(x):
            if is_dt:
                return pd.to_datetime(x)
            if is_num:
                return float(x)
            return str(x).strip().lower()
        cv = [conv(x) for x in vals]
        ss = s if (is_dt or is_num) else s.astype(str).str.strip().str.lower()
        if op == "eq":
            m = ss == cv[0]
        elif op == "ne":
            m = ss != cv[0]
        elif op == "in":
            m = ss.isin(cv)
        elif op == "not_in":
            m = ~ss.isin(cv)
        elif op in ("contains", "not_contains"):
            pat = "|".join(re.escape(str(x)) for x in cv)
            m = s.astype(str).str.lower().str.contains(pat, regex=True, na=False)
            m = ~m if op == "not_contains" else m
        elif op == "gt":
            m = ss > cv[0]
        elif op == "gte":
            m = ss >= cv[0]
        elif op == "lt":
            m = ss < cv[0]
        elif op == "lte":
            m = ss <= cv[0]
        elif op == "between":
            if len(cv) != 2:
                raise ValueError("between butuh 2 nilai [awal, akhir].")
            hi = cv[1] + pd.Timedelta(days=1) - pd.Timedelta(seconds=1) if is_dt else cv[1]
            m = ss.between(cv[0], hi)
        else:
            raise ValueError(f"Operator '{op}' tidak dikenal.")
        d = d[m.fillna(False)]
    gb = [g for g in (q.get("group_by") or [])]
    for g in gb:
        if g not in d.columns:
            raise ValueError(f"Kolom group_by '{g}' tidak ada.")
    aggs = {}
    for mt in q.get("metrics") or []:
        col, agg = mt.get("column"), mt.get("agg")
        if col not in d.columns:
            raise ValueError(f"Kolom metrik '{col}' tidak ada.")
        if agg not in ("sum", "mean", "median", "min", "max", "count", "nunique"):
            raise ValueError(f"Agregasi '{agg}' tidak dikenal.")
        name = mt.get("as") or f"{agg}({col})"
        src_col = col
        if col == "Invoice No":
            d = d.assign(**{"Invoice No": invoice_key(d["Invoice No"])})
        aggs[name] = (src_col, "size" if agg == "count" else agg)
    if not aggs:
        raise ValueError("Minimal satu metrik.")
    if gb:
        res = d.groupby(gb, dropna=False).agg(**aggs).reset_index()
    else:
        res = pd.DataFrame({k: [d[c].agg(a) if a != "size" else len(d)] for k, (c, a) in aggs.items()})
    sb = q.get("sort_by")
    if sb and sb in res.columns:
        res = res.sort_values(sb, ascending=bool(q.get("ascending", False)))
    lim = max(1, min(int(q.get("limit") or 20), CHAT_MAX_ROWS))
    res.attrs["total_rows"] = len(res)
    res.attrs["source_rows"] = len(d)
    return res.head(lim).reset_index(drop=True)


def data_schema(df: pd.DataFrame) -> str:
    """Ringkasan kolom untuk Claude: tipe data, contoh nilai, rentang tanggal."""
    skip = {"No", "Source File"}
    L_ = [f"Jumlah baris: {len(df)}"]
    if "Inv Date" in df.columns and df["Inv Date"].notna().any():
        L_.append(f"Rentang Inv Date: {df['Inv Date'].min():%Y-%m-%d} s/d {df['Inv Date'].max():%Y-%m-%d}")
    L_.append("Kolom:")
    for c in df.columns:
        if c in skip or str(c).startswith("_"):
            continue
        s = df[c]
        if pd.api.types.is_datetime64_any_dtype(s):
            L_.append(f"- {c} [tanggal]")
        elif pd.api.types.is_numeric_dtype(s):
            L_.append(f"- {c} [angka] total={s.sum():,.0f}")
        else:
            vc = s.astype(str).value_counts()
            top = ", ".join(f"{k}" for k in vc.index[:12])
            L_.append(f"- {c} [teks, {len(vc)} nilai unik] contoh: {top[:400]}")
    return "\n".join(L_)


CHAT_SYSTEM = """Kamu "Claude", asisten analis data penjualan hotel sebuah travel agent korporat di Indonesia.
Jawab pertanyaan pengguna HANYA berdasarkan hasil alat query_data (jangan menebak angka).
Panggil query_data sebanyak yang diperlukan (boleh beberapa kali) sebelum menjawab.

Kolom penting:
- 'Sales AR' = Revenue (Rp) · 'Profit' (Rp) · 'Room Night' (= Room × Night)
- Booking = nunique 'Invoice No'
- 'Customer Name' dan 'Invoice To' = nama customer · 'Supplier Name' · 'Supplier Category'
  (Wholesaler / OTA / PERTAMINA Corp Rate / Direct To Hotel)
- 'Hotel (Grup)' = nama hotel yang sudah distandarkan (pakai ini untuk analisis hotel)
- 'Hotel City' · 'Product Type' (Domestic/International) · 'Inv Date' · 'Bulan' (awal bulan)

Gaya jawaban: Bahasa Indonesia, ringkas dan langsung ke inti. Sebut angka penting dengan format
Indonesia (Rp 1,25 M; Rp 350 Jt; 1.234 RN). Pakai tabel markdown kecil bila membandingkan beberapa
item. Jika pertanyaan tidak bisa dijawab dari data, katakan terus terang dan sarankan pertanyaan
lain. Data yang tersedia sudah mengikuti filter aplikasi (lihat ringkasan di bawah).

Konteks: pertanyaan diajukan dari halaman "{page}". Angka yang sedang tampil di halaman itu
(boleh dikutip langsung; untuk angka lain pakai query_data):
{facts}

Ringkasan kolom data yang bisa di-query (sudah mengikuti semua filter halaman ini):
{schema}"""


def ask_claude(question: str, history: list, df: pd.DataFrame, model: str,
               page: str = "Dashboard", facts: str = "-"):
    """Loop tool-use: Claude memanggil query_data sampai siap menjawab.
    Return (jawaban, daftar (query, hasil DataFrame atau pesan error))."""
    import anthropic
    client = anthropic.Anthropic(api_key=_secret("ANTHROPIC_API_KEY"), timeout=90, max_retries=2)
    system = CHAT_SYSTEM.format(schema=data_schema(df), page=page, facts=facts[:6000])
    # Riwayat: hanya pesan teks yang tidak kosong, dan harus diawali pesan user
    msgs = [{"role": m["role"], "content": str(m["content"])} for m in history[-8:]
            if str(m.get("content", "")).strip()
            and not str(m["content"]).startswith(("⚠️", "(Claude tidak"))]
    while msgs and msgs[0]["role"] != "user":
        msgs.pop(0)
    msgs.append({"role": "user", "content": question})
    used, earlier_text, last_stop = [], [], None
    NUDGE = ("Berdasarkan hasil query di atas, tuliskan sekarang jawaban akhir untuk pertanyaan "
             "saya dalam Bahasa Indonesia (jangan memanggil alat lagi).")

    def add_user_text(text):
        """Tambahkan teks ke pesan user terakhir (atau buat pesan user baru)."""
        if msgs and msgs[-1]["role"] == "user":
            c = msgs[-1]["content"]
            msgs[-1]["content"] = ([{"type": "text", "text": c}] if isinstance(c, str) else c) + [
                {"type": "text", "text": text}]
        else:
            msgs.append({"role": "user", "content": text})

    for step in range(10):
        final_try = step >= 8  # langkah terakhir: paksa menjawab tanpa alat
        kw = {"tool_choice": {"type": "none"}} if final_try else {}
        r = client.messages.create(model=model, max_tokens=4096, system=system,
                                   tools=[QUERY_TOOL], messages=msgs, **kw)
        last_stop = r.stop_reason
        blocks = []
        for b in r.content:
            if b.type == "text" and b.text.strip():
                blocks.append({"type": "text", "text": b.text})
            elif b.type == "tool_use":
                blocks.append({"type": "tool_use", "id": b.id, "name": b.name, "input": b.input})
        if r.stop_reason == "pause_turn" and blocks:  # server minta dilanjutkan
            msgs.append({"role": "assistant", "content": blocks})
            continue
        calls = [b for b in blocks if b["type"] == "tool_use"]
        text = "\n\n".join(b["text"] for b in blocks if b["type"] == "text").strip()
        if calls and r.stop_reason == "tool_use" and not final_try:
            if text:
                earlier_text.append(text)
            msgs.append({"role": "assistant", "content": blocks})
            results = []
            for c in calls:
                try:
                    res = run_query(df, c["input"])
                    used.append((c["input"], res))
                    body = (f"{res.attrs.get('total_rows', len(res))} baris hasil "
                            f"(ditampilkan {len(res)}), dari {res.attrs.get('source_rows')} baris data:\n"
                            + res.round(2).to_csv(index=False))
                    results.append({"type": "tool_result", "tool_use_id": c["id"], "content": body})
                except Exception as e:  # noqa: BLE001
                    used.append((c["input"], f"{type(e).__name__}: {e}"))
                    results.append({"type": "tool_result", "tool_use_id": c["id"],
                                    "content": f"Error: {e}", "is_error": True})
            msgs.append({"role": "user", "content": results})
            continue
        if text:
            if r.stop_reason == "max_tokens":
                text += "\n\n_(jawaban terpotong karena terlalu panjang — coba persempit pertanyaan)_"
            return text, used
        # Jawaban kosong: minta Claude menuliskan jawaban akhir (tanpa memanggil alat lagi)
        if blocks:
            msgs.append({"role": "assistant", "content": blocks})
        add_user_text(NUDGE)
        if step < 8:
            r2 = client.messages.create(model=model, max_tokens=4096, system=system,
                                        tools=[QUERY_TOOL], tool_choice={"type": "none"},
                                        messages=msgs)
            last_stop = r2.stop_reason
            t2 = "\n\n".join(b.text for b in r2.content
                               if b.type == "text" and b.text.strip()).strip()
            if t2:
                return t2, used
        break
    if earlier_text:
        return "\n\n".join(earlier_text), used
    return (f"Maaf, Claude belum memberikan jawaban (status: {last_stop}). Coba ulangi atau "
            "persempit pertanyaannya."), used


def ask_claude_box(key: str, df: pd.DataFrame, page: str, facts: str, suggestions: list):
    """Kotak 'Ask Claude' (teks bebas) di bagian bawah sebuah tab. Data yang bisa ditanyakan =
    `df` (sudah mengikuti semua filter tab tersebut). Riwayat chat terpisah per tab."""
    with st.container(border=True, key=f"card_ask_{key}"):
        c1_, c2_ = st.columns([5, 1], vertical_alignment="center")
        section("💬 Ask Claude", f"tanya apa saja tentang data di halaman ini · "
                f"{angka(len(df))} baris sesuai filter", box=c1_)
        hk = f"chat_{key}"
        if c2_.button("🗑️ Hapus chat", key=f"chat_clear_{key}", width="stretch"):
            st.session_state[hk] = []
        if not _secret("ANTHROPIC_API_KEY"):
            st.info("Ask Claude belum aktif. Tambahkan `ANTHROPIC_API_KEY` di **Streamlit Secrets** "
                    "(⋮ → Settings → Secrets), lalu jalankan ulang aplikasi.")
            return
        chat = st.session_state.setdefault(hk, [])
        sug_box, msg_box = st.container(), st.container()
        q_in = st.chat_input("Tulis pertanyaan… (mis. supplier mana yang paling menguntungkan "
                             "bulan lalu?)", key=f"chat_q_{key}")

        pending = None
        if not chat and not q_in:
            with sug_box:
                st.caption("Contoh pertanyaan (klik untuk bertanya):")
                sc = st.columns(2)
                for i_, s_ in enumerate(suggestions):
                    if sc[i_ % 2].button(s_, key=f"chat_sug_{key}_{i_}", width="stretch"):
                        pending = s_
        question = q_in or pending
        with msg_box:
            for m_ in chat:
                with st.chat_message(m_["role"], avatar="🧑‍💼" if m_["role"] == "user" else "✨"):
                    st.markdown(m_["content"])
            if question:
                with st.chat_message("user", avatar="🧑‍💼"):
                    st.markdown(question)
                with st.chat_message("assistant", avatar="✨"):
                    with st.spinner("Claude sedang menghitung dari data…"):
                        try:
                            hist = [{"role": m_["role"], "content": m_["content"]} for m_ in chat]
                            ans, used = ask_claude(question, hist, df,
                                                   _secret("ANTHROPIC_MODEL") or AI_MODEL_DEFAULT,
                                                   page=page, facts=facts)
                        except Exception as e:  # noqa: BLE001
                            ans, used = f"⚠️ Gagal menghubungi Claude: {type(e).__name__}: {e}", []
                    st.markdown(ans)
                chat.append({"role": "user", "content": question})
                chat.append({"role": "assistant", "content": ans, "used": used})
                if pending:  # sembunyikan contoh pertanyaan setelah dipakai
                    st.rerun()


ASK_SUGGEST_DASH = [
    "Siapa 5 customer dengan profit terbesar, dan berapa margin masing-masing?",
    "Bulan apa revenue tertinggi dan terendah? Berapa selisihnya?",
    "Bandingkan Room Night per kategori supplier per bulan.",
    "Kota mana dengan revenue terbesar dan hotel apa yang paling berkontribusi?",
]
ASK_SUGGEST_HOTEL = [
    "Supplier mana yang paling banyak menyumbang Room Night di sini?",
    "Bulan apa Room Night tertinggi dan terendah?",
    "Customer siapa yang paling sering menginap di hotel-hotel ini?",
    "Hotel mana dengan rata-rata revenue per room night tertinggi?",
]


def col_sum(df, c):
    return df[c].sum() if c in df.columns else 0


with tab_dash:
    dash = {}  # isi dashboard, dipakai juga untuk ekspor PDF
    pdf_box = st.container()  # tombol ekspor PDF (diisi setelah semua bagian dihitung)
    sales = col_sum(fdata, "Sales AR")
    profit = col_sum(fdata, "Profit")
    n_inv = invoice_key(fdata["Invoice No"]).nunique() if "Invoice No" in fdata else len(fdata)
    rn = col_sum(fdata, "Room Night")
    margin = profit / sales * 100 if sales else 0
    n_hotel = fdata["Hotel (Grup)"].nunique() if "Hotel (Grup)" in fdata.columns else 0
    n_city = (fdata["Hotel City"].dropna().astype(str).str.strip().str.title().nunique()
              if "Hotel City" in fdata.columns else 0)
    periode = "-"
    if "Inv Date" in fdata.columns and fdata["Inv Date"].notna().any():
        periode = f"{fdata['Inv Date'].min():%d %b %Y} – {fdata['Inv Date'].max():%d %b %Y}"
    kpi_cards([
        dict(icon="💰", label="Total Sales AR", value=rupiah_full(sales), tone="blue",
             sub=f"Rata-rata {rupiah_full(sales / n_inv) if n_inv else '-'} / invoice"),
        dict(icon="📈", label="Total Profit", value=rupiah_full(profit), tone="green",
             sub=f"Rata-rata {rupiah_full(profit / n_inv) if n_inv else '-'} / invoice"),
        dict(icon="🎯", label="Margin", value=f"{margin:.2f}%", tone="orange",
             sub="Total Profit ÷ Total Sales AR", bar=max(min(margin / 20 * 100, 100), 0),
             tip="Bar penuh = margin 20%"),
        dict(icon="🧾", label="Jumlah Invoice (unik)", value=angka(n_inv), tone="purple",
             sub=f"dari {angka(len(fdata))} baris transaksi",
             tip="Satu invoice bisa berisi beberapa baris (beberapa kamar, penyesuaian, dll.)"),
        dict(icon="🛏️", label="Room Night", value=angka(rn), tone="teal",
             sub=f"Rata-rata {rupiah_full(sales / rn) if rn else '-'} / room night"),
        dict(icon="🏨", label="Jumlah Hotel", value=angka(n_hotel), tone="rose",
             sub=f"di {angka(n_city)} kota"),
    ], caption=f"Periode Inv Date: {periode}")
    dash["periode"] = periode
    dash["info"] = f"{angka(len(fdata))} baris transaksi"
    dash["kpi"] = [
        ("Total Sales AR", rupiah_full(sales),
         f"Rata-rata {rupiah_full(sales / n_inv) if n_inv else '-'} / invoice", "#0D7FCC"),
        ("Total Profit", rupiah_full(profit),
         f"Rata-rata {rupiah_full(profit / n_inv) if n_inv else '-'} / invoice", "#1BA0E2"),
        ("Margin", f"{margin:.2f}%", "Total Profit ÷ Total Sales AR", "#1494C6"),
        ("Jumlah Invoice (unik)", angka(n_inv), f"dari {angka(len(fdata))} baris transaksi", "#434343"),
        ("Room Night", angka(rn), f"Rata-rata {rupiah_full(sales / rn) if rn else '-'} / room night",
         "#0D7FCC"),
        ("Jumlah Hotel", angka(n_hotel), f"di {angka(n_city)} kota", "#1BA0E2"),
    ]

    # Tren bulanan
    if "Bulan" in fdata.columns:
        m = (fdata.groupby("Bulan")
             .agg(Sales=("Sales AR", "sum"), Profit=("Profit", "sum"),
                  RoomNight=("Room Night", "sum"))
             .reset_index())
        m["Label"] = m["Bulan"].dt.month.map(lambda x: BULAN[x - 1]) + " " + m["Bulan"].dt.strftime("%y")
        dash["trend"] = m
        with st.container(border=True, key="card_trend"):
            section("Tren Bulanan", "Sales AR (batang) & Profit (garis)")
            fig = go.Figure()
            fig.add_bar(x=m["Label"], y=m["Sales"], name="Sales AR", marker_color=PALETTE[0],
                        opacity=.9, hovertemplate="Sales AR: Rp %{y:,.0f}<extra></extra>")
            try:
                fig.update_traces(marker_cornerradius=6, selector=dict(type="bar"))
            except (ValueError, TypeError):  # plotly lama belum mendukung sudut membulat
                pass
            fig.add_scatter(x=m["Label"], y=m["Profit"], name="Profit", yaxis="y2",
                            mode="lines+markers", line=dict(color=PALETTE[2], width=3),
                            marker=dict(size=7), hovertemplate="Profit: Rp %{y:,.0f}<extra></extra>")
            clean_fig(fig, 340)
            fig.update_layout(hovermode="x unified", bargap=.35,
                              yaxis=dict(title=None, tickprefix="Rp "),
                              yaxis2=dict(title=None, overlaying="y", side="right", showgrid=False,
                                          tickprefix="Rp "))
            st.plotly_chart(fig, width="stretch", config={"displayModeBar": False})

    def top_bar(col, title, color, metric="Sales AR"):
        if col not in fdata.columns or metric not in fdata.columns:
            return None
        d = fdata[fdata[col].astype(str).str.strip().ne("-")]
        key = d[col].astype(str).str.strip()
        if col == "Hotel City":  # satukan 'JAKARTA' / 'Jakarta'
            key = key.str.title()
        aggs = {"Sales AR": ("Sales AR", "sum"), "Room Night": ("Room Night", "sum")}
        if "Profit" in d.columns:
            aggs["Profit"] = ("Profit", "sum")
        agg = d.groupby(key).agg(**aggs)
        agg.index.name = col
        t = agg.nlargest(top_n, metric).sort_values(metric).reset_index()
        t["Sales / Room Night"] = t["Sales AR"] / t["Room Night"].where(t["Room Night"] > 0)
        f = px.bar(t, x=metric, y=col, orientation="h", title=title,
                   color_discrete_sequence=[color], height=max(350, 28 * len(t) + 100),
                   text_auto=",.0f" if metric == "Room Night" else False,
                   hover_data={c: ":,.0f" for c in
                               ["Sales AR", "Profit", "Room Night", "Sales / Room Night"]
                               if c in t.columns})
        f.update_traces(textposition="outside", cliponaxis=False)
        f.update_layout(yaxis_title=None, margin=dict(t=50, b=20, r=40))
        return f

    def four_charts(metric, label, colors, key, second=None):
        """second: (kolom, nama, metric) untuk grafik ke-2; default grafik hotel dengan metric sama."""
        second = second or ("Hotel (Grup)", "Hotel", metric)
        c1, c2 = st.columns(2)
        for col_box, (col, name, m_), color in zip(
            [c1, c2, c1, c2],
            [("Customer Name", "Customer", metric), second,
             ("Hotel City", "Kota", metric), ("Supplier Name", "Supplier", metric)], colors):
            fig = top_bar(col, f"Top {top_n} {name} ({m_ if m_ != metric else label})", color, m_)
            if fig is not None:
                col_box.plotly_chart(fig, width="stretch", key=f"{key}_{col}_{m_}")

    def customer_share_table(box, metric: str, key: str, with_avg: bool = False,
                             name_col: str = "Customer Name"):
        """Tabel Top N Customer berdasarkan metric.
        with_avg=False: kolom Share % = porsi dari total semua customer.
        with_avg=True : kolom Booking (Invoice No unik), Avg Room Night / Bulan (total Room Night
                        ÷ jumlah bulan periode) + Avg Profit % = rata-rata margin (Profit ÷ Sales AR)
                        dari setiap booking customer."""
        if name_col not in fdata.columns or metric not in fdata.columns:
            return
        d = fdata[fdata[name_col].notna() & fdata[name_col].astype(str).str.strip().ne("-")]
        names_ = d[name_col].astype(str).str.strip()
        if name_col != "Customer Name":  # tampilan Proper seperti Customer Name
            names_ = proper_case(names_)
        d = d.assign(_cust=names_,
                     _inv=(invoice_key(d["Invoice No"]) if "Invoice No" in d.columns
                           else pd.Series(d.index.astype(str), index=d.index)))
        agg = d.groupby("_cust")[metric].sum().rename_axis("Customer").to_frame()
        total = agg[metric].sum()
        avg_col, pct_col = "Avg Room Night / Bulan", "Avg Profit %"
        # Jumlah bulan dalam periode data (mengikuti filter tanggal), dihitung inklusif
        n_months = 1
        if "Inv Date" in d.columns and d["Inv Date"].notna().any():
            p0, p1 = d["Inv Date"].min().to_period("M"), d["Inv Date"].max().to_period("M")
            n_months = max((p1 - p0).n + 1, 1)
        if with_avg and "Sales AR" in d.columns:
            # 1 baris per booking: jumlahkan Profit & Sales AR per Invoice No
            bk = (d.dropna(subset=["_inv"]).groupby(["_cust", "_inv"])[["Profit", "Sales AR"]].sum())
            bk["pct"] = bk["Profit"] / bk["Sales AR"].where(bk["Sales AR"] != 0) * 100
            per_c = bk.groupby(level=0).agg(n=("Profit", "size"), pct=("pct", "mean"))
            n_bk = per_c["n"].reindex(agg.index)
            agg["Booking"] = n_bk
            # rata-rata Room Night per bulan = total Room Night customer ÷ jumlah bulan periode
            rn_c = (d.groupby("_cust")["Room Night"].sum() if "Room Night" in d.columns
                    else pd.Series(0, index=agg.index))
            agg[avg_col] = rn_c.reindex(agg.index).fillna(0) / n_months
            agg[pct_col] = per_c["pct"].reindex(agg.index)
            all_pct = bk["pct"].mean()
        else:
            with_avg = False
            agg["Share %"] = agg[metric] / total * 100 if total else 0.0
        t = agg.nlargest(top_n, metric).reset_index()
        t[metric] = round_money(t[metric].astype(float))
        bar_col = pct_col if with_avg else "Share %"
        if with_avg:
            # dibulatkan ke atas (round up) menjadi bilangan bulat, mis. 854,1 -> 855
            import numpy as np
            t[avg_col] = np.ceil(t[avg_col].astype(float).fillna(0) - 1e-9).astype("int64")
            t["Booking"] = t["Booking"].fillna(0).astype(int)
            t = t[["Customer", "Booking", avg_col, pct_col]]
            # Top N dipilih dari total Profit, lalu diurutkan dari Avg Profit % tertinggi
            t = t.sort_values(pct_col, ascending=False, na_position="last").reset_index(drop=True)
        t.insert(0, "Rank", range(1, len(t) + 1))
        vmax = float(t[bar_col].max()) if len(t) and t[bar_col].max() > 0 else 100.0
        sub_txt = "diurutkan dari Avg Profit % tertinggi" if with_avg else "porsi dari total"
        if name_col != "Customer Name":
            sub_txt += f" · berdasarkan kolom {name_col}"
        section(f"Top {top_n} Customer · {metric}", sub_txt, box=box)
        box.dataframe(
            t, hide_index=True, width="stretch", height=38 + 35 * len(t),
            column_config={
                "Rank": st.column_config.NumberColumn("#", width=40),
                "Customer": st.column_config.TextColumn(width="medium"),
                metric: st.column_config.NumberColumn(format="localized", width="small"),
                "Booking": st.column_config.NumberColumn(
                    "Booking", format="localized", width="small",
                    help="Jumlah booking = jumlah Invoice No. unik"),
                avg_col: st.column_config.NumberColumn(
                    "Avg RN/Bln", format="localized", width="small",
                    help=f"Rata-rata Room Night per bulan = total Room Night customer ÷ "
                         f"{n_months} bulan periode data, dibulatkan ke atas"),
                pct_col: st.column_config.ProgressColumn(
                    "Avg Profit %", format="%.2f%%", min_value=0, max_value=vmax, width="small",
                    help="Rata-rata margin per booking: (Profit ÷ Sales AR) tiap Invoice No., "
                         "lalu dirata-rata per customer"),
                "Share %": st.column_config.ProgressColumn(
                    "Share %", format="%.2f%%", min_value=0, max_value=vmax, width="small"),
            })
        if with_avg:
            per_txt = (f"{d['Inv Date'].min():%b %Y} – {d['Inv Date'].max():%b %Y}"
                       if "Inv Date" in d.columns and d["Inv Date"].notna().any() else "-")
            cap = (f"Periode: {per_txt} ({n_months} bulan) · "
                   f"Top {top_n} berdasarkan total {metric} ({rupiah_full(total)}), "
                   f"diurutkan dari Avg Profit % tertinggi · "
                   f"rata-rata profit per booking semua customer: {all_pct:.2f}%.")
        else:
            cap = (f"Top {top_n} = {t['Share %'].sum():.2f}% dari total {metric} "
                   f"({rupiah_full(total)}).")
        box.caption(cap)
        dash[key] = (f"Top {top_n} Customer · {metric}", sub_txt, t.copy(), cap)

    t1, t2 = st.columns(2)
    with t1.container(border=True, key="card_cust_sales"):
        # Customer di tabel Sales AR diambil dari kolom Invoice To (bila ada)
        customer_share_table(st, "Sales AR", "cust_sales", name_col=inv_to_col or "Customer Name")
    with t2.container(border=True, key="card_cust_profit"):
        customer_share_table(st, "Profit", "cust_profit", with_avg=True)

    # Komposisi berdasarkan Room Night (Room × Night); Sales AR tetap tampil saat hover
    c3, c4 = st.columns(2)
    for box_, col_, label_ in [(c3, "Supplier Category", "Kategori Supplier"),
                               (c4, "Product Type", "Product Type")]:
        if col_ in fdata.columns and "Room Night" in fdata.columns:
            comp = (fdata.groupby(col_)
                    .agg(**{"Room Night": ("Room Night", "sum"), "Sales AR": ("Sales AR", "sum")})
                    .reset_index().sort_values("Room Night", ascending=False))
            fig_ = px.pie(comp, names=col_, values="Room Night", hole=0.62,
                          hover_data={"Sales AR": ":,.0f"},
                          color_discrete_sequence=PALETTE)
            fig_.update_traces(sort=False, textinfo="percent", textfont_size=12,
                               marker=dict(line=dict(color="rgba(255,255,255,.9)", width=2)),
                               hovertemplate="<b>%{label}</b><br>Room Night: %{value:,.0f}"
                                             " (%{percent})<br>Sales AR: Rp %{customdata[0]:,.0f}"
                                             "<extra></extra>")
            clean_fig(fig_, 330, legend_bottom=True)
            fig_.add_annotation(text=f"<b>{angka(comp['Room Night'].sum())}</b><br>"
                                     "<span style='font-size:11px'>Room Night</span>",
                                showarrow=False, font=dict(size=18))
            dash.setdefault("pies", []).append(
                (f"Room Night per {label_}", comp[[col_, "Room Night"]].copy(), col_))
            with box_.container(border=True, key=f"card_pie_{col_}"):
                section(f"Room Night per {label_}", "komposisi Room × Night")
                st.plotly_chart(fig_, width="stretch", key=f"pie_rn_{col_}",
                                config={"displayModeBar": False})

    # Domestic vs International (dari kolom Product Type), ditaruh setelah pie chart
    if "Product Type" in fdata.columns:
        inv_all = (invoice_key(fdata["Invoice No"]) if "Invoice No" in fdata.columns
                   else pd.Series(fdata.index.astype(str), index=fdata.index))
        dvi = (fdata.assign(_m=market_type(fdata["Product Type"]).values, _inv=inv_all.values)
               .groupby("_m").agg(Revenue=("Sales AR", "sum"),
                                  **{"Room Night": ("Room Night", "sum"),
                                     "Booking": ("_inv", "nunique")}))
        if len(dvi):
            dvi["Share Revenue %"] = dvi["Revenue"] / dvi["Revenue"].sum() * 100
            dvi["Revenue"] = round_money(dvi["Revenue"].astype(float))
            dash["dvi"] = dvi.reset_index().rename(columns={"_m": "Kategori"})
            with st.container(border=True, key="card_dvi"):
                section("Domestic vs International", "berdasarkan kolom Product Type")
                st.dataframe(dvi.reset_index().rename(columns={"_m": "Kategori"}),
                             hide_index=True, width="stretch",
                             column_config={
                                 "Revenue": st.column_config.NumberColumn(format="localized"),
                                 "Room Night": st.column_config.NumberColumn(format="localized"),
                                 "Booking": st.column_config.NumberColumn(format="localized"),
                                 "Share Revenue %": st.column_config.ProgressColumn(
                                     format="%.1f%%", min_value=0, max_value=100)})

    # Ask Claude di bagian paling bawah tab (setelah Domestic vs International)
    _facts_dash = dash_facts(dash)
    ask_claude_box("dash", fdata, "Dashboard", _facts_dash, ASK_SUGGEST_DASH)

    # Tombol ekspor PDF ditaruh di bagian atas dashboard
    with pdf_box:
        _e1, _e2 = st.columns([5, 1], vertical_alignment="center")
        _e1.caption("Ringkasan halaman ini bisa diunduh sebagai PDF satu halaman (A4).")
        _per = re.sub(r"[^A-Za-z0-9]+", "_", dash.get("periode", ""))[:40].strip("_")
        file_download(_e2, "⬇️ Export PDF", lambda d=dash: dashboard_pdf(d),
                      f"dashboard_penjualan_{_per or 'hotel'}.pdf", "application/pdf",
                      "dl_dash_pdf", width="stretch")

# Warna tetap per supplier (dipakai di grafik perbandingan Room Night antar supplier)
SUPPLIER_COLORS = {
    "Expedia": "#0D7FCC", "Fit Ruums": "#F2A33A", "Destination Of The World": "#2CA58D",
    "MG Holiday": "#8E5BD6", "Channel Manager": "#434343", "Kliknbook": "#E4572E",
    "Tiket.com": "#1BA0E2", "Traveloka": "#7FB800", "Webbeds": "#C2416B", "Other": "#A6A6A6",
}


def supplier_group(df: pd.DataFrame) -> pd.Series:
    """Kelompok supplier: nama di SUPPLIER_RN_CARDS (cocok bila mengandung nama tsb), lainnya 'Other'."""
    key = df["Supplier Name"].astype(str).map(_compact)
    grp = pd.Series("Other", index=df.index, dtype=object)
    for name in SUPPLIER_RN_CARDS:
        grp[(grp == "Other") & key.str.contains(_compact(name), regex=False)] = name
    return grp


SUP_CSS = """<style>
.sup-grid {display:grid; grid-template-columns:repeat(5,minmax(0,1fr)); gap:12px; margin:6px 0 12px;}
@media (max-width:1200px) {.sup-grid {grid-template-columns:repeat(3,minmax(0,1fr));}}
@media (max-width:700px)  {.sup-grid {grid-template-columns:repeat(2,minmax(0,1fr));}}
.sup {background:#fff; border:1px solid #e6eaef; border-radius:14px; padding:12px 14px 12px;
      transition:box-shadow .15s ease, transform .15s ease;}
.sup:hover {box-shadow:0 6px 18px rgba(13,127,204,.10); transform:translateY(-1px);}
.sup.zero {opacity:.5;}
.sup .nm {display:flex; align-items:center; gap:8px; font-size:.78rem; font-weight:600;
          color:#434343; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;}
.sup .dot {width:9px; height:9px; border-radius:50%; flex:none;}
.sup .v {font-size:1.45rem; font-weight:700; color:#434343; margin:6px 0 2px;
         font-variant-numeric:tabular-nums; line-height:1.1;}
.sup .v small {font-size:.72rem; font-weight:500; color:#8c8c8c; margin-left:4px;}
.sup .s {font-size:.74rem; color:#8c8c8c;}
.sup .b {height:4px; border-radius:99px; background:#eef2f6; margin-top:9px; overflow:hidden;}
.sup .b > span {display:block; height:100%; border-radius:99px;}
.chips {display:flex; flex-wrap:wrap; gap:8px; margin:2px 0 10px;}
.chip {background:#F1F8FD; color:#0D7FCC; border-radius:99px; padding:4px 12px;
       font-size:.8rem; font-weight:600;}
.chip.g {background:#f2f2f2; color:#434343; font-weight:500;}
@media (prefers-color-scheme: dark) {
  .sup {background:#1b1f2a; border-color:#2b3140;} .sup .nm, .sup .v {color:#f1f3f7;}
  .sup .b {background:#2b3140;} .chip.g {background:#2b3140; color:#e6e6e6;}
}
</style>"""


def chips(items: list):
    """Baris 'chip' kecil berisi ringkasan (teks, gaya) - gaya '' = biru, 'g' = abu-abu."""
    st.markdown(SUP_CSS + '<div class="chips">' + "".join(
        f'<span class="chip {c}">{html.escape(t)}</span>' for t, c in items) + "</div>",
        unsafe_allow_html=True)


# Kartu supplier bisa diklik: tombol transparan menutupi seluruh kartu
SUP_CLICK_CSS = """<style>
[class*="st-key-supcard_"] {position: relative; margin-bottom: 4px;}
[class*="st-key-supcard_"] [data-testid="stMarkdownContainer"],
[class*="st-key-supcard_"] [data-testid="stMarkdown"] {margin-bottom: 0 !important; height: auto !important;}
[class*="st-key-supcard_"] .sup {margin: 0;}
[class*="st-key-supbtn_"] {position: absolute !important; inset: 0; z-index: 3; margin: 0 !important;
    width: 100% !important; height: 100% !important;}
[class*="st-key-supbtn_"] .stButton, [class*="st-key-supbtn_"] button {
    width: 100% !important; height: 100% !important; min-height: 100% !important;}
[class*="st-key-supbtn_"] button {opacity: 0; cursor: pointer; border: none !important;}
.sup {cursor: pointer;}
.sup.sel {border-width: 2px; box-shadow: 0 6px 18px rgba(13,127,204,.18);}
.sup.dim {opacity: .45;}
</style>"""


def _toggle_supplier(name: str):
    """Klik kartu yang sama lagi = hapus filter."""
    cur = st.session_state.get("find_supplier")
    st.session_state["find_supplier"] = None if cur == name else name


def supplier_rn_cards(df: pd.DataFrame, inv: pd.Series, selected=None):
    """Kartu total Room Night per supplier (SUPPLIER_RN_CARDS) + 'Other'. Setiap kartu bisa
    diklik untuk memfilter seluruh isi tab ini ke supplier tersebut."""
    if "Supplier Name" not in df.columns or "Room Night" not in df.columns:
        return
    grp = supplier_group(df)
    g = df.assign(_g=grp.values, _inv=inv.values).groupby("_g").agg(
        rn=("Room Night", "sum"), bk=("_inv", "nunique"))
    tot = g["rn"].sum() or 1
    cat_of = dict(zip(SUPPLIER_RN_CARDS, supplier_category(pd.Series(SUPPLIER_RN_CARDS))))
    names = SUPPLIER_RN_CARDS + ["Other"]
    stats = pd.DataFrame({"name": names,
                          "rn": [g["rn"].get(n, 0) for n in names],
                          "bk": [g["bk"].get(n, 0) for n in names],
                          "color": [SUPPLIER_COLORS.get(n, "#7A7A7A") for n in names]})
    st.markdown(SUP_CSS, unsafe_allow_html=True)
    st.markdown(SUP_CLICK_CSS, unsafe_allow_html=True)
    for r0 in range(0, len(names), 5):
        cols = st.columns(5, gap="small")
        for col, name in zip(cols, names[r0:r0 + 5]):
            i = names.index(name)
            rn_, bk_ = g["rn"].get(name, 0), g["bk"].get(name, 0)
            c_ = SUPPLIER_COLORS.get(name, "#7A7A7A")
            cls = "sup"
            if selected == name:
                cls += " sel"
            elif selected or not rn_:
                cls += " dim"
            tip = html.escape(f"{name} · kategori: {cat_of.get(name, 'supplier di luar daftar')} · "
                              + ("klik lagi untuk menghapus filter" if selected == name
                                 else "klik untuk memfilter"))
            style = f' style="border-color:{c_};background:{c_}0F"' if selected == name else ""
            with col.container(key=f"supcard_{i}"):
                st.markdown(
                    f'<div class="{cls}" title="{tip}"{style}>'
                    f'<div class="nm"><span class="dot" style="background:{c_}"></span>'
                    f'{html.escape(name)}</div>'
                    f'<div class="v">{angka(rn_)}<small>RN</small></div>'
                    f'<div class="s">{rn_ / tot * 100:.1f}% · {angka(bk_)} booking</div>'
                    f'<div class="b"><span style="width:{rn_ / tot * 100:.1f}%;background:{c_}">'
                    f'</span></div></div>', unsafe_allow_html=True)
                st.button(name, key=f"supbtn_{i}", on_click=_toggle_supplier, args=(name,),
                          disabled=not rn_ and selected != name)
    return stats


SCORECARD_CSS = """<style>
.sc-wrap {display:grid; grid-template-columns:repeat(auto-fit,minmax(220px,1fr)); gap:14px; margin:6px 0 18px;}
.sc {border-radius:14px; padding:16px 18px; color:#fff; box-shadow:0 2px 8px rgba(0,0,0,.12);}
.sc .lbl {font-size:.85rem; opacity:.9; text-transform:uppercase; letter-spacing:.04em;}
.sc .val {font-size:clamp(1.35rem,2.3vw,2rem); font-weight:700; margin:4px 0 2px; word-break:break-word;}
.sc .sub {font-size:.8rem; opacity:.85;}
.sc.rev {background:linear-gradient(135deg,#0A6AAC,#0D7FCC);}
.sc.rn  {background:linear-gradient(135deg,#0F7AA6,#1494C6);}
.sc.bk  {background:linear-gradient(135deg,#2B2B2B,#434343);}
</style>"""


def scorecards(revenue: float, room_night: float, booking: int, n_rows: int):
    adr = rupiah_full(revenue / room_night) if room_night else "-"
    per_bk = rupiah_full(revenue / booking) if booking else "-"
    st.markdown(SCORECARD_CSS + f"""
<div class="sc-wrap">
  <div class="sc rev"><div class="lbl">💰 Revenue</div>
    <div class="val">{rupiah_full(revenue)}</div><div class="sub">Total Sales AR</div></div>
  <div class="sc rn"><div class="lbl">🛏️ Total Room Night</div>
    <div class="val">{angka(room_night)}</div><div class="sub">Room × Night · rata-rata {adr} / RN</div></div>
  <div class="sc bk"><div class="lbl">🧾 Total Booking</div>
    <div class="val">{angka(booking)}</div>
    <div class="sub">Invoice No. unik · {angka(n_rows)} baris · {per_bk} / booking</div></div>
</div>""", unsafe_allow_html=True)


with tab_find:
    hcol = "Hotel (Grup)"
    fx = {}  # isi tab untuk ekspor PDF
    if hcol not in fdata.columns:
        st.info("Kolom Hotel Name tidak tersedia pada data.")
    else:
        find_pdf_box = st.container()  # tombol Export PDF (diisi setelah semua bagian dihitung)
        # ---------------- Kartu filter ----------------
        base = fdata
        with st.container(border=True, key="card_find_filter"):
            section("Pilih Hotel", "kosongkan untuk melihat semua hotel · mengikuti filter di sidebar")
            fa, fb, fc = st.columns([1.7, 3, 1])
            if "Product Type" in fdata.columns:
                mkt = market_type(fdata["Product Type"])
                m_opts = ["Semua"] + [x for x in ["Domestic", "International", "Lainnya"]
                                      if (mkt == x).any()]
                grp = (fdata.assign(_m=mkt).groupby("_m")["Product Type"].unique()
                       .map(lambda v: ", ".join(sorted(map(str, v)))))
                m_help = "Dari kolom Product Type — " + " · ".join(
                    f"{k}: {v}" for k, v in grp.items())
                if hasattr(st, "segmented_control"):  # tombol pilihan ringkas (Streamlit baru)
                    sel_m = fa.segmented_control("Pasar", m_opts, default="Semua",
                                                 key="find_market", help=m_help) or "Semua"
                else:
                    sel_m = fa.radio("Pasar", m_opts, horizontal=True, key="find_market",
                                     help=m_help)
                if sel_m != "Semua":
                    base = fdata[mkt == sel_m]
            else:
                fa.caption("Kolom Product Type tidak ada.")

            h_opts = (base.groupby(hcol)["Sales AR"].sum().sort_values(ascending=False).index
                      .astype(str).tolist())
            urut = fc.selectbox("Urutkan", ["Revenue terbesar", "A–Z"], key="find_sort")
            if urut == "A–Z":
                h_opts = sorted(h_opts)
            # Buang pilihan hotel yang tidak tersedia lagi setelah filter pasar berubah
            if "find_hotel" in st.session_state:
                st.session_state["find_hotel"] = [h for h in st.session_state["find_hotel"]
                                                  if h in h_opts]
            sel_h = fb.multiselect("Nama hotel", h_opts, key="find_hotel",
                                   placeholder="Ketik nama hotel… (kosong = semua hotel)")
            hd_all = base[base[hcol].astype(str).isin(sel_h)] if sel_h else base
            inv_all = (invoice_key(hd_all["Invoice No"]) if "Invoice No" in hd_all
                       else pd.Series(hd_all.index, index=hd_all.index))
            # Filter supplier dari kartu yang diklik (berlaku untuk semua isi tab ini)
            sel_sup = st.session_state.get("find_supplier")
            if sel_sup and "Supplier Name" in hd_all.columns and not hd_all.empty:
                hd = hd_all[(supplier_group(hd_all) == sel_sup).values]
            else:
                sel_sup, hd = None, hd_all

            if not hd.empty:
                inv = invoice_key(hd["Invoice No"]) if "Invoice No" in hd else pd.Series(hd.index)
                info = [(f"{angka(hd[hcol].nunique())} hotel", ""),
                        (f"{angka(col_sum(hd, 'Room Night'))} Room Night", ""),
                        (f"{angka(inv.nunique())} booking", ""),
                        (rupiah_full(col_sum(hd, "Sales AR")), "g")]
                if sel_sup:
                    info.insert(0, (f"Supplier: {sel_sup}", ""))
                if len(sel_h) == 1:  # info hotel tunggal: kota, chain, rating
                    for c, lbl in [("Hotel City", ""), ("Hotel Chain", "Chain: "),
                                   ("Hotel Rating", "")]:
                        if c in hd.columns and hd[c].notna().any():
                            v = hd[c].dropna().astype(str).mode().iat[0]
                            info.append(("⭐" * int(float(v)) if c == "Hotel Rating"
                                         else f"{lbl}{str(v).title() if c == 'Hotel City' else v}",
                                         "g"))
                chips(info)
                fx["kpi"] = [("Hotel", angka(hd[hcol].nunique()), "#0D7FCC"),
                             ("Room Night", angka(col_sum(hd, "Room Night")), "#1BA0E2"),
                             ("Booking", angka(inv.nunique()), "#1494C6"),
                             ("Revenue", rupiah_full(col_sum(hd, "Sales AR")), "#434343")]
                if "Inv Date" in hd.columns and hd["Inv Date"].notna().any():
                    fx["periode"] = (f"{hd['Inv Date'].min():%d %b %Y} – "
                                     f"{hd['Inv Date'].max():%d %b %Y}")
                if len(sel_h) == 1 and "Hotel Name (Asli)" in hd.columns:
                    var = sorted(hd["Hotel Name (Asli)"].dropna().astype(str).unique())
                    if len(var) > 1:
                        st.caption(f"Penulisan asli ({len(var)}): " + "; ".join(var[:10])
                                   + (" …" if len(var) > 10 else ""))

        if hd_all.empty:
            st.warning("Tidak ada transaksi untuk hotel yang dipilih.")
        else:
            # ---------------- Room Night per supplier (kartu = filter) ----------------
            with st.container(border=True, key="card_find_sup"):
                s1, s2 = st.columns([5, 1])
                section("Room Night per Supplier",
                        f"filter aktif: {sel_sup} — semua isi tab mengikuti" if sel_sup else
                        "klik kartu untuk memfilter semua isi tab · Other = supplier di luar daftar",
                        box=s1)
                if sel_sup:
                    s2.button("✕ Hapus filter", key="clear_sup", on_click=_toggle_supplier,
                              args=(sel_sup,), width="stretch")
                fx["sup"] = supplier_rn_cards(hd_all, inv_all, selected=sel_sup)

        if hd.empty and not hd_all.empty:
            st.info("Tidak ada transaksi untuk supplier ini pada pilihan sekarang.")
        if not hd.empty:

            # ---------------- Tren bulanan per supplier ----------------
            if "Bulan" in hd.columns:
                mm = (hd.assign(_inv=inv.values).groupby("Bulan")
                      .agg(RoomNight=("Room Night", "sum"), Booking=("_inv", "nunique"),
                           Revenue=("Sales AR", "sum")).reset_index())
                mm["Label"] = (mm["Bulan"].dt.month.map(lambda x: BULAN[x - 1]) + " "
                               + mm["Bulan"].dt.strftime("%y"))
                with st.container(border=True, key="card_find_trend"):
                    # --- Total Room Night per bulan (semua supplier) ---
                    _h1, _h2 = st.columns([5, 1], vertical_alignment="center")
                    section("Total Room Night per Bulan",
                            "semua supplier" + (f" · filter supplier: {sel_sup}" if sel_sup else ""),
                            box=_h1)
                    _trend_btn = _h2.empty()  # tombol Excel, diisi setelah tabel supplier dihitung
                    ft = go.Figure()
                    ft.add_scatter(x=mm["Label"], y=mm["RoomNight"], name="Total Room Night",
                                   mode="lines+markers+text", text=[angka(v) for v in mm["RoomNight"]],
                                   textposition="top center", textfont=dict(size=11, color="#434343"),
                                   line=dict(color=PALETTE[0], width=3.5, shape="spline", smoothing=0.6),
                                   marker=dict(size=8), fill="tozeroy",
                                   fillcolor="rgba(13,127,204,0.08)",
                                   hovertemplate="Total: %{y:,.0f} RN<extra></extra>")
                    clean_fig(ft, 260)
                    _ymax = float(mm["RoomNight"].max() or 1)
                    ft.update_layout(showlegend=False, hovermode="x unified",
                                     yaxis=dict(title=None, range=[0, _ymax * 1.22]),
                                     margin=dict(t=20, b=10, l=30, r=30))
                    ft.update_traces(cliponaxis=False)
                    st.plotly_chart(ft, width="stretch", key="find_trend_total",
                                    config={"displayModeBar": False})

                    section("Tren Room Night per Supplier",
                            "klik nama di legenda untuk menyembunyikan / menampilkan garis")
                    fg = go.Figure()
                    if "Supplier Name" in hd.columns:
                        sm = (hd.assign(_g=supplier_group(hd).values).groupby(["Bulan", "_g"])
                              ["Room Night"].sum().unstack(fill_value=0)
                              .reindex(mm["Bulan"], fill_value=0))
                        order = [s for s in SUPPLIER_RN_CARDS + ["Other"]
                                 if s in sm.columns and sm[s].sum() > 0]
                        fx["trend"] = pd.DataFrame({"Label": mm["Label"].values,
                                                    **{s: sm[s].values for s in order}})
                        for s in order:
                            fg.add_scatter(x=mm["Label"], y=sm[s].values, name=s,
                                           mode="lines+markers",
                                           line=dict(color=SUPPLIER_COLORS.get(s, "#7A7A7A"),
                                                     width=2.5, shape="spline", smoothing=0.6,
                                                     dash="dot" if s == "Other" else "solid"),
                                           marker=dict(size=6),
                                           hovertemplate=f"{s}: " + "%{y:,.0f} RN<extra></extra>")
                    clean_fig(fg, 380)
                    fg.update_layout(hovermode="x unified",
                                     yaxis=dict(title=None, rangemode="tozero"),
                                     legend=dict(orientation="h", y=1.1, x=0, xanchor="left",
                                                 title=None, font=dict(size=12)))
                    st.plotly_chart(fg, width="stretch", key="find_trend",
                                    config={"displayModeBar": False})

                    # --- Unduh tabel tren (Excel) ---
                    def make_trend_xlsx(mm=mm.copy(), sm=(sm.copy() if "Supplier Name" in hd.columns
                                                         else None),
                                        order=(list(order) if "Supplier Name" in hd.columns else []),
                                        sel_h=list(sel_h), sel_sup=sel_sup,
                                        pasar=st.session_state.get("find_market") or "Semua",
                                        periode=fx.get("periode", "-")):
                        tot = pd.DataFrame({
                            "Bulan": mm["Label"].values,
                            "Total Room Night": mm["RoomNight"].astype(int).values,
                            "Booking (Invoice unik)": mm["Booking"].astype(int).values,
                            "Revenue (Sales AR)": round_money(mm["Revenue"].astype(float)).values,
                        })
                        tot.loc[len(tot)] = ["TOTAL", tot["Total Room Night"].sum(),
                                             int(inv.nunique()),
                                             round_money(float(mm["Revenue"].sum()))]
                        sheets = {"Total per Bulan": tot}
                        if sm is not None and len(order):
                            piv = sm[order].astype(int).copy()
                            piv.insert(0, "Bulan", mm["Label"].values)
                            piv["Total Semua Supplier"] = sm.sum(axis=1).astype(int).values
                            piv = piv.reset_index(drop=True)
                            piv.loc[len(piv)] = ["TOTAL"] + [int(piv[c].sum()) for c in piv.columns[1:]]
                            sheets["Room Night per Supplier"] = piv
                        card = [("Tren Room Night per Bulan", None, "title"),
                                ("Filter", None, None),
                                ("Periode Inv Date", periode, "text"),
                                ("Pasar", pasar, "text"),
                                ("Supplier", sel_sup or "Semua", "text"),
                                ("Hotel", "Semua hotel" if not sel_h else "; ".join(sel_h), "text"),
                                ("Diunduh", pd.Timestamp.now().strftime("%d %b %Y %H:%M"), "text")]
                        return excel_book(sheets, scorecard=card)

                    _fn_t = "tren_room_night" + (f"_{re.sub(r'[^A-Za-z0-9]+', '_', sel_sup)}"
                                                 if sel_sup else "")
                    excel_download(_trend_btn, "⬇️ Download Excel", make_trend_xlsx,
                                   f"{_fn_t}.xlsx", "dl_find_trend_xlsx", width="stretch")

            # ---------------- Ringkasan per hotel ----------------
            g = hd.assign(_inv=inv.values).groupby(hcol)
            per_h = g.agg(Revenue=("Sales AR", "sum"), **{"Room Night": ("Room Night", "sum"),
                                                          "Booking": ("_inv", "nunique")})
            for c, lbl in [("Hotel City", "Kota"), ("Hotel Chain", "Chain"),
                           ("Hotel Rating", "Rating")]:
                if c in hd.columns:
                    per_h[lbl] = g[c].agg(lambda s: s.dropna().mode().iat[0]
                                          if s.notna().any() else None)
            if "Kota" in per_h.columns:
                per_h["Kota"] = per_h["Kota"].map(lambda v: str(v).title() if pd.notna(v) else v)
            per_h = per_h[[c for c in ["Kota", "Chain", "Rating"] if c in per_h.columns]
                          + ["Revenue", "Room Night", "Booking"]]
            per_h["Rata-rata / Room Night"] = per_h["Revenue"] / per_h["Room Night"].where(
                per_h["Room Night"] > 0)
            per_h["Share RN %"] = per_h["Room Night"] / (per_h["Room Night"].sum() or 1) * 100
            for c in ("Revenue", "Rata-rata / Room Night"):
                per_h[c] = round_money(per_h[c].astype(float))
            per_h = (per_h.sort_values("Revenue", ascending=False).reset_index()
                     .rename(columns={hcol: "Hotel"}))
            fx["table"] = per_h
            with st.container(border=True, key="card_find_table"):
                t1_, t2_ = st.columns([5, 1], vertical_alignment="center")
                section("Ringkasan per Hotel", f"{angka(len(per_h))} hotel · urut Revenue terbesar",
                        box=t1_)

                def make_hotel_xlsx(per_h=per_h, fx=dict(fx), sel_h=list(sel_h), sel_sup=sel_sup,
                                    pasar=st.session_state.get("find_market") or "Semua"):
                    card = [("Ringkasan per Hotel", None, "title"),
                            ("Filter", None, None),
                            ("Periode Inv Date", fx.get("periode", "-"), "text"),
                            ("Pasar", pasar, "text"),
                            ("Supplier", sel_sup or "Semua", "text"),
                            ("Hotel", "Semua hotel" if not sel_h else "; ".join(sel_h), "text"),
                            ("Diunduh", pd.Timestamp.now().strftime("%d %b %Y %H:%M"), "text")]
                    out = per_h.rename(columns={"Share RN %": "Share Room Night %"})
                    return excel_book({"Ringkasan per Hotel": out}, scorecard=card)

                _fn = "ringkasan_per_hotel" + (f"_{re.sub(r'[^A-Za-z0-9]+', '_', sel_sup)}"
                                               if sel_sup else "")
                excel_download(t2_, "⬇️ Download Excel", make_hotel_xlsx, f"{_fn}.xlsx",
                               "dl_find_xlsx", width="stretch")
                st.dataframe(
                    per_h, hide_index=True, width="stretch",
                    height=min(460, 38 + 35 * len(per_h)),
                    column_config={
                        "Hotel": st.column_config.TextColumn(width="large"),
                        "Revenue": st.column_config.NumberColumn(format="localized"),
                        "Room Night": st.column_config.NumberColumn(format="localized"),
                        "Booking": st.column_config.NumberColumn(format="localized"),
                        "Rata-rata / Room Night": st.column_config.NumberColumn(
                            "Avg / RN", format="localized",
                            help="Rata-rata Revenue per Room Night"),
                        "Share RN %": st.column_config.ProgressColumn(
                            "Share RN", format="%.1f%%", min_value=0,
                            max_value=float(per_h["Share RN %"].max() or 100)),
                        "Rating": st.column_config.NumberColumn(format="%d ⭐"),
                    })

    if fx.get("kpi"):
        fx["pasar"] = st.session_state.get("find_market") or "Semua"
        fx["supplier"] = st.session_state.get("find_supplier")
        fx["hotels"] = list(st.session_state.get("find_hotel") or [])
        # Ask Claude di bagian paling bawah tab (setelah Ringkasan per Hotel)
        _facts_find = hotel_facts(fx)
        # Data yang bisa ditanyakan = hasil filter pasar + hotel + supplier di tab ini
        ask_claude_box("find", hd, "Filter per Hotel", _facts_find, ASK_SUGGEST_HOTEL)
        with find_pdf_box:
            _p1, _p2 = st.columns([5, 1], vertical_alignment="center")
            _p1.caption("Ringkasan tab ini (sesuai filter & supplier yang dipilih) bisa diunduh "
                        "sebagai PDF satu halaman (A4).")
            _per = re.sub(r"[^A-Za-z0-9]+", "_", fx.get("periode", ""))[:40].strip("_")
            file_download(_p2, "⬇️ Export PDF", lambda f=dict(fx): hotel_pdf(f),
                          f"ringkasan_hotel_{_per or 'periode'}.pdf", "application/pdf",
                          "dl_find_pdf", width="stretch")

with tab_data:
    st.subheader("Preview data gabungan (setelah filter)")
    st.caption(f"{angka(len(fdata))} baris. Tabel menampilkan maksimal 5.000 baris pertama.")
    st.dataframe(fdata.head(5000), width="stretch", height=450)

    st.subheader("Unduh hasil merge")
    scope = st.radio("Data yang diekspor", ["Sesuai filter", "Semua data gabungan"], horizontal=True)
    export_df = fdata if scope == "Sesuai filter" else data
    cols = st.multiselect("Kolom yang diekspor (kosong = semua)", list(export_df.columns))
    if cols:
        export_df = export_df[cols]

    fmt = st.radio("Format", ["CSV", "Excel (.xlsx)"], horizontal=True)
    # Tanda pengenal ekspor: file lama tidak dipakai lagi jika filter/kolom berubah
    sig = (scope, fmt, tuple(export_df.columns), len(export_df),
           float(col_sum(export_df, "Sales AR")))

    if fmt.startswith("Excel") and len(export_df) > 1_048_575:
        st.error("Melebihi batas baris Excel (1.048.576). Gunakan CSV.")
    elif st.button("⚙️ Siapkan file unduhan"):
        with st.spinner("Menyiapkan file... (data besar bisa butuh beberapa saat)"):
            if fmt == "CSV":
                payload = export_df.to_csv(index=False).encode("utf-8-sig")
            else:
                payload = to_excel_bytes(export_df, summary,
                                         {"Hotel Belum Standar": unmatched_hotels()})
            st.session_state["export"] = (sig, payload)

    if st.session_state.get("export", (None,))[0] == sig:
        is_csv = fmt == "CSV"
        st.download_button(
            "⬇️ Unduh " + ("CSV" if is_csv else "Excel"),
            st.session_state["export"][1],
            file_name="merged_data.csv" if is_csv else "merged_data.xlsx",
            mime="text/csv" if is_csv else
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary")

with tab_hotel:
    if std_table is not None:
        st.subheader("Standarisasi nama hotel (master STADARD NAME)")
        um = unmatched_hotels()
        n_tot = len(std_table)
        n_ok = int(std_table["Standard Name"].notna().sum())
        u1, u2, u3 = st.columns(3)
        u1.metric("Nama hotel asli", angka(n_tot))
        u2.metric("Cocok dengan master", angka(n_ok),
                  help=f"Persis: {angka((std_table['Cocok'] == 'Persis').sum())} · "
                       f"beda huruf besar/kecil saja: "
                       f"{angka((std_table['Cocok'] == 'Huruf besar-kecil').sum())}")
        u3.metric("Belum ada di master", angka(n_tot - n_ok))
        with st.expander(f"⚠️ Nama hotel yang belum ada di master ({angka(len(um))})",
                         expanded=False):
            st.caption("Nama ini tetap memakai penulisan asli. Isi kolom STADARD NAME di file CSV "
                       "unduhan, lalu salin ke file master agar ikut distandarkan.")
            st.dataframe(um, hide_index=True, width="stretch", height=320,
                         column_config={"Sales AR": st.column_config.NumberColumn(
                             format="localized")})
            x1, x2 = st.columns(2)
            excel_download(x1, "⬇️ Unduh daftar hotel belum standar (Excel)",
                           lambda um=um: excel_book({"Hotel Belum Standar": um}),
                           "hotel_belum_standar.xlsx", "dlx_unmatched")
            x2.download_button("⬇️ Unduh daftar hotel belum standar (CSV)",
                               um.to_csv(index=False).encode("utf-8-sig"),
                               file_name="hotel_belum_standar.csv", mime="text/csv",
                               key="dl_unmatched")
        st.divider()

    if not hotel_on or mapping is None:
        st.info("Aktifkan **Gabungkan nama hotel yang mirip** di sidebar untuk melihat "
                "dan mengoreksi pengelompokan hotel.")
    else:
        st.subheader("Pengelompokan nama hotel yang mirip")
        st.caption("Logika: nama dinormalisasi (huruf kecil, tanpa tanda baca/isi kurung, "
                   "'&' = 'and', tanpa nama kota & kata umum), lalu dibandingkan dengan skor "
                   "kemiripan di dalam kota yang sama. Kata pertama (merek) harus mirip dan "
                   "angka di nama harus sama, supaya mis. 'Santika Pasir Kaliki' tidak "
                   "tergabung dengan 'Santika Pasir Koja'. Nama grup = varian dengan transaksi "
                   "terbanyak.")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Nama hotel asli", angka(len(mapping)))
        m2.metric("Hotel setelah digabung", angka(mapping["Hotel (Grup)"].nunique()))
        m3.metric("Varian yang tergabung", angka((mapping["Jumlah Varian"] > 1).sum()))
        m4.metric("Koreksi manual", angka(len(overrides)))

        s1, s2 = st.columns([3, 1])
        q = s1.text_input("Cari nama hotel / kota", placeholder="mis. crowne, marriott, bandung")
        only_multi = s2.checkbox("Hanya grup > 1 varian", value=True)
        view = mapping
        if only_multi:
            view = view[view["Jumlah Varian"] > 1]
        if q:
            ql = q.lower()
            view = view[view["Hotel Name"].str.lower().str.contains(ql, regex=False, na=False)
                        | view["Hotel (Grup)"].str.lower().str.contains(ql, regex=False, na=False)
                        | view["Hotel City"].astype(str).str.lower().str.contains(ql, regex=False)]
        view = view.sort_values(["Hotel (Grup)", "Transaksi"], ascending=[True, False])

        st.markdown("**Koreksi manual:** ubah isi kolom *Hotel (Grup)* lalu klik *Terapkan*. "
                    "Untuk **memisahkan** varian yang salah gabung, isi dengan nama aslinya; "
                    "untuk **menggabungkan**, isi dengan nama grup tujuan.")
        cols_show = ["Hotel City", "Hotel Name", "Hotel (Grup)", "Skor", "Transaksi",
                     "Jumlah Varian", "Diubah Manual", "Key"]
        with st.form("hotel_edit"):
            edited = st.data_editor(
                view[cols_show], hide_index=True, width="stretch", height=450,
                disabled=[c for c in cols_show if c != "Hotel (Grup)"],
                column_config={"Key": None,
                               "Skor": st.column_config.NumberColumn("Skor Kemiripan",
                                                                     format="%.1f"),
                               "Hotel (Grup)": st.column_config.TextColumn(
                                   "Hotel (Grup) ✏️", required=True)},
                key="hotel_editor")
            if st.form_submit_button("💾 Terapkan perubahan", type="primary"):
                changed = edited[edited["Hotel (Grup)"].str.strip()
                                 != view["Hotel (Grup)"].values]
                for k_, g_ in zip(changed["Key"], changed["Hotel (Grup)"]):
                    overrides[k_] = g_.strip()
                st.rerun()

        b1, b2, b3 = st.columns(3)
        if b1.button("↩️ Hapus semua koreksi manual", disabled=not overrides):
            overrides.clear()
            st.rerun()
        b2.download_button(
            "⬇️ Unduh mapping hotel (CSV)",
            mapping[["Hotel City", "Hotel Name", "Hotel (Grup)", "Skor", "Transaksi",
                     "Diubah Manual"]].to_csv(index=False).encode("utf-8-sig"),
            file_name="mapping_hotel.csv", mime="text/csv")
        up = b3.file_uploader("Muat mapping hotel (CSV)", type=["csv"],
                              help="CSV hasil unduhan di samping (boleh diedit di Excel). "
                                   "Wajib berisi kolom Hotel City, Hotel Name, Hotel (Grup).")
        if up is not None and st.session_state.get("map_file_id") != up.file_id:
            mp = pd.read_csv(up, dtype=str, keep_default_na=False)
            need = {"Hotel City", "Hotel Name", "Hotel (Grup)"}
            if need <= set(mp.columns):
                base = mapping.set_index("Key")["Hotel (Grup)"]
                keys = hotel_key(mp["Hotel City"].replace("", "nan"), mp["Hotel Name"])
                for k_, g_ in zip(keys, mp["Hotel (Grup)"]):
                    if g_.strip() and k_ in base.index and base[k_] != g_.strip():
                        overrides[k_] = g_.strip()
                st.session_state["map_file_id"] = up.file_id
                st.rerun()
            else:
                st.error(f"Kolom wajib tidak lengkap: {', '.join(sorted(need - set(mp.columns)))}")

with tab_file:
    st.subheader("Ringkasan per file")
    per_file = (data.assign(_inv=invoice_key(data["Invoice No"]) if "Invoice No" in data
                            else data.index)
                .groupby("Source File")
                .agg(Baris=("Source File", "size"), **{"Invoice Unik": ("_inv", "nunique")},
                     Sales_AR=("Sales AR", "sum"), Profit=("Profit", "sum"))
                .reset_index())
    st.dataframe(summary.merge(per_file, left_on="File", right_on="Source File", how="left")
                 .drop(columns=["Source File", "Baris"]),
                 width="stretch", hide_index=True)
    st.caption("Kolom 'Periode (dari judul)' diambil dari baris judul di dalam file, "
               "berguna untuk mengecek apakah nama file sesuai isinya.")
