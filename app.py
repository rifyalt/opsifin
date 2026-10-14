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

import html
import io
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
</style>""", unsafe_allow_html=True)

EXCEL_EXT = (".xlsx", ".xlsm", ".xls")

DATE_COLS = ["Inv Date", "Due Date", "Issued Date", "Check In", "Check Out",
             "Departure Date", "Arrival Date"]

NUMERIC_COLS = ["Room", "Night", "Base Fare", "Fare Tax", "IWJR", "Add Charge",
                "Insurance", "PSC", "Other Charge", "Incentive", "Agent Comm", "NTA",
                "Travel Services", "Sales Net", "VAT", "Stamp Fee", "MDR", "Sales AR",
                "Extra Disc", "Profit", "Profit %", "Rounding", "Base Sell"]

BULAN = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"]


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
            .str.replace(" ", " ").str.replace(r"\s+", " ", regex=True).str.strip())


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
def to_excel_bytes(df: pd.DataFrame, summary: pd.DataFrame, extra: dict = None) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter",
                        engine_kwargs={"options": {"strings_to_urls": False}}) as w:
        df.to_excel(w, sheet_name="Merged Data", index=False)
        summary.to_excel(w, sheet_name="Ringkasan File", index=False)
        for name, sheet in (extra or {}).items():
            if sheet is not None and len(sheet):
                sheet.to_excel(w, sheet_name=name[:31], index=False)
    return buf.getvalue()


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
]


def customer_norm(sr: pd.Series) -> pd.Series:
    """Normalisasi untuk pencocokan: huruf besar, spasi rapi, tanpa spasi sebelum '('.
    Jadi 'CBT PERTAMINA (HOTEL)', 'CBT PERTAMINA(HOTEL)' dan 'cbt pertamina (hotel)' sama."""
    s_ = sr.astype(str).str.upper().str.replace(r"\s+", " ", regex=True).str.strip()
    return s_.str.replace(r"\s*\(\s*", "(", regex=True).str.replace(r"\s*\)", ")", regex=True)


cx = st.expander("👤 Aturan Merge Customer (gabungkan beberapa nama customer menjadi satu)")
if "Customer Name" in data.columns:
    cust_on = cx.toggle("Gabungkan nama customer", value=True,
                                help="Nilai customer di kolom kiri dijumlahkan ke customer di "
                                     "kolom kanan, dan yang ditampilkan hanya nama di kolom kanan.")
    if "customer_rules" not in st.session_state:
        st.session_state["customer_rules"] = pd.DataFrame(DEFAULT_CUSTOMER_RULES)
    cust_opts = sorted(data["Customer Name"].dropna().astype(str).unique())
    rules = cx.data_editor(
        st.session_state["customer_rules"], num_rows="dynamic", hide_index=True,
        width="stretch", key="customer_rules_editor", disabled=not cust_on,
        column_config={
            "Nama Customer Asli": st.column_config.SelectboxColumn(
                "Nama customer asli", options=cust_opts, required=True, width="large"),
            "Digabung Menjadi": st.column_config.TextColumn("Digabung menjadi (nama yang ditampilkan)",
                                                             required=True, width="large"),
        })
    cx.caption("Tambah baris (+) untuk aturan baru, mis. "
                       "'CBT PERTAMINA(HOTEL CM)' → 'CBT PERTAMINA'.")

    if cust_on:
        r_ = rules.dropna(subset=["Nama Customer Asli", "Digabung Menjadi"])
        r_ = r_[r_["Digabung Menjadi"].astype(str).str.strip().ne("")]
        rule_map = dict(zip(customer_norm(r_["Nama Customer Asli"]),
                            r_["Digabung Menjadi"].astype(str).str.strip()))
        if rule_map:
            orig = data["Customer Name"]
            mapped = customer_norm(orig).map(rule_map)
            n_rows = int(mapped.notna().sum())
            data = data.assign(**{"Customer Name": mapped.fillna(orig),
                                  "Customer Name (Asli)": orig})
            cx.caption(f"{angka(n_rows)} baris digabung ke {len(set(rule_map.values()))} "
                               f"customer.")

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
fdata = multi_filter(fdata, "Hotel City", "Kota Hotel")
top_n = st.sidebar.slider("Top N pada grafik", 5, 30, 10)

# Filter utama di area dashboard: Supplier Name & Hotel
st.markdown("##### 🔎 Filter Dashboard")
fc1, fc2 = st.columns(2)


def include_exclude_filter(df: pd.DataFrame, col: str, label: str, box, key: str) -> pd.DataFrame:
    """Filter dengan mode Include (hanya yang dipilih) atau Exclude (kecualikan yang dipilih),
    ditambah pengecualian berdasarkan kata kunci (mis. 'channel manage')."""
    if col not in df.columns:
        return df
    mode = box.radio(f"Mode filter {label}", ["Include", "Exclude"], horizontal=True,
                     key=f"mode_{key}",
                     help="Include: tampilkan HANYA yang dipilih. "
                          "Exclude: tampilkan semua KECUALI yang dipilih.")
    opts = sorted(data[col].dropna().astype(str).unique())
    sel = box.multiselect(label, opts, key=f"sel_{key}",
                          placeholder="Semua" if mode == "Include" else "Tidak ada yang dikecualikan")
    kw = box.text_input(f"Kecualikan {label} yang mengandung kata (pisahkan dengan koma)",
                        key=f"kw_{key}", placeholder="mis. channel manage, expedia")

    vals = df[col].astype(str)
    if sel:
        df = df[vals.isin(sel)] if mode == "Include" else df[~vals.isin(sel)]
        vals = df[col].astype(str)
    words = [w.strip().lower() for w in kw.split(",") if w.strip()]
    if words:
        pattern = "|".join(re.escape(w) for w in words)
        hit = vals.str.lower().str.contains(pattern, regex=True, na=False)
        excluded = sorted(vals[hit].unique())
        df = df[~hit]
        box.caption(f"Dikecualikan ({len(excluded)}): " + (", ".join(excluded[:8]) or "-")
                    + (" …" if len(excluded) > 8 else ""))
    return df


fdata = include_exclude_filter(fdata, "Supplier Name", "Supplier Name", fc1, "supplier")
fdata = include_exclude_filter(fdata, "Hotel (Grup)",
                               "Hotel (sudah digabung)" if hotel_on else "Hotel", fc2, "hotel")

if fdata.empty:
    st.warning("Tidak ada data yang cocok dengan filter.")
    st.stop()

# ---------------------------------------------------------------------------
# Tabs
# ---------------------------------------------------------------------------
tab_dash, tab_top, tab_find, tab_data, tab_hotel, tab_file = st.tabs(
    ["📈 Dashboard", "🏆 Top Hotel", "🔍 Filter per Hotel", "🗂️ Data Gabungan", "🏨 Grup Hotel",
     "📁 Ringkasan File"])


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


def col_sum(df, c):
    return df[c].sum() if c in df.columns else 0


with tab_dash:
    sales = col_sum(fdata, "Sales AR")
    profit = col_sum(fdata, "Profit")
    n_inv = invoice_key(fdata["Invoice No"]).nunique() if "Invoice No" in fdata else len(fdata)
    rn = col_sum(fdata, "Room Night")
    k = st.columns(3) + st.columns(3)
    k[0].metric("Total Sales AR", rupiah_full(sales))
    k[1].metric("Total Profit", rupiah_full(profit))
    k[2].metric("Margin", f"{(profit / sales * 100 if sales else 0):.2f}%")
    k[3].metric("Jumlah Invoice (unik)", angka(n_inv),
                help=f"Dihitung dari nomor invoice unik. {angka(len(fdata))} baris transaksi "
                     f"berasal dari {angka(n_inv)} invoice, karena satu invoice bisa berisi "
                     f"beberapa baris (beberapa kamar, penyesuaian, dll.).")
    k[4].metric("Room Night", angka(rn))
    if "Hotel (Grup)" in fdata.columns:
        k[5].metric("Jumlah Hotel", angka(fdata["Hotel (Grup)"].nunique()))

    # Tren bulanan
    if "Bulan" in fdata.columns:
        m = (fdata.groupby("Bulan")
             .agg(Sales=("Sales AR", "sum"), Profit=("Profit", "sum"),
                  RoomNight=("Room Night", "sum"))
             .reset_index())
        m["Label"] = m["Bulan"].dt.month.map(lambda x: BULAN[x - 1]) + " " + m["Bulan"].dt.strftime("%y")
        fig = go.Figure()
        fig.add_bar(x=m["Label"], y=m["Sales"], name="Sales AR", marker_color="#2E86AB")
        fig.add_scatter(x=m["Label"], y=m["Profit"], name="Profit", yaxis="y2",
                        mode="lines+markers", line=dict(color="#F18F01", width=3))
        fig.update_layout(title="Tren Bulanan: Sales AR & Profit", height=420,
                          yaxis=dict(title="Sales AR (Rp)"),
                          yaxis2=dict(title="Profit (Rp)", overlaying="y", side="right"),
                          legend=dict(orientation="h", y=1.1), margin=dict(t=70, b=30))
        st.plotly_chart(fig, width="stretch")

    def top_bar(col, title, color, metric="Sales AR"):
        if col not in fdata.columns or metric not in fdata.columns:
            return None
        d = fdata[fdata[col].astype(str).str.strip().ne("-")]
        key = d[col].astype(str).str.strip()
        if col == "Hotel City":  # satukan 'JAKARTA' / 'Jakarta'
            key = key.str.title()
        agg = d.groupby(key).agg(**{"Sales AR": ("Sales AR", "sum"),
                                     "Room Night": ("Room Night", "sum")})
        agg.index.name = col
        t = agg.nlargest(top_n, metric).sort_values(metric).reset_index()
        t["Sales / Room Night"] = t["Sales AR"] / t["Room Night"].where(t["Room Night"] > 0)
        f = px.bar(t, x=metric, y=col, orientation="h", title=title,
                   color_discrete_sequence=[color], height=max(350, 28 * len(t) + 100),
                   text_auto=",.0f" if metric == "Room Night" else False,
                   hover_data={"Sales AR": ":,.0f", "Room Night": ":,.0f",
                               "Sales / Room Night": ":,.0f"})
        f.update_traces(textposition="outside", cliponaxis=False)
        f.update_layout(yaxis_title=None, margin=dict(t=50, b=20, r=40))
        return f

    def four_charts(metric, label, colors, key):
        c1, c2 = st.columns(2)
        for col_box, (col, name, color) in zip(
            [c1, c2, c1, c2],
            [("Customer Name", "Customer", colors[0]), ("Hotel (Grup)", "Hotel", colors[1]),
             ("Hotel City", "Kota", colors[2]), ("Supplier Name", "Supplier", colors[3])]):
            fig = top_bar(col, f"Top {top_n} {name} ({label})", color, metric)
            if fig is not None:
                col_box.plotly_chart(fig, width="stretch", key=f"{key}_{col}")

    st.markdown("#### 💰 Analisa Sales AR")
    four_charts("Sales AR", "Sales AR", ["#2E86AB", "#A23B72", "#3B8B5A", "#C73E1D"], "sales")

    # ---------------- Analisa Room Night (Room × Night) ----------------
    if "Room Night" in fdata.columns:
        st.divider()
        st.markdown("#### 🛏️ Analisa Room Night (Room × Night)")
        tot_rn = fdata["Room Night"].sum()
        tot_room = fdata["Room"].sum() if "Room" in fdata.columns else 0
        r = st.columns(3)
        r[0].metric("Total Room Night", angka(tot_rn))
        r[1].metric("Rata-rata Sales per Room Night",
                    rupiah_full(sales / tot_rn) if tot_rn else "-")
        r[2].metric("Rata-rata lama menginap", f"{tot_rn / tot_room:.2f} malam" if tot_room else "-",
                    help="Total Room Night ÷ total Room")
        four_charts("Room Night", "Room Night", ["#1B6F8A", "#7B2D5B", "#2A6B45", "#9C2F15"], "rn")
        st.divider()

    c3, c4 = st.columns(2)
    if "Branch" in fdata.columns:
        b = fdata.groupby("Branch")["Sales AR"].sum().reset_index()
        c3.plotly_chart(px.pie(b, names="Branch", values="Sales AR", hole=0.5,
                               title="Komposisi Sales per Branch"), width="stretch")
    if "Product Type" in fdata.columns:
        p = fdata.groupby("Product Type")["Sales AR"].sum().reset_index()
        c4.plotly_chart(px.pie(p, names="Product Type", values="Sales AR", hole=0.5,
                               title="Komposisi Sales per Product Type"), width="stretch")

@st.cache_data(show_spinner=False)
def hotel_ranking(df: pd.DataFrame) -> pd.DataFrame:
    """Ringkasan per hotel (grup): Sales AR, Profit, Room Night, invoice, ADR, kota utama."""
    hcol = "Hotel (Grup)"
    d = df[df[hcol].notna() & df[hcol].astype(str).str.strip().ne("-")]
    d = d.assign(**{"Invoice No": invoice_key(d["Invoice No"])})
    agg = d.groupby(hcol).agg(
        **{"Sales AR": ("Sales AR", "sum"), "Profit": ("Profit", "sum"),
           "Room Night": ("Room Night", "sum"), "Invoice (unik)": ("Invoice No", "nunique")})
    # Kota utama = kota dengan transaksi terbanyak untuk hotel tsb
    city = (d.groupby([hcol, "Hotel City"]).size().reset_index(name="n")
            .sort_values("n").drop_duplicates(hcol, keep="last").set_index(hcol)["Hotel City"])
    agg.insert(0, "Kota", city)
    agg["Rata-rata / Room Night"] = (agg["Sales AR"] / agg["Room Night"].where(agg["Room Night"] > 0))
    agg["Margin %"] = agg["Profit"] / agg["Sales AR"].where(agg["Sales AR"] != 0) * 100
    agg["Share Sales %"] = agg["Sales AR"] / agg["Sales AR"].sum() * 100
    agg["Share RN %"] = agg["Room Night"] / agg["Room Night"].sum() * 100
    return agg.reset_index().rename(columns={hcol: "Hotel"})


def ranking_section(rank: pd.DataFrame, metric: str, share_col: str, color: str,
                    unit_fmt, key: str):
    r = rank.sort_values(metric, ascending=False).reset_index(drop=True)
    r.insert(0, "Rank", range(1, len(r) + 1))
    top = r.head(top_n)

    total = r[metric].sum()
    c = st.columns(4)
    c[0].metric(f"Total {metric}", unit_fmt(total))
    c[1].metric(f"Kontribusi Top {top_n}", f"{top[metric].sum() / total * 100:.1f}%" if total else "-")
    if len(top):
        c[2].metric("Hotel #1", unit_fmt(top[metric].iat[0]))
        c[2].caption(f"🥇 {top['Hotel'].iat[0]} ({top['Kota'].iat[0]})")
    c[3].metric("Jumlah hotel", angka(len(r)))

    g1, g2 = st.columns([3, 2])
    chart = top.sort_values(metric)
    fig = px.bar(chart, x=metric, y="Hotel", orientation="h", color_discrete_sequence=[color],
                 hover_data={"Kota": True, "Sales AR": ":,.0f", "Room Night": ":,.0f",
                             "Rata-rata / Room Night": ":,.0f"},
                 title=f"Top {top_n} Hotel berdasarkan {metric}",
                 height=max(380, 30 * len(chart) + 110))
    fig.update_layout(yaxis_title=None, margin=dict(t=50, b=20))
    g1.plotly_chart(fig, width="stretch", key=f"bar_{key}")

    # Rata-rata harga per room night untuk hotel yang sama (urutan sama dengan grafik kiri)
    adr = chart.assign(ADR=chart["Rata-rata / Room Night"].fillna(0))
    f_adr = px.bar(adr, x="ADR", y="Hotel", orientation="h", color_discrete_sequence=["#F18F01"],
                   title="Rata-rata Sales per Room Night", text_auto=",.0f",
                   height=max(380, 30 * len(chart) + 110))
    f_adr.update_layout(yaxis_title=None, xaxis_title="Rp / room night",
                        yaxis=dict(showticklabels=False), margin=dict(t=50, b=20))
    g2.plotly_chart(f_adr, width="stretch", key=f"adr_{key}")

    # Tren bulanan Top 5
    if "Bulan" in fdata.columns:
        top5 = top["Hotel"].head(5).tolist()
        src = "Sales AR" if metric == "Sales AR" else "Room Night"
        tr = (fdata[fdata["Hotel (Grup)"].isin(top5)]
              .groupby(["Bulan", "Hotel (Grup)"])[src].sum().reset_index())
        f2 = px.line(tr, x="Bulan", y=src, color="Hotel (Grup)", markers=True,
                     title=f"Tren Bulanan {metric} — Top 5 Hotel", height=400)
        f2.update_layout(legend=dict(orientation="h", y=-0.25, title=None),
                         xaxis_title=None, margin=dict(t=50))
        st.plotly_chart(f2, width="stretch", key=f"trend_{key}")

    st.markdown(f"**Peringkat lengkap ({angka(len(r))} hotel)**")
    st.dataframe(
        r[["Rank", "Hotel", "Kota", metric, share_col] +
          [x for x in ["Sales AR", "Room Night", "Invoice (unik)", "Rata-rata / Room Night",
                       "Profit", "Margin %"] if x != metric]],
        hide_index=True, width="stretch", height=420,
        column_config={
            "Sales AR": st.column_config.NumberColumn(format="localized"),
            "Profit": st.column_config.NumberColumn(format="localized"),
            "Room Night": st.column_config.NumberColumn(format="localized"),
            "Rata-rata / Room Night": st.column_config.NumberColumn(format="localized"),
            "Margin %": st.column_config.NumberColumn(format="%.2f%%"),
            share_col: st.column_config.ProgressColumn(
                "Share %", format="%.2f%%", min_value=0,
                max_value=float(r[share_col].max()) if len(r) else 100.0),
        })
    st.download_button(f"⬇️ Unduh peringkat {metric} (CSV)",
                       r.to_csv(index=False).encode("utf-8-sig"),
                       file_name=f"top_hotel_{key}.csv", mime="text/csv", key=f"dl_{key}")


with tab_top:
    if not {"Hotel (Grup)", "Sales AR", "Room Night"} <= set(fdata.columns):
        st.info("Kolom hotel / Sales AR / Room Night tidak tersedia pada data.")
    else:
        st.caption("Peringkat memakai nama hotel yang sudah digabung (tab Grup Hotel) dan "
                   "mengikuti semua filter. Room Night = Room × Night. "
                   "Jumlah Top N diatur dari slider di sidebar.")
        rank = hotel_ranking(fdata[["Hotel (Grup)", "Hotel City", "Sales AR", "Profit",
                                    "Room Night", "Invoice No"]])
        sub1, sub2 = st.tabs(["💰 Top Spender Hotel (Sales AR)", "🛏️ Top Room Night"])
        with sub1:
            ranking_section(rank, "Sales AR", "Share Sales %", "#2E86AB", rupiah_full, "spender")
        with sub2:
            ranking_section(rank, "Room Night", "Share RN %", "#3B8B5A", angka, "roomnight")

SCORECARD_CSS = """<style>
.sc-wrap {display:grid; grid-template-columns:repeat(auto-fit,minmax(220px,1fr)); gap:14px; margin:6px 0 18px;}
.sc {border-radius:14px; padding:16px 18px; color:#fff; box-shadow:0 2px 8px rgba(0,0,0,.12);}
.sc .lbl {font-size:.85rem; opacity:.9; text-transform:uppercase; letter-spacing:.04em;}
.sc .val {font-size:clamp(1.35rem,2.3vw,2rem); font-weight:700; margin:4px 0 2px; word-break:break-word;}
.sc .sub {font-size:.8rem; opacity:.85;}
.sc.rev {background:linear-gradient(135deg,#1f6f9f,#2E86AB);}
.sc.rn  {background:linear-gradient(135deg,#2a6b45,#3B8B5A);}
.sc.bk  {background:linear-gradient(135deg,#a8580a,#F18F01);}
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
    if hcol not in fdata.columns:
        st.info("Kolom Hotel Name tidak tersedia pada data.")
    else:
        st.caption("Pilih satu atau beberapa hotel (nama sudah distandarkan). Data tetap mengikuti "
                   "filter di sidebar & Filter Dashboard. Kosongkan pilihan untuk melihat semua hotel.")
        h_opts = (fdata.groupby(hcol)["Sales AR"].sum().sort_values(ascending=False).index
                  .astype(str).tolist())
        f1, f2 = st.columns([3, 1])
        urut = f2.radio("Urutan pilihan", ["Revenue terbesar", "A–Z"], key="find_sort")
        if urut == "A–Z":
            h_opts = sorted(h_opts)
        sel_h = f1.multiselect("Nama hotel", h_opts, key="find_hotel",
                               placeholder="Ketik nama hotel... (kosong = semua hotel)")
        hd = fdata[fdata[hcol].astype(str).isin(sel_h)] if sel_h else fdata

        if hd.empty:
            st.warning("Tidak ada transaksi untuk hotel yang dipilih.")
        else:
            inv = invoice_key(hd["Invoice No"]) if "Invoice No" in hd else pd.Series(hd.index)
            scorecards(col_sum(hd, "Sales AR"), col_sum(hd, "Room Night"),
                       int(inv.nunique()), len(hd))

            # Info hotel (kota, chain, rating) bila hanya 1 hotel dipilih
            if len(sel_h) == 1:
                info = []
                for c, lbl in [("Hotel City", "Kota"), ("Hotel Chain", "Chain"),
                               ("Hotel Rating", "Rating")]:
                    if c in hd.columns and hd[c].notna().any():
                        v = hd[c].dropna().astype(str).mode().iat[0]
                        info.append(f"**{lbl}:** {'⭐' * int(float(v)) if c == 'Hotel Rating' else v}")
                if "Hotel Name (Asli)" in hd.columns:
                    var = sorted(hd["Hotel Name (Asli)"].dropna().astype(str).unique())
                    info.append(f"**Penulisan asli ({len(var)}):** " + "; ".join(var[:10])
                                + (" …" if len(var) > 10 else ""))
                st.markdown(" · ".join(info))

            # Tren bulanan
            if "Bulan" in hd.columns:
                mm = (hd.assign(_inv=inv.values).groupby("Bulan")
                      .agg(Revenue=("Sales AR", "sum"), RoomNight=("Room Night", "sum"),
                           Booking=("_inv", "nunique")).reset_index())
                mm["Label"] = (mm["Bulan"].dt.month.map(lambda x: BULAN[x - 1]) + " "
                               + mm["Bulan"].dt.strftime("%y"))
                fg = go.Figure()
                fg.add_bar(x=mm["Label"], y=mm["Revenue"], name="Revenue", marker_color="#2E86AB")
                fg.add_scatter(x=mm["Label"], y=mm["RoomNight"], name="Room Night", yaxis="y2",
                               mode="lines+markers", line=dict(color="#3B8B5A", width=3))
                fg.add_scatter(x=mm["Label"], y=mm["Booking"], name="Booking", yaxis="y2",
                               mode="lines+markers", line=dict(color="#F18F01", width=2, dash="dot"))
                fg.update_layout(title="Tren Bulanan: Revenue, Room Night & Booking", height=420,
                                 yaxis=dict(title="Revenue (Rp)"),
                                 yaxis2=dict(title="Room Night / Booking", overlaying="y",
                                             side="right"),
                                 legend=dict(orientation="h", y=1.12), margin=dict(t=80, b=30))
                st.plotly_chart(fg, width="stretch", key="find_trend")

            # Ringkasan per hotel
            g = hd.assign(_inv=inv.values).groupby(hcol)
            per_h = g.agg(Revenue=("Sales AR", "sum"), **{"Room Night": ("Room Night", "sum"),
                                                          "Booking": ("_inv", "nunique")})
            for c, lbl in [("Hotel City", "Kota"), ("Hotel Chain", "Chain"),
                           ("Hotel Rating", "Rating")]:
                if c in hd.columns:
                    per_h[lbl] = g[c].agg(lambda s: s.dropna().mode().iat[0]
                                          if s.notna().any() else None)
            per_h = per_h[[c for c in ["Kota", "Chain", "Rating"] if c in per_h.columns]
                          + ["Revenue", "Room Night", "Booking"]]
            per_h["Rata-rata / Room Night"] = per_h["Revenue"] / per_h["Room Night"].where(
                per_h["Room Night"] > 0)
            per_h = (per_h.sort_values("Revenue", ascending=False).reset_index()
                     .rename(columns={hcol: "Hotel"}))
            st.markdown(f"**Ringkasan per hotel ({angka(len(per_h))} hotel)**")
            st.dataframe(per_h, hide_index=True, width="stretch", height=min(420, 38 + 35 * len(per_h)),
                         column_config={
                             "Revenue": st.column_config.NumberColumn(format="localized"),
                             "Room Night": st.column_config.NumberColumn(format="localized"),
                             "Booking": st.column_config.NumberColumn(format="localized"),
                             "Rata-rata / Room Night": st.column_config.NumberColumn(
                                 format="localized"),
                             "Rating": st.column_config.NumberColumn(format="%d ⭐"),
                         })

            with st.expander(f"📄 Detail transaksi ({angka(len(hd))} baris)"):
                st.dataframe(hd.head(5000), width="stretch", height=380)
            d1, d2 = st.columns(2)
            d1.download_button("⬇️ Unduh ringkasan per hotel (CSV)",
                               per_h.to_csv(index=False).encode("utf-8-sig"),
                               file_name="ringkasan_per_hotel.csv", mime="text/csv",
                               key="dl_find_sum")
            d2.download_button("⬇️ Unduh detail transaksi (CSV)",
                               hd.to_csv(index=False).encode("utf-8-sig"),
                               file_name="detail_transaksi_hotel.csv", mime="text/csv",
                               key="dl_find_det")

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
            st.download_button("⬇️ Unduh daftar hotel belum standar (CSV)",
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
