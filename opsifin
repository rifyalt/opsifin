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
import time
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from hotel_grouping import DEFAULT_STOPWORDS, build_hotel_groups

# ---------------------------------------------------------------------------
# Konfigurasi
# ---------------------------------------------------------------------------
st.set_page_config(page_title="Merge Excel & Dashboard", page_icon="📊", layout="wide")

EXCEL_EXT = (".xlsx", ".xlsm", ".xls")

DATE_COLS = ["Inv Date", "Due Date", "Issued Date", "Check In", "Check Out",
             "Departure Date", "Arrival Date"]

NUMERIC_COLS = ["Room", "Night", "Base Fare", "Fare Tax", "IWJR", "Add Charge",
                "Insurance", "PSC", "Other Charge", "Incentive", "Agent Comm", "NTA",
                "Travel Services", "Sales Net", "VAT", "Stamp Fee", "MDR", "Sales AR",
                "Extra Disc", "Profit", "Profit %", "Rounding", "Base Sell"]

BULAN = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"]


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
def to_excel_bytes(df: pd.DataFrame, summary: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter",
                        engine_kwargs={"options": {"constant_memory": True,
                                                   "strings_to_urls": False}}) as w:
        df.to_excel(w, sheet_name="Merged Data", index=False)
        summary.to_excel(w, sheet_name="Ringkasan File", index=False)
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
                       and not f.name.startswith("~$"))
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
fdata = multi_filter(fdata, "Supplier Name", "Supplier Name", fc1, "Semua supplier")
fdata = multi_filter(fdata, "Hotel (Grup)", "Hotel (sudah digabung)" if hotel_on else "Hotel",
                     fc2, "Semua hotel")

if fdata.empty:
    st.warning("Tidak ada data yang cocok dengan filter.")
    st.stop()

# ---------------------------------------------------------------------------
# Tabs
# ---------------------------------------------------------------------------
tab_dash, tab_top, tab_data, tab_hotel, tab_file = st.tabs(
    ["📈 Dashboard", "🏆 Top Hotel", "🗂️ Data Gabungan", "🏨 Grup Hotel", "📁 Ringkasan File"])


def col_sum(df, c):
    return df[c].sum() if c in df.columns else 0


with tab_dash:
    sales = col_sum(fdata, "Sales AR")
    profit = col_sum(fdata, "Profit")
    n_inv = fdata["Invoice No"].nunique() if "Invoice No" in fdata else len(fdata)
    rn = col_sum(fdata, "Room Night")
    k = st.columns(3) + st.columns(3)
    k[0].metric("Total Sales AR", rupiah(sales))
    k[1].metric("Total Profit", rupiah(profit))
    k[2].metric("Margin", f"{(profit / sales * 100 if sales else 0):.2f}%")
    k[3].metric("Jumlah Invoice", angka(n_inv))
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

    def top_bar(col, title, color):
        if col not in fdata.columns:
            return None
        t = (fdata[fdata[col].astype(str).str.strip().ne("-")]
             .groupby(col)["Sales AR"].sum().nlargest(top_n).sort_values().reset_index())
        f = px.bar(t, x="Sales AR", y=col, orientation="h", title=title,
                   color_discrete_sequence=[color], height=max(350, 28 * len(t) + 100))
        f.update_layout(yaxis_title=None, margin=dict(t=50, b=20))
        return f

    c1, c2 = st.columns(2)
    for col_box, (col, title, color) in zip(
        [c1, c2, c1, c2],
        [("Customer Name", f"Top {top_n} Customer (Sales AR)", "#2E86AB"),
         ("Hotel (Grup)", f"Top {top_n} Hotel (Sales AR)", "#A23B72"),
         ("Hotel City", f"Top {top_n} Kota (Sales AR)", "#3B8B5A"),
         ("Supplier Name", f"Top {top_n} Supplier (Sales AR)", "#C73E1D")]):
        fig = top_bar(col, title, color)
        if fig is not None:
            col_box.plotly_chart(fig, width="stretch")

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
    agg = d.groupby(hcol).agg(
        **{"Sales AR": ("Sales AR", "sum"), "Profit": ("Profit", "sum"),
           "Room Night": ("Room Night", "sum"), "Invoice": ("Invoice No", "nunique")})
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
          [x for x in ["Sales AR", "Room Night", "Invoice", "Rata-rata / Room Night",
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
            ranking_section(rank, "Sales AR", "Share Sales %", "#2E86AB", rupiah, "spender")
        with sub2:
            ranking_section(rank, "Room Night", "Share RN %", "#3B8B5A", angka, "roomnight")

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
                payload = to_excel_bytes(export_df, summary)
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
    per_file = (data.groupby("Source File")
                .agg(Baris=("Source File", "size"),
                     Sales_AR=("Sales AR", "sum"), Profit=("Profit", "sum"))
                .reset_index())
    st.dataframe(summary.merge(per_file, left_on="File", right_on="Source File", how="left")
                 .drop(columns=["Source File", "Baris"]),
                 width="stretch", hide_index=True)
    st.caption("Kolom 'Periode (dari judul)' diambil dari baris judul di dalam file, "
               "berguna untuk mengecek apakah nama file sesuai isinya.")
