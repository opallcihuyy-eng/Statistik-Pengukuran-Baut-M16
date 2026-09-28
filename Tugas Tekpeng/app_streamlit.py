"""
==================================================================================
 STATISTIK PENGUKURAN REPLIKASI -- VERNIER CALIPER
 Tampilan: lembar kerja teknik (blueprint / drafting sheet), tiap orang berwarna
==================================================================================
Cara jalankan di laptop:
    1. pip install -r requirements.txt
    2. streamlit run app_streamlit.py

App ini HANYA alat bantu HITUNG. Pengukuran fisik dengan vernier caliper
dilakukan dulu oleh 5 anggota kelompok (5 orang x 5 ulangan = 25 data).
Objek pengukuran: diameter Baut M16 (diameter nominal 16 mm).
==================================================================================
"""

import html
import os
import statistics
import tempfile

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from kalkulator_statistik_pengukuran import buat_workbook, hitung_statistik

st.set_page_config(
    page_title="Lembar Kerja Pengukuran -- Baut M16",
    page_icon="⌀",
    layout="wide",
    initial_sidebar_state="expanded",
)

INK = "#12293F"
PAPER = "#F4EFE4"
CARD = "#FFFDF8"
LINE = "#C7BFA9"
RUST = "#C2540A"
OK_GREEN = "#2E8B57"

# satu warna tetap untuk tiap orang -- dipakai di kartu input, tabel, dan grafik
PERSON_COLORS = ["#D9541E", "#1F6FB2", "#2E8B57", "#B03A6F", "#A67C00"]

# ----------------------------------------------------------------------------------
# CSS -- semua warna dipaksa eksplisit supaya aman di dark mode maupun light mode
# ----------------------------------------------------------------------------------
BASE_CSS = f"""
@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=IBM+Plex+Mono:wght@400;500;600&family=Inter:wght@400;500&display=swap');

html, body, [class*="css"] {{ font-family: 'Inter', sans-serif; }}

.stApp, [data-testid="stAppViewContainer"] {{
    background-color: {PAPER};
    background-image:
        linear-gradient(rgba(18,41,63,0.05) 1px, transparent 1px),
        linear-gradient(90deg, rgba(18,41,63,0.05) 1px, transparent 1px);
    background-size: 26px 26px;
}}
header[data-testid="stHeader"] {{
    background-color: rgba(244,239,228,0.95) !important;
    border-bottom: 1px solid {LINE};
}}
header[data-testid="stHeader"] * {{ color: {INK} !important; }}
[data-testid="stMainBlockContainer"], .block-container {{
    padding-top: 3.4rem !important;
    max-width: 1320px;
}}

/* ---------- teks utama ---------- */
[data-testid="stMarkdownContainer"] p,
[data-testid="stMarkdownContainer"] li,
[data-testid="stWidgetLabel"] p {{ color: {INK}; }}
[data-testid="stCaptionContainer"], [data-testid="stCaptionContainer"] p {{
    color: #4F5B66 !important;
}}

/* ---------- input (number, text, textarea) ---------- */
div[data-baseweb="input"], div[data-baseweb="base-input"], div[data-baseweb="textarea"] {{
    background-color: {CARD} !important;
    border-radius: 2px !important;
}}
div[data-baseweb="input"] input, div[data-baseweb="base-input"] input, textarea {{
    color: {INK} !important;
    -webkit-text-fill-color: {INK} !important;
    background-color: {CARD} !important;
    font-family: 'IBM Plex Mono', monospace !important;
}}
[data-testid="stNumberInputContainer"] button {{ background-color: #EFE8D6 !important; }}
[data-testid="stNumberInputContainer"] button * {{ color: {INK} !important; fill: {INK} !important; }}

/* ---------- sidebar ---------- */
section[data-testid="stSidebar"] {{
    background-color: {INK};
    border-right: 3px solid {RUST};
}}
[data-testid="stSidebarHeader"] {{
    padding: 0.5rem 1rem 0 1rem !important;
    height: auto !important;
    min-height: 0 !important;
    margin: 0 !important;
}}
[data-testid="stSidebarUserContent"] {{ padding-top: 0.3rem !important; }}
section[data-testid="stSidebar"] * {{ color: #E7E1D2 !important; }}
section[data-testid="stSidebar"] [data-testid="stCaptionContainer"] p {{ color: #AEBBC7 !important; }}
section[data-testid="stSidebar"] hr {{ border-color: rgba(231,225,210,0.25); }}
section[data-testid="stSidebar"] div[data-baseweb="input"] input,
section[data-testid="stSidebar"] textarea {{
    color: {INK} !important; -webkit-text-fill-color: {INK} !important;
}}
section[data-testid="stSidebar"] [data-testid="stNumberInputContainer"] button * {{
    color: {INK} !important; fill: {INK} !important;
}}
.side-title {{
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.78rem; letter-spacing: 0.14em; text-transform: uppercase;
    color: {RUST}; border-bottom: 1px solid rgba(231,225,210,0.3);
    padding-bottom: 6px; margin-bottom: 4px;
}}

/* ---------- kop lembar (title block) ---------- */
.title-block {{
    border: 1.5px solid {INK}; background: {CARD};
    display: grid; grid-template-columns: 2.4fr 1fr 1fr;
    margin-bottom: 1.6rem;
}}
.title-block .main {{ grid-row: 1 / span 2; padding: 20px 24px; }}
.title-block .main h1 {{
    font-family: 'Space Grotesk', sans-serif; font-size: 1.6rem;
    margin: 0 0 6px 0; padding: 0; color: {INK}; line-height: 1.2;
}}
.title-block .main p {{ margin: 0; color: #4F5B66; font-size: 0.88rem; line-height: 1.5; }}
.title-block .field {{
    padding: 12px 18px; border-left: 1.5px solid {INK};
    display: flex; flex-direction: column; justify-content: center;
}}
.title-block .field:nth-child(4), .title-block .field:nth-child(5) {{ border-top: 1.5px solid {INK}; }}
.title-block .field span.k {{
    font-family: 'IBM Plex Mono', monospace; font-size: 0.62rem;
    letter-spacing: 0.12em; color: #8A6D3B; text-transform: uppercase;
}}
.title-block .field span.v {{
    font-family: 'IBM Plex Mono', monospace; font-size: 0.95rem;
    color: {INK}; font-weight: 600;
}}

/* ---------- judul bagian ---------- */
.sheet-heading {{
    display: flex; align-items: baseline; gap: 12px;
    margin: 0.6rem 0 0.9rem 0; border-bottom: 1.5px solid {INK}; padding-bottom: 6px;
}}
.sheet-heading .num {{ font-family: 'IBM Plex Mono', monospace; font-size: 1.6rem; font-weight: 600; color: {RUST}; }}
.sheet-heading .txt {{
    font-family: 'Space Grotesk', sans-serif; font-size: 1.05rem;
    letter-spacing: 0.04em; text-transform: uppercase; color: {INK};
}}
.sheet-heading .sub {{ margin-left: auto; font-family: 'IBM Plex Mono', monospace; font-size: 0.75rem; color: #6E6858; }}

/* ---------- kartu ringkasan ---------- */
.plate {{ border: 1.3px solid {INK}; background: {CARD}; padding: 14px 16px 12px 16px; height: 100%; }}
.plate .k {{
    font-family: 'IBM Plex Mono', monospace; font-size: 0.68rem; letter-spacing: 0.09em;
    text-transform: uppercase; color: #8A6D3B; margin-bottom: 6px; display: block;
}}
.plate .v {{ font-family: 'IBM Plex Mono', monospace; font-size: 1.7rem; font-weight: 600; color: {INK}; line-height: 1.1; }}
.plate .u {{ font-size: 0.85rem; color: #6E6858; margin-left: 4px; font-weight: 400; }}
.plate.accent {{ border-color: {RUST}; }}
.plate.accent .v {{ color: {RUST}; }}

/* ---------- kartu per orang ---------- */
.person-head {{
    font-family: 'Space Grotesk', sans-serif; font-weight: 700; font-size: 1rem;
    color: #FFFFFF; padding: 7px 11px; border-radius: 2px; margin-bottom: 6px;
    display: flex; justify-content: space-between; align-items: center;
}}
.person-head small {{ font-family: 'IBM Plex Mono', monospace; font-weight: 500; font-size: 0.68rem; opacity: 0.9; letter-spacing: 0.08em; }}
.person-foot {{
    font-family: 'IBM Plex Mono', monospace; font-size: 0.78rem; font-weight: 600;
    border-top: 1.5px dashed currentColor; padding-top: 8px; margin-top: 4px; text-align: center;
}}

/* ---------- tombol ---------- */
.stButton > button, .stDownloadButton > button {{
    background-color: {INK}; border: 1.3px solid {INK}; border-radius: 2px;
    padding: 0.55rem 1.3rem;
}}
.stButton > button *, .stDownloadButton > button * {{
    color: {PAPER} !important; font-family: 'IBM Plex Mono', monospace;
    letter-spacing: 0.06em; text-transform: uppercase; font-size: 0.82rem;
}}
.stButton > button:hover, .stDownloadButton > button:hover {{ background-color: {RUST}; border-color: {RUST}; }}

/* ---------- upload file ---------- */
[data-testid="stFileUploaderDropzone"] {{
    background-color: {CARD} !important; border: 1.5px dashed {INK} !important; border-radius: 2px !important;
}}
[data-testid="stFileUploaderDropzone"] * {{ color: {INK} !important; }}
[data-testid="stFileUploaderDropzone"] button {{ background-color: {INK} !important; border-radius: 2px !important; }}
[data-testid="stFileUploaderDropzone"] button * {{ color: {PAPER} !important; }}

/* ---------- tab ---------- */
.stTabs [data-baseweb="tab-list"] {{ gap: 4px; }}
.stTabs [data-baseweb="tab"] {{ background: transparent; border-radius: 0; padding: 8px 16px; }}
.stTabs [data-baseweb="tab"] p {{
    font-family: 'IBM Plex Mono', monospace; font-size: 0.82rem; text-transform: uppercase;
    letter-spacing: 0.05em; color: #4F5B66 !important;
}}
.stTabs [aria-selected="true"] p {{ color: {RUST} !important; font-weight: 600; }}
.stTabs [data-baseweb="tab-highlight"] {{ background-color: {RUST} !important; height: 3px !important; }}
.stTabs [data-baseweb="tab-border"] {{ background-color: {INK} !important; }}

/* ---------- notifikasi ---------- */
[data-testid="stAlert"], [data-testid="stAlertContainer"] {{
    background-color: {CARD} !important; border: 1.3px solid {INK}; border-radius: 2px;
}}
[data-testid="stAlert"] * {{ color: {INK} !important; }}

/* ---------- tabel HTML ---------- */
.tbl-scroll {{ max-height: 440px; overflow: auto; border: 1.3px solid {INK}; background: {CARD}; }}
.sheet-table {{ width: 100%; border-collapse: collapse; font-family: 'IBM Plex Mono', monospace; font-size: 0.82rem; }}
.sheet-table th {{
    position: sticky; top: 0; background: {INK}; color: {PAPER}; text-align: left; padding: 8px 10px;
    font-weight: 500; letter-spacing: 0.05em; text-transform: uppercase; font-size: 0.68rem;
}}
.sheet-table td {{ padding: 7px 10px; border-bottom: 1px solid #E4DCC8; color: {INK}; }}
.sheet-table tr:nth-child(even) td {{ background: #F7F2E6; }}
.sheet-table td.num {{ text-align: right; font-weight: 600; }}
.sheet-table .dot {{ display: inline-block; width: 10px; height: 10px; border-radius: 50%; margin-right: 8px; }}

.footnote {{
    font-family: 'IBM Plex Mono', monospace; font-size: 0.72rem; color: #6E6858;
    border-top: 1px solid {LINE}; padding-top: 8px; margin-top: 1.5rem;
}}
"""

# CSS tiap kolom orang -- warna berbeda per indeks
PERSON_CSS = ""
for _i, _c in enumerate(PERSON_COLORS):
    PERSON_CSS += f"""
.st-key-kolom_{_i} {{
    background-color: {_c}1A; border: 1.5px solid {_c}; border-top: 6px solid {_c};
    border-radius: 2px; padding: 0.8rem 0.8rem 0.6rem 0.8rem; gap: 0.35rem;
}}
.st-key-kolom_{_i} [data-testid="stWidgetLabel"] p {{
    color: {_c} !important; font-family: 'IBM Plex Mono', monospace;
    font-weight: 600; font-size: 0.74rem; letter-spacing: 0.06em; text-transform: uppercase;
}}
.st-key-kolom_{_i} div[data-baseweb="input"] {{ border: 1.5px solid {_c} !important; }}
.st-key-kolom_{_i} input {{ font-weight: 600 !important; }}
.st-key-kolom_{_i} [data-testid="stNumberInputContainer"] button {{ background-color: {_c}26 !important; }}
"""

st.markdown(f"<style>{BASE_CSS}{PERSON_CSS}</style>", unsafe_allow_html=True)


# ----------------------------------------------------------------------------------
# HELPER
# ----------------------------------------------------------------------------------
def heading(num, teks, sub=""):
    st.markdown(
        f'<div class="sheet-heading"><span class="num">{num}</span>'
        f'<span class="txt">{teks}</span><span class="sub">{sub}</span></div>',
        unsafe_allow_html=True,
    )


def html_table(df, colors=None):
    th = "".join(f"<th>{html.escape(str(c))}</th>" for c in df.columns)
    body = []
    for r, (_, row) in enumerate(df.iterrows()):
        tds = []
        for k, v in enumerate(row.tolist()):
            txt = html.escape(str(v))
            if k == 0:
                dot = f'<span class="dot" style="background:{colors[r]}"></span>' if colors else ""
                tds.append(f"<td>{dot}{txt}</td>")
            else:
                tds.append(f'<td class="num">{txt}</td>')
        body.append("<tr>" + "".join(tds) + "</tr>")
    return (f'<div class="tbl-scroll"><table class="sheet-table"><thead><tr>{th}</tr></thead>'
            f'<tbody>{"".join(body)}</tbody></table></div>')


# ----------------------------------------------------------------------------------
# KOP LEMBAR
# ----------------------------------------------------------------------------------
st.markdown(
    '<div class="title-block">'
    '<div class="main"><h1>⌀ Lembar Kerja Statistik Pengukuran Replikasi</h1>'
    '<p>Alat bantu hitung untuk Teknik Pengukuran. Data fisik diukur dulu dengan vernier caliper, '
    'baru diketik di sini.</p></div>'
    '<div class="field"><span class="k">Objek</span><span class="v">Baut M16</span></div>'
    '<div class="field"><span class="k">Metode</span><span class="v">5 x 5 ulangan</span></div>'
    '<div class="field"><span class="k">Alat</span><span class="v">Vernier caliper</span></div>'
    '<div class="field"><span class="k">n data</span><span class="v">25</span></div>'
    '</div>',
    unsafe_allow_html=True,
)

# ----------------------------------------------------------------------------------
# SIDEBAR
# ----------------------------------------------------------------------------------
with st.sidebar:
    st.markdown('<div class="side-title">Pengaturan instrumen</div>', unsafe_allow_html=True)
    nama_orang = st.text_area(
        "Anggota kelompok (satu nama per baris)",
        value="Orang 1\nOrang 2\nOrang 3\nOrang 4\nOrang 5",
        height=125,
    )
    orang_list = [x.strip() for x in nama_orang.splitlines() if x.strip()]

    st.markdown("---")
    pakai_acuan = st.checkbox("Gunakan nilai acuan / referensi", value=True)
    nilai_acuan = None
    if pakai_acuan:
        nilai_acuan = st.number_input(
            "Nilai acuan (mm)", value=16.00, step=0.01, format="%.4f",
            help="Diameter nominal baut M16 = 16 mm.",
        )
    resolusi_alat = st.number_input(
        "Resolusi vernier caliper (mm)", value=0.02, step=0.01, format="%.2f",
        help="Skala terkecil alat -- umumnya 0.05 atau 0.02 mm.",
    )
    st.markdown("---")
    st.caption("Data hanya tersimpan sementara di sesi browser. Simpan hasilnya lewat tab Ekspor.")

if len(orang_list) != 5:
    st.error(f"Butuh tepat 5 nama orang di sidebar -- saat ini ada {len(orang_list)}.")
    st.stop()
if len(set(orang_list)) != 5:
    st.error("Nama anggota tidak boleh ada yang sama.")
    st.stop()

# ----------------------------------------------------------------------------------
# STATE INPUT (dikunci per posisi, jadi ganti nama tidak mereset angka)
# ----------------------------------------------------------------------------------
DATA_AWAL_M16 = [
    [15.98, 16.00, 15.98, 16.02, 16.00],
    [16.04, 16.02, 16.06, 16.04, 16.02],
    [15.92, 15.94, 15.92, 15.96, 15.94],
    [16.00, 15.98, 16.02, 16.00, 15.98],
    [15.96, 15.98, 15.96, 16.00, 15.98],
]
for i in range(5):
    for j in range(5):
        st.session_state.setdefault(f"in_{i}_{j}", float(DATA_AWAL_M16[i][j]))
st.session_state["orang_cache"] = orang_list


def _muat_file():
    """Callback uploader: isi 25 kotak input dari CSV/Excel (kolom Orang, Ulangan, Diameter)."""
    up = st.session_state.get("uploader_data")
    if up is None:
        return
    try:
        df = pd.read_csv(up) if up.name.lower().endswith(".csv") else pd.read_excel(up)
        df = df.rename(columns=lambda c: str(c).strip().capitalize())
        if not {"Orang", "Ulangan", "Diameter"}.issubset(df.columns):
            raise ValueError("kolom harus bernama Orang, Ulangan, Diameter")
        df["Orang"] = df["Orang"].astype(str).str.strip()
        nama_file = list(dict.fromkeys(df["Orang"]))
        aktif = st.session_state["orang_cache"]
        urutan = aktif if set(nama_file) == set(aktif) else nama_file
        if len(urutan) != 5:
            raise ValueError(f"file berisi {len(urutan)} orang, seharusnya 5")
        for i, nm in enumerate(urutan):
            vals = df[df["Orang"] == nm].sort_values("Ulangan")["Diameter"].astype(float).tolist()
            if len(vals) != 5:
                raise ValueError(f"'{nm}' punya {len(vals)} data, seharusnya 5")
            for j, v in enumerate(vals):
                st.session_state[f"in_{i}_{j}"] = v
        st.session_state["upload_msg"] = ("ok", "Data dari file berhasil dimuat ke 25 kotak input.")
    except Exception as e:
        st.session_state["upload_msg"] = ("err", f"Gagal membaca file: {e}")


# ----------------------------------------------------------------------------------
# TAB
# ----------------------------------------------------------------------------------
tab_data, tab_hasil, tab_grafik, tab_unduh = st.tabs(
    ["01 -- Data", "02 -- Hasil", "03 -- Grafik", "04 -- Ekspor"]
)

# ============================================================ TAB 01: DATA
with tab_data:
    heading("01", "Input Data Pengukuran", "5 orang × 5 ulangan · satuan mm")
    st.caption("Tiap orang punya kolom dengan warna sendiri. Ketik langsung diameter hasil ukur (mm).")

    kolom = st.columns(5, gap="small")
    data_per_orang = {}
    for i, (o, kol) in enumerate(zip(orang_list, kolom)):
        warna = PERSON_COLORS[i]
        with kol:
            with st.container(key=f"kolom_{i}"):
                st.markdown(
                    f'<div class="person-head" style="background:{warna}">'
                    f'{html.escape(o)}<small>P{i + 1}</small></div>',
                    unsafe_allow_html=True,
                )
                vals = [
                    st.number_input(f"Ulangan {j + 1}", key=f"in_{i}_{j}", min_value=0.0,
                                    step=0.01, format="%.2f")
                    for j in range(5)
                ]
                if all(v is not None for v in vals):
                    ringkas = f"x̄ {statistics.mean(vals):.4f} · s {statistics.stdev(vals):.4f}"
                else:
                    ringkas = "data belum lengkap"
                st.markdown(f'<div class="person-foot" style="color:{warna}">{ringkas}</div>',
                            unsafe_allow_html=True)
        data_per_orang[o] = vals

    semua_data = [v for o in orang_list for v in data_per_orang[o]]
    valid = all(v is not None for v in semua_data)

    st.write("")
    kiri, kanan = st.columns([1, 1.4], gap="large")
    with kiri:
        st.markdown(
            f'<div class="plate" style="margin-bottom:12px;"><span class="k">Cek cepat</span>'
            f'<span class="v" style="font-size:1.2rem;">{sum(v is not None for v in semua_data)}'
            f'<span class="u">/ 25 data terisi</span></span></div>',
            unsafe_allow_html=True,
        )
        hitung = st.button("HITUNG STATISTIK →", type="primary")
        if hitung:
            if valid:
                st.session_state["hasil"] = hitung_statistik(semua_data, nilai_acuan)
                st.session_state["data_final"] = data_per_orang
                st.session_state["semua_data"] = semua_data
                st.session_state["nilai_acuan_used"] = nilai_acuan
                st.success("Selesai. Lihat tab 02, 03, dan 04.")
            else:
                st.warning("Lengkapi seluruh 25 kotak sebelum menghitung.")
        elif "hasil" in st.session_state and st.session_state["semua_data"] != semua_data:
            st.warning("Data berubah sejak hitungan terakhir. Klik HITUNG STATISTIK lagi.")
    with kanan:
        st.caption("Atau muat 25 data sekaligus dari file (kolom: Orang, Ulangan, Diameter)")
        st.file_uploader("Upload file", type=["csv", "xlsx"], key="uploader_data",
                         on_change=_muat_file, label_visibility="collapsed")
        pesan = st.session_state.pop("upload_msg", None)
        if pesan:
            (st.success if pesan[0] == "ok" else st.error)(pesan[1])

ada_hasil = "hasil" in st.session_state
BELUM = "Belum ada hasil. Isi data di tab 01 lalu klik HITUNG STATISTIK."

if ada_hasil:
    hasil = st.session_state["hasil"]
    data_final = st.session_state["data_final"]
    semua_hit = st.session_state["semua_data"]
    acuan = st.session_state["nilai_acuan_used"]
    warna_orang = {o: PERSON_COLORS[i] for i, o in enumerate(data_final)}

# ============================================================ TAB 02: HASIL
with tab_hasil:
    heading("02", "Ringkasan Hasil", "n = 25")
    if not ada_hasil:
        st.info(BELUM)
    else:
        c1, c2, c3, c4 = st.columns(4)
        plates = [
            (c1, "Rata-rata", f"{hasil.rata_rata:.4f}", "mm", False),
            (c2, "Std. Deviasi", f"{hasil.stdev:.4f}", "mm", False),
            (c3, "Presisi", f"{hasil.presisi_persen:.2f}", "%", True),
            (c4, "Akurasi",
             f"{hasil.akurasi_persen:.2f}" if hasil.akurasi_persen is not None else "n/a",
             "%" if hasil.akurasi_persen is not None else "", True),
        ]
        for kol, label, val, unit, aksen in plates:
            with kol:
                st.markdown(
                    f'<div class="{"plate accent" if aksen else "plate"}"><span class="k">{label}</span>'
                    f'<span class="v">{val}<span class="u">{unit}</span></span></div>',
                    unsafe_allow_html=True,
                )

        st.write("")
        colA, colB = st.columns([1.2, 1], gap="large")
        with colA:
            st.markdown("**Besaran statistik lengkap**")
            tabel = {
                "Jumlah data (n)": hasil.n,
                "Rata-rata (mm)": round(hasil.rata_rata, 4),
                "Standar Deviasi (mm)": round(hasil.stdev, 4),
                "Varians": round(hasil.varians, 6),
                "Standard Error of Mean (mm)": round(hasil.sem, 4),
                "Median (mm)": round(hasil.median, 4),
                "Modus (mm)": round(hasil.modus, 4),
                "Minimum (mm)": hasil.minimum,
                "Maksimum (mm)": hasil.maksimum,
                "Rentang / Range (mm)": round(hasil.rentang, 4),
                "Koefisien Variasi (%)": round(hasil.cv_persen, 2),
                "Presisi (%)": round(hasil.presisi_persen, 2),
                "Derajat kebebasan (df)": hasil.df,
                "Nilai t (95% CI)": hasil.t95,
                "Batas bawah CI 95% (mm)": round(hasil.ci95_lower, 4),
                "Batas atas CI 95% (mm)": round(hasil.ci95_upper, 4),
                "Ketidakpastian Pengukuran, U (mm)": round(hasil.ketidakpastian, 4),
            }
            if hasil.error_mutlak is not None:
                tabel.update({
                    "Nilai Acuan (mm)": acuan,
                    "Error Mutlak (mm)": round(hasil.error_mutlak, 4),
                    "Error Relatif (%)": round(hasil.error_relatif_persen, 2),
                    "Akurasi (%)": round(hasil.akurasi_persen, 2),
                })
            st.markdown(html_table(pd.DataFrame(tabel.items(), columns=["Besaran", "Nilai"])),
                        unsafe_allow_html=True)
        with colB:
            st.markdown("**Statistik per orang (repeatability)**")
            rows = []
            for o, vals in data_final.items():
                rows.append({
                    "Orang": o,
                    "Rata-rata": f"{statistics.mean(vals):.4f}",
                    "SD": f"{statistics.stdev(vals):.4f}",
                    "Min": f"{min(vals):.2f}",
                    "Max": f"{max(vals):.2f}",
                    "Rentang": f"{max(vals) - min(vals):.2f}",
                })
            st.markdown(html_table(pd.DataFrame(rows), colors=list(warna_orang.values())),
                        unsafe_allow_html=True)

# ============================================================ TAB 03: GRAFIK
with tab_grafik:
    heading("03", "Visualisasi Data", "interaktif · arahkan kursor untuk detail")
    if not ada_hasil:
        st.info(BELUM)
    else:
        grid = "#E7E0CE"
        tata = dict(
            template="simple_white",
            font=dict(family="IBM Plex Mono, monospace", size=12, color=INK),
            title_font=dict(family="Space Grotesk, sans-serif", size=15, color=INK),
            paper_bgcolor=CARD, plot_bgcolor=CARD,
            margin=dict(l=10, r=10, t=48, b=10),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
        )

        def finish(fig, judul, xt, yt):
            fig.update_layout(title=judul, xaxis_title=xt, yaxis_title=yt, **tata)
            fig.update_xaxes(gridcolor=grid, linecolor=INK)
            fig.update_yaxes(gridcolor=grid, linecolor=INK)
            return fig

        g1, g2 = st.columns(2)
        with g1:
            f1 = go.Figure()
            for o, vals in data_final.items():
                f1.add_trace(go.Scatter(x=list(range(1, 6)), y=vals, mode="lines+markers", name=o,
                                        line=dict(color=warna_orang[o], width=2.5), marker=dict(size=8)))
            if acuan is not None:
                f1.add_hline(y=acuan, line_dash="dash", line_color=INK, annotation_text="acuan")
            st.plotly_chart(finish(f1, "Hasil Tiap Ulangan per Orang", "Ulangan ke-", "Diameter (mm)"), theme=None)
        with g2:
            nama = list(data_final)
            f2 = go.Figure(go.Bar(
                x=nama, y=[statistics.mean(data_final[o]) for o in nama],
                error_y=dict(type="data", array=[statistics.stdev(data_final[o]) for o in nama], color=INK),
                marker_color=[warna_orang[o] for o in nama]))
            if acuan is not None:
                f2.add_hline(y=acuan, line_dash="dash", line_color=INK, annotation_text="acuan")
            st.plotly_chart(finish(f2, "Rata-rata ± SD per Orang", "", "Diameter (mm)"), theme=None)

        g3, g4 = st.columns(2)
        with g3:
            f3 = go.Figure(go.Histogram(x=semua_hit, nbinsx=8, marker_color=INK, opacity=0.85))
            f3.add_vline(x=hasil.rata_rata, line_color=RUST, line_width=2, annotation_text=f"mean {hasil.rata_rata:.3f}")
            f3.add_vline(x=hasil.ci95_lower, line_dash="dash", line_color=OK_GREEN)
            f3.add_vline(x=hasil.ci95_upper, line_dash="dash", line_color=OK_GREEN, annotation_text="CI 95%")
            if acuan is not None:
                f3.add_vline(x=acuan, line_dash="dot", line_color="#B03A6F", annotation_text="acuan")
            st.plotly_chart(finish(f3, "Distribusi Seluruh 25 Data", "Diameter (mm)", "Frekuensi"), theme=None)
        with g4:
            f4 = go.Figure()
            f4.add_trace(go.Scatter(x=list(range(1, 26)), y=semua_hit, mode="lines",
                                    line=dict(color="#B9B2A0", width=1), showlegend=False, hoverinfo="skip"))
            idx = 1
            for o, vals in data_final.items():
                f4.add_trace(go.Scatter(x=list(range(idx, idx + 5)), y=vals, mode="markers", name=o,
                                        marker=dict(color=warna_orang[o], size=9)))
                idx += 5
            f4.add_hline(y=hasil.rata_rata, line_color=INK, line_width=1.5, annotation_text="CL")
            f4.add_hline(y=hasil.rata_rata + 3 * hasil.stdev, line_dash="dash", line_color=RUST, annotation_text="UCL")
            f4.add_hline(y=hasil.rata_rata - 3 * hasil.stdev, line_dash="dash", line_color=RUST, annotation_text="LCL")
            st.plotly_chart(finish(f4, "Control Chart · 25 Data Berurutan", "Urutan pengukuran", "Diameter (mm)"), theme=None)

        st.markdown(
            '<div class="footnote">01 repeatability per orang · 02 reproducibility antar orang · '
            '03 sebaran terhadap rata-rata &amp; acuan · 04 batas kontrol 3-sigma</div>',
            unsafe_allow_html=True,
        )

# ============================================================ TAB 04: EKSPOR
with tab_unduh:
    heading("04", "Ekspor Laporan", ".xlsx dengan rumus")
    if not ada_hasil:
        st.info(BELUM)
    else:
        st.write("Berkas Excel berisi rumus (bukan angka mati), lengkap dengan grafik siap pakai untuk laporan.")
        tmp_path = os.path.join(tempfile.gettempdir(), "hasil_statistik_pengukuran.xlsx")
        buat_workbook(data_final, acuan, resolusi_alat, output_path=tmp_path)
        with open(tmp_path, "rb") as f:
            st.download_button(
                "UNDUH HASIL_STATISTIK_PENGUKURAN.XLSX",
                data=f.read(),
                file_name="hasil_statistik_pengukuran.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        st.caption("Buka di Excel/LibreOffice. Tekan Ctrl+Shift+F9 untuk recalc jika ada sel kosong saat pertama dibuka.")
