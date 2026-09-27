"""
==================================================================================
 STATISTIK PENGUKURAN REPLIKASI -- VERNIER CALIPER
 Tampilan: gaya lembar kerja teknik (blueprint / drafting sheet)
==================================================================================
Cara jalankan di laptop:
    1. pip install -r requirements.txt
    2. streamlit run app_streamlit.py
    3. Browser otomatis terbuka di http://localhost:8501

App ini HANYA alat bantu HITUNG. Pengukuran fisik dengan vernier caliper
tetap dilakukan dulu di lab oleh 5 anggota kelompok (5 orang x 5 ulangan
= 25 data), baru angkanya dimasukkan di sini.

Objek pengukuran: diameter Baut M16 (diameter nominal 16 mm).
==================================================================================
"""

import io
import os
import tempfile
import statistics

import pandas as pd
import streamlit as st
import plotly.graph_objects as go

from kalkulator_statistik_pengukuran import hitung_statistik, buat_workbook

# ----------------------------------------------------------------------------------
# KONFIGURASI HALAMAN
# ----------------------------------------------------------------------------------
st.set_page_config(
    page_title="Lembar Kerja Pengukuran -- Baut M16",
    page_icon="⌀",
    layout="wide",
    initial_sidebar_state="expanded",
)

INK = "#12293F"        # navy tinta
PAPER = "#F4EFE4"      # kertas
LINE = "#C7BFA9"       # garis kisi
RUST = "#C2540A"       # amber/rust aksen
RUST_SOFT = "#E9C9A6"
OK_GREEN = "#3E6B4F"

PALETTE = ["#C2540A", "#12293F", "#3E6B4F", "#8A6D3B", "#5B7B9A"]

# ----------------------------------------------------------------------------------
# CSS -- lembar gambar teknik: kertas krem, kisi tipis, tipografi mono untuk angka
# ----------------------------------------------------------------------------------
st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=IBM+Plex+Mono:wght@400;500;600&family=Inter:wght@400;500&display=swap');

html, body, [class*="css"] {{
    font-family: 'Inter', sans-serif;
}}

.stApp {{
    background-color: {PAPER};
    background-image:
        linear-gradient(rgba(18,41,63,0.05) 1px, transparent 1px),
        linear-gradient(90deg, rgba(18,41,63,0.05) 1px, transparent 1px);
    background-size: 26px 26px;
}}

section[data-testid="stSidebar"] {{
    background-color: {INK};
    border-right: 2px solid {RUST};
}}
section[data-testid="stSidebar"] * {{
    color: #E7E1D2 !important;
}}
section[data-testid="stSidebar"] input, section[data-testid="stSidebar"] textarea {{
    color: {INK} !important;
}}
section[data-testid="stSidebar"] .stCheckbox label p {{ color:#E7E1D2 !important; }}

/* -- Title block, model kop gambar teknik -- */
.title-block {{
    border: 1.5px solid {INK};
    background: #FFFDF8;
    padding: 0;
    margin-bottom: 1.6rem;
    display: grid;
    grid-template-columns: 2.4fr 1fr 1fr;
}}
.title-block .main {{
    padding: 18px 22px;
    border-right: 1.5px solid {INK};
}}
.title-block .main h1 {{
    font-family: 'Space Grotesk', sans-serif;
    font-size: 1.55rem;
    margin: 0 0 4px 0;
    color: {INK};
    letter-spacing: 0.01em;
}}
.title-block .main p {{
    margin: 0;
    color: #55606B;
    font-size: 0.88rem;
}}
.title-block .field {{
    padding: 10px 16px;
    border-bottom: 1px solid {LINE};
    display: flex;
    flex-direction: column;
    justify-content: center;
}}
.title-block .field:nth-child(3) {{ border-right: 1.5px solid {INK}; }}
.title-block .field span.k {{
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.62rem;
    letter-spacing: 0.12em;
    color: #8A6D3B;
    text-transform: uppercase;
}}
.title-block .field span.v {{
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.92rem;
    color: {INK};
    font-weight: 600;
}}

/* -- Section heading bergaya nomor lembar -- */
.sheet-heading {{
    display: flex;
    align-items: baseline;
    gap: 12px;
    margin: 1.7rem 0 0.9rem 0;
    border-bottom: 1.5px solid {INK};
    padding-bottom: 6px;
}}
.sheet-heading .num {{
    font-family: 'IBM Plex Mono', monospace;
    font-size: 1.6rem;
    font-weight: 600;
    color: {RUST};
}}
.sheet-heading .txt {{
    font-family: 'Space Grotesk', sans-serif;
    font-size: 1.05rem;
    letter-spacing: 0.04em;
    text-transform: uppercase;
    color: {INK};
}}
.sheet-heading .sub {{
    margin-left: auto;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.75rem;
    color: #8A8272;
}}

/* -- kartu data / metric plate -- */
.plate {{
    border: 1.3px solid {INK};
    background: #FFFDF8;
    padding: 14px 16px 12px 16px;
    height: 100%;
}}
.plate .k {{
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.68rem;
    letter-spacing: 0.09em;
    text-transform: uppercase;
    color: #8A6D3B;
    margin-bottom: 6px;
    display: block;
}}
.plate .v {{
    font-family: 'IBM Plex Mono', monospace;
    font-size: 1.7rem;
    font-weight: 600;
    color: {INK};
    line-height: 1.1;
}}
.plate .u {{
    font-size: 0.85rem;
    color: #8A8272;
    margin-left: 4px;
    font-weight: 400;
}}
.plate.accent {{ border-color: {RUST}; }}
.plate.accent .v {{ color: {RUST}; }}

/* buttons: square, tegas, tidak bulat gembung */
.stButton > button, .stDownloadButton > button {{
    background-color: {INK};
    color: #F4EFE4;
    border: 1.3px solid {INK};
    border-radius: 2px;
    font-family: 'IBM Plex Mono', monospace;
    letter-spacing: 0.06em;
    text-transform: uppercase;
    font-size: 0.82rem;
    padding: 0.55rem 1.3rem;
}}
.stButton > button:hover, .stDownloadButton > button:hover {{
    background-color: {RUST};
    border-color: {RUST};
    color: #FFFDF8;
}}

/* tabs */
.stTabs [data-baseweb="tab-list"] {{
    gap: 4px;
    border-bottom: 1.3px solid {INK};
}}
.stTabs [data-baseweb="tab"] {{
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.82rem;
    text-transform: uppercase;
    letter-spacing: 0.05em;
    color: #55606B;
    background: transparent;
    border-radius: 0;
    padding: 8px 16px;
}}
.stTabs [aria-selected="true"] {{
    color: {RUST} !important;
    border-bottom: 2.5px solid {RUST} !important;
    background: transparent !important;
}}

.footnote {{
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.72rem;
    color: #8A8272;
    border-top: 1px solid {LINE};
    padding-top: 8px;
    margin-top: 2rem;
}}

hr {{ border-color: {LINE}; }}
</style>
""", unsafe_allow_html=True)

# ----------------------------------------------------------------------------------
# TITLE BLOCK
# ----------------------------------------------------------------------------------
st.markdown(f"""
<div class="title-block">
  <div class="main">
    <h1>⌀ Lembar Kerja Statistik Pengukuran Replikasi</h1>
    <p>Alat bantu hitung -- Teknik Pengukuran. Data fisik diukur dulu dengan vernier caliper, baru diketik di sini.</p>
  </div>
  <div class="field"><span class="k">Objek</span><span class="v">Baut M16</span></div>
  <div class="field"><span class="k">Metode</span><span class="v">5 x 5 ulangan</span></div>
  <div class="field"><span class="k">Alat</span><span class="v">Vernier caliper</span></div>
  <div class="field"><span class="k">n data</span><span class="v">25</span></div>
</div>
""", unsafe_allow_html=True)

# ----------------------------------------------------------------------------------
# SIDEBAR -- PENGATURAN INSTRUMEN
# ----------------------------------------------------------------------------------
with st.sidebar:
    st.markdown("##### PENGATURAN INSTRUMEN")
    st.caption("Konfigurasi kelompok & alat ukur")

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
    st.caption("Deploy gratis ke internet: lihat panduan yang menyertai file ini, atau tanya ke Claude.")

if len(orang_list) != 5:
    st.error(f"Butuh tepat 5 nama orang di sidebar -- saat ini ada {len(orang_list)}.")
    st.stop()

# ----------------------------------------------------------------------------------
# STATE: tabel data 5 ulangan x 5 orang
# ----------------------------------------------------------------------------------
DATA_AWAL_M16 = {
    "Orang 1": [15.98, 16.00, 15.98, 16.02, 16.00],
    "Orang 2": [16.04, 16.02, 16.06, 16.04, 16.02],
    "Orang 3": [15.92, 15.94, 15.92, 15.96, 15.94],
    "Orang 4": [16.00, 15.98, 16.02, 16.00, 15.98],
    "Orang 5": [15.96, 15.98, 15.96, 16.00, 15.98],
}


def _default_table(orang_list):
    kolom = {o: DATA_AWAL_M16.get(o, [16.00] * 5) for o in orang_list}
    df = pd.DataFrame(kolom, index=[f"Ulangan {i+1}" for i in range(5)])
    return df


if "tabel_pengukuran" not in st.session_state or st.session_state.get("orang_list_cache") != orang_list:
    st.session_state.tabel_pengukuran = _default_table(orang_list)
    st.session_state.orang_list_cache = orang_list
    st.session_state.pop("editor_pengukuran", None)

# ----------------------------------------------------------------------------------
# TABS
# ----------------------------------------------------------------------------------
tab_data, tab_hasil, tab_grafik, tab_unduh = st.tabs(["01 -- Data", "02 -- Hasil", "03 -- Grafik", "04 -- Ekspor"])

# ==================================================================== TAB 01: DATA
with tab_data:
    st.markdown("""
    <div class="sheet-heading">
        <span class="num">01</span><span class="txt">Input Data Pengukuran</span>
        <span class="sub">5 orang &times; 5 ulangan, satuan mm</span>
    </div>
    """, unsafe_allow_html=True)

    colL, colR = st.columns([2.3, 1])
    with colL:
        st.caption("Klik sel untuk mengubah nilai langsung -- seperti mengisi lembar data.")
        edited = st.data_editor(
            st.session_state.tabel_pengukuran,
            key="editor_pengukuran",
            use_container_width=True,
            column_config={
                o: st.column_config.NumberColumn(o, format="%.4f", step=0.01) for o in orang_list
            },
        )
        st.session_state.tabel_pengukuran = edited

    with colR:
        st.caption("Atau muat 25 data sekaligus dari file")
        up = st.file_uploader("Kolom wajib: Orang, Ulangan, Diameter", type=["csv", "xlsx"], label_visibility="collapsed")
        if up is not None:
            try:
                df_up = pd.read_csv(up) if up.name.endswith(".csv") else pd.read_excel(up)
                pivot = df_up.pivot(index="Ulangan", columns="Orang", values="Diameter")
                pivot = pivot.reindex(columns=orang_list)
                pivot.index = [f"Ulangan {i}" for i in pivot.index]
                if pivot.isnull().values.any() or pivot.shape != (5, 5):
                    st.error("File harus berisi tepat 25 baris (5 orang x 5 ulangan) dengan nama orang yang cocok.")
                else:
                    st.session_state.tabel_pengukuran = pivot
                    st.session_state.pop("editor_pengukuran", None)
                    st.success("Data dimuat. Tabel diperbarui.")
                    st.rerun()
            except Exception as e:
                st.error(f"Gagal membaca file: {e}")

        st.markdown(
            f"""<div class="plate" style="margin-top:14px;">
            <span class="k">Cek cepat</span>
            <span class="v" style="font-size:1.1rem;">{edited.shape[0]*edited.shape[1]} <span class="u">data terisi</span></span>
            </div>""",
            unsafe_allow_html=True,
        )

    data_per_orang = {o: edited[o].tolist() for o in orang_list}
    valid = all(len(v) == 5 and not any(pd.isna(v)) for v in data_per_orang.values())

    if not valid:
        st.warning("Lengkapi seluruh 25 sel sebelum menghitung.")
        st.stop()

    semua_data = [v for o in orang_list for v in data_per_orang[o]]
    st.button("HITUNG STATISTIK →", type="primary", key="btn_hitung")
    if st.session_state.get("btn_hitung"):
        st.session_state["hasil"] = hitung_statistik(semua_data, nilai_acuan)
        st.session_state["data_final"] = data_per_orang
        st.session_state["semua_data"] = semua_data
        st.session_state["nilai_acuan_used"] = nilai_acuan

if "hasil" not in st.session_state:
    st.stop()

hasil = st.session_state["hasil"]
data_final = st.session_state["data_final"]
semua_data = st.session_state["semua_data"]
nilai_acuan = st.session_state["nilai_acuan_used"]

# ==================================================================== TAB 02: HASIL
with tab_hasil:
    st.markdown("""
    <div class="sheet-heading">
        <span class="num">02</span><span class="txt">Ringkasan Hasil</span>
        <span class="sub">n = 25</span>
    </div>
    """, unsafe_allow_html=True)

    c1, c2, c3, c4 = st.columns(4)
    plates = [
        (c1, "Rata-rata", f"{hasil.rata_rata:.4f}", "mm", False),
        (c2, "Std. Deviasi", f"{hasil.stdev:.4f}", "mm", False),
        (c3, "Presisi", f"{hasil.presisi_persen:.2f}", "%", True),
        (c4, "Akurasi", f"{hasil.akurasi_persen:.2f}" if hasil.akurasi_persen is not None else "n/a",
         "%" if hasil.akurasi_persen is not None else "", True),
    ]
    for col, label, val, unit, accent in plates:
        with col:
            cls = "plate accent" if accent else "plate"
            st.markdown(
                f"""<div class="{cls}"><span class="k">{label}</span>
                <span class="v">{val}<span class="u">{unit}</span></span></div>""",
                unsafe_allow_html=True,
            )

    st.write("")
    colA, colB = st.columns([1.3, 1])

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
                "Nilai Acuan (mm)": nilai_acuan,
                "Error Mutlak (mm)": round(hasil.error_mutlak, 4),
                "Error Relatif (%)": round(hasil.error_relatif_persen, 2),
                "Akurasi (%)": round(hasil.akurasi_persen, 2),
            })
        df_tabel = pd.DataFrame(tabel.items(), columns=["Besaran", "Nilai"])
        st.dataframe(df_tabel, use_container_width=True, hide_index=True, height=420)

    with colB:
        st.markdown("**Statistik per orang (repeatability)**")
        rows = []
        for o in orang_list:
            vals = data_final[o]
            rows.append({
                "Orang": o,
                "Rata-rata": round(statistics.mean(vals), 4),
                "SD": round(statistics.stdev(vals), 4),
                "Min": min(vals),
                "Max": max(vals),
                "Rentang": round(max(vals) - min(vals), 4),
            })
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True, height=420)

# ==================================================================== TAB 03: GRAFIK
with tab_grafik:
    st.markdown("""
    <div class="sheet-heading">
        <span class="num">03</span><span class="txt">Visualisasi Data</span>
        <span class="sub">interaktif -- arahkan kursor untuk detail</span>
    </div>
    """, unsafe_allow_html=True)

    layout_common = dict(
        template="simple_white",
        font=dict(family="IBM Plex Mono, monospace", size=12, color=INK),
        paper_bgcolor="#FFFDF8",
        plot_bgcolor="#FFFDF8",
        margin=dict(l=10, r=10, t=42, b=10),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
    )

    g1, g2 = st.columns(2)

    # --- Grafik 1: tiap ulangan per orang ---
    with g1:
        fig1 = go.Figure()
        for i, o in enumerate(orang_list):
            fig1.add_trace(go.Scatter(
                x=list(range(1, 6)), y=data_final[o], mode="lines+markers",
                name=o, line=dict(color=PALETTE[i % len(PALETTE)], width=2),
                marker=dict(size=7),
            ))
        if nilai_acuan is not None:
            fig1.add_hline(y=nilai_acuan, line_dash="dash", line_color=RUST, annotation_text="acuan")
        fig1.update_layout(title="Hasil Tiap Ulangan per Orang", xaxis_title="Ulangan ke-",
                            yaxis_title="Diameter (mm)", **layout_common)
        st.plotly_chart(fig1, use_container_width=True)

    # --- Grafik 2: rata-rata +/- SD per orang ---
    with g2:
        means = [statistics.mean(data_final[o]) for o in orang_list]
        stds = [statistics.stdev(data_final[o]) for o in orang_list]
        fig2 = go.Figure(go.Bar(
            x=orang_list, y=means, error_y=dict(type="data", array=stds, color=INK),
            marker_color=[PALETTE[i % len(PALETTE)] for i in range(len(orang_list))],
        ))
        if nilai_acuan is not None:
            fig2.add_hline(y=nilai_acuan, line_dash="dash", line_color=RUST, annotation_text="acuan")
        fig2.update_layout(title="Rata-rata ± SD per Orang", yaxis_title="Diameter (mm)", **layout_common)
        st.plotly_chart(fig2, use_container_width=True)

    g3, g4 = st.columns(2)

    # --- Grafik 3: histogram ---
    with g3:
        fig3 = go.Figure(go.Histogram(x=semua_data, nbinsx=8, marker_color=PALETTE[1], opacity=0.85))
        fig3.add_vline(x=hasil.rata_rata, line_color=INK, line_width=2,
                        annotation_text=f"mean {hasil.rata_rata:.3f}")
        fig3.add_vline(x=hasil.ci95_lower, line_dash="dash", line_color=OK_GREEN)
        fig3.add_vline(x=hasil.ci95_upper, line_dash="dash", line_color=OK_GREEN, annotation_text="CI 95%")
        if nilai_acuan is not None:
            fig3.add_vline(x=nilai_acuan, line_dash="dot", line_color=RUST, annotation_text="acuan")
        fig3.update_layout(title="Distribusi Seluruh 25 Data", xaxis_title="Diameter (mm)",
                            yaxis_title="Frekuensi", **layout_common)
        st.plotly_chart(fig3, use_container_width=True)

    # --- Grafik 4: control chart ---
    with g4:
        fig4 = go.Figure()
        idx = 1
        for i, o in enumerate(orang_list):
            xs = list(range(idx, idx + 5))
            fig4.add_trace(go.Scatter(x=xs, y=data_final[o], mode="markers", name=o,
                                       marker=dict(color=PALETTE[i % len(PALETTE)], size=8)))
            idx += 5
        fig4.add_trace(go.Scatter(x=list(range(1, 26)), y=semua_data, mode="lines",
                                   line=dict(color="#B9B2A0", width=1), showlegend=False))
        ucl = hasil.rata_rata + 3 * hasil.stdev
        lcl = hasil.rata_rata - 3 * hasil.stdev
        fig4.add_hline(y=hasil.rata_rata, line_color=INK, line_width=1.5, annotation_text="CL")
        fig4.add_hline(y=ucl, line_dash="dash", line_color=RUST, annotation_text="UCL")
        fig4.add_hline(y=lcl, line_dash="dash", line_color=RUST, annotation_text="LCL")
        fig4.update_layout(title="Control Chart -- 25 Data Berurutan", xaxis_title="Urutan pengukuran",
                            yaxis_title="Diameter (mm)", **layout_common)
        st.plotly_chart(fig4, use_container_width=True)

    st.markdown("""
    <div class="footnote">
    01 repeatability per orang &middot; 02 reproducibility antar orang &middot;
    03 sebaran terhadap rata-rata &amp; acuan &middot; 04 kontrol batas 3-sigma
    </div>
    """, unsafe_allow_html=True)

# ==================================================================== TAB 04: EKSPOR
with tab_unduh:
    st.markdown("""
    <div class="sheet-heading">
        <span class="num">04</span><span class="txt">Ekspor Laporan</span>
        <span class="sub">.xlsx dengan rumus</span>
    </div>
    """, unsafe_allow_html=True)

    st.write("Berkas Excel berisi rumus (bukan angka mati), lengkap dengan grafik siap pakai untuk laporan.")

    tmp_path = os.path.join(tempfile.gettempdir(), "hasil_statistik_pengukuran.xlsx")
    buat_workbook(data_final, nilai_acuan, resolusi_alat, output_path=tmp_path)
    with open(tmp_path, "rb") as f:
        st.download_button(
            "UNDUH HASIL_STATISTIK_PENGUKURAN.XLSX",
            data=f.read(),
            file_name="hasil_statistik_pengukuran.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    st.caption("Buka di Excel/LibreOffice -- tekan Ctrl+Shift+F9 untuk recalc jika ada sel kosong saat pertama dibuka.")
