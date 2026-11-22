from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

EXCEL_NAME = "Laporan Pencapaian Prestasi Fizikal Program CIDB Q1 2026 - DASHBOARD.xlsx"
SHEET_NAME = "DATA DASHBOARD OKTOBER CLEAN"
GREEN_OVERRIDES_OCT = {163, 174}

PERIODS = {
    "Suku Pertama": ("PERATUS SASARAN\n(WEIGHTAGE)\nQ1", "PERATUS PENCAPAIAN Q1", "% PENCAPAIAN (100%)", "CATATAN\n(JUSTIFIKASI)"),
    "Suku Kedua": ("SASARAN Q2\n(KUMULATIF)", "PENCAPAIAN Q2\n(KUMULATIF)", "% PENCAPAIAN (100%).0", "CATATAN\n(JUSTIFIKASI)"),
    "Julai": ("SASARAN JULAI\n(KUMULATIF)", "PENCAPAIAN JULAI\n(KUMULATIF)", "% PENCAPAIAN (100%).1", "CATATAN\n(JUSTIFIKASI)\nJULAI"),
    "Ogos": ("SASARAN OGOS\n(KUMULATIF)", "PENCAPAIAN OGOS\n(KUMULATIF)", "% PENCAPAIAN (100%).2", "CATATAN\n(JUSTIFIKASI)\nOGOS"),
    "Oktober": ("SASARAN OKTOBER\n(KUMULATIF)", "PENCAPAIAN OKTOBER\n(KUMULATIF)", "% PENCAPAIAN (100%).3", "CATATAN\n(JUSTIFIKASI)\nOKTOBER"),
}

# Sasaran kumulatif panel mengikut tempoh. Nilai ini bukan purata kolum
# sasaran setiap program; ia ialah sasaran masa tahunan dashboard asal.
PANEL_TARGETS = {
    "Suku Pertama": 25.00,
    "Suku Kedua": 50.00,
    "Julai": 58.33,
    "Ogos": 66.67,
    "Oktober": 83.33,
}

st.set_page_config(page_title="Dashboard Prestasi Program CIDB", page_icon="📊", layout="wide")
st.markdown("""
<style>
.stApp{background:radial-gradient(circle at 60% 18%,#fff 0,#f2f7fb 48%,#eaf1f8 100%);color:#12233c}
.main .block-container{padding:1rem 2rem 1.5rem;max-width:1480px}
header[data-testid="stHeader"]{background:transparent}
section[data-testid="stSidebar"]{background:linear-gradient(180deg,#263a56,#192b43);border-right:1px solid #7890aa}
section[data-testid="stSidebar"] *{color:#fff!important}
section[data-testid="stSidebar"] div[data-testid="stExpander"]{border:1px solid #d8e1ea;border-radius:20px;background:rgba(255,255,255,.035);margin-bottom:12px;overflow:hidden}
section[data-testid="stSidebar"] div[data-testid="stExpander"] details summary{font-weight:850;padding:12px 15px}
section[data-testid="stSidebar"] div[data-testid="stExpander"] div[role="radiogroup"]{display:flex;flex-direction:column;gap:8px;background:transparent;border:0;box-shadow:none;padding:2px 8px 12px}
section[data-testid="stSidebar"] div[data-testid="stExpander"] div[role="radiogroup"] label{border:1px solid #667b94;border-radius:999px;background:#30455f;padding:8px 12px;font-size:12px;font-weight:800;line-height:1.25;width:100%}
section[data-testid="stSidebar"] div[data-testid="stExpander"] div[role="radiogroup"] label:has(input:checked){background:#f7f9fc;border-color:#fff}
section[data-testid="stSidebar"] div[data-testid="stExpander"] div[role="radiogroup"] label:has(input:checked) *{color:#213752!important}
section[data-testid="stSidebar"] div[data-testid="stExpander"] div[role="radiogroup"] label>div:first-child{display:none}
section[data-testid="stSidebar"] hr{border-color:#60758d}
.title{font-size:36px;font-weight:900;line-height:1.22;color:#10243e;margin:12px 0 10px;letter-spacing:-.7px}
.download-note{font-size:12px;color:#5b6b7d;margin-top:-8px}
.circle-wrap{display:flex;justify-content:space-around;gap:34px;align-items:flex-start;padding:50px 8px 30px}
.circle-item{text-align:center;min-width:150px}.threshold{font-size:13px;font-weight:800;color:#111827;margin-bottom:18px}
.circle{width:126px;height:126px;border-radius:50%;display:flex;align-items:center;justify-content:center;margin:auto;color:#fff;font-size:36px;font-weight:900;border:3px solid rgba(255,255,255,.38)}
.green{background:linear-gradient(145deg,#6cf17e,#00a92e);box-shadow:0 0 28px 11px rgba(38,232,91,.35),inset 0 -8px 15px rgba(0,80,15,.25)}
.yellow{background:linear-gradient(145deg,#ffe96a,#d9ad00);box-shadow:0 0 28px 11px rgba(255,213,45,.34),inset 0 -8px 15px rgba(120,85,0,.24)}
.red{background:linear-gradient(145deg,#ff7b74,#c51f1f);box-shadow:0 0 28px 11px rgba(244,65,65,.32),inset 0 -8px 15px rgba(100,0,0,.25)}
.summary-box{padding:38px 8px 18px;margin-top:18px}.active-no{font-size:42px;font-weight:900;color:#1a2d48}.active-label{font-size:15px;font-weight:800;margin:4px 0 14px}
.line{border-top:1px solid #bdc9d6;margin:12px 0}.stats{font-size:16px;line-height:1.8}.achievement{color:#16b954;font-size:22px;font-weight:900;margin-top:20px}
.bottom-stat{text-align:center;font-size:19px;font-weight:800;color:#1d3657;margin:0 0 8px;word-spacing:5px}
.section-title{font-size:19px;font-weight:900;color:#17365d;margin:12px 0 8px}
div[data-testid="stDownloadButton"] button{border-radius:9px;border:1px solid #c9d5e3;background:#fff;color:#233b59;font-weight:700}
div[role="radiogroup"]{background:rgba(255,255,255,.78);border:1px solid #d6e0eb;border-radius:13px;padding:8px 14px;box-shadow:0 4px 14px rgba(35,58,86,.05)}
#MainMenu,footer{visibility:hidden}
</style>""", unsafe_allow_html=True)

@st.cache_data(show_spinner=False)
def load_raw():
    path = Path(__file__).resolve().parent / EXCEL_NAME
    if not path.exists():
        raise FileNotFoundError(f"Fail '{EXCEL_NAME}' tidak ditemui. Letakkan fail Excel dan program.py dalam folder yang sama.")
    df = pd.read_excel(path, sheet_name=SHEET_NAME, header=3)
    df = df[df["BIL."].notna()].copy()
    df["BIL."] = pd.to_numeric(df["BIL."], errors="coerce")
    df = df[df["BIL."].notna()].copy(); df["BIL."] = df["BIL."].astype(int)
    # Populasi tetap ialah 234 program aktif. Sebanyak 28 program yang telah
    # digugurkan dikecualikan berdasarkan status muktamad Oktober.
    oct_pct = pd.to_numeric(df["% PENCAPAIAN (100%).3"], errors="coerce")
    df = df[oct_pct.notna()].copy()
    return df

def period_view(raw, period):
    target, actual, pct, note = PERIODS[period]
    # Pandas may retain the second duplicate Q2 percentage header without a suffix.
    if period == "Suku Kedua" and pct not in raw.columns: pct = "% PENCAPAIAN (100%)"
    out = pd.DataFrame({
        "BIL.":raw["BIL."], "Kod Program":raw["KOD PROGRAM"], "Sektor":raw["SEKTOR"],
        "Bahagian":raw["BAHAGIAN"], "Deskripsi Program":raw["DESKRIPSI PROGRAM"],
        "Sasaran":pd.to_numeric(raw[target], errors="coerce"),
        "Pencapaian":pd.to_numeric(raw[actual], errors="coerce"),
        "% Pencapaian":pd.to_numeric(raw[pct], errors="coerce"),
        "Catatan/Justifikasi":raw[note], "Tambahan":raw["Bil. Program\nDiTambah"], "Gugur Asal":raw["Bil. Program\nGugur"],
    })
    def status(row):
        if period == "Ogos":
            if row["BIL."] in {99, 122}: return "Bermula Q3"
            if pd.isna(row["% Pencapaian"]) or row["BIL."] == 248: return "Bermula Q4"
        if period == "Oktober" and "BERMULA NOVEMBER" in str(row["Catatan/Justifikasi"]).upper():
            return "Bermula November"
        if period == "Oktober" and row["BIL."] in GREEN_OVERRIDES_OCT: return "Hijau"
        if pd.isna(row["% Pencapaian"]): return "Belum Bermula"
        if row["% Pencapaian"] >= 85: return "Hijau"
        if row["% Pencapaian"] >= 60: return "Kuning"
        return "Merah"
    out["Status"] = out.apply(status, axis=1)
    # Program ini kekal sebahagian daripada kategori Merah Oktober, tetapi
    # ditandakan berasingan supaya bilangannya jelas pada panel dashboard.
    out["Bermula November"] = out["Status"].eq("Bermula November")
    return out

try: raw = load_raw()
except Exception as exc: st.error(str(exc)); st.stop()

period = st.radio("Tempoh Pelaporan", list(PERIODS), index=4, horizontal=True, label_visibility="collapsed")
df = period_view(raw, period)

with st.sidebar:
    sectors = sorted(df["Sektor"].dropna().astype(str).unique())
    with st.expander("Pilih Sektor", expanded=True):
        sector_choice = st.radio("Sektor", ["Semua Sektor"] + sectors, label_visibility="collapsed", key=f"sector_{period}")
    base = df if sector_choice == "Semua Sektor" else df[df["Sektor"].astype(str) == sector_choice]
    divisions = sorted(base["Bahagian"].dropna().astype(str).unique())
    with st.expander("Pilih Bahagian", expanded=False):
        division_choice = st.radio("Bahagian", ["Semua Bahagian"] + divisions, label_visibility="collapsed", key=f"division_{period}_{sector_choice}")
    search = st.text_input("Carian", placeholder="Kod atau nama program")

filtered = df.copy()
if sector_choice != "Semua Sektor": filtered = filtered[filtered["Sektor"].astype(str) == sector_choice]
if division_choice != "Semua Bahagian": filtered = filtered[filtered["Bahagian"].astype(str) == division_choice]
if search.strip():
    key = search.strip().lower()
    filtered = filtered[filtered["Kod Program"].fillna("").astype(str).str.lower().str.contains(key,regex=False) | filtered["Deskripsi Program"].fillna("").astype(str).str.lower().str.contains(key,regex=False)]

st.markdown(f'<div class="title">LAPORAN PRESTASI PROGRAM SEHINGGA {period.upper()}<br>2026</div>', unsafe_allow_html=True)
st.download_button("📄 Muat Turun Data", filtered.to_csv(index=False).encode("utf-8-sig"), f"Prestasi_Program_{period}_2026.csv", "text/csv")

counts = df["Status"].value_counts(); active = len(df)
active_df = df[df["Status"].isin(["Hijau","Kuning","Merah"])]
target_avg = PANEL_TARGETS[period]
achievement = active_df["% Pencapaian"].dropna().mean() if not active_df.empty else 0
# Formula dashboard asal: PRESTASI = PENCAPAIAN × SASARAN KUMULATIF.
actual_avg = (achievement / 100) * target_avg

main, summary = st.columns([3.1,1.05], gap="large")
with main:
    st.markdown(f'''<div class="circle-wrap">
      <div class="circle-item"><div class="threshold">≥ 85%</div><div class="circle green">{counts.get('Hijau',0)}</div></div>
      <div class="circle-item"><div class="threshold">60% - 84.99%</div><div class="circle yellow">{counts.get('Kuning',0)}</div></div>
      <div class="circle-item"><div class="threshold">&lt; 60%</div><div class="circle red">{counts.get('Merah',0)}</div></div>
    </div>''', unsafe_allow_html=True)
with summary:
    st.markdown(f'''<div class="summary-box"><div class="active-no">{active}</div><div class="active-label">Jumlah Program Aktif</div>
    <div class="line"></div><div class="stats"><b>SASARAN</b> {target_avg:.1f}%<br><b>PRESTASI</b> {actual_avg:.1f}%</div>
    <div class="line"></div><div class="achievement">PENCAPAIAN {achievement:.1f}%</div></div>''', unsafe_allow_html=True)

q3 = int(counts.get("Bermula Q3", 0)); q4 = int(counts.get("Bermula Q4", 0))
november_start = int(df["Bermula November"].sum())
period_stats = ""
if q3 or q4:
    period_stats = f'Q3- <span style="color:#1e6df1">{q3}</span> &nbsp; | &nbsp; Q4- <span style="color:#1e6df1">{q4}</span>'
elif period == "Oktober" and november_start:
    period_stats = f'BERMULA NOVEMBER- <span style="color:#1e6df1">{november_start}</span>'
if period_stats:
    st.markdown(f'<div class="bottom-stat">{period_stats}</div>', unsafe_allow_html=True)

st.markdown('<div class="section-title">Pecahan Mengikut Sektor</div>', unsafe_allow_html=True)
colors={"Hijau":"#24bd50","Kuning":"#f2c318","Merah":"#e33c3c","Bermula November":"#7c5ce7","Bermula Q3":"#2f80ed","Bermula Q4":"#6b7c93","Belum Bermula":"#aab3c0"}
sector_chart=df.groupby(["Sektor","Status"],dropna=False).size().reset_index(name="Bilangan")
fig=px.bar(sector_chart,y="Sektor",x="Bilangan",color="Status",orientation="h",color_discrete_map=colors,
           category_orders={"Status":["Hijau","Kuning","Merah","Bermula November","Bermula Q3","Bermula Q4","Belum Bermula"]})
fig.update_layout(barmode="stack",height=390,margin=dict(l=10,r=10,t=10,b=10),legend_title_text="",yaxis_title="",xaxis_title="Bilangan Program")
st.plotly_chart(fig,use_container_width=True)

st.markdown(f'<div class="section-title">Senarai Program ({len(filtered)} rekod)</div>',unsafe_allow_html=True)
cols=["BIL.","Kod Program","Sektor","Bahagian","Deskripsi Program","Sasaran","Pencapaian","% Pencapaian","Status","Bermula November","Catatan/Justifikasi"]
st.dataframe(filtered[cols],use_container_width=True,hide_index=True,height=570,
             column_config={"% Pencapaian":st.column_config.NumberColumn("% Pencapaian",format="%.1f%%"),
                            "Deskripsi Program":st.column_config.TextColumn("Deskripsi Program",width="large"),
                            "Catatan/Justifikasi":st.column_config.TextColumn("Catatan/Justifikasi",width="large")})
