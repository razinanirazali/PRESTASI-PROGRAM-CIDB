
import streamlit as st
import pandas as pd
import plotly.express as px
from pathlib import Path
import re
import hmac

# =====================================================
# PAGE SETUP
# =====================================================
st.set_page_config(
    page_title="Dashboard Status Prestasi Fizikal Program CIDB 2026",
    layout="wide",
    initial_sidebar_state="expanded"
)


# =====================================================
# PASSWORD PROTECTION
# =====================================================
# Password dashboard.
# Boleh tukar nilai ini terus jika tidak menggunakan Streamlit Secrets.
DEFAULT_PASSWORD = "Progr@m123"

# Nama session baharu digunakan supaya sesi login lama tidak melepasi
# paparan password selepas fail ini dikemas kini.
AUTH_SESSION_KEY = "program_dashboard_authenticated_v2"


def get_app_password():
    """
    Ambil password daripada Streamlit Secrets jika APP_PASSWORD tersedia.
    Jika tiada, gunakan DEFAULT_PASSWORD.
    """
    try:
        secret_password = st.secrets.get("APP_PASSWORD", "")
        if str(secret_password).strip():
            return str(secret_password)
    except Exception:
        pass

    return DEFAULT_PASSWORD


def check_password():
    """Sekat semua kandungan dashboard sehingga password yang betul dimasukkan."""
    if st.session_state.get(AUTH_SESSION_KEY, False):
        return

    # Pastikan sidebar tidak kelihatan pada halaman login.
    st.markdown(
        """
        <style>
        section[data-testid="stSidebar"] {
            display: none !important;
        }

        header[data-testid="stHeader"] {
            background: transparent !important;
        }

        .stApp {
            background:
                radial-gradient(circle at 15% 10%, rgba(191,219,254,0.55) 0%, transparent 30%),
                radial-gradient(circle at 85% 18%, rgba(220,252,231,0.42) 0%, transparent 32%),
                linear-gradient(135deg, #f8fafc 0%, #eef3f8 50%, #ffffff 100%) !important;
        }

        .main .block-container,
        .block-container {
            max-width: 560px !important;
            padding-top: 5rem !important;
            padding-left: 1.5rem !important;
            padding-right: 1.5rem !important;
        }

        .login-card {
            background: rgba(255,255,255,0.94);
            border: 1px solid rgba(203,213,225,0.85);
            border-radius: 26px;
            padding: 34px 36px 24px 36px;
            box-shadow:
                0 22px 55px rgba(15,23,42,0.16),
                inset 0 1px 0 rgba(255,255,255,0.95);
            text-align: center;
            margin-bottom: 20px;
        }

        .login-icon {
            font-size: 48px;
            line-height: 1;
            margin-bottom: 12px;
        }

        .login-title {
            font-size: 32px;
            font-weight: 900;
            color: #1f2937;
            margin-bottom: 7px;
        }

        .login-subtitle {
            font-size: 15px;
            color: #64748b;
            margin-bottom: 0;
        }

        div[data-testid="stForm"] {
            background: rgba(255,255,255,0.92);
            border: 1px solid rgba(203,213,225,0.80);
            border-radius: 22px;
            padding: 22px 24px 18px 24px;
            box-shadow: 0 12px 32px rgba(15,23,42,0.10);
        }
        </style>

        <div class="login-card">
            <div class="login-icon">🔒</div>
            <div class="login-title">Dashboard Progam CIDB</div>
            <div class="login-subtitle">
                Sila masukkan kata laluan untuk membuka dashboard
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    with st.form("program_dashboard_login_form", clear_on_submit=False):
        entered_password = st.text_input(
            "Kata Laluan",
            type="password",
            placeholder="Masukkan kata laluan",
            key="program_dashboard_password_input"
        )

        submit = st.form_submit_button(
            "Masuk ke Dashboard",
            use_container_width=True
        )

        if submit:
            if hmac.compare_digest(
                str(entered_password),
                str(get_app_password())
            ):
                st.session_state[AUTH_SESSION_KEY] = True
                st.session_state.pop("program_dashboard_password_input", None)
                st.rerun()
            else:
                st.error("❌ Kata laluan tidak betul.")

    # Sangat penting: hentikan pelaksanaan sebelum mana-mana kod dashboard dijalankan.
    st.stop()


check_password()

QUARTER_CONFIG = {
    "Suku Pertama": {
        "code": "Q1",
        "title": "PENCAPAIAN PRESTASI FIZIKAL PROGRAM CIDB SUKU PERTAMA 2026",
        "sheet_options": [
            " DATA DASHBOARD",
            "DATA DASHBOARD",
            " DATA DASHBOARD Q1",
            "DATA DASHBOARD Q1"
        ],
        "sasaran_panel": 25.00,
    },
    "Suku Kedua": {
        "code": "Q2",
        "title": "PENCAPAIAN PRESTASI FIZIKAL PROGRAM CIDB SUKU KEDUA 2026",
        "sheet_options": [
            "DATA DASHBOARD Q2 CLEAN",
            " DATA DASHBOARD Q2 CLEAN"
        ],
        "sasaran_panel": 50.00,
    },
}

# Nilai ini ditetapkan selepas pengguna memilih tab suku tahun.
ACTIVE_QUARTER_LABEL = "Suku Pertama"
ACTIVE_QUARTER = "Q1"
ACTIVE_SHEET_OPTIONS = QUARTER_CONFIG[ACTIVE_QUARTER_LABEL]["sheet_options"]
ACTIVE_SASARAN_PANEL = QUARTER_CONFIG[ACTIVE_QUARTER_LABEL]["sasaran_panel"]

# =====================================================
# LOKASI FAIL EXCEL
# =====================================================
# Streamlit Community Cloud
DATA_FOLDER = Path(__file__).parent
EXCEL_FILENAME = "Laporan Pencapaian Prestasi Fizikal Program CIDB Q1 2026 - DASHBOARD.xlsx"
EXCEL_PATH = DATA_FOLDER / EXCEL_FILENAME

# Excel Column:
# L = WEIGHTAGE = index 11
# M = % PENCAPAIAN = index 12
WEIGHTAGE_COL_INDEX = 11
PENCAPAIAN_COL_INDEX = 12

# Column I = DATA DARI BAHAGIAN / PENCAPAIAN FIZIKAL
PENCAPAIAN_FIZIKAL_COL_INDEX = 8

# Column K = STATUS KHAS / MAKLUMAT BERMULA Q2, Q3, Q4
STATUS_TEXT_COL_INDEX = 10

# Column AD / kolum CATATAN (JUSTIFIKASI) Q2 digunakan untuk mengesan
# status khas BERMULA Q3 dan BERMULA Q4.
Q2_JUSTIFIKASI_COL_INDEX = 30
Q2_JUSTIFIKASI_COL = None


# =====================================================
# CSS
# =====================================================
st.markdown(
    """
    <style>
    .traffic-container {
        background: #eef3f8;
        border-radius: 18px;
        padding: 34px 34px 24px 34px;
        margin-bottom: 20px;
        text-align: center;
    }

    .traffic-title {
        text-align: center;
        font-size: 18px;
        font-weight: 900;
        color: #111827;
        margin-bottom: 18px;
        min-height: 50px;
        display: flex;
        align-items: center;
        justify-content: center;
    }

    .traffic-range {
        text-align: center;
        font-size: 15px;
        font-weight: 800;
        color: #374151;
        margin-top: 14px;
        min-height: 26px;
        display: flex;
        align-items: center;
        justify-content: center;
    }

    .st-key-btn_hijau button,
    .st-key-btn_kuning button,
    .st-key-btn_merah button,
    .st-key-btn_gugur button,
    .st-key-btn_tidak button {
        width: 135px !important;
        height: 135px !important;
        border-radius: 50% !important;
        border: none !important;
        font-size: 42px !important;
        font-weight: 900 !important;
        margin-left: auto !important;
        margin-right: auto !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        box-shadow: 0 18px 35px rgba(0,0,0,0.16) !important;
        transition: 0.15s ease-in-out !important;
    }

    .st-key-btn_hijau button {
        background: #2fb463 !important;
        color: white !important;
    }

    .st-key-btn_kuning button {
        background: #f6c90e !important;
        color: #263042 !important;
    }

    .st-key-btn_merah button {
        background: #ef463b !important;
        color: white !important;
    }

    .st-key-btn_gugur button,
    .st-key-btn_tidak button {
        background: #3c3948 !important;
        color: white !important;
    }

    .st-key-btn_hijau button:hover,
    .st-key-btn_kuning button:hover,
    .st-key-btn_merah button:hover,
    .st-key-btn_gugur button:hover,
    .st-key-btn_tidak button:hover {
        transform: scale(1.04);
        border: 3px solid #ffffff !important;
    }

    .st-key-btn_hijau button p,
    .st-key-btn_kuning button p,
    .st-key-btn_merah button p,
    .st-key-btn_gugur button p,
    .st-key-btn_tidak button p {
        font-size: 42px !important;
        font-weight: 900 !important;
    }

    /* FINAL INLINE GUGUR / TIDAK DILAKSANAKAN */
    .inline-status-wrap {
        text-align: center;
        margin-top: 18px;
        margin-bottom: 4px;
        color: #2f3b4d;
        font-size: 18px;
        font-weight: 900;
    }

    .inline-status-label {
        display: inline-block;
        vertical-align: middle;
        padding-top: 6px;
    }

    .inline-status-separator {
        display: inline-block;
        vertical-align: middle;
        padding: 0 12px;
        color: #8a94a3;
        font-weight: 900;
    }

    .st-key-btn_gugur_value button,
    .st-key-btn_tidak_value button {
        background: transparent !important;
        border: none !important;
        box-shadow: none !important;
        color: #2f3b4d !important;
        font-size: 26px !important;
        font-weight: 900 !important;
        min-height: 0 !important;
        height: auto !important;
        width: auto !important;
        padding: 0 4px !important;
        margin: 0 !important;
        display: inline-flex !important;
        align-items: center !important;
        justify-content: center !important;
    }

    .st-key-btn_gugur_value button:hover,
    .st-key-btn_tidak_value button:hover {
        color: #ef463b !important;
        text-decoration: underline !important;
        transform: none !important;
        border: none !important;
    }

    .st-key-btn_gugur_value button p,
    .st-key-btn_tidak_value button p {
        font-size: 26px !important;
        font-weight: 900 !important;
    }


    /* FINAL CLICKABLE CIRCLE BUTTONS - NO CLICK TEXT */
    .traffic-title {
        width: 135px !important;
        margin-left: auto !important;
        margin-right: auto !important;
        text-align: center !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
    }

    .traffic-range {
        width: 135px !important;
        margin-left: auto !important;
        margin-right: auto !important;
        text-align: center !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
    }

    .st-key-btn_hijau,
    .st-key-btn_kuning,
    .st-key-btn_merah {
        display: flex !important;
        justify-content: center !important;
        align-items: center !important;
    }

    .st-key-btn_hijau button,
    .st-key-btn_kuning button,
    .st-key-btn_merah button {
        width: 135px !important;
        height: 135px !important;
        min-height: 135px !important;
        border-radius: 50% !important;
        border: none !important;
        font-size: 42px !important;
        font-weight: 900 !important;
        padding: 0 !important;
        margin-left: auto !important;
        margin-right: auto !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        box-shadow: 0 18px 35px rgba(0,0,0,0.16) !important;
    }

    .st-key-btn_hijau button {
        background: #2fb463 !important;
        color: white !important;
    }

    .st-key-btn_kuning button {
        background: #f6c90e !important;
        color: #263042 !important;
    }

    .st-key-btn_merah button {
        background: #ef463b !important;
        color: white !important;
    }

    .st-key-btn_hijau button:hover,
    .st-key-btn_kuning button:hover,
    .st-key-btn_merah button:hover {
        transform: scale(1.04);
        border: 3px solid #ffffff !important;
    }

    .st-key-btn_hijau button p,
    .st-key-btn_kuning button p,
    .st-key-btn_merah button p {
        font-size: 42px !important;
        font-weight: 900 !important;
        margin: 0 !important;
        padding: 0 !important;
    }


    /* FINAL LEFT ALIGN LABELS */
    .traffic-title {
        width: 135px !important;
        margin-left: auto !important;
        margin-right: auto !important;
        text-align: left !important;
        display: block !important;
        align-items: unset !important;
        justify-content: unset !important;
    }

    .traffic-range {
        width: 135px !important;
        margin-left: auto !important;
        margin-right: auto !important;
        text-align: left !important;
        display: block !important;
        align-items: unset !important;
        justify-content: unset !important;
    }

    .inline-status-wrap {
        text-align: left !important;
        margin-top: 18px !important;
        margin-bottom: 4px !important;
        color: #2f3b4d !important;
        font-size: 18px !important;
        font-weight: 900 !important;
        white-space: nowrap !important;
    }

    .st-key-btn_gugur_value,
    .st-key-btn_tidak_value {
        text-align: left !important;
    }

    .st-key-btn_gugur_value button,
    .st-key-btn_tidak_value button {
        margin-left: 0 !important;
        margin-right: auto !important;
        background: transparent !important;
        border: none !important;
        box-shadow: none !important;
        color: #2f3b4d !important;
        font-size: 26px !important;
        font-weight: 900 !important;
        min-height: 0 !important;
        height: auto !important;
        width: auto !important;
        padding: 0 4px !important;
    }


    /* RIGHT SUMMARY PANEL */
    .summary-panel {
        text-align: center;
        color: #2f3b4d;
        padding-top: 18px;
        padding-left: 10px;
        padding-right: 10px;
    }

    .summary-total {
        font-size: 46px;
        font-weight: 900;
        line-height: 1;
        margin-bottom: 4px;
    }

    .summary-label {
        font-size: 17px;
        font-weight: 900;
        margin-bottom: 12px;
    }

    .summary-line {
        border-top: 1px solid #c8d0da;
        margin: 10px 0 12px 0;
    }

    .summary-row {
        font-size: 21px;
        font-weight: 900;
        line-height: 1.35;
    }

    .summary-row span {
        font-weight: 500;
    }

    .summary-achievement {
        font-size: 25px;
        font-weight: 900;
        color: #2fb463;
        margin-top: 14px;
    }


    /* CLEAN TRAFFIC LAYOUT - SMALL BUTTONS TOP, GUGUR ONLY BELOW */
    .traffic-container {
        background: #eef3f8;
        border-radius: 18px;
        padding: 28px 34px 28px 34px;
        margin-bottom: 20px;
        text-align: center;
    }

    .top-click-btn-label {
        text-align: center;
        margin-bottom: 10px;
    }

    .st-key-btn_hijau_top button,
    .st-key-btn_kuning_top button,
    .st-key-btn_merah_top button {
        width: 62px !important;
        height: 62px !important;
        min-height: 62px !important;
        border-radius: 50% !important;
        border: 2px solid rgba(255,255,255,0.75) !important;
        color: transparent !important;
        padding: 0 !important;
        margin: 0 auto !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        position: relative !important;
        cursor: pointer !important;
        transition: all 0.12s ease-in-out !important;
        transform: translateY(0) !important;
    }

    .st-key-btn_hijau_top button {
        background: radial-gradient(circle at 30% 25%, #c8ffd7 0%, #39c978 38%, #1f8f4e 100%) !important;
        box-shadow:
            inset 0 5px 9px rgba(255,255,255,0.55),
            inset 0 -8px 12px rgba(0,0,0,0.22),
            0 8px 0 #16683a,
            0 14px 24px rgba(0,0,0,0.22) !important;
    }

    .st-key-btn_kuning_top button {
        background: radial-gradient(circle at 30% 25%, #fff7b8 0%, #f6c90e 42%, #c49100 100%) !important;
        box-shadow:
            inset 0 5px 9px rgba(255,255,255,0.65),
            inset 0 -8px 12px rgba(0,0,0,0.20),
            0 8px 0 #8f6900,
            0 14px 24px rgba(0,0,0,0.22) !important;
    }

    .st-key-btn_merah_top button {
        background: radial-gradient(circle at 30% 25%, #ffb3b7 0%, #ef463b 40%, #b51f24 100%) !important;
        box-shadow:
            inset 0 5px 9px rgba(255,255,255,0.55),
            inset 0 -8px 12px rgba(0,0,0,0.24),
            0 8px 0 #81171a,
            0 14px 24px rgba(0,0,0,0.24) !important;
    }

    .st-key-btn_hijau_top button:after,
    .st-key-btn_kuning_top button:after,
    .st-key-btn_merah_top button:after {
        content: "" !important;
        position: absolute !important;
        top: 10px !important;
        left: 14px !important;
        width: 18px !important;
        height: 10px !important;
        border-radius: 50% !important;
        background: rgba(255,255,255,0.65) !important;
        filter: blur(0.2px) !important;
        display: block !important;
    }

    .st-key-btn_hijau_top button:hover,
    .st-key-btn_kuning_top button:hover,
    .st-key-btn_merah_top button:hover {
        transform: translateY(-2px) scale(1.04) !important;
        border: 2px solid #ffffff !important;
    }

    .st-key-btn_hijau_top button:active,
    .st-key-btn_kuning_top button:active,
    .st-key-btn_merah_top button:active {
        transform: translateY(6px) scale(0.98) !important;
        box-shadow:
            inset 0 3px 7px rgba(255,255,255,0.35),
            inset 0 -4px 8px rgba(0,0,0,0.26),
            0 2px 0 rgba(0,0,0,0.45),
            0 6px 12px rgba(0,0,0,0.18) !important;
    }

    .static-traffic-card {
        text-align: center;
        width: 100%;
    }

    .static-circle {
        width: 135px;
        height: 135px;
        border-radius: 50%;
        margin: 14px auto 14px auto;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 42px;
        font-weight: 900;
        box-shadow: 0 18px 35px rgba(0,0,0,0.16);
    }

    .static-green {
        background: #2fb463;
        color: #ffffff;
    }

    .static-yellow {
        background: #f6c90e;
        color: #263042;
    }

    .static-red {
        background: #ef463b;
        color: #ffffff;
    }

    .static-range {
        text-align: center;
        font-size: 16px;
        font-weight: 900;
        color: #111827;
        margin-bottom: 6px;
    }

    .bottom-special-wrap {
        margin-top: 24px;
        text-align: center;
    }

    .bottom-special-label {
        text-align: center;
        font-size: 14px;
        font-weight: 900;
        color: #2f3b4d;
        margin-bottom: 3px;
        white-space: nowrap;
    }

    .st-key-btn_gugur_value button,
    .st-key-btn_tidak_value button {
        background: transparent !important;
        border: none !important;
        box-shadow: none !important;
        color: #2f3b4d !important;
        font-size: 24px !important;
        font-weight: 900 !important;
        min-height: 0 !important;
        height: auto !important;
        width: auto !important;
        padding: 0 !important;
        margin: 0 auto !important;
        display: block !important;
    }

    .st-key-btn_gugur_value button:hover,
    .st-key-btn_tidak_value button:hover {
        color: #ef463b !important;
        text-decoration: underline !important;
        transform: none !important;
        border: none !important;
    }

    .st-key-btn_gugur_value button p,
    .st-key-btn_tidak_value button p {
        font-size: 24px !important;
        font-weight: 900 !important;
        margin: 0 !important;
        padding: 0 !important;
    }


    /* INLINE GUGUR / TIDAK DILAKSANAKAN */
    .inline-gugur-label {
        text-align: right !important;
        font-size: 18px !important;
        font-weight: 900 !important;
        color: #2f3b4d !important;
        padding-top: 3px !important;
        white-space: nowrap !important;
    }

    .inline-separator {
        text-align: center !important;
        font-size: 22px !important;
        font-weight: 900 !important;
        color: #2f3b4d !important;
        padding-top: 1px !important;
    }

    .st-key-btn_gugur_value button,
    .st-key-btn_tidak_value button {
        background: transparent !important;
        border: none !important;
        box-shadow: none !important;
        color: #2f3b4d !important;
        font-size: 22px !important;
        font-weight: 900 !important;
        padding: 0 !important;
        margin: 0 !important;
        min-height: 0 !important;
        height: auto !important;
        width: auto !important;
        display: block !important;
        text-align: left !important;
    }

    .st-key-btn_gugur_value button:hover,
    .st-key-btn_tidak_value button:hover {
        color: #ef463b !important;
        text-decoration: underline !important;
        transform: none !important;
        border: none !important;
    }

    .st-key-btn_gugur_value button p,
    .st-key-btn_tidak_value button p {
        font-size: 22px !important;
        font-weight: 900 !important;
        margin: 0 !important;
        padding: 0 !important;
    }


    /* PREMIUM LED CIRCLE VALUE DISPLAY */
    .static-traffic-card {
        text-align: center !important;
        width: 100% !important;
        position: relative !important;
    }

    .static-circle {
        width: 110px !important;
        height: 110px !important;
        border-radius: 50% !important;
        margin: 14px auto 14px auto !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        font-size: 36px !important;
        font-weight: 900 !important;
        position: relative !important;
        overflow: hidden !important;
        border: 4px solid rgba(255,255,255,0.35) !important;
        outline: none !important;
        transition: all 0.18s ease-in-out !important;
        transform: translateY(0) !important;
    }

    .static-circle::before {
        content: "" !important;
        position: absolute !important;
        top: 13px !important;
        left: 22px !important;
        width: 78px !important;
        height: 30px !important;
        border-radius: 50% !important;
        background: rgba(255,255,255,0.35) !important;
        transform: rotate(-18deg) !important;
        z-index: 1 !important;
        pointer-events: none !important;
    }

    .static-circle::after {
        content: "" !important;
        position: absolute !important;
        inset: 0 !important;
        border-radius: 50% !important;
        background-image:
            linear-gradient(rgba(255,255,255,0.08) 1px, transparent 1px),
            linear-gradient(90deg, rgba(255,255,255,0.08) 1px, transparent 1px) !important;
        background-size: 8px 8px !important;
        opacity: 0.45 !important;
        z-index: 1 !important;
        pointer-events: none !important;
    }

    .static-circle:hover {
        transform: translateY(-4px) scale(1.035) !important;
    }

    .static-green {
        color: #ffffff !important;
        background:
            radial-gradient(circle at 30% 25%, #9dffad 0%, #2ee45a 42%, #07912b 100%) !important;
        box-shadow:
            0 0 18px rgba(46,228,90,0.60),
            0 0 38px rgba(46,228,90,0.36),
            0 16px 30px rgba(0,0,0,0.22),
            inset 0 9px 13px rgba(255,255,255,0.30),
            inset 0 -14px 20px rgba(0,0,0,0.28) !important;
        text-shadow:
            0 3px 4px rgba(0,0,0,0.45),
            0 0 12px rgba(255,255,255,0.28) !important;
    }

    .static-yellow {
        color: #263042 !important;
        background:
            radial-gradient(circle at 30% 25%, #fff9b5 0%, #f6d21e 43%, #b98a00 100%) !important;
        box-shadow:
            0 0 18px rgba(246,210,30,0.62),
            0 0 38px rgba(246,210,30,0.36),
            0 16px 30px rgba(0,0,0,0.22),
            inset 0 9px 13px rgba(255,255,255,0.42),
            inset 0 -14px 20px rgba(0,0,0,0.20) !important;
        text-shadow:
            0 1px 2px rgba(255,255,255,0.50),
            0 2px 4px rgba(0,0,0,0.18) !important;
    }

    .static-red {
        color: #ffffff !important;
        background:
            radial-gradient(circle at 30% 25%, #ffaaaa 0%, #f04a42 42%, #a51218 100%) !important;
        box-shadow:
            0 0 18px rgba(240,74,66,0.62),
            0 0 38px rgba(240,74,66,0.36),
            0 16px 30px rgba(0,0,0,0.22),
            inset 0 9px 13px rgba(255,255,255,0.28),
            inset 0 -14px 20px rgba(0,0,0,0.30) !important;
        text-shadow:
            0 3px 4px rgba(0,0,0,0.45),
            0 0 12px rgba(255,255,255,0.25) !important;
    }

    .static-range {
        text-align: center !important;
        font-size: 14px !important;
        font-weight: 900 !important;
        color: #111827 !important;
        margin-bottom: 8px !important;
        position: relative !important;
        z-index: 2 !important;
    }


    /* ACCORDION / EXPANDER STYLE */
    div[data-testid="stExpander"] {
        background: #f8fafc !important;
        border: 1px solid #d7dde6 !important;
        border-radius: 22px !important;
        box-shadow:
            0 8px 18px rgba(15, 23, 42, 0.08),
            inset 0 1px 0 rgba(255,255,255,0.85) !important;
        margin-bottom: 18px !important;
        overflow: visible !important;
    }

    div[data-testid="stExpander"] details {
        border-radius: 22px !important;
    }

    div[data-testid="stExpander"] summary {
        min-height: 56px !important;
        padding: 0 20px !important;
        font-size: 17px !important;
        font-weight: 900 !important;
        color: #374151 !important;
        letter-spacing: 0.2px !important;
    }

    div[data-testid="stExpander"] summary:hover {
        background: #eef3f8 !important;
    }

    div[data-testid="stExpander"] div[data-testid="stExpanderDetails"] {
        background: #ffffff !important;
        border-top: 1px solid #e5e7eb !important;
        padding-top: 18px !important;
    }


    /* CLICKABLE JUMLAH PROGRAM BUTTON */
    .st-key-btn_jumlah_program button {
        background: transparent !important;
        border: none !important;
        box-shadow: none !important;
        color: #2f3b4d !important;
        font-size: 46px !important;
        font-weight: 900 !important;
        line-height: 1 !important;
        padding: 0 !important;
        margin: 0 auto 4px auto !important;
        min-height: 0 !important;
        height: auto !important;
        width: auto !important;
        display: block !important;
    }

    .st-key-btn_jumlah_program button:hover {
        color: #245be8 !important;
        transform: scale(1.04) !important;
        text-decoration: underline !important;
        border: none !important;
    }

    .st-key-btn_jumlah_program button p {
        font-size: 46px !important;
        font-weight: 900 !important;
        margin: 0 !important;
        padding: 0 !important;
    }


    /* CLICKABLE LARGE TRAFFIC LIGHT VALUE BUTTONS */
    .st-key-btn_hijau_value button,
    .st-key-btn_kuning_value button,
    .st-key-btn_merah_value button {
        width: 110px !important;
        height: 110px !important;
        min-height: 110px !important;
        border-radius: 50% !important;
        margin: 14px auto 14px auto !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        font-size: 44px !important;
        font-weight: 900 !important;
        position: relative !important;
        overflow: hidden !important;
        border: 4px solid rgba(255,255,255,0.35) !important;
        transition: all 0.18s ease-in-out !important;
        cursor: pointer !important;
        padding: 0 !important;
    }

    .st-key-btn_hijau_value button {
        color: #ffffff !important;
        background:
            radial-gradient(circle at 30% 25%, #9dffad 0%, #2ee45a 42%, #07912b 100%) !important;
        box-shadow:
            0 0 18px rgba(46,228,90,0.60),
            0 0 38px rgba(46,228,90,0.36),
            0 16px 30px rgba(0,0,0,0.22),
            inset 0 9px 13px rgba(255,255,255,0.30),
            inset 0 -14px 20px rgba(0,0,0,0.28) !important;
        text-shadow:
            0 3px 4px rgba(0,0,0,0.45),
            0 0 12px rgba(255,255,255,0.28) !important;
    }

    .st-key-btn_kuning_value button {
        color: #263042 !important;
        background:
            radial-gradient(circle at 30% 25%, #fff9b5 0%, #f6d21e 43%, #b98a00 100%) !important;
        box-shadow:
            0 0 18px rgba(246,210,30,0.62),
            0 0 38px rgba(246,210,30,0.36),
            0 16px 30px rgba(0,0,0,0.22),
            inset 0 9px 13px rgba(255,255,255,0.42),
            inset 0 -14px 20px rgba(0,0,0,0.20) !important;
        text-shadow:
            0 1px 2px rgba(255,255,255,0.50),
            0 2px 4px rgba(0,0,0,0.18) !important;
    
        text-shadow:
            0 1px 0 rgba(255,255,255,0.70),
            0 3px 4px rgba(0,0,0,0.28),
            0 0 8px rgba(255,255,255,0.30) !important;
    }

    .st-key-btn_merah_value button {
        color: #ffffff !important;
        background:
            radial-gradient(circle at 30% 25%, #ffaaaa 0%, #f04a42 42%, #a51218 100%) !important;
        box-shadow:
            0 0 18px rgba(240,74,66,0.62),
            0 0 38px rgba(240,74,66,0.36),
            0 16px 26px rgba(0,0,0,0.22),
            inset 0 9px 13px rgba(255,255,255,0.28),
            inset 0 -14px 20px rgba(0,0,0,0.30) !important;
        text-shadow:
            0 3px 4px rgba(0,0,0,0.45),
            0 0 12px rgba(255,255,255,0.25) !important;
    }

    .st-key-btn_hijau_value button:hover,
    .st-key-btn_kuning_value button:hover,
    .st-key-btn_merah_value button:hover {
        transform: translateY(-4px) scale(1.035) !important;
        border: 4px solid rgba(255,255,255,0.70) !important;
    }

    .st-key-btn_hijau_value button:active,
    .st-key-btn_kuning_value button:active,
    .st-key-btn_merah_value button:active {
        transform: translateY(4px) scale(0.98) !important;
    }

    .st-key-btn_hijau_value button p,
    .st-key-btn_kuning_value button p,
    .st-key-btn_merah_value button p {
        font-size: 36px !important;
        font-weight: 900 !important;
        margin: 0 !important;
        padding: 0 !important;
    }


    /* HIDE SMALL TOP TRAFFIC BUTTONS */
    .st-key-btn_hijau_top,
    .st-key-btn_kuning_top,
    .st-key-btn_merah_top {
        display: none !important;
        height: 0 !important;
        min-height: 0 !important;
        margin: 0 !important;
        padding: 0 !important;
    }



    /* CLEAN SAME-LINE TRAFFIC + SUMMARY - NO BOX */
    .traffic-container {
        background: transparent !important;
        border: none !important;
        box-shadow: none !important;
        border-radius: 0 !important;
        padding: 0 !important;
        margin: 0 !important;
        min-height: auto !important;
        text-align: center !important;
    }

    .summary-panel {
        background: transparent !important;
        border: none !important;
        box-shadow: none !important;
        border-radius: 0 !important;
        padding: 0 0 0 8px !important;
        margin: 0 !important;
        min-height: auto !important;
        height: auto !important;
        text-align: left !important;
        color: #2f3b4d !important;
    }

    .summary-label {
        text-align: left !important;
        font-size: 16px !important;
        font-weight: 900 !important;
        margin-bottom: 14px !important;
    }

    .summary-line {
        border-top: 1px solid #c8d0da !important;
        margin: 14px 0 14px 0 !important;
    }

    .summary-row {
        text-align: left !important;
        font-size: 19px !important;
        font-weight: 900 !important;
        line-height: 1.45 !important;
    }

    .summary-row span {
        font-weight: 500 !important;
    }

    .summary-achievement {
        text-align: left !important;
        font-size: 24px !important;
        font-weight: 900 !important;
        color: #2fb463 !important;
        margin-top: 16px !important;
    }

    .st-key-btn_jumlah_program button {
        margin-left: 0 !important;
        margin-right: auto !important;
        text-align: left !important;
        background: transparent !important;
        box-shadow: none !important;
        border: none !important;
        padding: 0 !important;
    }

    .st-key-btn_jumlah_program button p {
        font-size: 42px !important;
        font-weight: 900 !important;
        margin: 0 !important;
        padding: 0 !important;
    }

    .static-range {
        margin-top: 0 !important;
        margin-bottom: 12px !important;
    }

    div[data-testid="column"] > div {
        padding-top: 0 !important;
    }


    /* CENTER TRAFFIC LIGHT VALUE BUTTONS UNDER LABELS */
    .st-key-btn_hijau_value,
    .st-key-btn_kuning_value,
    .st-key-btn_merah_value {
        width: 100% !important;
        display: flex !important;
        justify-content: center !important;
        align-items: center !important;
        text-align: center !important;
    }

    .st-key-btn_hijau_value button,
    .st-key-btn_kuning_value button,
    .st-key-btn_merah_value button {
        margin-left: auto !important;
        margin-right: auto !important;
        display: flex !important;
        justify-content: center !important;
        align-items: center !important;
    }

    .static-traffic-card {
        width: 100% !important;
        display: flex !important;
        justify-content: center !important;
        align-items: center !important;
        text-align: center !important;
    }

    .static-range {
        width: 100% !important;
        text-align: center !important;
        display: block !important;
        margin-left: auto !important;
        margin-right: auto !important;
    }

    /* More balanced spacing between the three traffic lights */
    .traffic-container [data-testid="column"] {
        display: flex !important;
        flex-direction: column !important;
        align-items: center !important;
        justify-content: flex-start !important;
    }


    /* REDUCE TOP GAP ABOVE MAIN TITLE */
    .block-container {
        padding-top: 1.2rem !important;
        padding-bottom: 2rem !important;
    }

    header[data-testid="stHeader"] {
        background: transparent !important;
    }

    h1 {
        margin-top: 0 !important;
        padding-top: 0 !important;
    }


    /* SOFT MIRROR / 3D BACKGROUND - MAIN PAGE + SIDEBAR */
    .stApp {
        background:
            radial-gradient(circle at 15% 10%, rgba(191, 219, 254, 0.42) 0%, transparent 28%),
            radial-gradient(circle at 85% 18%, rgba(220, 252, 231, 0.34) 0%, transparent 30%),
            radial-gradient(circle at 50% 92%, rgba(226, 232, 240, 0.72) 0%, transparent 42%),
            linear-gradient(135deg, #f8fafc 0%, #eef3f8 46%, #ffffff 100%) !important;
        background-attachment: fixed !important;
    }

    .main .block-container {
        background:
            linear-gradient(135deg, rgba(255,255,255,0.72), rgba(248,250,252,0.58)) !important;
        border-radius: 28px !important;
        box-shadow:
            0 24px 60px rgba(15, 23, 42, 0.08),
            inset 0 1px 0 rgba(255,255,255,0.90) !important;
        backdrop-filter: blur(14px) !important;
        -webkit-backdrop-filter: blur(14px) !important;
        padding-left: 3rem !important;
        padding-right: 3rem !important;
    }

    section[data-testid="stSidebar"] {
        background:
            radial-gradient(circle at 20% 0%, rgba(96, 165, 250, 0.28), transparent 35%),
            linear-gradient(180deg, #1f2937 0%, #334155 52%, #475569 100%) !important;
        box-shadow:
            16px 0 40px rgba(15, 23, 42, 0.18),
            inset -1px 0 0 rgba(255,255,255,0.12) !important;
    }

    section[data-testid="stSidebar"] * {
        color: #f8fafc !important;
    }

    section[data-testid="stSidebar"] label,
    section[data-testid="stSidebar"] p,
    section[data-testid="stSidebar"] span {
        color: #f8fafc !important;
        font-weight: 800 !important;
    }

    section[data-testid="stSidebar"] div[data-baseweb="select"] > div,
    section[data-testid="stSidebar"] div[data-baseweb="tag"] {
        background: rgba(255,255,255,0.16) !important;
        border: 1px solid rgba(255,255,255,0.22) !important;
        border-radius: 14px !important;
        box-shadow:
            inset 0 1px 0 rgba(255,255,255,0.22),
            0 8px 20px rgba(15,23,42,0.18) !important;
        backdrop-filter: blur(10px) !important;
    }

    section[data-testid="stSidebar"] input {
        color: #111827 !important;
    }

    /* Mirror card effect for expanders */
    div[data-testid="stExpander"] {
        background: rgba(255,255,255,0.72) !important;
        border: 1px solid rgba(203, 213, 225, 0.72) !important;
        box-shadow:
            0 14px 34px rgba(15, 23, 42, 0.10),
            inset 0 1px 0 rgba(255,255,255,0.92) !important;
        backdrop-filter: blur(14px) !important;
        -webkit-backdrop-filter: blur(14px) !important;
    }

    div[data-testid="stExpander"] summary {
        background:
            linear-gradient(135deg, rgba(255,255,255,0.78), rgba(241,245,249,0.82)) !important;
    }

    /* Softer premium divider */
    hr {
        border: none !important;
        height: 1px !important;
        background: linear-gradient(90deg, transparent, rgba(100,116,139,0.35), transparent) !important;
    }

    /* Keep title readable on glass background */
    h1 {
        color: #1f2937 !important;
        text-shadow: 0 1px 0 rgba(255,255,255,0.60) !important;
    }


    /* FIX SIDEBAR FILTER PILL TEXT VISIBILITY */
    section[data-testid="stSidebar"] div[data-testid="stMarkdownContainer"] p,
    section[data-testid="stSidebar"] label,
    section[data-testid="stSidebar"] [data-testid="stWidgetLabel"] {
        color: #ffffff !important;
        font-weight: 900 !important;
    }

    section[data-testid="stSidebar"] button,
    section[data-testid="stSidebar"] button p,
    section[data-testid="stSidebar"] button span {
        color: #1f2937 !important;
        font-weight: 900 !important;
        text-shadow: none !important;
    }

    section[data-testid="stSidebar"] button {
        background: rgba(255,255,255,0.92) !important;
        border: 1px solid rgba(255,255,255,0.72) !important;
        box-shadow:
            inset 0 1px 0 rgba(255,255,255,0.90),
            0 6px 14px rgba(15,23,42,0.18) !important;
    }

    section[data-testid="stSidebar"] button:hover {
        background: #ffffff !important;
        color: #111827 !important;
        transform: translateY(-1px) !important;
    }

    section[data-testid="stSidebar"] button[aria-pressed="true"],
    section[data-testid="stSidebar"] button[data-baseweb="button"][aria-pressed="true"] {
        background: linear-gradient(135deg, #dbeafe, #ffffff) !important;
        border: 2px solid #93c5fd !important;
        color: #0f172a !important;
    }

    section[data-testid="stSidebar"] button[aria-pressed="true"] p,
    section[data-testid="stSidebar"] button[aria-pressed="true"] span {
        color: #0f172a !important;
    }


    /* NAVY GLASS FILTER PILLS */
    section[data-testid="stSidebar"] button {
        background: rgba(51, 65, 85, 0.55) !important;
        border: 1px solid rgba(255,255,255,0.18) !important;
        color: #ffffff !important;
        backdrop-filter: blur(12px) !important;
        -webkit-backdrop-filter: blur(12px) !important;
        box-shadow:
            inset 0 1px 0 rgba(255,255,255,0.12),
            0 6px 16px rgba(0,0,0,0.18) !important;
    }

    section[data-testid="stSidebar"] button p,
    section[data-testid="stSidebar"] button span {
        color: #ffffff !important;
        font-weight: 800 !important;
        text-shadow: none !important;
    }

    section[data-testid="stSidebar"] button:hover {
        background: rgba(71, 85, 105, 0.75) !important;
        color: #ffffff !important;
    }

    section[data-testid="stSidebar"] button[aria-pressed="true"] {
        background: rgba(96,165,250,0.45) !important;
        border: 1px solid rgba(147,197,253,0.85) !important;
        box-shadow:
            0 0 15px rgba(96,165,250,0.35),
            inset 0 1px 0 rgba(255,255,255,0.20) !important;
    }

    section[data-testid="stSidebar"] button[aria-pressed="true"] p,
    section[data-testid="stSidebar"] button[aria-pressed="true"] span {
        color: #ffffff !important;
    }


    /* CLEAR SELECTED FILTER PILL STATE */
    section[data-testid="stSidebar"] button {
        background: rgba(51, 65, 85, 0.58) !important;
        border: 1px solid rgba(255,255,255,0.22) !important;
        color: #ffffff !important;
        border-radius: 999px !important;
        backdrop-filter: blur(12px) !important;
        -webkit-backdrop-filter: blur(12px) !important;
        box-shadow:
            inset 0 1px 0 rgba(255,255,255,0.14),
            0 6px 16px rgba(0,0,0,0.18) !important;
    }

    section[data-testid="stSidebar"] button p,
    section[data-testid="stSidebar"] button span {
        color: #ffffff !important;
        font-weight: 900 !important;
        text-shadow: none !important;
    }

    section[data-testid="stSidebar"] button:hover {
        background: rgba(71, 85, 105, 0.82) !important;
        border: 1px solid rgba(255,255,255,0.45) !important;
        transform: translateY(-1px) !important;
    }

    section[data-testid="stSidebar"] button[aria-pressed="true"],
    section[data-testid="stSidebar"] button[data-selected="true"],
    section[data-testid="stSidebar"] button[kind="primary"],
    section[data-testid="stSidebar"] button:focus {
        background:
            linear-gradient(135deg, rgba(37,99,235,0.96), rgba(59,130,246,0.96)) !important;
        border: 2px solid rgba(255,255,255,0.95) !important;
        box-shadow:
            0 0 0 2px rgba(255,255,255,0.22),
            0 0 20px rgba(59,130,246,0.72),
            0 0 38px rgba(59,130,246,0.38),
            inset 0 1px 0 rgba(255,255,255,0.28) !important;
    }

    section[data-testid="stSidebar"] button[aria-pressed="true"] p,
    section[data-testid="stSidebar"] button[data-selected="true"] p,
    section[data-testid="stSidebar"] button[kind="primary"] p,
    section[data-testid="stSidebar"] button:focus p,
    section[data-testid="stSidebar"] button[aria-pressed="true"] span,
    section[data-testid="stSidebar"] button[data-selected="true"] span,
    section[data-testid="stSidebar"] button[kind="primary"] span,
    section[data-testid="stSidebar"] button:focus span {
        color: #ffffff !important;
        font-weight: 900 !important;
    }

    


    /* PERSISTENT CUSTOM SIDEBAR PILLS */
    section[data-testid="stSidebar"] button {
        border-radius: 999px !important;
        background: rgba(51, 65, 85, 0.58) !important;
        border: 1px solid rgba(255,255,255,0.22) !important;
        color: #ffffff !important;
        min-height: 38px !important;
        padding: 0.38rem 0.82rem !important;
        box-shadow:
            inset 0 1px 0 rgba(255,255,255,0.14),
            0 6px 16px rgba(0,0,0,0.18) !important;
    }

    section[data-testid="stSidebar"] button p {
        color: #ffffff !important;
        font-weight: 900 !important;
        font-size: 13px !important;
        line-height: 1.15 !important;
        white-space: normal !important;
        overflow: visible !important;
        text-overflow: unset !important;
        word-break: break-word !important;
        text-align: left !important;
        width: 100% !important;
    }

    section[data-testid="stSidebar"] button {
        justify-content: flex-start !important;
        text-align: left !important;
    }

    section[data-testid="stSidebar"] button:hover {
        background: rgba(71, 85, 105, 0.82) !important;
        border: 1px solid rgba(255,255,255,0.45) !important;
        transform: translateY(-1px) !important;
    }


    /* KEEP ALL SELECTED SIDEBAR PILLS BLUE */
    section[data-testid="stSidebar"] button[kind="primary"],
    section[data-testid="stSidebar"] button[data-testid="baseButton-primary"] {
        background:
            linear-gradient(135deg, rgba(37,99,235,0.98), rgba(59,130,246,0.98)) !important;
        border: 2px solid rgba(255,255,255,0.95) !important;
        color: #ffffff !important;
        box-shadow:
            0 0 0 2px rgba(255,255,255,0.22),
            0 0 20px rgba(59,130,246,0.75),
            0 0 38px rgba(59,130,246,0.40),
            inset 0 1px 0 rgba(255,255,255,0.28) !important;
    }

    section[data-testid="stSidebar"] button[kind="primary"] p,
    section[data-testid="stSidebar"] button[data-testid="baseButton-primary"] p,
    section[data-testid="stSidebar"] button[kind="primary"] span,
    section[data-testid="stSidebar"] button[data-testid="baseButton-primary"] span {
        color: #ffffff !important;
        font-weight: 900 !important;
    }

    section[data-testid="stSidebar"] button[kind="secondary"],
    section[data-testid="stSidebar"] button[data-testid="baseButton-secondary"] {
        background: rgba(51, 65, 85, 0.58) !important;
        border: 1px solid rgba(255,255,255,0.22) !important;
        color: #ffffff !important;
        box-shadow:
            inset 0 1px 0 rgba(255,255,255,0.14),
            0 6px 16px rgba(0,0,0,0.18) !important;
    }

    section[data-testid="stSidebar"] button[kind="secondary"] p,
    section[data-testid="stSidebar"] button[data-testid="baseButton-secondary"] p,
    section[data-testid="stSidebar"] button[kind="secondary"] span,
    section[data-testid="stSidebar"] button[data-testid="baseButton-secondary"] span {
        color: #ffffff !important;
        font-weight: 900 !important;
    }

    section[data-testid="stSidebar"] button:focus {
        outline: none !important;
    }



    /* CLICKABLE Q2/Q3/Q4/TIDAK DILAKSANAKAN STATUS ROW */
    .st-key-btn_q2_status button,
    .st-key-btn_q3_status button,
    .st-key-btn_q4_status button,
    .st-key-btn_tidak_status button {
        background: transparent !important;
        border: none !important;
        box-shadow: none !important;
        color: #245be8 !important;
        font-size: 26px !important;
        font-weight: 900 !important;
        line-height: 1 !important;
        padding: 0 !important;
        margin-top: -7px !important;
        margin-bottom: 0 !important;
        margin-left: 0 !important;
        margin-right: 0 !important;
        min-height: 0 !important;
        height: auto !important;
        width: auto !important;
        text-align: left !important;
    }

    .st-key-btn_q2_status button:hover,
    .st-key-btn_q3_status button:hover,
    .st-key-btn_q4_status button:hover,
    .st-key-btn_tidak_status button:hover {
        color: #1d4ed8 !important;
        text-decoration: underline !important;
        transform: none !important;
        border: none !important;
    }

    .st-key-btn_q2_status button p,
    .st-key-btn_q3_status button p,
    .st-key-btn_q4_status button p,
    .st-key-btn_tidak_status button p {
        color: #245be8 !important;
        font-size: 30px !important;
        font-weight: 900 !important;
        line-height: 1 !important;
        margin: 0 !important;
        padding: 0 !important;
    }

    .status-separator-text {
        text-align: center;
        font-size: 26px;
        font-weight: 900;
        color: #2f3b4d;
        line-height: 1;
        padding-top: 1px;
    }

    .status-label-text {
        font-size: 25px;
        font-weight: 900;
        color: #2f3b4d;
        text-align: right;
        line-height: 1;
        padding-top: 2px;
        white-space: nowrap;
    }


    /* SIDEBAR EXPANDER FOR KOD PROGRAM */
    section[data-testid="stSidebar"] div[data-testid="stExpander"] {
        background: rgba(51, 65, 85, 0.45) !important;
        border: 1px solid rgba(255,255,255,0.20) !important;
        border-radius: 18px !important;
        box-shadow:
            inset 0 1px 0 rgba(255,255,255,0.12),
            0 8px 20px rgba(0,0,0,0.16) !important;
        margin-bottom: 14px !important;
    }

    section[data-testid="stSidebar"] div[data-testid="stExpander"] summary {
        color: #ffffff !important;
        font-weight: 900 !important;
        min-height: 42px !important;
        padding: 0 12px !important;
        background: transparent !important;
    }

    section[data-testid="stSidebar"] div[data-testid="stExpanderDetails"] {
        background: transparent !important;
        padding-top: 8px !important;
    }


    /* KOD PROGRAM FILTER - SMALL FONT */
    section[data-testid="stSidebar"] div[class*="st-key-filter_kod_program"] button {
        padding: 0.26rem 0.42rem !important;
        min-height: 30px !important;
    }

    section[data-testid="stSidebar"] div[class*="st-key-filter_kod_program"] button p {
        font-size: 10px !important;
        line-height: 1.05 !important;
        text-align: center !important;
        font-weight: 900 !important;
    }



    /* FOCUS TRAFFIC LIGHT: SAME APPEARANCE AS MAIN PAGE */
    .st-key-focus_btn_hijau_value,
    .st-key-focus_btn_kuning_value,
    .st-key-focus_btn_merah_value {
        width: 100% !important;
        display: flex !important;
        justify-content: center !important;
        align-items: center !important;
        text-align: center !important;
    }

    .st-key-focus_btn_hijau_value button,
    .st-key-focus_btn_kuning_value button,
    .st-key-focus_btn_merah_value button {
        width: 110px !important;
        height: 110px !important;
        min-height: 110px !important;
        border-radius: 50% !important;
        margin: 14px auto 14px auto !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        font-size: 44px !important;
        font-weight: 900 !important;
        position: relative !important;
        overflow: hidden !important;
        border: 4px solid rgba(255,255,255,0.35) !important;
        transition: all 0.18s ease-in-out !important;
        cursor: pointer !important;
        padding: 0 !important;
    }

    .st-key-focus_btn_hijau_value button {
        color: #ffffff !important;
        background: radial-gradient(circle at 30% 25%, #9dffad 0%, #2ee45a 42%, #07912b 100%) !important;
        box-shadow: 0 0 18px rgba(46,228,90,0.60), 0 0 38px rgba(46,228,90,0.36), 0 16px 30px rgba(0,0,0,0.22), inset 0 9px 13px rgba(255,255,255,0.30), inset 0 -14px 20px rgba(0,0,0,0.28) !important;
        text-shadow: 0 3px 4px rgba(0,0,0,0.45), 0 0 12px rgba(255,255,255,0.28) !important;
    }

    .st-key-focus_btn_kuning_value button {
        color: #263042 !important;
        background: radial-gradient(circle at 30% 25%, #fff9b5 0%, #f6d21e 43%, #b98a00 100%) !important;
        box-shadow: 0 0 18px rgba(246,210,30,0.62), 0 0 38px rgba(246,210,30,0.36), 0 16px 30px rgba(0,0,0,0.22), inset 0 9px 13px rgba(255,255,255,0.42), inset 0 -14px 20px rgba(0,0,0,0.20) !important;
        text-shadow: 0 1px 0 rgba(255,255,255,0.70), 0 3px 4px rgba(0,0,0,0.28), 0 0 8px rgba(255,255,255,0.30) !important;
    }

    .st-key-focus_btn_merah_value button {
        color: #ffffff !important;
        background: radial-gradient(circle at 30% 25%, #ffaaaa 0%, #f04a42 42%, #a51218 100%) !important;
        box-shadow: 0 0 18px rgba(240,74,66,0.62), 0 0 38px rgba(240,74,66,0.36), 0 16px 26px rgba(0,0,0,0.22), inset 0 9px 13px rgba(255,255,255,0.28), inset 0 -14px 20px rgba(0,0,0,0.30) !important;
        text-shadow: 0 3px 4px rgba(0,0,0,0.45), 0 0 12px rgba(255,255,255,0.25) !important;
    }

    .st-key-focus_btn_hijau_value button:hover,
    .st-key-focus_btn_kuning_value button:hover,
    .st-key-focus_btn_merah_value button:hover {
        transform: translateY(-4px) scale(1.035) !important;
        border: 4px solid rgba(255,255,255,0.70) !important;
    }

    .st-key-focus_btn_hijau_value button p,
    .st-key-focus_btn_kuning_value button p,
    .st-key-focus_btn_merah_value button p {
        font-size: 36px !important;
        font-weight: 900 !important;
        margin: 0 !important;
        padding: 0 !important;
    }

    .st-key-focus_btn_q2_status button,
    .st-key-focus_btn_q3_status button,
    .st-key-focus_btn_q4_status button,
    .st-key-focus_btn_gugur_status button {
        background: transparent !important;
        border: none !important;
        box-shadow: none !important;
        color: #245be8 !important;
        font-size: 26px !important;
        font-weight: 900 !important;
        line-height: 1 !important;
        padding: 0 !important;
        margin-top: -7px !important;
        margin-bottom: 0 !important;
        margin-left: 0 !important;
        margin-right: 0 !important;
        min-height: 0 !important;
        height: auto !important;
        width: auto !important;
        text-align: left !important;
    }

    .st-key-focus_btn_q2_status button:hover,
    .st-key-focus_btn_q3_status button:hover,
    .st-key-focus_btn_q4_status button:hover,
    .st-key-focus_btn_gugur_status button:hover {
        color: #1d4ed8 !important;
        text-decoration: underline !important;
        transform: none !important;
        border: none !important;
    }

    .st-key-focus_btn_q2_status button p,
    .st-key-focus_btn_q3_status button p,
    .st-key-focus_btn_q4_status button p,
    .st-key-focus_btn_gugur_status button p {
        color: #245be8 !important;
        font-size: 30px !important;
        font-weight: 900 !important;
        line-height: 1 !important;
        margin: 0 !important;
        padding: 0 !important;
    }

    .st-key-focus_btn_jumlah_program button {
        margin-left: 0 !important;
        margin-right: auto !important;
        text-align: left !important;
        background: transparent !important;
        box-shadow: none !important;
        border: none !important;
        padding: 0 !important;
        color: #2f3b4d !important;
        font-size: 46px !important;
        font-weight: 900 !important;
        line-height: 1 !important;
        min-height: 0 !important;
        height: auto !important;
        width: auto !important;
        display: block !important;
    }

    .st-key-focus_btn_jumlah_program button:hover {
        color: #245be8 !important;
        transform: scale(1.04) !important;
        text-decoration: underline !important;
        border: none !important;
    }

    .st-key-focus_btn_jumlah_program button p {
        font-size: 42px !important;
        font-weight: 900 !important;
        margin: 0 !important;
        padding: 0 !important;
    }


    /* TAB SUKU PERTAMA / SUKU KEDUA */
    div[data-testid="stRadio"] > div {
        display: flex !important;
        gap: 8px !important;
        padding: 5px !important;
        width: fit-content !important;
        border-radius: 14px !important;
        background: rgba(226,232,240,0.82) !important;
        border: 1px solid rgba(148,163,184,0.45) !important;
        box-shadow: inset 0 1px 0 rgba(255,255,255,0.85) !important;
    }

    div[data-testid="stRadio"] label {
        margin: 0 !important;
        padding: 8px 18px !important;
        border-radius: 10px !important;
        font-weight: 900 !important;
        cursor: pointer !important;
    }

    div[data-testid="stRadio"] label:has(input:checked) {
        background: linear-gradient(135deg, #2563eb, #3b82f6) !important;
        color: #ffffff !important;
        box-shadow: 0 6px 16px rgba(37,99,235,0.28) !important;
    }

    div[data-testid="stRadio"] label:has(input:checked) p {
        color: #ffffff !important;
    }

    div[data-testid="stRadio"] input {
        display: none !important;
    }

    /* COMPACT Q2 / Q3 / Q4 / GUGUR - KEKAL SATU BARIS */
    .compact-status-row [data-testid="stHorizontalBlock"] {
        flex-wrap: nowrap !important;
        gap: 0.20rem !important;
        align-items: center !important;
    }

    .compact-status-row [data-testid="column"] {
        min-width: 0 !important;
        width: auto !important;
        flex: 0 0 auto !important;
    }

    .compact-status-label {
        font-size: 25px !important;
        font-weight: 900 !important;
        color: #2f3b4d !important;
        line-height: 1 !important;
        white-space: nowrap !important;
        text-align: right !important;
        padding-top: 2px !important;
    }

    .compact-status-separator {
        font-size: 26px !important;
        font-weight: 900 !important;
        color: #2f3b4d !important;
        line-height: 1 !important;
        white-space: nowrap !important;
        text-align: center !important;
        padding: 0 1px !important;
    }

    .st-key-btn_q2_status button,
    .st-key-btn_q3_status button,
    .st-key-btn_q4_status button,
    .st-key-btn_tidak_status button,
    .st-key-focus_btn_q2_status button,
    .st-key-focus_btn_q3_status button,
    .st-key-focus_btn_q4_status button,
    .st-key-focus_btn_gugur_status button {
        white-space: nowrap !important;
        min-width: 38px !important;
        width: auto !important;
        padding-left: 0 !important;
        padding-right: 0 !important;
    }

    .st-key-btn_q2_status button p,
    .st-key-btn_q3_status button p,
    .st-key-btn_q4_status button p,
    .st-key-btn_tidak_status button p,
    .st-key-focus_btn_q2_status button p,
    .st-key-focus_btn_q3_status button p,
    .st-key-focus_btn_q4_status button p,
    .st-key-focus_btn_gugur_status button p {
        white-space: nowrap !important;
    }

    </style>
    """,
    unsafe_allow_html=True
)


# =====================================================
# HELPER FUNCTIONS
# =====================================================
def clean_col(col):
    col = str(col).strip()
    col = col.replace("\n", " ")
    col = col.replace("\r", " ")
    col = " ".join(col.split())
    return col


def clean_upper(text):
    return clean_col(text).upper()


def resolve_sheet_name(xls_or_sheet_names):
    """Cari nama sheet sebenar walaupun ada ruang depan/belakang atau tambahan Q1."""
    sheet_names = (
        xls_or_sheet_names.sheet_names
        if hasattr(xls_or_sheet_names, "sheet_names")
        else list(xls_or_sheet_names)
    )

    # Cuba padanan tepat dahulu.
    if SHEET_NAME in sheet_names:
        return SHEET_NAME

    target = SHEET_NAME.strip().upper()

    # Cuba padanan selepas buang ruang depan/belakang.
    for sheet in sheet_names:
        if str(sheet).strip().upper() == target:
            return sheet

    # Cuba padanan yang mengandungi nama asas, contoh DATA DASHBOARD Q1.
    for sheet in sheet_names:
        sheet_clean = str(sheet).strip().upper()
        if target in sheet_clean or sheet_clean in target:
            return sheet

    # Fallback khusus untuk fail Q1.
    for sheet in sheet_names:
        if "DATA DASHBOARD" in str(sheet).strip().upper():
            return sheet

    return None


def to_number(series):
    return pd.to_numeric(
        series.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("%", "", regex=False)
        .str.strip()
        .replace({
            "": pd.NA,
            "-": pd.NA,
            "nan": pd.NA,
            "None": pd.NA,
            "#DIV/0!": pd.NA,
            "#VALUE!": pd.NA
        }),
        errors="coerce"
    )


def resolve_sheet_name(xls, sheet_options=None):
    """Cari nama sheet sebenar secara fleksibel mengikut suku dipilih."""
    available = list(xls.sheet_names)
    wanted_sheets = sheet_options or ACTIVE_SHEET_OPTIONS

    for wanted in wanted_sheets:
        for sheet in available:
            if sheet == wanted:
                return sheet

    for wanted in wanted_sheets:
        wanted_clean = str(wanted).strip().upper()
        for sheet in available:
            if str(sheet).strip().upper() == wanted_clean:
                return sheet

    return None


def find_header_row(uploaded_file, sheet_name):
    raw = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None)

    for i in range(min(50, len(raw))):
        row_text = " ".join(
            clean_upper(x)
            for x in raw.iloc[i].tolist()
            if pd.notna(x)
        )

        if "SEKTOR" in row_text and "BAHAGIAN" in row_text:
            return i

    return 0

def find_col(df, possible_names):
    for col in df.columns:
        col_upper = clean_upper(col)

        for name in possible_names:
            if name.upper() in col_upper:
                return col

    return None


def find_col_exact(df, possible_names):
    """Cari kolum dengan padanan nama tepat selepas pembersihan teks."""
    wanted = {clean_upper(name) for name in possible_names}

    for col in df.columns:
        if clean_upper(col) in wanted:
            return col

    return None


def safe_key(text):
    text = str(text)
    text = re.sub(r"[^A-Za-z0-9_]+", "_", text)
    text = text.strip("_")
    return text[:80] if text else "item"


def sidebar_pill(label, options, key, n_cols=None, use_expander=False, expanded=False):
    """
    Custom multi-select pill.
    Semua item yang dipilih akan kekal:
    - warna biru menyala
    - tanda ✓
    - state disimpan dalam st.session_state
    """
    options = sorted([
        str(x).strip()
        for x in options
        if pd.notna(x) and str(x).strip() != ""
    ])

    state_key = f"{key}_selected"

    if state_key not in st.session_state:
        st.session_state[state_key] = []

    selected_values = st.session_state[state_key]

    if n_cols is None:
        max_cols = 3 if key in ["filter_bahagian", "filter_kod_program"] else 1
    else:
        max_cols = n_cols

    def render_pills():
        for i in range(0, len(options), max_cols):
            cols = st.columns(max_cols) if use_expander else st.sidebar.columns(max_cols)

            for j, option in enumerate(options[i:i + max_cols]):
                pill_key = f"{key}_{safe_key(option)}"
                is_selected = option in selected_values
                label_text = f"✓ {option}" if is_selected else option
                button_type = "primary" if is_selected else "secondary"

                with cols[j]:
                    if st.button(label_text, key=pill_key, type=button_type):
                        if option in st.session_state[state_key]:
                            st.session_state[state_key].remove(option)
                        else:
                            st.session_state[state_key].append(option)
                        st.rerun()

    if use_expander:
        with st.sidebar.expander(label, expanded=expanded):
            render_pills()
    else:
        st.sidebar.markdown(f"**{label}**")
        render_pills()

    return st.session_state[state_key]


def detect_status_khas(row):
    """Kesan status khas mengikut suku tahun yang sedang dipaparkan."""
    row_texts = [
        clean_upper(value)
        for value in row.tolist()
        if pd.notna(value) and str(value).strip() != ""
    ]
    combined_text = " | ".join(row_texts)

    tidak_patterns = [
        "TIDAK DILAKSANAKAN",
        "TIDAK AKAN DILAKSANAKAN",
        "TIDAK DILAKSANA",
        "TIDAK AKAN DILAKSANA",
        "TAK DILAKSANAKAN",
        "TAK AKAN DILAKSANAKAN",
        "TAK DILAKSANA",
        "TAK AKAN DILAKSANA",
        "GUGUR",
    ]

    # Sheet Q2 mempunyai kolum khusus Bil. Program Gugur.
    if ACTIVE_QUARTER == "Q2":
        gugur_col = find_col_exact(
            pd.DataFrame(columns=row.index),
            ["BIL. PROGRAM GUGUR", "BIL PROGRAM GUGUR"]
        )
        if gugur_col is not None:
            gugur_value = pd.to_numeric(row.get(gugur_col), errors="coerce")
            if pd.notna(gugur_value) and gugur_value > 0:
                return "TIDAK DILAKSANAKAN"

    if any(pattern in combined_text for pattern in tidak_patterns):
        return "TIDAK DILAKSANAKAN"

    if ACTIVE_QUARTER == "Q1":
        try:
            status_text = clean_upper(row.iloc[STATUS_TEXT_COL_INDEX])
        except Exception:
            status_text = ""

        if "BERMULA Q2" in status_text or "MULA Q2" in status_text or status_text == "Q2":
            return "BERMULA Q2"
        if "BERMULA Q3" in status_text or "MULA Q3" in status_text or status_text == "Q3":
            return "BERMULA Q3"
        if "BERMULA Q4" in status_text or "MULA Q4" in status_text or status_text == "Q4":
            return "BERMULA Q4"

        try:
            weightage_value = pd.to_numeric(
                str(row.get(weightage_col, "")).replace(",", "").replace("%", "").strip(),
                errors="coerce"
            )
            pencapaian_value = pd.to_numeric(
                str(row.get(pencapaian_col, "")).replace(",", "").replace("%", "").strip(),
                errors="coerce"
            )
        except Exception:
            weightage_value = pd.NA
            pencapaian_value = pd.NA

        if pd.isna(weightage_value) and pd.isna(pencapaian_value):
            return "BERMULA Q2"

        return ""

    # Q2: Q3 dan Q4 MESTI dirujuk pada kolum CATATAN (JUSTIFIKASI) Q2 sahaja.
    # Jangan cari pada seluruh baris kerana teks Q3/Q4 di kolum lain boleh menyebabkan
    # rekod tersalah klasifikasi.
    try:
        # Guna nama kolum sebenar, bukan nombor index.
        # Index berubah apabila kolum kosong dibuang oleh load_data().
        q2_justifikasi_text = clean_upper(row.get(Q2_JUSTIFIKASI_COL, ""))
    except Exception:
        q2_justifikasi_text = ""

    if re.search(r"\bBERMULA\s*Q3\b|\bMULA\s*Q3\b", q2_justifikasi_text):
        return "BERMULA Q3"
    if re.search(r"\bBERMULA\s*Q4\b|\bMULA\s*Q4\b", q2_justifikasi_text):
        return "BERMULA Q4"

    return ""

def traffic_light_status(value):
    # Traffic light berdasarkan KPI Pencapaian = Prestasi / Sasaran x 100
    if pd.isna(value):
        return ""
    if value >= 85:
        return "🟢 Hijau"
    if value >= 60:
        return "🟡 Kuning"
    return "🔴 Merah"


def highlight_traffic_light(row):
    """
    Warnakan baris Senarai Status Prestasi ikut kolum TRAFFIC_LIGHT.
    Nota: display_df hanya papar kolum utama, jadi styling mesti rujuk
    nilai TRAFFIC_LIGHT yang memang dipaparkan dalam jadual.
    """
    traffic = str(row.get("TRAFFIC_LIGHT", ""))

    if "Hijau" in traffic:
        return ["background-color: #d9ead3"] * len(row)

    if "Kuning" in traffic:
        return ["background-color: #fff2cc"] * len(row)

    if "Merah" in traffic:
        return ["background-color: #f4cccc"] * len(row)

    if "Q2" in traffic or "Q3" in traffic or "Q4" in traffic:
        return ["background-color: #d9d2e9"] * len(row)

    if "Gugur" in traffic:
        return ["background-color: #e7e6e6"] * len(row)

    return [""] * len(row)


@st.cache_data
def load_data(uploaded_file, sheet_options):
    xls = pd.ExcelFile(uploaded_file)
    actual_sheet = resolve_sheet_name(xls, sheet_options)

    if actual_sheet is None:
        return None, xls.sheet_names, None

    header_row = find_header_row(uploaded_file, actual_sheet)

    df = pd.read_excel(
        uploaded_file,
        sheet_name=actual_sheet,
        header=header_row
    )

    df = df.dropna(axis=1, how="all")
    df = df.dropna(how="all")

    df.columns = [
        f"KOLUM_{i}" if str(col).startswith("Unnamed") else clean_col(col)
        for i, col in enumerate(df.columns)
    ]

    return df, xls.sheet_names, header_row




def render_selected_list(df_list, sektor_col, bahagian_col, program_col, pencapaian_col, title):
    st.markdown(f"### SENARAI STATUS PRESTASI - {title.upper()}")

    if df_list.empty:
        st.warning("Tiada rekod untuk kategori ini.")
        return

    # Papar kolum utama sahaja sehingga TRAFFIC_LIGHT.
    # Kolum tambahan selepas TRAFFIC_LIGHT tidak akan dipaparkan.
    display_cols = [
        sektor_col,
        bahagian_col,
        program_col,
        pencapaian_fizikal_col,
        weightage_col,
        pencapaian_col,
        "KPI_PENCAPAIAN_NUM",
        "TRAFFIC_LIGHT"
    ]

    display_cols = [
        c for c in display_cols
        if c is not None and c in df_list.columns
    ]

    display_df = df_list[display_cols].copy()

    if "KPI_PENCAPAIAN_NUM" in display_df.columns:
        display_df = display_df.rename(
            columns={"KPI_PENCAPAIAN_NUM": "PENCAPAIAN KPI (%)"}
        )

    st.dataframe(
        display_df.style.apply(highlight_traffic_light, axis=1),
        use_container_width=True,
        hide_index=True
    )

    csv = display_df.to_csv(index=False).encode("utf-8-sig")

    st.download_button(
        label=f"⬇️ Download Senarai {title}",
        data=csv,
        file_name=f"senarai_{title.lower().replace(' ', '_')}.csv",
        mime="text/csv"
    )




def build_summary_program_bahagian(source_df, bahagian_col):
    """Bina ringkasan bilangan program dan % pencapaian mengikut bahagian."""
    if source_df.empty:
        return pd.DataFrame()

    summary = (
        source_df
        .groupby([bahagian_col, "KATEGORI_TRAFFIC"], dropna=False)
        .size()
        .unstack(fill_value=0)
        .reset_index()
    )

    # Kolum Q2 dibuang daripada summary dan digantikan dengan % PENCAPAIAN.
    status_columns = ["Hijau", "Kuning", "Merah", "Q3", "Q4", "Gugur"]

    for col in status_columns:
        if col not in summary.columns:
            summary[col] = 0

    summary = summary[[bahagian_col] + status_columns].copy()

    # Suku Kedua: % PENCAPAIAN mesti merujuk terus kepada Column AD
    # iaitu KPI_PENCAPAIAN_NUM, bukan jumlah Pencapaian / jumlah Sasaran.
    # Kaedah lama menyebabkan nilai melebihi 100% lalu semuanya diclip kepada 100%.
    if ACTIVE_QUARTER == "Q2":
        valid_pencapaian = source_df[
            (source_df["STATUS_KHAS"] == "")
            & source_df["KPI_PENCAPAIAN_NUM"].notna()
        ].copy()

        if not valid_pencapaian.empty:
            pencapaian_bahagian = (
                valid_pencapaian
                .groupby(bahagian_col, dropna=False)["KPI_PENCAPAIAN_NUM"]
                .mean()
                .reset_index(name="% PENCAPAIAN")
            )
            summary = summary.merge(
                pencapaian_bahagian,
                on=bahagian_col,
                how="left"
            )
            total_pencapaian = float(
                valid_pencapaian["KPI_PENCAPAIAN_NUM"].mean()
            )
        else:
            summary["% PENCAPAIAN"] = 0.0
            total_pencapaian = 0.0
    else:
        valid_pencapaian = source_df[
            source_df["WEIGHTAGE_L_NUM"].notna()
            & (source_df["WEIGHTAGE_L_NUM"] > 0)
            & source_df["PENCAPAIAN_M_NUM"].notna()
        ].copy()

        if not valid_pencapaian.empty:
            pencapaian_bahagian = (
                valid_pencapaian
                .groupby(bahagian_col, dropna=False)
                .agg(
                    JUMLAH_SASARAN=("WEIGHTAGE_L_NUM", "sum"),
                    JUMLAH_PRESTASI=("PENCAPAIAN_M_NUM", "sum")
                )
                .reset_index()
            )
            pencapaian_bahagian["% PENCAPAIAN"] = (
                pencapaian_bahagian["JUMLAH_PRESTASI"]
                / pencapaian_bahagian["JUMLAH_SASARAN"]
                * 100
            )
            summary = summary.merge(
                pencapaian_bahagian[[bahagian_col, "% PENCAPAIAN"]],
                on=bahagian_col,
                how="left"
            )
            total_pencapaian = float(
                valid_pencapaian["PENCAPAIAN_M_NUM"].sum()
                / valid_pencapaian["WEIGHTAGE_L_NUM"].sum()
                * 100
            )
        else:
            summary["% PENCAPAIAN"] = 0.0
            total_pencapaian = 0.0

    summary["% PENCAPAIAN"] = (
        pd.to_numeric(summary["% PENCAPAIAN"], errors="coerce")
        .fillna(0.0)
    )

    summary = summary.rename(columns={bahagian_col: "BAHAGIAN"})
    summary["JUMLAH"] = summary[status_columns].sum(axis=1)

    total_row = {
        "BAHAGIAN": "JUMLAH KESELURUHAN",
        "Hijau": int(summary["Hijau"].sum()),
        "Kuning": int(summary["Kuning"].sum()),
        "Merah": int(summary["Merah"].sum()),
        "Q3": int(summary["Q3"].sum()),
        "Q4": int(summary["Q4"].sum()),
        "Gugur": int(summary["Gugur"].sum()),
        "% PENCAPAIAN": float(total_pencapaian),
        "JUMLAH": int(summary["JUMLAH"].sum()),
    }

    summary = summary.sort_values(
        by=["JUMLAH", "BAHAGIAN"],
        ascending=[False, True]
    ).reset_index(drop=True)

    summary = pd.concat(
        [summary, pd.DataFrame([total_row])],
        ignore_index=True
    )

    return summary

def highlight_summary_table(row):
    """Warnakan baris jumlah keseluruhan dalam jadual ringkasan."""
    if str(row.get("BAHAGIAN", "")) == "JUMLAH KESELURUHAN":
        return [
            "background-color: #dbeafe; font-weight: 900; color: #0f172a"
        ] * len(row)

    return [""] * len(row)


def build_bahagian_chart(filtered_df, bahagian_col, chart_height=None):
    """
    Bina carta stacked bar mengikut bahagian.
    Fungsi ini digunakan untuk paparan normal dan Paparan Carta Besar.
    """
    bahagian_status = (
        filtered_df
        .groupby([bahagian_col, "KATEGORI_TRAFFIC"], as_index=False)
        .size()
        .rename(columns={"size": "JUMLAH"})
    )

    status_order = ["Hijau", "Kuning", "Merah"]
    if ACTIVE_QUARTER == "Q1":
        status_order.append("Q2")
    status_order.extend(["Q3", "Q4", "Gugur"])

    bahagian_status["KATEGORI_TRAFFIC"] = pd.Categorical(
        bahagian_status["KATEGORI_TRAFFIC"],
        categories=status_order,
        ordered=True
    )

    bahagian_total = (
        bahagian_status
        .groupby(bahagian_col, as_index=False)["JUMLAH"]
        .sum()
        .rename(columns={"JUMLAH": "TOTAL"})
    )

    bahagian_status = bahagian_status.merge(
        bahagian_total,
        on=bahagian_col,
        how="left"
    )

    bahagian_status = bahagian_status.sort_values(
        ["TOTAL", bahagian_col, "KATEGORI_TRAFFIC"],
        ascending=[True, True, True]
    )

    if chart_height is None:
        chart_height = max(520, filtered_df[bahagian_col].nunique() * 42)

    fig_bahagian_stack = px.bar(
        bahagian_status,
        x="JUMLAH",
        y=bahagian_col,
        color="KATEGORI_TRAFFIC",
        orientation="h",
        text="JUMLAH",
        title="Bilangan Program Mengikut Status Prestasi Bagi Setiap Bahagian",
        category_orders={"KATEGORI_TRAFFIC": status_order},
        color_discrete_map={
            "Hijau": "#2fb463",
            "Kuning": "#f6c90e",
            "Merah": "#ef463b",
            "Q2": "#8e7cc3",
            "Q3": "#674ea7",
            "Q4": "#351c75",
            "Gugur": "#7a7788"
        }
    )

    fig_bahagian_stack.update_traces(
        textposition="inside",
        insidetextanchor="middle"
    )

    fig_bahagian_stack.update_layout(
        barmode="stack",
        xaxis_title="Jumlah Program",
        yaxis_title="Bahagian",
        legend_title="Status",
        height=chart_height,
        uniformtext_minsize=10,
        uniformtext_mode="hide",
        margin=dict(l=70, r=40, t=80, b=60)
    )

    return fig_bahagian_stack



def get_chart_source_df(filtered_df):
    """
    Carta ikut pilihan Traffic Light.
    Jika klik nilai Traffic Light, carta bahagian turut ditapis mengikut status tersebut.
    """
    selected = st.session_state.get("selected_traffic")

    if selected is None or selected == "Semua":
        return filtered_df.copy()

    status_map = {
        "Hijau": "Hijau",
        "Kuning": "Kuning",
        "Merah": "Merah",
        "Q2": "Q2",
        "Bermula Q2": "Q2",
        "Q3": "Q3",
        "Bermula Q3": "Q3",
        "Q4": "Q4",
        "Bermula Q4": "Q4",
        "Gugur": "Gugur"
    }

    selected_status = status_map.get(selected)

    if selected_status and "KATEGORI_TRAFFIC" in filtered_df.columns:
        return filtered_df[filtered_df["KATEGORI_TRAFFIC"] == selected_status].copy()

    return filtered_df.copy()


# =====================================================
# SIDEBAR FILTER ONLY
# =====================================================


# =====================================================
# TAB SUKU TAHUN + MAIN TITLE
# =====================================================
quarter_tab = st.radio(
    "Pilih Suku Tahun",
    options=["Suku Pertama", "Suku Kedua"],
    index=1,
    horizontal=True,
    label_visibility="collapsed",
    key="quarter_tab_selector"
)

ACTIVE_QUARTER_LABEL = quarter_tab
ACTIVE_QUARTER = QUARTER_CONFIG[quarter_tab]["code"]
ACTIVE_SHEET_OPTIONS = QUARTER_CONFIG[quarter_tab]["sheet_options"]
ACTIVE_SASARAN_PANEL = QUARTER_CONFIG[quarter_tab]["sasaran_panel"]

# Bersihkan pilihan status apabila pengguna bertukar suku tahun.
if st.session_state.get("last_quarter_tab") != ACTIVE_QUARTER:
    st.session_state["last_quarter_tab"] = ACTIVE_QUARTER
    st.session_state["selected_traffic"] = "Semua"
    st.session_state["focus_page"] = None

st.title(QUARTER_CONFIG[quarter_tab]["title"])

if not EXCEL_PATH.exists():
    st.sidebar.error("Fail Excel tidak dijumpai.")
    st.error("Fail Excel tidak dijumpai di folder yang ditetapkan.")

    st.write("Folder yang digunakan:")
    st.code(str(DATA_FOLDER))

    st.write("Nama fail yang dicari:")
    st.code(EXCEL_FILENAME)

    if DATA_FOLDER.exists():
        st.write("Fail Excel yang ada dalam folder ini:")
        excel_files = list(DATA_FOLDER.glob("*.xlsx"))
        if excel_files:
            for f in excel_files:
                st.write(f"- {f.name}")
        else:
            st.warning("Tiada fail .xlsx dijumpai dalam folder ini.")
    else:
        st.warning("Folder tidak wujud. Sila semak path folder.")

    st.stop()

uploaded_file = EXCEL_PATH


# =====================================================
# LOAD DATA
# =====================================================
df, sheet_list, header_row = load_data(uploaded_file, ACTIVE_SHEET_OPTIONS)

if df is None:
    st.error(f"Worksheet untuk **{quarter_tab}** tidak dijumpai dalam fail Excel.")
    st.write("Worksheet yang tersedia:")
    st.write(sheet_list)
    st.stop()


# =====================================================
# DETECT IMPORTANT COLUMNS
# =====================================================
sektor_col = find_col(df, ["SEKTOR"])
bahagian_col = find_col(df, ["BAHAGIAN"])
program_col = find_col(df, ["DESKRIPSI PROGRAM", "NAMA PROGRAM"])

if program_col is None:
    program_col = find_col(df, ["PROGRAM", "AKTIVITI", "TAJUK"])

if sektor_col is None or bahagian_col is None:
    st.error("Kolum **SEKTOR** atau **BAHAGIAN** tidak dijumpai dalam sheet DATA DASHBOARD.")
    st.write("Kolum yang dibaca:")
    st.write(list(df.columns))
    st.stop()

if program_col is None:
    program_col = df.columns[0]

if ACTIVE_QUARTER == "Q1":
    weightage_col = find_col_exact(df, [
        "PERATUS SASARAN (WEIGHTAGE) Q1",
        "PERATUS SASARAN WEIGHTAGE Q1"
    ])
    pencapaian_col = find_col_exact(df, ["PERATUS PENCAPAIAN Q1"])
    pencapaian_fizikal_col = find_col_exact(df, ["DATA DARI BAHAGIAN"])

    # Fallback kepada kedudukan asal Q1 jika nama kolum berubah sedikit.
    if weightage_col is None and len(df.columns) > WEIGHTAGE_COL_INDEX:
        weightage_col = df.columns[WEIGHTAGE_COL_INDEX]
    if pencapaian_col is None and len(df.columns) > PENCAPAIAN_COL_INDEX:
        pencapaian_col = df.columns[PENCAPAIAN_COL_INDEX]
    if pencapaian_fizikal_col is None and len(df.columns) > PENCAPAIAN_FIZIKAL_COL_INDEX:
        pencapaian_fizikal_col = df.columns[PENCAPAIAN_FIZIKAL_COL_INDEX]
else:
    weightage_col = find_col_exact(df, [
        "SASARAN Q2 (KUMULATIF)",
        "SASARAN Q2 KUMULATIF"
    ])
    pencapaian_col = find_col_exact(df, [
        "PENCAPAIAN Q2 (KUMULATIF)",
        "PENCAPAIAN Q2 KUMULATIF"
    ])
    pencapaian_fizikal_col = find_col_exact(df, ["DATA DARI BAHAGIAN Q2"])

    # Traffic Light Suku Kedua merujuk terus Column AD.
    q2_kpi_col = find_col_exact(df, [
        "% PENCAPAIAN (100%)",
        "PENCAPAIAN (100%)",
        "% PENCAPAIAN 100%"
    ])

    # Kolum justifikasi khusus Q2. Selepas kolum kosong dibuang, kedudukan
    # index tidak lagi semestinya 30. Oleh itu, rujuk nama kolum sebenar.
    Q2_JUSTIFIKASI_COL = find_col_exact(df, [
        "CATATAN (JUSTIFIKASI).1",
        "CATATAN (JUSTIFIKASI) Q2",
        "JUSTIFIKASI Q2"
    ])

if weightage_col is None or pencapaian_col is None:
    st.error(
        f"Kolum sasaran atau pencapaian untuk {quarter_tab} tidak dijumpai dalam "
        f"sheet {resolve_sheet_name(pd.ExcelFile(uploaded_file), ACTIVE_SHEET_OPTIONS)}."
    )
    st.write("Kolum yang dibaca:")
    st.write(list(df.columns))
    st.stop()

if ACTIVE_QUARTER == "Q2" and q2_kpi_col is None:
    st.error(
        "Kolum **% PENCAPAIAN (100%)** untuk Traffic Light Suku Kedua "
        "tidak dijumpai dalam sheet DATA DASHBOARD Q2 CLEAN."
    )
    st.write("Kolum yang dibaca:")
    st.write(list(df.columns))
    st.stop()

if ACTIVE_QUARTER == "Q2" and Q2_JUSTIFIKASI_COL is None:
    st.error(
        "Kolum **CATATAN (JUSTIFIKASI) Q2** tidak dijumpai. "
        "Kolum ini diperlukan untuk menentukan program Bermula Q3 dan Bermula Q4."
    )
    st.write("Kolum yang dibaca:")
    st.write(list(df.columns))
    st.stop()


# =====================================================
# CLEAN DATA
# =====================================================
df[sektor_col] = df[sektor_col].astype("string").str.strip()
df[bahagian_col] = df[bahagian_col].astype("string").str.strip()
df[program_col] = df[program_col].astype("string").str.strip()

df["STATUS_KHAS"] = df.apply(detect_status_khas, axis=1)

# Dalam Suku Kedua, program yang bermula Q2 telah masuk tempoh penilaian.
# Oleh itu, status BERMULA Q2 ditukar kepada status biasa untuk dinilai dalam Traffic Light.
if ACTIVE_QUARTER == "Q2":
    df.loc[df["STATUS_KHAS"] == "BERMULA Q2", "STATUS_KHAS"] = ""

df["WEIGHTAGE_L_NUM"] = to_number(df[weightage_col])
df["PENCAPAIAN_M_NUM"] = to_number(df[pencapaian_col])

# KPI PENCAPAIAN UNTUK TRAFFIC LIGHT:
# - Suku Pertama kekal seperti asal: PRESTASI / SASARAN x 100.
# - Suku Kedua merujuk terus Column AD: % PENCAPAIAN (100%).
if ACTIVE_QUARTER == "Q1":
    df["KPI_PENCAPAIAN_NUM"] = 0.0

    valid_kpi_mask = (
        df["WEIGHTAGE_L_NUM"].notna()
        & (df["WEIGHTAGE_L_NUM"] > 0)
        & df["PENCAPAIAN_M_NUM"].notna()
    )

    df.loc[valid_kpi_mask, "KPI_PENCAPAIAN_NUM"] = (
        df.loc[valid_kpi_mask, "PENCAPAIAN_M_NUM"]
        / df.loc[valid_kpi_mask, "WEIGHTAGE_L_NUM"]
    ) * 100
else:
    # Nilai kosong dalam Column AD dianggap 0% supaya program aktif
    # tidak tercicir daripada kategori Traffic Light.
    df["KPI_PENCAPAIAN_NUM"] = to_number(df[q2_kpi_col])

df = df[
    df[sektor_col].notna()
    & df[bahagian_col].notna()
    & (df[sektor_col].astype(str).str.strip() != "")
    & (df[bahagian_col].astype(str).str.strip() != "")
].copy()

jumlah_asal = len(df)


# =====================================================
# RULES
# =====================================================
df_bermula_q2 = df[df["STATUS_KHAS"] == "BERMULA Q2"].copy()
df_bermula_q3 = df[df["STATUS_KHAS"] == "BERMULA Q3"].copy()
df_bermula_q4 = df[df["STATUS_KHAS"] == "BERMULA Q4"].copy()
df_tidak_dilaksanakan = df[df["STATUS_KHAS"] == "TIDAK DILAKSANAKAN"].copy()

if ACTIVE_QUARTER == "Q1":
    # Suku Pertama kekalkan syarat penilaian asal.
    df_dinilai = df[
        (df["STATUS_KHAS"] == "")
        & df["WEIGHTAGE_L_NUM"].notna()
        & (df["WEIGHTAGE_L_NUM"] > 0)
    ].copy()
else:
    # Suku Kedua: semua program aktif dinilai berdasarkan Column AD.
    # Q3, Q4 dan Gugur telah diasingkan melalui STATUS_KHAS.
    df_dinilai = df[
        df["STATUS_KHAS"] == ""
    ].copy()

jumlah_bermula_q2_asal = len(df_bermula_q2)
jumlah_bermula_q3_asal = len(df_bermula_q3)
jumlah_bermula_q4_asal = len(df_bermula_q4)
jumlah_tidak_dilaksanakan_asal = len(df_tidak_dilaksanakan)
jumlah_dinilai = len(df_dinilai)
jumlah_ignore_weightage = (
    jumlah_asal
    - jumlah_bermula_q2_asal
    - jumlah_bermula_q3_asal
    - jumlah_bermula_q4_asal
    - jumlah_tidak_dilaksanakan_asal
    - jumlah_dinilai
)

df_dinilai["TRAFFIC_LIGHT"] = df_dinilai["KPI_PENCAPAIAN_NUM"].apply(traffic_light_status)
df_bermula_q2["TRAFFIC_LIGHT"] = "🟣 Q2"
df_bermula_q3["TRAFFIC_LIGHT"] = "🟣 Q3"
df_bermula_q4["TRAFFIC_LIGHT"] = "🟣 Q4"
df_tidak_dilaksanakan["TRAFFIC_LIGHT"] = "⚫ Gugur"

df_paparan = pd.concat(
    [
        df_dinilai,
        df_bermula_q2,
        df_bermula_q3,
        df_bermula_q4,
        df_tidak_dilaksanakan
    ],
    ignore_index=True
)


# =====================================================
# SIDEBAR FILTER
# =====================================================
selected_sektor = sidebar_pill(
    "Pilih Sektor",
    df_paparan[sektor_col].dropna().unique(),
    f"filter_sektor_{ACTIVE_QUARTER.lower()}",
    n_cols=1,
    use_expander=True,
    expanded=True
)

if selected_sektor:
    filtered_df = df_paparan[df_paparan[sektor_col].isin(selected_sektor)].copy()
else:
    filtered_df = df_paparan.copy()

selected_bahagian = sidebar_pill(
    "Pilih Bahagian",
    filtered_df[bahagian_col].dropna().unique(),
    f"filter_bahagian_{ACTIVE_QUARTER.lower()}",
    n_cols=2,
    use_expander=True,
    expanded=False
)

if selected_bahagian:
    filtered_df = filtered_df[filtered_df[bahagian_col].isin(selected_bahagian)].copy()


# =====================================================
# FILTER KOD PROGRAM (SHORT FORM)
# Contoh: PI.N0.NA.10110003.01 -> NA
# =====================================================
kod_program_col = find_col(df_paparan, ["KOD PROGRAM", "KOD"])
selected_kod = []

if kod_program_col is not None:
    # Wajib cipta kolum short form dalam kedua-dua dataframe:
    # df_paparan untuk pilihan filter, filtered_df untuk tapisan sebenar.
    df_paparan["KOD_PROGRAM_SHORT"] = (
        df_paparan[kod_program_col]
        .astype(str)
        .str.extract(r"PI\.[^.]+\.([^.]+)\.", expand=False)
        .fillna("LAIN-LAIN")
        .str.strip()
    )

    if kod_program_col in filtered_df.columns:
        filtered_df["KOD_PROGRAM_SHORT"] = (
            filtered_df[kod_program_col]
            .astype(str)
            .str.extract(r"PI\.[^.]+\.([^.]+)\.", expand=False)
            .fillna("LAIN-LAIN")
            .str.strip()
        )

    selected_kod = sidebar_pill(
        "Pilih Kod Program",
        df_paparan["KOD_PROGRAM_SHORT"].dropna().unique(),
        f"filter_kod_program_{ACTIVE_QUARTER.lower()}",
        n_cols=4,
        use_expander=True,
        expanded=False
    )

    if selected_kod and "KOD_PROGRAM_SHORT" in filtered_df.columns:
        filtered_df = filtered_df[
            filtered_df["KOD_PROGRAM_SHORT"].isin(selected_kod)
        ].copy()


# =====================================================
# RESET FILTER - LETAK DI BAWAH SIDEBAR
# =====================================================
st.sidebar.markdown("---")

if st.sidebar.button(
    "🔄 Reset Semua Filter",
    use_container_width=True,
    key="reset_filter_btn"
):
    st.session_state[f"filter_sektor_{ACTIVE_QUARTER.lower()}_selected"] = []
    st.session_state[f"filter_bahagian_{ACTIVE_QUARTER.lower()}_selected"] = []
    st.session_state[f"filter_kod_program_{ACTIVE_QUARTER.lower()}_selected"] = []

    if "selected_traffic" in st.session_state:
        st.session_state.selected_traffic = "Semua"

    st.rerun()

if st.sidebar.button(
    "🔃 Refresh Data Excel",
    use_container_width=True,
    key="refresh_excel_btn"
):
    st.cache_data.clear()
    st.session_state[f"filter_sektor_{ACTIVE_QUARTER.lower()}_selected"] = []
    st.session_state[f"filter_bahagian_{ACTIVE_QUARTER.lower()}_selected"] = []
    st.session_state[f"filter_kod_program_{ACTIVE_QUARTER.lower()}_selected"] = []

    if "selected_traffic" in st.session_state:
        st.session_state.selected_traffic = "Semua"

    st.rerun()

st.sidebar.markdown("---")

if st.sidebar.button(
    "🔒 Log Keluar",
    use_container_width=True,
    key="logout_dashboard_btn"
):
    st.session_state[AUTH_SESSION_KEY] = False
    st.rerun()

st.sidebar.markdown("---")

if filtered_df.empty:
    st.warning("Tiada data berdasarkan filter yang dipilih.")
    st.stop()

filtered_dinilai = filtered_df[filtered_df["STATUS_KHAS"] == ""].copy()

# Debug ringkas untuk semakan jika nilai panel tidak keluar.
# Boleh comment / padam selepas dashboard disahkan betul.
# st.write("Jumlah rekod selepas filter:", len(filtered_df))
# st.write("Jumlah rekod dinilai:", len(filtered_dinilai))

# Sidebar hanya paparkan pilihan filter sahaja.


# =====================================================
# MAIN PAGE - TRAFFIC LIGHT
# =====================================================
# Q1 kekal menggunakan rekod yang sedang dinilai seperti logik asal.
# Q2 mengira Traffic Light terus daripada Column AD (% PENCAPAIAN (100%)),
# tetapi Q3, Q4 dan Gugur dikeluarkan daripada warna supaya setiap program
# hanya mempunyai SATU status.
traffic_source_df = filtered_dinilai[
    filtered_dinilai["KPI_PENCAPAIAN_NUM"].notna()
].copy()

df_hijau = traffic_source_df[
    traffic_source_df["KPI_PENCAPAIAN_NUM"] >= 85
].copy()

df_kuning = traffic_source_df[
    (traffic_source_df["KPI_PENCAPAIAN_NUM"] >= 60)
    & (traffic_source_df["KPI_PENCAPAIAN_NUM"] < 85)
].copy()

df_merah = traffic_source_df[
    traffic_source_df["KPI_PENCAPAIAN_NUM"] < 60
].copy()

df_bermula_q2_filtered = filtered_df[filtered_df["STATUS_KHAS"] == "BERMULA Q2"].copy()
df_bermula_q3_filtered = filtered_df[filtered_df["STATUS_KHAS"] == "BERMULA Q3"].copy()
df_bermula_q4_filtered = filtered_df[filtered_df["STATUS_KHAS"] == "BERMULA Q4"].copy()
df_tidak_filtered = filtered_df[filtered_df["STATUS_KHAS"] == "TIDAK DILAKSANAKAN"].copy()

hijau = len(df_hijau)
kuning = len(df_kuning)
merah = len(df_merah)
bermula_q2 = len(df_bermula_q2_filtered)
bermula_q3 = len(df_bermula_q3_filtered)
bermula_q4 = len(df_bermula_q4_filtered)
tidak_dilaksanakan = len(df_tidak_filtered)

# Panel ringkasan:
# BACA PADA SHEET DATA DASHBOARD SAHAJA.
# Kiraan berubah ikut filter Sektor / Bahagian.
#
# Formula:
# - Column L = WEIGHTAGE / SASARAN Q1
# - Column M = % PENCAPAIAN Q1
# - SASARAN % = jumlah Column L / jumlah Column L x 100 = 100%
# - PRESTASI % = jumlah Column M / jumlah Column L x 100
# - PENCAPAIAN % = PRESTASI % / SASARAN % x 100
#
# Contoh:
# SASARAN = 20,700 / 20,700 x 100 = 100.00%
# PRESTASI = 19,587 / 20,700 x 100 = 94.62%
# PENCAPAIAN = 94.62% / 100.00% x 100 = 94.62%

panel_df = filtered_df.copy()

# Jumlah Program mesti dikira daripada SEMUA rekod program dalam sheet,
# termasuk rekod yang belum mempunyai weightage/pencapaian.
# Sebelum ini kiraan menggunakan filtered_df (df_paparan), menyebabkan
# 6 program tanpa weightage tidak dimasukkan: 263 - 6 - 11 = 246.
# Kiraan betul Q2 ialah 263 - 11 Gugur = 252.
panel_count_df = df.copy()

if selected_sektor:
    panel_count_df = panel_count_df[
        panel_count_df[sektor_col].isin(selected_sektor)
    ].copy()

if selected_bahagian:
    panel_count_df = panel_count_df[
        panel_count_df[bahagian_col].isin(selected_bahagian)
    ].copy()

if selected_kod and kod_program_col is not None and kod_program_col in panel_count_df.columns:
    panel_count_df["KOD_PROGRAM_SHORT"] = (
        panel_count_df[kod_program_col]
        .astype(str)
        .str.extract(r"PI\.[^.]+\.([^.]+)\.", expand=False)
        .fillna("LAIN-LAIN")
        .str.strip()
    )
    panel_count_df = panel_count_df[
        panel_count_df["KOD_PROGRAM_SHORT"].isin(selected_kod)
    ].copy()

jumlah_program_panel = len(
    panel_count_df[
        panel_count_df["STATUS_KHAS"] != "TIDAK DILAKSANAKAN"
    ]
)

panel_df["SASARAN_PANEL_NUM"] = (
    panel_df["WEIGHTAGE_L_NUM"]
    .clip(lower=0, upper=100)
)

panel_df["PRESTASI_PANEL_NUM"] = (
    panel_df["PENCAPAIAN_M_NUM"]
    .clip(lower=0, upper=100)
)

# Kiraan panel mengikut suku dipilih.
# Q1 kekal menggunakan formula asal jumlah pencapaian / jumlah sasaran.
# Q2 mesti sama dengan % PENCAPAIAN dalam Summary:
#   - ambil terus Column AD (% PENCAPAIAN (100%));
#   - hanya program aktif/status biasa dimasukkan;
#   - Q3, Q4 dan Gugur tidak dimasukkan;
#   - PRESTASI = % PENCAPAIAN x sasaran suku (50%).
if ACTIVE_QUARTER == "Q2":
    panel_valid = panel_df[
        (panel_df["STATUS_KHAS"] == "")
        & panel_df["KPI_PENCAPAIAN_NUM"].notna()
    ].copy()

    sasaran_panel = ACTIVE_SASARAN_PANEL

    if not panel_valid.empty:
        pencapaian_panel = float(
            panel_valid["KPI_PENCAPAIAN_NUM"].mean()
        )
        prestasi_panel = (
            pencapaian_panel / 100
        ) * sasaran_panel
    else:
        prestasi_panel = 0.0
        pencapaian_panel = 0.0
else:
    panel_valid = panel_df[
        panel_df["WEIGHTAGE_L_NUM"].notna()
        & (panel_df["WEIGHTAGE_L_NUM"] > 0)
        & panel_df["PENCAPAIAN_M_NUM"].notna()
    ].copy()

    if not panel_valid.empty:
        sasaran_panel = ACTIVE_SASARAN_PANEL

        jumlah_l = panel_valid["WEIGHTAGE_L_NUM"].sum()
        jumlah_m = panel_valid["PENCAPAIAN_M_NUM"].sum()

        if jumlah_l > 0:
            prestasi_panel = (jumlah_m / jumlah_l) * sasaran_panel
            prestasi_panel = min(prestasi_panel, sasaran_panel)
            pencapaian_panel = (prestasi_panel / sasaran_panel) * 100
        else:
            prestasi_panel = 0.0
            pencapaian_panel = 0.0
    else:
        sasaran_panel = 0.0
        prestasi_panel = 0.0
        pencapaian_panel = 0.0

# Elakkan paparan negatif sahaja. Nilai melebihi 100 dikekalkan jika memang ada dalam data.
pencapaian_panel = max(pencapaian_panel, 0.0)

# Kategori ringkas untuk stacked bar chart
filtered_df["KATEGORI_TRAFFIC"] = filtered_df["TRAFFIC_LIGHT"].replace({
    "🟢 Hijau": "Hijau",
    "🟡 Kuning": "Kuning",
    "🔴 Merah": "Merah",
    "🟣 Q2": "Q2",
    "🟣 Q3": "Q3",
    "🟣 Q4": "Q4",
    "⚫ Gugur": "Gugur"
})

if "selected_traffic" not in st.session_state:
    st.session_state.selected_traffic = "Semua"

chart_filtered_df = get_chart_source_df(filtered_df)

# =====================================================
# PAPARAN BESAR MINIMUM - ASAL DIKEKALKAN
# =====================================================
FOCUS_PAGES = ["traffic", "chart", "report"]
FOCUS_TITLES = {
    "traffic": "🚦 TRAFFIC LIGHT",
    "chart": "📊 PENCAPAIAN MENGIKUT BAHAGIAN",
    "report": "📋 LAPORAN STATUS PRESTASI",
}

if "focus_page" not in st.session_state:
    st.session_state.focus_page = None


def focus_previous():
    current = st.session_state.get("focus_page") or "traffic"
    index = FOCUS_PAGES.index(current)
    st.session_state.focus_page = FOCUS_PAGES[(index - 1) % len(FOCUS_PAGES)]


def focus_next():
    current = st.session_state.get("focus_page") or "traffic"
    index = FOCUS_PAGES.index(current)
    st.session_state.focus_page = FOCUS_PAGES[(index + 1) % len(FOCUS_PAGES)]


def focus_home():
    st.session_state.focus_page = None


if st.session_state.focus_page in FOCUS_PAGES:
    nav_left, nav_title, nav_right = st.columns([1, 10, 1])

    with nav_left:
        if st.button("‹", key="focus_previous_btn", use_container_width=True):
            focus_previous()
            st.rerun()

    with nav_title:
        if st.button(
            FOCUS_TITLES[st.session_state.focus_page],
            key="focus_home_btn",
            help="Kembali ke paparan utama",
            use_container_width=True
        ):
            focus_home()
            st.rerun()

    with nav_right:
        if st.button("›", key="focus_next_btn", use_container_width=True):
            focus_next()
            st.rerun()

    if st.session_state.focus_page == "chart":
        if chart_filtered_df.empty:
            st.warning("Tiada data carta untuk status yang dipilih.")
        else:
            fig_bahagian_besar = build_bahagian_chart(
                chart_filtered_df,
                bahagian_col,
                chart_height=max(760, chart_filtered_df[bahagian_col].nunique() * 58)
            )
            fig_bahagian_besar.update_layout(
                margin=dict(l=20, r=25, t=65, b=35),
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)"
            )
            st.plotly_chart(
                fig_bahagian_besar,
                use_container_width=True,
                key="bahagian_chart_focus",
                config={"displaylogo": False, "responsive": True, "displayModeBar": True}
            )

    elif st.session_state.focus_page == "traffic":
        focus_left, focus_right = st.columns([3.0, 1.05], gap="medium")

        with focus_left:
            st.markdown('<div class="traffic-container">', unsafe_allow_html=True)
            fc1, fc2, fc3 = st.columns(3)

            with fc1:
                st.markdown('<div class="static-traffic-card"><div class="static-range">≥ 85%</div></div>', unsafe_allow_html=True)
                if st.button(f"{hijau}", key="focus_btn_hijau_value"):
                    st.session_state.selected_traffic = "Hijau"
                    st.rerun()

            with fc2:
                st.markdown('<div class="static-traffic-card"><div class="static-range">60% - 84.99%</div></div>', unsafe_allow_html=True)
                if st.button(f"{kuning}", key="focus_btn_kuning_value"):
                    st.session_state.selected_traffic = "Kuning"
                    st.rerun()

            with fc3:
                st.markdown('<div class="static-traffic-card"><div class="static-range">< 60%</div></div>', unsafe_allow_html=True)
                if st.button(f"{merah}", key="focus_btn_merah_value"):
                    st.session_state.selected_traffic = "Merah"
                    st.rerun()

            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown('<div class="compact-status-row" style="margin-top:22px; margin-bottom:10px;">', unsafe_allow_html=True)

            if ACTIVE_QUARTER == "Q1":
                fq2_label, fq2_value, fsep1, fq3_label, fq3_value, fsep2, fq4_label, fq4_value, fsep3, fg_label, fg_value = st.columns(
                    [0.55, 0.42, 0.16, 0.55, 0.42, 0.16, 0.55, 0.42, 0.16, 1.05, 0.42]
                )
                with fq2_label:
                    st.markdown('<div class="compact-status-label">Q2-</div>', unsafe_allow_html=True)
                with fq2_value:
                    if st.button(str(bermula_q2), key="focus_btn_q2_status"):
                        st.session_state.selected_traffic = "Q2"
                        st.rerun()
                with fsep1:
                    st.markdown('<div class="compact-status-separator">|</div>', unsafe_allow_html=True)
            else:
                fq3_label, fq3_value, fsep2, fq4_label, fq4_value, fsep3, fg_label, fg_value = st.columns(
                    [0.55, 0.42, 0.16, 0.55, 0.42, 0.16, 1.05, 0.42]
                )

            with fq3_label:
                st.markdown('<div class="compact-status-label">Q3-</div>', unsafe_allow_html=True)
            with fq3_value:
                if st.button(str(bermula_q3), key="focus_btn_q3_status"):
                    st.session_state.selected_traffic = "Q3"
                    st.rerun()
            with fsep2:
                st.markdown('<div class="compact-status-separator">|</div>', unsafe_allow_html=True)
            with fq4_label:
                st.markdown('<div class="compact-status-label">Q4-</div>', unsafe_allow_html=True)
            with fq4_value:
                if st.button(str(bermula_q4), key="focus_btn_q4_status"):
                    st.session_state.selected_traffic = "Q4"
                    st.rerun()
            with fsep3:
                st.markdown('<div class="compact-status-separator">|</div>', unsafe_allow_html=True)
            with fg_label:
                st.markdown('<div class="compact-status-label">GUGUR-</div>', unsafe_allow_html=True)
            with fg_value:
                if st.button(str(tidak_dilaksanakan), key="focus_btn_gugur_status"):
                    st.session_state.selected_traffic = "Gugur"
                    st.rerun()

            st.markdown("</div>", unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

        with focus_right:
            st.markdown('<div class="summary-panel">', unsafe_allow_html=True)
            if st.button(f"{jumlah_program_panel:,.0f}", key="focus_btn_jumlah_program"):
                st.session_state.selected_traffic = "Semua"
                st.rerun()
            st.markdown(
                f"""
                    <div class="summary-label">Jumlah Program Aktif</div>
                    <div class="summary-line"></div>
                    <div class="summary-row">SASARAN <span>{sasaran_panel:.2f}%</span></div>
                    <div class="summary-row">PRESTASI <span>{prestasi_panel:.2f}%</span></div>
                    <div class="summary-line"></div>
                    <div class="summary-achievement">PENCAPAIAN {pencapaian_panel:.2f}%</div>
                </div>
                """,
                unsafe_allow_html=True
            )

    elif st.session_state.focus_page == "report":
        tab_senarai_focus, tab_summary_focus = st.tabs([
            "📋 SENARAI STATUS PRESTASI",
            "📑 SUMMARY PROGRAM MENGIKUT BAHAGIAN"
        ])

        with tab_senarai_focus:
            selected = st.session_state.selected_traffic
            if selected == "Semua":
                render_selected_list(filtered_df, sektor_col, bahagian_col, program_col, pencapaian_col, "Semua")
            elif selected == "Hijau":
                render_selected_list(df_hijau, sektor_col, bahagian_col, program_col, pencapaian_col, "Hijau")
            elif selected == "Kuning":
                render_selected_list(df_kuning, sektor_col, bahagian_col, program_col, pencapaian_col, "Kuning")
            elif selected == "Merah":
                render_selected_list(df_merah, sektor_col, bahagian_col, program_col, pencapaian_col, "Merah")
            elif selected in ["Bermula Q2", "Q2"]:
                render_selected_list(df_bermula_q2_filtered, sektor_col, bahagian_col, program_col, pencapaian_col, "Q2")
            elif selected in ["Bermula Q3", "Q3"]:
                render_selected_list(df_bermula_q3_filtered, sektor_col, bahagian_col, program_col, pencapaian_col, "Q3")
            elif selected in ["Bermula Q4", "Q4"]:
                render_selected_list(df_bermula_q4_filtered, sektor_col, bahagian_col, program_col, pencapaian_col, "Q4")
            elif selected == "Gugur":
                render_selected_list(df_tidak_filtered, sektor_col, bahagian_col, program_col, pencapaian_col, "Gugur")
            else:
                st.info("Klik Jumlah Program Aktif, bulatan Traffic Light, atau nilai Q2/Q3/Q4/Gugur untuk lihat senarai.")

        with tab_summary_focus:
            st.markdown("### SUMMARY PROGRAM MENGIKUT BAHAGIAN")
            st.caption("Ringkasan ini berubah mengikut filter Sektor, Bahagian dan Kod Program.")
            summary_df_focus = build_summary_program_bahagian(filtered_df, bahagian_col)
            if summary_df_focus.empty:
                st.warning("Tiada data untuk dipaparkan dalam summary.")
            else:
                summary_display_focus = summary_df_focus.rename(columns={
                    "Hijau": "🟢 HIJAU", "Kuning": "🟡 KUNING", "Merah": "🔴 MERAH",
                    "Q3": "Q3", "Q4": "Q4", "Gugur": "GUGUR",
                    "% PENCAPAIAN": "% PENCAPAIAN",
                    "JUMLAH": "JUMLAH PROGRAM"
                })
                integer_cols_focus = ["🟢 HIJAU", "🟡 KUNING", "🔴 MERAH", "Q3", "Q4", "GUGUR", "JUMLAH PROGRAM"]
                for col in integer_cols_focus:
                    if col in summary_display_focus.columns:
                        summary_display_focus[col] = pd.to_numeric(summary_display_focus[col], errors="coerce").fillna(0).astype(int)
                if "% PENCAPAIAN" in summary_display_focus.columns:
                    summary_display_focus["% PENCAPAIAN"] = pd.to_numeric(
                        summary_display_focus["% PENCAPAIAN"], errors="coerce"
                    ).fillna(0.0)
                focus_format = {
                    col: "{:,.0f}" for col in integer_cols_focus if col in summary_display_focus.columns
                }
                if "% PENCAPAIAN" in summary_display_focus.columns:
                    focus_format["% PENCAPAIAN"] = "{:.2f}%"
                st.dataframe(
                    summary_display_focus.style.apply(highlight_summary_table, axis=1).format(focus_format),
                    use_container_width=True,
                    hide_index=True,
                    height=min(720, 80 + (len(summary_display_focus) * 35))
                )

    st.stop()

traffic_focus_spacer, traffic_focus_btn = st.columns([18, 1])
with traffic_focus_btn:
    if st.button("⛶", key="btn_buka_traffic_focus", help="Buka paparan besar Traffic Light", use_container_width=True):
        st.session_state.focus_page = "traffic"
        st.rerun()

main_left, main_right = st.columns([3.0, 1.05], gap="medium")

with main_left:
    st.markdown('<div class="traffic-container">', unsafe_allow_html=True)

    c1, c2, c3 = st.columns(3)

    with c1:
        st.markdown(
            """
            <div class="static-traffic-card">
                <div class="static-range">≥ 85%</div>
            </div>
            """,
            unsafe_allow_html=True
        )

        if st.button(f"{hijau}", key="btn_hijau_value"):
            st.session_state.selected_traffic = "Hijau"
            if "selected_chart_bahagian" in st.session_state:
                st.session_state.selected_chart_bahagian = None
            st.rerun()

    with c2:
        st.markdown(
            """
            <div class="static-traffic-card">
                <div class="static-range">60% - 84.99%</div>
            </div>
            """,
            unsafe_allow_html=True
        )

        if st.button(f"{kuning}", key="btn_kuning_value"):
            st.session_state.selected_traffic = "Kuning"
            if "selected_chart_bahagian" in st.session_state:
                st.session_state.selected_chart_bahagian = None
            st.rerun()

    with c3:
        st.markdown(
            """
            <div class="static-traffic-card">
                <div class="static-range">< 60%</div>
            </div>
            """,
            unsafe_allow_html=True
        )

        if st.button(f"{merah}", key="btn_merah_value"):
            st.session_state.selected_traffic = "Merah"
            if "selected_chart_bahagian" in st.session_state:
                st.session_state.selected_chart_bahagian = None
            st.rerun()

    st.markdown("<br>", unsafe_allow_html=True)

    # Baris status khas: Suku Pertama papar Q2/Q3/Q4/Gugur; Suku Kedua papar Q3/Q4/Gugur sahaja.
    st.markdown('<div class="compact-status-row" style="margin-top:22px; margin-bottom:10px;">', unsafe_allow_html=True)

    if ACTIVE_QUARTER == "Q1":
        q2_label_col, q2_value_col, q_sep1, q3_label_col, q3_value_col, q_sep2, q4_label_col, q4_value_col, q_sep3, gugur_label_col, gugur_value_col = st.columns(
            [0.55, 0.42, 0.16, 0.55, 0.42, 0.16, 0.55, 0.42, 0.16, 1.05, 0.42]
        )
        with q2_label_col:
            st.markdown('<div class="compact-status-label">Q2-</div>', unsafe_allow_html=True)
        with q2_value_col:
            if st.button(str(bermula_q2), key="btn_q2_status"):
                st.session_state.selected_traffic = "Q2"
                if "selected_chart_bahagian" in st.session_state:
                    st.session_state.selected_chart_bahagian = None
                st.rerun()
        with q_sep1:
            st.markdown('<div class="compact-status-separator">|</div>', unsafe_allow_html=True)
    else:
        q3_label_col, q3_value_col, q_sep2, q4_label_col, q4_value_col, q_sep3, gugur_label_col, gugur_value_col = st.columns(
            [0.55, 0.42, 0.16, 0.55, 0.42, 0.16, 1.05, 0.42]
        )

    with q3_label_col:
        st.markdown('<div class="compact-status-label">Q3-</div>', unsafe_allow_html=True)
    with q3_value_col:
        if st.button(str(bermula_q3), key="btn_q3_status"):
            st.session_state.selected_traffic = "Q3"
            if "selected_chart_bahagian" in st.session_state:
                st.session_state.selected_chart_bahagian = None
            st.rerun()
    with q_sep2:
        st.markdown('<div class="compact-status-separator">|</div>', unsafe_allow_html=True)
    with q4_label_col:
        st.markdown('<div class="compact-status-label">Q4-</div>', unsafe_allow_html=True)
    with q4_value_col:
        if st.button(str(bermula_q4), key="btn_q4_status"):
            st.session_state.selected_traffic = "Q4"
            if "selected_chart_bahagian" in st.session_state:
                st.session_state.selected_chart_bahagian = None
            st.rerun()
    with q_sep3:
        st.markdown('<div class="compact-status-separator">|</div>', unsafe_allow_html=True)
    with gugur_label_col:
        st.markdown('<div class="compact-status-label">GUGUR-</div>', unsafe_allow_html=True)
    with gugur_value_col:
        if st.button(str(tidak_dilaksanakan), key="btn_tidak_status"):
            st.session_state.selected_traffic = "Gugur"
            if "selected_chart_bahagian" in st.session_state:
                st.session_state.selected_chart_bahagian = None
            st.rerun()

    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)

with main_right:
    st.markdown('<div class="summary-panel">', unsafe_allow_html=True)

    if st.button(f"{jumlah_program_panel:,.0f}", key="btn_jumlah_program"):
        st.session_state.selected_traffic = "Semua"
        if "selected_chart_bahagian" in st.session_state:
            st.session_state.selected_chart_bahagian = None
        st.rerun()

    st.markdown(
        f"""
            <div class="summary-label">Jumlah Program Aktif</div>
            <div class="summary-line"></div>
            <div class="summary-row">SASARAN <span>{sasaran_panel:.2f}%</span></div>
            <div class="summary-row">PRESTASI <span>{prestasi_panel:.2f}%</span></div>
            <div class="summary-line"></div>
            <div class="summary-achievement">PENCAPAIAN {pencapaian_panel:.2f}%</div>
        </div>
        """,
        unsafe_allow_html=True
    )

st.divider()


# =====================================================
# CARTA MENGIKUT BAHAGIAN
# Ikon kecil di sebelah tajuk membuka Paparan Carta Besar.
# Dalam mod tersebut, sidebar kekal kelihatan dan filter masih aktif.
# =====================================================
chart_title_col, chart_expand_col = st.columns([18, 1])

with chart_title_col:
    st.markdown("### 📊 PENCAPAIAN MENGIKUT BAHAGIAN")

with chart_expand_col:
    if st.button(
        "⛶",
        key="btn_buka_paparan_carta_besar",
        help="Buka Paparan Carta Besar dengan sidebar",
        use_container_width=True
    ):
        st.session_state.focus_page = "chart"
        st.rerun()

if chart_filtered_df.empty:
    st.warning("Tiada data carta untuk status yang dipilih.")
else:
    fig_bahagian_stack = build_bahagian_chart(
        chart_filtered_df,
        bahagian_col
    )

    st.plotly_chart(
        fig_bahagian_stack,
        use_container_width=True,
        key="bahagian_chart_main",
        config={
            "displaylogo": False,
            "responsive": True,
            "displayModeBar": True
        }
    )

st.divider()

# =====================================================
# LAPORAN STATUS PRESTASI - PAPARAN TERUS SEPERTI E-FILING
# Dataframe masih menggunakan ikon fullscreen asal Streamlit.
# =====================================================
report_title_col, report_expand_col = st.columns([18, 1])
with report_title_col:
    st.markdown("### 📋 LAPORAN STATUS PRESTASI")
with report_expand_col:
    if st.button("⛶", key="btn_buka_report_focus", help="Buka paparan besar Laporan", use_container_width=True):
        st.session_state.focus_page = "report"
        st.rerun()

tab_senarai, tab_summary = st.tabs([
    "📋 SENARAI STATUS PRESTASI",
    "📑 SUMMARY PROGRAM MENGIKUT BAHAGIAN"
])

with tab_senarai:
    selected = st.session_state.selected_traffic

    if selected == "Semua":
        render_selected_list(
            filtered_df,
            sektor_col,
            bahagian_col,
            program_col,
            pencapaian_col,
            "Semua"
        )

    elif selected == "Hijau":
        render_selected_list(
            df_hijau,
            sektor_col,
            bahagian_col,
            program_col,
            pencapaian_col,
            "Hijau"
        )

    elif selected == "Kuning":
        render_selected_list(
            df_kuning,
            sektor_col,
            bahagian_col,
            program_col,
            pencapaian_col,
            "Kuning"
        )

    elif selected == "Merah":
        render_selected_list(
            df_merah,
            sektor_col,
            bahagian_col,
            program_col,
            pencapaian_col,
            "Merah"
        )

    elif selected in ["Bermula Q2", "Q2"]:
        render_selected_list(
            df_bermula_q2_filtered,
            sektor_col,
            bahagian_col,
            program_col,
            pencapaian_col,
            "Q2"
        )

    elif selected in ["Bermula Q3", "Q3"]:
        render_selected_list(
            df_bermula_q3_filtered,
            sektor_col,
            bahagian_col,
            program_col,
            pencapaian_col,
            "Q3"
        )

    elif selected in ["Bermula Q4", "Q4"]:
        render_selected_list(
            df_bermula_q4_filtered,
            sektor_col,
            bahagian_col,
            program_col,
            pencapaian_col,
            "Q4"
        )

    elif selected == "Gugur":
        render_selected_list(
            df_tidak_filtered,
            sektor_col,
            bahagian_col,
            program_col,
            pencapaian_col,
            "Gugur"
        )

    else:
        st.info(
            "Klik Jumlah Program Aktif, bulatan Traffic Light, atau nilai "
            "Q2/Q3/Q4/Gugur untuk lihat senarai."
        )

with tab_summary:
    st.markdown("### SUMMARY PROGRAM MENGIKUT BAHAGIAN")
    st.caption(
        "Ringkasan ini berubah mengikut filter Sektor, Bahagian dan Kod Program."
    )

    summary_df = build_summary_program_bahagian(
        filtered_df,
        bahagian_col
    )

    if summary_df.empty:
        st.warning("Tiada data untuk dipaparkan dalam summary.")
    else:
        summary_display = summary_df.rename(columns={
            "Hijau": "🟢 HIJAU",
            "Kuning": "🟡 KUNING",
            "Merah": "🔴 MERAH",
            "Q3": "Q3",
            "Q4": "Q4",
            "Gugur": "GUGUR",
            "% PENCAPAIAN": "% PENCAPAIAN",
            "JUMLAH": "JUMLAH PROGRAM"
        })

        integer_cols = [
            "🟢 HIJAU",
            "🟡 KUNING",
            "🔴 MERAH",
            "Q3",
            "Q4",
            "GUGUR",
            "JUMLAH PROGRAM"
        ]

        for col in integer_cols:
            if col in summary_display.columns:
                summary_display[col] = (
                    pd.to_numeric(summary_display[col], errors="coerce")
                    .fillna(0)
                    .astype(int)
                )

        if "% PENCAPAIAN" in summary_display.columns:
            summary_display["% PENCAPAIAN"] = pd.to_numeric(
                summary_display["% PENCAPAIAN"], errors="coerce"
            ).fillna(0.0)

        summary_format = {
            col: "{:,.0f}"
            for col in integer_cols
            if col in summary_display.columns
        }
        if "% PENCAPAIAN" in summary_display.columns:
            summary_format["% PENCAPAIAN"] = "{:.2f}%"

        st.dataframe(
            summary_display.style.apply(
                highlight_summary_table,
                axis=1
            ).format(summary_format),
            use_container_width=True,
            hide_index=True,
            height=min(720, 80 + (len(summary_display) * 35))
        )

        summary_csv = summary_display.to_csv(
            index=False
        ).encode("utf-8-sig")

        st.download_button(
            label="⬇️ Download Summary Program Mengikut Bahagian",
            data=summary_csv,
            file_name="summary_program_mengikut_bahagian.csv",
            mime="text/csv",
            use_container_width=True,
            key="download_summary_bahagian_btn"
        )
