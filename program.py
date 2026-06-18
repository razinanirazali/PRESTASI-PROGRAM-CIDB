
import streamlit as st
import pandas as pd
import plotly.express as px
from pathlib import Path
import re

# =====================================================
# PAGE SETUP
# =====================================================
st.set_page_config(
    page_title="Dashboard Status Prestasi Q1 2026",
    layout="wide",
    initial_sidebar_state="expanded"
)

SHEET_NAME = " DATA DASHBOARD"

SHEET_NAME_OPTIONS = [
    " DATA DASHBOARD",
    "DATA DASHBOARD",
    " DATA DASHBOARD Q1",
    "DATA DASHBOARD Q1"
]

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
        width: 138px !important;
        height: 138px !important;
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
        font-size: 16px !important;
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
        width: 138px !important;
        height: 138px !important;
        min-height: 138px !important;
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
        font-size: 44px !important;
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


def resolve_sheet_name(xls):
    """Cari nama sheet sebenar secara fleksibel."""
    available = list(xls.sheet_names)

    for wanted in SHEET_NAME_OPTIONS:
        for sheet in available:
            if sheet == wanted:
                return sheet

    for wanted in SHEET_NAME_OPTIONS:
        wanted_clean = wanted.strip().upper()
        for sheet in available:
            if sheet.strip().upper() == wanted_clean:
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
    """
    Kesan status khas daripada Column K dan fallback untuk rekod yang tiada weightage.

    Logik Q1 2026:
    - Column K digunakan untuk status khas seperti Q2/Q3/Q4/Tidak dilaksanakan.
    - Variasi TIDAK DILAKSANAKAN / TIDAK AKAN DILAKSANAKAN / TIDAK DILAKSANA
      semuanya dikira sebagai TIDAK DILAKSANAKAN.
    - Jika rekod tiada status, tetapi Weightage L dan % Pencapaian M kosong,
      rekod tersebut dimasukkan ke Q2 supaya jumlah keseluruhan tetap sama
      dengan jumlah program sebenar.
    """
    try:
        status_text = row.iloc[STATUS_TEXT_COL_INDEX]
    except Exception:
        status_text = ""

    if pd.isna(status_text):
        status_text = ""

    status_text = clean_upper(status_text)
    status_text = " ".join(status_text.split())

    tidak_patterns = [
        "TIDAK DILAKSANAKAN",
        "TIDAK AKAN DILAKSANAKAN",
        "TIDAK DILAKSANA",
        "TIDAK AKAN DILAKSANA",
        "TAK DILAKSANAKAN",
        "TAK AKAN DILAKSANAKAN",
        "TAK DILAKSANA",
        "TAK AKAN DILAKSANA",
    ]

    if any(pattern in status_text for pattern in tidak_patterns):
        return "TIDAK DILAKSANAKAN"

    if "BERMULA Q2" in status_text or "MULA Q2" in status_text or status_text.strip() == "Q2":
        return "BERMULA Q2"

    if "BERMULA Q3" in status_text or "MULA Q3" in status_text or status_text.strip() == "Q3":
        return "BERMULA Q3"

    if "BERMULA Q4" in status_text or "MULA Q4" in status_text or status_text.strip() == "Q4":
        return "BERMULA Q4"

    # Fallback untuk 4 rekod yang tiada Weightage L dan tiada % Pencapaian M.
    # Tanpa fallback ini, jumlah detect menjadi 254 walaupun jumlah program 258.
    try:
        weightage_value = pd.to_numeric(
            str(row.iloc[WEIGHTAGE_COL_INDEX]).replace(",", "").replace("%", "").strip(),
            errors="coerce"
        )
    except Exception:
        weightage_value = pd.NA

    try:
        pencapaian_value = pd.to_numeric(
            str(row.iloc[PENCAPAIAN_COL_INDEX]).replace(",", "").replace("%", "").strip(),
            errors="coerce"
        )
    except Exception:
        pencapaian_value = pd.NA

    if pd.isna(weightage_value) and pd.isna(pencapaian_value):
        return "BERMULA Q2"

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

    if "Tidak Dilaksanakan" in traffic:
        return ["background-color: #e7e6e6"] * len(row)

    return [""] * len(row)


@st.cache_data
def load_data(uploaded_file):
    xls = pd.ExcelFile(uploaded_file)
    actual_sheet = resolve_sheet_name(xls)

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




def build_bahagian_chart(filtered_df, bahagian_col, chart_height=None):
    """
    Bina carta stacked bar mengikut bahagian.
    Fungsi ini digunakan untuk paparan normal dan paparan fullscreen custom.
    """
    bahagian_status = (
        filtered_df
        .groupby([bahagian_col, "KATEGORI_TRAFFIC"], as_index=False)
        .size()
        .rename(columns={"size": "JUMLAH"})
    )

    status_order = ["Hijau", "Kuning", "Merah", "Q2", "Q3", "Q4", "Tidak Dilaksanakan"]

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
            "Tidak Dilaksanakan": "#7a7788"
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


# =====================================================
# SIDEBAR FILTER ONLY
# =====================================================


# =====================================================
# MAIN TITLE
# =====================================================
st.title("PENCAPAIAN PRESTASI FIZIKAL PROGRAM CIDB Q1 2026")

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
df, sheet_list, header_row = load_data(uploaded_file)

if df is None:
    st.error("Worksheet **DATA DASHBOARD / DATA DASHBOARD Q1** tidak dijumpai dalam fail Excel.")
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

if len(df.columns) <= PENCAPAIAN_COL_INDEX:
    st.error("Column L atau Column M tidak wujud dalam sheet DATA DASHBOARD selepas header dibaca.")
    st.write("Kolum yang dibaca:")
    st.write(list(df.columns))
    st.stop()

weightage_col = df.columns[WEIGHTAGE_COL_INDEX]
pencapaian_col = df.columns[PENCAPAIAN_COL_INDEX]
pencapaian_fizikal_col = df.columns[PENCAPAIAN_FIZIKAL_COL_INDEX] if len(df.columns) > PENCAPAIAN_FIZIKAL_COL_INDEX else None


# =====================================================
# CLEAN DATA
# =====================================================
df[sektor_col] = df[sektor_col].astype("string").str.strip()
df[bahagian_col] = df[bahagian_col].astype("string").str.strip()
df[program_col] = df[program_col].astype("string").str.strip()

df["STATUS_KHAS"] = df.apply(detect_status_khas, axis=1)

df["WEIGHTAGE_L_NUM"] = to_number(df[weightage_col])
df["PENCAPAIAN_M_NUM"] = to_number(df[pencapaian_col])

# KPI PENCAPAIAN:
# Column L = SASARAN
# Column M = PRESTASI
# Traffic light dikira berdasarkan PRESTASI / SASARAN x 100.
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

df_dinilai = df[
    (df["STATUS_KHAS"] == "")
    & df["WEIGHTAGE_L_NUM"].notna()
    & (df["WEIGHTAGE_L_NUM"] > 0)
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
df_tidak_dilaksanakan["TRAFFIC_LIGHT"] = "⚫ Tidak Dilaksanakan"

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
    "filter_sektor",
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
    "filter_bahagian",
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
        "filter_kod_program",
        n_cols=3,
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
    st.session_state["filter_sektor_selected"] = []
    st.session_state["filter_bahagian_selected"] = []
    st.session_state["filter_kod_program_selected"] = []

    if "selected_traffic" in st.session_state:
        st.session_state.selected_traffic = None

    if "fullscreen_chart" in st.session_state:
        st.session_state.fullscreen_chart = False

    if "fullscreen_list" in st.session_state:
        st.session_state.fullscreen_list = False

    st.rerun()

if st.sidebar.button(
    "🔃 Refresh Data Excel",
    use_container_width=True,
    key="refresh_excel_btn"
):
    st.cache_data.clear()
    st.session_state["filter_sektor_selected"] = []
    st.session_state["filter_bahagian_selected"] = []
    st.session_state["filter_kod_program_selected"] = []

    if "selected_traffic" in st.session_state:
        st.session_state.selected_traffic = None

    if "fullscreen_chart" in st.session_state:
        st.session_state.fullscreen_chart = False

    if "fullscreen_list" in st.session_state:
        st.session_state.fullscreen_list = False

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
df_hijau = filtered_dinilai[filtered_dinilai["KPI_PENCAPAIAN_NUM"] >= 85].copy()

df_kuning = filtered_dinilai[
    (filtered_dinilai["KPI_PENCAPAIAN_NUM"] >= 60)
    & (filtered_dinilai["KPI_PENCAPAIAN_NUM"] < 85)
].copy()

df_merah = filtered_dinilai[filtered_dinilai["KPI_PENCAPAIAN_NUM"] < 60].copy()

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

jumlah_program_panel = len(panel_df)

panel_df["SASARAN_PANEL_NUM"] = (
    panel_df["WEIGHTAGE_L_NUM"]
    .clip(lower=0, upper=100)
)

panel_df["PRESTASI_PANEL_NUM"] = (
    panel_df["PENCAPAIAN_M_NUM"]
    .clip(lower=0, upper=100)
)

# Kiraan panel Q1 (formula betul):
# - Column L = nilai sasaran sebenar setiap program.
# - Column M = nilai pencapaian/prestasi sebenar setiap program.
# - Sasaran panel tetap 25%.
# - Prestasi panel = jumlah Column M / jumlah Column L x 25.
# - Pencapaian panel = Prestasi / Sasaran x 100.
# - Hanya rekod yang ada nilai Column L dan Column M dikira.
panel_valid = panel_df[
    panel_df["WEIGHTAGE_L_NUM"].notna()
    & (panel_df["WEIGHTAGE_L_NUM"] > 0)
    & panel_df["PENCAPAIAN_M_NUM"].notna()
].copy()

if not panel_valid.empty:
    sasaran_panel = 25.00

    jumlah_l = panel_valid["WEIGHTAGE_L_NUM"].sum()
    jumlah_m = panel_valid["PENCAPAIAN_M_NUM"].sum()

    if jumlah_l > 0:
        prestasi_panel = (jumlah_m / jumlah_l) * sasaran_panel
        pencapaian_panel = (prestasi_panel / sasaran_panel) * 100
    else:
        prestasi_panel = 0
        pencapaian_panel = 0
else:
    sasaran_panel = 0
    prestasi_panel = 0
    pencapaian_panel = 0

pencapaian_panel = min(pencapaian_panel, 100)

# Kategori ringkas untuk stacked bar chart
filtered_df["KATEGORI_TRAFFIC"] = filtered_df["TRAFFIC_LIGHT"].replace({
    "🟢 Hijau": "Hijau",
    "🟡 Kuning": "Kuning",
    "🔴 Merah": "Merah",
    "🟣 Q2": "Q2",
    "🟣 Q3": "Q3",
    "🟣 Q4": "Q4",
    "⚫ Tidak Dilaksanakan": "Tidak Dilaksanakan"
})

if "selected_traffic" not in st.session_state:
    st.session_state.selected_traffic = None

if "fullscreen_chart" not in st.session_state:
    st.session_state.fullscreen_chart = False

if "fullscreen_list" not in st.session_state:
    st.session_state.fullscreen_list = False

# =====================================================
# CUSTOM FULLSCREEN PAGE - CARTA SAHAJA
# =====================================================
if st.session_state.fullscreen_chart:
    st.markdown(
        """
        <style>
        section[data-testid="stSidebar"] {
            display: none !important;
        }

        .main .block-container {
            max-width: 100vw !important;
            width: 100vw !important;
            min-height: 100vh !important;
            padding: 1rem 1.5rem 1.5rem 1.5rem !important;
            margin: 0 !important;
            border-radius: 0 !important;
            box-shadow: none !important;
            background: #ffffff !important;
            backdrop-filter: none !important;
            -webkit-backdrop-filter: none !important;
        }

        .stApp {
            background: #ffffff !important;
        }

        header[data-testid="stHeader"] {
            background: #ffffff !important;
        }
        </style>
        """,
        unsafe_allow_html=True
    )

    top_left, top_right = st.columns([0.78, 0.22])

    with top_left:
        st.markdown("## 📊 PENCAPAIAN MENGIKUT BAHAGIAN")

    with top_right:
        if st.button("⬅️ Kembali ke Dashboard", use_container_width=True):
            st.session_state.fullscreen_chart = False
            st.rerun()

    fig_fullscreen = build_bahagian_chart(
        filtered_df,
        bahagian_col,
        chart_height=max(720, filtered_df[bahagian_col].nunique() * 60)
    )

    st.plotly_chart(
        fig_fullscreen,
        use_container_width=True,
        config={"displaylogo": False}
    )

    st.stop()

# =====================================================
# CUSTOM FULLSCREEN PAGE - SENARAI STATUS PRESTASI SAHAJA
# =====================================================
if st.session_state.fullscreen_list:
    st.markdown(
        """
        <style>
        section[data-testid="stSidebar"] {
            display: none !important;
        }

        .main .block-container {
            max-width: 100vw !important;
            width: 100vw !important;
            min-height: 100vh !important;
            padding: 1rem 1.5rem 1.5rem 1.5rem !important;
            margin: 0 !important;
            border-radius: 0 !important;
            box-shadow: none !important;
            background: #ffffff !important;
            backdrop-filter: none !important;
            -webkit-backdrop-filter: none !important;
        }

        .stApp {
            background: #ffffff !important;
        }

        header[data-testid="stHeader"] {
            background: #ffffff !important;
        }
        </style>
        """,
        unsafe_allow_html=True
    )

    top_left, top_right = st.columns([0.78, 0.22])

    with top_left:
        st.markdown("## 📋 SENARAI STATUS PRESTASI")

    with top_right:
        if st.button("⬅️ Kembali ke Dashboard", use_container_width=True):
            st.session_state.fullscreen_list = False
            st.rerun()

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

    elif selected == "Tidak Dilaksanakan":
        render_selected_list(
            df_tidak_filtered,
            sektor_col,
            bahagian_col,
            program_col,
            pencapaian_col,
            "Tidak Dilaksanakan"
        )

    else:
        st.info("Klik Jumlah Program, bulatan Traffic Light, atau nilai Q2/Q3/Q4/Tidak Dilaksanakan untuk lihat senarai.")

    st.stop()

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

    st.markdown("<br>", unsafe_allow_html=True)

    # Baris khas: Q2 / Q3 / Q4 / TIDAK DILAKSANAKAN
    st.markdown("<div style='margin-top:22px; margin-bottom:10px;'>", unsafe_allow_html=True)
    q_col1, q_sep1, q_col2, q_sep2, q_col3, q_sep3, q_col4 = st.columns([1.45, 0.12, 1.45, 0.12, 1.45, 0.12, 3.25])

    with q_col1:
        q2_label, q2_value = st.columns([2.5, 1])
        with q2_label:
            st.markdown('<div class="status-label-text">Q2 -</div>', unsafe_allow_html=True)
        with q2_value:
            if st.button(str(bermula_q2), key="btn_q2_status"):
                st.session_state.selected_traffic = "Q2"

    with q_sep1:
        st.markdown('<div class="status-separator-text">|</div>', unsafe_allow_html=True)

    with q_col2:
        q3_label, q3_value = st.columns([2.5, 1])
        with q3_label:
            st.markdown('<div class="status-label-text">Q3 -</div>', unsafe_allow_html=True)
        with q3_value:
            if st.button(str(bermula_q3), key="btn_q3_status"):
                st.session_state.selected_traffic = "Q3"

    with q_sep2:
        st.markdown('<div class="status-separator-text">|</div>', unsafe_allow_html=True)

    with q_col3:
        q4_label, q4_value = st.columns([2.5, 1])
        with q4_label:
            st.markdown('<div class="status-label-text">Q4 -</div>', unsafe_allow_html=True)
        with q4_value:
            if st.button(str(bermula_q4), key="btn_q4_status"):
                st.session_state.selected_traffic = "Q4"

    with q_sep3:
        st.markdown('<div class="status-separator-text">|</div>', unsafe_allow_html=True)

    with q_col4:
        tidak_label, tidak_value = st.columns([6.5, 1])
        with tidak_label:
            st.markdown('<div class="status-label-text">TIDAK DILAKSANAKAN -</div>', unsafe_allow_html=True)
        with tidak_value:
            if st.button(str(tidak_dilaksanakan), key="btn_tidak_status"):
                st.session_state.selected_traffic = "Tidak Dilaksanakan"

    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)

with main_right:
    st.markdown('<div class="summary-panel">', unsafe_allow_html=True)

    if st.button(f"{jumlah_program_panel:,.0f}", key="btn_jumlah_program"):
        st.session_state.selected_traffic = "Semua"

    st.markdown(
        f"""
            <div class="summary-label">Jumlah Program</div>
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
# ACCORDION 1 - CARTA MENGIKUT BAHAGIAN
# =====================================================
with st.expander("📊 PENCAPAIAN MENGIKUT BAHAGIAN", expanded=False):

    btn_col1, btn_col2 = st.columns([0.78, 0.22])

    with btn_col2:
        if st.button("🖥️ Full Screen Carta", use_container_width=True):
            st.session_state.fullscreen_chart = True
            st.rerun()

    fig_bahagian_stack = build_bahagian_chart(filtered_df, bahagian_col)

    st.plotly_chart(
        fig_bahagian_stack,
        use_container_width=True,
        config={"displaylogo": False}
    )


# =====================================================
# ACCORDION 2 - SENARAI APABILA KLIK TRAFFIC LIGHT
# =====================================================
with st.expander("📋 SENARAI STATUS PRESTASI", expanded=False):

    list_btn_col1, list_btn_col2 = st.columns([0.78, 0.22])

    with list_btn_col2:
        if st.button("🖥️ Full Screen Senarai", use_container_width=True):
            st.session_state.fullscreen_list = True
            st.rerun()

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

    elif selected == "Tidak Dilaksanakan":
        render_selected_list(
            df_tidak_filtered,
            sektor_col,
            bahagian_col,
            program_col,
            pencapaian_col,
            "Tidak Dilaksanakan"
        )

    else:
        st.info("Klik Jumlah Program, bulatan Traffic Light, atau nilai Q2/Q3/Q4/Tidak Dilaksanakan untuk lihat senarai.")


