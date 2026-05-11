# ============================================================
# SISTEM KLASIFIKASI PENERIMA BSM - NAIVE BAYES
# Dashboard Machine Learning Modern | Streamlit App
# Semua fitur dalam 1 file: app.py
# Jalankan: streamlit run app.py
# ============================================================

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from sklearn.preprocessing import StandardScaler, LabelEncoder
from sklearn.model_selection import train_test_split
from sklearn.naive_bayes import GaussianNB
from sklearn.metrics import accuracy_score, confusion_matrix
import io
import time
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────────────────────────
# KONFIGURASI HALAMAN
# ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Klasifikasi BSM | Naive Bayes",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ─────────────────────────────────────────────────────────────
# CUSTOM CSS MODERN (Emerald, Putih, Abu Soft, Dark Mode Ready)
# ─────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;600&display=swap');

/* ===== VARIABEL ===== */
:root {
    --em-50:#ecfdf5;--em-100:#d1fae5;--em-200:#a7f3d0;--em-300:#6ee7b7;
    --em-400:#34d399;--em-500:#10b981;--em-600:#059669;--em-700:#047857;
    --em-800:#065f46;--em-900:#064e3b;
    --sl-50:#f8fafc;--sl-100:#f1f5f9;--sl-200:#e2e8f0;--sl-400:#94a3b8;
    --sl-500:#64748b;--sl-600:#475569;--sl-700:#334155;--sl-800:#1e293b;
    --white:#ffffff;
    --r:14px;--r-sm:8px;
    --sh:0 2px 12px rgba(0,0,0,0.07);--sh-md:0 6px 24px rgba(0,0,0,0.12);
}

/* ===== FONT GLOBAL ===== */
html,body{font-family:'Plus Jakarta Sans',sans-serif;}
.main,.main *{font-family:'Plus Jakarta Sans',sans-serif;}
.main .block-container{padding:1.5rem 2rem 3rem;max-width:1400px;}

/* ===== SIDEBAR — hanya background, TIDAK sentuh overflow/scroll ===== */
[data-testid="stSidebar"]>div:first-child{
    background:linear-gradient(175deg,#064e3b 0%,#065f46 55%,#1e293b 100%);
}
[data-testid="stSidebar"] p,
[data-testid="stSidebar"] span {
    color:#ffffff;
}
[data-testid="stSidebar"] label{
    color:#6ee7b7!important;font-size:0.78rem;font-weight:600;
    letter-spacing:0.05em;text-transform:uppercase;
}
[data-testid="stSidebar"] .stRadio>div{gap:4px;flex-direction:column;}
[data-testid="stSidebar"] .stRadio label{
    background:rgba(255,255,255,0.06)!important;
    border-radius:8px!important;padding:9px 12px!important;
    font-size:0.88rem!important;font-weight:500!important;
    color:#ffffff!important;text-transform:none!important;
    letter-spacing:0!important;border:1px solid transparent!important;
    cursor:pointer;transition:background 0.15s,border-color 0.15s;
}
[data-testid="stSidebar"] .stRadio label:hover{
    background:rgba(52,211,153,0.14)!important;
    border-color:rgba(52,211,153,0.3)!important;
}
[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p{
    color:#a7f3d0!important;font-size:0.82rem;
}

/* ===== TOMBOL COLLAPSE — jangan disembunyikan ===== */
[data-testid="collapsedControl"]{
    visibility:visible!important;
    pointer-events:auto!important;
    opacity:1!important;
}

/* ===== HERO ===== */
.hero-header{
    background:linear-gradient(135deg,#065f46 0%,#059669 55%,#34d399 100%);
    border-radius:var(--r);padding:2rem 2.4rem;margin-bottom:1.6rem;
    position:relative;overflow:hidden;
    box-shadow:0 4px 20px rgba(16,185,129,0.2);
}
.hero-header::before{
    content:'';position:absolute;top:-50px;right:-50px;
    width:200px;height:200px;background:rgba(255,255,255,0.05);
    border-radius:50%;pointer-events:none;
}
.hero-badge{
    display:inline-block;background:rgba(255,255,255,0.16);color:#fff;
    border:1px solid rgba(255,255,255,0.25);border-radius:999px;
    padding:3px 12px;font-size:0.72rem;font-weight:600;
    letter-spacing:0.07em;text-transform:uppercase;margin-bottom:10px;
}
.hero-title{font-size:1.85rem;font-weight:800;color:#fff!important;margin:0 0 4px;line-height:1.2;}
.hero-sub{font-size:0.9rem;color:#d1fae5!important;margin:0;}

/* ===== METRIC CARDS ===== */
.metric-card{
    background:var(--white);border-radius:var(--r);padding:1.3rem 1.4rem;
    box-shadow:var(--sh);border:1px solid var(--sl-200);
    position:relative;overflow:hidden;transition:transform 0.18s,box-shadow 0.18s;
}
.metric-card::before{
    content:'';position:absolute;top:0;left:0;right:0;height:3px;
    background:linear-gradient(90deg,var(--em-600),var(--em-400));
}
.metric-card:hover{transform:translateY(-2px);box-shadow:var(--sh-md);}
.metric-icon{font-size:1.6rem;margin-bottom:6px;}
.metric-label{font-size:0.7rem;font-weight:700;color:var(--sl-400);text-transform:uppercase;letter-spacing:0.08em;margin-bottom:4px;}
.metric-value{font-size:1.9rem;font-weight:800;color:var(--em-700);line-height:1;font-family:'JetBrains Mono',monospace;}
.metric-desc{font-size:0.74rem;color:var(--sl-400);margin-top:4px;}

/* ===== SECTION HEADER ===== */
.section-header{display:flex;align-items:center;gap:10px;margin:1.6rem 0 1rem;}
.section-icon{width:36px;height:36px;background:var(--em-100);border-radius:var(--r-sm);display:flex;align-items:center;justify-content:center;font-size:1.05rem;flex-shrink:0;}
.section-title{font-size:1.1rem;font-weight:700;color:var(--sl-800);margin:0;}
.section-divider{height:2px;border-radius:2px;margin-bottom:1.2rem;background:linear-gradient(90deg,var(--em-500) 0%,transparent 100%);opacity:0.35;}

/* ===== CARDS ===== */
.card{background:var(--white);border-radius:var(--r);padding:1.4rem;box-shadow:var(--sh);border:1px solid var(--sl-200);margin-bottom:1rem;}
.card-title{font-size:0.78rem;font-weight:700;color:var(--sl-600);text-transform:uppercase;letter-spacing:0.07em;margin-bottom:0.9rem;padding-bottom:0.6rem;border-bottom:1px solid var(--sl-100);}

/* ===== STEP CARDS ===== */
.step-card{background:var(--white);border-radius:var(--r);padding:1.1rem 1.4rem;border-left:4px solid var(--em-500);box-shadow:var(--sh);margin-bottom:0.7rem;}
.step-title{font-weight:700;color:var(--sl-700);font-size:0.88rem;margin-bottom:2px;}
.step-desc{font-size:0.78rem;color:var(--sl-400);}

/* ===== ALERTS ===== */
.alert-success{background:#f0fdf4;border:1px solid #bbf7d0;border-radius:var(--r-sm);padding:10px 14px;color:#166534;font-size:0.84rem;font-weight:500;margin:6px 0;}
.alert-info{background:#eff6ff;border:1px solid #bfdbfe;border-radius:var(--r-sm);padding:10px 14px;color:#1e40af;font-size:0.84rem;font-weight:500;margin:6px 0;}
.alert-warning{background:#fffbeb;border:1px solid #fde68a;border-radius:var(--r-sm);padding:10px 14px;color:#92400e;font-size:0.84rem;font-weight:500;margin:6px 0;}

/* ===== BADGES ===== */
.badge{display:inline-block;border-radius:999px;padding:2px 10px;font-size:0.7rem;font-weight:700;}
.badge-green{background:var(--em-100);color:var(--em-800);}
.badge-red{background:#fee2e2;color:#991b1b;}
.badge-blue{background:#dbeafe;color:#1e40af;}
.badge-slate{background:var(--sl-100);color:var(--sl-600);}

/* ===== ABOUT CARDS ===== */
.about-card{background:var(--white);border-radius:var(--r);padding:1.3rem 1.5rem;border:1px solid var(--sl-200);box-shadow:var(--sh);height:100%;}
.about-card-icon{font-size:1.8rem;margin-bottom:8px;}
.about-card-title{font-size:0.92rem;font-weight:700;color:var(--sl-800);margin-bottom:6px;}
.about-card-desc{font-size:0.8rem;color:var(--sl-500);line-height:1.6;}

/* ===== SIDEBAR HTML HELPERS ===== */
.sidebar-logo{text-align:center;padding:1.2rem 1rem 0.8rem;}
.sidebar-logo-icon{font-size:2.5rem;margin-bottom:6px;}
.sidebar-logo-title{font-size:0.95rem;font-weight:800;color:#fff;line-height:1.2;}
.sidebar-logo-sub{font-size:0.7rem;color:#6ee7b7;margin-top:2px;}
.sidebar-divider{height:1px;background:rgba(255,255,255,0.1);margin:10px 0;}
.sidebar-section-label{font-size:0.62rem;color:#34d399;text-transform:uppercase;letter-spacing:0.14em;font-weight:700;padding:0 4px;margin-bottom:4px;}

/* ===== FOOTER ===== */
.footer{background:var(--sl-800);color:#cbd5e1;border-radius:var(--r);padding:1.4rem 2rem;text-align:center;margin-top:2.5rem;font-size:0.8rem;line-height:1.8;}
.footer-brand{color:var(--em-400);font-weight:700;font-size:0.95rem;}
.footer-divider{width:36px;height:2px;background:var(--em-500);margin:8px auto;border-radius:2px;}

/* ===== STREAMLIT KOMPONEN ===== */
.stTabs [data-baseweb="tab-list"]{gap:4px;background:var(--sl-100);padding:4px;border-radius:10px;}
.stTabs [data-baseweb="tab"]{border-radius:8px!important;padding:6px 14px!important;font-weight:600!important;font-size:0.82rem!important;color:var(--sl-500)!important;}
.stTabs [aria-selected="true"]{background:var(--white)!important;color:var(--em-700)!important;box-shadow:var(--sh)!important;}

.stButton>button{
    background:linear-gradient(135deg,var(--em-700),var(--em-500))!important;
    color:#fff!important;border:none!important;border-radius:var(--r-sm)!important;
    font-weight:600!important;transition:all 0.18s!important;
    box-shadow:0 2px 8px rgba(16,185,129,0.25)!important;
}
.stButton>button:hover{transform:translateY(-1px)!important;box-shadow:0 5px 18px rgba(16,185,129,0.35)!important;}

.stDownloadButton>button{
    background:var(--white)!important;color:var(--em-700)!important;
    border:2px solid var(--em-500)!important;border-radius:var(--r-sm)!important;
    font-weight:600!important;transition:all 0.18s!important;
}
.stDownloadButton>button:hover{background:var(--em-50)!important;transform:translateY(-1px)!important;}

.stProgress>div>div>div{background:linear-gradient(90deg,var(--em-600),var(--em-400))!important;border-radius:999px!important;}

code,pre{font-family:'JetBrains Mono',monospace!important;}

::-webkit-scrollbar{width:5px;height:5px;}
::-webkit-scrollbar-track{background:var(--sl-100);}
::-webkit-scrollbar-thumb{background:var(--em-300);border-radius:3px;}
::-webkit-scrollbar-thumb:hover{background:var(--em-500);}

/* Sembunyikan branding Streamlit saja */
#MainMenu{visibility:hidden;}
footer{visibility:hidden;}
.stDeployButton{display:none;}
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────
# HELPER FUNCTIONS
# ─────────────────────────────────────────────────────────────

def render_hero(title: str, subtitle: str, badge: str = "Machine Learning"):
    """Render hero header di atas setiap halaman."""
    st.markdown(f"""
    <div class="hero-header">
        <div class="hero-badge">🎓 {badge}</div>
        <div class="hero-title">{title}</div>
        <div class="hero-sub">{subtitle}</div>
    </div>
    """, unsafe_allow_html=True)


def render_metric_cards(metrics: list):
    """
    metrics: list of dict {icon, label, value, desc}
    """
    cols = st.columns(len(metrics))
    for col, m in zip(cols, metrics):
        with col:
            st.markdown(f"""
            <div class="metric-card">
                <div class="metric-icon">{m['icon']}</div>
                <div class="metric-label">{m['label']}</div>
                <div class="metric-value">{m['value']}</div>
                <div class="metric-desc">{m['desc']}</div>
            </div>
            """, unsafe_allow_html=True)


def render_section(icon: str, title: str):
    """Render section header dengan ikon."""
    st.markdown(f"""
    <div class="section-header">
        <div class="section-icon">{icon}</div>
        <div class="section-title">{title}</div>
    </div>
    <div class="section-divider"></div>
    """, unsafe_allow_html=True)


def alert(msg: str, kind: str = "success"):
    icons = {"success": "✅", "info": "ℹ️", "warning": "⚠️"}
    st.markdown(f'<div class="alert-{kind}">{icons.get(kind,"•")} {msg}</div>',
                unsafe_allow_html=True)


def make_plotly_theme():
    """Layout dasar Plotly agar konsisten dengan tema hijau-putih."""
    return dict(
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Plus Jakarta Sans, sans-serif", color="#334155"),
        margin=dict(l=20, r=20, t=40, b=20),
    )


# ─────────────────────────────────────────────────────────────
# PREPROCESSING ENGINE
# ─────────────────────────────────────────────────────────────

def run_preprocessing(df_raw: pd.DataFrame):
    """
    Jalankan seluruh pipeline preprocessing:
    missing value → duplikasi → tipe data → encoding → standardisasi
    Kembalikan (df_processed, info_dict, le_pekerjaan, scaler)
    """
    df = df_raw.copy()

    # --- Rename kolom ---
    expected_cols = [
        'NO', 'NAMA PESERTA DIDIK', 'SEKOLAH', 'KELAS',
        'NIK', 'NISN', 'NAMA AYAH/IBU', 'KECAMATAN',
        'BESARAN BIAYA', 'PENDAPATAN ORANG TUA',
        'PEKERJAAN ORANG TUA', 'JUMLAH TANGGUNGAN',
        'STATUS RUMAH', 'LABEL'
    ]
    if len(df.columns) == len(expected_cols):
        df.columns = expected_cols

    info = {}

    # 1. Missing value (sebelum konversi)
    info['missing_before'] = df.isnull().sum().to_dict()

    # 2. Konversi tipe data numerik
    df['PENDAPATAN ORANG TUA'] = pd.to_numeric(df['PENDAPATAN ORANG TUA'], errors='coerce')
    df['JUMLAH TANGGUNGAN']    = pd.to_numeric(df['JUMLAH TANGGUNGAN'],    errors='coerce')

    # 3. Isi missing value
    df['PENDAPATAN ORANG TUA'] = df['PENDAPATAN ORANG TUA'].fillna(df['PENDAPATAN ORANG TUA'].mean())
    df['JUMLAH TANGGUNGAN']    = df['JUMLAH TANGGUNGAN'].fillna(df['JUMLAH TANGGUNGAN'].median())

    info['missing_after'] = df.isnull().sum().to_dict()

    # 4. Hapus duplikasi
    before_dup = len(df)
    df = df.drop_duplicates()
    after_dup  = len(df)
    info['dup_removed'] = before_dup - after_dup
    info['row_after_dup'] = after_dup

    # 5. KELAS → numerik
    df['KELAS_NUM'] = df['KELAS'].astype(str).str.extract(r'(\d+)')
    df['KELAS_NUM'] = pd.to_numeric(df['KELAS_NUM'], errors='coerce').fillna(0).astype(int)

    # 6. PEKERJAAN ORANG TUA → Label Encoding
    le_pekerjaan = LabelEncoder()
    df['PEKERJAAN ORANG TUA_ENC'] = le_pekerjaan.fit_transform(
        df['PEKERJAAN ORANG TUA'].astype(str)
    )

    # 7. STATUS RUMAH → biner
    status_rumah_mapping = {"Milik Sendiri": 1, "Kontrak/sewa": 0}
    df['STATUS RUMAH_ENC'] = df['STATUS RUMAH'].map(status_rumah_mapping)
    df['STATUS RUMAH_ENC'] = df['STATUS RUMAH_ENC'].fillna(0).astype(int)

    # 8. PENDAPATAN → Z-Score
    scaler = StandardScaler()
    df['PENDAPATAN_ZSCORE'] = scaler.fit_transform(df[['PENDAPATAN ORANG TUA']])

    # 9. LABEL → biner
    df['LABEL'] = df['LABEL'].astype(str).str.strip().str.capitalize()
    label_mapping = {"Ya": 1, "Tidak": 0}
    df['LABEL_ENC'] = df['LABEL'].map(label_mapping)

    info['le_pekerjaan'] = le_pekerjaan
    info['scaler']       = scaler
    info['total_rows']   = len(df)

    return df, info


# ─────────────────────────────────────────────────────────────
# MODEL TRAINING ENGINE
# ─────────────────────────────────────────────────────────────

def run_training(df: pd.DataFrame, test_size: float = 0.2):
    """
    Latih model Gaussian Naive Bayes dan kembalikan seluruh artefak.
    """
    feature_cols = [
        'KELAS_NUM',
        'PENDAPATAN_ZSCORE',
        'PEKERJAAN ORANG TUA_ENC',
        'JUMLAH TANGGUNGAN',
        'STATUS RUMAH_ENC'
    ]

    X = df[feature_cols].copy()
    y = df['LABEL_ENC'].copy()

    # Hapus baris NaN pada fitur/target
    valid_idx = X.notna().all(axis=1) & y.notna()
    X = X[valid_idx]
    y = y[valid_idx]

    X_train, X_test, y_train, y_test = train_test_split(
        X, y, test_size=test_size, random_state=42, stratify=y
    )

    model = GaussianNB()
    model.fit(X_train, y_train)

    y_train_pred  = model.predict(X_train)
    y_pred        = model.predict(X_test)
    y_pred_proba  = model.predict_proba(X_test)

    train_acc = accuracy_score(y_train, y_train_pred)
    test_acc  = accuracy_score(y_test, y_pred)
    cm        = confusion_matrix(y_test, y_pred)

    # Metrik manual per kelas
    TP = cm[1, 1]; TN = cm[0, 0]
    FP = cm[0, 1]; FN = cm[1, 0]

    prec0 = TN / (TN + FN) if (TN + FN) > 0 else 0
    rec0  = TN / (TN + FP) if (TN + FP) > 0 else 0
    f1_0  = 2 * prec0 * rec0 / (prec0 + rec0) if (prec0 + rec0) > 0 else 0

    prec1 = TP / (TP + FP) if (TP + FP) > 0 else 0
    rec1  = TP / (TP + FN) if (TP + FN) > 0 else 0
    f1_1  = 2 * prec1 * rec1 / (prec1 + rec1) if (prec1 + rec1) > 0 else 0

    sup0 = int((y_test == 0).sum())
    sup1 = int((y_test == 1).sum())
    tot  = sup0 + sup1

    macro_prec = (prec0 + prec1) / 2
    macro_rec  = (rec0  + rec1)  / 2
    macro_f1   = (f1_0  + f1_1)  / 2
    w_prec     = (prec0 * sup0 + prec1 * sup1) / tot
    w_rec      = (rec0  * sup0 + rec1  * sup1) / tot
    w_f1       = (f1_0  * sup0 + f1_1  * sup1) / tot

    results_df = pd.DataFrame({
        'No': range(1, len(y_test) + 1),
        'Actual':    ['Penerima' if v == 1 else 'Tidak Penerima' for v in y_test.values],
        'Predicted': ['Penerima' if v == 1 else 'Tidak Penerima' for v in y_pred],
        'Prob Tidak (%)': (y_pred_proba[:, 0] * 100).round(2),
        'Prob Ya (%)':    (y_pred_proba[:, 1] * 100).round(2),
        'Status': ['✓ Benar' if a == p else '✗ Salah'
                   for a, p in zip(y_test.values, y_pred)]
    })

    return {
        'model': model,
        'X': X, 'y': y,
        'X_train': X_train, 'X_test': X_test,
        'y_train': y_train, 'y_test': y_test,
        'y_pred': y_pred, 'y_pred_proba': y_pred_proba,
        'train_acc': train_acc, 'test_acc': test_acc,
        'cm': cm,
        'prec0': prec0, 'rec0': rec0,  'f1_0': f1_0,
        'prec1': prec1, 'rec1': rec1,  'f1_1': f1_1,
        'macro_prec': macro_prec, 'macro_rec': macro_rec, 'macro_f1': macro_f1,
        'w_prec': w_prec, 'w_rec': w_rec, 'w_f1': w_f1,
        'sup0': sup0, 'sup1': sup1,
        'feature_cols': feature_cols,
        'results_df': results_df,
    }


# ─────────────────────────────────────────────────────────────
# EXCEL EXPORT HELPERS
# ─────────────────────────────────────────────────────────────

def _style_header(cell, bg="1F4E79", fg="FFFFFF"):
    cell.font      = Font(name='Arial', bold=True, color=fg, size=10)
    cell.fill      = PatternFill("solid", start_color=bg)
    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    thin = Side(style='thin', color='BFBFBF')
    cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)


def _style_data(cell, bg="FFFFFF", bold=False, align='center'):
    cell.font      = Font(name='Arial', size=9, bold=bold)
    cell.fill      = PatternFill("solid", start_color=bg)
    cell.alignment = Alignment(horizontal=align, vertical='center')
    thin = Side(style='thin', color='D9D9D9')
    cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)


def export_preprocessing_excel(df: pd.DataFrame, le_pekerjaan: LabelEncoder) -> bytes:
    """Buat file Excel data preprocessing dengan styling rapi."""
    df_exp = df[[
        'NO', 'NAMA PESERTA DIDIK', 'SEKOLAH', 'KELAS', 'KELAS_NUM',
        'PENDAPATAN ORANG TUA', 'PENDAPATAN_ZSCORE',
        'PEKERJAAN ORANG TUA', 'PEKERJAAN ORANG TUA_ENC',
        'JUMLAH TANGGUNGAN', 'STATUS RUMAH', 'STATUS RUMAH_ENC',
        'LABEL', 'LABEL_ENC'
    ]].copy()

    for col in ['KELAS_NUM', 'PEKERJAAN ORANG TUA_ENC', 'STATUS RUMAH_ENC', 'LABEL_ENC']:
        df_exp[col] = pd.to_numeric(df_exp[col], errors='coerce').astype('Int64')

    df_exp.columns = [
        'No', 'Nama Peserta Didik', 'Sekolah', 'Kelas', 'Kelas (Angka)',
        'Pendapatan Orang Tua (Rp)', 'Pendapatan (Z-Score)',
        'Pekerjaan Orang Tua', 'Pekerjaan (Encoded)',
        'Jumlah Tanggungan', 'Status Rumah', 'Status Rumah (Encoded)',
        'Label', 'Label (Encoded)'
    ]

    wb = Workbook()
    ws = wb.active
    ws.title = "Data Preprocessing"

    # Judul
    ws.merge_cells('A1:N1')
    tc = ws['A1']
    tc.value = "DATA HASIL PREPROCESSING - BSM 2022-2024"
    tc.font = Font(name='Arial', bold=True, size=13, color='FFFFFF')
    tc.fill = PatternFill("solid", start_color="065f46")
    tc.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 28

    ws.merge_cells('A2:N2')
    sc = ws['A2']
    sc.value = f"Total Data: {len(df_exp)} baris | Fitur: 5 | Target: LABEL (Ya=1, Tidak=0)"
    sc.font  = Font(name='Arial', size=9, italic=True, color='595959')
    sc.fill  = PatternFill("solid", start_color="d1fae5")
    sc.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[2].height = 16

    headers = list(df_exp.columns)
    for ci, h in enumerate(headers, 1):
        _style_header(ws.cell(row=3, column=ci, value=h))
    ws.row_dimensions[3].height = 30

    encoded_int = {'Kelas (Angka)', 'Pekerjaan (Encoded)', 'Status Rumah (Encoded)', 'Label (Encoded)'}
    row_colors = ["FFFFFF", "ecfdf5"]
    for ri, row in enumerate(df_exp.itertuples(index=False), 4):
        bg = row_colors[(ri - 4) % 2]
        for ci, val in enumerate(row, 1):
            cn = headers[ci - 1]
            if cn in encoded_int:
                try: val = int(val)
                except: pass
            cell = ws.cell(row=ri, column=ci, value=val)
            if 'Encoded' in cn or 'Z-Score' in cn or 'Angka' in cn:
                abg = "FFF9C4" if bg == "FFFFFF" else "FFF3CD"
            else:
                abg = bg
            _style_data(cell, abg)
            if cn == 'Pendapatan Orang Tua (Rp)':
                cell.number_format = '#,##0'
                cell.alignment = Alignment(horizontal='right', vertical='center')
            elif cn == 'Pendapatan (Z-Score)':
                cell.number_format = '0.000000'
                cell.alignment = Alignment(horizontal='right', vertical='center')
            elif cn in ('No', 'Kelas (Angka)', 'Pekerjaan (Encoded)',
                        'Status Rumah (Encoded)', 'Label (Encoded)', 'Jumlah Tanggungan'):
                cell.alignment = Alignment(horizontal='center', vertical='center')
            else:
                cell.alignment = Alignment(horizontal='left', vertical='center')

    for i, w in enumerate([5, 28, 28, 10, 12, 22, 18, 22, 16, 16, 18, 18, 8, 14], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # Sheet Mapping
    wm = wb.create_sheet("Mapping Encoding")
    wm.merge_cells('A1:B1')
    wm['A1'].value = "Mapping Encoding"
    wm['A1'].font  = Font(name='Arial', bold=True, size=12, color='FFFFFF')
    wm['A1'].fill  = PatternFill("solid", start_color="065f46")
    wm['A1'].alignment = Alignment(horizontal='center', vertical='center')
    wm.row_dimensions[1].height = 24

    _style_header(wm.cell(row=3, column=1, value="Pekerjaan Orang Tua"), "059669")
    _style_header(wm.cell(row=3, column=2, value="Kode"), "059669")
    for i, (kd, pk) in enumerate(
            zip(le_pekerjaan.transform(le_pekerjaan.classes_), le_pekerjaan.classes_), 4):
        bg = "FFFFFF" if i % 2 == 0 else "ecfdf5"
        _style_data(wm.cell(row=i, column=1, value=pk), bg, align='left')
        _style_data(wm.cell(row=i, column=2, value=int(kd)), bg)

    rs = 4 + len(le_pekerjaan.classes_) + 2
    _style_header(wm.cell(row=rs, column=1, value="Status Rumah"), "059669")
    _style_header(wm.cell(row=rs, column=2, value="Kode"), "059669")
    for i, (s, k) in enumerate([("Kontrak/sewa", 0), ("Milik Sendiri", 1)], rs + 1):
        bg = "FFFFFF" if i % 2 == 0 else "ecfdf5"
        _style_data(wm.cell(row=i, column=1, value=s), bg, align='left')
        _style_data(wm.cell(row=i, column=2, value=k), bg)

    rs2 = rs + 4
    _style_header(wm.cell(row=rs2, column=1, value="Label"), "059669")
    _style_header(wm.cell(row=rs2, column=2, value="Kode"), "059669")
    for i, (lb, k) in enumerate([("Tidak", 0), ("Ya", 1)], rs2 + 1):
        bg = "FFFFFF" if i % 2 == 0 else "ecfdf5"
        _style_data(wm.cell(row=i, column=1, value=lb), bg, align='left')
        _style_data(wm.cell(row=i, column=2, value=k), bg)

    wm.column_dimensions['A'].width = 24
    wm.column_dimensions['B'].width = 10

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def export_results_excel(res: dict) -> bytes:
    """Export hasil prediksi ke Excel."""
    buf = io.BytesIO()
    res['results_df'].to_excel(buf, index=False)
    return buf.getvalue()


def export_evaluation_excel(res: dict) -> bytes:
    """Export ringkasan evaluasi ke Excel."""
    rows = [
        ('Accuracy Training',         res['train_acc']),
        ('Accuracy Testing',          res['test_acc']),
        ('Precision (Tidak Penerima)', res['prec0']),
        ('Recall (Tidak Penerima)',    res['rec0']),
        ('F1-Score (Tidak Penerima)', res['f1_0']),
        ('Precision (Penerima)',       res['prec1']),
        ('Recall (Penerima)',          res['rec1']),
        ('F1-Score (Penerima)',        res['f1_1']),
        ('Macro Avg Precision',        res['macro_prec']),
        ('Macro Avg Recall',           res['macro_rec']),
        ('Macro Avg F1-Score',         res['macro_f1']),
        ('Weighted Avg Precision',     res['w_prec']),
        ('Weighted Avg Recall',        res['w_rec']),
        ('Weighted Avg F1-Score',      res['w_f1']),
    ]
    df_ev = pd.DataFrame(rows, columns=['Metrik', 'Nilai'])
    df_ev['Persentase'] = df_ev['Nilai'].apply(lambda x: f"{x*100:.2f}%")
    buf = io.BytesIO()
    df_ev.to_excel(buf, index=False)
    return buf.getvalue()


def export_params_excel(res: dict) -> bytes:
    """Export parameter Naive Bayes ke Excel."""
    model = res['model']
    fc    = res['feature_cols']
    buf   = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        pd.DataFrame({
            'Kelas': ['Tidak Penerima (0)', 'Penerima (1)'],
            'Prior Probability': model.class_prior_,
            'Persentase': [f"{p*100:.2f}%" for p in model.class_prior_]
        }).to_excel(writer, sheet_name='Prior Probability', index=False)

        pd.DataFrame({
            'Fitur': fc,
            'Tidak Penerima (0)': model.theta_[0],
            'Penerima (1)': model.theta_[1]
        }).to_excel(writer, sheet_name='Mean', index=False)

        pd.DataFrame({
            'Fitur': fc,
            'Tidak Penerima (0)': np.sqrt(model.var_[0]),
            'Penerima (1)': np.sqrt(model.var_[1])
        }).to_excel(writer, sheet_name='Std Deviation', index=False)

        pd.DataFrame({
            'Fitur': fc,
            'Tidak Penerima (0)': model.var_[0],
            'Penerima (1)': model.var_[1]
        }).to_excel(writer, sheet_name='Variance', index=False)
    return buf.getvalue()


# ─────────────────────────────────────────────────────────────
# SIDEBAR NAVIGATION
# ─────────────────────────────────────────────────────────────

with st.sidebar:
    st.markdown("""
    <div class="sidebar-logo">
        <div class="sidebar-logo-icon">🎓</div>
        <div class="sidebar-logo-title">Klasifikasi BSM</div>
        <div class="sidebar-logo-sub">Naive Bayes · Dashboard ML</div>
    </div>
    <div class="sidebar-divider"></div>
    """, unsafe_allow_html=True)

    st.markdown('<div class="sidebar-section-label">Navigasi</div>', unsafe_allow_html=True)

    page = st.radio(
        "Menu",
        [
            "🏠  Dashboard",
            "📂  Upload Dataset",
            "⚙️  Preprocessing",
            "📊  Visualisasi",
            "🤖  Training Model",
            "📈  Evaluasi",
            "🗂️  Hasil Prediksi",
            "💾  Export Hasil",
            "ℹ️  Tentang Sistem",
        ],
        label_visibility="collapsed"
    )

    st.markdown('<div class="sidebar-divider"></div>', unsafe_allow_html=True)

    # Status dataset
    st.markdown('<div class="sidebar-section-label">Status</div>', unsafe_allow_html=True)
    ds_loaded  = "df_processed" in st.session_state
    model_done = "train_results" in st.session_state

    st.markdown(f"""
    <div style="display:flex;flex-direction:column;gap:6px;padding:0 4px;">
        <div style="display:flex;align-items:center;gap:8px;font-size:0.8rem;">
            {'<span style="color:#34d399;">●</span>' if ds_loaded else '<span style="color:#94a3b8;">○</span>'}
            <span>Dataset {'dimuat' if ds_loaded else 'belum dimuat'}</span>
        </div>
        <div style="display:flex;align-items:center;gap:8px;font-size:0.8rem;">
            {'<span style="color:#34d399;">●</span>' if ds_loaded else '<span style="color:#94a3b8;">○</span>'}
            <span>Preprocessing {'selesai' if ds_loaded else 'menunggu'}</span>
        </div>
        <div style="display:flex;align-items:center;gap:8px;font-size:0.8rem;">
            {'<span style="color:#34d399;">●</span>' if model_done else '<span style="color:#94a3b8;">○</span>'}
            <span>Model {'terlatih' if model_done else 'belum dilatih'}</span>
        </div>
    </div>
    """, unsafe_allow_html=True)

    if ds_loaded:
        st.markdown('<div class="sidebar-divider"></div>', unsafe_allow_html=True)
        st.markdown('<div class="sidebar-section-label">Pengaturan Training</div>', unsafe_allow_html=True)
        test_size = st.slider("Ukuran Data Uji (%)", 10, 40, 20, step=5) / 100
        st.session_state['test_size'] = test_size


# ─────────────────────────────────────────────────────────────
# PAGE: DASHBOARD
# ─────────────────────────────────────────────────────────────

if page == "🏠  Dashboard":
    render_hero(
        "Dashboard Klasifikasi BSM",
        "Sistem Machine Learning Naive Bayes untuk Klasifikasi Penerima Bantuan Siswa Miskin",
        "Machine Learning · Naive Bayes ·"
    )

    # Quick stats jika dataset sudah dimuat
    if "df_processed" in st.session_state:
        df_p = st.session_state['df_processed']
        n_total = len(df_p)
        n_ya    = int((df_p['LABEL_ENC'] == 1).sum())
        n_tdk   = int((df_p['LABEL_ENC'] == 0).sum())
        render_metric_cards([
            {"icon": "🗃️", "label": "Total Data",      "value": str(n_total), "desc": "baris setelah preprocessing"},
            {"icon": "✅", "label": "Penerima BSM",    "value": str(n_ya),    "desc": f"{n_ya/n_total*100:.1f}% dari total"},
            {"icon": "❌", "label": "Tidak Penerima",  "value": str(n_tdk),   "desc": f"{n_tdk/n_total*100:.1f}% dari total"},
            {"icon": "📐", "label": "Jumlah Fitur",    "value": "5",          "desc": "fitur input model"},
        ])
        if "train_results" in st.session_state:
            res = st.session_state['train_results']
            st.markdown("<br>", unsafe_allow_html=True)
            render_metric_cards([
                {"icon": "🎯", "label": "Akurasi Training", "value": f"{res['train_acc']*100:.1f}%", "desc": "data latih"},
                {"icon": "🏁", "label": "Akurasi Testing",  "value": f"{res['test_acc']*100:.1f}%",  "desc": "data uji"},
                {"icon": "📊", "label": "F1-Score (Macro)", "value": f"{res['macro_f1']*100:.1f}%",  "desc": "rata-rata semua kelas"},
                {"icon": "🔬", "label": "Algoritma",        "value": "GNB",                          "desc": "Gaussian Naive Bayes"},
            ])
    else:
        render_metric_cards([
            {"icon": "🗃️", "label": "Total Data",     "value": "--", "desc": "Upload dataset terlebih dahulu"},
            {"icon": "✅", "label": "Penerima BSM",   "value": "--", "desc": ""},
            {"icon": "📐", "label": "Jumlah Fitur",   "value": "5",  "desc": "fitur input model"},
            {"icon": "🔬", "label": "Algoritma",      "value": "GNB","desc": "Gaussian Naive Bayes"},
        ])

    st.markdown("<br>", unsafe_allow_html=True)
    render_section("📋", "Alur Kerja Sistem")

    steps = [
        ("1", "Upload Dataset",   "Unggah file Excel (.xlsx) data BSM 2022-2024"),
        ("2", "Preprocessing",    "Missing value, duplikasi, encoding, standardisasi Z-score"),
        ("3", "Visualisasi",      "Eksplorasi distribusi data dengan grafik Plotly interaktif"),
        ("4", "Training Model",   "Latih Gaussian Naive Bayes dengan pembagian train/test"),
        ("5", "Evaluasi",         "Accuracy, Precision, Recall, F1-Score, Confusion Matrix"),
        ("6", "Hasil Prediksi",   "Tabel prediksi beserta probabilitas dan status benar/salah"),
        ("7", "Export Hasil",     "Unduh seluruh hasil ke file Excel terformat"),
    ]
    for num, title, desc in steps:
        st.markdown(f"""
        <div class="step-card">
            <div style="display:flex;align-items:flex-start;gap:12px;">
                <div style="width:28px;height:28px;background:linear-gradient(135deg,#059669,#34d399);
                            border-radius:50%;display:flex;align-items:center;justify-content:center;
                            font-size:0.78rem;font-weight:800;color:white;flex-shrink:0;">{num}</div>
                <div>
                    <div class="step-title">{title}</div>
                    <div class="step-desc">{desc}</div>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)

    if "df_processed" not in st.session_state:
        st.markdown('<br>', unsafe_allow_html=True)
        st.markdown("""
        <div class="alert-info">
        ℹ️ Mulai dengan membuka menu <strong>Upload Dataset</strong> di sidebar untuk mengunggah file Excel Anda.
        </div>
        """, unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────
# PAGE: UPLOAD DATASET
# ─────────────────────────────────────────────────────────────

elif page == "📂  Upload Dataset":
    render_hero(
        "Upload Dataset",
        "Unggah file Excel (.xlsx) data BSM — maksimal 200MB",
        "Data Upload"
    )

    render_section("📁", "Unggah File Excel")
    uploaded = st.file_uploader(
        "Drag and drop file Excel di sini, atau klik untuk memilih",
        type=["xlsx"],
        help="Format: .xlsx | Sheet: Table 1 | Header di baris ke-2"
    )

    col1, col2 = st.columns([2, 1])

    with col1:
        if uploaded is not None:
            # Validasi & baca file
            with st.spinner("Membaca file Excel..."):
                try:
                    # Coba baca dengan format notebook asli
                    df_raw = pd.read_excel(uploaded, sheet_name=0, header=1, skiprows=[1])
                    if len(df_raw.columns) < 14:
                        # fallback: header default
                        uploaded.seek(0)
                        df_raw = pd.read_excel(uploaded, sheet_name=0)

                    # Hapus baris yang sepenuhnya kosong
                    df_raw = df_raw.dropna(how='all')
                    st.session_state['df_raw'] = df_raw

                    alert(f"File berhasil dimuat — {len(df_raw)} baris, {len(df_raw.columns)} kolom")

                    render_section("🔍", "Preview Dataset (10 Baris Pertama)")
                    st.dataframe(df_raw.head(10), use_container_width=True, height=300)

                    # Statistik dasar
                    render_section("📊", "Statistik Awal")
                    c1, c2, c3, c4 = st.columns(4)
                    with c1: st.metric("Total Baris",  len(df_raw))
                    with c2: st.metric("Total Kolom",  len(df_raw.columns))
                    with c3: st.metric("Missing Values", int(df_raw.isnull().sum().sum()))
                    with c4: st.metric("Duplikat",     int(df_raw.duplicated().sum()))

                    # Tipe data
                    with st.expander("📋 Tipe Data Kolom"):
                        dtype_df = pd.DataFrame({
                            'Kolom': df_raw.dtypes.index,
                            'Tipe Data': df_raw.dtypes.values.astype(str),
                            'Non-Null': df_raw.count().values,
                            'Null': df_raw.isnull().sum().values,
                        })
                        st.dataframe(dtype_df, use_container_width=True)

                    st.markdown("<br>", unsafe_allow_html=True)
                    if st.button("▶ Lanjut ke Preprocessing", use_container_width=True):
                        with st.spinner("Menjalankan preprocessing..."):
                            progress = st.progress(0)
                            for i in range(1, 101):
                                time.sleep(0.01)
                                progress.progress(i)
                            df_p, info = run_preprocessing(df_raw)
                            st.session_state['df_processed'] = df_p
                            st.session_state['preprocess_info'] = info
                        st.success("✅ Preprocessing selesai! Buka menu Preprocessing untuk detail.")

                except Exception as e:
                    st.error(f"❌ Error membaca file: {e}")
                    st.info("Pastikan file adalah .xlsx dengan format yang sesuai.")

    with col2:
        st.markdown("""
        <div class="card">
            <div class="card-title">📌 Format File</div>
            <div style="font-size:0.83rem;color:#475569;line-height:1.8;">
                <b>Extension:</b> .xlsx<br>
                <b>Sheet:</b> Table 1 (atau sheet pertama)<br>
                <b>Header:</b> Baris ke-2<br>
                <b>Kolom wajib:</b><br>
                · NO<br>
                · NAMA PESERTA DIDIK<br>
                · KELAS<br>
                · PENDAPATAN ORANG TUA<br>
                · PEKERJAAN ORANG TUA<br>
                · JUMLAH TANGGUNGAN<br>
                · STATUS RUMAH<br>
                · LABEL (Ya/Tidak)
            </div>
        </div>
        """, unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────
# PAGE: PREPROCESSING
# ─────────────────────────────────────────────────────────────

elif page == "⚙️  Preprocessing":
    render_hero(
        "Preprocessing Data",
        "Transformasi, encoding, dan standardisasi data sebelum pelatihan model",
        "Data Preprocessing"
    )

    if "df_processed" not in st.session_state:
        st.markdown("""
        <div class="alert-warning">
        ⚠️ Dataset belum dimuat. Silakan upload file Excel terlebih dahulu di menu <b>Upload Dataset</b>.
        </div>
        """, unsafe_allow_html=True)
    else:
        df   = st.session_state['df_processed']
        info = st.session_state['preprocess_info']
        le   = info['le_pekerjaan']

        # --- Progress steps ---
        render_section("🔄", "Tahapan Preprocessing")
        steps_done = [
            ("✅ Missing Value Check",   "Deteksi dan penanganan nilai kosong"),
            ("✅ Hapus Duplikasi",        f"{info['dup_removed']} baris duplikat dihapus"),
            ("✅ Konversi Tipe Data",     "PENDAPATAN & TANGGUNGAN → numerik"),
            ("✅ Encoding KELAS",         "Kelas 7/8/9 → 7/8/9 (ordinal)"),
            ("✅ Label Encoding Pekerjaan", f"{len(le.classes_)} kategori pekerjaan → 0–{len(le.classes_)-1}"),
            ("✅ Biner Status Rumah",     "Milik Sendiri=1, Kontrak/sewa=0"),
            ("✅ Standardisasi Z-Score", "PENDAPATAN ORANG TUA → skala standar"),
            ("✅ Encoding Label",         "Ya=1, Tidak=0 (variabel target)"),
        ]
        for i, (s, d) in enumerate(steps_done):
            pct = int((i + 1) / len(steps_done) * 100)
            st.markdown(f"""
            <div class="step-card">
                <div style="display:flex;justify-content:space-between;align-items:center;">
                    <div>
                        <div class="step-title">{s}</div>
                        <div class="step-desc">{d}</div>
                    </div>
                    <div class="badge badge-green">{pct}%</div>
                </div>
            </div>
            """, unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        tabs = st.tabs([
            "📋 Missing Value",
            "📊 Statistik",
            "🔡 Mapping Encoding",
            "📄 Data Hasil",
        ])

        # Tab 1: Missing Value
        with tabs[0]:
            render_section("❓", "Cek Missing Value")
            mv_before = pd.Series(info['missing_before'])
            mv_after  = pd.Series(info['missing_after'])
            mv_df = pd.DataFrame({
                'Kolom': mv_before.index,
                'Missing (Sebelum)': mv_before.values,
                'Missing (Sesudah)': mv_after.reindex(mv_before.index, fill_value=0).values,
            })
            st.dataframe(mv_df, use_container_width=True)

            total_mv = mv_df['Missing (Sebelum)'].sum()
            if total_mv == 0:
                alert("Tidak ada missing value dalam dataset original!")
            else:
                alert(f"Total {total_mv} missing value ditemukan dan telah ditangani.", "warning")

            st.markdown("<br>", unsafe_allow_html=True)
            c1, c2, c3 = st.columns(3)
            with c1: st.metric("Data Sebelum Dedup", info['row_after_dup'] + info['dup_removed'])
            with c2: st.metric("Duplikat Dihapus",   info['dup_removed'])
            with c3: st.metric("Data Sesudah Dedup", info['row_after_dup'])

        # Tab 2: Statistik
        with tabs[1]:
            render_section("📈", "Statistik Deskriptif")
            num_cols = ['PENDAPATAN ORANG TUA', 'JUMLAH TANGGUNGAN', 'PENDAPATAN_ZSCORE',
                        'KELAS_NUM', 'PEKERJAAN ORANG TUA_ENC', 'STATUS RUMAH_ENC', 'LABEL_ENC']
            num_cols_avail = [c for c in num_cols if c in df.columns]
            st.dataframe(df[num_cols_avail].describe().round(4), use_container_width=True)

        # Tab 3: Mapping
        with tabs[2]:
            render_section("🔡", "Mapping Encoding")
            c1, c2, c3 = st.columns(3)
            with c1:
                st.markdown("**Pekerjaan Orang Tua**")
                pek_df = pd.DataFrame({
                    'Pekerjaan': le.classes_,
                    'Kode': le.transform(le.classes_)
                })
                st.dataframe(pek_df, use_container_width=True, hide_index=True)
            with c2:
                st.markdown("**Status Rumah**")
                st.dataframe(pd.DataFrame({
                    'Status': ['Kontrak/sewa', 'Milik Sendiri'],
                    'Kode': [0, 1]
                }), use_container_width=True, hide_index=True)
            with c3:
                st.markdown("**Label (Target)**")
                st.dataframe(pd.DataFrame({
                    'Label': ['Tidak', 'Ya'],
                    'Kode': [0, 1]
                }), use_container_width=True, hide_index=True)

        # Tab 4: Data Hasil
        with tabs[3]:
            render_section("📄", "Data Setelah Preprocessing")
            show_cols = [
                'NO', 'NAMA PESERTA DIDIK', 'KELAS', 'KELAS_NUM',
                'PENDAPATAN ORANG TUA', 'PENDAPATAN_ZSCORE',
                'PEKERJAAN ORANG TUA', 'PEKERJAAN ORANG TUA_ENC',
                'JUMLAH TANGGUNGAN', 'STATUS RUMAH', 'STATUS RUMAH_ENC',
                'LABEL', 'LABEL_ENC'
            ]
            sc = [c for c in show_cols if c in df.columns]
            st.dataframe(df[sc].head(20), use_container_width=True, height=400)
            st.caption(f"Menampilkan 20 dari {len(df)} baris. Gunakan Export untuk data lengkap.")


# ─────────────────────────────────────────────────────────────
# PAGE: VISUALISASI
# ─────────────────────────────────────────────────────────────

elif page == "📊  Visualisasi":
    render_hero(
        "Visualisasi Data",
        "Eksplorasi distribusi dan pola data dengan grafik interaktif Plotly",
        "Data Visualization"
    )

    if "df_processed" not in st.session_state:
        st.markdown('<div class="alert-warning">⚠️ Dataset belum dimuat.</div>', unsafe_allow_html=True)
    else:
        df  = st.session_state['df_processed']
        th  = make_plotly_theme()
        colors = ["#059669", "#34d399", "#065f46", "#6ee7b7", "#a7f3d0", "#047857"]

        tabs = st.tabs(["🥧 Distribusi Label", "💰 Pendapatan", "💼 Pekerjaan",
                         "🏠 Status Rumah", "🎓 Kelas", "📊 Korelasi"])

        # Tab 1: Pie distribusi label
        with tabs[0]:
            render_section("🥧", "Distribusi Label Penerima BSM")
            label_counts = df['LABEL'].value_counts()
            fig = px.pie(
                names=label_counts.index,
                values=label_counts.values,
                color_discrete_sequence=["#059669", "#f87171"],
                hole=0.45,
                title="Distribusi Label: Penerima vs Tidak Penerima BSM"
            )
            fig.update_traces(
                textposition='outside',
                textinfo='label+percent+value',
                textfont_size=13,
                pull=[0.03, 0]
            )
            fig.update_layout(**th, title_font_size=15, height=420)
            st.plotly_chart(fig, use_container_width=True)

        # Tab 2: Histogram pendapatan
        with tabs[1]:
            render_section("💰", "Distribusi Pendapatan Orang Tua")
            col1, col2 = st.columns(2)
            with col1:
                fig1 = px.histogram(
                    df, x='PENDAPATAN ORANG TUA',
                    nbins=30,
                    color_discrete_sequence=["#059669"],
                    title="Histogram Pendapatan (Rupiah)",
                    labels={'PENDAPATAN ORANG TUA': 'Pendapatan (Rp)'}
                )
                fig1.update_layout(**th, height=380)
                st.plotly_chart(fig1, use_container_width=True)
            with col2:
                fig2 = px.histogram(
                    df, x='PENDAPATAN_ZSCORE',
                    nbins=30,
                    color_discrete_sequence=["#34d399"],
                    title="Histogram Pendapatan (Z-Score)",
                    labels={'PENDAPATAN_ZSCORE': 'Z-Score'}
                )
                fig2.update_layout(**th, height=380)
                st.plotly_chart(fig2, use_container_width=True)

            # Box plot per label
            fig3 = px.box(
                df, x='LABEL', y='PENDAPATAN ORANG TUA',
                color='LABEL',
                color_discrete_map={'Ya': '#059669', 'Tidak': '#f87171'},
                title="Box Plot Pendapatan per Status Penerima",
                labels={'PENDAPATAN ORANG TUA': 'Pendapatan (Rp)', 'LABEL': 'Status'}
            )
            fig3.update_layout(**th, height=380)
            st.plotly_chart(fig3, use_container_width=True)

        # Tab 3: Bar chart pekerjaan
        with tabs[2]:
            render_section("💼", "Distribusi Pekerjaan Orang Tua")
            pek_cnt = df['PEKERJAAN ORANG TUA'].value_counts().reset_index()
            pek_cnt.columns = ['Pekerjaan', 'Jumlah']
            fig = px.bar(
                pek_cnt, y='Pekerjaan', x='Jumlah',
                orientation='h',
                color='Jumlah',
                color_continuous_scale='Greens',
                title="Frekuensi Pekerjaan Orang Tua",
                text='Jumlah'
            )
            fig.update_traces(textposition='outside')
            fig.update_layout(**th, height=max(350, len(pek_cnt) * 38),
                              coloraxis_showscale=False, yaxis={'categoryorder': 'total ascending'})
            st.plotly_chart(fig, use_container_width=True)

            # Stacked per label
            pek_label = df.groupby(['PEKERJAAN ORANG TUA', 'LABEL']).size().reset_index(name='Jumlah')
            fig2 = px.bar(
                pek_label, y='PEKERJAAN ORANG TUA', x='Jumlah', color='LABEL',
                orientation='h',
                color_discrete_map={'Ya': '#059669', 'Tidak': '#f87171'},
                barmode='stack',
                title="Pekerjaan Orang Tua berdasarkan Status Penerima",
            )
            fig2.update_layout(**th, height=max(350, len(pek_cnt) * 38),
                               yaxis={'categoryorder': 'total ascending'})
            st.plotly_chart(fig2, use_container_width=True)

        # Tab 4: Status Rumah
        with tabs[3]:
            render_section("🏠", "Distribusi Status Rumah")
            col1, col2 = st.columns(2)
            with col1:
                sr_cnt = df['STATUS RUMAH'].value_counts()
                fig1 = px.pie(
                    names=sr_cnt.index, values=sr_cnt.values,
                    color_discrete_sequence=["#059669", "#34d399"],
                    hole=0.4,
                    title="Distribusi Status Rumah"
                )
                fig1.update_layout(**th, height=350)
                st.plotly_chart(fig1, use_container_width=True)
            with col2:
                sr_label = df.groupby(['STATUS RUMAH', 'LABEL']).size().reset_index(name='Jumlah')
                fig2 = px.bar(
                    sr_label, x='STATUS RUMAH', y='Jumlah', color='LABEL',
                    barmode='group',
                    color_discrete_map={'Ya': '#059669', 'Tidak': '#f87171'},
                    title="Status Rumah vs Status Penerima"
                )
                fig2.update_layout(**th, height=350)
                st.plotly_chart(fig2, use_container_width=True)

        # Tab 5: Kelas
        with tabs[4]:
            render_section("🎓", "Distribusi Kelas")
            kl_cnt = df.groupby(['KELAS', 'LABEL']).size().reset_index(name='Jumlah')
            fig = px.bar(
                kl_cnt, x='KELAS', y='Jumlah', color='LABEL',
                barmode='group',
                color_discrete_map={'Ya': '#059669', 'Tidak': '#f87171'},
                title="Distribusi Kelas berdasarkan Status Penerima"
            )
            fig.update_layout(**th, height=400)
            st.plotly_chart(fig, use_container_width=True)

            # Tanggungan
            fig2 = px.histogram(
                df, x='JUMLAH TANGGUNGAN', color='LABEL',
                barmode='overlay', opacity=0.75,
                color_discrete_map={'Ya': '#059669', 'Tidak': '#f87171'},
                title="Distribusi Jumlah Tanggungan per Status Penerima"
            )
            fig2.update_layout(**th, height=380)
            st.plotly_chart(fig2, use_container_width=True)

        # Tab 6: Korelasi
        with tabs[5]:
            render_section("📊", "Heatmap Korelasi Fitur")
            num_cols = ['KELAS_NUM', 'PENDAPATAN_ZSCORE', 'PEKERJAAN ORANG TUA_ENC',
                        'JUMLAH TANGGUNGAN', 'STATUS RUMAH_ENC', 'LABEL_ENC']
            num_avail = [c for c in num_cols if c in df.columns]
            corr = df[num_avail].corr().round(3)
            fig = px.imshow(
                corr,
                text_auto=True,
                color_continuous_scale='Greens',
                title="Correlation Heatmap",
                aspect='auto'
            )
            fig.update_layout(**th, height=500)
            st.plotly_chart(fig, use_container_width=True)


# ─────────────────────────────────────────────────────────────
# PAGE: TRAINING MODEL
# ─────────────────────────────────────────────────────────────

elif page == "🤖  Training Model":
    render_hero(
        "Training Model",
        "Pelatihan Gaussian Naive Bayes dengan pembagian data train/test",
        "Model Training"
    )

    if "df_processed" not in st.session_state:
        st.markdown('<div class="alert-warning">⚠️ Dataset belum dimuat.</div>', unsafe_allow_html=True)
    else:
        df        = st.session_state['df_processed']
        test_size = st.session_state.get('test_size', 0.2)

        # Info konfigurasi
        n_valid = df['LABEL_ENC'].notna().sum()
        n_train = int(n_valid * (1 - test_size))
        n_test  = n_valid - n_train

        render_section("⚙️", "Konfigurasi Model")
        render_metric_cards([
            {"icon": "📦", "label": "Total Data Valid",  "value": str(n_valid), "desc": "setelah drop NaN"},
            {"icon": "🏋️", "label": "Data Training",    "value": str(n_train), "desc": f"{(1-test_size)*100:.0f}% dari total"},
            {"icon": "🧪", "label": "Data Testing",      "value": str(n_test),  "desc": f"{test_size*100:.0f}% dari total"},
            {"icon": "🔬", "label": "Algoritma",         "value": "GNB",        "desc": "Gaussian Naive Bayes"},
        ])

        st.markdown("<br>", unsafe_allow_html=True)
        render_section("📐", "Fitur Input Model")
        feat_info = pd.DataFrame({
            'Fitur': ['KELAS_NUM', 'PENDAPATAN_ZSCORE', 'PEKERJAAN ORANG TUA_ENC',
                      'JUMLAH TANGGUNGAN', 'STATUS RUMAH_ENC'],
            'Deskripsi': [
                'Tingkat kelas (7, 8, atau 9) — ordinal',
                'Pendapatan orang tua yang telah di-standardisasi Z-score',
                'Pekerjaan orang tua — label encoding',
                'Jumlah tanggungan keluarga — numerik',
                'Status kepemilikan rumah — biner (0/1)'
            ],
            'Tipe': ['Ordinal', 'Kontinu (Z-Score)', 'Nominal', 'Kontinu', 'Biner']
        })
        st.dataframe(feat_info, use_container_width=True, hide_index=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # Tombol train
        if st.button("🚀 Mulai Training Model", use_container_width=True):
            with st.spinner("Melatih model Gaussian Naive Bayes..."):
                progress = st.progress(0)
                time.sleep(0.2); progress.progress(20)
                res = run_training(df, test_size=test_size)
                time.sleep(0.2); progress.progress(60)
                st.session_state['train_results'] = res
                time.sleep(0.2); progress.progress(100)

            st.success("✅ Model berhasil dilatih! Buka menu Evaluasi untuk melihat hasil.")

            # Tampil ringkasan cepat
            st.markdown("<br>", unsafe_allow_html=True)
            render_section("🎯", "Hasil Training (Ringkasan)")
            render_metric_cards([
                {"icon": "🏋️", "label": "Akurasi Training", "value": f"{res['train_acc']*100:.2f}%", "desc": f"{len(res['X_train'])} data latih"},
                {"icon": "🧪", "label": "Akurasi Testing",  "value": f"{res['test_acc']*100:.2f}%",  "desc": f"{len(res['X_test'])} data uji"},
                {"icon": "✓",  "label": "Prediksi Benar",   "value": str((res['y_test'].values == res['y_pred']).sum()), "desc": "dari data uji"},
                {"icon": "✗",  "label": "Prediksi Salah",   "value": str((res['y_test'].values != res['y_pred']).sum()), "desc": "dari data uji"},
            ])

        elif "train_results" in st.session_state:
            res = st.session_state['train_results']
            alert("Model sudah terlatih. Klik tombol di atas untuk melatih ulang.")
            render_metric_cards([
                {"icon": "🏋️", "label": "Akurasi Training", "value": f"{res['train_acc']*100:.2f}%", "desc": ""},
                {"icon": "🧪", "label": "Akurasi Testing",  "value": f"{res['test_acc']*100:.2f}%",  "desc": ""},
                {"icon": "🔬", "label": "Algoritma",        "value": "GNB",                          "desc": "Gaussian Naive Bayes"},
                {"icon": "📐", "label": "Fitur",            "value": "5",                             "desc": "variabel input"},
            ])


# ─────────────────────────────────────────────────────────────
# PAGE: EVALUASI
# ─────────────────────────────────────────────────────────────

elif page == "📈  Evaluasi":
    render_hero(
        "Evaluasi Model",
        "Akurasi, Precision, Recall, F1-Score, dan Confusion Matrix",
        "Model Evaluation"
    )

    if "train_results" not in st.session_state:
        st.markdown('<div class="alert-warning">⚠️ Model belum dilatih. Buka menu Training Model.</div>',
                    unsafe_allow_html=True)
    else:
        res = st.session_state['train_results']
        th  = make_plotly_theme()

        # Metric cards utama
        render_section("🎯", "Metrik Utama")
        render_metric_cards([
            {"icon": "🏋️", "label": "Akurasi Training", "value": f"{res['train_acc']*100:.2f}%",  "desc": "data latih"},
            {"icon": "🧪", "label": "Akurasi Testing",  "value": f"{res['test_acc']*100:.2f}%",   "desc": "data uji"},
            {"icon": "📊", "label": "F1-Score Macro",   "value": f"{res['macro_f1']*100:.2f}%",   "desc": "macro average"},
            {"icon": "⚖️", "label": "F1-Score Weighted","value": f"{res['w_f1']*100:.2f}%",       "desc": "weighted average"},
        ])

        st.markdown("<br>", unsafe_allow_html=True)
        tabs = st.tabs(["📋 Tabel Metrik", "🗂️ Confusion Matrix", "📊 Grafik Evaluasi", "🧮 Parameter NB"])

        # Tab 1: Tabel
        with tabs[0]:
            render_section("📋", "Metrik Evaluasi Lengkap")
            col1, col2 = st.columns(2)
            with col1:
                st.markdown("**Per Kelas**")
                mk_df = pd.DataFrame({
                    'Metrik': ['Precision', 'Recall', 'F1-Score'],
                    'Tidak Penerima (0)': [f"{res['prec0']:.4f}", f"{res['rec0']:.4f}", f"{res['f1_0']:.4f}"],
                    'Penerima (1)':       [f"{res['prec1']:.4f}", f"{res['rec1']:.4f}", f"{res['f1_1']:.4f}"],
                })
                st.dataframe(mk_df, use_container_width=True, hide_index=True)
            with col2:
                st.markdown("**Average**")
                av_df = pd.DataFrame({
                    'Tipe': ['Macro', 'Weighted'],
                    'Precision': [f"{res['macro_prec']:.4f}", f"{res['w_prec']:.4f}"],
                    'Recall':    [f"{res['macro_rec']:.4f}", f"{res['w_rec']:.4f}"],
                    'F1-Score':  [f"{res['macro_f1']:.4f}", f"{res['w_f1']:.4f}"],
                })
                st.dataframe(av_df, use_container_width=True, hide_index=True)

            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown("**Ringkasan Lengkap**")
            full_summary = pd.DataFrame({
                'Metrik': [
                    'Accuracy Training', 'Accuracy Testing',
                    'Precision – Tidak Penerima', 'Recall – Tidak Penerima', 'F1-Score – Tidak Penerima',
                    'Precision – Penerima',       'Recall – Penerima',       'F1-Score – Penerima',
                    'Macro Avg Precision', 'Macro Avg Recall', 'Macro Avg F1-Score',
                    'Weighted Avg Precision', 'Weighted Avg Recall', 'Weighted Avg F1-Score',
                ],
                'Nilai': [
                    res['train_acc'], res['test_acc'],
                    res['prec0'], res['rec0'], res['f1_0'],
                    res['prec1'], res['rec1'], res['f1_1'],
                    res['macro_prec'], res['macro_rec'], res['macro_f1'],
                    res['w_prec'],     res['w_rec'],     res['w_f1'],
                ],
            })
            full_summary['Persentase'] = full_summary['Nilai'].apply(lambda x: f"{x*100:.2f}%")
            full_summary['Nilai'] = full_summary['Nilai'].round(4)
            st.dataframe(full_summary, use_container_width=True, hide_index=True)

        # Tab 2: Confusion Matrix
        with tabs[1]:
            render_section("🗂️", "Confusion Matrix")
            cm = res['cm']
            fig = go.Figure(data=go.Heatmap(
                z=cm,
                x=['Predicted: Tidak', 'Predicted: Ya'],
                y=['Actual: Tidak', 'Actual: Ya'],
                text=[[str(v) for v in row] for row in cm],
                texttemplate="%{text}",
                textfont={"size": 22, "family": "JetBrains Mono", "color": "white"},
                colorscale=[
                    [0, "#ecfdf5"], [0.4, "#34d399"], [1, "#065f46"]
                ],
                showscale=True,
                colorbar=dict(title="Jumlah")
            ))
            fig.update_layout(
                **th,
                title="Confusion Matrix – Gaussian Naive Bayes",
                title_font_size=15,
                height=420,
                xaxis_title="Prediksi",
                yaxis_title="Aktual"
            )
            st.plotly_chart(fig, use_container_width=True)

            # Interpretasi
            TP = cm[1,1]; TN = cm[0,0]; FP = cm[0,1]; FN = cm[1,0]
            c1, c2, c3, c4 = st.columns(4)
            with c1: st.metric("True Positive (TP)",  TP, help="Penerima diprediksi Penerima")
            with c2: st.metric("True Negative (TN)",  TN, help="Tidak diprediksi Tidak")
            with c3: st.metric("False Positive (FP)", FP, help="Tidak diprediksi Penerima")
            with c4: st.metric("False Negative (FN)", FN, help="Penerima diprediksi Tidak")

        # Tab 3: Grafik Evaluasi
        with tabs[2]:
            render_section("📊", "Grafik Evaluasi")

            # Bar chart semua metrik
            all_metrics = {
                'Accuracy Training': res['train_acc'],
                'Accuracy Testing':  res['test_acc'],
                'Precision (0)': res['prec0'], 'Recall (0)': res['rec0'], 'F1 (0)': res['f1_0'],
                'Precision (1)': res['prec1'], 'Recall (1)': res['rec1'], 'F1 (1)': res['f1_1'],
                'Macro P':   res['macro_prec'], 'Macro R': res['macro_rec'], 'Macro F1': res['macro_f1'],
            }
            fig1 = go.Figure(go.Bar(
                x=list(all_metrics.keys()),
                y=[v * 100 for v in all_metrics.values()],
                marker_color=[
                    "#065f46", "#059669",
                    "#34d399", "#34d399", "#34d399",
                    "#10b981", "#10b981", "#10b981",
                    "#6ee7b7", "#6ee7b7", "#6ee7b7",
                ],
                text=[f"{v*100:.1f}%" for v in all_metrics.values()],
                textposition='outside'
            ))
            fig1.update_layout(**th, title="Perbandingan Semua Metrik Evaluasi (%)",
                               yaxis_title="Nilai (%)", height=400)
            st.plotly_chart(fig1, use_container_width=True)

            # Radar chart
            categories = ['Precision (0)', 'Recall (0)', 'F1 (0)', 'Precision (1)', 'Recall (1)', 'F1 (1)']
            vals = [res['prec0'], res['rec0'], res['f1_0'], res['prec1'], res['rec1'], res['f1_1']]
            fig2 = go.Figure(go.Scatterpolar(
                r=vals + [vals[0]],
                theta=categories + [categories[0]],
                fill='toself',
                line_color='#059669',
                fillcolor='rgba(5,150,105,0.2)',
                name='Model Performance'
            ))
            fig2.update_layout(
                **th, polar=dict(radialaxis=dict(range=[0, 1], showticklabels=True)),
                title="Radar Chart Metrik Evaluasi", height=400
            )
            st.plotly_chart(fig2, use_container_width=True)

        # Tab 4: Parameter NB
        with tabs[3]:
            render_section("🧮", "Parameter Gaussian Naive Bayes")
            model = res['model']
            fc    = res['feature_cols']

            st.markdown("**Prior Probability**")
            st.dataframe(pd.DataFrame({
                'Kelas': ['Tidak Penerima (0)', 'Penerima (1)'],
                'Prior Probability': model.class_prior_.round(4),
                'Persentase': [f"{p*100:.2f}%" for p in model.class_prior_]
            }), use_container_width=True, hide_index=True)

            col1, col2, col3 = st.columns(3)
            with col1:
                st.markdown("**Mean (μ) per Fitur**")
                st.dataframe(pd.DataFrame({
                    'Fitur': fc,
                    'Tidak (0)': model.theta_[0].round(4),
                    'Ya (1)':    model.theta_[1].round(4),
                }), use_container_width=True, hide_index=True)
            with col2:
                st.markdown("**Std Dev (σ) per Fitur**")
                st.dataframe(pd.DataFrame({
                    'Fitur': fc,
                    'Tidak (0)': np.sqrt(model.var_[0]).round(4),
                    'Ya (1)':    np.sqrt(model.var_[1]).round(4),
                }), use_container_width=True, hide_index=True)
            with col3:
                st.markdown("**Variance (σ²) per Fitur**")
                st.dataframe(pd.DataFrame({
                    'Fitur': fc,
                    'Tidak (0)': model.var_[0].round(4),
                    'Ya (1)':    model.var_[1].round(4),
                }), use_container_width=True, hide_index=True)


# ─────────────────────────────────────────────────────────────
# PAGE: HASIL PREDIKSI
# ─────────────────────────────────────────────────────────────

elif page == "🗂️  Hasil Prediksi":
    render_hero(
        "Hasil Prediksi",
        "Tabel prediksi lengkap dengan probabilitas dan status benar/salah",
        "Prediction Results"
    )

    if "train_results" not in st.session_state:
        st.markdown('<div class="alert-warning">⚠️ Model belum dilatih.</div>', unsafe_allow_html=True)
    else:
        res = st.session_state['train_results']
        df_r = res['results_df']

        # Summary
        benar = (df_r['Status'] == '✓ Benar').sum()
        salah = len(df_r) - benar
        render_metric_cards([
            {"icon": "📋", "label": "Total Data Uji",    "value": str(len(df_r)), "desc": ""},
            {"icon": "✓",  "label": "Prediksi Benar",   "value": str(benar),     "desc": f"{benar/len(df_r)*100:.1f}%"},
            {"icon": "✗",  "label": "Prediksi Salah",   "value": str(salah),     "desc": f"{salah/len(df_r)*100:.1f}%"},
            {"icon": "🎯", "label": "Akurasi",           "value": f"{benar/len(df_r)*100:.2f}%", "desc": ""},
        ])

        st.markdown("<br>", unsafe_allow_html=True)
        render_section("🗂️", "Tabel Hasil Prediksi")

        # Filter
        col1, col2, col3 = st.columns(3)
        with col1:
            filter_status = st.selectbox("Filter Status", ["Semua", "✓ Benar", "✗ Salah"])
        with col2:
            filter_actual = st.selectbox("Filter Actual", ["Semua", "Penerima", "Tidak Penerima"])
        with col3:
            filter_pred = st.selectbox("Filter Predicted", ["Semua", "Penerima", "Tidak Penerima"])

        df_disp = df_r.copy()
        if filter_status != "Semua":
            df_disp = df_disp[df_disp['Status'] == filter_status]
        if filter_actual != "Semua":
            df_disp = df_disp[df_disp['Actual'] == filter_actual]
        if filter_pred != "Semua":
            df_disp = df_disp[df_disp['Predicted'] == filter_pred]

        st.caption(f"Menampilkan {len(df_disp)} dari {len(df_r)} prediksi")
        st.dataframe(df_disp, use_container_width=True, height=450, hide_index=True)

        # Distribusi probabilitas
        st.markdown("<br>", unsafe_allow_html=True)
        render_section("📊", "Distribusi Probabilitas Prediksi")
        th = make_plotly_theme()

        col1, col2 = st.columns(2)
        with col1:
            fig1 = px.histogram(
                df_r, x='Prob Ya (%)', color='Actual',
                nbins=20, opacity=0.75, barmode='overlay',
                color_discrete_map={'Penerima': '#059669', 'Tidak Penerima': '#f87171'},
                title="Distribusi Probabilitas 'Ya' per Kelas Aktual",
                labels={'Prob Ya (%)': 'Probabilitas Ya (%)'}
            )
            fig1.update_layout(**th, height=350)
            st.plotly_chart(fig1, use_container_width=True)
        with col2:
            fig2 = px.scatter(
                df_r, x='Prob Tidak (%)', y='Prob Ya (%)', color='Status',
                color_discrete_map={'✓ Benar': '#059669', '✗ Salah': '#f87171'},
                title="Scatter: Probabilitas Tidak vs Ya",
            )
            fig2.update_layout(**th, height=350)
            st.plotly_chart(fig2, use_container_width=True)


# ─────────────────────────────────────────────────────────────
# PAGE: EXPORT HASIL
# ─────────────────────────────────────────────────────────────

elif page == "💾  Export Hasil":
    render_hero(
        "Export Hasil",
        "Unduh seluruh hasil ke file Excel terformat dan siap presentasi",
        "Export & Download"
    )

    if "df_processed" not in st.session_state:
        st.markdown('<div class="alert-warning">⚠️ Dataset belum dimuat.</div>', unsafe_allow_html=True)
    else:
        df   = st.session_state['df_processed']
        info = st.session_state['preprocess_info']
        le   = info['le_pekerjaan']

        render_section("💾", "Pilihan Export")

        c1, c2 = st.columns(2)

        with c1:
            st.markdown("""
            <div class="card">
                <div class="card-title">📊 Data Preprocessing</div>
                <p style="font-size:0.84rem;color:#475569;margin-bottom:1rem;">
                    Seluruh data setelah preprocessing beserta mapping encoding dalam 2 sheet Excel terformat.
                </p>
            </div>
            """, unsafe_allow_html=True)
            preproc_bytes = export_preprocessing_excel(df, le)
            st.download_button(
                "⬇️ Download Data Preprocessing (.xlsx)",
                data=preproc_bytes,
                file_name="data_bsm_preprocessed.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        with c2:
            if "train_results" in st.session_state:
                res = st.session_state['train_results']

                st.markdown("""
                <div class="card">
                    <div class="card-title">🎯 Hasil Prediksi</div>
                    <p style="font-size:0.84rem;color:#475569;margin-bottom:1rem;">
                        Tabel prediksi dengan kolom Actual, Predicted, probabilitas, dan status benar/salah.
                    </p>
                </div>
                """, unsafe_allow_html=True)
                st.download_button(
                    "⬇️ Download Hasil Prediksi (.xlsx)",
                    data=export_results_excel(res),
                    file_name="hasil_prediksi_naive_bayes.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            else:
                st.markdown('<div class="alert-info">ℹ️ Latih model terlebih dahulu untuk mengekspor hasil prediksi.</div>',
                            unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        c3, c4 = st.columns(2)

        with c3:
            if "train_results" in st.session_state:
                res = st.session_state['train_results']
                st.markdown("""
                <div class="card">
                    <div class="card-title">📈 Ringkasan Evaluasi</div>
                    <p style="font-size:0.84rem;color:#475569;margin-bottom:1rem;">
                        Seluruh metrik evaluasi: accuracy, precision, recall, F1-score, macro, dan weighted average.
                    </p>
                </div>
                """, unsafe_allow_html=True)
                st.download_button(
                    "⬇️ Download Ringkasan Evaluasi (.xlsx)",
                    data=export_evaluation_excel(res),
                    file_name="ringkasan_evaluasi_naive_bayes.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

        with c4:
            if "train_results" in st.session_state:
                res = st.session_state['train_results']
                st.markdown("""
                <div class="card">
                    <div class="card-title">🧮 Parameter Model</div>
                    <p style="font-size:0.84rem;color:#475569;margin-bottom:1rem;">
                        Prior probability, mean, standard deviation, dan variance per fitur per kelas.
                    </p>
                </div>
                """, unsafe_allow_html=True)
                st.download_button(
                    "⬇️ Download Parameter Model (.xlsx)",
                    data=export_params_excel(res),
                    file_name="parameter_naive_bayes.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

        st.markdown("<br>", unsafe_allow_html=True)
        alert("File Excel yang diunduh sudah diformat rapi.")


# ─────────────────────────────────────────────────────────────
# PAGE: TENTANG SISTEM
# ─────────────────────────────────────────────────────────────

elif page == "ℹ️  Tentang Sistem":
    render_hero(
        "Tentang Sistem",
        "Informasi lengkap mengenai sistem klasifikasi penerima BSM",
        "About"
    )

    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("""
        <div class="about-card">
            <div class="about-card-icon">🎓</div>
            <div class="about-card-title">Tujuan Sistem</div>
            <div class="about-card-desc">
                Membantu pengambilan keputusan pemberian Bantuan Siswa Miskin (BSM) menggunakan 
                algoritma Machine Learning Gaussian Naive Bayes secara otomatis, transparan, dan terukur.
            </div>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown("""
        <div class="about-card">
            <div class="about-card-icon">🔬</div>
            <div class="about-card-title">Algoritma</div>
            <div class="about-card-desc">
                <b>Gaussian Naive Bayes</b> — algoritma probabilistik berbasis Teorema Bayes yang 
                mengasumsikan distribusi Gaussian pada setiap fitur numerik. Cocok untuk dataset 
                dengan fitur kontinu dan kategorikal.
            </div>
        </div>
        """, unsafe_allow_html=True)
    with col3:
        st.markdown("""
        <div class="about-card">
            <div class="about-card-icon">📊</div>
            <div class="about-card-title">Fitur Sistem</div>
            <div class="about-card-desc">
                Upload Excel → Preprocessing otomatis → Visualisasi interaktif → 
                Training & Evaluasi → Tabel prediksi → Export Excel terformat.
                Semua dalam satu dashboard modern.
            </div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    render_section("🛠️", "Teknologi yang Digunakan")

    tech = [
        ("🐍 Python 3.x",        "Bahasa pemrograman utama"),
        ("⚡ Streamlit",          "Framework web dashboard interaktif"),
        ("🐼 Pandas",             "Manipulasi dan analisis data"),
        ("🔢 NumPy",              "Komputasi numerik"),
        ("🤖 Scikit-learn",       "Implementasi Gaussian Naive Bayes & metrik evaluasi"),
        ("📈 Plotly",             "Visualisasi data interaktif"),
        ("📗 Openpyxl",           "Pembuatan dan ekspor file Excel terformat"),
    ]
    for name, desc in tech:
        st.markdown(f"""
        <div class="step-card">
            <div style="display:flex;align-items:center;gap:12px;">
                <div style="font-size:1.2rem;">{name.split(' ')[0]}</div>
                <div>
                    <div class="step-title">{' '.join(name.split(' ')[1:])}</div>
                    <div class="step-desc">{desc}</div>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    render_section("📚", "Fitur Input Model")
    feat_df = pd.DataFrame({
        'No.': range(1, 6),
        'Nama Fitur': ['KELAS_NUM', 'PENDAPATAN_ZSCORE', 'PEKERJAAN ORANG TUA_ENC',
                       'JUMLAH TANGGUNGAN', 'STATUS RUMAH_ENC'],
        'Tipe Transformasi': ['Ordinal Encoding', 'Z-Score Standardization',
                              'Label Encoding', 'Numerik (asli)', 'Binary Encoding'],
        'Keterangan': [
            'Tingkat kelas siswa (7, 8, atau 9)',
            'Pendapatan orang tua yang distandarisasi',
            'Jenis pekerjaan orang tua (diurutkan alfabet)',
            'Jumlah anggota keluarga yang ditanggung',
            'Milik Sendiri=1, Kontrak/sewa=0'
        ]
    })
    st.dataframe(feat_df, use_container_width=True, hide_index=True)

    st.markdown("<br>", unsafe_allow_html=True)
    render_section("📖", "Cara Penggunaan")
    cara = [
        "Buka menu **Upload Dataset** → unggah file Excel BSM (.xlsx)",
        "Klik tombol **Lanjut ke Preprocessing** — sistem akan otomatis memproses data",
        "Buka **Visualisasi** untuk eksplorasi distribusi data",
        "Buka **Training Model** → sesuaikan ukuran data uji → klik **Mulai Training**",
        "Buka **Evaluasi** untuk melihat accuracy, confusion matrix, dan metrik lengkap",
        "Buka **Hasil Prediksi** untuk tabel prediksi data uji",
        "Buka **Export Hasil** → unduh semua hasil ke file Excel",
    ]
    for i, c in enumerate(cara, 1):
        st.markdown(f"""
        <div class="step-card">
            <div style="display:flex;align-items:flex-start;gap:12px;">
                <div style="width:24px;height:24px;background:#059669;border-radius:50%;
                            display:flex;align-items:center;justify-content:center;
                            font-size:0.72rem;font-weight:800;color:white;flex-shrink:0;">{i}</div>
                <div style="font-size:0.88rem;color:#334155;padding-top:2px;">{c}</div>
            </div>
        </div>
        """, unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────
# FOOTER MODERN
# ─────────────────────────────────────────────────────────────

st.markdown("""
<div class="footer">
    <div class="footer-brand">🎓 Sistem Klasifikasi Penerima BSM</div>
    <div class="footer-divider"></div>
    <div>Gaussian Naive Bayes · Machine Learning Dashboard · Streamlit</div>
    <div style="margin-top:0.4rem;color:#64748b;font-size:0.75rem;">
        Made with ❤️
    </div>
</div>
""", unsafe_allow_html=True)
