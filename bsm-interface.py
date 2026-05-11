# ============================================================
# SISTEM KLASIFIKASI PENERIMA BSM - NAIVE BAYES
# Dashboard Machine Learning Modern | Streamlit App
# Dengan Detail Perhitungan di UI + Terminal
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
# CUSTOM CSS
# ─────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;600&display=swap');

:root {
    --em-50:#ecfdf5;--em-100:#d1fae5;--em-200:#a7f3d0;--em-300:#6ee7b7;
    --em-400:#34d399;--em-500:#10b981;--em-600:#059669;--em-700:#047857;
    --em-800:#065f46;--em-900:#064e3b;
    --sl-50:#f8fafc;--sl-100:#f1f5f9;--sl-200:#e2e8f0;--sl-400:#94a3b8;
    --sl-500:#64748b;--sl-600:#475569;--sl-700:#334155;--sl-800:#1e293b;
    --white:#ffffff;
    --r:14px;--r-sm:8px;
    --sh:0 2px 12px rgba(0,0,0,0.07);
}

html,body{font-family:'Plus Jakarta Sans',sans-serif;}
.main .block-container{padding:1.5rem 2rem 3rem;max-width:1400px;}

[data-testid="stSidebar"]>div:first-child{
    background:linear-gradient(175deg,#064e3b 0%,#065f46 55%,#1e293b 100%);
}
[data-testid="stSidebar"] p, [data-testid="stSidebar"] span {color:#ffffff;}
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

.hero-header{
    background:linear-gradient(135deg,#065f46 0%,#059669 55%,#34d399 100%);
    border-radius:var(--r);padding:2rem 2.4rem;margin-bottom:1.6rem;
    position:relative;overflow:hidden;
    box-shadow:0 4px 20px rgba(16,185,129,0.2);
}
.hero-badge{
    display:inline-block;background:rgba(255,255,255,0.16);color:#fff;
    border:1px solid rgba(255,255,255,0.25);border-radius:999px;
    padding:3px 12px;font-size:0.72rem;font-weight:600;
    letter-spacing:0.07em;text-transform:uppercase;margin-bottom:10px;
}
.hero-title{font-size:1.85rem;font-weight:800;color:#fff!important;margin:0 0 4px;line-height:1.2;}
.hero-sub{font-size:0.9rem;color:#d1fae5!important;margin:0;}

.metric-card{
    background:var(--white);border-radius:var(--r);padding:1.3rem 1.4rem;
    box-shadow:var(--sh);border:1px solid var(--sl-200);
    position:relative;overflow:hidden;transition:transform 0.18s,box-shadow 0.18s;
}
.metric-card::before{
    content:'';position:absolute;top:0;left:0;right:0;height:3px;
    background:linear-gradient(90deg,var(--em-600),var(--em-400));
}
.metric-card:hover{transform:translateY(-2px);box-shadow:0 6px 24px rgba(0,0,0,0.12);}
.metric-icon{font-size:1.6rem;margin-bottom:6px;}
.metric-label{font-size:0.7rem;font-weight:700;color:var(--sl-400);text-transform:uppercase;letter-spacing:0.08em;margin-bottom:4px;}
.metric-value{font-size:1.9rem;font-weight:800;color:var(--em-700);line-height:1;font-family:'JetBrains Mono',monospace;}
.metric-desc{font-size:0.74rem;color:var(--sl-400);margin-top:4px;}

.section-header{display:flex;align-items:center;gap:10px;margin:1.6rem 0 1rem;}
.section-icon{width:36px;height:36px;background:var(--em-100);border-radius:var(--r-sm);display:flex;align-items:center;justify-content:center;font-size:1.05rem;flex-shrink:0;}
.section-title{font-size:1.1rem;font-weight:700;color:var(--sl-800);margin:0;}
.section-divider{height:2px;border-radius:2px;margin-bottom:1.2rem;background:linear-gradient(90deg,var(--em-500) 0%,transparent 100%);opacity:0.35;}

.card{background:var(--white);border-radius:var(--r);padding:1.4rem;box-shadow:var(--sh);border:1px solid var(--sl-200);margin-bottom:1rem;}
.card-title{font-size:0.78rem;font-weight:700;color:var(--sl-600);text-transform:uppercase;letter-spacing:0.07em;margin-bottom:0.9rem;padding-bottom:0.6rem;border-bottom:1px solid var(--sl-100);}

.step-card{background:var(--white);border-radius:var(--r);padding:1.1rem 1.4rem;border-left:4px solid var(--em-500);box-shadow:var(--sh);margin-bottom:0.7rem;}
.step-title{font-weight:700;color:var(--sl-700);font-size:0.88rem;margin-bottom:2px;}
.step-desc{font-size:0.78rem;color:var(--sl-400);}

.alert-success{background:#f0fdf4;border:1px solid #bbf7d0;border-radius:var(--r-sm);padding:10px 14px;color:#166534;font-size:0.84rem;font-weight:500;margin:6px 0;}
.alert-info{background:#eff6ff;border:1px solid #bfdbfe;border-radius:var(--r-sm);padding:10px 14px;color:#1e40af;font-size:0.84rem;font-weight:500;margin:6px 0;}
.alert-warning{background:#fffbeb;border:1px solid #fde68a;border-radius:var(--r-sm);padding:10px 14px;color:#92400e;font-size:0.84rem;font-weight:500;margin:6px 0;}

.badge{display:inline-block;border-radius:999px;padding:2px 10px;font-size:0.7rem;font-weight:700;}
.badge-green{background:var(--em-100);color:var(--em-800);}
.badge-red{background:#fee2e2;color:#991b1b;}

/* Style untuk terminal output */
.terminal-output {
    background: #1a1a2e;
    color: #00ff88;
    font-family: 'JetBrains Mono', 'Courier New', monospace;
    font-size: 0.82rem;
    padding: 1.2rem;
    border-radius: 8px;
    line-height: 1.6;
    white-space: pre-wrap;
    word-break: break-all;
    max-height: 600px;
    overflow-y: auto;
    border: 1px solid #333;
}

.terminal-output .highlight {
    color: #ffcc00;
    font-weight: bold;
}

.terminal-output .error {
    color: #ff4444;
}

.terminal-output .success {
    color: #00ff88;
    font-weight: bold;
}

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

.stProgress>div>div>div{background:linear-gradient(90deg,var(--em-600),var(--em-400))!important;border-radius:999px!important;}

code,pre{font-family:'JetBrains Mono',monospace!important;}

::-webkit-scrollbar{width:5px;height:5px;}
::-webkit-scrollbar-track{background:var(--sl-100);}
::-webkit-scrollbar-thumb{background:var(--em-300);border-radius:3px;}

#MainMenu{visibility:hidden;}
footer{visibility:hidden;}
.stDeployButton{display:none;}

.sidebar-logo{text-align:center;padding:1.2rem 1rem 0.8rem;}
.sidebar-logo-icon{font-size:2.5rem;margin-bottom:6px;}
.sidebar-logo-title{font-size:0.95rem;font-weight:800;color:#fff;line-height:1.2;}
.sidebar-logo-sub{font-size:0.7rem;color:#6ee7b7;margin-top:2px;}
.sidebar-divider{height:1px;background:rgba(255,255,255,0.1);margin:10px 0;}
.sidebar-section-label{font-size:0.62rem;color:#34d399;text-transform:uppercase;letter-spacing:0.14em;font-weight:700;padding:0 4px;margin-bottom:4px;}

.footer{background:var(--sl-800);color:#cbd5e1;border-radius:var(--r);padding:1.4rem 2rem;text-align:center;margin-top:2.5rem;font-size:0.8rem;line-height:1.8;}
.footer-brand{color:var(--em-400);font-weight:700;font-size:0.95rem;}
.footer-divider{width:36px;height:2px;background:var(--em-500);margin:8px auto;border-radius:2px;}
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────
# HELPER FUNCTIONS
# ─────────────────────────────────────────────────────────────

def render_hero(title: str, subtitle: str, badge: str = "Machine Learning"):
    st.markdown(f"""
    <div class="hero-header">
        <div class="hero-badge">🎓 {badge}</div>
        <div class="hero-title">{title}</div>
        <div class="hero-sub">{subtitle}</div>
    </div>
    """, unsafe_allow_html=True)


def render_metric_cards(metrics: list):
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
    return dict(
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Plus Jakarta Sans, sans-serif", color="#334155"),
        margin=dict(l=20, r=20, t=40, b=20),
    )


# ═════════════════════════════════════════════════════════════
# FUNGSI LOG COLLECTOR (Menggabungkan terminal + UI)
# ═════════════════════════════════════════════════════════════

class LogCollector:
    """Mengumpulkan output log untuk ditampilkan di terminal DAN UI Streamlit."""
    
    def __init__(self):
        self.logs = []
    
    def add(self, text: str):
        self.logs.append(text)
        print(text)  # Tetap print ke terminal
    
    def add_separator(self, char="=", length=80):
        self.add(char * length)
    
    def add_header(self, title: str):
        self.add_separator("=")
        self.add(title)
        self.add_separator("=")
    
    def add_blank(self):
        self.add("")
    
    def get_all(self) -> str:
        return "\n".join(self.logs)
    
    def clear(self):
        self.logs = []


# ═════════════════════════════════════════════════════════════
# FUNGSI PERHITUNGAN DENGAN LOGGING
# ═════════════════════════════════════════════════════════════

def print_preprocessing_info(df, info, log: LogCollector):
    """Menampilkan detail preprocessing ke terminal dan mengumpulkan log."""
    log.add_header("PREPROCESSING DATA")
    
    log.add(f"\n[1] DATA AWAL")
    log.add(f"    Jumlah data awal : {info['row_after_dup'] + info['dup_removed']} baris")
    
    log.add(f"\n[2] MISSING VALUE")
    log.add(f"    Missing value sebelum penanganan:")
    for col, missing in info['missing_before'].items():
        if missing > 0:
            log.add(f"    - {col}: {missing}")
    
    log.add(f"\n    Missing value setelah penanganan:")
    for col, missing in info['missing_after'].items():
        if missing > 0:
            log.add(f"    - {col}: {missing}")
    
    log.add(f"\n[3] DUPLIKASI DATA")
    log.add(f"    Data duplikat dihapus : {info['dup_removed']} baris")
    log.add(f"    Data setelah dedup   : {info['row_after_dup']} baris")
    
    log.add(f"\n[4] MAPPING ENCODING")
    
    le = info['le_pekerjaan']
    log.add(f"\n    Pekerjaan Orang Tua → Label Encoding:")
    for kd, pk in zip(le.transform(le.classes_), le.classes_):
        log.add(f"    {pk:30s} → {kd}")
    
    log.add(f"\n    Status Rumah → Binary Encoding:")
    log.add(f"    {'Kontrak/sewa':30s} → 0")
    log.add(f"    {'Milik Sendiri':30s} → 1")
    
    log.add(f"\n    Label → Binary Encoding:")
    log.add(f"    {'Tidak':30s} → 0")
    log.add(f"    {'Ya':30s} → 1")
    
    log.add(f"\n[5] STANDARDISASI Z-SCORE (PENDAPATAN ORANG TUA)")
    scaler = info['scaler']
    log.add(f"    Mean   : {scaler.mean_[0]:.6f}")
    log.add(f"    Std    : {scaler.scale_[0]:.6f}")
    
    log.add(f"\n[6] 5 BARIS PERTAMA HASIL PREPROCESSING")
    cols_show = ['NO', 'NAMA PESERTA DIDIK', 'KELAS', 'KELAS_NUM',
                 'PENDAPATAN ORANG TUA', 'PENDAPATAN_ZSCORE',
                 'PEKERJAAN ORANG TUA', 'PEKERJAAN ORANG TUA_ENC',
                 'JUMLAH TANGGUNGAN', 'STATUS RUMAH', 'STATUS RUMAH_ENC',
                 'LABEL', 'LABEL_ENC']
    log.add(df[cols_show].head(5).to_string(index=False))
    
    log.add_header("PREPROCESSING SELESAI")


def print_split_info(X, y, X_train, X_test, y_train, y_test, log: LogCollector):
    """Menampilkan info pembagian data."""
    log.add_header("PEMBAGIAN DATA TRAINING DAN TESTING")
    
    log.add(f"\n    Total data valid : {len(X)}")
    log.add(f"    Data Training    : {len(X_train)} ({(len(X_train)/len(X))*100:.1f}%)")
    log.add(f"    Data Testing     : {len(X_test)} ({(len(X_test)/len(X))*100:.1f}%)")
    
    train_dist = y_train.value_counts()
    test_dist = y_test.value_counts()
    
    log.add(f"\n    Distribusi Kelas pada Data Training:")
    log.add(f"    - Tidak Penerima (0) : {train_dist.get(0, 0)}")
    log.add(f"    - Penerima (1)       : {train_dist.get(1, 0)}")
    
    log.add(f"\n    Distribusi Kelas pada Data Testing:")
    log.add(f"    - Tidak Penerima (0) : {test_dist.get(0, 0)}")
    log.add(f"    - Penerima (1)       : {test_dist.get(1, 0)}")
    
    log.add_header("PEMBAGIAN DATA SELESAI")


def print_training_info(model, feature_cols, log: LogCollector):
    """Menampilkan parameter model."""
    log.add_header("TRAINING GAUSSIAN NAIVE BAYES")
    
    log.add(f"\n[1] PRIOR PROBABILITY")
    log.add(f"    P(Tidak) = {model.class_prior_[0]:.6f} ({model.class_prior_[0]*100:.2f}%)")
    log.add(f"    P(Ya)    = {model.class_prior_[1]:.6f} ({model.class_prior_[1]*100:.2f}%)")
    
    log.add(f"\n[2] MEAN (THETA_) PER FITUR")
    log.add(f"    {'Fitur':35s} {'Tidak (0)':>12s} {'Ya (1)':>12s}")
    log.add(f"    {'-'*35} {'-'*12} {'-'*12}")
    for i, feat in enumerate(feature_cols):
        log.add(f"    {feat:35s} {model.theta_[0][i]:12.6f} {model.theta_[1][i]:12.6f}")
    
    log.add(f"\n[3] VARIANCE (VAR_) PER FITUR")
    log.add(f"    {'Fitur':35s} {'Tidak (0)':>12s} {'Ya (1)':>12s}")
    log.add(f"    {'-'*35} {'-'*12} {'-'*12}")
    for i, feat in enumerate(feature_cols):
        log.add(f"    {feat:35s} {model.var_[0][i]:12.6f} {model.var_[1][i]:12.6f}")
    
    log.add(f"\n[4] STANDARD DEVIATION (SIGMA) PER FITUR")
    log.add(f"    {'Fitur':35s} {'Tidak (0)':>12s} {'Ya (1)':>12s}")
    log.add(f"    {'-'*35} {'-'*12} {'-'*12}")
    for i, feat in enumerate(feature_cols):
        log.add(f"    {feat:35s} {np.sqrt(model.var_[0][i]):12.6f} {np.sqrt(model.var_[1][i]):12.6f}")
    
    log.add_header("TRAINING SELESAI")


def print_prediction_details(model, X_test, y_test, log: LogCollector, max_display: int = 10):
    """Menampilkan detail prediksi per data testing. max_display membatasi output."""
    log.add_header("PROSES PREDIKSI PER DATA TESTING")
    
    feature_cols = X_test.columns.tolist()
    X_test_array = X_test.values
    y_test_array = y_test.values
    
    n_display = min(len(X_test_array), max_display)
    
    for idx in range(n_display):
        log.add(f"\n{'─'*60}")
        log.add(f"DATA TEST KE-{idx+1}")
        log.add(f"{'─'*60}")
        
        x = X_test_array[idx]
        actual = y_test_array[idx]
        
        log.add(f"\n    Fitur:")
        for i, feat in enumerate(feature_cols):
            log.add(f"    - {feat} = {x[i]:.6f}")
        
        theta = model.theta_
        var = model.var_
        class_prior = model.class_prior_
        
        posteriors = []
        for c in range(2):
            log_prior = np.log(class_prior[c])
            log_likelihood = 0
            for i in range(len(feature_cols)):
                coef = 1.0 / np.sqrt(2 * np.pi * var[c][i])
                exponent = np.exp(-((x[i] - theta[c][i]) ** 2) / (2 * var[c][i]))
                likelihood = coef * exponent
                log_likelihood += np.log(max(likelihood, 1e-300))
            
            posterior_log = log_prior + log_likelihood
            posteriors.append(posterior_log)
        
        max_log = max(posteriors)
        exp_post = [np.exp(p - max_log) for p in posteriors]
        total = sum(exp_post)
        prob = [e / total for e in exp_post]
        
        log.add(f"\n    Posterior Log:")
        log.add(f"    - Log P(Tidak) = {posteriors[0]:.6f}")
        log.add(f"    - Log P(Ya)    = {posteriors[1]:.6f}")
        
        log.add(f"\n    Probabilitas Akhir:")
        log.add(f"    - Prob Tidak = {prob[0]*100:.2f}%")
        log.add(f"    - Prob Ya    = {prob[1]*100:.2f}%")
        
        pred = np.argmax(prob)
        pred_label = "Ya" if pred == 1 else "Tidak"
        actual_label = "Ya" if actual == 1 else "Tidak"
        status = "BENAR" if pred == actual else "SALAH"
        
        log.add(f"\n    Hasil:")
        log.add(f"    - Prediksi = {pred_label}")
        log.add(f"    - Aktual   = {actual_label}")
        log.add(f"    - Status   = {status}")
    
    if len(X_test_array) > max_display:
        log.add(f"\n    ... dan {len(X_test_array) - max_display} data lainnya (gunakan terminal untuk melihat semua)")
    
    log.add(f"\n{'─'*60}")
    log.add_header("PROSES PREDIKSI SELESAI")


def print_evaluation_results(results, log: LogCollector):
    """Menampilkan hasil evaluasi."""
    log.add_header("EVALUASI MODEL")
    
    cm = results['cm']
    log.add(f"\n[1] CONFUSION MATRIX")
    log.add(f"    [[TN  FP]     [[{cm[0][0]:4d}  {cm[0][1]:4d}]")
    log.add(f"     [FN  TP]]  =  [{cm[1][0]:4d}  {cm[1][1]:4d}]]")
    
    TN = cm[0][0]; FP = cm[0][1]; FN = cm[1][0]; TP = cm[1][1]
    total = TN + FP + FN + TP
    benar = TP + TN; salah = FP + FN
    
    log.add(f"\n    True Positive  (TP) : {TP}")
    log.add(f"    True Negative  (TN) : {TN}")
    log.add(f"    False Positive (FP) : {FP}")
    log.add(f"    False Negative (FN) : {FN}")
    
    accuracy = (TP + TN) / total
    precision_0 = TN / (TN + FN) if (TN + FN) > 0 else 0
    recall_0 = TN / (TN + FP) if (TN + FP) > 0 else 0
    f1_0 = 2 * precision_0 * recall_0 / (precision_0 + recall_0) if (precision_0 + recall_0) > 0 else 0
    
    precision_1 = TP / (TP + FP) if (TP + FP) > 0 else 0
    recall_1 = TP / (TP + FN) if (TP + FN) > 0 else 0
    f1_1 = 2 * precision_1 * recall_1 / (precision_1 + recall_1) if (precision_1 + recall_1) > 0 else 0
    
    log.add(f"\n[2] METRIK EVALUASI")
    log.add(f"    Accuracy : {accuracy:.6f} ({accuracy*100:.2f}%)")
    
    log.add(f"\n    Kelas Tidak Penerima (0):")
    log.add(f"    - Precision : {precision_0:.6f} ({precision_0*100:.2f}%)")
    log.add(f"    - Recall    : {recall_0:.6f} ({recall_0*100:.2f}%)")
    log.add(f"    - F1-Score  : {f1_0:.6f} ({f1_0*100:.2f}%)")
    
    log.add(f"\n    Kelas Penerima (1):")
    log.add(f"    - Precision : {precision_1:.6f} ({precision_1*100:.2f}%)")
    log.add(f"    - Recall    : {recall_1:.6f} ({recall_1*100:.2f}%)")
    log.add(f"    - F1-Score  : {f1_1:.6f} ({f1_1*100:.2f}%)")
    
    macro_prec = (precision_0 + precision_1) / 2
    macro_rec = (recall_0 + recall_1) / 2
    macro_f1 = (f1_0 + f1_1) / 2
    
    sup0 = TN + FP; sup1 = TP + FN
    w_prec = (precision_0 * sup0 + precision_1 * sup1) / total
    w_rec = (recall_0 * sup0 + recall_1 * sup1) / total
    w_f1 = (f1_0 * sup0 + f1_1 * sup1) / total
    
    log.add(f"\n    Macro Average:")
    log.add(f"    - Precision : {macro_prec:.6f} ({macro_prec*100:.2f}%)")
    log.add(f"    - Recall    : {macro_rec:.6f} ({macro_rec*100:.2f}%)")
    log.add(f"    - F1-Score  : {macro_f1:.6f} ({macro_f1*100:.2f}%)")
    
    log.add(f"\n    Weighted Average:")
    log.add(f"    - Precision : {w_prec:.6f} ({w_prec*100:.2f}%)")
    log.add(f"    - Recall    : {w_rec:.6f} ({w_rec*100:.2f}%)")
    log.add(f"    - F1-Score  : {w_f1:.6f} ({w_f1*100:.2f}%)")
    
    log.add(f"\n[3] RINGKASAN AKHIR")
    log.add(f"    Total prediksi benar : {benar} dari {total} data")
    log.add(f"    Total prediksi salah : {salah} dari {total} data")
    log.add(f"    Akurasi akhir        : {accuracy*100:.2f}%")
    
    log.add_header("EVALUASI SELESAI")


# ─────────────────────────────────────────────────────────────
# PREPROCESSING ENGINE
# ─────────────────────────────────────────────────────────────

def run_preprocessing(df_raw: pd.DataFrame):
    """Jalankan pipeline preprocessing."""
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

    # 1. Missing value
    info['missing_before'] = df.isnull().sum().to_dict()

    # 2. Konversi tipe data
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

    # LOGGING
    log = LogCollector()
    print_preprocessing_info(df, info, log)
    
    # Simpan log ke session state
    st.session_state['preprocess_log'] = log.get_all()

    return df, info


# ─────────────────────────────────────────────────────────────
# MODEL TRAINING ENGINE
# ─────────────────────────────────────────────────────────────

def run_training(df: pd.DataFrame, test_size: float = 0.2):
    """Latih model Gaussian Naive Bayes."""
    feature_cols = [
        'KELAS_NUM',
        'PENDAPATAN_ZSCORE',
        'PEKERJAAN ORANG TUA_ENC',
        'JUMLAH TANGGUNGAN',
        'STATUS RUMAH_ENC'
    ]

    X = df[feature_cols].copy()
    y = df['LABEL_ENC'].copy()

    valid_idx = X.notna().all(axis=1) & y.notna()
    X = X[valid_idx]
    y = y[valid_idx]

    X_train, X_test, y_train, y_test = train_test_split(
        X, y, test_size=test_size, random_state=42, stratify=y
    )

    # Log collector
    log = LogCollector()
    
    # Log pembagian data
    print_split_info(X, y, X_train, X_test, y_train, y_test, log)

    model = GaussianNB()
    model.fit(X_train, y_train)

    # Log parameter model
    print_training_info(model, feature_cols, log)

    y_train_pred  = model.predict(X_train)
    y_pred        = model.predict(X_test)
    y_pred_proba  = model.predict_proba(X_test)

    # Log prediksi (batasi 10 data pertama)
    print_prediction_details(model, X_test, y_test, log, max_display=10)

    train_acc = accuracy_score(y_train, y_train_pred)
    test_acc  = accuracy_score(y_test, y_pred)
    cm        = confusion_matrix(y_test, y_pred)

    TP = cm[1, 1]; TN = cm[0, 0]; FP = cm[0, 1]; FN = cm[1, 0]

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

    results_dict = {
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
    }

    # Log evaluasi
    print_evaluation_results(results_dict, log)
    
    # Simpan log ke session state
    st.session_state['training_log'] = log.get_all()

    results_df = pd.DataFrame({
        'No': range(1, len(y_test) + 1),
        'Actual':    ['Penerima' if v == 1 else 'Tidak Penerima' for v in y_test.values],
        'Predicted': ['Penerima' if v == 1 else 'Tidak Penerima' for v in y_pred],
        'Prob Tidak (%)': (y_pred_proba[:, 0] * 100).round(2),
        'Prob Ya (%)':    (y_pred_proba[:, 1] * 100).round(2),
        'Status': ['✓ Benar' if a == p else '✗ Salah'
                   for a, p in zip(y_test.values, y_pred)]
    })

    results_dict['results_df'] = results_df
    
    return results_dict


# ─────────────────────────────────────────────────────────────
# EXCEL EXPORT HELPERS (tidak berubah dari versi sebelumnya)
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
            abg = "FFF9C4" if bg == "FFFFFF" and ('Encoded' in cn or 'Z-Score' in cn or 'Angka' in cn) else ("FFF3CD" if ('Encoded' in cn or 'Z-Score' in cn or 'Angka' in cn) else bg)
            _style_data(cell, abg)
            if cn == 'Pendapatan Orang Tua (Rp)':
                cell.number_format = '#,##0'
            elif cn == 'Pendapatan (Z-Score)':
                cell.number_format = '0.000000'

    for i, w in enumerate([5, 28, 28, 10, 12, 22, 18, 22, 16, 16, 18, 18, 8, 14], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    wm = wb.create_sheet("Mapping Encoding")
    wm.merge_cells('A1:B1')
    wm['A1'].value = "Mapping Encoding"
    wm['A1'].font  = Font(name='Arial', bold=True, size=12, color='FFFFFF')
    wm['A1'].fill  = PatternFill("solid", start_color="065f46")
    
    _style_header(wm.cell(row=3, column=1, value="Pekerjaan Orang Tua"), "059669")
    _style_header(wm.cell(row=3, column=2, value="Kode"), "059669")
    for i, (kd, pk) in enumerate(zip(le_pekerjaan.transform(le_pekerjaan.classes_), le_pekerjaan.classes_), 4):
        _style_data(wm.cell(row=i, column=1, value=pk), "FFFFFF" if i % 2 == 0 else "ecfdf5", align='left')
        _style_data(wm.cell(row=i, column=2, value=int(kd)), "FFFFFF" if i % 2 == 0 else "ecfdf5")

    rs = 4 + len(le_pekerjaan.classes_) + 2
    _style_header(wm.cell(row=rs, column=1, value="Status Rumah"), "059669")
    _style_header(wm.cell(row=rs, column=2, value="Kode"), "059669")
    for i, (s, k) in enumerate([("Kontrak/sewa", 0), ("Milik Sendiri", 1)], rs + 1):
        _style_data(wm.cell(row=i, column=1, value=s), "FFFFFF" if i % 2 == 0 else "ecfdf5", align='left')
        _style_data(wm.cell(row=i, column=2, value=k), "FFFFFF" if i % 2 == 0 else "ecfdf5")

    wm.column_dimensions['A'].width = 24
    wm.column_dimensions['B'].width = 10

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def export_results_excel(res: dict) -> bytes:
    buf = io.BytesIO()
    res['results_df'].to_excel(buf, index=False)
    return buf.getvalue()


def export_evaluation_excel(res: dict) -> bytes:
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
    model = res['model']
    fc    = res['feature_cols']
    buf   = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        pd.DataFrame({
            'Kelas': ['Tidak Penerima (0)', 'Penerima (1)'],
            'Prior Probability': model.class_prior_,
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
            "📋  Log Perhitungan",
            "💾  Export Hasil",
            "ℹ️  Tentang Sistem",
        ],
        label_visibility="collapsed"
    )

    st.markdown('<div class="sidebar-divider"></div>', unsafe_allow_html=True)

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
        "Machine Learning · Naive Bayes"
    )

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
        ("6", "Log Perhitungan",  "Lihat detail perhitungan Naive Bayes step-by-step"),
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
        help="Format: .xlsx | Sheet: Table 1"
    )

    col1, col2 = st.columns([2, 1])

    with col1:
        if uploaded is not None:
            with st.spinner("Membaca file Excel..."):
                try:
                    df_raw = pd.read_excel(uploaded, sheet_name=0, header=1, skiprows=[1])
                    if len(df_raw.columns) < 14:
                        uploaded.seek(0)
                        df_raw = pd.read_excel(uploaded, sheet_name=0)

                    df_raw = df_raw.dropna(how='all')
                    st.session_state['df_raw'] = df_raw

                    alert(f"File berhasil dimuat — {len(df_raw)} baris, {len(df_raw.columns)} kolom")

                    render_section("🔍", "Preview Dataset (10 Baris Pertama)")
                    st.dataframe(df_raw.head(10), use_container_width=True, height=300)

                    render_section("📊", "Statistik Awal")
                    c1, c2, c3, c4 = st.columns(4)
                    with c1: st.metric("Total Baris",  len(df_raw))
                    with c2: st.metric("Total Kolom",  len(df_raw.columns))
                    with c3: st.metric("Missing Values", int(df_raw.isnull().sum().sum()))
                    with c4: st.metric("Duplikat",     int(df_raw.duplicated().sum()))

                    st.markdown("<br>", unsafe_allow_html=True)
                    if st.button("▶ Lanjut ke Preprocessing", use_container_width=True):
                        with st.spinner("Menjalankan preprocessing..."):
                            progress = st.progress(0)
                            for i in range(1, 101):
                                time.sleep(0.005)
                                progress.progress(i)
                            df_p, info = run_preprocessing(df_raw)
                            st.session_state['df_processed'] = df_p
                            st.session_state['preprocess_info'] = info
                        st.success("✅ Preprocessing selesai! Buka menu Preprocessing atau Log Perhitungan untuk detail.")

                except Exception as e:
                    st.error(f"❌ Error membaca file: {e}")

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
# PAGE: PREPROCESSING (singkat)
# ─────────────────────────────────────────────────────────────

elif page == "⚙️  Preprocessing":
    render_hero(
        "Preprocessing Data",
        "Transformasi, encoding, dan standardisasi data sebelum pelatihan model",
        "Data Preprocessing"
    )

    if "df_processed" not in st.session_state:
        st.markdown('<div class="alert-warning">⚠️ Dataset belum dimuat.</div>', unsafe_allow_html=True)
    else:
        df   = st.session_state['df_processed']
        info = st.session_state['preprocess_info']
        le   = info['le_pekerjaan']

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
                <div>
                    <div class="step-title">{s}</div>
                    <div class="step-desc">{d}</div>
                </div>
            </div>
            """, unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        st.info("📋 Buka menu **Log Perhitungan** untuk melihat detail preprocessing di UI.")


# ─────────────────────────────────────────────────────────────
# PAGE: VISUALISASI (ringkas)
# ─────────────────────────────────────────────────────────────

elif page == "📊  Visualisasi":
    render_hero("Visualisasi Data", "Eksplorasi distribusi data dengan grafik interaktif Plotly", "Data Visualization")

    if "df_processed" not in st.session_state:
        st.markdown('<div class="alert-warning">⚠️ Dataset belum dimuat.</div>', unsafe_allow_html=True)
    else:
        df = st.session_state['df_processed']
        th = make_plotly_theme()

        tabs = st.tabs(["🥧 Label", "💰 Pendapatan", "💼 Pekerjaan", "🏠 Status Rumah", "🎓 Kelas", "📊 Korelasi"])

        with tabs[0]:
            label_counts = df['LABEL'].value_counts()
            fig = px.pie(names=label_counts.index, values=label_counts.values,
                         color_discrete_sequence=["#059669", "#f87171"], hole=0.45)
            fig.update_layout(**th, height=400)
            st.plotly_chart(fig, use_container_width=True)

        with tabs[1]:
            fig = px.histogram(df, x='PENDAPATAN_ZSCORE', nbins=30, color='LABEL',
                              color_discrete_map={'Ya': '#059669', 'Tidak': '#f87171'}, barmode='overlay', opacity=0.7)
            fig.update_layout(**th, height=400)
            st.plotly_chart(fig, use_container_width=True)

        with tabs[2]:
            pek_cnt = df['PEKERJAAN ORANG TUA'].value_counts().reset_index()
            pek_cnt.columns = ['Pekerjaan', 'Jumlah']
            fig = px.bar(pek_cnt, y='Pekerjaan', x='Jumlah', orientation='h',
                        color='Jumlah', color_continuous_scale='Greens')
            fig.update_layout(**th, height=max(350, len(pek_cnt)*38))
            st.plotly_chart(fig, use_container_width=True)

        with tabs[3]:
            sr_cnt = df['STATUS RUMAH'].value_counts()
            fig = px.pie(names=sr_cnt.index, values=sr_cnt.values,
                        color_discrete_sequence=["#059669", "#34d399"], hole=0.4)
            fig.update_layout(**th, height=350)
            st.plotly_chart(fig, use_container_width=True)

        with tabs[4]:
            kl_cnt = df.groupby(['KELAS', 'LABEL']).size().reset_index(name='Jumlah')
            fig = px.bar(kl_cnt, x='KELAS', y='Jumlah', color='LABEL', barmode='group',
                        color_discrete_map={'Ya': '#059669', 'Tidak': '#f87171'})
            fig.update_layout(**th, height=400)
            st.plotly_chart(fig, use_container_width=True)

        with tabs[5]:
            num_cols = ['KELAS_NUM', 'PENDAPATAN_ZSCORE', 'PEKERJAAN ORANG TUA_ENC',
                       'JUMLAH TANGGUNGAN', 'STATUS RUMAH_ENC', 'LABEL_ENC']
            num_avail = [c for c in num_cols if c in df.columns]
            corr = df[num_avail].corr().round(3)
            fig = px.imshow(corr, text_auto=True, color_continuous_scale='Greens', aspect='auto')
            fig.update_layout(**th, height=500)
            st.plotly_chart(fig, use_container_width=True)


# ─────────────────────────────────────────────────────────────
# PAGE: TRAINING MODEL
# ─────────────────────────────────────────────────────────────

elif page == "🤖  Training Model":
    render_hero("Training Model", "Pelatihan Gaussian Naive Bayes dengan pembagian data train/test", "Model Training")

    if "df_processed" not in st.session_state:
        st.markdown('<div class="alert-warning">⚠️ Dataset belum dimuat.</div>', unsafe_allow_html=True)
    else:
        df        = st.session_state['df_processed']
        test_size = st.session_state.get('test_size', 0.2)

        n_valid = df['LABEL_ENC'].notna().sum()
        n_train = int(n_valid * (1 - test_size))
        n_test  = n_valid - n_train

        render_section("⚙️", "Konfigurasi Model")
        render_metric_cards([
            {"icon": "📦", "label": "Total Data Valid",  "value": str(n_valid), "desc": "setelah drop NaN"},
            {"icon": "🏋️", "label": "Data Training",    "value": str(n_train), "desc": f"{(1-test_size)*100:.0f}%"},
            {"icon": "🧪", "label": "Data Testing",      "value": str(n_test),  "desc": f"{test_size*100:.0f}%"},
            {"icon": "🔬", "label": "Algoritma",         "value": "GNB",        "desc": "Gaussian Naive Bayes"},
        ])

        st.markdown("<br>", unsafe_allow_html=True)

        if st.button("🚀 Mulai Training Model", use_container_width=True):
            with st.spinner("Melatih model..."):
                progress = st.progress(0)
                time.sleep(0.1); progress.progress(20)
                res = run_training(df, test_size=test_size)
                time.sleep(0.1); progress.progress(60)
                st.session_state['train_results'] = res
                time.sleep(0.1); progress.progress(100)

            st.success("✅ Model berhasil dilatih!")
            st.info("📋 Buka menu **Log Perhitungan** untuk melihat detail perhitungan di UI.")

            st.markdown("<br>", unsafe_allow_html=True)
            render_metric_cards([
                {"icon": "🏋️", "label": "Akurasi Training", "value": f"{res['train_acc']*100:.2f}%", "desc": ""},
                {"icon": "🧪", "label": "Akurasi Testing",  "value": f"{res['test_acc']*100:.2f}%",  "desc": ""},
                {"icon": "✓",  "label": "Prediksi Benar",   "value": str((res['y_test'].values == res['y_pred']).sum()), "desc": ""},
                {"icon": "✗",  "label": "Prediksi Salah",   "value": str((res['y_test'].values != res['y_pred']).sum()), "desc": ""},
            ])

        elif "train_results" in st.session_state:
            res = st.session_state['train_results']
            alert("Model sudah terlatih. Klik tombol di atas untuk melatih ulang.")


# ─────────────────────────────────────────────────────────────
# PAGE: EVALUASI (ringkas)
# ─────────────────────────────────────────────────────────────

elif page == "📈  Evaluasi":
    render_hero("Evaluasi Model", "Akurasi, Precision, Recall, F1-Score, dan Confusion Matrix", "Model Evaluation")

    if "train_results" not in st.session_state:
        st.markdown('<div class="alert-warning">⚠️ Model belum dilatih.</div>', unsafe_allow_html=True)
    else:
        res = st.session_state['train_results']
        th  = make_plotly_theme()

        render_section("🎯", "Metrik Utama")
        render_metric_cards([
            {"icon": "🏋️", "label": "Akurasi Training", "value": f"{res['train_acc']*100:.2f}%",  "desc": ""},
            {"icon": "🧪", "label": "Akurasi Testing",  "value": f"{res['test_acc']*100:.2f}%",   "desc": ""},
            {"icon": "📊", "label": "F1-Score Macro",   "value": f"{res['macro_f1']*100:.2f}%",   "desc": ""},
            {"icon": "⚖️", "label": "F1-Score Weighted","value": f"{res['w_f1']*100:.2f}%",       "desc": ""},
        ])

        st.markdown("<br>", unsafe_allow_html=True)

        # Confusion Matrix
        cm = res['cm']
        fig = go.Figure(data=go.Heatmap(
            z=cm,
            x=['Predicted: Tidak', 'Predicted: Ya'],
            y=['Actual: Tidak', 'Actual: Ya'],
            text=[[str(v) for v in row] for row in cm],
            texttemplate="%{text}",
            textfont={"size": 22, "color": "white"},
            colorscale=[[0, "#ecfdf5"], [0.4, "#34d399"], [1, "#065f46"]]
        ))
        fig.update_layout(**th, title="Confusion Matrix", height=400)
        st.plotly_chart(fig, use_container_width=True)


# ─────────────────────────────────────────────────────────────
# PAGE: HASIL PREDIKSI (ringkas)
# ─────────────────────────────────────────────────────────────

elif page == "🗂️  Hasil Prediksi":
    render_hero("Hasil Prediksi", "Tabel prediksi dengan probabilitas dan status", "Prediction Results")

    if "train_results" not in st.session_state:
        st.markdown('<div class="alert-warning">⚠️ Model belum dilatih.</div>', unsafe_allow_html=True)
    else:
        res = st.session_state['train_results']
        df_r = res['results_df']

        benar = (df_r['Status'] == '✓ Benar').sum()
        render_metric_cards([
            {"icon": "📋", "label": "Total Data Uji",  "value": str(len(df_r)), "desc": ""},
            {"icon": "✓",  "label": "Prediksi Benar", "value": str(benar),     "desc": f"{benar/len(df_r)*100:.1f}%"},
            {"icon": "🎯", "label": "Akurasi",         "value": f"{benar/len(df_r)*100:.2f}%", "desc": ""},
        ])

        st.dataframe(df_r, use_container_width=True, height=450, hide_index=True)


# ─────────────────────────────────────────────────────────────
# PAGE: LOG PERHITUNGAN (FITUR BARU UNTUK STREAMLIT CLOUD)
# ─────────────────────────────────────────────────────────────

elif page == "📋  Log Perhitungan":
    render_hero(
        "Log Perhitungan Naive Bayes",
        "Detail lengkap setiap langkah perhitungan — preprocessing hingga evaluasi",
        "Calculation Log"
    )

    tabs = st.tabs(["📋 Preprocessing Log", "🤖 Training & Prediksi Log"])

    with tabs[0]:
        render_section("📋", "Log Preprocessing")
        if "preprocess_log" in st.session_state:
            log_text = st.session_state['preprocess_log']
            st.markdown(f'<div class="terminal-output">{log_text}</div>', unsafe_allow_html=True)
            
            st.download_button(
                "⬇️ Download Log Preprocessing (.txt)",
                data=log_text,
                file_name="log_preprocessing.txt",
                mime="text/plain",
                use_container_width=True
            )
        else:
            st.info("📂 Upload dan preprocessing dataset terlebih dahulu untuk melihat log.")

    with tabs[1]:
        render_section("🤖", "Log Training & Prediksi")
        if "training_log" in st.session_state:
            log_text = st.session_state['training_log']
            
            # Tampilkan dalam expander karena panjang
            with st.expander("🔍 Klik untuk melihat/menyembunyikan log training", expanded=True):
                st.markdown(f'<div class="terminal-output">{log_text}</div>', unsafe_allow_html=True)
            
            st.download_button(
                "⬇️ Download Log Training (.txt)",
                data=log_text,
                file_name="log_training_naive_bayes.txt",
                mime="text/plain",
                use_container_width=True
            )
            
            st.info("💡 Log ini menampilkan maksimal 10 data testing pertama. Untuk melihat semua data, gunakan `streamlit run app.py` di terminal lokal.")
        else:
            st.info("🤖 Training model terlebih dahulu untuk melihat log training.")


# ─────────────────────────────────────────────────────────────
# PAGE: EXPORT HASIL
# ─────────────────────────────────────────────────────────────

elif page == "💾  Export Hasil":
    render_hero("Export Hasil", "Unduh seluruh hasil ke file Excel terformat", "Export & Download")

    if "df_processed" not in st.session_state:
        st.markdown('<div class="alert-warning">⚠️ Dataset belum dimuat.</div>', unsafe_allow_html=True)
    else:
        df   = st.session_state['df_processed']
        info = st.session_state['preprocess_info']
        le   = info['le_pekerjaan']

        c1, c2 = st.columns(2)
        with c1:
            st.markdown('<div class="card"><div class="card-title">📊 Data Preprocessing</div></div>', unsafe_allow_html=True)
            preproc_bytes = export_preprocessing_excel(df, le)
            st.download_button("⬇️ Download (.xlsx)", data=preproc_bytes,
                              file_name="data_bsm_preprocessed.xlsx", use_container_width=True)

        with c2:
            if "train_results" in st.session_state:
                res = st.session_state['train_results']
                st.markdown('<div class="card"><div class="card-title">🎯 Hasil Prediksi</div></div>', unsafe_allow_html=True)
                st.download_button("⬇️ Download (.xlsx)", data=export_results_excel(res),
                                  file_name="hasil_prediksi.xlsx", use_container_width=True)

        st.markdown("<br>", unsafe_allow_html=True)
        c3, c4 = st.columns(2)
        with c3:
            if "train_results" in st.session_state:
                res = st.session_state['train_results']
                st.download_button("⬇️ Evaluasi (.xlsx)", data=export_evaluation_excel(res),
                                  file_name="evaluasi.xlsx", use_container_width=True)
        with c4:
            if "train_results" in st.session_state:
                res = st.session_state['train_results']
                st.download_button("⬇️ Parameter (.xlsx)", data=export_params_excel(res),
                                  file_name="parameter_nb.xlsx", use_container_width=True)


# ─────────────────────────────────────────────────────────────
# PAGE: TENTANG SISTEM
# ─────────────────────────────────────────────────────────────

elif page == "ℹ️  Tentang Sistem":
    render_hero("Tentang Sistem", "Informasi lengkap mengenai sistem klasifikasi penerima BSM", "About")

    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("""
        <div class="card">
            <div style="font-size:1.8rem;">🎓</div>
            <div style="font-weight:700;color:#1e293b;margin:8px 0;">Tujuan Sistem</div>
            <p style="font-size:0.85rem;color:#64748b;">Membantu keputusan pemberian BSM menggunakan Machine Learning Gaussian Naive Bayes.</p>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown("""
        <div class="card">
            <div style="font-size:1.8rem;">🔬</div>
            <div style="font-weight:700;color:#1e293b;margin:8px 0;">Algoritma</div>
            <p style="font-size:0.85rem;color:#64748b;">Gaussian Naive Bayes — algoritma probabilistik berbasis Teorema Bayes.</p>
        </div>
        """, unsafe_allow_html=True)
    with col3:
        st.markdown("""
        <div class="card">
            <div style="font-size:1.8rem;">📊</div>
            <div style="font-weight:700;color:#1e293b;margin:8px 0;">Fitur</div>
            <p style="font-size:0.85rem;color:#64748b;">Upload → Preprocessing → Training → Evaluasi → Export. Semua dalam 1 dashboard.</p>
        </div>
        """, unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────
# FOOTER
# ─────────────────────────────────────────────────────────────

st.markdown("""
<div class="footer">
    <div class="footer-brand">🎓 Sistem Klasifikasi Penerima BSM</div>
    <div class="footer-divider"></div>
    <div>Gaussian Naive Bayes · Machine Learning Dashboard · Streamlit</div>
    <div style="margin-top:0.4rem;color:#64748b;font-size:0.75rem;">Made with ❤️</div>
</div>
""", unsafe_allow_html=True)
