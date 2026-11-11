# -*- coding: utf-8 -*-
"""
Kurumsal Satış Analiz Platformu
=============================================
Aylık/dönemsel satış Excel verisini yükleyip çok boyutlu interaktif
analiz sunan, beyaz etiketlenebilir (white-label) Streamlit uygulaması.
Farklı şirketlere marka adı/logo/renk değiştirilerek yeniden satılabilir.

Çalıştırmak için:
    pip install streamlit pandas plotly openpyxl xlsxwriter
    streamlit run reporting_analysis_app.py
"""

import io
import re
import unicodedata

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import plotly.io as pio
import streamlit as st

# --------------------------------------------------------------------------------------
# SAYFA AYARLARI
# --------------------------------------------------------------------------------------
st.set_page_config(
    page_title="Satış Analiz Platformu",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# --------------------------------------------------------------------------------------
# BEYAZ ETİKETLEME (WHITE-LABEL) — MARKA AYARLARI
# --------------------------------------------------------------------------------------
# Bu blok, uygulamanın tek bir kod tabanıyla farklı müşterilere (şirketlere)
# yeniden satılabilmesini sağlar: şirket adı, logo ve marka rengi tek yerden
# değiştirilir, kod içinde hiçbir yerde sabit (hardcoded) marka referansı yoktur.
if "brand_name" not in st.session_state:
    st.session_state.brand_name = "Şirketiniz"
if "brand_accent" not in st.session_state:
    st.session_state.brand_accent = "#116BBE"  # Brass Gold — varsayılan vurgu rengi
if "brand_logo" not in st.session_state:
    st.session_state.brand_logo = None

with st.sidebar.expander("⚙️ Marka Ayarları (White-Label)", expanded=False):
    st.session_state.brand_name = st.text_input("Şirket / Marka Adı", value=st.session_state.brand_name)
    st.session_state.brand_accent = st.color_picker("Vurgu Rengi", value=st.session_state.brand_accent)
    logo_file = st.file_uploader("Logo (opsiyonel, PNG/JPG)", type=["png", "jpg", "jpeg"], key="logo_uploader")
    if logo_file is not None:
        st.session_state.brand_logo = logo_file.getvalue()
    st.caption("Bu ayarlar sadece görünümü değiştirir; veri işleme mantığı etkilenmez. Farklı bir müşteriye teslim ederken bu üç alanı güncellemeniz yeterlidir.")

BRAND_NAME = st.session_state.brand_name
BRAND_ACCENT = st.session_state.brand_accent
BRAND_LOGO = st.session_state.brand_logo

# --- Kurumsal renk paleti (marka rengi hariç sabit, nötr zemin tonları) ---
NAVY_INK = "#0E1A2B"
CLOUD_GRAY = "#F5F6F8"
SLATE = "#4A5568"
HARBOR_TEAL = "#2F6F8F"
SIGNAL_GREEN = "#2E7D5B"
SIGNAL_RED = "#C0403A"
BORDER_GRAY = "#E5E9EF"

# --------------------------------------------------------------------------------------
# KURUMSAL CSS ENJEKSİYONU
# --------------------------------------------------------------------------------------
st.markdown(
    f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600;700&family=IBM+Plex+Mono:wght@500;600;700&display=swap');

    html, body, [class*="css"] {{
        font-family: 'IBM Plex Sans', sans-serif;
    }}

    /* --- Ana zemin --- */
    .stApp {{
        background-color: {CLOUD_GRAY};
    }}

    /* --- Sidebar: koyu 'kontrol paneli' hissi --- */
    [data-testid="stSidebar"] {{
        background-color: {NAVY_INK};
    }}
    [data-testid="stSidebar"] * {{
        color: #E7ECF3 !important;
    }}
    [data-testid="stSidebar"] .stSelectbox label,
    [data-testid="stSidebar"] .stMultiSelect label,
    [data-testid="stSidebar"] .stRadio label,
    [data-testid="stSidebar"] .stTextInput label {{
        font-size: 0.78rem;
        text-transform: uppercase;
        letter-spacing: 0.06em;
        color: #93A2B8 !important;
        font-weight: 600;
    }}
    [data-testid="stSidebar"] hr {{
        border-color: rgba(255,255,255,0.12);
    }}
    [data-testid="stSidebar"] [data-baseweb="tag"] {{
        background-color: {BRAND_ACCENT} !important;
    }}

    /* --- Sekmeler (tabs): pill navigasyon --- */
    [data-testid="stTabs"] button[role="tab"] {{
        font-family: 'IBM Plex Sans', sans-serif;
        font-weight: 600;
        color: {SLATE};
        border-radius: 8px 8px 0 0;
    }}
    [data-testid="stTabs"] button[aria-selected="true"] {{
        color: {NAVY_INK} !important;
        border-bottom: 3px solid {BRAND_ACCENT} !important;
    }}

    /* --- KPI kartları (st.metric otomatik olarak 'gösterge paneli' stiline döner) --- */
    [data-testid="stMetric"] {{
        background-color: #FFFFFF;
        border: 1px solid {BORDER_GRAY};
        border-top: 4px solid {BRAND_ACCENT};
        border-radius: 10px;
        padding: 1rem 1.1rem 0.9rem 1.1rem;
        box-shadow: 0 1px 3px rgba(14,26,43,0.06);
    }}
    [data-testid="stMetricLabel"] {{
        font-size: 0.78rem;
        text-transform: uppercase;
        letter-spacing: 0.06em;
        color: {SLATE} !important;
        font-weight: 600;
    }}
    [data-testid="stMetricValue"] {{
        font-family: 'IBM Plex Mono', monospace;
        color: {NAVY_INK} !important;
        font-weight: 700;
    }}

    /* --- Butonlar --- */
    .stButton button, .stDownloadButton button {{
        background-color: {NAVY_INK};
        color: #FFFFFF;
        border-radius: 8px;
        border: none;
        font-weight: 600;
    }}
    .stButton button:hover, .stDownloadButton button:hover {{
        background-color: {BRAND_ACCENT};
        color: {NAVY_INK};
    }}

    /* --- Dosya yükleme alanı --- */
    [data-testid="stFileUploaderDropzone"] {{
        border: 2px dashed {HARBOR_TEAL};
        background-color: #FFFFFF;
        border-radius: 10px;
    }}

    /* --- Başlıklar --- */
    h1, h2, h3 {{
        color: {NAVY_INK};
        font-family: 'IBM Plex Sans', sans-serif;
    }}

    /* --- Marka header banner --- */
    .brand-header {{
        display: flex;
        align-items: center;
        gap: 0.9rem;
        padding: 0.9rem 1.2rem;
        background-color: {NAVY_INK};
        border-radius: 12px;
        margin-bottom: 1.2rem;
    }}
    .brand-header .brand-title {{
        color: #FFFFFF;
        font-size: 1.35rem;
        font-weight: 700;
        letter-spacing: 0.01em;
        margin: 0;
    }}
    .brand-header .brand-subtitle {{
        color: #93A2B8;
        font-size: 0.85rem;
        margin: 0;
    }}
    .brand-header .brand-bar {{
        width: 4px;
        align-self: stretch;
        background-color: {BRAND_ACCENT};
        border-radius: 4px;
    }}
    </style>
    """,
    unsafe_allow_html=True,
)

# --- Plotly için kurumsal grafik teması (tüm px./go. grafiklerine otomatik uygulanır) ---
_corporate_template = go.layout.Template(
    layout=go.Layout(
        font=dict(family="IBM Plex Sans, sans-serif", color=NAVY_INK, size=13),
        title=dict(font=dict(family="IBM Plex Sans, sans-serif", size=16, color=NAVY_INK)),
        colorway=[BRAND_ACCENT, HARBOR_TEAL, SLATE, SIGNAL_GREEN, SIGNAL_RED, "#8899A6", "#B7C3D0"],
        paper_bgcolor="#FFFFFF",
        plot_bgcolor="#FFFFFF",
        xaxis=dict(gridcolor=BORDER_GRAY, zerolinecolor=BORDER_GRAY, linecolor=BORDER_GRAY),
        yaxis=dict(gridcolor=BORDER_GRAY, zerolinecolor=BORDER_GRAY, linecolor=BORDER_GRAY),
        legend=dict(bgcolor="rgba(0,0,0,0)"),
    )
)
pio.templates["corporate"] = _corporate_template
pio.templates.default = "corporate"

# --- Yardımcı fonksiyonlar: grafik indirme, bellek optimizasyonu, anomali tespiti ---
def save_plotly_buttons(fig, name: str = "chart"):
    """Plotly figürünü PNG/SVG/PDF olarak indirmeye hazır hale getirir. Kaleido gerekebilir."""
    formats = [("png", "image/png"), ("svg", "image/svg+xml"), ("pdf", "application/pdf")]
    for fmt, mime in formats:
        try:
            img = pio.to_image(fig, format=fmt, scale=2)
            st.download_button(label=f"📥 {name}.{fmt}", data=img, file_name=f"{name}.{fmt}", mime=mime)
        except Exception:
            # Kaleido yoksa uyar ve sonraki formatları deneme
            if fmt == "png":
                st.info("Grafik indirmek için 'kaleido' paketini yükleyin: pip install kaleido")
            break


def reduce_memory(df: pd.DataFrame) -> pd.DataFrame:
    """Basit downcast ve kategorileştirme ile bellek kullanımını azaltır."""
    for col in df.columns:
        try:
            if pd.api.types.is_integer_dtype(df[col]):
                df[col] = pd.to_numeric(df[col], downcast="integer")
            elif pd.api.types.is_float_dtype(df[col]):
                df[col] = pd.to_numeric(df[col], downcast="float")
            elif pd.api.types.is_object_dtype(df[col]):
                nuniq = df[col].nunique(dropna=True)
                if len(df) > 0 and (nuniq / len(df)) < 0.5:
                    df[col] = df[col].astype("category")
        except Exception:
            pass
    return df


def detect_anomalies(df: pd.DataFrame, dims: list, period: str, threshold_pct: float, primary_sales_col: str, noise_floor: float = 0):
    """Basit anomali tespiti: son dönem ile önceki dönem arasındaki % değişimi kontrol eder.
    period: 'YOY' veya 'MoM'
    Dönen DataFrame sütunları: Dim, Name, Period_Old, Period_New, Old, New, Pct
    """
    anomalies = []
    if not dims or primary_sales_col not in df.columns:
        return pd.DataFrame()

    if period.upper() == "YOY":
        years_present = sorted(df["Year"].dropna().unique().tolist())
        if len(years_present) < 2:
            return pd.DataFrame()
        y_new, y_old = years_present[-1], years_present[-2]
        for dim in dims:
            if dim not in df.columns:
                continue
            g_new = df[df["Year"] == y_new].groupby(dim)[primary_sales_col].sum()
            g_old = df[df["Year"] == y_old].groupby(dim)[primary_sales_col].sum()
            joined = pd.concat([g_new.rename("new"), g_old.rename("old")], axis=1).fillna(0)
            joined = joined[(joined["new"] + joined["old"]) >= noise_floor]
            for name, row in joined.iterrows():
                old = row.get("old", 0)
                new = row.get("new", 0)
                if old == 0:
                    continue
                pct = (new - old) / abs(old) * 100
                if abs(pct) >= threshold_pct:
                    anomalies.append((dim, name, y_old, y_new, old, new, pct))

    elif period.upper() in ("MOM", "MO", "M"):
        if "Month" not in df.columns:
            return pd.DataFrame()
        mdf = df.copy()
        mdf["YearMonth"] = mdf["Year"].astype(str) + "-" + mdf["Month"].astype(str)
        ym_sorted = sorted(mdf["YearMonth"].dropna().unique().tolist())
        if len(ym_sorted) < 2:
            return pd.DataFrame()
        ym_new, ym_old = ym_sorted[-1], ym_sorted[-2]
        for dim in dims:
            if dim not in mdf.columns:
                continue
            g_new = mdf[mdf["YearMonth"] == ym_new].groupby(dim)[primary_sales_col].sum()
            g_old = mdf[mdf["YearMonth"] == ym_old].groupby(dim)[primary_sales_col].sum()
            joined = pd.concat([g_new.rename("new"), g_old.rename("old")], axis=1).fillna(0)
            joined = joined[(joined["new"] + joined["old"]) >= noise_floor]
            for name, row in joined.iterrows():
                old = row.get("old", 0)
                new = row.get("new", 0)
                if old == 0:
                    continue
                pct = (new - old) / abs(old) * 100
                if abs(pct) >= threshold_pct:
                    anomalies.append((dim, name, ym_old, ym_new, old, new, pct))

    if not anomalies:
        return pd.DataFrame()
    return pd.DataFrame(anomalies, columns=["Dim", "Name", "Period_Old", "Period_New", "Old", "New", "Pct"])


# --- Marka header banner'ı çiz ---
_header_cols = st.columns([0.06, 0.55, 2.5]) if BRAND_LOGO else None
if BRAND_LOGO:
    with st.container():
        hc1, hc2 = st.columns([1, 8])
        with hc1:
            st.image(BRAND_LOGO, width=56)
        with hc2:
            st.markdown(
                f"""<div class="brand-header" style="background:none;padding:0.2rem 0;">
                <div class="brand-bar"></div>
                <div>
                    <p class="brand-title" style="color:{NAVY_INK};">{BRAND_NAME}</p>
                    <p class="brand-subtitle" style="color:{SLATE};">Kurumsal Satış Analiz Platformu</p>
                </div>
                </div>""",
                unsafe_allow_html=True,
            )
else:
    st.markdown(
        f"""
        <div class="brand-header">
            <div class="brand-bar"></div>
            <div>
                <p class="brand-title">{BRAND_NAME}</p>
                <p class="brand-subtitle">Kurumsal Satış Analiz Platformu</p>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

# --------------------------------------------------------------------------------------
# BEKLENEN SÜTUNLAR VE ESNEK EŞLEŞTİRME
# --------------------------------------------------------------------------------------
# Ham veride sütun adları bazen fazladan boşluk / farklı yazım içerebilir
# (örn: " Quantity ", "Net Sales EUR "). Bu sözlük, normalize edilmiş
# (boşluksuz, küçük harf) adları standart isimlere eşler.
COLUMN_ALIASES = {
    "year": "Year",
    "scenario": "Scenario",
    "month": "Month",
    "region": "Region",
    "countrygroup": "Country Group",
    "country": "Country",
    "ph1code": "PH1 Code",
    "ph1": "PH1",
    "ph2code": "PH2 Code",
    "ph2": "PH2",
    "pgroup": "Pgroup",
    "ph3code": "PH3 Code",
    "ph3": "PH3",
    "ph4code": "PH4 Code",
    "ph4": "PH4",
    "brand": "Brand",
    "bifs": "BI/FS",
    "bi/fs": "BI/FS",
    "source": "Source",
    "quantity": "Quantity",
    "netsaleseur": "Net Sales EUR",
    "netsalesusd": "Net Sales USD",
}

BIFS_LABELS = {"BI": "Built-In", "FS": "Free-Standing"}

MONTH_NAMES = {
    "01": "Oca", "02": "Şub", "03": "Mar", "04": "Nis", "05": "May", "06": "Haz",
    "07": "Tem", "08": "Ağu", "09": "Eyl", "10": "Eki", "11": "Kas", "12": "Ara",
}
MONTH_ORDER = [f"{i:02d}" for i in range(1, 13)]

REQUIRED_MIN_COLS = {"Year", "Quantity", "Net Sales EUR", "Net Sales USD"}


def normalize_key(col: str) -> str:
    """Sütun adını karşılaştırma için normalize eder: boşluk/alt çizgi kaldır, küçük harfe çevir."""
    col = unicodedata.normalize("NFKD", str(col))
    col = col.strip().lower()
    col = re.sub(r"[\s_]+", "", col)
    return col


@st.cache_data(show_spinner=False)
def load_and_clean(file_bytes: bytes, filename: str) -> pd.DataFrame:
    """Excel dosyasını okur, sütunları standartlaştırır ve tipleri temizler."""
    if filename.lower().endswith(".csv"):
        df = pd.read_csv(io.BytesIO(file_bytes))
    else:
        df = pd.read_excel(io.BytesIO(file_bytes))

    # Sütun adlarını eşle
    rename_map = {}
    for col in df.columns:
        key = normalize_key(col)
        if key in COLUMN_ALIASES:
            rename_map[col] = COLUMN_ALIASES[key]
    df = df.rename(columns=rename_map)

    # Metin sütunlarındaki baştaki/sondaki boşlukları temizle.
    # Bu adım kritik: "Azerbaijan" ile "Azerbaijan " (sonunda boşluklu) Excel'de
    # farklı iki grup gibi okunup filtrelerde/gruplamalarda veri kaybına yol açabilir.
    text_cols = [
        "Region", "Country Group", "Country", "PH1 Code", "PH1", "PH2 Code", "PH2",
        "Pgroup", "PH3 Code", "PH3", "PH4 Code", "PH4", "Brand", "BI/FS", "Source", "Scenario",
    ]
    for col in text_cols:
        if col in df.columns and not pd.api.types.is_numeric_dtype(df[col]):
            df[col] = df[col].astype(str).str.strip()
            df[col] = df[col].replace({"nan": pd.NA, "None": pd.NA})

    # Sayısal sütunları temizle (binlik ayraç, virgül, boşluk olabilir)
    for num_col in ["Quantity", "Net Sales EUR", "Net Sales USD"]:
        if num_col in df.columns:
            if not pd.api.types.is_numeric_dtype(df[num_col]):
                df[num_col] = (
                    df[num_col]
                    .astype(str)
                    .str.replace(",", "", regex=False)
                    .str.replace(" ", "", regex=False)
                    .replace({"": "0", "nan": "0"})
                )
            df[num_col] = pd.to_numeric(df[num_col], errors="coerce").fillna(0)

    # Year'ı stringe çevir (2024 vs "2024" karışıklığını önlemek için)
    if "Year" in df.columns:
        df["Year"] = df["Year"].astype(str).str.extract(r"(\d{4})")[0]

    # Month'u iki haneli string yap
    if "Month" in df.columns:
        df["Month"] = df["Month"].astype(str).str.zfill(2)

    # BI/FS okunabilir hale getir
    if "BI/FS" in df.columns:
        df["BI/FS_Label"] = df["BI/FS"].map(BIFS_LABELS).fillna(df["BI/FS"])

    # Birim fiyat hesapla
    if "Quantity" in df.columns and "Net Sales EUR" in df.columns:
        df["Unit Price EUR"] = df.apply(
            lambda r: (r["Net Sales EUR"] / r["Quantity"]) if r["Quantity"] else 0, axis=1
        )
    if "Quantity" in df.columns and "Net Sales USD" in df.columns:
        df["Unit Price USD"] = df.apply(
            lambda r: (r["Net Sales USD"] / r["Quantity"]) if r["Quantity"] else 0, axis=1
        )

    return df


def fmt_num(x: float) -> str:
    return f"{x:,.0f}".replace(",", ".")


def fmt_money(x: float) -> str:
    return f"{x:,.0f}".replace(",", ".")


def pct_change(new: float, old: float):
    if old == 0 or pd.isna(old):
        return None
    return (new - old) / abs(old) * 100


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Data")
    return output.getvalue()


# --------------------------------------------------------------------------------------
# BAŞLIK VE DOSYA YÜKLEME
# --------------------------------------------------------------------------------------
st.caption("Excel raporunuzu yükleyin, sütunlar otomatik tanınsın, çok boyutlu analiz görüntülensin.")

uploaded_file = st.file_uploader(
    "Excel / CSV dosyanızı yükleyin",
    type=["xlsx", "xls", "csv"],
    help="Beklenen sütunlar: Year, Scenario, Month, Region, Country, PH1-PH4, Brand, BI/FS, Source, Quantity, Net Sales EUR, Net Sales USD",
)

if uploaded_file is None:
    st.info("Devam etmek için bir dosya yükleyin.")
    st.stop()

with st.spinner("Veri okunuyor ve temizleniyor..."):
    df = load_and_clean(uploaded_file.getvalue(), uploaded_file.name)

# --- Gelişmiş seçenekler (sidebar) ---
with st.sidebar.expander("⚙️ Gelişmiş İşlevler", expanded=False):
    memory_opt = st.checkbox("Bellek optimizasyonu (downcast/categorize)", value=False,
                             help="Büyük dosyalarda bellek kullanımını azaltmak için sütunları downcast eder ve uygun olanları category yapar.")
    anomaly_enable = st.checkbox("Anomali Tespiti Etkinleştir", value=True,
                                help="Seçilen boyutlarda YOY/MoM bazında büyük sapmaları otomatik tespit eder.")
    anomaly_period = st.selectbox("Anomali Periyodu", ["YOY", "MoM"], index=0)
    anomaly_threshold = st.slider("Anomali Eşiği (%)", 5, 100, 15, step=1)

    st.markdown("---")
    st.caption("Grafik indirme: PNG/SVG/PDF (kaleido gerekebilir).")

    # Bilgilendirme: required package
    st.write("Not: Grafik indirme için `kaleido` kurulu olmalı: `pip install kaleido`")

if memory_opt:
    before = df.memory_usage(deep=True).sum()
    df = reduce_memory(df)
    after = df.memory_usage(deep=True).sum()
    st.info(f"Bellek optimize edildi: {before:,} → {after:,} bytes".replace(",", "."))

missing = REQUIRED_MIN_COLS - set(df.columns)
if missing:
    st.error(
        f"Aşağıdaki gerekli sütunlar bulunamadı: {', '.join(missing)}. "
        f"Lütfen sütun başlıklarını kontrol edin."
    )
    with st.expander("Bulunan sütunlar"):
        st.write(list(df.columns))
    st.stop()

st.success(f"Veri yüklendi: {len(df):,} satır, {len(df.columns)} sütun.".replace(",", "."))

# --------------------------------------------------------------------------------------
# SIDEBAR FİLTRELER
# --------------------------------------------------------------------------------------
st.sidebar.header("🔎 Filtreler")

# --- Döviz dönüşümü (sidebar) ---
with st.sidebar.expander("💱 Döviz / Kur Ayarları", expanded=False):
    fx_source = st.selectbox("Kur Kaynağı", ["Manuel", "Merkez Bankası (API yok)"] , index=0)
    manual_eur_to_usd = st.number_input("EUR → USD manuel kur", value=1.08, step=0.001, format="%.4f")
    apply_fx = st.checkbox("Eksik para birimlerini dönüştür (Net Sales EUR↔USD)", value=True)



def multiselect_filter(label, col):
    if col not in df.columns:
        return None
    options = sorted(df[col].dropna().astype(str).unique().tolist())
    selected = st.sidebar.multiselect(label, options, default=[])
    return selected if selected else options


f_year = multiselect_filter("Yıl", "Year")

# Ay filtresi: seçenekler ay adı olarak gösterilir, arka planda "01".."12" koduna eşlenir
f_month = None
if "Month" in df.columns:
    month_codes_present = sorted(df["Month"].dropna().astype(str).unique().tolist())
    month_code_to_label = {c: MONTH_NAMES.get(c, c) for c in month_codes_present}
    label_to_code = {v: k for k, v in month_code_to_label.items()}
    month_labels_present = [month_code_to_label[c] for c in month_codes_present]
    selected_month_labels = st.sidebar.multiselect("Ay", month_labels_present, default=[])
    f_month = (
        [label_to_code[lbl] for lbl in selected_month_labels]
        if selected_month_labels
        else month_codes_present
    )

f_scenario = multiselect_filter("Senaryo", "Scenario")
f_region = multiselect_filter("Bölge", "Region")
f_country = multiselect_filter("Ülke", "Country")
f_brand = multiselect_filter("Marka", "Brand")
f_bifs = multiselect_filter("BI/FS", "BI/FS_Label") if "BI/FS_Label" in df.columns else None
f_ph1 = multiselect_filter("Ürün Grubu (PH1)", "PH1")
f_pgroup = multiselect_filter("Ürün Grubu (Pgroup)", "Pgroup")

mask = pd.Series(True, index=df.index)
if f_year is not None:
    mask &= df["Year"].astype(str).isin(f_year)
if f_month is not None and "Month" in df.columns:
    mask &= df["Month"].astype(str).isin(f_month)
if f_scenario is not None and "Scenario" in df.columns:
    mask &= df["Scenario"].astype(str).isin(f_scenario)
if f_region is not None and "Region" in df.columns:
    mask &= df["Region"].astype(str).isin(f_region)
if f_country is not None and "Country" in df.columns:
    mask &= df["Country"].astype(str).isin(f_country)
if f_brand is not None and "Brand" in df.columns:
    mask &= df["Brand"].astype(str).isin(f_brand)
if f_bifs is not None:
    mask &= df["BI/FS_Label"].astype(str).isin(f_bifs)
if f_ph1 is not None and "PH1" in df.columns:
    mask &= df["PH1"].astype(str).isin(f_ph1)
if f_pgroup is not None and "Pgroup" in df.columns:
    mask &= df["Pgroup"].astype(str).isin(f_pgroup)

fdf = df[mask].copy()

if fdf.empty:
    st.warning("Seçilen filtrelerle eşleşen veri yok. Lütfen filtreleri gözden geçirin.")
    st.stop()

# Para birimi: EUR, USD veya ikisi birlikte seçilebilir
selected_currencies = st.sidebar.multiselect(
    "Para Birimi", ["EUR", "USD"], default=["EUR"]
)
if not selected_currencies:
    selected_currencies = ["EUR"]
sales_cols = [f"Net Sales {c}" for c in selected_currencies]
sales_col = sales_cols[0]           # geriye dönük uyumluluk / tekli grafiklerde birincil para birimi
currency = selected_currencies[0]   # birincil para birimi etiketi (tekli grafik başlıklarında kullanılır)
currency_label = " + ".join(selected_currencies)  # her iki para birimi seçiliyse "EUR + USD"
dual_currency = len(sales_cols) > 1

# --- Basit döviz dönüşümü: eksik Net Sales sütunlarını diğer sütundan manuel kur ile üret
if apply_fx:
    # Eğer EUR seçili ama USD sütunu yoksa, USD üret
    if "Net Sales EUR" in df.columns and "Net Sales USD" not in df.columns:
        df["Net Sales USD"] = (df["Net Sales EUR"] * manual_eur_to_usd).round(2)
    # Eğer USD var ama EUR yoksa, EUR üret
    if "Net Sales USD" in df.columns and "Net Sales EUR" not in df.columns:
        df["Net Sales EUR"] = (df["Net Sales USD"] / manual_eur_to_usd).round(2)
    # Yeniden oluşturulan sütunlar için tip temizliği
    for col in ["Net Sales EUR", "Net Sales USD"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)



def sales_agg_kwargs():
    """groupby(...).agg(...) içine geçirilecek, seçili tüm para birimlerini içeren kwargs sözlüğü."""
    return {f"Sales_{c}": (f"Net Sales {c}", "sum") for c in selected_currencies}


def sales_value_cols():
    return [f"Sales_{c}" for c in selected_currencies]


st.sidebar.markdown("---")
st.sidebar.caption(f"Filtrelenmiş kayıt sayısı: {len(fdf):,}".replace(",", "."))

# --------------------------------------------------------------------------------------
# SEKMELER (10 ANALİZ)
# --------------------------------------------------------------------------------------
tabs = st.tabs([
    "📊 Genel Özet",
    "🧠 Otomatik İçgörüler",
    "📈 YOY Karşılaştırma",
    "🗂️ Senaryo Kırılımı",
    "📉 Çoklu Yıl Trend",
    "🌍 Bölge/Ülke",
    "🧱 Ürün Hiyerarşisi",
    "🏷️ Marka",
    "🔧 BI / FS",
    "🥧 Pareto (80/20)",
    "📋 Detaylı Tablo",
    "📅 Aylık Trend & MoM",
    "🎯 Actual vs Budget",
])

# ---------------------------------------------------------------- 1. GENEL ÖZET
with tabs[0]:
    st.subheader("Genel KPI Özeti")
    total_qty = fdf["Quantity"].sum()
    n_countries = fdf["Country"].nunique() if "Country" in fdf.columns else None

    cols = st.columns(1 + 2 * len(selected_currencies) + (1 if n_countries is not None else 0))
    cols[0].metric("Toplam Adet", fmt_num(total_qty))
    idx = 1
    for c in selected_currencies:
        total_sales_c = fdf[f"Net Sales {c}"].sum()
        cols[idx].metric(f"Toplam Ciro ({c})", fmt_money(total_sales_c))
        idx += 1
    for c in selected_currencies:
        total_sales_c = fdf[f"Net Sales {c}"].sum()
        avg_price_c = total_sales_c / total_qty if total_qty else 0
        cols[idx].metric(f"Ort. Birim Fiyat ({c})", f"{avg_price_c:,.2f}".replace(",", "."))
        idx += 1
    if n_countries is not None:
        cols[idx].metric("Ülke Sayısı", n_countries)

    st.markdown("---")
    col1, col2 = st.columns(2)
    if "Country" in fdf.columns:
        top_country = (
            fdf.groupby("Country").agg(**sales_agg_kwargs()).sort_values(sales_value_cols()[0], ascending=False).head(10)
        )
        fig = px.bar(
            top_country[::-1].reset_index(),
            x=sales_value_cols(),
            y="Country",
            orientation="h",
            barmode="group",
            title=f"Ciroya Göre İlk 10 Ülke ({currency_label})",
            labels={"value": "Net Satış", "Country": "Ülke"},
        )
        col1.plotly_chart(fig, width='stretch')
        with col1:
            save_plotly_buttons(fig, "top_countries")

    if "PH1" in fdf.columns:
        top_ph1 = fdf.groupby("PH1").agg(**sales_agg_kwargs())
        primary_col = sales_value_cols()[0]
        fig2 = px.pie(top_ph1, values=primary_col, names=top_ph1.index, title=f"Ürün Grubuna (PH1) Göre Ciro Payı ({currency})")
        col2.plotly_chart(fig2, width='stretch')
        with col2:
            save_plotly_buttons(fig2, "ph1_pie")

# ---------------------------------------------------------------- 2. OTOMATİK İÇGÖRÜLER
with tabs[1]:
    st.subheader("🧠 Otomatik İçgörüler")
    st.caption("Sistem, filtrelenmiş veriyi tarayarak dikkat çekici değişimleri otomatik tespit eder ve önem sırasına göre listeler.")

    ic1, ic2 = st.columns(2)
    insight_threshold = ic1.slider("Değişim eşiği (%)", min_value=5, max_value=50, value=15, step=5,
                                    help="Bu yüzdenin üzerindeki değişimler içgörü olarak raporlanır.")
    insight_dims = ic2.multiselect(
        "Taranacak boyutlar",
        [c for c in ["Country", "PH1", "Pgroup", "Brand", "Region"] if c in fdf.columns],
        default=[c for c in ["Country", "PH1", "Pgroup"] if c in fdf.columns],
    )

    primary_c = selected_currencies[0]
    primary_sales_col = f"Net Sales {primary_c}"
    total_portfolio = fdf[primary_sales_col].sum()
    noise_floor = total_portfolio * 0.01 if total_portfolio else 0  # portföyün <%1'i olan kalemleri gürültü sayıp ele

    insights = []  # (severity_rank, icon, text)

    # --- 1) YOY İçgörüleri ---
    years_present = sorted(fdf["Year"].dropna().unique().tolist())
    if len(years_present) >= 2:
        y_new, y_old = years_present[-1], years_present[-2]
        for dim in insight_dims:
            g_new = fdf[fdf["Year"] == y_new].groupby(dim)[primary_sales_col].sum()
            g_old = fdf[fdf["Year"] == y_old].groupby(dim)[primary_sales_col].sum()
            joined = pd.concat([g_new.rename("new"), g_old.rename("old")], axis=1).fillna(0)
            joined = joined[(joined["new"] + joined["old"]) >= noise_floor]
            for name, row in joined.iterrows():
                pct = pct_change(row["new"], row["old"])
                if pct is None or abs(pct) < insight_threshold:
                    continue
                direction = "arttı" if pct > 0 else "düştü"
                icon = "🟢" if pct > 0 else "🔴"
                insights.append((
                    abs(pct),
                    icon,
                    f"**{name}** ({dim}): {y_new} cirosu {y_old}'e göre **%{abs(pct):.1f} {direction}** "
                    f"({row['old']:,.0f} → {row['new']:,.0f} {primary_c})".replace(",", "."),
                ))

    # --- 2) Ay/Ay (MoM) İçgörüleri ---
    if "Month" in fdf.columns:
        mdf = fdf.copy()
        mdf["YearMonth"] = mdf["Year"].astype(str) + "-" + mdf["Month"].astype(str)
        ym_sorted = sorted(mdf["YearMonth"].dropna().unique().tolist())
        if len(ym_sorted) >= 2:
            ym_new, ym_old = ym_sorted[-1], ym_sorted[-2]
            for dim in insight_dims:
                g_new = mdf[mdf["YearMonth"] == ym_new].groupby(dim)[primary_sales_col].sum()
                g_old = mdf[mdf["YearMonth"] == ym_old].groupby(dim)[primary_sales_col].sum()
                joined = pd.concat([g_new.rename("new"), g_old.rename("old")], axis=1).fillna(0)
                joined = joined[(joined["new"] + joined["old"]) >= noise_floor]
                for name, row in joined.iterrows():
                    pct = pct_change(row["new"], row["old"])
                    if pct is None or abs(pct) < insight_threshold:
                        continue
                    direction = "arttı" if pct > 0 else "düştü"
                    icon = "🟢" if pct > 0 else "🔴"
                    insights.append((
                        abs(pct) * 0.9,  # MoM dalgalanmaları YOY'a göre biraz daha az ağırlıklı (mevsimsellik gürültüsü olabilir)
                        icon,
                        f"**{name}** ({dim}): **{ym_new}** cirosu bir önceki aya ({ym_old}) göre "
                        f"**%{abs(pct):.1f} {direction}** ({row['old']:,.0f} → {row['new']:,.0f} {primary_c})".replace(",", "."),
                    ))

    # --- 3) Actual vs Budget Sapma İçgörüleri ---
    if "Scenario" in fdf.columns:
        def _scen_type(s):
            s_low = str(s).lower()
            if "actual" in s_low:
                return "Actual"
            if "budget" in s_low:
                return "Budget"
            if "5+7" in s_low:
                return "5+7 Forecast"
            return "Diğer"

        sdf = fdf.copy()
        sdf["ScenType"] = sdf["Scenario"].map(_scen_type)
        sdf["ScenYear"] = sdf["Scenario"].astype(str).str.extract(r"(\d{4})")[0].fillna(sdf["Year"])
        avail_types = sdf["ScenType"].unique().tolist()
        base_t = "Budget" if "Budget" in avail_types else ("5+7 Forecast" if "5+7 Forecast" in avail_types else None)

        if base_t and "Actual" in avail_types:
            for yr in sorted(sdf["ScenYear"].dropna().unique().tolist()):
                act = sdf[(sdf["ScenType"] == "Actual") & (sdf["ScenYear"] == yr)]
                bud = sdf[(sdf["ScenType"] == base_t) & (sdf["ScenYear"] == yr)]
                if act.empty or bud.empty:
                    continue
                for dim in insight_dims:
                    a_g = act.groupby(dim)[primary_sales_col].sum()
                    b_g = bud.groupby(dim)[primary_sales_col].sum()
                    joined = pd.concat([a_g.rename("actual"), b_g.rename("budget")], axis=1).fillna(0)
                    joined = joined[(joined["actual"] + joined["budget"]) >= noise_floor]
                    for name, row in joined.iterrows():
                        pct = pct_change(row["actual"], row["budget"])
                        if pct is None or abs(pct) < insight_threshold:
                            continue
                        direction = "üzerinde" if pct > 0 else "altında"
                        icon = "🟢" if pct > 0 else "🔴"
                        insights.append((
                            abs(pct) * 1.1,  # Bütçe sapmaları en yüksek öncelikli (yönetim için en kritik sinyal)
                            icon,
                            f"**{name}** ({dim}, {yr}): Gerçekleşen (Actual), {base_t} hedefinin "
                            f"**%{abs(pct):.1f} {direction}** ({row['budget']:,.0f} → {row['actual']:,.0f} {primary_c})".replace(",", "."),
                        ))

    # --- Sonuçları göster ---
    # --- Ek: Otomatik anomali tespiti (basit) ---
    if anomaly_enable:
        anomalies_df = detect_anomalies(
            fdf,
            dims=insight_dims,
            period=anomaly_period,
            threshold_pct=anomaly_threshold,
            primary_sales_col=primary_sales_col,
            noise_floor=noise_floor,
        )
        if not anomalies_df.empty:
            st.markdown("---")
            st.subheader("⚠️ Otomatik Anomali Tespiti")
            st.write(f"Eşik: %{anomaly_threshold} — Periyot: {anomaly_period}")
            st.dataframe(anomalies_df.style.format({"Old": "{:,.0f}", "New": "{:,.0f}", "Pct": "{:+.2f}%"}), width='stretch')
            st.warning(f"{len(anomalies_df)} anomali tespit edildi. Lütfen listeyi inceleyin.")
    if not insights:
        st.success("Seçilen eşik ve boyutlarda dikkat çekici bir sapma bulunamadı. Eşiği düşürerek tekrar deneyebilirsiniz.")
    else:
        insights.sort(key=lambda x: x[0], reverse=True)
        n_pos = sum(1 for i in insights if i[1] == "🟢")
        n_neg = sum(1 for i in insights if i[1] == "🔴")
        m1, m2, m3 = st.columns(3)
        m1.metric("Toplam İçgörü", len(insights))
        m2.metric("🟢 Pozitif Sinyal", n_pos)
        m3.metric("🔴 Negatif Sinyal", n_neg)
        st.markdown("---")

        for _, icon, text in insights[:30]:
            border_color = SIGNAL_GREEN if icon == "🟢" else SIGNAL_RED
            st.markdown(
                f"""<div style="border-left:4px solid {border_color}; background:#FFFFFF;
                padding:0.6rem 0.9rem; margin-bottom:0.5rem; border-radius:6px;
                box-shadow:0 1px 2px rgba(14,26,43,0.05);">
                {icon} {text}
                </div>""",
                unsafe_allow_html=True,
            )
        if len(insights) > 30:
            st.caption(f"İlk 30 içgörü gösteriliyor (toplam {len(insights)} tespit edildi). Eşiği yükselterek listeyi daraltabilirsiniz.")

# ---------------------------------------------------------------- 3. YOY KARŞILAŞTIRMA
with tabs[2]:
    st.subheader("Yıllık Karşılaştırma (YOY)")
    years_available = sorted(fdf["Year"].dropna().unique().tolist())
    if len(years_available) < 2:
        st.info("YOY karşılaştırması için en az 2 farklı yıl gereklidir. Filtreleri genişletin.")
    else:
        c1, c2 = st.columns(2)
        year_new = c1.selectbox("Bu Yıl", years_available, index=len(years_available) - 1)
        year_old = c2.selectbox(
            "Karşılaştırılan Yıl",
            years_available,
            index=len(years_available) - 2,
        )
        dim = st.selectbox(
            "Kırılım Boyutu",
            [c for c in ["Country", "PH1", "Pgroup", "Brand", "BI/FS_Label", "Region"] if c in fdf.columns],
        )

        g_new = fdf[fdf["Year"] == year_new].groupby(dim).agg(
            Quantity=("Quantity", "sum"), **sales_agg_kwargs()
        )
        g_old = fdf[fdf["Year"] == year_old].groupby(dim).agg(
            Quantity=("Quantity", "sum"), **sales_agg_kwargs()
        )
        comp = g_new.join(g_old, lsuffix=f"_{year_new}", rsuffix=f"_{year_old}", how="outer").fillna(0)
        comp["QTY %"] = comp.apply(
            lambda r: pct_change(r[f"Quantity_{year_new}"], r[f"Quantity_{year_old}"]), axis=1
        )
        fmt_dict = {
            f"Quantity_{year_new}": "{:,.0f}",
            f"Quantity_{year_old}": "{:,.0f}",
            "QTY %": "{:+.2f}%",
        }
        for c in selected_currencies:
            new_col, old_col = f"Sales_{c}_{year_new}", f"Sales_{c}_{year_old}"
            comp[f"{c} %"] = comp.apply(lambda r: pct_change(r[new_col], r[old_col]), axis=1)
            fmt_dict[new_col] = "{:,.0f}"
            fmt_dict[old_col] = "{:,.0f}"
            fmt_dict[f"{c} %"] = "{:+.2f}%"

        primary_new = f"Sales_{selected_currencies[0]}_{year_new}"
        comp = comp.sort_values(primary_new, ascending=False)
        st.dataframe(comp.style.format(fmt_dict), width='stretch')

        fig = px.bar(
            comp.reset_index(),
            x=dim,
            y=[f"Sales_{selected_currencies[0]}_{year_old}", primary_new],
            barmode="group",
            title=f"{year_old} vs {year_new} — {dim} Bazında Ciro ({selected_currencies[0]})",
        )
        st.plotly_chart(fig, width='stretch')

# ---------------------------------------------------------------- 4. SENARYO KIRILIMI
with tabs[3]:
    st.subheader("Senaryo Kırılımı (Actual / Budget / 5+7 vb.)")
    if "Scenario" in fdf.columns:
        scen = fdf.groupby("Scenario").agg(
            Quantity=("Quantity", "sum"), **sales_agg_kwargs()
        ).sort_values(sales_value_cols()[0], ascending=False)
        c1, c2 = st.columns([2, 1])
        fig = px.bar(
            scen.reset_index(), x="Scenario", y=sales_value_cols(), barmode="group",
            title=f"Senaryo Bazında Ciro ({currency_label})", text_auto=".2s"
        )
        c1.plotly_chart(fig, width='stretch')
        fmt_dict = {"Quantity": "{:,.0f}", **{c: "{:,.0f}" for c in sales_value_cols()}}
        c2.dataframe(scen.style.format(fmt_dict), width='stretch')
    else:
        st.info("Veride 'Scenario' sütunu bulunamadı.")

# ---------------------------------------------------------------- 5. ÇOKLU YIL TREND
with tabs[4]:
    st.subheader("Çoklu Yıl Trend")
    group_dim = None
    if "Scenario" in fdf.columns:
        group_dim = "Scenario"
    trend = fdf.groupby(["Year"] + ([group_dim] if group_dim else [])).agg(
        Quantity=("Quantity", "sum"), **sales_agg_kwargs()
    ).reset_index()
    trend = trend.sort_values("Year")

    trend_cols = st.columns(len(selected_currencies))
    for i, c in enumerate(selected_currencies):
        fig = px.line(
            trend, x="Year", y=f"Sales_{c}", color=group_dim, markers=True,
            title=f"Yıllara Göre Ciro Trendi ({c})"
        )
        trend_cols[i].plotly_chart(fig, width='stretch')

    fig2 = px.bar(
        trend, x="Year", y="Quantity", color=group_dim,
        title="Yıllara Göre Adet Trendi", barmode="group"
    )
    st.plotly_chart(fig2, width='stretch')

# ---------------------------------------------------------------- 6. BÖLGE/ÜLKE
with tabs[5]:
    st.subheader("Bölge / Ülke Kırılımı")
    if "Region" in fdf.columns:
        fig = px.treemap(
            fdf, path=[px.Constant("Tümü"), "Region", "Country"] if "Country" in fdf.columns else [px.Constant("Tümü"), "Region"],
            values=sales_col, title=f"Bölge → Ülke Kırılımı ({currency}, treemap boyutu birincil para birimine göre)"
        )
        st.plotly_chart(fig, width='stretch')
    if "Country" in fdf.columns:
        country_g = fdf.groupby("Country").agg(
            Quantity=("Quantity", "sum"), **sales_agg_kwargs()
        ).sort_values(sales_value_cols()[0], ascending=False)
        fmt_dict = {"Quantity": "{:,.0f}", **{c: "{:,.0f}" for c in sales_value_cols()}}
        st.dataframe(country_g.style.format(fmt_dict), width='stretch')

# ---------------------------------------------------------------- 7. ÜRÜN HİYERARŞİSİ
with tabs[6]:
    st.subheader("Ürün Hiyerarşisi (PH1 → PH4)")
    path_cols = [c for c in ["PH1", "PH2", "PH3", "PH4"] if c in fdf.columns]
    if path_cols:
        fig = px.treemap(
            fdf, path=[px.Constant("Tümü")] + path_cols, values=sales_col,
            title=f"Ürün Hiyerarşisi Ciro Dağılımı ({currency}, treemap boyutu birincil para birimine göre)"
        )
        fig.update_traces(root_color="lightgrey")
        st.plotly_chart(fig, width='stretch')
    else:
        st.info("Ürün hiyerarşisi sütunları (PH1-PH4) bulunamadı.")

# ---------------------------------------------------------------- 8. MARKA
with tabs[7]:
    st.subheader("Marka Bazında Performans")
    if "Brand" in fdf.columns:
        brand_g = fdf.groupby("Brand").agg(
            Quantity=("Quantity", "sum"), **sales_agg_kwargs()
        ).sort_values(sales_value_cols()[0], ascending=False)
        c1, c2 = st.columns(2)
        fig = px.bar(
            brand_g.reset_index(), x="Brand", y=sales_value_cols(), barmode="group",
            title=f"Marka Bazında Ciro ({currency_label})"
        )
        c1.plotly_chart(fig, width='stretch')
        fig2 = px.pie(brand_g, values="Quantity", names=brand_g.index, title="Marka Bazında Adet Payı")
        c2.plotly_chart(fig2, width='stretch')
        fmt_dict = {"Quantity": "{:,.0f}", **{c: "{:,.0f}" for c in sales_value_cols()}}
        st.dataframe(brand_g.style.format(fmt_dict), width='stretch')
    else:
        st.info("Veride 'Brand' sütunu bulunamadı.")

# ---------------------------------------------------------------- 9. BI / FS
with tabs[8]:
    st.subheader("Built-In / Free-Standing Kırılımı")
    if "BI/FS_Label" in fdf.columns:
        bifs_g = fdf.groupby("BI/FS_Label").agg(
            Quantity=("Quantity", "sum"), **sales_agg_kwargs()
        ).sort_values(sales_value_cols()[0], ascending=False)
        c1, c2 = st.columns(2)
        fig = px.pie(bifs_g, values=sales_value_cols()[0], names=bifs_g.index, title=f"BI/FS Ciro Payı ({currency})", hole=0.4)
        c1.plotly_chart(fig, width='stretch')

        bifs_trend = fdf.groupby(["Year", "BI/FS_Label"]).agg(**sales_agg_kwargs()).reset_index()
        fig2 = px.line(
            bifs_trend, x="Year", y=sales_value_cols()[0], color="BI/FS_Label", markers=True,
            title=f"Yıllara Göre BI/FS Trendi ({selected_currencies[0]})"
        )
        c2.plotly_chart(fig2, width='stretch')
        fmt_dict = {"Quantity": "{:,.0f}", **{c: "{:,.0f}" for c in sales_value_cols()}}
        st.dataframe(bifs_g.style.format(fmt_dict), width='stretch')
    else:
        st.info("Veride 'BI/FS' sütunu bulunamadı.")

# ---------------------------------------------------------------- 10. PARETO (80/20)
with tabs[9]:
    st.subheader("Pareto Analizi (80/20)")
    dim_options = [c for c in ["Country", "PH1", "Pgroup", "PH4", "Brand"] if c in fdf.columns]
    if dim_options:
        if dual_currency:
            st.caption(f"Not: Pareto analizi tek para birimi üzerinden hesaplanır. Şu an **{currency}** kullanılıyor.")
        pareto_dim = st.selectbox("Boyut Seç", dim_options, key="pareto_dim")
        g = fdf.groupby(pareto_dim)[sales_col].sum().sort_values(ascending=False)
        cum_pct = g.cumsum() / g.sum() * 100
        pareto_df = pd.DataFrame({"Sales": g, "Cumulative %": cum_pct})

        fig = go.Figure()
        fig.add_bar(x=pareto_df.index, y=pareto_df["Sales"], name=f"Ciro ({currency})")
        fig.add_trace(go.Scatter(
            x=pareto_df.index, y=pareto_df["Cumulative %"], name="Kümülatif %",
            yaxis="y2", mode="lines+markers"
        ))
        fig.update_layout(
            title=f"{pareto_dim} Bazında Pareto Analizi",
            yaxis=dict(title=f"Ciro ({currency})"),
            yaxis2=dict(title="Kümülatif %", overlaying="y", side="right", range=[0, 105]),
            legend=dict(orientation="h", y=1.1),
        )
        st.plotly_chart(fig, width='stretch')

        n80 = (cum_pct <= 80).sum() + 1
        st.info(f"Cironun %80'i, {pareto_dim} bazında ilk **{n80}** kalemden geliyor (toplam {len(g)} kalem).")
    else:
        st.info("Pareto analizi için uygun bir boyut bulunamadı.")

# ---------------------------------------------------------------- 11. DETAYLI TABLO
with tabs[10]:
    st.subheader("Detaylı Filtrelenebilir Tablo")
    st.dataframe(fdf, width='stretch', height=500)

    excel_bytes = to_excel_bytes(fdf)
    st.download_button(
        label="📥 Filtrelenmiş Veriyi Excel Olarak İndir",
        data=excel_bytes,
        file_name="filtrelenmis_veri.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# ---------------------------------------------------------------- 12. AYLIK TREND & MoM
with tabs[11]:
    st.subheader("Aylık Mevsimsellik ve Ay/Ay (MoM) Büyüme")
    if "Month" not in fdf.columns:
        st.info("Veride 'Month' sütunu bulunamadı, bu analiz için gereklidir.")
    else:
        mdf = fdf.copy()
        # Map month names; ensure we don't assign categorical values into categorical series
        mdf["Month_Name"] = (
            mdf["Month"].map(MONTH_NAMES)
            .astype(object)
            .fillna(mdf["Month"].astype(object))
        )

        # --- Mevsimsellik: yıllar üst üste, ay bazında ---
        seasonal = mdf.groupby(["Year", "Month", "Month_Name"]).agg(
            Quantity=("Quantity", "sum"), **sales_agg_kwargs()
        ).reset_index().sort_values(["Year", "Month"])

        primary_sales = sales_value_cols()[0]
        fig = px.line(
            seasonal, x="Month_Name", y=primary_sales, color="Year", markers=True,
            category_orders={"Month_Name": [MONTH_NAMES[m] for m in MONTH_ORDER]},
            title=f"Aylık Mevsimsellik — Yıllar Üst Üste ({selected_currencies[0]})",
        )
        st.plotly_chart(fig, width='stretch')

        st.markdown("---")
        st.markdown("**Ay/Ay (MoM) Büyüme Tablosu**")

        # Kronolojik sırada Year-Month bazında toplam
        chron = mdf.groupby(["Year", "Month"]).agg(
            Quantity=("Quantity", "sum"), **sales_agg_kwargs()
        ).reset_index()
        chron["YearMonth"] = chron["Year"].astype(str) + "-" + chron["Month"].astype(str)
        chron = chron.sort_values("YearMonth").reset_index(drop=True)
        chron["MoM Quantity %"] = chron["Quantity"].pct_change() * 100
        display_cols = ["YearMonth", "Quantity", "MoM Quantity %"]
        fmt_dict = {"Quantity": "{:,.0f}", "MoM Quantity %": "{:+.2f}%"}
        for c in selected_currencies:
            sc = f"Sales_{c}"
            chron[f"MoM {c} %"] = chron[sc].pct_change() * 100
            chron[f"3-Ay Ort. ({c})"] = chron[sc].rolling(3, min_periods=1).mean()
            display_cols += [sc, f"MoM {c} %", f"3-Ay Ort. ({c})"]
            fmt_dict[sc] = "{:,.0f}"
            fmt_dict[f"MoM {c} %"] = "{:+.2f}%"
            fmt_dict[f"3-Ay Ort. ({c})"] = "{:,.0f}"

        st.dataframe(chron[display_cols].style.format(fmt_dict), width='stretch')

        fig2 = go.Figure()
        fig2.add_bar(x=chron["YearMonth"], y=chron[primary_sales], name=f"Aylık Ciro ({selected_currencies[0]})")
        fig2.add_trace(go.Scatter(
            x=chron["YearMonth"], y=chron[f"3-Ay Ort. ({selected_currencies[0]})"],
            name="3-Ay Hareketli Ortalama", mode="lines", line=dict(width=3),
        ))
        fig2.update_layout(title="Kronolojik Ciro Trendi + 3-Aylık Hareketli Ortalama")
        st.plotly_chart(fig2, width='stretch')

# ---------------------------------------------------------------- 13. ACTUAL vs BUDGET SAPMA
with tabs[12]:
    st.subheader("Actual vs Budget / Forecast Sapma Analizi")
    if "Scenario" not in fdf.columns:
        st.info("Veride 'Scenario' sütunu bulunamadı, bu analiz için gereklidir.")
    else:
        scen_series = fdf["Scenario"].astype(str)

        def scen_type(s):
            s_low = s.lower()
            if "actual" in s_low:
                return "Actual"
            if "budget" in s_low:
                return "Budget"
            if "5+7" in s_low or "5 + 7" in s_low:
                return "5+7 Forecast"
            return "Diğer"

        sdf = fdf.copy()
        sdf["ScenType"] = scen_series.map(scen_type)
        sdf["ScenYear"] = sdf["Scenario"].astype(str).str.extract(r"(\d{4})")[0].fillna(sdf["Year"])

        avail_types = sorted(sdf["ScenType"].unique().tolist())
        if "Actual" not in avail_types or len(avail_types) < 2:
            st.info("Karşılaştırma için hem 'Actual' hem de 'Budget'/'5+7' senaryoları gereklidir.")
        else:
            c1, c2 = st.columns(2)
            base_type = c1.selectbox("Karşılaştırılacak Senaryo", [t for t in avail_types if t != "Actual"])
            comp_years = sorted(sdf["ScenYear"].dropna().unique().tolist())
            sel_year = c2.selectbox("Yıl", comp_years, index=len(comp_years) - 1)

            dim2 = st.selectbox(
                "Kırılım Boyutu",
                [c for c in ["Country", "PH1", "Pgroup", "Brand", "Region"] if c in sdf.columns],
                key="variance_dim",
            )

            act = sdf[(sdf["ScenType"] == "Actual") & (sdf["ScenYear"] == sel_year)]
            bud = sdf[(sdf["ScenType"] == base_type) & (sdf["ScenYear"] == sel_year)]

            primary_c = selected_currencies[0]
            var_frames = []
            fmt_dict = {}
            for c in selected_currencies:
                col = f"Net Sales {c}"
                a_g = act.groupby(dim2)[col].sum().rename(f"Actual_{c}")
                b_g = bud.groupby(dim2)[col].sum().rename(f"{base_type}_{c}")
                merged = pd.concat([a_g, b_g], axis=1).fillna(0)
                merged[f"Sapma_{c}"] = merged[f"Actual_{c}"] - merged[f"{base_type}_{c}"]
                merged[f"Sapma % ({c})"] = merged.apply(
                    lambda r: pct_change(r[f"Actual_{c}"], r[f"{base_type}_{c}"]), axis=1
                )
                var_frames.append(merged)
                fmt_dict[f"Actual_{c}"] = "{:,.0f}"
                fmt_dict[f"{base_type}_{c}"] = "{:,.0f}"
                fmt_dict[f"Sapma_{c}"] = "{:+,.0f}"
                fmt_dict[f"Sapma % ({c})"] = "{:+.2f}%"

            var_df = pd.concat(var_frames, axis=1)
            var_df = var_df.sort_values(f"Actual_{primary_c}", ascending=False)

            st.dataframe(var_df.style.format(fmt_dict), width='stretch')

            fig3 = px.bar(
                var_df.reset_index(), x=dim2, y=[f"Actual_{primary_c}", f"{base_type}_{primary_c}"], barmode="group",
                title=f"{sel_year} — Actual vs {base_type} ({dim2} Bazında, {primary_c})",
            )
            st.plotly_chart(fig3, width='stretch')

            metric_cols = st.columns(len(selected_currencies))
            for i, c in enumerate(selected_currencies):
                total_actual = var_df[f"Actual_{c}"].sum()
                total_budget = var_df[f"{base_type}_{c}"].sum()
                total_var_pct = pct_change(total_actual, total_budget)
                if total_var_pct is not None:
                    metric_cols[i].metric(
                        f"Genel Sapma ({sel_year}: Actual vs {base_type}, {c})",
                        f"{total_var_pct:+.2f}%",
                        delta=f"{total_actual - total_budget:+,.0f} {c}",
                    )

st.markdown("---")
st.caption(f"⚙️ {BRAND_NAME} · Kurumsal Satış Analiz Platformu — veriler yalnızca bu oturumda, tarayıcınızda işlenir.")

# py -m streamlit run reporting_analysis_app.py
