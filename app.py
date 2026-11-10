from pandas.errors import ClosedFileError
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import os
import json
import base64
import re
from datetime import datetime
from io import BytesIO
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# --- Configuration & Paths ---
FEEDBACK_DIR = "feedback"
CONFIG_DIR = "config"
os.makedirs(FEEDBACK_DIR, exist_ok=True)
os.makedirs(CONFIG_DIR, exist_ok=True)

CATEGORIES_JSON = os.path.join(CONFIG_DIR, "categories.json")
LOGO_PNG = os.path.join("assets", "deen_logo.png")

# Default Category Mapping Rules
DEFAULT_CATEGORY_MAPPING = {
    'Boxer': ['boxer'],
    'Tank Top': ['tank top', 'tanktop', 'tank', 'top'],
    'Cuban Shirt': ['Cuban Collar Shirt', 'cuban'],
    'Panjabi': ['panjabi', 'punjabi'],
    'Jeans': ['jeans'],
    'Formal Shirt': ['executive', 'formal'],
    'Denim Shirt': ['denim shirt'],
    'Flannel Shirt': ['flannel'],
    'Polo Shirt': ['polo'],
    'Cargo Trousers': ['cargo', 'cargo trousers'],
    'Trousers': ['trousers', 'pant', 'trcargo', 'trouser', 'joggers', 'track pant', 'jogger'],
    'Twill': ['twill chino', 'twill', 'chino'],
    'Mask': ['mask'],
    'Water Bottle': ['water bottle'],
    'Contrast Shirt': ['contrast'],
    'Turtleneck': ['turtleneck', 'mock neck'],
    'Drop Shoulder': ['drop', 'shoulder'],
    'Wallet': ['wallet'],
    'Card Holder': ['card holder'],
    'Passport Holder': ['passport holder'],
    'Kaftan Shirt': ['kaftan'],
    'Active Wear': ['active wear'],
    'Jersy': ['jersy'],
    'Sweatshirt': ['sweatshirt', 'hoodie', 'pullover'],
    'Jacket': ['jacket', 'outerwear', 'coat'],
    'Belt': ['belt'],
    'Sweater': ['sweater', 'cardigan', 'knitwear'],
    'Cap': ['cap'],
    'Leather Bag': ['bag', 'backpack'],
    'Mug': ['mug'],
    'Hoodie': ['hoodie'],
    'Knit Cap': ['knit cap']
}

def load_category_mapping():
    """Loads custom categories if available, else writes & returns default."""
    if os.path.exists(CATEGORIES_JSON):
        try:
            with open(CATEGORIES_JSON, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            pass
    with open(CATEGORIES_JSON, 'w', encoding='utf-8') as f:
        json.dump(DEFAULT_CATEGORY_MAPPING, f, indent=4)
    return DEFAULT_CATEGORY_MAPPING.copy()

CATEGORY_MAPPING = load_category_mapping()

# Synchronized Vibrant Chart Palette System
VIBRANT_PALETTE = [
    '#2563EB', '#10B981', '#F59E0B', '#8B5CF6', '#EC4899', '#06B6D4',
    '#F97316', '#14B8A6', '#6366F1', '#EF4444', '#84CC16', '#D946EF',
    '#0EA5E9', '#F43F5E', '#EAB308', '#3B82F6', '#059669', '#A855F7',
    '#FB923C', '#2DD4BF', '#FB7185', '#4ADE80', '#38BDF8', '#C084FC',
    '#F87171', '#34D399', '#FBBF24', '#818CF8', '#F472B6', '#22D3EE',
    '#A3E635', '#E879F9', '#60A5FA', '#1E40AF', '#EA580C', '#0D9488',
    '#9333EA', '#BE123C', '#CA8A04', '#4F46E5'
]

CATEGORY_PALETTE = (
    VIBRANT_PALETTE +
    list(px.colors.qualitative.Bold) +
    list(px.colors.qualitative.Vivid) +
    list(px.colors.qualitative.Prism)
)

ALL_KNOWN_CATEGORIES = sorted(list(set(
    list(DEFAULT_CATEGORY_MAPPING.keys()) +
    ['FS T-Shirt', 'T-Shirt', 'FS Shirt', 'Casual Shirt', 'Others']
)))

CATEGORY_COLOR_MAP = {
    cat: CATEGORY_PALETTE[i % len(CATEGORY_PALETTE)]
    for i, cat in enumerate(ALL_KNOWN_CATEGORIES)
}
CATEGORY_COLOR_MAP['Others'] = '#94A3B8'

def get_category_color_map(categories):
    """Returns a deterministic, consistent vibrant color mapping for any subset of categories."""
    color_map = {}
    for cat in sorted(categories):
        if cat == 'Others':
            color_map[cat] = '#94A3B8'
        elif cat in CATEGORY_COLOR_MAP:
            color_map[cat] = CATEGORY_COLOR_MAP[cat]
        else:
            idx = abs(hash(cat)) % len(CATEGORY_PALETTE)
            color_map[cat] = CATEGORY_PALETTE[idx]
    return color_map

def load_logo():
    """Loads brand logo as base64 string for safe embedding."""
    if os.path.exists(LOGO_PNG):
        try:
            with open(LOGO_PNG, "rb") as image_file:
                return base64.b64encode(image_file.read()).decode()
        except Exception:
            return None
    return None

# --- Consolidated Premium Styling ---

def apply_custom_styles():
    """Applies a consolidated, high-contrast, theme-adaptive CSS design system."""
    st.markdown("""
        <style>
        /* Import Modern Typography */
        @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');

        /* Theme Color Variables */
        :root {
            --primary-color: #2563EB;
            --background-color: #ffffff;
            --secondary-background-color: #f8fafc;
            --text-color: #0f172a;
            --text-muted: #475569;
            --card-border: #e2e8f0;
        }

        @media (prefers-color-scheme: dark) {
            :root {
                --primary-color: #3B82F6;
                --background-color: #0e1117;
                --secondary-background-color: #1e2430;
                --text-color: #f8fafc;
                --text-muted: #cbd5e1;
                --card-border: rgba(255, 255, 255, 0.16);
            }
        }

        /* Base Typography */
        html, body, [class*="css"] {
            font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, sans-serif;
        }
        .main { background-color: transparent; }

        /* Uniform Column & Card Dimensions */
        div[data-testid="column"] {
            display: flex !important;
            flex-direction: column !important;
            flex: 1 1 0px !important;
        }

        div[data-testid="column"] > div {
            display: flex !important;
            flex-direction: column !important;
            flex: 1 1 auto !important;
            height: 100% !important;
        }

        /* Metric Styling - Identical Uniform Height, Baseline Alignment & High Contrast */
        div[data-testid="stMetric"] { 
            background-color: var(--secondary-background-color) !important; 
            padding: 16px 16px !important; 
            border-radius: 16px !important; 
            box-shadow: 0 4px 15px rgba(0,0,0,0.05) !important;
            border: 1px solid var(--card-border) !important;
            transition: all 0.3s ease !important;
            height: 120px !important;
            min-height: 120px !important;
            max-height: 120px !important;
            display: flex !important;
            flex-direction: row !important;
            flex-wrap: wrap !important;
            align-content: center !important;
            align-items: baseline !important;
            box-sizing: border-box !important;
        }
        
        div[data-testid="stMetric"]:hover {
            transform: translateY(-4px) !important;
            box-shadow: 0 8px 25px rgba(0,0,0,0.12) !important;
            border-color: var(--primary-color) !important;
        }

        [data-testid="stMetricLabel"] {
            width: 100% !important;
            flex: 0 0 100% !important;
            font-weight: 600 !important;
            font-size: 0.95rem !important;
            margin-bottom: 6px !important;
            color: var(--text-muted) !important;
            white-space: nowrap !important;
            overflow: hidden !important;
            text-overflow: ellipsis !important;
        }

        [data-testid="stMetricValue"], 
        [data-testid="stMetricValue"] > div {
            display: inline-flex !important;
            align-items: baseline !important;
            font-weight: 800 !important;
            font-size: clamp(1.05rem, 1.3vw, 1.35rem) !important;
            color: var(--text-color) !important;
            white-space: nowrap !important;
            overflow: visible !important;
            text-overflow: clip !important;
            line-height: 1.2 !important;
        }

        /* In-Line Smaller Margin % in Metric Card - High Contrast Green */
        [data-testid="stMetricDelta"],
        [data-testid="stMetricDelta"] > div,
        [data-testid="stMetricDelta"] span {
            display: inline-flex !important;
            align-items: baseline !important;
            margin-left: 6px !important;
            font-size: 0.82rem !important;
            font-weight: 700 !important;
            color: #10b981 !important;
            background: transparent !important;
            padding: 0 !important;
            line-height: 1.2 !important;
        }

        [data-testid="stMetricDelta"] svg {
            display: none !important;
        }

        /* Donut Chart Transparent Slice Separator - Adaptive Theme Stroke */
        div[data-testid="stPlotlyChart"] .js-plotly-plot .slice path,
        div[data-testid="stPlotlyChart"] .js-plotly-plot path.surface {
            stroke: var(--background-color, #ffffff) !important;
            stroke-width: 3px !important;
            stroke-linejoin: round !important;
        }

        /* UI Components - Inputs & Selectboxes */
        div[data-testid="stSelectbox"] > div, 
        div[data-testid="stTextInput"] > div,
        div[data-testid="stTextArea"] > div {
            background-color: var(--secondary-background-color) !important;
            color: var(--text-color) !important;
            border-radius: 12px !important;
            border: 1px solid var(--card-border) !important;
        }

        /* Buttons - High Contrast Primary */
        .stButton>button { 
            width: 100%; 
            border-radius: 12px; 
            height: 3.5em; 
            background-color: var(--primary-color) !important;
            color: white !important;
            font-weight: 700;
            border: none;
            transition: all 0.2s ease;
        }
        .stButton>button:hover {
            opacity: 0.9;
            transform: scale(1.01);
        }

        /* Tabs & Dividers */
        div[data-testid="stTabs"] button {
            color: var(--text-muted) !important;
            opacity: 0.85;
        }
        div[data-testid="stTabs"] button[aria-selected="true"] {
            color: var(--primary-color) !important;
            opacity: 1;
            font-weight: 700 !important;
            border-bottom: 2px solid var(--primary-color) !important;
        }
        hr { border-color: rgba(128, 128, 128, 0.2) !important; }

        /* Expander & Notifications */
        div[data-testid="stExpander"], div[data-testid="stNotification"] {
            background-color: var(--secondary-background-color) !important;
            border: 1px solid var(--card-border) !important;
            border-radius: 14px !important;
            color: var(--text-color) !important;
        }

        /* Main Container Padding to Avoid Sticky Footer Overlap */
        .block-container {
            padding-bottom: 150px !important;
            padding-top: 3rem !important;
        }

        /* Sticky Footer - Dynamic Contrast & Always Fixed */
        .sticky-footer {
            position: fixed !important;
            left: 0 !important;
            bottom: 0 !important;
            right: 0 !important;
            width: 100% !important;
            background-color: var(--background-color) !important;
            color: var(--text-color) !important;
            text-align: center;
            padding: 18px 0;
            border-top: 1px solid rgba(128, 128, 128, 0.2);
            z-index: 999999;
            font-size: 0.9rem;
            backdrop-filter: blur(10px);
            box-shadow: 0 -5px 25px rgba(0,0,0,0.08);
        }

        /* High-Contrast Dark View Styling (OS dark mode) */
        @media (prefers-color-scheme: dark) {
            body, .stApp {
                color: #f8fafc !important;
            }
            h1, h2, h3, h4, h5, h6 {
                color: #ffffff !important;
            }
            p, span, label, .stMarkdown {
                color: #f1f5f9 !important;
            }
            .stCaption, [data-testid="stCaptionContainer"] {
                color: #cbd5e1 !important;
            }
            div[data-testid="stMetric"] {
                background-color: #1e2430 !important;
                border: 1px solid rgba(255, 255, 255, 0.18) !important;
                box-shadow: 0 4px 20px rgba(0,0,0,0.3) !important;
            }
            [data-testid="stMetricLabel"] {
                color: #cbd5e1 !important;
            }
            [data-testid="stMetricValue"], [data-testid="stMetricValue"] > div {
                color: #ffffff !important;
            }
            [data-testid="stMetricDelta"], [data-testid="stMetricDelta"] > div, [data-testid="stMetricDelta"] span {
                color: #34d399 !important;
            }
            div[data-testid="stSelectbox"] > div, 
            div[data-testid="stTextInput"] > div,
            div[data-testid="stTextArea"] > div {
                background-color: #1e2430 !important;
                color: #f8fafc !important;
                border: 1px solid rgba(255, 255, 255, 0.25) !important;
            }
            div[data-testid="stExpander"], div[data-testid="stNotification"] {
                background-color: #1e2430 !important;
                border: 1px solid rgba(255, 255, 255, 0.18) !important;
                color: #f8fafc !important;
            }
            .sticky-footer {
                background-color: #0e1117 !important;
                color: #f8fafc !important;
                border-top: 1px solid rgba(255, 255, 255, 0.18) !important;
            }
            .brand-name {
                color: #f8fafc !important;
            }
            div[data-testid="stPlotlyChart"] .js-plotly-plot .slice path,
            div[data-testid="stPlotlyChart"] .js-plotly-plot path.surface {
                stroke: #0e1117 !important;
            }
        }

        /* Streamlit In-App Dark Theme Overrides (when toggled via Streamlit menu) */
        [data-theme="dark"] body, [data-theme="dark"] .stApp {
            color: #f8fafc !important;
        }
        [data-theme="dark"] h1, [data-theme="dark"] h2, [data-theme="dark"] h3,
        [data-theme="dark"] h4, [data-theme="dark"] h5, [data-theme="dark"] h6 {
            color: #ffffff !important;
        }
        [data-theme="dark"] p, [data-theme="dark"] span, [data-theme="dark"] label, [data-theme="dark"] .stMarkdown {
            color: #f1f5f9 !important;
        }
        [data-theme="dark"] .stCaption, [data-theme="dark"] [data-testid="stCaptionContainer"] {
            color: #cbd5e1 !important;
        }
        [data-theme="dark"] div[data-testid="stMetric"] {
            background-color: #1e2430 !important;
            border: 1px solid rgba(255, 255, 255, 0.18) !important;
            box-shadow: 0 4px 20px rgba(0,0,0,0.3) !important;
        }
        [data-theme="dark"] [data-testid="stMetricLabel"] {
            color: #cbd5e1 !important;
        }
        [data-theme="dark"] [data-testid="stMetricValue"], [data-theme="dark"] [data-testid="stMetricValue"] > div {
            color: #ffffff !important;
        }
        [data-theme="dark"] [data-testid="stMetricDelta"], [data-theme="dark"] [data-testid="stMetricDelta"] > div, [data-theme="dark"] [data-testid="stMetricDelta"] span {
            color: #34d399 !important;
        }
        [data-theme="dark"] div[data-testid="stSelectbox"] > div, 
        [data-theme="dark"] div[data-testid="stTextInput"] > div,
        [data-theme="dark"] div[data-testid="stTextArea"] > div {
            background-color: #1e2430 !important;
            color: #f8fafc !important;
            border: 1px solid rgba(255, 255, 255, 0.25) !important;
        }
        [data-theme="dark"] div[data-testid="stExpander"], [data-theme="dark"] div[data-testid="stNotification"] {
            background-color: #1e2430 !important;
            border: 1px solid rgba(255, 255, 255, 0.18) !important;
            color: #f8fafc !important;
        }
        [data-theme="dark"] .sticky-footer {
            background-color: #0e1117 !important;
            color: #f8fafc !important;
            border-top: 1px solid rgba(255, 255, 255, 0.18) !important;
        }
        [data-theme="dark"] .brand-name {
            color: #f8fafc !important;
        }
        [data-theme="dark"] div[data-testid="stPlotlyChart"] .js-plotly-plot .slice path,
        [data-theme="dark"] div[data-testid="stPlotlyChart"] .js-plotly-plot path.surface {
            stroke: #0e1117 !important;
        }
        .footer-content-inner {
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 6px;
        }
        .brand-wrapper {
            display: flex;
            align-items: center;
            gap: 2px;
        }
        .small-logo {
            height: 24px;
            width: auto;
            filter: grayscale(0%);
        }
        .brand-name {
            font-weight: 700;
            color: var(--primary-color);
        }

        /* Responsive Mobile Adjustments */
        @media (max-width: 768px) {
            .block-container {
                padding-left: 1rem !important;
                padding-right: 1rem !important;
                padding-bottom: 220px !important;
            }
            .footer-content-inner {
                flex-direction: column;
                gap: 8px;
            }
            [data-testid="stMetricValue"] {
                font-size: 1.3rem !important;
            }
            .sticky-footer {
                padding: 15px 5px !important;
            }
        }
        </style>
        """, unsafe_allow_html=True)
    
    # Active theme synchronization helper
    import streamlit.components.v1 as components
    components.html(
        """
        <script>
        function syncTheme() {
            try {
                const doc = window.parent.document;
                const app = doc.querySelector('.stApp');
                if (app) {
                    const bg = window.parent.getComputedStyle(app).backgroundColor;
                    const m = bg.match(/rgba?\\((\\d+),\\s*(\\d+),\\s*(\\d+)/);
                    if (m) {
                        const lum = (0.299 * parseInt(m[1]) + 0.587 * parseInt(m[2]) + 0.114 * parseInt(m[3])) / 255;
                        const theme = lum < 0.5 ? 'dark' : 'light';
                        if (doc.documentElement.getAttribute('data-theme') !== theme) {
                            doc.documentElement.setAttribute('data-theme', theme);
                            app.setAttribute('data-theme', theme);
                        }
                    }
                }
            } catch(e) {}
        }
        syncTheme();
        setInterval(syncTheme, 800);
        </script>
        """,
        height=0,
        width=0
    )

# --- Business Logic & Data Processing ---

def log_event(event_type, details):
    """Logs system events to feedback/system_logs.json safely."""
    log_file = os.path.join(FEEDBACK_DIR, "system_logs.json")
    entry = {
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "type": event_type,
        "details": details
    }
    try:
        logs = []
        if os.path.exists(log_file):
            with open(log_file, "r", encoding="utf-8") as f:
                logs = json.load(f)
        logs.append(entry)
        with open(log_file, "w", encoding="utf-8") as f:
            json.dump(logs[-100:], f, indent=4)
    except (IOError, json.JSONDecodeError, OSError):
        pass

def matches_keyword(keyword, text_lower):
    """Matches a keyword against lowercase text safely, using word boundaries for short terms."""
    kw = keyword.strip().lower()
    if not kw:
        return False
    # For short terms (<= 4 chars like 'top', 'cap', 'tee', 'fs'), use word boundary
    if len(kw) <= 4:
        pattern = r'(?<![a-zA-Z0-9])' + re.escape(kw) + r'(?![a-zA-Z0-9])'
        return bool(re.search(pattern, text_lower))
    return kw in text_lower

def get_product_category(name, mapping_dict=None):
    """Categorizes product based on keywords with length prioritization and word boundaries."""
    if mapping_dict is None:
        mapping_dict = load_category_mapping()
    name_str = str(name).lower()
    
    # 1. Match configured categories (longer, more specific keywords first)
    for cat, keywords in mapping_dict.items():
        sorted_kw = sorted(keywords, key=len, reverse=True)
        if any(matches_keyword(kw, name_str) for kw in sorted_kw):
            return cat
    
    # 2. Special handling for T-Shirts and Shirts
    fs_keywords = ['full sleeve', 'long sleeve', 'fs', 'l/s']
    is_fs = any(matches_keyword(kw, name_str) for kw in fs_keywords)
    
    if any(matches_keyword(kw, name_str) for kw in ['t-shirt', 't shirt', 'tee']):
        return 'FS T-Shirt' if is_fs else 'T-Shirt'
    if matches_keyword('shirt', name_str):
        return 'FS Shirt' if is_fs else 'Casual Shirt'
        
    return 'Others'

def find_columns(df):
    """Auto-detects required and optional columns with priority-based alias matching."""
    mapping = {
        'name': ['item name', 'product name', 'product', 'item', 'title', 'description', 'name'],
        'cost': ['unit price', 'selling price', 'item price', 'price', 'mrp', 'rate', 'item cost', 'cost'],
        'qty': ['quantity', 'qty', 'units', 'sold', 'count', 'total quantity'],
        'cogs': [
            'cogs', 'total cogs', 'unit cogs', 'cogs total', 'cogs unit', 'item cogs', 'product cogs',
            'cost of goods sold', 'cost of good sold', 'cost of goods',
            'c.o.g.s', 'c.o.g.s.', 'cogs(tk)', 'cogs (tk)', 'cogs (bdt)', 'cogs(bdt)', 'cogs amount',
            'cost price', 'buying price', 'purchase price', 'purchase cost', 'buying rate', 
            'cost per item', 'unit cost', 'base cost', 'cost'
        ],
        'date': ['date', 'order date', 'month', 'time', 'created at'],
        'order_id': ['order id', 'order #', 'invoice number', 'invoice #', 'order number', 'transaction id', 'id'],
        'phone': ['phone', 'contact', 'mobile', 'cell', 'phone number', 'customer phone']
    }
    found = {}
    actual_cols = list(df.columns)
    lower_cols = [str(c).strip().lower() for c in actual_cols]
    
    for key, aliases in mapping.items():
        used_cols = set(found.values())
        # Exact match
        for alias in aliases:
            if alias in lower_cols:
                candidate = actual_cols[lower_cols.index(alias)]
                if candidate not in used_cols:
                    found[key] = candidate
                    break
        # Partial match
        if key not in found:
            for i, col in enumerate(lower_cols):
                candidate = actual_cols[i]
                if candidate in used_cols:
                    continue
                if key == 'cogs' and any(skip in col for skip in ['shipping', 'delivery', 'courier']):
                    continue
                if key == 'cogs':
                    cogs_partial_aliases = [a for a in aliases if a != 'cost']
                    if any(alias in col for alias in cogs_partial_aliases):
                        found[key] = candidate
                        break
                elif any(alias in col for alias in aliases):
                    found[key] = candidate
                    break
    return found

def clean_numeric_series(series):
    """Fast vectorized numeric cleaning removing currency, commas, and whitespace while preserving negative numbers."""
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series.fillna(0), errors='coerce').fillna(0)
    
    s_str = series.astype(str).str.strip()
    is_paren_negative = s_str.str.startswith('(') & s_str.str.endswith(')')
    
    cleaned = s_str.str.replace(r'[^0-9.-]', '', regex=True)
    cleaned = cleaned.str.replace(r'(?!^)-', '', regex=True)
    
    numeric = pd.to_numeric(cleaned, errors='coerce').fillna(0)
    numeric = np.where(is_paren_negative, -numeric.abs(), numeric)
    return pd.Series(numeric, index=series.index)

@st.cache_data(show_spinner="Loading sales data...")
def load_data(file_bytes, filename):
    """Loads CSV or Excel data from uploaded bytes with caching."""
    buf = BytesIO(file_bytes)
    if filename.lower().endswith('.csv'):
        return pd.read_csv(buf)
    else:
        return pd.read_excel(buf)

@st.cache_data(show_spinner="Processing analytics & categorizing products...")
def process_analytics(df, mapping):
    """Core data processing, COGS heuristics, profit margins, and metric aggregations."""
    df = df.copy()
    
    # 1. Clean Data (Vectorized)
    df['Clean_Name'] = df[mapping['name']].fillna('Unknown').astype(str)
    exclude_keywords = 'choose any|shoose any|combo|bundle|bundel'
    df = df[~df['Clean_Name'].str.contains(exclude_keywords, case=False, na=False)]

    df['Clean_Cost'] = clean_numeric_series(df[mapping['cost']])
    df['Clean_Qty'] = clean_numeric_series(df[mapping['qty']])
    df.loc[df['Clean_Qty'] < 0, 'Clean_Qty'] = 0
    df['Total Amount'] = df['Clean_Cost'] * df['Clean_Qty']
    
    # 2. Process COGS (Cost of Goods Sold) if mapped
    has_cogs = False
    total_cogs = None
    if mapping.get('cogs') and mapping['cogs'] in df.columns:
        cogs_col = mapping['cogs']
        raw_cogs = clean_numeric_series(df[cogs_col])
        cogs_col_lower = str(cogs_col).strip().lower()
        
        is_total_cogs = any(term in cogs_col_lower for term in ['total cogs', 'total_cogs', 'cogs total', 'cogs_total', 'total cost'])
        
        # Heuristic: if not explicitly named total, check if values reflect unit or total COGS
        if not is_total_cogs:
            multi_qty = df['Clean_Qty'] > 1
            if multi_qty.any():
                cost_nonzero = df.loc[multi_qty, 'Clean_Cost'].replace(0, pd.NA).dropna()
                if not cost_nonzero.empty:
                    valid_idx = cost_nonzero.index
                    ratio = raw_cogs.loc[valid_idx] / df.loc[valid_idx, 'Clean_Cost']
                    if (ratio > 1.1).mean() > 0.5:
                        is_total_cogs = True
        
        if is_total_cogs:
            df['Total COGS'] = raw_cogs.copy()
            df.loc[df['Clean_Qty'] == 0, 'Total COGS'] = 0
            df['Unit COGS'] = (df['Total COGS'] / df['Clean_Qty'].replace(0, 1)).round(2)
        else:
            df['Unit COGS'] = raw_cogs.copy()
            df['Total COGS'] = df['Unit COGS'] * df['Clean_Qty']
            
        has_cogs = True
        total_cogs = float(df['Total COGS'].sum())
    
    # 3. Fast unique-name categorization
    active_mapping = load_category_mapping()
    unique_names = df['Clean_Name'].unique()
    cat_lookup = {name: get_product_category(name, active_mapping) for name in unique_names}
    df['Category'] = df['Clean_Name'].map(cat_lookup).fillna('Others')
    
    # 4. Timeframe Detection
    timeframe = ""
    if mapping.get('date') and mapping['date'] in df.columns:
        try:
            dates = pd.to_datetime(df[mapping['date']], errors='coerce').dropna()
            if not dates.empty:
                if dates.dt.to_period('M').nunique() == 1:
                    timeframe = dates.iloc[0].strftime("%B_%Y")
                else:
                    timeframe = f"{dates.min().strftime('%d%b')}_to_{dates.max().strftime('%d%b_%y')}"
        except Exception:
            timeframe = "Report"

    # 5. Aggregations
    summary_aggs = {'Clean_Qty': 'sum', 'Total Amount': 'sum'}
    if has_cogs:
        summary_aggs['Total COGS'] = 'sum'
    summary = df.groupby('Category').agg(summary_aggs).reset_index()
    if has_cogs:
        summary.columns = ['Category', 'Total Qty', 'Total Amount', 'Total COGS']
        summary['Profit Margin %'] = np.where(
            summary['Total Amount'] > 0,
            ((summary['Total Amount'] - summary['Total COGS']) / summary['Total Amount'] * 100).round(1),
            0.0
        )
    else:
        summary.columns = ['Category', 'Total Qty', 'Total Amount']
    
    t_rev = summary['Total Amount'].sum()
    t_qty = summary['Total Qty'].sum()
    
    gross_profit = float(t_rev - total_cogs) if total_cogs is not None else None
    profit_margin = float((gross_profit / t_rev * 100)) if (gross_profit is not None and t_rev > 0) else None
    
    drilldown = df.groupby(['Category', 'Clean_Cost']).agg({'Clean_Qty': 'sum', 'Total Amount': 'sum'}).reset_index()
    drilldown.columns = ['Category', 'Price', 'Total Qty', 'Total Amount']
    
    top_items = df.groupby('Clean_Name').agg({'Clean_Qty': 'sum', 'Total Amount': 'sum', 'Category': 'first'}).reset_index()
    top_items.columns = ['Product Name', 'Total Qty', 'Total Amount', 'Category']
    top_items = top_items.sort_values('Total Amount', ascending=False)

    prod_aggs = {'Clean_Qty': 'sum', 'Total Amount': 'sum'}
    if has_cogs:
        prod_aggs['Total COGS'] = 'sum'
    product_summary = df.groupby(['Category', 'Clean_Name']).agg(prod_aggs).reset_index()
    if has_cogs:
        product_summary.columns = ['Category', 'Product Name', 'Total Qty', 'Total Amount', 'Total COGS']
        product_summary['Profit Margin %'] = np.where(
            product_summary['Total Amount'] > 0,
            ((product_summary['Total Amount'] - product_summary['Total COGS']) / product_summary['Total Amount'] * 100).round(1),
            0.0
        )
    else:
        product_summary.columns = ['Category', 'Product Name', 'Total Qty', 'Total Amount']
    
    # 6. Basket Metrics & Order Count
    avg_basket_value = 0
    group_cols = [c for c in [mapping.get('order_id'), mapping.get('phone')] if c and c in df.columns]
    
    if group_cols:
        order_groups = df.groupby(group_cols).agg({'Total Amount': 'sum'})
        avg_basket_value = order_groups['Total Amount'].mean()
        total_orders = len(order_groups)
    else:
        total_orders = len(df)
        
    # 7. Unmapped items analysis ("Others")
    unmapped_df = df[df['Category'] == 'Others']
    if not unmapped_df.empty:
        unmapped_items = unmapped_df.groupby('Clean_Name').agg({
            'Clean_Qty': 'sum',
            'Total Amount': 'sum'
        }).reset_index()
        unmapped_items.columns = ['Product Name', 'Total Qty', 'Total Amount']
        unmapped_items = unmapped_items.sort_values('Total Amount', ascending=False)
    else:
        unmapped_items = pd.DataFrame(columns=['Product Name', 'Total Qty', 'Total Amount'])

    unmapped_qty = unmapped_items['Total Qty'].sum() if not unmapped_items.empty else 0
    unmapped_rev = unmapped_items['Total Amount'].sum() if not unmapped_items.empty else 0

    return {
        'drilldown': drilldown,
        'summary': summary,
        'top_items': top_items,
        'product_summary': product_summary,
        'unmapped_items': unmapped_items,
        'unmapped_qty': unmapped_qty,
        'unmapped_rev': unmapped_rev,
        'timeframe': timeframe,
        'avg_basket_value': avg_basket_value,
        'total_qty': t_qty,
        'total_rev': t_rev,
        'total_orders': total_orders,
        'total_cogs': total_cogs,
        'has_cogs': has_cogs,
        'gross_profit': gross_profit,
        'profit_margin': profit_margin
    }

# --- UI & Sidebar Components ---

@st.fragment
def render_sidebar(logo_b64=None):
    with st.sidebar:
        if not logo_b64:
            logo_b64 = load_logo()
        if logo_b64:
            st.markdown(
                f"""
                <div style="display: flex; align-items: center; gap: 12px; margin-bottom: 8px;">
                    <img src="data:image/png;base64,{logo_b64}" width="42" height="42" style="object-fit: contain; border-radius: 8px; filter: drop-shadow(0 2px 4px rgba(0,0,0,0.15)); flex-shrink: 0;">
                    <div style="display: flex; flex-direction: column; justify-content: center;">
                        <h2 style="margin: 0; font-size: 1.3rem; font-weight: 800; letter-spacing: -0.3px; line-height: 1.15;">DEEN Commerce</h2>
                        <div style="margin-top: 3px; font-size: 0.8rem; opacity: 0.8; font-weight: 500; line-height: 1.1;">Daily Sales Report</div>
                    </div>
                </div>
                """,
                unsafe_allow_html=True
            )
        else:
            st.markdown(
                """
                <div style="margin-bottom: 8px;">
                    <h2 style="margin: 0; font-size: 1.3rem; font-weight: 800; letter-spacing: -0.3px; line-height: 1.15;">DEEN Commerce</h2>
                    <div style="margin-top: 3px; font-size: 0.8rem; opacity: 0.8; font-weight: 500; line-height: 1.1;">Daily Sales Report</div>
                </div>
                """,
                unsafe_allow_html=True
            )
        st.divider()

        with st.expander(":material/tune: Category Rule Manager", expanded=False):
            st.caption("Add or adjust keyword matching rules:")
            cat_list = sorted(list(CATEGORY_MAPPING.keys()))
            selected_cat = st.selectbox("Category", cat_list)
            
            curr_keywords = CATEGORY_MAPPING.get(selected_cat, [])
            st.write(f"Current keywords: `{', '.join(curr_keywords)}`")
            
            c_add, c_del = st.tabs(["Add", "Remove"])
            with c_add:
                new_kw = st.text_input("New Keyword", key="new_kw_input")
                if st.button("Add Keyword", type="secondary"):
                    if new_kw and new_kw.lower() not in [k.lower() for k in curr_keywords]:
                        CATEGORY_MAPPING[selected_cat].append(new_kw.lower())
                        with open(CATEGORIES_JSON, 'w', encoding='utf-8') as f:
                            json.dump(CATEGORY_MAPPING, f, indent=4)
                        st.session_state['category_rules_version'] = st.session_state.get('category_rules_version', 0) + 1
                        st.success(f"Added '{new_kw}' to {selected_cat}!")
                        st.rerun()
            with c_del:
                if curr_keywords:
                    del_kw = st.selectbox("Select Keyword to Remove", curr_keywords)
                    if st.button("Delete Keyword", type="secondary"):
                        CATEGORY_MAPPING[selected_cat].remove(del_kw)
                        with open(CATEGORIES_JSON, 'w', encoding='utf-8') as f:
                            json.dump(CATEGORY_MAPPING, f, indent=4)
                        st.session_state['category_rules_version'] = st.session_state.get('category_rules_version', 0) + 1
                        st.success(f"Removed '{del_kw}'!")
                        st.rerun()

            if st.button("Reset Rules to Defaults", icon=":material/restart_alt:"):
                with open(CATEGORIES_JSON, 'w', encoding='utf-8') as f:
                    json.dump(DEFAULT_CATEGORY_MAPPING, f, indent=4)
                st.session_state['category_rules_version'] = st.session_state.get('category_rules_version', 0) + 1
                st.success("Reset all category rules to default!")
                st.rerun()

        with st.expander(":material/cleaning_services: Clean Items & Exclusions", expanded=False):
            st.caption("Lines containing bundle keywords are filtered automatically:")
            st.code("choose any | shoose any | combo | bundle | bundel", language="text")

        with st.expander(":material/bug_report: Feedback & Support", expanded=False):
            f_type = st.selectbox("Issue Type", ["Bug Report", "Feature Request", "Category Keyword Suggestion"])
            f_text = st.text_area("Details", placeholder="Describe the issue or suggested keyword...")
            if st.button("Submit Feedback"):
                if f_text.strip():
                    log_event(f"FEEDBACK_{f_type.upper().replace(' ', '_')}", f_text.strip())
                    st.success("Thank you! Your feedback has been recorded.")
                else:
                    st.warning("Please enter feedback details.")

def render_footer(logo_b64=None):
    """Renders fixed sticky footer."""
    logo_img = f'<img src="data:image/png;base64,{logo_b64}" class="small-logo">' if logo_b64 else '🐾'
    footer_html = f"""
    <div class="sticky-footer">
        <div class="footer-content-inner">
            <span>Powered by</span>
            <div class="brand-wrapper">
                {logo_img}
                <span class="brand-name">Catwise Analytics</span>
            </div>
        </div>
    </div>
    """
    st.markdown(footer_html, unsafe_allow_html=True)

# --- Excel Report Generator ---

def generate_excel_report(sheet_dict):
    """Generates an optimized, beautifully styled multi-sheet Excel workbook with native number formats."""
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        for sheet_name, dframe in sheet_dict.items():
            dframe.to_excel(writer, sheet_name=sheet_name, index=False)
        
        wb = writer.book
        header_fill = PatternFill(start_color='007BFF', end_color='007BFF', fill_type='solid')
        header_font = Font(bold=True, color='FFFFFF')
        total_fill = PatternFill(start_color='ECECEC', end_color='ECECEC', fill_type='solid')
        thin_border = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='thin')
        )
        
        for sheet_name, dframe in sheet_dict.items():
            ws = wb[sheet_name]
            ws.freeze_panes = 'A2'
            
            # Style Header Row
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = thin_border
            
            # Pre-compute column configurations
            col_alignments = [
                Alignment(horizontal='left') if col in ['Category', 'Product Name', 'Product'] 
                else Alignment(horizontal='right')
                for col in dframe.columns
            ]
            is_numeric_col = [any(x in col for x in ['Qty', 'Price', 'Amount', 'COGS']) for col in dframe.columns]
            is_margin_col = ['Margin' in col or '%' in col for col in dframe.columns]
            
            # Auto-adjust column widths
            for i, col_name in enumerate(dframe.columns):
                col_letter = ws.cell(row=1, column=i+1).column_letter
                content_len = dframe[col_name].astype(str).str.len().max() if not dframe.empty else 0
                ws.column_dimensions[col_letter].width = max(int(content_len), len(str(col_name))) + 3
            
            # Fast single-pass row styling
            for row_cells in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=len(dframe.columns)):
                first_cell_str = str(row_cells[0].value).upper()
                is_total = 'TOTAL' in first_cell_str
                for idx, cell in enumerate(row_cells):
                    cell.border = thin_border
                    cell.alignment = col_alignments[idx]
                    if is_numeric_col[idx]:
                        cell.number_format = '#,##0'
                    elif is_margin_col[idx]:
                        cell.number_format = '0.0"%"'
                    if is_total:
                        cell.font = Font(bold=True)
                        cell.fill = total_fill
    buf.seek(0)
    return buf.getvalue()

# --- Main App ---

def main():
    st.set_page_config(page_title="Sales Report", page_icon=":material/monitoring:", layout="wide")
    apply_custom_styles()
    
    logo_b64 = load_logo()
    render_sidebar(logo_b64)

    st.title("Sales Report", anchor=False)
    st.caption("Upload order reports to visualize volume, revenue, profit margins, and category drilldowns.")

    uploaded_file = st.file_uploader("Upload Sales Data (Excel or CSV)", type=['xlsx', 'csv'])

    if uploaded_file is not None:
        try:
            file_bytes = uploaded_file.getvalue()
            file_key = f"{uploaded_file.name}_{len(file_bytes)}"
            
            if st.session_state.get('current_file_key') != file_key:
                st.session_state['current_file_key'] = file_key
                st.session_state['results'] = None

            df = load_data(file_bytes, uploaded_file.name)
            auto_cols = find_columns(df)
            all_cols = list(df.columns)
            
            def get_idx(key):
                col_name = auto_cols.get(key)
                return all_cols.index(col_name) if col_name in all_cols else 0

            with st.container(border=True):
                is_mapped = all(k in auto_cols for k in ['name', 'cost', 'qty'])
                if is_mapped:
                    detected_extra = []
                    if 'cogs' in auto_cols:
                        detected_extra.append("COGS")
                    if 'date' in auto_cols:
                        detected_extra.append("Date")
                    extra_str = f" (including {', '.join(detected_extra)})" if detected_extra else ""
                    st.caption(f":material/check_circle: Mandatory columns{extra_str} were auto-detected. Adjust if needed:")
                else:
                    st.caption(":material/info: Please match the required columns from your uploaded file:")

                st.markdown("**:material/check_circle: Required Fields** *(Mandatory for calculations)*")
                rc1, rc2, rc3 = st.columns(3)
                m_name = rc1.selectbox("Product Name *", all_cols, index=get_idx('name'), help="Column containing the product or item description")
                m_cost = rc2.selectbox("Unit Price *", all_cols, index=get_idx('cost'), help="Column containing unit price or cost")
                m_qty = rc3.selectbox("Quantity *", all_cols, index=get_idx('qty'), help="Column containing quantity sold")

                st.markdown("**:material/add_circle: Optional Dimensions** *(Enables timeframe, basket size & COGS metrics)*")
                oc1, oc2, oc3, oc4 = st.columns(4)
                m_cogs = oc1.selectbox("COGS (Opt)", ["None"] + all_cols, index=get_idx('cogs')+1 if 'cogs' in auto_cols else 0, help="Column containing Cost of Goods Sold (per unit or total)")
                m_date = oc2.selectbox("Date (Opt)", ["None"] + all_cols, index=get_idx('date')+1 if 'date' in auto_cols else 0, help="Order date column for report timeframe")
                m_order = oc3.selectbox("Order ID (Opt)", ["None"] + all_cols, index=get_idx('order_id')+1 if 'order_id' in auto_cols else 0, help="Order number or invoice ID for basket size")
                m_phone = oc4.selectbox("Phone (Opt)", ["None"] + all_cols, index=get_idx('phone')+1 if 'phone' in auto_cols else 0, help="Customer contact for multi-item order grouping")

                mapping = {
                    'name': m_name, 'cost': m_cost, 'qty': m_qty,
                    'cogs': m_cogs if m_cogs != "None" else None,
                    'date': m_date if m_date != "None" else None,
                    'order_id': m_order if m_order != "None" else None,
                    'phone': m_phone if m_phone != "None" else None
                }

                st.write("")
                if st.button("Generate Analytics", icon=":material/analytics:", type="primary"):
                    st.session_state['results'] = process_analytics(df, mapping)

            results = st.session_state.get('results')
            if results is not None:
                has_cogs = results.get('has_cogs') and results.get('total_cogs') is not None
                has_basket = results.get('avg_basket_value', 0) > 0
                
                if has_cogs:
                    cogs_val = results['total_cogs']
                    profit_val = results.get('gross_profit', results['total_rev'] - cogs_val)
                    margin_pct = results.get('profit_margin', (profit_val / results['total_rev'] * 100) if results['total_rev'] > 0 else 0)
                    
                    if has_basket:
                        m1, m2, m3, m4, m5 = st.columns(5)
                    else:
                        m1, m2, m3, m4 = st.columns(4)
                        
                    m1.metric("Total Orders", f"{results['total_orders']:,.0f}", border=True)
                    m2.metric("Units Sold", f"{results['total_qty']:,.0f}", border=True)
                    m3.metric(
                        "Gross Revenue", 
                        f"TK {results['total_rev']:,.0f}", 
                        delta=f"({margin_pct:.1f}%)",
                        delta_arrow="off",
                        delta_color="normal",
                        border=True, 
                        help=f"Gross Revenue: TK {results['total_rev']:,.0f} | Margin: {margin_pct:.1f}%\nTotal COGS: TK {cogs_val:,.0f} | Gross Profit: TK {profit_val:,.0f}"
                    )
                    m4.metric("Total COGS", f"TK {cogs_val:,.0f}", border=True, help="Total Cost of Goods Sold")
                    if has_basket:
                        m5.metric("Basket Size", f"TK {results['avg_basket_value']:,.0f}", border=True, help="Average gross spend per order or customer phone")
                else:
                    if has_basket:
                        m1, m2, m3, m4 = st.columns(4)
                        m1.metric("Total Orders", f"{results['total_orders']:,.0f}", border=True)
                        m2.metric("Units Sold", f"{results['total_qty']:,.0f}", border=True)
                        m3.metric("Gross Revenue", f"TK {results['total_rev']:,.0f}", border=True, help="Total gross revenue from sales")
                        m4.metric("Basket Size", f"TK {results['avg_basket_value']:,.0f}", border=True, help="Average gross spend per order or customer phone")
                    else:
                        m1, m2, m3 = st.columns(3)
                        m1.metric("Total Orders", f"{results['total_orders']:,.0f}", border=True)
                        m2.metric("Units Sold", f"{results['total_qty']:,.0f}", border=True)
                        m3.metric("Gross Revenue", f"TK {results['total_rev']:,.0f}", border=True, help="Total gross revenue from sales")

                st.divider()
                
                # Visuals
                v1, v2 = st.columns(2)
                
                summ_sorted = results['summary'].sort_values('Total Amount', ascending=False)
                summ_qty_sorted = results['summary'].sort_values('Total Qty', ascending=False)
                cat_color_map = get_category_color_map(results['summary']['Category'].unique())
                
                with v1:
                    with st.container(border=True):
                        st.subheader(":material/pie_chart: Revenue by Category", anchor=False)
                        fig_pie = px.pie(
                            summ_sorted, 
                            values='Total Amount', 
                            names='Category', 
                            color='Category', 
                            hole=0.5, 
                            color_discrete_map=cat_color_map
                        )
                        fig_pie.update_traces(
                            marker=dict(
                                line=dict(color='#000038', width=3)
                            )
                        )
                        fig_pie.update_layout(
                            paper_bgcolor='rgba(0,0,0,0)',
                            plot_bgcolor='rgba(0,0,0,0)',
                            margin=dict(t=10, b=10, l=10, r=10),
                            legend=dict(orientation="v", yanchor="middle", y=0.5, xanchor="left", x=1.0)
                        )
                        st.plotly_chart(fig_pie, width="stretch")
                
                with v2:
                    with st.container(border=True):
                        st.subheader(":material/bar_chart: Volume by Category", anchor=False)
                        fig_bar = px.bar(
                            summ_qty_sorted, 
                            x='Category', 
                            y='Total Qty', 
                            color='Category', 
                            color_discrete_map=cat_color_map
                        )
                        fig_bar.update_layout(
                            xaxis={'categoryorder': 'total descending'},
                            paper_bgcolor='rgba(0,0,0,0)',
                            plot_bgcolor='rgba(0,0,0,0)',
                            margin=dict(t=10, b=10, l=10, r=10),
                            showlegend=False
                        )
                        st.plotly_chart(fig_bar, width="stretch")
                
                # Data Tables
                t1, t2, t3 = st.tabs([":material/table_chart: Category Summary", ":material/payments: Price-wise Category", ":material/inventory_2: Category-wise Product"])
                
                df_breakdown = results['summary'].sort_values('Category', ascending=True).copy()
                total_qty = df_breakdown['Total Qty'].sum()
                total_amount = df_breakdown['Total Amount'].sum()
                
                if results.get('has_cogs'):
                    total_cogs_sum = df_breakdown['Total COGS'].sum()
                    total_margin_pct = round(((total_amount - total_cogs_sum) / total_amount * 100), 1) if total_amount > 0 else 0.0
                    df_breakdown.loc[len(df_breakdown) + 1] = ['TOTAL SALES (Summary)', total_qty, total_amount, total_cogs_sum, total_margin_pct]
                else:
                    df_breakdown.loc[len(df_breakdown) + 1] = ['TOTAL SALES (Summary)', total_qty, total_amount]
                
                df_breakdown.index = range(1, len(df_breakdown) + 1)
                
                df_drill = results['drilldown'].sort_values(['Category', 'Price'], ascending=[True, False]).copy()
                df_drill.columns = ['Category', 'Price', 'Qty', 'Total Amount']
                df_drill.loc[len(df_drill) + 1] = ['TOTAL (All Categories)', None, df_drill['Qty'].sum(), df_drill['Total Amount'].sum()]
                df_drill.index = range(1, len(df_drill) + 1)

                df_product = results['product_summary'].sort_values(['Category', 'Total Amount'], ascending=[True, False]).copy()
                if results.get('has_cogs'):
                    total_prod_qty = df_product['Total Qty'].sum()
                    total_prod_amt = df_product['Total Amount'].sum()
                    total_prod_cogs = df_product['Total COGS'].sum()
                    total_prod_margin = round(((total_prod_amt - total_prod_cogs) / total_prod_amt * 100), 1) if total_prod_amt > 0 else 0.0
                    df_product.loc[len(df_product) + 1] = ['TOTAL (All Products)', None, total_prod_qty, total_prod_amt, total_prod_cogs, total_prod_margin]
                else:
                    df_product.columns = ['Category', 'Product Name', 'Total Qty', 'Total Amount']
                    df_product.loc[len(df_product) + 1] = ['TOTAL (All Products)', None, df_product['Total Qty'].sum(), df_product['Total Amount'].sum()]
                df_product.index = range(1, len(df_product) + 1)

                with t1: 
                    t1_cols = ['Category', 'Total Qty', 'Total Amount']
                    t1_cfg = {
                        "Category": st.column_config.TextColumn("Product Category"),
                        "Total Qty": st.column_config.NumberColumn("Total Units", format="%,.0f pcs"),
                        "Total Amount": st.column_config.NumberColumn("Total Revenue", format="TK %,.0f"),
                    }
                    if results.get('has_cogs'):
                        t1_cols.extend(['Total COGS', 'Profit Margin %'])
                        t1_cfg["Total COGS"] = st.column_config.NumberColumn("Total COGS", format="TK %,.0f")
                        t1_cfg["Profit Margin %"] = st.column_config.NumberColumn("Profit Margin %", format="%.1f%%")
                    st.dataframe(
                        df_breakdown[t1_cols],
                        column_config=t1_cfg,
                        width="stretch",
                        hide_index=True
                    )
                
                with t2: 
                    st.dataframe(
                        df_drill[['Category', 'Price', 'Qty', 'Total Amount']],
                        column_config={
                            "Category": st.column_config.TextColumn("Product Category"),
                            "Price": st.column_config.NumberColumn("Unit Price", format="TK %,.0f"),
                            "Qty": st.column_config.NumberColumn("Units Sold", format="%,.0f pcs"),
                            "Total Amount": st.column_config.NumberColumn("Total Revenue", format="TK %,.0f"),
                        },
                        width="stretch",
                        hide_index=True
                    )

                with t3:
                    c_search, c_filter = st.columns([3, 2])
                    with c_search:
                        p_search = st.text_input(
                            "Search Product", 
                            placeholder="Search products by title or keyword...", 
                            label_visibility="collapsed",
                            key="tab3_product_search"
                        )
                    with c_filter:
                        cat_choices = sorted([c for c in results['product_summary']['Category'].unique() if c])
                        selected_cats = st.multiselect(
                            "Filter Category", 
                            options=cat_choices, 
                            default=[], 
                            placeholder="All Categories",
                            label_visibility="collapsed",
                            key="tab3_category_filter"
                        )

                    filtered_prod = results['product_summary'].copy()
                    if selected_cats:
                        filtered_prod = filtered_prod[filtered_prod['Category'].isin(selected_cats)]
                    if p_search.strip():
                        filtered_prod = filtered_prod[
                            filtered_prod['Product Name'].str.contains(p_search.strip(), case=False, na=False)
                        ]

                    df_filtered_view = filtered_prod.sort_values(['Category', 'Total Amount'], ascending=[True, False]).copy()
                    t3_cols = ['Category', 'Product Name', 'Total Qty', 'Total Amount']
                    t3_cfg = {
                        "Category": st.column_config.TextColumn("Category"),
                        "Product Name": st.column_config.TextColumn("Product Name"),
                        "Total Qty": st.column_config.NumberColumn("Total Units", format="%,.0f pcs"),
                        "Total Amount": st.column_config.NumberColumn("Total Revenue", format="TK %,.0f"),
                    }
                    if results.get('has_cogs'):
                        t3_cols.extend(['Total COGS', 'Profit Margin %'])
                        t3_cfg["Total COGS"] = st.column_config.NumberColumn("Total COGS", format="TK %,.0f")
                        t3_cfg["Profit Margin %"] = st.column_config.NumberColumn("Profit Margin %", format="%.1f%%")
                        f_total_qty = df_filtered_view['Total Qty'].sum()
                        f_total_amt = df_filtered_view['Total Amount'].sum()
                        f_total_cogs = df_filtered_view['Total COGS'].sum()
                        f_total_margin = round(((f_total_amt - f_total_cogs) / f_total_amt * 100), 1) if f_total_amt > 0 else 0.0
                        df_filtered_view.loc[len(df_filtered_view) + 1] = [
                            f"TOTAL ({len(df_filtered_view)} items)", None, f_total_qty, f_total_amt, f_total_cogs, f_total_margin
                        ]
                    else:
                        f_total_qty = df_filtered_view['Total Qty'].sum()
                        f_total_amt = df_filtered_view['Total Amount'].sum()
                        df_filtered_view.loc[len(df_filtered_view) + 1] = [
                            f"TOTAL ({len(df_filtered_view)} items)", None, f_total_qty, f_total_amt
                        ]
                    df_filtered_view.index = range(1, len(df_filtered_view) + 1)

                    st.dataframe(
                        df_filtered_view[t3_cols],
                        column_config=t3_cfg,
                        width="stretch",
                        hide_index=True
                    )

                # Unmapped Items Inspector ("Others" Review)
                unmapped_df = results.get('unmapped_items')
                unmapped_count = len(unmapped_df) if unmapped_df is not None else 0
                with st.expander(f"Unmapped Products Inspector ('Others' — {unmapped_count} unique items)", icon=":material/warning:", expanded=False):
                    if unmapped_df is None or unmapped_df.empty:
                        st.success("All products were successfully categorized! Zero unmapped items.", icon=":material/check_circle:")
                    else:
                        u1, u2 = st.columns(2)
                        u1.metric("Unmapped Units", f"{results['unmapped_qty']:,.0f}", border=True)
                        u2.metric("Unmapped Revenue", f"TK {results['unmapped_rev']:,.0f}", border=True)
                        st.caption("Review these products to identify missing keywords for your category rules in `config/categories.json`:")
                        st.dataframe(
                            unmapped_df,
                            column_config={
                                "Product Name": st.column_config.TextColumn("Product Description"),
                                "Total Qty": st.column_config.NumberColumn("Quantity", format="%,.0f pcs"),
                                "Total Amount": st.column_config.NumberColumn("Revenue Impact", format="TK %,.0f"),
                            },
                            width="stretch",
                            hide_index=True
                        )

                # Export Multi-sheet Workbook
                sheet_dict = {
                    'Category Summary': df_breakdown,
                    'Price-wise Category': df_drill,
                    'Category-wise Product': df_product
                }

                excel_data = generate_excel_report(sheet_dict)
                fname = f"Sales_Report_{results['timeframe']}.xlsx"
                st.download_button("Download Report", data=excel_data, file_name=fname, icon=":material/download:")
                
        except Exception as e:
            st.error(f"Processing Error: {e}")
            log_event("CRASH", str(e))
    else:
        st.session_state['current_file_key'] = None
        st.session_state['results'] = None

    render_footer(logo_b64)

if __name__ == "__main__":
    main()
