"""
Catwise Analytics - Enhanced Version with Dynamic Column & Category Detection
All UI/UX improvements implemented while preserving and enhancing pie/donut and bar charts.
"""

from pandas.errors import ClosedFileError
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os
import json
import base64
import re
from datetime import datetime, timedelta
from io import BytesIO
from collections import Counter
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# --- Configuration & Constants ---
FEEDBACK_DIR = "feedback"
PRESETS_DIR = "presets"
os.makedirs(FEEDBACK_DIR, exist_ok=True)
os.makedirs(PRESETS_DIR, exist_ok=True)

LOGO_PNG = "assets/deen_logo.png"
DEEN_BRAND_COLOR = "#007BFF"
CATEGORY_COLORS = px.colors.qualitative.Pastel + px.colors.qualitative.Set3

# --- Dynamic Category Patterns (Replaces hardcoded CATEGORY_MAPPING) ---
CATEGORY_PATTERNS = {
    'FS T-Shirt': [
        r'full[- ]?sleeve.*t[- ]?shirt', r'fs.*t[- ]?shirt', r't[- ]?shirt.*full[- ]?sleeve',
        r'full[- ]?sleeve.*tee', r'long[- ]?sleeve.*t[- ]?shirt', r'l/s.*t[- ]?shirt'
    ],
    'HS T-Shirt': [
        r'half[- ]?sleeve.*t[- ]?shirt', r'hs.*t[- ]?shirt', r't[- ]?shirt.*half[- ]?sleeve',
        r'short[- ]?sleeve.*t[- ]?shirt', r's/s.*t[- ]?shirt'
    ],
    'T-Shirt': [
        r't[- ]?shirt', r'tee', r't[- ]?shirt$', r'^tee$',
        r'plain.*t[- ]?shirt', r'basic.*t[- ]?shirt'
    ],
    'FS Shirt': [
        r'full[- ]?sleeve.*shirt', r'fs.*shirt', r'shirt.*full[- ]?sleeve',
        r'long[- ]?sleeve.*shirt', r'l/s.*shirt'
    ],
    'HS Shirt': [
        r'half[- ]?sleeve.*shirt', r'hs.*shirt', r'shirt.*half[- ]?sleeve',
        r'short[- ]?sleeve.*shirt', r's/s.*shirt'
    ],
    'Formal Shirt': [
        r'formal.*shirt', r'executive.*shirt', r'dress.*shirt',
        r'business.*shirt', r'office.*shirt'
    ],
    'Cuban Shirt': [r'cuban.*shirt', r'cuban[- ]?collar.*shirt', r'cuban collar'],
    'Denim Shirt': [r'denim.*shirt', r'jeans.*shirt'],
    'Flannel Shirt': [r'flannel.*shirt', r'flannel'],
    'Polo Shirt': [r'polo.*shirt', r'polo', r'polo t[- ]?shirt'],
    'Contrast Shirt': [r'contrast.*shirt', r'contrast'],
    'Kaftan Shirt': [r'kaftan.*shirt', r'kaftan'],
    'Boxer': [r'boxer', r'boxers', r'boxer short'],
    'Jeans': [r'jeans', r'jean', r'denim.*pant', r'denim.*trouser'],
    'Cargo Trousers': [r'cargo.*trouser', r'cargo.*pant', r'cargo', r'trousers.*cargo'],
    'Trousers': [r'trouser', r'trousers', r'pant', r'pants', r'jogger', r'joggers',
                r'track.*pant', r'track.*pants', r'trcargo'],
    'Twill Chino': [r'twill.*chino', r'chino.*twill', r'twill'],
    'Sweatshirt': [r'sweatshirt', r'hoodie', r'pullover', r'hooded.*sweatshirt'],
    'Sweater': [r'sweater', r'cardigan', r'knitwear', r'knitted.*sweater'],
    'Jacket': [r'jacket', r'outerwear', r'coat', r'blazer', r'bomber.*jacket'],
    'Tank Top': [r'tank.*top', r'tanktop', r'tank', r'vest', r'sleeveless.*top'],
    'Panjabi': [r'panjabi', r'punjabi', r'kurta', r'sherwani'],
    'Mask': [r'mask', r'face.*mask', r'protective.*mask'],
    'Water Bottle': [r'water.*bottle', r'bottle', r'hydration.*bottle'],
    'Wallet': [r'wallet', r'leather.*wallet', r'card.*holder'],
    'Passport Holder': [r'passport.*holder', r'passport cover'],
    'Cap': [r'cap', r'baseball.*cap', r'hat', r'beanie'],
    'Belt': [r'belt', r'leather.*belt'],
    'Leather Bag': [r'leather.*bag', r'bag', r'backpack', r'sling.*bag', r'crossbody.*bag'],
    'Drop Shoulder': [r'drop.*shoulder', r'shoulder.*drop', r'oversized.*shoulder'],
    'Turtleneck': [r'turtleneck', r'mock.*neck', r'roll.*neck'],
    'Active Wear': [r'active.*wear', r'sports.*wear', r'gym.*wear', r'athleisure'],
    'Jersy': [r'jersy', r'jersey', r'sports.*jersey'],
}

# --- Priority Column Mapping (Enhanced) ---
PRIORITY_COLUMN_MAPPING = {
    'name': {
        'priority': 1,
        'exact_aliases': ['item name', 'product name', 'product', 'item', 'title', 'description', 'name', 'product_title'],
        'partial_aliases': ['name', 'product', 'item', 'title'],
        'keywords': ['name', 'product', 'item', 'title', 'description'],
        'data_type': 'string'
    },
    'cost': {
        'priority': 2,
        'exact_aliases': ['item cost', 'price', 'unit price', 'cost', 'rate', 'mrp', 'selling price', 'unit_cost', 'sale_price'],
        'partial_aliases': ['price', 'cost', 'rate', 'mrp', 'selling'],
        'keywords': ['price', 'cost', 'rate', 'value', 'amount', 'charge'],
        'data_type': 'numeric'
    },
    'qty': {
        'priority': 3,
        'exact_aliases': ['quantity', 'qty', 'units', 'sold', 'count', 'total quantity', 'unit', 'item_qty'],
        'partial_aliases': ['qty', 'quantity', 'unit', 'count', 'sold'],
        'keywords': ['quantity', 'qty', 'unit', 'count', 'number', 'amount', 'pieces'],
        'data_type': 'numeric'
    },
    'date': {
        'priority': 4,
        'exact_aliases': ['date', 'order date', 'month', 'time', 'created at', 'created_date', 'order_date', 'sale_date'],
        'partial_aliases': ['date', 'order', 'created', 'sale'],
        'keywords': ['date', 'time', 'day', 'month', 'year', 'created', 'order'],
        'data_type': 'datetime'
    },
    'order_id': {
        'priority': 5,
        'exact_aliases': ['order id', 'order #', 'invoice number', 'invoice #', 'order number', 'transaction id', 'id', 'order_no', 'invoice_no'],
        'partial_aliases': ['order', 'invoice', 'transaction', 'id'],
        'keywords': ['order', 'invoice', 'transaction', 'id', 'number', 'ref', 'reference'],
        'data_type': 'string'
    },
    'phone': {
        'priority': 6,
        'exact_aliases': ['phone', 'contact', 'mobile', 'cell', 'phone number', 'customer phone', 'customer_mobile'],
        'partial_aliases': ['phone', 'mobile', 'contact', 'cell'],
        'keywords': ['phone', 'mobile', 'contact', 'cell', 'customer'],
        'data_type': 'string'
    }
}

COLUMN_DISPLAY_ORDER = ['name', 'cost', 'qty', 'date', 'order_id', 'phone']


def load_logo():
    if os.path.exists(LOGO_PNG):
        try:
            with open(LOGO_PNG, "rb") as f:
                return base64.b64encode(f.read()).decode()
        except:
            pass
    return ""


def log_event(event_type, details):
    log_file = os.path.join(FEEDBACK_DIR, "system_logs.json")
    entry = {
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "type": event_type,
        "details": details
    }
    try:
        logs = []
        if os.path.exists(log_file):
            with open(log_file, "r") as f:
                logs = json.load(f)
        logs.append(entry)
        with open(log_file, "w") as f:
            json.dump(logs[-100:], f, indent=4)
    except:
        pass


def categorize_product_dynamically(name):
    if pd.isna(name) or str(name).strip() == '':
        return 'Others'
    name_str = str(name).lower()
    for category, patterns in CATEGORY_PATTERNS.items():
        for pattern in patterns:
            if re.search(pattern, name_str, re.IGNORECASE):
                return category
    fs_keywords = ['full sleeve', 'full-sleeve', 'fs', 'l/s', 'long sleeve', 'long-sleeve']
    hs_keywords = ['half sleeve', 'half-sleeve', 'hs', 's/s', 'short sleeve', 'short-sleeve']
    is_fs = any(kw in name_str for kw in fs_keywords)
    is_hs = any(kw in name_str for kw in hs_keywords)
    if any(kw in name_str for kw in ['t-shirt', 't shirt', 'tee', 't shirt']):
        if is_fs:
            return 'FS T-Shirt'
        elif is_hs:
            return 'HS T-Shirt'
        return 'T-Shirt'
    if 'shirt' in name_str:
        if is_fs:
            return 'FS Shirt'
        elif is_hs:
            return 'HS Shirt'
        return 'Shirt'
    return 'Others'


def find_columns(df):
    if df is None or df.empty:
        return {}
    found = {}
    used_indices = set()
    all_cols = list(df.columns)
    lower_cols = [str(c).strip().lower() for c in all_cols]
    sorted_keys = sorted(
        PRIORITY_COLUMN_MAPPING.keys(),
        key=lambda k: PRIORITY_COLUMN_MAPPING[k]['priority']
    )
    for key in sorted_keys:
        config = PRIORITY_COLUMN_MAPPING[key]
        for alias in config['exact_aliases']:
            alias_lower = alias.lower()
            if alias_lower in lower_cols:
                idx = lower_cols.index(alias_lower)
                if idx not in used_indices:
                    found[key] = all_cols[idx]
                    used_indices.add(idx)
                    break
    for key in sorted_keys:
        if key in found:
            continue
        config = PRIORITY_COLUMN_MAPPING[key]
        for i, col_lower in enumerate(lower_cols):
            if i in used_indices:
                continue
            for alias in config['partial_aliases']:
                alias_lower = alias.lower()
                if alias_lower in col_lower:
                    found[key] = all_cols[i]
                    used_indices.add(i)
                    break
            if key in found:
                break
    for key in sorted_keys:
        if key in found:
            continue
        config = PRIORITY_COLUMN_MAPPING[key]
        for i, col_lower in enumerate(lower_cols):
            if i in used_indices:
                continue
            for keyword in config['keywords']:
                if keyword.lower() in col_lower:
                    found[key] = all_cols[i]
                    used_indices.add(i)
                    break
            if key in found:
                break
    if 'cost' not in found:
        for i, col in enumerate(all_cols):
            if i in used_indices:
                continue
            if df[col].dtype in ['int64', 'float64', 'int32', 'float32']:
                sample_values = df[col].dropna().head(10)
                if len(sample_values) > 0 and all(v > 0 for v in sample_values):
                    found['cost'] = col
                    used_indices.add(i)
                    break
    if 'qty' not in found:
        for i, col in enumerate(all_cols):
            if i in used_indices:
                continue
            if df[col].dtype in ['int64', 'float64', 'int32', 'float32']:
                sample_values = df[col].dropna().head(10)
                if len(sample_values) > 0:
                    if all(v >= 0 for v in sample_values):
                        found['qty'] = col
                        used_indices.add(i)
                        break
    return found


def get_category_color(category):
    hash_val = hash(category) % len(CATEGORY_COLORS)
    return CATEGORY_COLORS[hash_val]


def clean_numeric(val):
    if pd.isna(val):
        return 0
    if isinstance(val, (int, float)):
        return val
    clean_val = ''.join(c for c in str(val) if c.isdigit() or c == '.')
    try:
        return float(clean_val) if clean_val else 0
    except ValueError:
        return 0


def process_analytics(df, mapping):
    df = df.copy()
    df['Clean_Name'] = df[mapping['name']].fillna('Unknown').astype(str)
    exclude_keywords = 'choose any|shoose any|combo|bundle|bundel'
    df = df[~df['Clean_Name'].str.contains(exclude_keywords, case=False, na=False)]
    df['Clean_Cost'] = df[mapping['cost']].apply(clean_numeric)
    df['Clean_Qty'] = df[mapping['qty']].apply(clean_numeric)
    df.loc[df['Clean_Qty'] < 0, 'Clean_Qty'] = 0
    df['Total Amount'] = df['Clean_Cost'] * df['Clean_Qty']
    df['Category'] = df['Clean_Name'].apply(categorize_product_dynamically)
    timeframe = ""
    if mapping.get('date') and mapping['date'] in df.columns:
        try:
            dates = pd.to_datetime(df[mapping['date']], errors='coerce').dropna()
            if not dates.empty:
                if dates.dt.to_period('M').nunique() == 1:
                    timeframe = dates.iloc[0].strftime("%B_%Y")
                else:
                    timeframe = f"{dates.min().strftime('%d%b')}_to_{dates.max().strftime('%d%b_%y')}"
        except:
            timeframe = "Report"
    timeseries_data = None
    if mapping.get('date') and mapping['date'] in df.columns:
        try:
            df['Parsed_Date'] = pd.to_datetime(df[mapping['date']], errors='coerce')
            daily_sales = df.groupby('Parsed_Date').agg({'Total Amount': 'sum', 'Clean_Qty': 'sum'}).reset_index()
            daily_sales = daily_sales.sort_values('Parsed_Date')
            if len(daily_sales) > 1:
                timeseries_data = daily_sales
        except:
            pass
    summary = df.groupby('Category').agg({'Clean_Qty': 'sum', 'Total Amount': 'sum'}).reset_index()
    summary.columns = ['Category', 'Total Qty', 'Total Amount']
    t_rev = summary['Total Amount'].sum()
    t_qty = summary['Total Qty'].sum()
    drilldown = df.groupby(['Category', 'Clean_Cost']).agg({'Clean_Qty': 'sum', 'Total Amount': 'sum'}).reset_index()
    drilldown.columns = ['Category', 'Price', 'Total Qty', 'Total Amount']
    top_items = df.groupby('Clean_Name').agg({'Clean_Qty': 'sum', 'Total Amount': 'sum', 'Category': 'first'}).reset_index()
    top_items.columns = ['Product Name', 'Total Qty', 'Total Amount', 'Category']
    top_items = top_items.sort_values('Total Amount', ascending=False)
    avg_basket_value = 0
    total_orders = 0
    group_cols = [c for c in [mapping.get('order_id'), mapping.get('phone')] if c and c in df.columns]
    if group_cols:
        order_groups = df.groupby(group_cols).agg({'Total Amount': 'sum'})
        avg_basket_value = order_groups['Total Amount'].mean()
        total_orders = len(order_groups)
    all_categories = sorted(df['Category'].unique().tolist())
    return {
        'drilldown': drilldown,
        'summary': summary,
        'top_items': top_items,
        'timeseries': timeseries_data,
        'timeframe': timeframe,
        'avg_basket_value': avg_basket_value,
        'total_qty': t_qty,
        'total_rev': t_rev,
        'total_orders': total_orders,
        'all_categories': all_categories
    }


def save_mapping_preset(preset_name, mapping):
    preset_file = os.path.join(PRESETS_DIR, f"{preset_name}.json")
    preset_data = {'name': preset_name, 'mapping': mapping, 'created_at': datetime.now().isoformat(), 'updated_at': datetime.now().isoformat()}
    try:
        with open(preset_file, 'w') as f:
            json.dump(preset_data, f, indent=4)
        return True
    except:
        return False


def load_mapping_presets():
    presets = []
    if os.path.exists(PRESETS_DIR):
        for filename in os.listdir(PRESETS_DIR):
            if filename.endswith('.json'):
                try:
                    with open(os.path.join(PRESETS_DIR, filename), 'r') as f:
                        preset_data = json.load(f)
                        presets.append(preset_data)
                except:
                    pass
    return presets


def delete_mapping_preset(preset_name):
    preset_file = os.path.join(PRESETS_DIR, f"{preset_name}.json")
    if os.path.exists(preset_file):
        try:
            os.remove(preset_file)
            return True
        except:
            return False
    return False


def apply_custom_styles():
    st.markdown("""
        <style>
        .main { background-color: transparent; }
        div[data-testid="stMetric"] { background-color: var(--background-secondary, #ffffff); padding: 24px; border-radius: 16px; box-shadow: 0 4px 20px rgba(0,0,0,0.06); border: 1px solid var(--secondary-background-color, #ececec); transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1); overflow: visible !important; }
        div[data-testid="stMetric"]:hover { transform: translateY(-4px); box-shadow: 0 12px 24px rgba(0,0,0,0.12); border-color: #007bff; }
        [data-testid="stMetricLabel"] { font-weight: 600 !important; font-size: 1.1rem !important; margin-bottom: 8px !important; color: var(--text-color, #1a1a1b) !important; }
        [data-testid="stMetricValue"], [data-testid="stMetricValue"] > div { font-weight: 800 !important; font-size: 1.6rem !important; color: var(--text-color, #1a1a1b) !important; white-space: nowrap !important; overflow: visible !important; text-overflow: clip !important; min-width: fit-content !important; }
        [data-theme="dark"] [data-testid="stMetricLabel"], [data-theme="dark"] [data-testid="stMetricValue"], [data-theme="dark"] [data-testid="stMetricValue"] > div { color: #ffffff !important; }
        [data-theme="dark"] div[data-testid="stMetric"] { background-color: #1e1e1e !important; box-shadow: 0 4px 20px rgba(0,0,0,0.4); border: 1px solid #333; }
        .stButton>button { width: 100%; border-radius: 12px; height: 3.5em; background: linear-gradient(135deg, #007bff 0%, #0056b3 100%); color: white; font-weight: 600; border: none; transition: all 0.3s ease; }
        .stButton>button:hover { transform: scale(1.02); background: linear-gradient(135deg, #0056b3 0%, #004085 100%); box-shadow: 0 6px 12px rgba(0,0,0,0.15); }
        div[data-testid="stExpander"] { border: 1px solid var(--secondary-background-color, #eee); box-shadow: 0 4px 12px rgba(0,0,0,0.05); background: var(--background-secondary, white); border-radius: 14px; margin-bottom: 1.5rem; padding: 5px; transition: border-color 0.3s ease; }
        [data-theme="dark"] div[data-testid="stExpander"] { background-color: #1a1a1a !important; border-color: #333 !important; box-shadow: 0 4px 15px rgba(0,0,0,0.3); }
        div[data-testid="stSelectbox"] > div, div[data-testid="stTextInput"] > div, div[data-testid="stTextArea"] > div { background-color: var(--secondary-background-color) !important; color: var(--text-color) !important; border-radius: 12px !important; border: 1px solid rgba(128, 128, 128, 0.2) !important; }
        div[data-testid="stTabs"] button { color: var(--text-color) !important; opacity: 0.7; border-radius: 8px 8px 0 0 !important; padding: 10px 20px !important; }
        div[data-testid="stTabs"] button[aria-selected="true"] { color: var(--primary-color) !important; opacity: 1; border-bottom: 3px solid var(--primary-color) !important; background-color: rgba(0, 123, 255, 0.1) !important; }
        hr { border-color: rgba(128, 128, 128, 0.2) !important; margin: 1rem 0 !important; }
        .sticky-footer { position: fixed !important; left: 0 !important; bottom: 0 !important; right: 0 !important; width: 100% !important; background-color: var(--background-color) !important; color: var(--text-color) !important; text-align: center; padding: 20px 0; border-top: 1px solid rgba(128, 128, 128, 0.2); z-index: 999999; font-size: 0.9rem; backdrop-filter: blur(10px); box-shadow: 0 -5px 25px rgba(0,0,0,0.1); }
        [data-theme="dark"] .sticky-footer { background-color: #0e1117 !important; box-shadow: 0 -5px 25px rgba(0,0,0,0.4); }
        [data-theme="dark"] .sticky-footer, [data-theme="dark"] .sticky-footer span, [data-theme="dark"] .brand-name { color: #f8f9fa !important; }
        .footer-content-inner { display: flex; align-items: center; justify-content: center; gap: 6px; flex-wrap: wrap; }
        .brand-wrapper { display: flex; align-items: center; gap: 6px; }
        .small-logo { height: 24px; width: auto; filter: grayscale(0%); }
        .brand-name { font-weight: 700; color: var(--primary-color); }
        .block-container { padding-bottom: 150px !important; padding-top: 3rem !important; }
        .stSpinner > div { border-top-color: #007bff !important; }
        [data-testid="stFileUploader"] { border: 2px dashed #007bff !important; border-radius: 12px !important; background-color: var(--secondary-background-color) !important; }
        @media (max-width: 768px) { .block-container { padding-left: 1rem !important; padding-right: 1rem !important; padding-bottom: 220px !important; } .footer-content-inner { flex-direction: column; gap: 8px; } [data-testid="stMetricValue"] { font-size: 1.3rem !important; } .sticky-footer { padding: 15px 5px !important; } .stButton>button { height: 3em !important; font-size: 0.9rem !important; } }
        </style>
    """, unsafe_allow_html=True)


def render_sidebar():
    with st.sidebar:
        st.markdown("### 📚 **Quick Tutorial**")
        st.markdown("""**Step 1:** Upload your sales data (Excel or CSV)  **Step 2:** Review auto-detected columns or adjust manually  **Step 3:** Click Generate Analytics  **💡 Pro Tips:** Use Auto-Detect All for instant column mapping. Save presets for repeated file formats. Use filters to explore specific data segments.""")
        st.divider()
        st.markdown("### 💾 **Column Mapping Presets**")
        presets = load_mapping_presets()
        preset_names = [p['name'] for p in presets]
        if preset_names:
            selected_preset = st.selectbox("Load Preset:", ["None"] + preset_names)
            if selected_preset != "None":
                preset = next(p for p in presets if p['name'] == selected_preset)
                if st.button(f"Load {selected_preset}"):
                    st.session_state.loaded_preset = preset['mapping']
                    st.session_state.preset_loaded = True
                    st.success(f"Loaded preset: {selected_preset}")
            st.divider()
            delete_preset = st.selectbox("Delete Preset:", ["None"] + preset_names)
            if delete_preset != "None" and st.button(f"Delete {delete_preset}"):
                if delete_mapping_preset(delete_preset):
                    st.success(f"Deleted preset: {delete_preset}")
                    st.rerun()
        else:
            st.info("No presets saved yet.")
        st.divider()
        if st.checkbox("Save Current Mapping as Preset"):
            preset_name = st.text_input("Preset Name:", placeholder="e.g., My Shopify Data")
            if preset_name and 'current_mapping' in st.session_state:
                if st.button(f"Save {preset_name}"):
                    if save_mapping_preset(preset_name, st.session_state.current_mapping):
                        st.success(f"Saved preset: {preset_name}")
                        st.rerun()
        st.divider()
        st.markdown("### 📝 **Feedback & Debug**")
        comment = st.text_area("Report Issues:", placeholder="Category Polo is incorrect...")
        if st.button("Submit Report"):
            feedback_file = os.path.join(FEEDBACK_DIR, "user_feedback.json")
            entry = {"timestamp": datetime.now().isoformat(), "comment": comment}
            try:
                data = []
                if os.path.exists(feedback_file):
                    with open(feedback_file, "r") as f:
                        data = json.load(f)
                data.append(entry)
                with open(feedback_file, "w") as f:
                    json.dump(data, f, indent=4)
                st.success("Feedback saved! Thank you!")
            except:
                st.error("Failed to save feedback.")
        if st.checkbox("View System Logs"):
            log_path = os.path.join(FEEDBACK_DIR, "system_logs.json")
            if os.path.exists(log_path):
                with open(log_path, "r") as f:
                    logs = json.load(f)
                    st.json(logs[-10:])
            else:
                st.info("No logs available.")


def render_footer(logo_b64):
    footer_html = f"""
    <div style="height: 120px;"></div>
    <div class="sticky-footer">
        <div class="footer-content-inner">
            <span>© {datetime.now().year} Sajid Islam. All rights reserved. | </span>
            <div class="brand-wrapper">
                <span style="opacity: 0.85;">Powered by</span>
                <img src="data:image/png;base64,{logo_b64}" class="small-logo">
                <span class="brand-name">DEEN Commerce Ltd.</span>
            </div>
        </div>
    </div>
    """
    st.markdown(footer_html, unsafe_allow_html=True)


def render_empty_state():
    st.markdown("""
        <div style='text-align: center; padding: 40px 20px;'>
            <h1 style='color: #007bff;'>Welcome to Catwise Analytics!</h1>
            <p style='font-size: 1.2rem; color: #666; margin: 20px 0;'>Upload your sales data to get started</p>
            <div style='background: linear-gradient(135deg, #e3f2fd 0%, #bbdefb 100%); padding: 30px; border-radius: 16px; margin: 30px 0;'>
                <h3 style='color: #007bff;'>What You Can Do</h3>
                <ul style='text-align: left; display: inline-block; color: #333;'>
                    <li>View revenue and sales metrics by category</li>
                    <li>Analyze trends with time-series charts</li>
                    <li>Filter data by category, date, and price</li>
                    <li>Save and load column mapping presets</li>
                    <li>Export beautiful Excel reports</li>
                </ul>
            </div>
            <p style='color: #999;'><strong>Supported Formats:</strong> Excel (.xlsx) and CSV files</p>
        </div>
    """, unsafe_allow_html=True)


def render_filters(df, results):
    if df is None or results is None:
        return df
    st.markdown("### Interactive Filters")
    col1, col2, col3 = st.columns(3)
    with col1:
        if results.get('all_categories'):
            selected_categories = st.multiselect("Filter by Category", options=results['all_categories'], default=results['all_categories'], help="Select categories to include in analysis")
        else:
            selected_categories = []
    with col2:
        date_col = None
        if 'date' in st.session_state.current_mapping and st.session_state.current_mapping['date']:
            date_col = st.session_state.current_mapping['date']
            if date_col in df.columns:
                try:
                    df['Parsed_Date'] = pd.to_datetime(df[date_col], errors='coerce')
                    min_date = df['Parsed_Date'].min()
                    max_date = df['Parsed_Date'].max()
                    date_range = st.date_input("Date Range", value=(min_date, max_date), min_value=min_date, max_value=max_date, help="Filter by date range")
                    if len(date_range) == 2:
                        df = df[(df['Parsed_Date'].dt.date >= date_range[0]) & (df['Parsed_Date'].dt.date <= date_range[1])]
                except:
                    pass
    with col3:
        if 'Clean_Cost' in df.columns:
            min_price = float(df['Clean_Cost'].min())
            max_price = float(df['Clean_Cost'].max())
            price_range = st.slider("Price Range (TK)", min_value=min_price, max_value=max_price, value=(min_price, max_price), help="Filter by product price")
            df = df[(df['Clean_Cost'] >= price_range[0]) & (df['Clean_Cost'] <= price_range[1])]
    if selected_categories:
        df = df[df['Category'].isin(selected_categories)]
    return df


def show_success_animation():
    st.markdown("""
        <div style='text-align: center; padding: 20px;'>
            <h3 style='color: #28a745;'>Analytics Generated Successfully!</h3>
            <div style='font-size: 3rem;'>🎈🎉🎈</div>
        </div>
    """, unsafe_allow_html=True)


def main():
    st.set_page_config(page_title="Catwise Analytics - Enhanced", page_icon="📊", layout="wide")
    if 'current_mapping' not in st.session_state:
        st.session_state.current_mapping = {}
    if 'loaded_preset' not in st.session_state:
        st.session_state.loaded_preset = None
    if 'preset_loaded' not in st.session_state:
        st.session_state.preset_loaded = False
    if 'show_filters' not in st.session_state:
        st.session_state.show_filters = False
    if 'results' not in st.session_state:
        st.session_state.results = None
    if 'processed_df' not in st.session_state:
        st.session_state.processed_df = None
    apply_custom_styles()
    logo_b64 = load_logo()
    render_sidebar()
    st.title("📊 Catwise Analytics - Enhanced Dashboard")
    uploaded_file = st.file_uploader("📁 Upload Sales Data (Excel or CSV)", type=['xlsx', 'csv'], help="Upload your sales data file. Supports Excel (.xlsx) and CSV formats.")
    if not uploaded_file:
        render_empty_state()
        render_footer(logo_b64)
        return
    try:
        df = pd.read_csv(uploaded_file) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file)
        st.success(f"File loaded: {uploaded_file.name}")
        with st.expander("🔍 Preview Data", expanded=False):
            st.dataframe(df.head(20), use_container_width=True)
            st.caption(f"Showing first 20 rows. Total rows: {len(df)}")
        st.divider()
        st.markdown("### Column Mapping")
        auto_cols = find_columns(df)
        all_cols = list(df.columns)
        if st.session_state.preset_loaded and st.session_state.loaded_preset:
            auto_cols = st.session_state.loaded_preset
            st.session_state.preset_loaded = False
        col_a, col_b = st.columns([1, 3])
        with col_a:
            if st.button("🔄 Auto-Detect All Columns", help="Automatically detect all column types"):
                auto_cols = find_columns(df)
                st.rerun()
        with col_b:
            is_mapped = all(k in auto_cols for k in ['name', 'cost', 'qty'])
            if is_mapped:
                st.success("All required columns detected!")
            else:
                st.warning("Some required columns missing. Please verify.")
        st.markdown("**Map your columns:**")
        cols_per_row = 3
        cols = list(COLUMN_DISPLAY_ORDER)
        for i in range(0, len(cols), cols_per_row):
            row_cols = cols[i:i+cols_per_row]
            st_cols = st.columns(cols_per_row)
            for j, col_key in enumerate(row_cols):
                with st_cols[j]:
                    config = PRIORITY_COLUMN_MAPPING[col_key]
                    current_val = auto_cols.get(col_key, "None") if col_key != 'date' else auto_cols.get(col_key, None)
                    if col_key in ['date', 'order_id', 'phone']:
                        options = ["None"] + all_cols
                        default_idx = options.index(current_val) if current_val in options else 0
                        selected = st.selectbox(f"{config['exact_aliases'][0].title()}", options=options, index=default_idx, help=f"Select {col_key} column or None")
                        mapping_value = selected if selected != "None" else None
                    else:
                        default_idx = all_cols.index(current_val) if current_val in all_cols else 0
                        selected = st.selectbox(f"{config['exact_aliases'][0].title()}", options=all_cols, index=default_idx, help=f"Select {col_key} column")
                        mapping_value = selected
                    auto_cols[col_key] = mapping_value
        st.session_state.current_mapping = auto_cols
        required_cols = ['name', 'cost', 'qty']
        missing_cols = [c for c in required_cols if not auto_cols.get(c)]
        st.divider()
        col1, col2, col3 = st.columns([1, 1, 1])
        with col2:
            if st.button("🚀 Generate Analytics", disabled=len(missing_cols) > 0, help="Click to generate analytics. Required columns must be mapped."):
                with st.spinner("Processing data... Please wait"):
                    results = process_analytics(df, auto_cols)
                    st.session_state.results = results
                    st.session_state.processed_df = df.copy()
                    st.session_state.show_filters = True
                    st.rerun()
        if missing_cols:
            st.error(f"Missing required columns: {', '.join(missing_cols)}")
        if st.session_state.results:
            results = st.session_state.results
            df_processed = st.session_state.processed_df
            show_success_animation()
            if st.session_state.show_filters:
                df_filtered = render_filters(df_processed, results)
                if len(df_filtered) != len(df_processed):
                    results = process_analytics(df_filtered, auto_cols)
                    st.session_state.results = results
            else:
                df_filtered = df_processed
            st.divider()
            st.markdown("### Key Metrics")
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Total Orders", f"{results['total_orders']:,.0f}", help="Total number of unique orders/transactions")
            m2.metric("Units Sold", f"{results['total_qty']:,.0f}", help="Total quantity of products sold")
            m3.metric("Gross Revenue", f"TK {results['total_rev']:,.0f}", help="Total revenue from all sales")
            if results['avg_basket_value'] > 0:
                m4.metric("Avg. Basket Size", f"TK {results['avg_basket_value']:,.0f}", help="Average value per order/basket")
            else:
                m4.metric("Avg. Basket Size", "0")
            st.divider()
            if results.get('timeseries') is not None and len(results['timeseries']) > 1:
                st.markdown("### Sales Trend Over Time")
                timeseries_df = results['timeseries']
                fig_ts = go.Figure()
                fig_ts.add_trace(go.Scatter(x=timeseries_df['Parsed_Date'], y=timeseries_df['Total Amount'], mode='lines+markers', name='Revenue', line=dict(color=DEEN_BRAND_COLOR, width=3), marker=dict(size=8), fill='tozeroy', fillcolor="rgba(0, 123, 255, 0.2)"))
                fig_ts.update_layout(height=400, hovermode='x unified', xaxis_title="Date", yaxis_title="Revenue (TK)", template="plotly_white", showlegend=True)
                st.plotly_chart(fig_ts, use_container_width=True)
                st.divider()
            st.markdown("### Category Analysis")
            v1, v2 = st.columns(2)
            summ_sorted = results['summary'].sort_values('Total Amount', ascending=False)
            color_seq = [get_category_color(cat) for cat in summ_sorted['Category']]
            with v1:
                fig_pie = px.pie(summ_sorted, values='Total Amount', names='Category', hole=0.5, title='Revenue by Category', color_discrete_sequence=color_seq)
                fig_pie.update_traces(textposition='inside', textinfo='percent+label', hovertemplate='<b>%{label}</b><br>Revenue: TK %{value:,.0f}<br>Percentage: %{percent}<extra></extra>')
                fig_pie.update_layout(height=500, showlegend=True, legend=dict(orientation="v", yanchor="middle", y=0.5, xanchor="left", x=1.05))
                st.plotly_chart(fig_pie, use_container_width=True)
            with v2:
                fig_bar = px.bar(summ_sorted, x='Category', y='Total Qty', color='Category', title='Volume by Category', color_discrete_sequence=color_seq)
                fig_bar.update_traces(hovertemplate='<b>%{x}</b><br>Quantity: %{y:,.0f}<extra></extra>')
                fig_bar.update_layout(height=500, xaxis_title="Category", yaxis_title="Total Quantity", showlegend=False, xaxis={'categoryorder': 'total descending'})
                st.plotly_chart(fig_bar, use_container_width=True)
            st.divider()
            t1, t2, t3 = st.tabs(["Category Summary", "Price-wise Breakdown", "Top Products"])
            with t1:
                df_breakdown = results['summary'].sort_values('Total Amount', ascending=False).copy()
                total_qty = df_breakdown['Total Qty'].sum()
                total_amount = df_breakdown['Total Amount'].sum()
                df_breakdown.loc[len(df_breakdown)] = ['**TOTAL SALES**', total_qty, total_amount]
                df_breakdown.index = range(1, len(df_breakdown) + 1)
                st.dataframe(df_breakdown[['Category', 'Total Qty', 'Total Amount']], use_container_width=True, height=400)
            with t2:
                df_drill = results['drilldown'].sort_values(['Category', 'Price'], ascending=[True, False]).copy()
                df_drill.columns = ['Category', 'Price', 'Qty', 'Total Amount']
                df_drill.loc[len(df_drill)] = ['**TOTAL (All Categories)**', None, df_drill['Qty'].sum(), df_drill['Total Amount'].sum()]
                df_drill.index = range(1, len(df_drill) + 1)
                st.dataframe(df_drill[['Category', 'Price', 'Qty', 'Total Amount']], use_container_width=True, height=400)
            with t3:
                df_top = results['top_items'].head(20).copy()
                df_top.index = range(1, len(df_top) + 1)
                st.dataframe(df_top[['Product Name', 'Category', 'Total Qty', 'Total Amount']], use_container_width=True, height=400)
            st.divider()
            st.markdown("### Export Data")
            col1, col2 = st.columns(2)
            with col1:
                buf = BytesIO()
                with pd.ExcelWriter(buf, engine='openpyxl') as writer:
                    df_breakdown = results['summary'].sort_values('Total Amount', ascending=False).copy()
                    total_qty = df_breakdown['Total Qty'].sum()
                    total_amount = df_breakdown['Total Amount'].sum()
                    df_breakdown.loc[len(df_breakdown)] = ['TOTAL SALES (Summary)', total_qty, total_amount]
                    df_breakdown.index = range(1, len(df_breakdown) + 1)
                    df_breakdown[['Category', 'Total Qty', 'Total Amount']].to_excel(writer, sheet_name='Category Summary', index=False)
                    df_drill = results['drilldown'].sort_values(['Category', 'Price'], ascending=[True, False]).copy()
                    df_drill.columns = ['Category', 'Price', 'Qty', 'Total Amount']
                    df_drill.loc[len(df_drill)] = ['TOTAL (All Categories)', None, df_drill['Qty'].sum(), df_drill['Total Amount'].sum()]
                    df_drill.index = range(1, len(df_drill) + 1)
                    df_drill[['Category', 'Price', 'Qty', 'Total Amount']].to_excel(writer, sheet_name='Price-wise Category', index=False)
                    wb = writer.book
                    header_fill = PatternFill(start_color='007BFF', end_color='007BFF', fill_type='solid')
                    header_font = Font(bold=True, color='FFFFFF')
                    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                    total_fill = PatternFill(start_color='ECECEC', end_color='ECECEC', fill_type='solid')
                    for sheet_name in ['Category Summary', 'Price-wise Category']:
                        ws = wb[sheet_name]
                        dframe = df_breakdown if sheet_name == 'Category Summary' else df_drill
                        for cell in ws[1]:
                            cell.fill = header_fill
                            cell.font = header_font
                            cell.alignment = Alignment(horizontal='center', vertical='center')
                            cell.border = thin_border
                        ws.freeze_panes = 'A2'
                        for i, col_name in enumerate(dframe.columns):
                            col_letter = ws.cell(row=1, column=i+1).column_letter
                            content_max = dframe[col_name].apply(lambda x: len(str(x)) if pd.notnull(x) else 0).max() if not dframe.empty else 0
                            max_len = max(content_max, len(str(col_name))) + 2
                            ws.column_dimensions[col_letter].width = max_len
                            for row_idx in range(2, ws.max_row + 1):
                                cell = ws.cell(row=row_idx, column=i+1)
                                cell.border = thin_border
                                first_cell_val = str(ws.cell(row=row_idx, column=1).value).upper()
                                if 'TOTAL' in first_cell_val:
                                    cell.font = Font(bold=True)
                                    cell.fill = total_fill
                                if col_name in ['Category', 'Product Name']:
                                    cell.alignment = Alignment(horizontal='left')
                                else:
                                    cell.alignment = Alignment(horizontal='right')
                                if any(x in col_name for x in ['Qty', 'Price', 'Amount']):
                                    cell.number_format = '#,##0'
                fname = f"Sales_Report_{results['timeframe']}.xlsx"
                st.download_button("Download Excel Report", data=buf.getvalue(), file_name=fname, help="Download comprehensive Excel report with styled sheets")
            with col2:
                buf_raw = BytesIO()
                df_filtered.to_excel(buf_raw, index=False, engine='openpyxl')
                st.download_button("Download Raw Data", data=buf_raw.getvalue(), file_name=f"Raw_Data_{results['timeframe']}.xlsx", help="Download the processed raw data")
            if results.get('all_categories'):
                st.markdown("### Category Insights")
                col1, col2 = st.columns(2)
                with col1:
                    st.metric("Total Categories", len(results['all_categories']))
                with col2:
                    top_cat = results['summary'].sort_values('Total Amount', ascending=False).iloc[0]
                    st.metric("Top Category", f"{top_cat['Category']}")
    except Exception as e:
        st.error(f"Processing Error: {e}")
        log_event("CRASH", str(e))
        st.error("Please check your file format and try again.")
    render_footer(logo_b64)


if __name__ == "__main__":
    main()
