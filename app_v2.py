from pandas.errors import ClosedFileError
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os
import json
import base64
from datetime import datetime, timedelta
from io import BytesIO
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import re
from collections import Counter

# --- Configuration & Styling ---
FEEDBACK_DIR = "feedback"
MAPPING_PRESETS_FILE = os.path.join(FEEDBACK_DIR, "mapping_presets.json")
os.makedirs(FEEDBACK_DIR, exist_ok=True)

# Dynamic category detection - no hardcoded mappings!
# We'll detect categories from the data itself

CATEGORY_COLORS = px.colors.qualitative.Set3 + px.colors.qualitative.Pastel + px.colors.qualitative.Set2
DEEN_BRAND_COLOR = '#007BFF'
LOGO_PNG = "assets/deen_logo.png"


def load_logo():
    """Loads logo as base64 string."""
    if os.path.exists(LOGO_PNG):
        try:
            with open(LOGO_PNG, "rb") as f:
                return base64.b64encode(f.read()).decode()
        except: pass
    return ""


def apply_custom_styles():
    """Applies premium CSS styles to the dashboard."""
    st.markdown("""
    <style>
    .main { background-color: transparent; }
    @media (max-width: 768px) {
        .mobile-only { display: block !important; }
        .desktop-only { display: none !important; }
    }
    @media (min-width: 769px) {
        .mobile-only { display: none !important; }
        .desktop-only { display: block !important; }
    }
    div[data-testid="stMetric"] {
        background-color: var(--background-secondary, #ffffff);
        padding: 24px;
        border-radius: 16px;
        box-shadow: 0 4px 20px rgba(0,0,0,0.06);
        border: 1px solid var(--secondary-background-color, #ececec);
        transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
        position: relative;
        overflow: hidden;
    }
    div[data-testid="stMetric"]:hover {
        transform: translateY(-4px);
        box-shadow: 0 12px 24px rgba(0,0,0,0.12);
        border-color: #007bff;
    }
    [data-testid="stMetricLabel"] {
        font-weight: 600 !important;
        font-size: 1.1rem !important;
        margin-bottom: 8px !important;
        color: var(--text-color, #1a1a1b) !important;
    }
    [data-testid="stMetricValue"], [data-testid="stMetricValue"] > div {
        font-weight: 800 !important;
        font-size: 1.8rem !important;
        color: var(--text-color, #1a1a1b) !important;
    }
    .trend-indicator {
        position: absolute; top: 10px; right: 10px;
        font-size: 0.9rem; padding: 4px 8px; border-radius: 12px;
        font-weight: 600;
    }
    .trend-up { background: rgba(40, 167, 69, 0.2); color: #28a745; }
    .trend-down { background: rgba(220, 53, 69, 0.2); color: #dc3545; }
    .trend-neutral { background: rgba(108, 117, 125, 0.2); color: #6c757d; }
    @media (prefers-color-scheme: dark) {
        [data-testid="stMetricLabel"], [data-testid="stMetricValue"] { color: #ffffff !important; }
        div[data-testid="stMetric"] { background-color: #1e1e1e !important; border-color: #333 !important; }
        .trend-up { background: rgba(40, 167, 69, 0.3); }
        .trend-down { background: rgba(220, 53, 69, 0.3); }
        .trend-neutral { background: rgba(108, 117, 125, 0.3); }
    }
    [data-theme="dark"] [data-testid="stMetricLabel"], [data-theme="dark"] [data-testid="stMetricValue"] { color: #ffffff !important; }
    [data-theme="dark"] div[data-testid="stMetric"] { background-color: #1e1e1e !important; border-color: #333; }
    [data-theme="dark"] .trend-up { background: rgba(40, 167, 69, 0.3); }
    [data-theme="dark"] .trend-down { background: rgba(220, 53, 69, 0.3); }
    [data-theme="dark"] .trend-neutral { background: rgba(108, 117, 125, 0.3); }
    .stButton>button {
        width: 100%; border-radius: 12px; height: 3.5em;
        background: linear-gradient(135deg, #007bff 0%, #0056b3 100%);
        color: white; font-weight: 600; border: none; transition: all 0.3s ease;
        box-shadow: 0 4px 12px rgba(0,123,255,0.3);
    }
    .stButton>button:hover { transform: scale(1.02); background: linear-gradient(135deg, #0056b3 0%, #004085 100%); box-shadow: 0 6px 16px rgba(0,123,255,0.4); }
    .stButton>button:active { transform: scale(0.98); }
    div[data-testid="stExpander"] { border: 1px solid #eee; box-shadow: 0 4px 12px rgba(0,0,0,0.05); background: white; border-radius: 14px; margin-bottom: 1.5rem; padding: 5px; transition: all 0.3s ease; }
    div[data-testid="stExpander"]:hover { border-color: #007bff; box-shadow: 0 6px 16px rgba(0,123,255,0.15); }
    @media (prefers-color-scheme: dark) { div[data-testid="stExpander"] { background-color: #1a1a1a !important; border-color: #333 !important; } }
    [data-theme="dark"] div[data-testid="stExpander"] { background-color: #1a1a1a !important; border-color: #333 !important; }
    .empty-state { text-align: center; padding: 80px 20px; background: var(--background-secondary); border-radius: 20px; border: 2px dashed #ddd; margin: 40px 0; }
    .empty-state h2 { color: var(--text-color); margin-bottom: 10px; }
    .empty-state p { color: var(--text-color); opacity: 0.7; }
    .empty-state-icon { font-size: 4rem; margin-bottom: 20px; opacity: 0.8; }
    .filter-section { background: var(--background-secondary); padding: 20px; border-radius: 16px; border: 1px solid #eee; margin-bottom: 20px; box-shadow: 0 4px 12px rgba(0,0,0,0.05); }
    .filter-title { font-size: 1.2rem; font-weight: 700; margin-bottom: 15px; color: var(--text-color); }
    .card { background: var(--background-secondary); padding: 24px; border-radius: 16px; border: 1px solid #eee; margin-bottom: 20px; box-shadow: 0 4px 12px rgba(0,0,0,0.05); transition: all 0.3s ease; }
    .card:hover { box-shadow: 0 8px 20px rgba(0,0,0,0.1); border-color: #007bff; }
    .chart-container { background: var(--background-secondary); padding: 20px; border-radius: 16px; border: 1px solid #eee; margin-bottom: 20px; box-shadow: 0 4px 12px rgba(0,0,0,0.05); }
    .sticky-footer { position: fixed; left: 0; bottom: 0; width: 100%; background-color: var(--background-color) !important; color: var(--text-color) !important; text-align: center; padding: 20px 0; border-top: 1px solid rgba(128, 128, 128, 0.2); z-index: 999999; font-size: 0.9rem; backdrop-filter: blur(10px); }
    [data-theme="dark"] .sticky-footer { background-color: #0e1117 !important; box-shadow: 0 -5px 25px rgba(0,0,0,0.4); }
    .footer-content-inner { display: flex; align-items: center; justify-content: center; gap: 6px; }
    .brand-wrapper { display: flex; align-items: center; gap: 2px; }
    .small-logo { height: 24px; width: auto; filter: grayscale(0%); }
    .brand-name { font-weight: 700; color: var(--primary-color); }
    .block-container { padding-bottom: 150px !important; padding-top: 3rem !important; }
    h1, h2, h3, h4, h5, h6, p, span, label, .stMarkdown { color: var(--text-color) !important; }
    div[data-testid="stSelectbox"] > div, div[data-testid="stTextInput"] > div, div[data-testid="stTextArea"] > div { background-color: var(--secondary-background-color) !important; color: var(--text-color) !important; border-radius: 12px !important; border: 1px solid rgba(128, 128, 128, 0.2) !important; transition: all 0.2s ease !important; }
    div[data-testid="stSelectbox"] > div:hover { border-color: #007bff !important; box-shadow: 0 0 0 3px rgba(0,123,255,0.1) !important; }
    .stButton>button { background-color: var(--primary-color) !important; color: white !important; font-weight: 700; box-shadow: 0 4px 12px rgba(0,123,255,0.3); }
    .stButton>button:hover { opacity: 0.9; transform: scale(1.02); box-shadow: 0 6px 16px rgba(0,123,255,0.4); }
    div[data-testid="stTabs"] button { color: var(--text-color) !important; opacity: 0.7; padding: 12px 24px !important; font-weight: 600 !important; }
    div[data-testid="stTabs"] button[aria-selected="true"] { color: var(--primary-color) !important; opacity: 1; border-bottom: 3px solid var(--primary-color) !important; background: rgba(0,123,255,0.05) !important; }
    hr { border-color: rgba(128, 128, 128, 0.2) !important; }
    @media (max-width: 768px) { .block-container { padding-left: 1rem !important; padding-right: 1rem !important; padding-bottom: 220px !important; } .footer-content-inner { flex-direction: column; gap: 8px; } [data-testid="stMetricValue"] { font-size: 1.4rem !important; } .sticky-footer { padding: 15px 5px !important; } }
    </style>
    """, unsafe_allow_html=True)


# --- Helper Functions ---

def log_event(event_type, details):
    """Logs system events to JSON."""
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
    except: pass


def categorize_product_dynamically(name):
    """
    Dynamically categorizes products based on their names without hardcoded mappings.
    Uses pattern matching and common naming conventions.
    """
    if pd.isna(name) or not name:
        return 'Unknown'
    
    name_lower = str(name).lower()
    
    # Define category patterns dynamically
    category_patterns = {
        # Clothing types
        'T-Shirt': [r't\s?shirt', r't\s?-?shirt', r'tee', r'tee shirt'],
        'FS T-Shirt': [r'full\s+sleeve\s+t\s?-?shirt', r'fs\s+t\s?-?shirt', r'long\s+sleeve\s+tee'],
        'HS T-Shirt': [r'half\s+sleeve\s+t\s?-?shirt', r'hs\s+t\s?-?shirt'],
        'Polo Shirt': [r'polo', r'polo\s+shirt'],
        'Formal Shirt': [r'formal\s+shirt', r'executive\s+shirt', r'dress\s+shirt'],
        'Casual Shirt': [r'casual\s+shirt'],
        'Denim Shirt': [r'denim\s+shirt'],
        'Flannel Shirt': [r'flannel\s+shirt', r'flannel'],
        'Cuban Shirt': [r'cuban\s+shirt', r'cuban\s+collar'],
        'Panjabi': [r'panjabi', r'punjabi'],
        
        # Bottoms
        'Jeans': [r'jeans', r'jeans\s+pant'],
        'Trousers': [r'trousers?', r'pant[s]?', r'trcargo', r'joggers?', r'track\s+pant'],
        'Cargo Trousers': [r'cargo\s+trousers?', r'cargo\s+pant[s]?'],
        'Twill Chino': [r'twill\s+chino'],
        'Shorts': [r'shorts', r'short\s+pant'],
        
        # Outerwear
        'Jacket': [r'jacket', r'outerwear', r'coat', r'blazer'],
        'Sweatshirt': [r'sweatshirt', r'hoodie', r'pullover'],
        'Sweater': [r'sweater', r'cardigan', r'knitwear', r'wool\s+sweater'],
        'Belt': [r'belt'],
        
        # Accessories
        'Cap': [r'cap', r'baseball\s+cap', r'beanie'],
        'Mask': [r'mask', r'face\s+mask'],
        'Wallet': [r'wallet'],
        'Belt': [r'belt'],
        'Turtleneck': [r'turtleneck', r'mock\s+neck'],
        'Drop Shoulder': [r'drop\s+shoulder'],
        
        # Bags
        'Leather Bag': [r'bag', r'backpack', r'leather\s+bag', r'sling\s+bag'],
        'Passport Holder': [r'passport\s+holder', r'passport\s+cover'],
        'Water Bottle': [r'water\s+bottle', r'bottle'],
        
        # Special items
        'Kaftan Shirt': [r'kaftan'],
        'Active Wear': [r'active\s+wear', r'sportswear'],
        'Jersy': [r'jersy', r'jersey'],
        'Boxer': [r'boxer'],
        'Tank Top': [r'tank\s+top', r'tanktop', r'tank', r'sleeveless'],
        'Contrast Shirt': [r'contrast'],
    }
    
    # Check patterns in order of specificity
    for category, patterns in category_patterns.items():
        for pattern in patterns:
            if re.search(pattern, name_lower):
                return category
    
    # Additional intelligent categorization
    # Check for sleeve types
    if re.search(r'full\s+sleeve|long\s+sleeve|fs|l/s', name_lower):
        if re.search(r'shirt', name_lower):
            return 'FS Shirt'
        elif re.search(r't\s?-?shirt|tee', name_lower):
            return 'FS T-Shirt'
        elif re.search(r'sweater|sweatshirt|hoodie', name_lower):
            return 'FS Sweatshirt'
    
    if re.search(r'half\s+sleeve|hs|short\s+sleeve', name_lower):
        if re.search(r'shirt', name_lower):
            return 'HS Shirt'
        elif re.search(r't\s?-?shirt|tee', name_lower):
            return 'HS T-Shirt'
    
    # Check for material
    if re.search(r'leather', name_lower):
        if re.search(r'bag|backpack', name_lower):
            return 'Leather Bag'
        return 'Leather Item'
    
    # Default category
    return 'Others'


def extract_categories_from_data(product_names):
    """
    Extract unique categories from product names dynamically.
    """
    categories = set()
    for name in product_names:
        category = categorize_product_dynamically(name)
        categories.add(category)
    return sorted(categories)


def find_columns(df):
    """Auto-detects columns from dataframe with dynamic and flexible matching."""
    column_types = {
        'name': {
            'aliases': ['item name', 'product name', 'product', 'item', 'title', 'description', 'name', 'product title', 'item title'],
            'keywords': ['name', 'product', 'item', 'title', 'desc'],
            'priority': 1
        },
        'cost': {
            'aliases': ['item cost', 'price', 'unit price', 'cost', 'rate', 'mrp', 'selling price', 'amount', 'unit cost', 'sell price'],
            'keywords': ['price', 'cost', 'rate', 'mrp', 'amount', 'sell', 'value', 'unit'],
            'priority': 1
        },
        'qty': {
            'aliases': ['quantity', 'qty', 'units', 'sold', 'count', 'total quantity', 'item quantity', 'product quantity'],
            'keywords': ['quantity', 'qty', 'unit', 'sold', 'count', 'pcs', 'piece'],
            'priority': 1
        },
        'date': {
            'aliases': ['date', 'order date', 'month', 'time', 'created at', 'order time', 'purchase date'],
            'keywords': ['date', 'time', 'month', 'year', 'day', 'created', 'purchase', 'order date'],
            'priority': 2
        },
        'order_id': {
            'aliases': ['order id', 'order #', 'invoice number', 'invoice #', 'order number', 'transaction id', 'id', 'invoice', 'order no'],
            'keywords': ['order', 'invoice', 'transaction', 'id', '#', 'no', 'number'],
            'priority': 2
        },
        'phone': {
            'aliases': ['phone', 'contact', 'mobile', 'cell', 'phone number', 'customer phone', 'customer contact'],
            'keywords': ['phone', 'contact', 'mobile', 'cell', 'customer', 'call'],
            'priority': 2
        }
    }
    
    found = {}
    actual_cols = list(df.columns)
    lower_cols = [str(c).strip().lower() for c in actual_cols]
    used_indices = set()
    
    sorted_types = sorted(column_types.items(), key=lambda x: x[1]['priority'])
    
    for key, config in sorted_types:
        aliases = config['aliases']
        keywords = config['keywords']
        
        for alias in aliases:
            if alias in lower_cols:
                idx = lower_cols.index(alias)
                if idx not in used_indices:
                    found[key] = actual_cols[idx]
                    used_indices.add(idx)
                    break
        
        if key not in found:
            for i, col in enumerate(lower_cols):
                if i not in used_indices:
                    if any(alias.lower() in col for alias in aliases):
                        found[key] = actual_cols[i]
                        used_indices.add(i)
                        break
        
        if key not in found:
            for i, col in enumerate(lower_cols):
                if i not in used_indices:
                    if any(kw.lower() in col for kw in keywords):
                        found[key] = actual_cols[i]
                        used_indices.add(i)
                        break
    
    for i, col in enumerate(lower_cols):
        if i not in used_indices:
            col_name = actual_cols[i]
            try:
                sample_data = df[col_name].dropna().head(10)
                if all(isinstance(x, (int, float)) or (isinstance(x, str) and any(c.isdigit() for c in str(x))) for x in sample_data):
                    if any(kw in col for kw in ['price', 'cost', 'rate', 'amount']):
                        if 'cost' not in found:
                            found['cost'] = col_name
                            used_indices.add(i)
                    elif any(kw in col for kw in ['qty', 'quantity', 'unit', 'pcs']):
                        if 'qty' not in found:
                            found['qty'] = col_name
                            used_indices.add(i)
                    elif any(kw in col for kw in ['order', 'invoice', 'id']):
                        if 'order_id' not in found:
                            found['order_id'] = col_name
                            used_indices.add(i)
            except: pass
    
    return found


def save_mapping_preset(preset_name, mapping):
    """Save a column mapping preset."""
    try:
        presets = {}
        if os.path.exists(MAPPING_PRESETS_FILE):
            with open(MAPPING_PRESETS_FILE, 'r') as f:
                presets = json.load(f)
        presets[preset_name] = mapping
        with open(MAPPING_PRESETS_FILE, 'w') as f:
            json.dump(presets, f, indent=4)
        return True
    except Exception as e:
        log_event("ERROR", f"Failed to save preset: {str(e)}")
        return False


def load_mapping_presets():
    """Load saved mapping presets."""
    try:
        if os.path.exists(MAPPING_PRESETS_FILE):
            with open(MAPPING_PRESETS_FILE, 'r') as f:
                return json.load(f)
        return {}
    except:
        return {}


def delete_mapping_preset(preset_name):
    """Delete a mapping preset."""
    try:
        if os.path.exists(MAPPING_PRESETS_FILE):
            with open(MAPPING_PRESETS_FILE, 'r') as f:
                presets = json.load(f)
            if preset_name in presets:
                del presets[preset_name]
                with open(MAPPING_PRESETS_FILE, 'w') as f:
                    json.dump(presets, f, indent=4)
        return True
    except:
        return False


def get_category_color(category):
    """Get consistent color for a category."""
    all_categories = ['T-Shirt', 'FS T-Shirt', 'HS T-Shirt', 'Polo Shirt', 'Formal Shirt',
                     'Casual Shirt', 'Denim Shirt', 'Flannel Shirt', 'Cuban Shirt', 'Panjabi',
                     'Jeans', 'Trousers', 'Cargo Trousers', 'Twill Chino', 'Shorts', 'Jacket',
                     'Sweatshirt', 'Sweater', 'Belt', 'Cap', 'Mask', 'Wallet', 'Leather Bag',
                     'Passport Holder', 'Water Bottle', 'Kaftan Shirt', 'Active Wear', 'Jersy',
                     'Boxer', 'Tank Top', 'Contrast Shirt', 'Turtleneck', 'Drop Shoulder', 'Others', 'Unknown']
    if category in all_categories:
        idx = all_categories.index(category) % len(CATEGORY_COLORS)
        return CATEGORY_COLORS[idx]
    return '#999999'


def process_analytics(df, mapping):
    """Core data processing and metric calculation."""
    df = df.copy()
    
    # 1. Clean Data
    df['Clean_Name'] = df[mapping['name']].fillna('Unknown').astype(str)
    exclude_keywords = 'choose any|shoose any|combo|bundle|bundel'
    df = df[~df['Clean_Name'].str.contains(exclude_keywords, case=False, na=False)]

    def clean_numeric(val):
        if pd.isna(val): return 0
        if isinstance(val, (int, float)): return val
        clean_val = ''.join(c for c in str(val) if c.isdigit() or c == '.')
        try: return float(clean_val) if clean_val else 0
        except ValueError: return 0

    df['Clean_Cost'] = df[mapping['cost']].apply(clean_numeric)
    df['Clean_Qty'] = df[mapping['qty']].apply(clean_numeric)
    df.loc[df['Clean_Qty'] < 0, 'Clean_Qty'] = 0
    df['Total Amount'] = df['Clean_Cost'] * df['Clean_Qty']
    
    # Dynamic category detection
    df['Category'] = df['Clean_Name'].apply(categorize_product_dynamically)
    
    # 2. Timeframe Detection
    timeframe = ""
    date_col = mapping.get('date')
    if date_col and date_col in df.columns:
        try:
            dates = pd.to_datetime(df[date_col], errors='coerce').dropna()
            if not dates.empty:
                if dates.dt.to_period('M').nunique() == 1:
                    timeframe = dates.iloc[0].strftime("%B_%Y")
                else:
                    timeframe = f"{dates.min().strftime('%d%b')}_to_{dates.max().strftime('%d%b_%y')}"
                df['Date'] = dates
                df['Date_Str'] = dates.dt.strftime('%Y-%m-%d')
        except: timeframe = "Report"
    else:
        df['Date'] = None
        df['Date_Str'] = None

    # 3. Aggregations
    summary = df.groupby('Category').agg({'Clean_Qty': 'sum', 'Total Amount': 'sum'}).reset_index()
    summary.columns = ['Category', 'Total Qty', 'Total Amount']
    
    t_rev = summary['Total Amount'].sum()
    t_qty = summary['Total Qty'].sum()
    
    drilldown = df.groupby(['Category', 'Clean_Cost']).agg({'Clean_Qty': 'sum', 'Total Amount': 'sum'}).reset_index()
    drilldown.columns = ['Category', 'Price', 'Total Qty', 'Total Amount']
    
    top_items = df.groupby('Clean_Name').agg({'Clean_Qty': 'sum', 'Total Amount': 'sum', 'Category': 'first'}).reset_index()
    top_items.columns = ['Product Name', 'Total Qty', 'Total Amount', 'Category']
    top_items = top_items.sort_values('Total Amount', ascending=False)
    
    # 4. Basket Metrics
    avg_basket_value = 0
    group_cols = [c for c in [mapping.get('order_id'), mapping.get('phone')] if c and c in df.columns]
    
    if group_cols:
        order_groups = df.groupby(group_cols).agg({'Total Amount': 'sum'})
        avg_basket_value = order_groups['Total Amount'].mean()
    
    # 5. Time Series Data
    timeseries_data = None
    if date_col and date_col in df.columns and df['Date'].notna().any():
        timeseries_data = df.groupby('Date_Str').agg({'Total Amount': 'sum', 'Clean_Qty': 'sum'}).reset_index()
        timeseries_data.columns = ['Date', 'Revenue', 'Quantity']
        timeseries_data = timeseries_data.sort_values('Date')
    
    return {
        'drilldown': drilldown,
        'summary': summary,
        'top_items': top_items,
        'timeseries': timeseries_data,
        'timeframe': timeframe,
        'avg_basket_value': avg_basket_value,
        'total_qty': t_qty,
        'total_rev': t_rev,
        'total_orders': len(order_groups) if group_cols else 0,
        'df': df,
        'date_col': date_col,
        'all_categories': sorted(df['Category'].unique())
    }


# --- UI Components ---

def render_sidebar():
    with st.sidebar:
        st.markdown("### 🎯 Quick Actions")
        
        with st.expander("❓ How to Use", expanded=False):
            st.markdown("""
            **Quick Start:**
            1. 📁 **Upload** your sales data (Excel/CSV)
            2. 🔍 **Auto-detect** columns or map manually
            3. ⚡ **Generate** analytics
            
            **Category Detection:**
            Categories are detected dynamically from product names using pattern matching.
            No hardcoded category list needed!
            
            **Tips:**
            - Use 🔍 Auto-Detect for quick setup
            - Save mappings as presets for reuse
            - Use filters to analyze specific data
            """)
        
        st.divider()
        st.markdown("### 💾 Mapping Presets")
        
        presets = load_mapping_presets()
        if presets:
            selected_preset = st.selectbox("Load Preset", ["None"] + list(presets.keys()))
            if selected_preset != "None":
                if st.button("Apply", key="apply_preset", use_container_width=True):
                    st.session_state.loaded_preset = presets[selected_preset]
                    st.session_state.preset_applied = True
                    st.success(f"✅ Preset loaded!")
            
            with st.expander("Manage"):
                preset_to_delete = st.selectbox("Delete", ["None"] + list(presets.keys()), key="del_preset")
                if preset_to_delete != "None" and st.button("Delete", key="del_btn"):
                    if delete_mapping_preset(preset_to_delete):
                        st.success(f"✅ Deleted '{preset_to_delete}'")
                        st.rerun()
        else:
            st.info("No presets yet. Save your first mapping below!")
        
        st.divider()
        st.markdown("### 🎨 Display Options")
        show_timeseries = st.checkbox("Show Time Series", value=True)
        show_pie = st.checkbox("Show Pie Chart", value=True)
        show_bar = st.checkbox("Show Bar Chart", value=True)
        st.session_state.chart_prefs = {'timeseries': show_timeseries, 'pie': show_pie, 'bar': show_bar}
        
        st.divider()
        st.markdown("### 🐛 Feedback")
        comment = st.text_area("Report Issues", placeholder="Category detection issue...")
        if st.button("Submit"):
            feedback_file = os.path.join(FEEDBACK_DIR, "user_feedback.json")
            entry = {"timestamp": datetime.now().isoformat(), "comment": comment}
            try:
                data = []
                if os.path.exists(feedback_file):
                    with open(feedback_file, "r") as f: data = json.load(f)
                data.append(entry)
                with open(feedback_file, "w") as f: json.dump(data, f, indent=4)
                st.success("Feedback saved!")
            except: st.error("Failed to save")
        
        if st.checkbox("View Logs"):
            log_path = os.path.join(FEEDBACK_DIR, "system_logs.json")
            if os.path.exists(log_path):
                with open(log_path, "r") as f: st.json(json.load(f)[-10:])


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
    <div class="empty-state">
        <div class="empty-state-icon">📊</div>
        <h2>Welcome to Sales Dashboard</h2>
        <p>Upload your sales data to get started</p>
        <p style='color: #666; margin-top: 10px;'>Supports: Excel (.xlsx) and CSV files</p>
        <p style='color: #666; font-size: 0.9rem; margin-top: 20px;'>
            <strong>Tip:</strong> Use auto-detect for instant column mapping!
        </p>
    </div>
    """, unsafe_allow_html=True)


def render_filters(df, results):
    if df is None or 'Category' not in df.columns:
        return df
    
    st.markdown('<div class="filter-section"><div class="filter-title">🎯 Filters</div>', unsafe_allow_html=True)
    col1, col2, col3 = st.columns(3)
    
    with col1:
        all_categories = ['All'] + sorted(df['Category'].unique().tolist())
        selected_categories = st.multiselect("Categories", options=all_categories, default=['All'], help="Filter by categories")
    
    with col2:
        if results and results.get('date_col') and 'Date' in df.columns and df['Date'].notna().any():
            min_date = df['Date'].min().to_pydatetime()
            max_date = df['Date'].max().to_pydatetime()
            date_range = st.date_input("Date Range", value=(min_date, max_date), min_value=min_date, max_value=max_date, help="Filter by date")
        else:
            date_range = None
    
    with col3:
        if df['Clean_Cost'].max() > 0:
            price_range = st.slider("Price Range (TK)", min_value=0, max_value=float(df['Clean_Cost'].max()), value=(0, float(df['Clean_Cost'].max())), help="Filter by price")
        else:
            price_range = (0, 1)
    
    st.markdown('</div>', unsafe_allow_html=True)
    
    filtered_df = df.copy()
    if 'All' not in selected_categories and selected_categories:
        filtered_df = filtered_df[filtered_df['Category'].isin(selected_categories)]
    if date_range and len(date_range) == 2 and 'Date' in filtered_df.columns:
        filtered_df = filtered_df[(filtered_df['Date'].dt.date >= date_range[0]) & (filtered_df['Date'].dt.date <= date_range[1])]
    if price_range and len(price_range) == 2:
        filtered_df = filtered_df[(filtered_df['Clean_Cost'] >= price_range[0]) & (filtered_df['Clean_Cost'] <= price_range[1])]
    
    return filtered_df


# --- Main App ---

def main():
    st.set_page_config(page_title="Sales Dashboard", page_icon="📊", layout="wide")
    apply_custom_styles()
    
    # Initialize session state
    for key in ['uploaded_file', 'auto_cols', 'mapping', 'results', 'filtered_df',
                'chart_prefs', 'loaded_preset', 'preset_applied', 'show_advanced']:
        if key not in st.session_state:
            st.session_state[key] = None if key not in ['chart_prefs', 'show_advanced'] else {}
    
    if st.session_state.chart_prefs is None:
        st.session_state.chart_prefs = {'timeseries': True, 'pie': True, 'bar': True}
    if st.session_state.show_advanced is None:
        st.session_state.show_advanced = False
    if st.session_state.preset_applied is None:
        st.session_state.preset_applied = False
    if st.session_state.loaded_preset is None:
        st.session_state.loaded_preset = None
    
    logo_b64 = load_logo()
    render_sidebar()
    
    st.title("🚀 Sales Performance Dashboard")
    
    if st.session_state.uploaded_file is None:
        render_empty_state()
    
    uploaded_file = st.file_uploader(
        "📁 Upload Sales Data", 
        type=['xlsx', 'csv'],
        help="Upload Excel (.xlsx) or CSV file"
    )
    
    if uploaded_file:
        st.session_state.uploaded_file = uploaded_file.name
        
        try:
            with st.spinner("📂 Uploading and parsing file..."):
                df = pd.read_csv(uploaded_file) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file)
            
            st.success(f"✅ Attached: {uploaded_file.name}")
            st.session_state.original_df = df.copy()
            
            with st.expander("🔍 Preview Data", expanded=False):
                st.dataframe(df.head(10), use_container_width=True)
                col1, col2, col3 = st.columns(3)
                col1.metric("Total Rows", len(df))
                col2.metric("Total Columns", len(df.columns))
            
            st.markdown("---")
            st.subheader("⚙️ Column Mapping")
            
            if st.session_state.preset_applied and st.session_state.loaded_preset:
                auto_cols = st.session_state.loaded_preset
                st.info("✅ Using preset mapping")
                st.session_state.preset_applied = False
            else:
                auto_cols = find_columns(df)
            
            all_cols = list(df.columns)
            mandatory_keys = ['name', 'cost', 'qty']
            is_mapped = all(k in auto_cols for k in mandatory_keys)
            
            col1, col2, col3 = st.columns([2, 1, 1])
            with col1:
                if st.button("🔍 Auto-Detect All", use_container_width=True):
                    auto_cols = find_columns(df)
                    is_mapped = all(k in auto_cols for k in mandatory_keys)
                    st.session_state.auto_cols = auto_cols
            with col2:
                if is_mapped:
                    preset_name = st.text_input("Preset Name", placeholder="My Data", key="preset_name")
                    if preset_name and st.button("💾 Save", key="save_preset"):
                        mapping_to_save = {k: auto_cols.get(k, '') for k in ['name', 'cost', 'qty', 'date', 'order_id', 'phone']}
                        if save_mapping_preset(preset_name.strip(), mapping_to_save):
                            st.success(f"✅ Saved '{preset_name}'")
            with col3:
                if st.button("🔄 Reset", use_container_width=True):
                    auto_cols = {}
                    st.rerun()
            
            if is_mapped:
                st.success("✅ Detected: " + ", ".join([f"**{k}**: `{auto_cols[k]}`" for k in mandatory_keys if k in auto_cols]))
            else:
                st.warning("⚠️ Select required columns manually")
            
            mc1, mc2, mc3, mc4, mc5, mc6 = st.columns(6)
            
            def get_idx(key):
                return all_cols.index(auto_cols[key]) if key in auto_cols else 0
            def get_optional_idx(key):
                return all_cols.index(auto_cols[key]) + 1 if key in auto_cols else 0
            
            m_name = mc1.selectbox("Product Name", all_cols, index=get_idx('name'), help="Product names")
            m_cost = mc2.selectbox("Price", all_cols, index=get_idx('cost'), help="Product prices")
            m_qty = mc3.selectbox("Quantity", all_cols, index=get_idx('qty'), help="Quantities")
            m_date = mc4.selectbox("Date", ["None"] + all_cols, index=get_optional_idx('date'), help="Order dates")
            m_order = mc5.selectbox("Order ID", ["None"] + all_cols, index=get_optional_idx('order_id'), help="Order IDs")
            m_phone = mc6.selectbox("Phone", ["None"] + all_cols, index=get_optional_idx('phone'), help="Customer phone")
            
            matched_cols = set(auto_cols.values())
            unmatched = [c for c in all_cols if c not in matched_cols]
            if unmatched and len(unmatched) <= 5:
                st.info(f"ℹ️ Unmapped: {', '.join(unmatched)}")
            
            mapping = {
                'name': m_name, 'cost': m_cost, 'qty': m_qty,
                'date': m_date if m_date != "None" else None,
                'order_id': m_order if m_order != "None" else None,
                'phone': m_phone if m_phone != "None" else None
            }
            st.session_state.mapping = mapping
            st.session_state.auto_cols = auto_cols
            
            required_valid = all([m_name, m_cost, m_qty])
            
            st.markdown("---")
            if st.button("🚀 Generate Analytics", disabled=not required_valid, use_container_width=True):
                if not required_valid:
                    st.error("❌ Select all required columns")
                else:
                    with st.spinner("📊 Processing data..."):
                        try:
                            results = process_analytics(df, mapping)
                            st.session_state.results = results
                            st.session_state.filtered_df = df.copy()
                            st.balloons()
                            st.success("✅ Analytics generated!")
                        except Exception as e:
                            st.error(f"❌ Error: {e}")
                            log_event("CRASH", str(e))
            
            if st.session_state.results:
                results = st.session_state.results
                if st.session_state.filtered_df is not None:
                    filtered_df = render_filters(st.session_state.filtered_df, results)
                    if filtered_df is not None and len(filtered_df) > 0:
                        results = process_analytics(filtered_df, mapping)
                    else:
                        st.warning("⚠️ No data matches filters. Showing all.")
                        filtered_df = df.copy()
                        results = st.session_state.results
                else:
                    filtered_df = df.copy()
                
                st.markdown("---")
                st.subheader("📈 Key Metrics")
                m1, m2, m3, m4 = st.columns(4)
                m1.metric("Total Orders", f"{results['total_orders']:,.0f}")
                m2.metric("Units Sold", f"{results['total_qty']:,.0f}")
                m3.metric("Gross Revenue", f"TK {results['total_rev']:,.0f}")
                m4.metric("Avg Basket", f"TK {results['avg_basket_value']:,.0f}" if results['avg_basket_value'] > 0 else "TK 0")
                
                st.divider()
                st.subheader("📊 Visual Analytics")
                
                chart_prefs = st.session_state.get('chart_prefs', {'timeseries': True, 'pie': True, 'bar': True})
                
                if chart_prefs.get('timeseries', True) and results.get('timeseries') is not None and len(results['timeseries']) > 1:
                    v1, v2 = st.columns(2)
                    with v1:
                        st.markdown('<div class="chart-container">', unsafe_allow_html=True)
                        st.markdown("### 📈 Revenue Over Time")
                        ts_data = results['timeseries']
                        fig_ts = go.Figure()
                        fig_ts.add_trace(go.Scatter(x=pd.to_datetime(ts_data['Date']), y=ts_data['Revenue'], mode='lines+markers', name='Revenue', line=dict(color=DEEN_BRAND_COLOR, width=3), marker=dict(size=8), fill='tozeroy', fillcolor=f'rgba(0, 123, 255, 0.2)'))
                        fig_ts.update_layout(xaxis_title="Date", yaxis_title="Revenue (TK)", hovermode='x unified', template='plotly_white', height=400, showlegend=True)
                        st.plotly_chart(fig_ts, use_container_width=True)
                        st.markdown('</div>', unsafe_allow_html=True)
                    
                    if chart_prefs.get('pie', True):
                        with v2:
                            st.markdown('<div class="chart-container">', unsafe_allow_html=True)
                            st.markdown("### 🍩 Revenue by Category")
                            summ_sorted = results['summary'].sort_values('Total Amount', ascending=False)
                            colors = [get_category_color(cat) for cat in summ_sorted['Category']]
                            fig_pie = px.pie(summ_sorted, values='Total Amount', names='Category', hole=0.5, color_discrete_sequence=colors)
                            fig_pie.update_traces(textposition='inside', textinfo='percent+label', hovertemplate='<b>%{label}</b><br>Revenue: TK %{value:,.0f}<br>Share: %{percent}<extra></extra>', marker=dict(line=dict(color='white', width=2)))
                            fig_pie.update_layout(showlegend=True, legend=dict(orientation="v", yanchor="middle", y=0.5, xanchor="left", x=1.05), height=400)
                            st.plotly_chart(fig_pie, use_container_width=True)
                            st.markdown('</div>', unsafe_allow_html=True)
                else:
                    v1, v2 = st.columns(2)
                    if chart_prefs.get('pie', True):
                        with v1:
                            st.markdown('<div class="chart-container">', unsafe_allow_html=True)
                            st.markdown("### 🍩 Revenue by Category")
                            summ_sorted = results['summary'].sort_values('Total Amount', ascending=False)
                            colors = [get_category_color(cat) for cat in summ_sorted['Category']]
                            fig_pie = px.pie(summ_sorted, values='Total Amount', names='Category', hole=0.5, color_discrete_sequence=colors)
                            fig_pie.update_traces(textposition='inside', textinfo='percent+label', hovertemplate='<b>%{label}</b><br>Revenue: TK %{value:,.0f}<br>Share: %{percent}<extra></extra>')
                            fig_pie.update_layout(showlegend=True, height=400)
                            st.plotly_chart(fig_pie, use_container_width=True)
                            st.markdown('</div>', unsafe_allow_html=True)
                    if chart_prefs.get('bar', True):
                        with v2:
                            st.markdown('<div class="chart-container">', unsafe_allow_html=True)
                            st.markdown("### 📊 Volume by Category")
                            summ_sorted = results['summary'].sort_values('Total Qty', ascending=False)
                            colors = [get_category_color(cat) for cat in summ_sorted['Category']]
                            fig_bar = px.bar(summ_sorted, x='Category', y='Total Qty', color='Category', color_discrete_sequence=colors)
                            fig_bar.update_traces(hovertemplate='<b>%{x}</b><br>Quantity: %{y:,.0f}<extra></extra>')
                            fig_bar.update_layout(xaxis_title="", yaxis_title="Units Sold", showlegend=False, height=400, xaxis={'categoryorder':'total descending'})
                            st.plotly_chart(fig_bar, use_container_width=True)
                            st.markdown('</div>', unsafe_allow_html=True)
                
                if chart_prefs.get('bar', True) and not (chart_prefs.get('timeseries', True) and results.get('timeseries') is not None):
                    st.markdown('<div class="chart-container">', unsafe_allow_html=True)
                    st.markdown("### 📊 Volume by Category")
                    summ_sorted = results['summary'].sort_values('Total Qty', ascending=False)
                    colors = [get_category_color(cat) for cat in summ_sorted['Category']]
                    fig_bar = px.bar(summ_sorted, x='Category', y='Total Qty', color='Category', color_discrete_sequence=colors)
                    fig_bar.update_traces(hovertemplate='<b>%{x}</b><br>Quantity: %{y:,.0f}<extra></extra>')
                    fig_bar.update_layout(xaxis_title="Category", yaxis_title="Units Sold", showlegend=False, height=400, xaxis={'categoryorder':'total descending'})
                    st.plotly_chart(fig_bar, use_container_width=True)
                    st.markdown('</div>', unsafe_allow_html=True)
                
                st.divider()
                st.subheader("📋 Detailed Reports")
                
                st.markdown('<div class="card">', unsafe_allow_html=True)
                t1, t2 = st.tabs(["📊 Category Summary", "💰 Price-wise Breakdown"])
                df_breakdown = results['summary'].sort_values('Category', ascending=True).copy()
                total_qty = df_breakdown['Total Qty'].sum()
                total_amount = df_breakdown['Total Amount'].sum()
                df_breakdown.loc[len(df_breakdown) + 1] = ['TOTAL', total_qty, total_amount]
                df_breakdown.index = range(1, len(df_breakdown) + 1)
                df_drill = results['drilldown'].sort_values(['Category', 'Price'], ascending=[True, False]).copy()
                df_drill.columns = ['Category', 'Price', 'Qty', 'Total Amount']
                df_drill.loc[len(df_drill) + 1] = ['TOTAL', None, df_drill['Qty'].sum(), df_drill['Total Amount'].sum()]
                df_drill.index = range(1, len(df_drill) + 1)
                
                with t1:
                    st.dataframe(df_breakdown[['Category', 'Total Qty', 'Total Amount']], use_container_width=True, hide_index=True)
                with t2:
                    st.dataframe(df_drill[['Category', 'Price', 'Qty', 'Total Amount']], use_container_width=True, hide_index=True)
                st.markdown('</div>', unsafe_allow_html=True)
                
                st.markdown('<div class="card">', unsafe_allow_html=True)
                st.subheader("🏆 Top Products")
                top_n = st.slider("Show top", 5, 20, 10, key="top_n")
                top_items_display = results['top_items'].head(top_n)
                st.dataframe(top_items_display[['Product Name', 'Category', 'Total Qty', 'Total Amount']], use_container_width=True, hide_index=True)
                st.markdown('</div>', unsafe_allow_html=True)
                
                st.markdown("---")
                st.subheader("📥 Export")
                col1, col2, col3 = st.columns(3)
                with col1:
                    buf = BytesIO()
                    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
                        df_breakdown.to_excel(writer, sheet_name='Summary', index=False)
                        df_drill.to_excel(writer, sheet_name='Drilldown', index=False)
                        top_items_display.to_excel(writer, sheet_name='Top Products', index=False)
                        wb = writer.book
                        header_fill = PatternFill(start_color='007BFF', end_color='007BFF', fill_type='solid')
                        header_font = Font(bold=True, color='FFFFFF')
                        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                        for sheet_name in ['Summary', 'Drilldown', 'Top Products']:
                            ws = wb[sheet_name]
                            for cell in ws[1]:
                                cell.fill = header_fill
                                cell.font = header_font
                                cell.alignment = Alignment(horizontal='center', vertical='center')
                                cell.border = thin_border
                            ws.freeze_panes = 'A2'
                    st.download_button("📥 Excel Report", data=buf.getvalue(), file_name=f"Sales_{results['timeframe']}.xlsx", use_container_width=True)
                with col2:
                    csv_buf = BytesIO()
                    df_breakdown.to_csv(csv_buf, index=False)
                    st.download_button("📥 CSV", data=csv_buf.getvalue(), file_name=f"Summary_{results['timeframe']}.csv", use_container_width=True)
                with col3:
                    top_csv = BytesIO()
                    top_items_display.to_csv(top_csv, index=False)
                    st.download_button("📥 Top Products", data=top_csv.getvalue(), file_name=f"Top_{results['timeframe']}.csv", use_container_width=True)
                
        except Exception as e:
            st.error(f"❌ Error: {e}")
            log_event("UPLOAD_ERROR", str(e))

    else:
        if st.session_state.uploaded_file is None:
            render_empty_state()

    render_footer(logo_b64)


if __name__ == "__main__":
    main()
