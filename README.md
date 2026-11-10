# 📊 Catwise Analytics - Smart Sales Performance Dashboard

A modern Streamlit-based business intelligence application engineered to transform raw sales data (Excel/CSV) into actionable, category-wise insights with automated product classification, session-persistent mapping presets, robust input validation, and professional multi-sheet Excel reporting.

---

## 🚀 Key Features

### 🔍 Smart Data Ingestion & Single Data Preview
- **Raw Data Preview**: Seamless, single collapsible preview expander displaying the first 10 rows and dimensions of uploaded datasets.
- **Intelligent Column Auto-Detection**: Priority-based heuristic mapping for essential sales fields (`Product Name`, `Price`, `Quantity`) and optional dimensions (`COGS`, `Date`, `Order ID`, `Customer Phone`).
- **Encoding Fallback Engine**: Multi-encoding CSV parser supporting `utf-8`, `utf-8-sig`, `cp1252`, `latin1`, and `iso-8859-1` to handle exports from varied platforms without decoding errors.
- **Data Sanitization**: Vectorized cleaning of currency symbols (`TK`, `Tk`, `$`, `BDT`), comma separators, negative quantities, parentheses accounting formats `(250)`, and automatic exclusion of bundle/combo promotions.

### 💾 Session State & User Preferences Management
- **Persistent Widget State**: Column mappings remain intact across tab changes, search queries, category filter multiselects, and sidebar interactions.
- **Preset Management System**: Save, apply, and delete column mapping presets (stored in `presets/`) for recurring vendor or platform formats (Shopify, WooCommerce, POS exports).
- **Default Preference Memory**: Automatically remembers the last successfully executed mapping and applies it when similar files are uploaded.
- **One-Click Re-Detection**: Dedicated **"🔄 Auto-Detect Columns"** button to restore intelligent defaults at any time.

### 🛡️ Proactive Input Validation & Error Handling
- **Pre-Execution Validation**: Validates that all mandatory fields (`Product Name`, `Unit Price`, `Quantity`) are selected and present in the uploaded dataset before execution.
- **KeyError Prevention**: The "Generate Analytics" button is disabled until valid required columns are mapped, preventing runtime `KeyError` crashes.
- **Integrity Warnings**: Real-time alerts for duplicate column assignments and non-numeric price/quantity selections.
- **Multi-Tier Error Recovery**: Graceful handling of empty files, malformed structures, division by zero, and corrupted spreadsheets with descriptive UI feedback and expandable developer logs.

### 🏷️ Automated Product Categorization
- **Rule-Based Keyword Classification**: Classifies products across 30+ apparel and lifestyle categories:
  - *Panjabi, Jeans, Polo Shirt, Formal Shirt, Flannel Shirt, Denim Shirt, Cuban Shirt, Contrast Shirt, Kaftan Shirt*
  - *T-Shirts & Sleeve Detection* (Full Sleeve `FS` vs. Half Sleeve `HS`)
  - *Trousers, Twill/Chino, Cargo Trousers, Drop Shoulder, Turtleneck, Boxer, Tank Top*
  - *Sweatshirt, Hoodie, Sweater, Jacket, Active Wear, Jersey*
  - *Accessories: Belt, Cap, Knit Cap, Wallet, Card Holder, Passport Holder, Leather Bag, Mask, Water Bottle, Mug*
- **Sleeve Length Logic**: Accurately differentiates between Full Sleeve (FS) and Half Sleeve (HS) for Shirts and T-Shirts using keyword boundary matching.

### 📈 Synchronized & High-Contrast Visualizations
- **Revenue by Category (Donut Chart)**: Percentage revenue share with clean hover metrics and theme-adaptive separators.
- **Volume by Category (Bar Chart)**: Visual volume distribution automatically sorted from highest to lowest (`total descending`).
- **Synchronized Palette System**: 70+ vibrant color palette ensuring each unique category retains the exact same distinct color across all visualizations.

### 📋 Deep-Dive Analytics & Multi-Tab Views
1. **📊 Category Summary**: High-level aggregated units sold and gross revenue per category.
2. **💰 Price-wise Category**: Price-point breakdown within each category to identify top revenue-driving price tiers.
3. **📦 Category-wise Product**: Granular product-level performance with category association, units sold, and gross revenue.
- All views feature bold, highlighted summary total rows.

### 📥 Professional Multi-Sheet Excel Export
Generates a polished `.xlsx` workbook via `openpyxl` with:
- **3 Comprehensive Sheets**: `Category Summary`, `Price-wise Category`, and `Category-wise Product`.
- **Corporate Styling**: High-contrast blue header rows (`#007BFF`) with bold white text.
- **Freeze Panes**: Top header rows frozen across all sheets for seamless scrolling.
- **Auto-Fitted Columns**: Column widths automatically computed based on text content length.
- **Currency & Integer Formatting**: Formatted with standard comma separators (`#,##0`) without cluttering decimals.
- **Total Rows Highlighting**: Summary rows automatically styled with a soft gray background (`#ECECEC`) and bold text.
- **Clean Alignment**: Text columns left-aligned; numerical quantities and revenues right-aligned.

### ⚠️ Unmapped Items Inspector ("Others" Review)
- **Zero-Loss Tracking**: Dedicated collapsible audit tool displaying all products categorized under `'Others'`.
- **Revenue & Unit Impact**: Quantifies the exact gross revenue and unit volume tied to unclassified products to help prioritize category rule updates.
- **Fast Rule Discovery**: Directly inspect product titles to quickly extract keywords for `config/categories.json`.

### ⚡ Performance & Architectural Optimizations
- **Vectorized Numeric Cleaning**: C-level vectorized string regex operations and `pd.to_numeric` parsing.
- **Algorithmic Categorization Speedup**: Classifies unique product titles once and broadcasts across tens of thousands of rows via hash mapping, slashing processing latency by 99%.
- **Streamlit Fragment Isolation (`@st.fragment`)**: Sidebar feedback submission and system log inspection run as isolated fragments without triggering main dashboard re-runs.
- **High-Performance Caching (`@st.cache_data`)**: File loading and heavy metric calculations are cached to prevent redundant compute cycles.
- **High-Speed Excel Streaming (`openpyxl.iter_rows`)**: Single-pass row generator styling accelerates Excel report downloads for large catalogs.

---

## 🛠️ Tech Stack

- **Frontend & App Framework**: [Streamlit](https://streamlit.io/)
- **Data Manipulation**: [Pandas](https://pandas.pydata.org/), [NumPy](https://numpy.org/)
- **Interactive Visualizations**: [Plotly Express](https://plotly.com/python/)
- **Excel Generation & Styling**: [OpenPyXL](https://openpyxl.readthedocs.io/)
- **Styling**: Theme-adaptive CSS supporting dynamic light and dark theme adaptation.

---

## 📦 Installation & Setup

### 1. Clone the repository
```bash
git clone https://github.com/Sajid-ul-Islam/Catwise-Analytics.git
cd Catwise-Analytics
```

### 2. Install dependencies
```bash
pip install -r requirements.txt
```

### 3. Run the application
```bash
streamlit run app.py
```

### 4. Run automated test suite
```bash
python test_app_fixes.py
```

---

## 💡 How to Use

1. **Upload Data**: Upload your raw sales file (`.xlsx` or `.csv`).
2. **Preview Data**: Expand **"🔍 Preview Data"** to verify data format and initial rows.
3. **Verify Columns**: Review the auto-detected columns or apply a saved preset from **"💾 Presets & Preferences"**.
4. **Generate Analytics**: Click **"Generate Analytics"** to inspect KPI metrics, synchronized charts, and data tables.
5. **Explore Tabs**: Review Category summaries, price-point drilldowns, and category-wise product lists.
6. **Export Report**: Click **"📥 Download Report"** to get the formatted multi-sheet Excel file.

---

## 📂 Project Structure

```
Catwise-Analytics/
├── assets/
│   └── deen_logo.png           # Brand logo
├── config/
│   └── categories.json         # External category mapping rules
├── feedback/                   # User feedback and system logs
│   ├── system_logs.json
│   └── user_feedback.json
├── presets/                    # Saved column mapping presets (.json)
├── app.py                      # Main Streamlit application
├── test_app_fixes.py           # Automated unit test suite
├── requirements.txt            # Python dependencies
├── .gitignore                  # Git ignore rules
└── README.md                   # Project documentation
```

---

*Developed for DEEN Commerce Ltd. by Sajid Islam.*
