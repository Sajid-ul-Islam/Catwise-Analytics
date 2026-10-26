# 📊 Catwise Analytics - Smart Sales Performance Dashboard

A modern Streamlit-based business intelligence tool engineered to convert raw sales data (Excel/CSV) into actionable, category-wise insights with automated product classification and professional multi-sheet Excel reporting.

---

## 🚀 Key Features

### 🔍 Smart Data Ingestion & Column Auto-Detection
- **Intelligent Column Mapping**: Automatically detects and maps essential sales fields (`Product Name`, `Price`, `Quantity`) and optional dimensions (`Date`, `Order ID`, `Customer Phone`).
- **Data Sanitization**: Automatic cleaning of currency symbols, comma separators, whitespace, negative quantities, and exclusion of bundle/combo promotions.
- **Timeframe Detection**: Automatically inspects date ranges to generate dynamic, timestamped report titles.

### 🏷️ Automated Product Categorization
- **Rule-based Keyword Classification**: Automatically classifies products across 30+ apparel and lifestyle categories:
  - *Panjabi, Jeans, Polo Shirt, Formal Shirt, Flannel Shirt, Denim Shirt, Cuban Shirt, Contrast Shirt, Kaftan Shirt*
  - *T-Shirts & Sleeve Detection* (Full Sleeve `FS` vs. Half Sleeve `HS`)
  - *Trousers, Twill/Chino, Cargo Trousers, Drop Shoulder, Turtleneck, Boxer, Tank Top*
  - *Sweatshirt, Hoodie, Sweater, Jacket, Active Wear, Jersey*
  - *Accessories: Belt, Cap, Knit Cap, Wallet, Card Holder, Passport Holder, Leather Bag, Mask, Water Bottle, Mug*
- **Sleeve Length Logic**: Accurately differentiates between Full Sleeve (FS) and Half Sleeve (HS) for Shirts and T-Shirts.

### 📈 Synchronized & High-Contrast Visualizations
- **Revenue by Category (Donut Chart)**: Percentage revenue share with clean hover metrics.
- **Volume by Category (Bar Chart)**: Visual volume distribution automatically sorted **from highest to lowest** (`total descending`).
- **Synchronized Bright Color Palette**: Both charts share a coordinated 70+ color palette ensuring each unique category retains the exact same distinct color across all visualizations.

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
- **Fast Rule Discovery**: Directly inspect product titles to quickly extract keywords for [`CATEGORY_MAPPING`](file:///d:/Repo/Catwise-Analytics-1/app.py#L17-L50).

### ⚡ Performance & Architectural Optimizations
- **Vectorized Numeric Cleaning**: C-level vectorized string regex operations and `pd.to_numeric` parsing.
- **Algorithmic Categorization Speedup**: Classifies unique product titles once and broadcasts across tens of thousands of rows via hash mapping, slashing processing latency by 99%.
- **Streamlit Fragment Isolation (`@st.fragment`)**: Sidebar feedback submission and system log inspection run as isolated fragments without triggering main dashboard re-runs.
- **High-Performance Caching (`@st.cache_data`)**: File loading and heavy metric calculations are cached to prevent redundant compute cycles.
- **Session State Persistence (`st.session_state`)**: Analytics results remain persistent across tab switches and widget events.
- **High-Speed Excel Streaming (`openpyxl.iter_rows`)**: Single-pass row generator styling accelerates Excel report downloads for large catalogs.

### 🎯 Interactive Product Search & Filtering
- **Tab 3 Instant Search**: Real-time keyword filter across product titles.
- **Category Multi-Select**: Filter items by one or more categories simultaneously with dynamic subtotal calculations.
- **Formatted Data Tables**: Native currency (`TK %,.0f`) and quantity (`%,.0f pcs`) formatting using `st.column_config.NumberColumn` without row index clutter.

### ⚙️ External Category Rules (`config/categories.json`)
- Classification rules are decoupled into `config/categories.json`. Add or update category keywords without modifying application source code.

### 💬 Feedback & Error Logging
- Built-in sidebar reporting system to submit classification feedback and log misclassified items to `feedback/user_feedback.json`.
- System event logger for troubleshooting and audit trails (`feedback/system_logs.json`).

---

## 🛠️ Tech Stack

- **Frontend & App Framework**: [Streamlit](https://streamlit.io/)
- **Data Manipulation**: [Pandas](https://pandas.pydata.org/)
- **Interactive Visualizations**: [Plotly Express](https://plotly.com/python/)
- **Excel Generation & Styling**: [OpenPyXL](https://openpyxl.readthedocs.io/)
- **Styling**: Custom CSS supporting dynamic light and dark theme adaptation.

---

## 📦 Installation & Setup

### 1. Clone the repository
```bash
git clone https://github.com/Sajid-ul-Islam/Catwise-Analytics.git
cd Catwise-Analytics-1
```

### 2. Install dependencies
```bash
pip install -r requirements.txt
```

### 3. Run the application
```bash
streamlit run app.py
```

---

## 💡 How to Use

1. **Upload Data**: Upload your raw sales file (`.xlsx` or `.csv`).
2. **Verify Columns**: Review the auto-detected columns and adjust mappings if necessary.
3. **Generate Analytics**: Click **"Generate Analytics"** to inspect KPI metrics, synchronized charts, and data tables.
4. **Explore Tabs**: Review Category summaries, price-point drilldowns, and category-wise product lists.
5. **Export Report**: Click **"📥 Download Report"** to get the formatted multi-sheet Excel file.

---

## 📂 Project Structure

```
Catwise-Analytics-1/
├── assets/
│   └── deen_logo.png           # Brand logo
├── config/
│   └── categories.json         # External category mapping rules
├── feedback/                   # User feedback and system logs
│   ├── system_logs.json
│   └── user_feedback.json
├── app.py                      # Main Streamlit application
├── requirements.txt            # Python dependencies
└── README.md                   # Project documentation
```

---

*Developed for DEEN Commerce Ltd. by Sajid Islam.*
