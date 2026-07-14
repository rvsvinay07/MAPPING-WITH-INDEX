import streamlit as st
import pandas as pd
from thefuzz import process, fuzz
import io
import os
import json

# Import the pre-mapped expert funds
try:
    from pre_mapped_dict import PRE_MAPPED_FUNDS
except ImportError:
    PRE_MAPPED_FUNDS = {}

# Set page config with high-quality title and logo emoji
st.set_page_config(
    page_title="AlphaNifty - Fund to Index Matcher",
    page_icon="🎯",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Inject Premium Custom CSS for Sleek Dark Theme, Glassmorphism, and Gold Accents
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;600;800&family=Playfair+Display:ital,wght@0,600;0,800;1,600&display=swap');
    
    /* Global Styles */
    .stApp {
        background: linear-gradient(135deg, #0e1117 0%, #1a1e29 100%);
        color: #f0f2f6;
        font-family: 'Outfit', sans-serif;
    }
    
    /* Title and Header Styles */
    h1, h2, h3 {
        font-family: 'Outfit', sans-serif;
        font-weight: 800;
        background: linear-gradient(90deg, #dfb857 0%, #b8860b 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        letter-spacing: -0.5px;
    }
    
    /* Cards (Glassmorphism effect) */
    .premium-card {
        background: rgba(255, 255, 255, 0.03);
        border: 1px solid rgba(223, 184, 87, 0.15);
        border-radius: 20px;
        padding: 30px;
        box-shadow: 0 8px 32px 0 rgba(0, 0, 0, 0.37);
        backdrop-filter: blur(12px);
        -webkit-backdrop-filter: blur(12px);
        margin-bottom: 25px;
        transition: transform 0.3s ease, border-color 0.3s ease;
    }
    .premium-card:hover {
        transform: translateY(-2px);
        border-color: rgba(223, 184, 87, 0.3);
    }
    
    /* KPI Metric Cards */
    .kpi-container {
        display: flex;
        justify-content: space-between;
        gap: 15px;
        margin-bottom: 25px;
    }
    .kpi-card {
        flex: 1;
        background: rgba(255, 255, 255, 0.02);
        border-left: 4px solid #dfb857;
        border-radius: 8px;
        padding: 15px 20px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.2);
    }
    .kpi-title {
        font-size: 13px;
        text-transform: uppercase;
        letter-spacing: 1px;
        color: #a0aec0;
        margin-bottom: 5px;
    }
    .kpi-value {
        font-size: 26px;
        font-weight: 800;
        color: #dfb857;
    }
    
    /* Interactive Button Customization */
    div.stButton > button {
        background: linear-gradient(90deg, #dfb857 0%, #b8860b 100%) !important;
        color: #0e1117 !important;
        font-family: 'Outfit', sans-serif !important;
        font-weight: 700 !important;
        border: none !important;
        border-radius: 30px !important;
        padding: 12px 35px !important;
        font-size: 16px !important;
        box-shadow: 0 4px 15px rgba(223, 184, 87, 0.3) !important;
        transition: all 0.3s ease !important;
    }
    div.stButton > button:hover {
        transform: translateY(-2px) scale(1.02) !important;
        box-shadow: 0 6px 20px rgba(223, 184, 87, 0.5) !important;
    }
    
    /* File Uploader Customization */
    .uploadedFile {
        background-color: rgba(223, 184, 87, 0.05) !important;
        border: 1px dashed rgba(223, 184, 87, 0.3) !important;
        border-radius: 12px !important;
    }
    
    /* Tab Styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background-color: rgba(255, 255, 255, 0.02);
        padding: 8px;
        border-radius: 15px;
        border: 1px solid rgba(255, 255, 255, 0.05);
    }
    .stTabs [data-baseweb="tab"] {
        font-family: 'Outfit', sans-serif !important;
        font-weight: 600 !important;
        border-radius: 10px !important;
        padding: 10px 24px !important;
        color: #a0aec0 !important;
        transition: all 0.3s ease !important;
        background-color: transparent !important;
        border: none !important;
    }
    .stTabs [data-baseweb="tab"]:hover {
        color: #dfb857 !important;
        background-color: rgba(255, 255, 255, 0.02) !important;
    }
    .stTabs [aria-selected="true"] {
        background-color: rgba(223, 184, 87, 0.12) !important;
        color: #dfb857 !important;
        border: 1px solid rgba(223, 184, 87, 0.25) !important;
    }
</style>
""", unsafe_allow_html=True)


# --- DATA LOADING FUNCTIONS ---

@st.cache_data
def load_index_database(file_name="VS_indices line names.xlsx"):
    """Loads benchmark index names from the columns of VS_indices line names.xlsx."""
    if os.path.exists(file_name):
        try:
            df = pd.read_excel(file_name)
            # Filter out DATE and clean index names
            indices = [str(col).strip() for col in df.columns if str(col).upper() != "DATE"]
            return indices
        except Exception as e:
            st.error(f"Error reading index Excel file: {e}")
    
    # Fallback list of common indices if file is missing
    return [
        "NIFTY 50", "NIFTY NEXT 50", "NIFTY 100", "NIFTY 200", "Nifty Total Market",
        "NIFTY 500", "NIFTY MIDCAP 150", "NIFTY MIDCAP 50", "NIFTY SMALLCAP 250",
        "NIFTY SMALLCAP 50", "NIFTY SMALLCAP 100", "NIFTY MICROCAP 250",
        "NIFTY LargeMidcap 250", "NIFTY MIDSMALLCAP 400", "NIFTY AUTO", "NIFTY BANK",
        "NIFTY FINANCIAL SERVICES", "NIFTY FMCG", "Nifty HEALTHCARE", "NIFTY IT",
        "NIFTY PHARMA", "NIFTY REALTY", "NIFTY SERVICES SECTOR", "NIFTY ENERGY",
        "NIFTY INFRASTRUCTURE", "NIFTY MNC", "NIFTY PSE", "NIFTY100 ESG",
        "Nifty EV & New Age Automotive", "Nifty India Defence", "Nifty India Digital",
        "NIFTY INDIA CONSUMPTION", "NIFTY INDIA NEW AGE CONSUMPTION"
    ]

# Mappings of exact index codes matching the client's detailed mapping structure
CUSTOM_INDEX_CODES = {
    "nifty india consumption": "T_Nifty Consumption",
    "nifty 50": "B_Nifty 50",
    "nifty200 momentum 30": "ST_Nifty200Momentum30",
    "nifty financial services": "S_Nifty Fin Service",
    "nifty fmcg": "S_Nifty FMCG",
    "nifty 200": "B_Nifty 200",
    "nifty alpha low volatility 30": "ST_NIFTY ALPHALOWVOL",
    "nifty healthcare": "S_NIFTY500 HEALTH",
    "nifty dividend opportunities 50": "ST_Nifty Div Opps 50",
    "nifty india digital": "T_NIFTY IND DIGITAL",
    "nifty it": "S_Nifty IT",
    "nifty bank": "S_Nifty Bank",
    "nifty private bank": "S_Nifty Pvt Bank",
    "nifty psu bank": "S_Nifty PSU Bank",
    "nifty auto": "S_Nifty Auto",
    "nifty media": "S_Nifty Media",
    "nifty metal": "S_Nifty Metal",
    "nifty pharma": "S_Nifty Pharma",
    "nifty realty": "S_Nifty Realty",
    "nifty cpse": "T_Nifty CPSE",
    "nifty energy": "T_Nifty Energy",
    "nifty housing": "T_Nifty Housing",
    "nifty mnc": "T_Nifty MNC",
    "nifty pse": "T_Nifty PSE",
    "nifty services sector": "T_Nifty Services",
    "nifty infrastructure": "T_Nifty Infra",
    "nifty next 50": "B_Nifty Next 50",
    "nifty 100": "B_Nifty 100",
    "nifty 500": "B_Nifty 500",
    "nifty midcap 150": "B_Nifty Midcap 150",
    "nifty smallcap 250": "B_Nifty Smallcap 250",
    "nifty microcap 250": "B_Nifty Microcap 250",
    "nifty largemidcap 250": "B_Nifty LargeMidcap 250",
    "nifty balanced advantage": "B_AN_Balanced Advantage",
    "nifty gold": "B_AN_Gold",
    "nifty tax saver": "B_AN_Tax Saver",
    "nifty 10y g-sec": "B_AN_10Y G-Sec",
    "nifty midsmallcap400 momentum quality 100": "ST_NiftyMS400 MQ 100",
    "nifty midsmall it & telecom": "S_Nfty MS IT Telcm",
    "nifty financial services ex bank": "S_Nifty FinSerExBnk",
    "nifty alpha low-volatility 30": "ST_NIFTY ALPHALOWVOL",
    "nifty smallcap250 momentum quality 100": "ST_NftySml250MQ 100",
    "nifty200 alpha 30": "ST_Nifty200 Alpha 30"
}

def get_index_details(matched_name, all_indices):
    """Determines the category and code of a matched index based on its name and slice position."""
    matched_name_lower = matched_name.lower().strip()
    
    # Check custom code overrides first
    if matched_name_lower in CUSTOM_INDEX_CODES:
        code = CUSTOM_INDEX_CODES[matched_name_lower]
        if code.startswith("B_"):
            category = "Broad based"
        elif code.startswith("S_"):
            category = "Sectoral Indices"
        elif code.startswith("ST_"):
            category = "Strategic Indices"
        elif code.startswith("T_"):
            category = "Thematic Indices"
        else:
            category = "Broad based"
        return category, code, matched_name

    # Otherwise determine by index column position in VS_indices sheet
    try:
        idx_lower = [x.lower() for x in all_indices]
        pos = idx_lower.index(matched_name_lower)
        original_name = all_indices[pos]
    except ValueError:
        # Default fallback
        return "Broad based", f"B_{matched_name}", matched_name

    if pos < 25:
        category = "Broad based"
        code = f"B_{original_name}"
    elif pos < 45:
        category = "Sectoral Indices"
        code = f"S_{original_name}"
    elif pos < 81:
        category = "Strategic Indices"
        code = f"ST_{original_name}"
    else:
        category = "Thematic Indices"
        code = f"T_{original_name}"

    return category, code, original_name


import re

def normalize_fund_name(name):
    """Normalize whitespace and clean up name strings for robust matching."""
    if not name:
        return ""
    name_str = str(name).strip().lower()
    # Replace '_x000d_' with space (often added by excel exports for line breaks)
    name_str = name_str.replace('_x000d_', ' ')
    # Collapse multiple spaces/newlines/tabs to a single space
    name_str = re.sub(r'\s+', ' ', name_str)
    return name_str.strip()

# Build a normalized dictionary for exact lookups
NORMALIZED_PRE_MAPPED_FUNDS = {}
for key, val in PRE_MAPPED_FUNDS.items():
    norm_key = normalize_fund_name(key)
    if norm_key:
        NORMALIZED_PRE_MAPPED_FUNDS[norm_key] = val


def check_keyword_constraints(fund_name, matched_index_name):
    """
    Returns True if the matched_index_name satisfies the keyword constraints in fund_name.
    Prevents mismatched index mappings (e.g. IT/pharma/banking mismatches).
    """
    fund_name_lower = fund_name.lower()
    idx_name_lower = matched_index_name.lower()
    
    # 1. Tech / IT / Digital
    if 'tech' in fund_name_lower or 'it' in fund_name_lower.split() or fund_name_lower.endswith('it') or 'digital' in fund_name_lower:
        # Candidate index must be related to IT/Digital/Telecom
        if not any(x in idx_name_lower for x in ['it', 'digital', 'telecom']):
            return False
            
    # 2. Banking / Financial / Finance / Private Bank / PSU Bank
    if 'bank' in fund_name_lower or 'financial' in fund_name_lower or 'fin ' in fund_name_lower or 'fln' in fund_name_lower:
        if not any(x in idx_name_lower for x in ['bank', 'financial', 'fin ', 'finser']):
            return False
            
    # 3. Consumption / FMCG
    if 'consum' in fund_name_lower or 'fmcg' in fund_name_lower:
        if not any(x in idx_name_lower for x in ['consumption', 'fmcg']):
            return False
            
    # 4. Pharma / Healthcare / PHD
    if 'pharma' in fund_name_lower or 'health' in fund_name_lower or 'phd' in fund_name_lower.split():
        if not any(x in idx_name_lower for x in ['pharma', 'health']):
            return False
            
    # 5. Infrastructure
    if 'infra' in fund_name_lower:
        if 'infra' not in idx_name_lower:
            return False
            
    # 6. Gold
    if 'gold' in fund_name_lower:
        if 'gold' not in idx_name_lower:
            return False
            
    # 7. ESG
    if 'esg' in fund_name_lower:
        if 'esg' not in idx_name_lower:
            return False
            
    # 8. Momentum
    if 'momentum' in fund_name_lower:
        if 'momentum' not in idx_name_lower:
            return False
            
    # 9. Smallcap
    if 'small' in fund_name_lower:
        if not any(x in idx_name_lower for x in ['small', 'sml']):
            if '500' not in idx_name_lower and 'total market' not in idx_name_lower:
                return False
                
    # 10. Midcap
    if 'mid' in fund_name_lower:
        if not any(x in idx_name_lower for x in ['mid', 'm150', 'ms400', 'midsmall']):
            if '500' not in idx_name_lower and 'total market' not in idx_name_lower:
                return False

    return True


def match_fund(fund_name, all_indices, pre_mapped_funds):
    """Hybrid matcher combining pre-mapped expert funds lookup with fuzzy matching."""
    fund_name_clean = normalize_fund_name(fund_name)
    
    if not fund_name_clean:
        return "", "", "", 0, "Unknown"

    # Step 1: Direct exact match in expert pre-mapped funds
    if fund_name_clean in NORMALIZED_PRE_MAPPED_FUNDS:
        mapped_info = NORMALIZED_PRE_MAPPED_FUNDS[fund_name_clean]
        idx_name = mapped_info["matched_index_name"]
        code = mapped_info.get("code")
        category = mapped_info.get("category")
        
        # Standardize category labels
        if category:
            cat_l = category.lower()
            if cat_l.startswith("hybrid") or cat_l.startswith("debt") or cat_l.startswith("retirement"):
                category = "Broad based"
            elif "sectoral" in cat_l or "sectorial" in cat_l:
                category = "Sectoral Indices"
            elif "thematic" in cat_l:
                category = "Thematic Indices"
            elif "strategic" in cat_l:
                category = "Strategic Indices"
        
        if not code or not category or idx_name.lower().strip() in CUSTOM_INDEX_CODES:
            category, code, idx_name = get_index_details(idx_name, all_indices)
            
        return idx_name, code, category, 100, "High (Exact Match)"

    # Step 2: Fuzzy match in expert pre-mapped funds list
    pre_mapped_names = list(NORMALIZED_PRE_MAPPED_FUNDS.keys())
    best_pre_mapped, score_pre_mapped = process.extractOne(
        fund_name_clean, pre_mapped_names, scorer=fuzz.WRatio
    )
    
    if score_pre_mapped >= 92:
        mapped_info = NORMALIZED_PRE_MAPPED_FUNDS[best_pre_mapped]
        idx_name = mapped_info["matched_index_name"]
        code = mapped_info.get("code")
        category = mapped_info.get("category")
        
        # Standardize categories
        if category:
            cat_l = category.lower()
            if cat_l.startswith("hybrid") or cat_l.startswith("debt") or cat_l.startswith("retirement"):
                category = "Broad based"
            elif "sectoral" in cat_l or "sectorial" in cat_l:
                category = "Sectoral Indices"
            elif "thematic" in cat_l:
                category = "Thematic Indices"
            elif "strategic" in cat_l:
                category = "Strategic Indices"
                
        if not code or not category or idx_name.lower().strip() in CUSTOM_INDEX_CODES:
            category, code, idx_name = get_index_details(idx_name, all_indices)
            
        return idx_name, code, category, score_pre_mapped, f"High (Pre-mapped - {score_pre_mapped}%)"

    # Step 3: General fuzzy match against the 125 benchmark indices
    matches = process.extract(
        fund_name, all_indices, scorer=fuzz.WRatio, limit=10
    )
    
    best_match_original = None
    best_score = 0
    
    # Pick the best match that satisfies keyword constraints
    for match, score in matches:
        if check_keyword_constraints(fund_name, match):
            best_match_original = match
            best_score = score
            break
            
    # Fallback to the top candidate if none satisfy constraints
    if not best_match_original and matches:
        best_match_original = matches[0][0]
        best_score = matches[0][1]
        
    category, code, best_match_original = get_index_details(best_match_original, all_indices)
    
    if best_score >= 85:
        confidence = "High"
    elif best_score >= 60:
        confidence = "Medium"
    else:
        confidence = "Low"
        
    return best_match_original, code, category, best_score, f"{confidence} ({best_score}%)"


# --- APP INTERFACE SETUP ---

# Load data sources
all_indices = load_index_database()

# Header Brand and Styling
st.write("")
col_logo1, col_logo2, col_logo3 = st.columns([1, 2, 1])
with col_logo2:
    if os.path.exists("logo.png"):
        st.image("logo.png", use_container_width=True)
    elif os.path.exists("logo.avif"):
        st.image("logo.avif", use_container_width=True)
    else:
        st.markdown("<h1 style='text-align: center; font-size: 42px;'>🎯 ALPHANIFTY</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; font-size: 18px; color: #a0aec0; margin-top: -15px; text-align: center;'>Portfolio Diagnostic Engine - Benchmark Index Mapping</p>", unsafe_allow_html=True)
st.write("")

# Create Navigation Tabs
tab1, tab2 = st.tabs(["🔎 Single Fund Matcher", "📂 Bulk Portfolio Matcher"])

# --- TAB 1: SINGLE FUND MATCHER ---
with tab1:
    st.markdown('<div class="premium-card">', unsafe_allow_html=True)
    st.markdown("### Match a Mutual Fund")
    st.write("Type or paste any mutual fund name below to discover its closest benchmark index and category mapping.")
    
    fund_input = st.text_input(
        "Enter Mutual Fund Name:",
        placeholder="e.g., Bandhan Nifty Alpha 50 Index Fund",
        key="single_fund"
    )
    
    if fund_input:
        st.write("")
        with st.spinner("Finding best benchmark index..."):
            idx_name, code, category, score, conf_label = match_fund(fund_input, all_indices, PRE_MAPPED_FUNDS)
            
        col1, col2 = st.columns(2)
        with col1:
            st.markdown(f"""
            <div class="kpi-card" style="border-left-color: #dfb857;">
                <div class="kpi-title">Matched Benchmark Index</div>
                <div class="kpi-value" style="font-size: 22px; margin-top: 5px;">{idx_name}</div>
            </div>
            """, unsafe_allow_html=True)
            
            st.write("")
            st.markdown(f"""
            <div class="kpi-card" style="border-left-color: #4299e1;">
                <div class="kpi-title">Index Category</div>
                <div class="kpi-value" style="font-size: 22px; margin-top: 5px; color: #4299e1;">{category}</div>
            </div>
            """, unsafe_allow_html=True)
            
        with col2:
            st.markdown(f"""
            <div class="kpi-card" style="border-left-color: #38a169;">
                <div class="kpi-title">Index Code</div>
                <div class="kpi-value" style="font-size: 22px; margin-top: 5px; color: #38a169;">{code}</div>
            </div>
            """, unsafe_allow_html=True)
            
            st.write("")
            
            color = "#38a169" if "High" in conf_label else ("#dd6b20" if "Medium" in conf_label else "#e53e3e")
            st.markdown(f"""
            <div class="kpi-card" style="border-left-color: {color};">
                <div class="kpi-title">Match Confidence</div>
                <div class="kpi-value" style="font-size: 22px; margin-top: 5px; color: {color};">{conf_label}</div>
            </div>
            """, unsafe_allow_html=True)
            
    st.markdown('</div>', unsafe_allow_html=True)


# --- TAB 2: BULK PORTFOLIO MATCHER (NEW EXTENSION) ---
with tab2:
    st.markdown('<div class="premium-card">', unsafe_allow_html=True)
    st.markdown("### Bulk Portfolio Matcher")
    st.write("Upload an Excel file (.xlsx, .xls) or a CSV file containing your clients' scheme names and invested amounts. The engine will match them to benchmark indices and generate a downloadable report.")
    
    uploaded_file = st.file_uploader(
        "Upload Portfolio Excel or CSV:",
        type=["xlsx", "xls", "csv"],
        key="portfolio_file"
    )
    
    if uploaded_file:
        try:
            # Load the file into a DataFrame
            file_extension = os.path.splitext(uploaded_file.name)[-1].lower()
            if file_extension == '.csv':
                df_raw = pd.read_csv(uploaded_file)
            else:
                df_raw = pd.read_excel(uploaded_file)
                
            st.success(f"Successfully loaded file: {uploaded_file.name} ({len(df_raw)} rows)")
            
            # --- Column Selection & Mapping ---
            # Try to auto-detect fund and amount columns
            auto_fund_col = None
            auto_amount_col = None
            for col in df_raw.columns:
                col_clean = str(col).strip().lower()
                if any(x in col_clean for x in ["fund", "scheme", "portfolio", "name"]):
                    if auto_fund_col is None:
                        auto_fund_col = col
                if any(x in col_clean for x in ["amount", "value", "inr", "allocation", "size"]):
                    if auto_amount_col is None:
                        auto_amount_col = col
            
            st.write("")
            col_map1, col_map2 = st.columns(2)
            with col_map1:
                fund_col_selected = st.selectbox(
                    "Select Fund Name Column:",
                    df_raw.columns,
                    index=list(df_raw.columns).index(auto_fund_col) if auto_fund_col in df_raw.columns else 0
                )
            with col_map2:
                amount_col_selected = st.selectbox(
                    "Select Amount Column:",
                    df_raw.columns,
                    index=list(df_raw.columns).index(auto_amount_col) if auto_amount_col in df_raw.columns else 0
                )
            
            # Match Button
            st.write("")
            run_match = st.button("🔥 Map Portfolio Benchmark Indices")
            
            if run_match:
                # Check for NaNs and clean data
                df_work = df_raw.dropna(subset=[fund_col_selected]).copy()
                df_work[amount_col_selected] = pd.to_numeric(df_work[amount_col_selected].fillna(0), errors='coerce').fillna(0)
                
                # Perform the matching row-by-row
                results = []
                progress_bar = st.progress(0)
                total_rows = len(df_work)
                
                for idx, (_, row) in enumerate(df_work.iterrows()):
                    fund_name = str(row[fund_col_selected]).strip()
                    amount = float(row[amount_col_selected])
                    
                    matched_idx, index_code, category, score, conf_label = match_fund(fund_name, all_indices, PRE_MAPPED_FUNDS)
                    
                    confidence_val = "Low"
                    if score >= 85:
                        confidence_val = "High"
                    elif score >= 60:
                        confidence_val = "Medium"
                        
                    results.append({
                        "Original Scheme/Fund Name": fund_name,
                        "Mapped Index Name": matched_idx,
                        "Index Code": index_code,
                        "Category": category,
                        "Amount (INR)": amount,
                        "Match Score (%)": score,
                        "Match Confidence": confidence_val
                    })
                    
                    progress_bar.progress((idx + 1) / total_rows)
                
                progress_bar.empty()
                df_results = pd.DataFrame(results)
                
                # Calculate allocation percentage
                total_amount = df_results["Amount (INR)"].sum()
                if total_amount > 0:
                    df_results["Allocation (%)"] = (df_results["Amount (INR)"] / total_amount) * 100
                else:
                    df_results["Allocation (%)"] = 0.0
                    
                df_results["Allocation (%)"] = df_results["Allocation (%)"].round(2)
                
                # Add Amount and Percentages aliases for compatibility
                df_results["Amount"] = df_results["Amount (INR)"]
                df_results["Percentages"] = df_results["Allocation (%)"]
                
                # Rearrange columns to match standard template and include user's requested columns
                df_results.insert(0, '#', range(1, len(df_results) + 1))
                cols_order = [
                    '#', 'Original Scheme/Fund Name', 'Mapped Index Name', 'Index Code', 
                    'Category', 'Amount (INR)', 'Allocation (%)', 'Amount', 'Percentages', 'Match Confidence', 'Match Score (%)'
                ]
                df_results = df_results[cols_order]
                
                # --- Display KPI Summary Cards ---
                st.write("")
                st.subheader("📊 Portfolio Diagnostics Summary")
                st.markdown(f"""
                <div class="kpi-container">
                    <div class="kpi-card" style="border-left-color: #dfb857;">
                        <div class="kpi-title">Total Portfolio Value</div>
                        <div class="kpi-value">₹ {total_amount:,.2f}</div>
                    </div>
                    <div class="kpi-card" style="border-left-color: #4299e1;">
                        <div class="kpi-title">Total Funds Matched</div>
                        <div class="kpi-value">{len(df_results)}</div>
                    </div>
                    <div class="kpi-card" style="border-left-color: #38a169;">
                        <div class="kpi-title">High Confidence Matches</div>
                        <div class="kpi-value">{len(df_results[df_results["Match Confidence"] == "High"])} / {len(df_results)}</div>
                    </div>
                </div>
                """, unsafe_allow_html=True)
                
                # --- Visualizations ---
                col_chart1, col_chart2 = st.columns(2)
                with col_chart1:
                    st.write("#### 🍰 Asset Class Category Allocation")
                    category_alloc = df_results.groupby("Category")["Amount (INR)"].sum().reset_index()
                    st.bar_chart(data=category_alloc, x="Category", y="Amount (INR)", color="#dfb857")
                    
                with col_chart2:
                    st.write("#### 📈 Top Matched Index Allocation")
                    index_alloc = df_results.groupby("Mapped Index Name")["Amount (INR)"].sum().reset_index()
                    index_alloc = index_alloc.sort_values(by="Amount (INR)", ascending=False).head(10)
                    st.bar_chart(data=index_alloc, x="Mapped Index Name", y="Amount (INR)", color="#4299e1")
                
                # --- Preview Table ---
                st.write("")
                st.write("#### 📋 Matched Results Preview")
                st.dataframe(
                    df_results.style.format({
                        "Amount (INR)": "₹{:,.2f}",
                        "Allocation (%)": "{:.2f}%",
                        "Amount": "₹{:,.2f}",
                        "Percentages": "{:.2f}%",
                        "Match Score (%)": "{:.0f}%"
                    }),
                    use_container_width=True
                )
                
                # --- Export & Download Excel Utility ---
                st.write("")
                # Generate styled excel in memory
                excel_buffer = io.BytesIO()
                with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                    df_results.to_excel(writer, index=False, sheet_name='Detailed Mapping')
                    
                    # Style spreadsheet layout slightly using openpyxl
                    workbook = writer.book
                    worksheet = writer.sheets['Detailed Mapping']
                    
                    # Auto-fit columns
                    for col in worksheet.columns:
                        max_len = max(len(str(cell.value or '')) for cell in col)
                        col_letter = col[0].column_letter
                        worksheet.column_dimensions[col_letter].width = max(max_len + 3, 12)
                
                excel_data = excel_buffer.getvalue()
                
                # Download Button
                st.download_button(
                    label="📥 Download Matched Excel Report",
                    data=excel_data,
                    file_name=f"Matched_Benchmark_Indices_{uploaded_file.name.split('.')[0]}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
        except Exception as e:
            st.error(f"Error processing file: {e}")
            st.exception(e)
            
    st.markdown('</div>', unsafe_allow_html=True)
