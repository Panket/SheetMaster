import streamlit as st
import pandas as pd
import polars as pl
import os
import re
import io
import shutil
import time
from datetime import datetime

# Set page config
st.set_page_config(
    page_title="ExcelFusion - Premium Sheet Master",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Premium Styling
st.markdown("""
    <style>
        /* Custom font import */
        @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;600;800&family=Inter:wght@300;400;600;700&display=swap');
        
        html, body, [class*="css"] {
            font-family: 'Inter', sans-serif;
        }
        
        .header-title {
            font-family: 'Outfit', sans-serif;
            font-weight: 800;
            background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            font-size: 3rem;
            margin-bottom: 0px;
        }
        
        .header-subtitle {
            font-family: 'Inter', sans-serif;
            font-weight: 300;
            color: #555;
            font-size: 1.1rem;
            margin-top: 0px;
            margin-bottom: 25px;
        }
        
        .metric-card {
            background-color: white;
            padding: 18px;
            border-radius: 12px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.05);
            border: 1px solid #eef2f5;
            text-align: center;
        }
        
        .metric-label {
            font-size: 0.85rem;
            color: #888888;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        
        .metric-value {
            font-size: 1.6rem;
            color: #2c3e50;
            font-weight: 700;
            margin-top: 5px;
        }
        
        /* Highlight tab headers */
        .stTabs [data-baseweb="tab-list"] {
            gap: 8px;
        }
        
        .stTabs [data-baseweb="tab"] {
            background-color: #f1f3f5;
            border-radius: 8px 8px 0px 0px;
            padding: 8px 16px;
            font-weight: 600;
            color: #495057;
            border: none;
            transition: all 0.2s ease;
        }
        
        .stTabs [data-baseweb="tab"]:hover {
            background-color: #e9ecef;
            color: #1e3c72;
        }
        
        .stTabs [aria-selected="true"] {
            background-color: #1e3c72 !important;
            color: white !important;
        }
    </style>
""", unsafe_allow_html=True)

# ----------------------------------------------------
# Directory Setup & Helper Functions
# ----------------------------------------------------
UPLOAD_DIR = "uploads"
OUTPUT_DIR = "outputs"
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

def clear_temp_files():
    """Removes all temporary files and directories."""
    try:
        if os.path.exists(UPLOAD_DIR):
            shutil.rmtree(UPLOAD_DIR)
        if os.path.exists(OUTPUT_DIR):
            shutil.rmtree(OUTPUT_DIR)
        os.makedirs(UPLOAD_DIR, exist_ok=True)
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        return True
    except Exception as e:
        st.error(f"Error clearing directories: {e}")
        return False

# ----------------------------------------------------
# Session State Init
# ----------------------------------------------------
if "datasets" not in st.session_state:
    st.session_state["datasets"] = {}
if "active_df_name" not in st.session_state:
    st.session_state["active_df_name"] = None
if "history" not in st.session_state:
    st.session_state["history"] = []
if "filter_conditions" not in st.session_state:
    st.session_state["filter_conditions"] = []

# Helpers to manage history
def push_history(df, name, action="Edit"):
    if df is not None:
        if len(st.session_state["history"]) >= 10:
            st.session_state["history"].pop(0)
        st.session_state["history"].append({
            "df": df.copy(),
            "name": name,
            "action": action,
            "timestamp": datetime.now().strftime("%H:%M:%S")
        })

def update_active_dataset(df, name=None, action="Edit"):
    if name is None:
        name = st.session_state["active_df_name"]
    
    if name:
        push_history(st.session_state["datasets"].get(name), name, action)
        st.session_state["datasets"][name] = df
        st.session_state["active_df_name"] = name
        st.toast(f"✅ Data updated: {action}")

# File reading helper
def read_file(file_path, file_name, use_polars=False):
    _, ext = os.path.splitext(file_name.lower())
    try:
        if use_polars:
            if ext == ".csv":
                return pl.read_csv(file_path).to_pandas()
            elif ext in [".xlsx", ".xls"]:
                try:
                    return pl.read_excel(file_path).to_pandas()
                except Exception:
                    # Fallback to pandas
                    return pd.read_excel(file_path)
        else:
            if ext == ".csv":
                return pd.read_csv(file_path)
            elif ext in [".xlsx", ".xls"]:
                return pd.read_excel(file_path)
    except Exception as e:
        st.error(f"⚠️ Error reading file '{file_name}': {str(e)}")
        return None

# Export helpers
def convert_df_to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    return output.getvalue()

def convert_df_to_csv(df):
    return df.to_csv(index=False).encode('utf-8')

# Formula evaluator
def evaluate_formula(df, formula_str):
    cols_in_formula = re.findall(r'\[(.*?)\]', formula_str)
    
    for c in cols_in_formula:
        if c not in df.columns:
            raise ValueError(f"Column '{c}' not found in the dataset.")
            
    expr = formula_str
    # Replace column names from longest to shortest to prevent sub-string replacement issues
    for c in sorted(cols_in_formula, key=len, reverse=True):
        expr = expr.replace(f"[{c}]", f"df['{c}']")
    
    allowed_globals = {
        "__builtins__": None,
        "df": df,
        "pd": pd,
    }
    try:
        result = eval(expr, allowed_globals)
        return result
    except Exception as e:
        raise ValueError(f"Evaluation error: {str(e)}")

# Text standardization helper
def standardize_text_col(df, col, options):
    df_clean = df.copy()
    s = df_clean[col].astype(str)
    if "lowercase" in options:
        s = s.str.lower()
    if "strip" in options:
        s = s.str.strip()
    if "remove_special" in options:
        s = s.apply(lambda x: re.sub(r'[^a-zA-Z0-9\s\.\,\-\_]', '', x))
    if "remove_digits" in options:
        s = s.apply(lambda x: re.sub(r'\d', '', x))
    df_clean[col] = s
    return df_clean

# Outlier handler
def handle_outliers_iqr(df, col, action):
    q1 = df[col].quantile(0.25)
    q3 = df[col].quantile(0.75)
    iqr = q3 - q1
    lower_bound = q1 - 1.5 * iqr
    upper_bound = q3 + 1.5 * iqr
    
    outliers_mask = (df[col] < lower_bound) | (df[col] > upper_bound)
    num_outliers = outliers_mask.sum()
    
    df_clean = df.copy()
    if action == "Drop Outliers":
        df_clean = df_clean[~outliers_mask]
    elif action == "Cap Outliers":
        df_clean[col] = df_clean[col].clip(lower=lower_bound, upper=upper_bound)
        
    return df_clean, num_outliers, lower_bound, upper_bound

# ----------------------------------------------------
# Layout & Header
# ----------------------------------------------------
col_title, col_logo = st.columns([8, 2])
with col_title:
    st.markdown('<h1 class="header-title">📊 ExcelFusion</h1>', unsafe_allow_html=True)
    st.markdown('<p class="header-subtitle">Excel & CSV Wrangling Masterclass — Merge, Clean, and Edit seamlessly.</p>', unsafe_allow_html=True)

# ----------------------------------------------------
# Sidebar Controls
# ----------------------------------------------------
st.sidebar.markdown("### 📥 Upload Datasets")
uploaded_files = st.sidebar.file_uploader(
    "Upload Excel (.xlsx, .xls) or CSV files",
    type=["xlsx", "xls", "csv"],
    accept_multiple_files=True
)

st.sidebar.markdown("---")
st.sidebar.markdown("### ⚙️ Processing Engine")
engine_option = st.sidebar.toggle("Use Polars Engine (Faster for large files)", value=False)

# Load uploaded files into session state
if uploaded_files:
    for uploaded_file in uploaded_files:
        if uploaded_file.name not in st.session_state["datasets"]:
            with st.spinner(f"Reading {uploaded_file.name}..."):
                # Save to uploads folder
                file_path = os.path.join(UPLOAD_DIR, uploaded_file.name)
                with open(file_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())
                
                # Load DataFrame
                df = read_file(file_path, uploaded_file.name, use_polars=engine_option)
                if df is not None:
                    st.session_state["datasets"][uploaded_file.name] = df
                    if st.session_state["active_df_name"] is None:
                        st.session_state["active_df_name"] = uploaded_file.name
                        
# Active Dataset Selection
if st.session_state["datasets"]:
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 🎯 Active Dataset")
    
    dataset_keys = list(st.session_state["datasets"].keys())
    # Ensure active_df_name exists in keys, otherwise pick the first
    if st.session_state["active_df_name"] not in dataset_keys:
        st.session_state["active_df_name"] = dataset_keys[0]
        
    selected_active = st.sidebar.selectbox(
        "Choose file to work on:",
        options=dataset_keys,
        index=dataset_keys.index(st.session_state["active_df_name"])
    )
    st.session_state["active_df_name"] = selected_active
    active_df = st.session_state["datasets"][selected_active]
    
    # Show active dataset info cards
    st.sidebar.markdown("#### Stats Summary")
    
    rows, cols = active_df.shape
    missing_cells = active_df.isnull().sum().sum()
    total_cells = rows * cols
    missing_pct = (missing_cells / total_cells * 100) if total_cells > 0 else 0
    
    st.sidebar.markdown(f"""
        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 8px; margin-bottom: 12px;">
            <div class="metric-card">
                <div class="metric-label">Rows</div>
                <div class="metric-value">{rows:,}</div>
            </div>
            <div class="metric-card">
                <div class="metric-label">Columns</div>
                <div class="metric-value">{cols}</div>
            </div>
        </div>
        <div class="metric-card" style="margin-bottom: 12px;">
            <div class="metric-label">Missing Cells</div>
            <div class="metric-value">{missing_cells:,} <span style="font-size:0.9rem; font-weight:normal; color:#888;">({missing_pct:.1f}%)</span></div>
        </div>
    """, unsafe_allow_html=True)
    
    # Undo Action if history is available
    if st.session_state["history"]:
        st.sidebar.markdown("#### History Operations")
        if st.sidebar.button("↩️ Undo Last Change", use_container_width=True):
            pop_history()
            st.rerun()

st.sidebar.markdown("---")
if st.sidebar.button("🧹 Clear All Cache & Temp Files", type="secondary", use_container_width=True):
    clear_temp_files()
    st.session_state["datasets"] = {}
    st.session_state["active_df_name"] = None
    st.session_state["history"] = []
    st.session_state["filter_conditions"] = []
    st.sidebar.success("Cache and temporary files cleared!")
    time.sleep(1)
    st.rerun()

# ----------------------------------------------------
# Main UI Tabs
# ----------------------------------------------------
tab1, tab2, tab3, tab4 = st.tabs([
    "🔗 Merge & Combine", 
    "🔍 Dynamic Filtering", 
    "✏️ Edit Data", 
    "🧹 Advanced Cleaning"
])

# ----------------------------------------------------
# Tab 1: Merge & Combine
# ----------------------------------------------------
with tab1:
    st.markdown("### 🔗 Join or Stack Multiple Datasets")
    if not st.session_state["datasets"]:
        st.warning("Please upload one or more files in the sidebar to get started.")
    else:
        # File selector preview
        st.markdown("#### 📄 Available Datasets Overview")
        preview_cols = st.columns(min(len(st.session_state["datasets"]), 3))
        for idx, (name, df) in enumerate(st.session_state["datasets"].items()):
            col_ui = preview_cols[idx % 3]
            with col_ui:
                with st.expander(f"👁️ Preview: {name}", expanded=False):
                    st.write(f"Dimensions: {df.shape[0]} rows × {df.shape[1]} columns")
                    st.dataframe(df.head(5), use_container_width=True)

        st.markdown("---")
        combine_method = st.radio("Choose Combine Method:", ["Concatenate (Vertical Stack)", "Merge (Horizontal Join)"], horizontal=True)

        if combine_method == "Concatenate (Vertical Stack)":
            st.info("Concatenation stacks multiple datasets vertically matching columns with the same headers.")
            selected_files = st.multiselect(
                "Select files to concatenate (in ordering sequence):",
                options=list(st.session_state["datasets"].keys()),
                default=list(st.session_state["datasets"].keys())
            )
            
            col_con1, col_con2 = st.columns(2)
            with col_con1:
                ignore_index = st.checkbox("Ignore index (Reset indexes sequentially)", value=True)
            with col_con2:
                add_source_col = st.checkbox("Add file source column", value=True)
                
            if st.button("🚀 Process Concatenation", type="primary"):
                if len(selected_files) < 2:
                    st.error("Please select at least 2 datasets to concatenate.")
                else:
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    status_text.text("Initializing Concat Operations...")
                    progress_bar.progress(10)
                    
                    try:
                        dfs_to_concat = []
                        total_files = len(selected_files)
                        
                        for i, fname in enumerate(selected_files):
                            status_text.text(f"Processing {fname}...")
                            temp_df = st.session_state["datasets"][fname].copy()
                            if add_source_col:
                                temp_df["_source_file"] = fname
                            dfs_to_concat.append(temp_df)
                            progress_bar.progress(int(10 + (i + 1) / total_files * 70))
                        
                        status_text.text("Combining datasets...")
                        if engine_option:
                            # Use Polars for speed
                            pl_dfs = [pl.from_pandas(d) for d in dfs_to_concat]
                            # Polars needs same schema or diagonal join
                            result_pl = pl.concat(pl_dfs, how="diagonal")
                            combined_df = result_pl.to_pandas()
                        else:
                            combined_df = pd.concat(dfs_to_concat, ignore_index=ignore_index, sort=False)
                            
                        progress_bar.progress(100)
                        status_text.text("Concatenation successful!")
                        
                        # Store in session state as combined result
                        new_name = f"Concat_Result_{datetime.now().strftime('%H%M%S')}"
                        st.session_state["datasets"][new_name] = combined_df
                        st.session_state["active_df_name"] = new_name
                        
                        st.success(f"Successfully concatenated. Created: `{new_name}`")
                        
                    except Exception as e:
                        st.error(f"Failed to concatenate: {str(e)}")
                        
        else: # Merge/Join
            st.info("Merge performs a horizontal join of two tables based on matching key columns.")
            dataset_names = list(st.session_state["datasets"].keys())
            if len(dataset_names) < 2:
                st.warning("You need at least 2 datasets to perform a horizontal join.")
            else:
                col_left, col_right = st.columns(2)
                with col_left:
                    left_file = st.selectbox("Left Dataset (Base Table):", options=dataset_names, index=0)
                    left_df = st.session_state["datasets"][left_file]
                    left_keys = left_df.columns.tolist()
                
                with col_right:
                    right_file = st.selectbox("Right Dataset (Join Table):", options=dataset_names, index=1 if len(dataset_names) > 1 else 0)
                    right_df = st.session_state["datasets"][right_file]
                    right_keys = right_df.columns.tolist()

                col_join1, col_join2 = st.columns(2)
                with col_join1:
                    join_type = st.selectbox("Join Type:", ["inner", "left", "right", "outer"])
                with col_join2:
                    # Look for matching column names as default suggestion
                    common_cols = list(set(left_keys).intersection(set(right_keys)))
                    default_key = common_cols[0] if common_cols else left_keys[0]
                    join_key = st.selectbox("Join Column (Key):", options=left_keys, index=left_keys.index(default_key))
                    
                    if join_key not in right_keys:
                        st.warning(f"Warning: Join column '{join_key}' does not exist in the right dataset. Please ensure matching column names or values.")

                if st.button("🚀 Process Merge / Join", type="primary"):
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    status_text.text("Merging datasets...")
                    progress_bar.progress(30)
                    
                    try:
                        if engine_option:
                            pl_left = pl.from_pandas(left_df)
                            pl_right = pl.from_pandas(right_df)
                            # Map outer to full in Polars
                            pl_how = "full" if join_type == "outer" else join_type
                            merged_pl = pl_left.join(pl_right, on=join_key, how=pl_how)
                            merged_df = merged_pl.to_pandas()
                        else:
                            merged_df = pd.merge(left_df, right_df, on=join_key, how=join_type)
                            
                        progress_bar.progress(100)
                        status_text.text("Merge operation completed successfully!")
                        
                        new_name = f"Merge_Result_{datetime.now().strftime('%H%M%S')}"
                        st.session_state["datasets"][new_name] = merged_df
                        st.session_state["active_df_name"] = new_name
                        
                        st.success(f"Successfully joined. Created: `{new_name}`")
                    except Exception as e:
                        st.error(f"Failed to join datasets: {str(e)}")

        # Result preview & export
        if st.session_state["active_df_name"] and ("Result" in st.session_state["active_df_name"]):
            current_result_df = st.session_state["datasets"][st.session_state["active_df_name"]]
            st.markdown("#### 📈 Output Result Preview")
            st.dataframe(current_result_df.head(100), use_container_width=True)
            
            # Export Options
            col_exp1, col_exp2 = st.columns(2)
            with col_exp1:
                excel_data = convert_df_to_excel(current_result_df)
                st.download_button(
                    label="📥 Download Result as Excel (.xlsx)",
                    data=excel_data,
                    file_name=f"{st.session_state['active_df_name']}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            with col_exp2:
                csv_data = convert_df_to_csv(current_result_df)
                st.download_button(
                    label="📥 Download Result as CSV (.csv)",
                    data=csv_data,
                    file_name=f"{st.session_state['active_df_name']}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv",
                    use_container_width=True
                )

# ----------------------------------------------------
# Tab 2: Dynamic Filtering
# ----------------------------------------------------
with tab2:
    st.markdown("### 🔍 Filter and Query Data")
    if not st.session_state["datasets"]:
        st.warning("Please upload one or more files in the sidebar to get started.")
    else:
        df_name = st.session_state["active_df_name"]
        df = st.session_state["datasets"][df_name].copy()
        
        st.info(f"Filtering active dataset: **{df_name}** ({df.shape[0]} rows)")
        
        # Build filter conditions UI dynamically
        if st.button("➕ Add Filter Condition"):
            st.session_state["filter_conditions"].append({
                "column": df.columns[0],
                "operator": "equals",
                "value": ""
            })
            
        if st.session_state["filter_conditions"]:
            st.markdown("#### Set Filter Rules")
            
            # Use container for rule lines
            to_remove = []
            for idx, cond in enumerate(st.session_state["filter_conditions"]):
                r_col1, r_col2, r_col3, r_col4 = st.columns([3, 2, 4, 1])
                
                with r_col1:
                    cond["column"] = st.selectbox(
                        f"Column##{idx}",
                        options=df.columns,
                        index=df.columns.get_loc(cond["column"]) if cond["column"] in df.columns else 0,
                        label_visibility="collapsed"
                    )
                with r_col2:
                    ops = ["equals", "not equals", "contains", "greater than", "less than", "is in", "is null", "is not null"]
                    cond["operator"] = st.selectbox(
                        f"Operator##{idx}",
                        options=ops,
                        index=ops.index(cond["operator"]) if cond["operator"] in ops else 0,
                        label_visibility="collapsed"
                    )
                with r_col3:
                    # Provide text value or lists depending on operators
                    if cond["operator"] in ["is null", "is not null"]:
                        st.text("No value required")
                    elif cond["operator"] == "is in":
                        cond["value"] = st.text_input(
                            f"Value##{idx}",
                            value=str(cond["value"]),
                            placeholder="comma, separated, values",
                            label_visibility="collapsed"
                        )
                    else:
                        cond["value"] = st.text_input(
                            f"Value##{idx}",
                            value=str(cond["value"]),
                            placeholder="Enter filter keyword or number",
                            label_visibility="collapsed"
                        )
                with r_col4:
                    if st.button("❌", key=f"del_cond_{idx}"):
                        to_remove.append(idx)
                        
            # Apply deletes
            if to_remove:
                for r_idx in sorted(to_remove, reverse=True):
                    st.session_state["filter_conditions"].pop(r_idx)
                st.rerun()

            # Execute filters button
            col_act1, col_act2 = st.columns(2)
            with col_act1:
                apply_clicked = st.button("🔍 Run Filter", type="primary", use_container_width=True)
            with col_act2:
                if st.button("🧹 Clear All Filter Conditions", use_container_width=True):
                    st.session_state["filter_conditions"] = []
                    st.rerun()

            if apply_clicked:
                filtered_df = df.copy()
                
                try:
                    for cond in st.session_state["filter_conditions"]:
                        col = cond["column"]
                        op = cond["operator"]
                        val = cond["value"]
                        
                        if op == "equals":
                            # Auto cast val to column type if numeric
                            if pd.api.types.is_numeric_dtype(filtered_df[col]):
                                filtered_df = filtered_df[filtered_df[col] == float(val)]
                            else:
                                filtered_df = filtered_df[filtered_df[col].astype(str) == str(val)]
                        elif op == "not equals":
                            if pd.api.types.is_numeric_dtype(filtered_df[col]):
                                filtered_df = filtered_df[filtered_df[col] != float(val)]
                            else:
                                filtered_df = filtered_df[filtered_df[col].astype(str) != str(val)]
                        elif op == "contains":
                            filtered_df = filtered_df[filtered_df[col].astype(str).str.contains(str(val), case=False, na=False)]
                        elif op == "greater than":
                            filtered_df = filtered_df[filtered_df[col] > float(val)]
                        elif op == "less than":
                            filtered_df = filtered_df[filtered_df[col] < float(val)]
                        elif op == "is in":
                            vals_list = [v.strip() for v in val.split(",")]
                            if pd.api.types.is_numeric_dtype(filtered_df[col]):
                                vals_list = [float(v) for v in vals_list if v]
                            filtered_df = filtered_df[filtered_df[col].isin(vals_list)]
                        elif op == "is null":
                            filtered_df = filtered_df[filtered_df[col].isnull()]
                        elif op == "is not null":
                            filtered_df = filtered_df[filtered_df[col].notnull()]

                    st.success("Filtering complete!")
                    
                    # Display results comparisons
                    rows_before = df.shape[0]
                    rows_after = filtered_df.shape[0]
                    pct_rem = (rows_after / rows_before * 100) if rows_before > 0 else 0
                    
                    st.markdown(f"""
                        <div style="display: flex; gap: 20px; margin: 15px 0;">
                            <div class="metric-card" style="flex:1;">
                                <div class="metric-label">Original Rows</div>
                                <div class="metric-value">{rows_before:,}</div>
                            </div>
                            <div class="metric-card" style="flex:1; border-color:#2a5298;">
                                <div class="metric-label">Filtered Rows</div>
                                <div class="metric-value" style="color:#2a5298;">{rows_after:,} <span style="font-size:0.9rem; font-weight:normal; color:#888;">({pct_rem:.1f}%)</span></div>
                            </div>
                        </div>
                    """, unsafe_allow_html=True)
                    
                    st.dataframe(filtered_df.head(100), use_container_width=True)
                    
                    # Actions on filtered dataset
                    col_f1, col_f2, col_f3 = st.columns(3)
                    with col_f1:
                        if st.button("💾 Overwrite Active Dataset with Filtered Results"):
                            update_active_dataset(filtered_df, action="Filter Query")
                            st.rerun()
                    with col_f2:
                        excel_data = convert_df_to_excel(filtered_df)
                        st.download_button(
                            label="📥 Download Filtered File (.xlsx)",
                            data=excel_data,
                            file_name=f"Filtered_{df_name}_{datetime.now().strftime('%H%M%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                    with col_f3:
                        csv_data = convert_df_to_csv(filtered_df)
                        st.download_button(
                            label="📥 Download Filtered File (.csv)",
                            data=csv_data,
                            file_name=f"Filtered_{df_name}_{datetime.now().strftime('%H%M%S')}.csv",
                            mime="text/csv",
                            use_container_width=True
                        )
                        
                except Exception as e:
                    st.error(f"Failed to apply filters. Please verify numeric values or column choices. Detail: {str(e)}")
        else:
            st.markdown("No filter conditions added yet. Click **Add Filter Condition** above.")

# ----------------------------------------------------
# Tab 3: Edit Data
# ----------------------------------------------------
with tab3:
    st.markdown("### ✏️ Interactive Data & Schema Editor")
    if not st.session_state["datasets"]:
        st.warning("Please upload one or more files in the sidebar to get started.")
    else:
        df_name = st.session_state["active_df_name"]
        df = st.session_state["datasets"][df_name].copy()
        
        st.write(f"Editing Active File: **{df_name}**")
        
        edit_sub_tab1, edit_sub_tab2, edit_sub_tab3 = st.tabs([
            "➕ Add Column", 
            "➕ Add Row Manually", 
            "❌ Delete Row / Column"
        ])
        
        # 1. Add Column
        with edit_sub_tab1:
            st.subheader("Add a New Column")
            new_col_name = st.text_input("New Column Name:", placeholder="e.g. TotalPrice")
            col_val_mode = st.radio("Choose Column Assignment Value:", ["Static Value", "Formula / Expression"], horizontal=True)
            
            if col_val_mode == "Static Value":
                col_static_type = st.selectbox("Static Value Type:", ["Text", "Integer", "Float", "Boolean"])
                if col_static_type == "Text":
                    static_val = st.text_input("Enter text:", value="")
                elif col_static_type == "Integer":
                    static_val = st.number_input("Enter integer:", step=1, value=0)
                elif col_static_type == "Float":
                    static_val = st.number_input("Enter float:", step=0.1, value=0.0)
                else:
                    static_val = st.checkbox("Boolean Value", value=False)
            else:
                st.info("""
                    **Formula syntax:** Reference columns inside square brackets.
                    - Multiply columns: `[Qty] * [UnitPrice]`
                    - Add text columns: `[FirstName] + ' ' + [LastName]`
                    - Apply tax multiplier: `[Price] * 1.07`
                """)
                static_val = st.text_input("Enter formula:", placeholder="[Qty] * [Price]")
                
            if st.button("🚀 Add Column", key="btn_add_col"):
                if not new_col_name:
                    st.error("Please enter a name for the new column.")
                elif new_col_name in df.columns:
                    st.error(f"Column '{new_col_name}' already exists.")
                else:
                    try:
                        if col_val_mode == "Static Value":
                            df[new_col_name] = static_val
                        else:
                            # Formula calculation
                            res = evaluate_formula(df, static_val)
                            df[new_col_name] = res
                            
                        update_active_dataset(df, action=f"Added column {new_col_name}")
                        st.success(f"Column '{new_col_name}' added successfully!")
                        time.sleep(0.5)
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error executing column assignment: {str(e)}")

        # 2. Add Row
        with edit_sub_tab2:
            st.subheader("Add a New Row Manually")
            # Build Form dynamically based on active dataframe schema
            with st.form("add_row_form_tab3", clear_on_submit=True):
                new_row_inputs = {}
                col_layout = st.columns(3)
                
                for i, col_name in enumerate(df.columns):
                    col_ui = col_layout[i % 3]
                    col_type = df[col_name].dtype
                    
                    if pd.api.types.is_numeric_dtype(col_type):
                        if pd.api.types.is_integer_dtype(col_type):
                            new_row_inputs[col_name] = col_ui.number_input(f"{col_name} (Int)", step=1, value=0)
                        else:
                            new_row_inputs[col_name] = col_ui.number_input(f"{col_name} (Float)", step=0.01, value=0.0)
                    elif pd.api.types.is_datetime64_any_dtype(col_type):
                        new_row_inputs[col_name] = pd.to_datetime(col_ui.date_input(f"{col_name} (Date)"))
                    elif pd.api.types.is_bool_dtype(col_type):
                        new_row_inputs[col_name] = col_ui.checkbox(f"{col_name} (Bool)", value=False)
                    else:
                        new_row_inputs[col_name] = col_ui.text_input(f"{col_name} (Text)", value="")
                
                submit_row = st.form_submit_button("Append Row to End")
                if submit_row:
                    try:
                        new_row_df = pd.DataFrame([new_row_inputs])
                        # Match types
                        for c in df.columns:
                            try:
                                new_row_df[c] = new_row_df[c].astype(df[c].dtype)
                            except Exception:
                                pass
                        df = pd.concat([df, new_row_df], ignore_index=True)
                        update_active_dataset(df, action="Manually added a row")
                        st.success("Row appended successfully!")
                        time.sleep(0.5)
                        st.rerun()
                    except Exception as e:
                        st.error(f"Failed to add row: {str(e)}")
                        
        # 3. Delete Rows/Columns
        with edit_sub_tab3:
            st.subheader("Remove Rows or Columns")
            
            col_del1, col_del2 = st.columns(2)
            with col_del1:
                st.markdown("#### Delete Columns")
                cols_to_delete = st.multiselect("Select columns to drop:", options=df.columns)
                if st.button("❌ Drop Selected Columns", type="primary"):
                    if not cols_to_delete:
                        st.warning("No columns selected.")
                    else:
                        df = df.drop(columns=cols_to_delete)
                        update_active_dataset(df, action=f"Dropped columns {cols_to_delete}")
                        st.success(f"Successfully dropped columns: {cols_to_delete}")
                        time.sleep(0.5)
                        st.rerun()
                        
            with col_del2:
                st.markdown("#### Delete Rows")
                row_indices_str = st.text_input("Enter Row Index (or indices comma-separated, e.g., 0, 4, 10):")
                if st.button("❌ Drop Selected Rows", type="primary"):
                    if not row_indices_str:
                        st.warning("Please input at least one index.")
                    else:
                        try:
                            indices = [int(idx.strip()) for idx in row_indices_str.split(",") if idx.strip().isdigit()]
                            if not indices:
                                st.error("No valid numeric index input found.")
                            else:
                                # Filter indexes that are within range
                                valid_indices = [idx for idx in indices if idx in df.index]
                                if not valid_indices:
                                    st.error("Selected indices not found in dataframe index.")
                                else:
                                    df = df.drop(index=valid_indices).reset_index(drop=True)
                                    update_active_dataset(df, action=f"Dropped rows: {valid_indices}")
                                    st.success(f"Successfully dropped rows: {valid_indices}")
                                    time.sleep(0.5)
                                    st.rerun()
                        except Exception as e:
                            st.error(f"Failed to drop rows: {str(e)}")

        st.markdown("---")
        st.markdown("#### Preview Current Active Dataset")
        st.dataframe(df.head(100), use_container_width=True)

# ----------------------------------------------------
# Tab 4: Advanced Data Cleaning
# ----------------------------------------------------
with tab4:
    st.markdown("### 🧹 Advanced Data Cleaning Tools")
    if not st.session_state["datasets"]:
        st.warning("Please upload one or more files in the sidebar to get started.")
    else:
        df_name = st.session_state["active_df_name"]
        df = st.session_state["datasets"][df_name].copy()
        
        st.write(f"Cleaning Active File: **{df_name}**")
        
        clean_col1, clean_col2 = st.columns([1, 1])
        
        with clean_col1:
            # 1. Duplicates Remover
            with st.expander("👥 Handle Duplicates", expanded=True):
                dedup_cols = st.multiselect(
                    "Check duplicates based on specific columns (Leave empty to check entire row):",
                    options=df.columns
                )
                keep_pos = st.selectbox("Which duplicate occurrence to keep:", ["first", "last", "none"])
                
                if st.button("🧹 Remove Duplicates", use_container_width=True):
                    before_cnt = len(df)
                    keep_val = False if keep_pos == "none" else keep_pos
                    subset = dedup_cols if dedup_cols else None
                    df = df.drop_duplicates(subset=subset, keep=keep_val)
                    after_cnt = len(df)
                    removed = before_cnt - after_cnt
                    
                    update_active_dataset(df, action=f"Removed {removed} duplicates")
                    st.success(f"Removed {removed} duplicate rows. Current size: {after_cnt} rows.")
                    time.sleep(0.5)
                    st.rerun()

            # 2. Text Standardization
            with st.expander("🔤 Text Standardization", expanded=False):
                text_cols = [c for c in df.columns if pd.api.types.is_string_dtype(df[c]) or df[c].dtype == object]
                if not text_cols:
                    st.warning("No text/string columns detected.")
                else:
                    target_txt_col = st.selectbox("Select text column to clean:", options=text_cols)
                    std_options = st.multiselect(
                        "Select standardization options:",
                        options=[
                            "Convert to Lowercase",
                            "Trim Whitespaces (Strip)",
                            "Remove Special Characters",
                            "Remove Digits / Numbers"
                        ],
                        default=["Trim Whitespaces (Strip)"]
                    )
                    
                    if st.button("🧹 Standardize Text", use_container_width=True):
                        if not std_options:
                            st.warning("Please select at least one standardisation option.")
                        else:
                            clean_opts = []
                            if "Convert to Lowercase" in std_options:
                                clean_opts.append("lowercase")
                            if "Trim Whitespaces (Strip)" in std_options:
                                clean_opts.append("strip")
                            if "Remove Special Characters" in std_options:
                                clean_opts.append("remove_special")
                            if "Remove Digits / Numbers" in std_options:
                                clean_opts.append("remove_digits")
                            df = standardize_text_col(df, target_txt_col, clean_opts)
                            update_active_dataset(df, action=f"Standardized text column: {target_txt_col}")
                            st.success(f"Successfully cleaned '{target_txt_col}' text column!")
                            time.sleep(0.5)
                            st.rerun()

            # 3. Data Type Casting
            with st.expander("🔄 Convert Column Data Types", expanded=False):
                cast_col = st.selectbox("Select column to cast:", options=df.columns)
                target_type = st.selectbox("Target Data Type:", ["Integer", "Float", "String", "DateTime"])
                
                if st.button("🔄 Cast Data Type", use_container_width=True):
                    try:
                        if target_type == "Integer":
                            # fill na first or it will throw error
                            df[cast_col] = pd.to_numeric(df[cast_col], errors='coerce').fillna(0).astype(int)
                        elif target_type == "Float":
                            df[cast_col] = pd.to_numeric(df[cast_col], errors='coerce').astype(float)
                        elif target_type == "String":
                            df[cast_col] = df[cast_col].astype(str)
                        elif target_type == "DateTime":
                            df[cast_col] = pd.to_datetime(df[cast_col], errors='coerce')
                            
                        update_active_dataset(df, action=f"Cast type of {cast_col} to {target_type}")
                        st.success(f"Column '{cast_col}' casted to {target_type} successfully!")
                        time.sleep(0.5)
                        st.rerun()
                    except Exception as e:
                        st.error(f"Casting error: {str(e)}")
                        
        with clean_col2:
            # 4. Handle Missing Values
            with st.expander("❓ Missing Values Imputation", expanded=True):
                na_col = st.selectbox("Select column with missing values:", options=df.columns)
                null_count = df[na_col].isnull().sum()
                st.write(f"Nulls detected in **{na_col}**: **{null_count}** rows")
                
                impute_strategy = st.selectbox(
                    "Choose Strategy:",
                    ["Drop Rows with Nulls", "Fill with Mean (Numeric only)", "Fill with Median (Numeric only)", "Fill with Mode", "Fill with Custom Constant"]
                )
                
                constant_fill_val = ""
                if impute_strategy == "Fill with Custom Constant":
                    constant_fill_val = st.text_input("Enter fill value:")
                    
                if st.button("🧹 Clean Missing Values", use_container_width=True):
                    if null_count == 0:
                        st.info("No missing values in selected column.")
                    else:
                        try:
                            if impute_strategy == "Drop Rows with Nulls":
                                df = df.dropna(subset=[na_col])
                            elif impute_strategy == "Fill with Mean (Numeric only)":
                                if not pd.api.types.is_numeric_dtype(df[na_col]):
                                    st.error("Mean imputation is only valid for numerical columns.")
                                    st.stop()
                                df[na_col] = df[na_col].fillna(df[na_col].mean())
                            elif impute_strategy == "Fill with Median (Numeric only)":
                                if not pd.api.types.is_numeric_dtype(df[na_col]):
                                    st.error("Median imputation is only valid for numerical columns.")
                                    st.stop()
                                df[na_col] = df[na_col].fillna(df[na_col].median())
                            elif impute_strategy == "Fill with Mode":
                                mode_val = df[na_col].mode()
                                if len(mode_val) > 0:
                                    df[na_col] = df[na_col].fillna(mode_val[0])
                                else:
                                    st.warning("Could not calculate mode.")
                            elif impute_strategy == "Fill with Custom Constant":
                                # Attempt numeric conversion if applicable
                                if pd.api.types.is_numeric_dtype(df[na_col]):
                                    try:
                                        fill_v = float(constant_fill_val) if "." in constant_fill_val else int(constant_fill_val)
                                    except ValueError:
                                        fill_v = constant_fill_val
                                else:
                                    fill_v = constant_fill_val
                                df[na_col] = df[na_col].fillna(fill_v)
                                
                            update_active_dataset(df, action=f"Imputed missing in {na_col}")
                            st.success(f"Successfully cleaned missing values in '{na_col}'!")
                            time.sleep(0.5)
                            st.rerun()
                        except Exception as e:
                            st.error(f"Error resolving nulls: {str(e)}")

            # 5. Outlier Detection (IQR)
            with st.expander("📈 Outlier Detection & IQR Cleaning", expanded=False):
                numeric_cols = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]
                if not numeric_cols:
                    st.warning("No numeric columns found for IQR calculations.")
                else:
                    target_num_col = st.selectbox("Select numeric column:", options=numeric_cols)
                    outlier_action = st.selectbox("Choose action on outliers:", ["Detect Only", "Drop Outliers", "Cap Outliers"])
                    
                    # Compute stats without modifying
                    temp_df_clean, detected_outliers, lb, ub = handle_outliers_iqr(df, target_num_col, "Detect Only")
                    st.markdown(f"""
                        * **Detected Outliers**: {detected_outliers}
                        * **Lower Bound (Q1 - 1.5*IQR)**: {lb:.2f}
                        * **Upper Bound (Q3 + 1.5*IQR)**: {ub:.2f}
                    """)
                    
                    if st.button("🧹 Execute Outlier Clean", use_container_width=True):
                        df, det_cnt, lb, ub = handle_outliers_iqr(df, target_num_col, outlier_action)
                        update_active_dataset(df, action=f"Outlier handling ({outlier_action}) on {target_num_col}")
                        st.success(f"Success! {outlier_action} action processed. {det_cnt} outliers handled.")
                        time.sleep(0.5)
                        st.rerun()

        st.markdown("---")
        st.markdown("#### Preview Current Cleaned Active Dataset")
        st.dataframe(df.head(100), use_container_width=True)
