import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime

import os

st.set_page_config(page_title="MF Shorts Creator", layout="wide")

def load_data(uploaded_file):
    if uploaded_file is not None:
        try:
            # Try different headers to handle different file formats
            for h in [0, 4, 3, 2, 1]:
                if uploaded_file.name.endswith('.xls'):
                    temp_df = pd.read_excel(uploaded_file, engine='xlrd', header=h)
                else:
                    temp_df = pd.read_excel(uploaded_file, engine='openpyxl', header=h)
                
                if 'Scheme Name' in temp_df.columns:
                    return temp_df
            
            # If no 'Scheme Name' found, return with header 0
            uploaded_file.seek(0)
            if uploaded_file.name.endswith('.xls'):
                return pd.read_excel(uploaded_file, engine='xlrd')
            else:
                return pd.read_excel(uploaded_file, engine='openpyxl')
        except Exception as e:
            st.error(f"Error loading file: {e}")
            return None
    return None

def format_currency(val):
    if pd.isna(val):
        return "-"
    try:
        return f"{int(val):,}"
    except:
        return val

def format_percentage(val):
    if pd.isna(val):
        return "-"
    try:
        return f"{val:.2f}%"
    except:
        return val

def clean_scheme_name(name):
    if not isinstance(name, str):
        return name
    
    # Remove common suffixes
    # Order matters: longer ones first
    suffixes = ["Reg Gr", "Gr Gr", "Dir Gr", "Reg IDCW", "Dir IDCW", "Reg", "Gr", "Growth", "Direct"]
    
    for _ in range(3): # Try removing up to 3 suffixes
        found = False
        for suffix in suffixes:
            if name.endswith(" " + suffix) or name.endswith("-" + suffix):
                name = name[:-len(suffix)-1].strip()
                found = True
                break
            elif name.endswith(suffix):
                name = name[:-len(suffix)].strip()
                found = True
                break
        if not found:
            break
            
    # If it doesn't end with "Fund", append " Fund"
    if not name.lower().endswith(" fund"):
        name += " Fund"
        
    return name

def display_table_with_copy(df, table_id):
    """Display a DataFrame as an HTML table with custom copy button to avoid shortcut conflicts"""
    # Convert DataFrame to HTML table
    html_table = df.to_html(index=False, border=1, classes='dataframe')
    
    # Create copy button with JavaScript that copies the table
    copy_button_js = f'''
        <div style="margin-bottom: 10px;">
            <button onclick="copyTable_{table_id}()" 
                    style="background-color: #0066cc; color: white; padding: 8px 16px; border: none; 
                           border-radius: 4px; cursor: pointer; font-weight: bold;">
                📋 Copy to Clipboard
            </button>
        </div>
        <div id="table_{table_id}" style="overflow-x: auto;">
            {html_table}
        </div>
        <script>
            function copyTable_{table_id}() {{
                var table = document.getElementById('table_{table_id}');
                var range = document.createRange();
                range.selectNode(table);
                window.getSelection().removeAllRanges();
                window.getSelection().addRange(range);
                try {{
                    var successful = document.execCommand('copy');
                    if(successful) {{
                        alert('✅ Data copied to clipboard successfully!');
                    }}
                }} catch (err) {{
                    alert('❌ Failed to copy data');
                }}
                window.getSelection().removeAllRanges();
            }}
        </script>
    '''
    st.markdown(copy_button_js, unsafe_allow_html=True)

st.title("📊 Mutual Fund Shorts Data Tool")

# Sidebar for file upload and navigation
with st.sidebar:
    st.header("Settings")
    uploaded_file = st.file_uploader("Upload Mutual Fund Excel File", type=["xls", "xlsx"])
    
    persistent_file = "last_updated_sip_data.xls"
    
    # Logic to keep the last updated file
    if uploaded_file is not None:
        # Save the uploaded file locally to persist it
        with open(persistent_file, "wb") as f:
            f.write(uploaded_file.getbuffer())
        st.success("File uploaded and saved as default!")
        df = load_data(uploaded_file)
    else:
        # If no file is uploaded, try the last saved persistent file first
        if os.path.exists(persistent_file):
            try:
                for h in [4, 0, 3, 2, 1]:
                    temp_df = pd.read_excel(persistent_file, engine='xlrd', header=h)
                    if 'Scheme Name' in temp_df.columns:
                        df = temp_df
                        st.info("Using last updated data.")
                        break
            except:
                df = None
        
        # Fallback to original default files if no persistent file or if loading failed
        if df is None:
            default_files = ["SIP Returns (7).xls", "SIP Returns (6).xls"]
            for f in default_files:
                if os.path.exists(f):
                    try:
                        for h in [4, 0, 3, 2, 1]:
                            temp_df = pd.read_excel(f, engine='xlrd', header=h)
                            if 'Scheme Name' in temp_df.columns:
                                df = temp_df
                                st.info(f"Using default file: {f}")
                                break
                        if df is not None: break
                    except:
                        continue
        
        if df is None:
            st.warning("Please upload an Excel file.")

    if df is not None:
        sip_amount = st.number_input("Enter Monthly SIP Amount (₹)", min_value=500, value=10000, step=500)
        view_mode = st.radio("Select View Mode", ["Single Fund View", "Top Funds View"])

if df is not None:
    # Clean up column names (remove extra spaces)
    df.columns = [c.strip() if isinstance(c, str) else c for c in df.columns]
    
    # Clean up scheme names
    if 'Scheme Name' in df.columns:
        df['Scheme Name'] = df['Scheme Name'].apply(clean_scheme_name)
    
    if view_mode == "Single Fund View":
        st.header("🎯 Single Fund Analysis")
        
        fund_name = st.selectbox("Select Fund Name", df['Scheme Name'].unique())
        fund_data = df[df['Scheme Name'] == fund_name].iloc[0]
        
        # Dynamically detect available durations from columns containing 'XIRR(%)'
        # e.g., '3Y XIRR(%)' -> '3 YEAR'
        available_xirr_cols = [c for c in df.columns if isinstance(c, str) and "XIRR(%)" in c]
        all_durations = []
        for col in available_xirr_cols:
            prefix = col.split(' ')[0] # e.g., '3Y'
            if 'Y' in prefix:
                years = prefix.replace('Y', '')
                all_durations.append(f"{years} YEAR")
        
        # Sort durations numerically (1 YEAR, 3 YEAR, 5 YEAR, etc.)
        all_durations.sort(key=lambda x: int(x.split(' ')[0]))
        
        selected_durations = st.multiselect("Select Durations", all_durations, default=all_durations)
        
        table_rows = []
        
        # Fund-level metrics (Alpha, Beta, etc.)
        fund_metrics = [
            'Alpha', 'Beta', 'Sharpe Ratio', 'Standard Deviation', 'YTM', 
            'Average Maturity', 'Sortino Ratio', 'CY Quartile Rank', 
            'PY Quartile Rank', 'R-Squared', 'Information Ratio', 
            'Up Market Capture Ratio', 'Down Market Capture Ratio'
        ]
        available_fund_metrics = [m for m in fund_metrics if m in fund_data.index and not pd.isna(fund_data[m])]
        
        for dur in selected_durations:
            prefix = dur.split(' ')[0] + "Y"
            # Scale based on sip_amount (base data is for 10,000)
            scaling_factor = sip_amount / 10000.0
            
            invested = fund_data.get(f"{prefix} invested amount")
            current = fund_data.get(f"{prefix} Current value")
            xirr = fund_data.get(f"{prefix} XIRR(%)")
            
            if not pd.isna(invested) or not pd.isna(current) or not pd.isna(xirr):
                row = {
                    "DURATION": dur,
                    "INVESTED AMOUNT": invested * scaling_factor if not pd.isna(invested) else np.nan,
                    "CURRENT VALUE": current * scaling_factor if not pd.isna(current) else np.nan,
                    "XIRR %": xirr
                }
                # Add fund metrics to each row for easy display/copy
                for m in available_fund_metrics:
                    row[m] = fund_data[m]
                table_rows.append(row)
        
        display_df = pd.DataFrame(table_rows)
        
        # UI for display
        st.subheader(f"📈 {fund_name}")
        st.write(f"How ₹{sip_amount:,} Monthly SIP Grows Over Time")
        
        # Selectable columns for copy-paste
        all_cols = display_df.columns.tolist()
        # Default columns are the main growth ones
        default_cols = ["DURATION", "INVESTED AMOUNT", "CURRENT VALUE", "XIRR %"]
        selected_cols = st.multiselect("Select columns to display", all_cols, default=[c for c in default_cols if c in all_cols])
        
        if selected_cols:
            formatted_df = display_df[selected_cols].copy()
            
            # Apply formatting for display
            for col in formatted_df.columns:
                if col == "INVESTED AMOUNT" or col == "CURRENT VALUE":
                    formatted_df[col] = formatted_df[col].apply(format_currency)
                elif col == "XIRR %" or col in fund_metrics:
                    formatted_df[col] = formatted_df[col].apply(format_percentage)
                
            # Use our custom HTML table with copy button that won't conflict
            import hashlib
            table_hash = hashlib.md5(f"{fund_name}_{sip_amount}".encode()).hexdigest()[:8]
            display_table_with_copy(formatted_df, f"sip_single_{table_hash}")

        # Footer info like the image
        today_str = datetime.now().strftime("%d %b %Y")
        st.markdown(f"**AS ON {today_str}**")
        
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("🔴 **Subscribe**")
        with col2:
            st.markdown("🟢 **9047268800**")
            
        st.caption("Mutual Fund investments are subject to market risks, read all scheme related documents carefully. The past performance of the mutual funds is not necessarily indicative of future performance of the schemes.")

    elif view_mode == "Top Funds View":
        st.header("🏆 Top Funds Ranking")
        
        col1, col2 = st.columns(2)
        with col1:
            top_n = st.number_input("Select Top N", min_value=1, max_value=len(df), value=5, step=1)
        with col2:
            # Sort metrics: XIRR columns, Alpha, and other available metrics
            requested_metrics = [
                'Alpha', 'Beta', 'Sharpe Ratio', 'Standard Deviation', 'YTM', 
                'Average Maturity', 'Sortino Ratio', 'CY Quartile Rank', 
                'PY Quartile Rank', 'R-Squared', 'Information Ratio', 
                'Up Market Capture Ratio', 'Down Market Capture Ratio'
            ]
            
            # Combine return columns with other available metrics
            base_metrics = [c for c in df.columns if isinstance(c, str) and ("XIRR(%)" in c)]
            metric_options = base_metrics + [m for m in requested_metrics if m in df.columns]
            
            if not metric_options:
                st.error("No sorting metrics found in this file format.")
                st.stop()
            
            sort_metric = st.selectbox("Sort by Metric", metric_options, index=metric_options.index("Alpha") if "Alpha" in metric_options else 0)
            
        # Get top funds
        # Fix: Ensure sort metric column is numeric to avoid 'TypeError'
        df_sorted = df.copy()
        df_sorted[sort_metric] = pd.to_numeric(df_sorted[sort_metric], errors='coerce')
        
        # Determine sort order: Smallest to Largest for risk metrics (Beta, SD), Largest to Smallest for others
        is_ascending = sort_metric in ['Beta', 'Standard Deviation', 'Down Market Capture Ratio']
        top_funds = df_sorted.sort_values(by=sort_metric, ascending=is_ascending).head(top_n).copy()
        
        # Scale columns based on sip_amount (base data is for 10,000)
        scaling_factor = sip_amount / 10000.0
        for col in top_funds.columns:
            if isinstance(col, str) and ("invested amount" in col or "Current value" in col):
                top_funds[col] = top_funds[col] * scaling_factor
        default_display_cols = ['Category', 'Scheme Name', sort_metric]
        other_cols = [c for c in df.columns if c not in default_display_cols]
        selected_display_cols = st.multiselect("Select additional columns", other_cols, default=[])
        
        final_cols = default_display_cols + selected_display_cols
        
        # Apply formatting to top_funds
        for col in final_cols:
            if "XIRR(%)" in col or "Alpha" in col or col in requested_metrics:
                top_funds[col] = top_funds[col].apply(format_percentage)
            elif "invested amount" in col or "Current value" in col:
                top_funds[col] = top_funds[col].apply(format_currency)

        # Use our custom HTML table with copy button that won't conflict
        display_table_with_copy(top_funds[final_cols], "sip_top_funds")

else:
    st.info("Please upload an Excel file or ensure 'SIP Returns (6).xls' exists in the directory.")
