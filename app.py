import streamlit as st
import pandas as pd
import re

st.set_page_config(page_title="Service Invoicing Dashboard", layout="wide")

# Custom Styling
st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .stMetric { border: 1px solid #e1e4e8; padding: 15px; border-radius: 8px; background: white; }
    </style>
    """, unsafe_allow_html=True)

st.title("📊 Service Invoicing & Data Gap Analysis")
st.markdown("Reconciling SAP WIP with Irium Labor and New Timesheet data.")

# --- Constants from your problem description ---
INTERNAL_RATE = 23
EXTERNAL_RATE = 90

def convert_french_hours(time_str):
    """Converts formats like '7h15' or '0h45' to decimal (7.25, 0.75)"""
    if pd.isna(time_str) or not isinstance(time_str, str):
        return 0.0
    match = re.match(r'(\d+)h(\d+)', str(time_str))
    if match:
        hours = int(match.group(1))
        minutes = int(match.group(2))
        return hours + (minutes / 60)
    return 0.0

def load_data():
    with st.sidebar:
        st.header("📂 Upload Excel Files")
        file_wip = st.file_uploader("Upload WIP-FG.xlsx", type=['xlsx'])
        file_labor = st.file_uploader("Upload Labor_Hours_Irium.xlsx", type=['xlsx'])
        file_ts = st.file_uploader("Upload Rapport journalier des heures.xlsx", type=['xlsx'])
    
    return file_wip, file_labor, file_ts

wip_file, labor_file, ts_file = load_data()

if wip_file and labor_file:
    try:
        # 1. Load WIP Data (Checking both sheets usually present in your export)
        df_wip = pd.read_excel(wip_file) 
        # Ensure WO No is a string to prevent issues with leading zeros
        df_wip['WO No.'] = df_wip['WO No.'].astype(str)

        # 2. Load Labor Register
        df_labor = pd.read_excel(labor_file)
        df_labor['WO No.'] = df_labor['WO No.'].astype(str)
        
        # 3. Load New Timesheet if available
        df_ts_clean = pd.DataFrame(columns=['WO No.', 'Hours'])
        if ts_file:
            df_ts = pd.read_excel(ts_file)
            # Map specific column name from your file
            ts_wo_col = 'Numéro OR — Main d'œuvre (20)'
            ts_hrs_col = 'Heures travaillées'
            
            if ts_wo_col in df_ts.columns:
                df_ts_clean = df_ts[[ts_wo_col, ts_hrs_col]].copy()
                df_ts_clean.columns = ['WO No.', 'Hours_Raw']
                df_ts_clean['WO No.'] = df_ts_clean['WO No.'].astype(str)
                df_ts_clean['Hours'] = df_ts_clean['Hours_Raw'].apply(convert_french_hours)
                df_ts_clean = df_ts_clean.groupby('WO No.')['Hours'].sum().reset_index()

        # --- Data Integration (The Merging Logic) ---
        
        # Aggregate Irium Labor
        labor_summary = df_labor.groupby('WO No.').agg({
            'Time carried out': 'sum',
            'Hourly rate': 'max',
            'Sort': 'first'
        }).reset_index()

        # Merge WIP with Labor Register
        merged = pd.merge(df_wip, labor_summary, on='WO No.', how='outer', suffixes=('_wip', '_labor'))
        
        # Merge with New Timesheet
        merged = pd.merge(merged, df_ts_clean, on='WO No.', how='left')
        merged['Hours_New_TS'] = merged['Hours'].fillna(0)
        
        # Combined Total Hours
        merged['Total_Hours'] = merged['Time carried out'].fillna(0) + merged['Hours_New_TS']

        # --- Reasoning & Financial Logic ---
        def analyze_job(row):
            wip_amt = row.get('Amount', 0)
            hrs = row.get('Total_Hours', 0)
            # Determine rate based on 'Sort' column from your file (VTE vs CES)
            sort_val = str(row.get('Sort_wip', row.get('Sort_labor', '')))
            rate = INTERNAL_RATE if 'CES' in sort_val else EXTERNAL_RATE
            
            est_labor_val = hrs * rate
            inferred_parts = wip_amt - est_labor_val
            
            # Status Logic
            if wip_amt > 0 and hrs > 0:
                status = "✅ Ready (Data Matched)"
            elif wip_amt > 0 and hrs == 0:
                status = "⚠️ Missing Labor (Incomplete)"
            elif wip_amt == 0 and hrs > 0:
                status = "🚨 Admin Error (Labor but no WIP)"
            else:
                status = "🔍 Investigation Needed"
            
            return pd.Series([rate, est_labor_val, inferred_parts, status])

        merged[['Used_Rate', 'Est_Labor_Val', 'Est_Parts_Val', 'Invoicing_Status']] = merged.apply(analyze_job, axis=1)

        # --- Dashboard Display ---
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total Open WIP", f"€{merged['Amount'].sum():,.2f}")
        c2.metric("Total Labor Hours", f"{merged['Total_Hours'].sum():,.1f}h")
        c3.metric("Est. Invoicable Labor", f"€{merged['Est_Labor_Val'].sum():,.2f}")
        
        ready_count = len(merged[merged['Invoicing_Status'].str.contains("✅")])
        c4.metric("Jobs Ready to Invoice", ready_count)

        st.divider()
        
        # Table view
        st.subheader("Work Order Details & Gap Analysis")
        cols_to_show = ['WO No.', 'Customer name', 'Sort_wip', 'Amount', 'Total_Hours', 'Est_Labor_Val', 'Est_Parts_Val', 'Invoicing_Status']
        # Filter out rows with zero on both sides for cleaner view
        display_df = merged[(merged['Amount'] != 0) | (merged['Total_Hours'] != 0)][cols_to_show]
        
        st.dataframe(display_df.sort_values('Amount', ascending=False), use_container_width=True, hide_index=True)

        # Download
        csv = merged.to_csv(index=False).encode('utf-8')
        st.download_button("Download Full Analysis", csv, "service_analysis.csv", "text/csv")

    except Exception as e:
        st.error(f"Error processing files: {e}")
        st.info("Ensure files contain headers: 'WO No.', 'Amount', 'Time carried out', and 'Numéro OR — Main d'œuvre (20)'")

else:
    st.info("Please upload the WIP and Labor Register files in the sidebar to start the analysis.")
