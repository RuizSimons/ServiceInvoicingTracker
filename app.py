import streamlit as st
import pandas as pd
import re

st.set_page_config(page_title="Surmac Service Invoicing Insight", layout="wide")

# Custom Styling
st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .stMetric { border: 1px solid #e1e4e8; padding: 15px; border-radius: 8px; background: white; }
    </style>
    """, unsafe_allow_html=True)

st.title("📊 Surmac Service Invoicing Insight")
st.markdown("Reconciling SAP WIP with Irium Labor and New Timesheet data.")

# --- Mappings provided by the user ---
STATUS_MAP = {
    "AC": "AC - QUOTE ACCEPTED",
    "AF": "AF - TO INVOICE",
    "AP": "AP - TO INVOICE PARTIALLY",
    "CP": "CP - IN ACCOUNTING",
    "DE": "DE - QUOTE PRINTED",
    "EC": "EC - IN PROGRESS",
    "ED": "ED - DEVIS EDITE",
    "FC": "FC - INVOICED",
    "RE": "RE - QUOTE REFUSED",
    "TE": "TE - DEVIS TERMINÉ",
    "TP": "TP - FINISHED PARTIALLY",
    "TR": "TR - QUOTE TRANSFERED TO ORD",
    "TT": "TT - TOTALLY FINISHED"
}

TECH_MAP = {
    "2": "DILROY SEEBALACK",
    "4": "NICOLAS BOISSEAU",
    "5": "MICHEL FLORENTINE",
    "7": "ANSON HESTON",
    "8": "DITLANE JACOBS",
    "9": "ELIZEU DA SILVA",
    "11": "NAILI SAMIR",
    "13": "JESSE DE MORAES LOBATO",
    "14": "EDDER DOS SANTOS AMARAL",
    "15": "MATTHIEU DERAIN",
    "16": "JUNO CARVAJAL",
    "17": "PAOLO RAMOS",
    "18": "IBAN OBANDO",
    "19": "HERODE ADRIEN",
    "20": "Guevara Aguilar Jesus Alfonzo",
    "21": "Jurman VAN GENDEREN",
    "22": "BYRON LOPEZ"
}

# --- Constants ---
INTERNAL_RATE = 23
EXTERNAL_RATE = 90

def convert_french_hours(time_str):
    if pd.isna(time_str) or not isinstance(time_str, str):
        return 0.0
    match = re.match(r'(\d+)h(\d+)', str(time_str))
    if match:
        hours = int(match.group(1))
        minutes = int(match.group(2))
        return float(hours + (minutes / 60))
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
        # 1. Load WIP Data
        df_wip = pd.read_excel(wip_file)
        # Force data types and clean columns
        df_wip['WO No.'] = df_wip['WO No.'].astype(str).str.strip()
        df_wip['Amount'] = pd.to_numeric(df_wip['Amount'], errors='coerce').fillna(0.0)
        df_wip['Branch'] = df_wip['Branch'].astype(str).str.strip()
        df_wip['WO type'] = df_wip['WO type'].astype(str).str.strip()
        
        # Map Status
        df_wip['Status_Desc'] = df_wip['Status'].astype(str).map(STATUS_MAP).fillna(df_wip['Status'].astype(str))

        # 2. Load Labor Register (Irium)
        df_labor = pd.read_excel(labor_file)
        df_labor['WO No.'] = df_labor['WO No.'].astype(str).str.strip()
        df_labor['Time carried out'] = pd.to_numeric(df_labor['Time carried out'], errors='coerce').fillna(0.0)
        
        # Map Technicians in Labor file
        df_labor['Technician_Name'] = df_labor['Shre Salarie'].astype(str).map(TECH_MAP).fillna(df_labor['Shre Salarie'].astype(str))
        
        # 3. Load New Timesheet
        df_ts_clean = pd.DataFrame(columns=['WO No.', 'Hours', 'Technician_From_TS'])
        if ts_file:
            df_ts = pd.read_excel(ts_file)
            ts_wo_col = "Numéro OR — Main d'œuvre (20)"
            ts_hrs_col = "Heures travaillées"
            ts_tech_col = "Technicien"
            
            if ts_wo_col in df_ts.columns:
                df_ts_clean = df_ts[[ts_wo_col, ts_hrs_col, ts_tech_col]].copy()
                df_ts_clean.columns = ['WO No.', 'Hours_Raw', 'Technician_From_TS']
                df_ts_clean['WO No.'] = df_ts_clean['WO No.'].astype(str).str.strip()
                df_ts_clean['Hours'] = df_ts_clean['Hours_Raw'].apply(convert_french_hours)
                
                # Keep tech names for reference
                df_ts_clean = df_ts_clean.groupby('WO No.').agg({
                    'Hours': 'sum',
                    'Technician_From_TS': lambda x: ", ".join(set(x.astype(str)))
                }).reset_index()

        # --- Data Integration ---
        labor_summary = df_labor.groupby('WO No.').agg({
            'Time carried out': 'sum',
            'Hourly rate': 'max',
            'Sort': 'first',
            'Technician_Name': lambda x: ", ".join(set(x.astype(str)))
        }).reset_index()

        # Merging with consistent keys
        merged = pd.merge(df_wip, labor_summary, on='WO No.', how='outer', suffixes=('_wip', '_labor'))
        merged = pd.merge(merged, df_ts_clean, on='WO No.', how='left')
        
        # Ensure all numeric columns are float
        merged['Time carried out'] = merged['Time carried out'].fillna(0.0)
        merged['Hours'] = merged['Hours'].fillna(0.0)
        merged['Total_Hours'] = merged['Time carried out'] + merged['Hours']
        merged['Amount'] = merged['Amount'].fillna(0.0)
        
        # Consolidate Technician Names
        merged['Final_Technician'] = merged['Technician_Name'].fillna('') + ", " + merged['Technician_From_TS'].fillna('')
        merged['Final_Technician'] = merged['Final_Technician'].str.strip(', ')

        # --- Reasoning & Financial Logic ---
        def analyze_job(row):
            wip_amt = float(row.get('Amount', 0.0))
            hrs = float(row.get('Total_Hours', 0.0))
            sort_val = str(row.get('Sort_wip', row.get('Sort_labor', '')))
            rate = INTERNAL_RATE if 'CES' in sort_val else EXTERNAL_RATE
            
            est_labor_val = float(hrs * rate)
            inferred_parts = float(wip_amt - est_labor_val)
            
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

        # --- Side Bar Filters ---
        st.sidebar.divider()
        st.sidebar.header("🔍 Global Filters")
        
        # Branch Filter
        branches = sorted(merged['Branch'].dropna().unique().tolist())
        sel_branch = st.sidebar.multiselect("Filter by Branch", options=branches)
        
        # WO Type Filter
        wo_types = sorted(merged['WO type'].dropna().unique().tolist())
        sel_wo_type = st.sidebar.multiselect("Filter by WO Type", options=wo_types)

        # Technician Filter
        tech_string = merged['Final_Technician'].str.cat(sep=', ')
        all_techs = sorted(list(set([t.strip() for t in tech_string.split(',') if t.strip()])))
        sel_tech = st.sidebar.multiselect("Filter by Technician", options=all_techs)
        
        # Status Filter
        merged['Status_Desc'] = merged['Status_Desc'].fillna('UNKNOWN').astype(str)
        all_statuses = sorted(list(merged['Status_Desc'].unique()))
        sel_status = st.sidebar.multiselect("Filter by WO Status", options=all_statuses)

        # --- Apply All Filters ---
        filtered_df = merged.copy()
        if sel_branch:
            filtered_df = filtered_df[filtered_df['Branch'].isin(sel_branch)]
        if sel_wo_type:
            filtered_df = filtered_df[filtered_df['WO type'].isin(sel_wo_type)]
        if sel_tech:
            filtered_df = filtered_df[filtered_df['Final_Technician'].apply(lambda x: any(t in str(x) for t in sel_tech))]
        if sel_status:
            filtered_df = filtered_df[filtered_df['Status_Desc'].isin(sel_status)]

        # --- Dashboard Display ---
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("WIP Value", f"€{filtered_df['Amount'].sum():,.2f}")
        c2.metric("Labor Hours", f"{filtered_df['Total_Hours'].sum():,.1f}h")
        c3.metric("Est. Labor Value", f"€{filtered_df['Est_Labor_Val'].sum():,.2f}")
        
        ready_count = len(filtered_df[filtered_df['Invoicing_Status'].str.contains("✅")])
        c4.metric("Ready Jobs", ready_count)

        st.divider()
        
        st.subheader("Work Order Details")
        cols_to_show = [
            'WO No.', 'Branch', 'WO type', 'Customer name', 'Status_Desc', 'Final_Technician', 
            'Amount', 'Total_Hours', 'Est_Labor_Val', 'Est_Parts_Val', 'Invoicing_Status'
        ]
        
        # Clean data for final table
        final_display = filtered_df[(filtered_df['Amount'] != 0) | (filtered_df['Total_Hours'] != 0)][cols_to_show].copy()
        final_display['Amount'] = final_display['Amount'].astype(float)
        
        st.dataframe(final_display.sort_values('Amount', ascending=False), use_container_width=True, hide_index=True)

        csv = filtered_df.to_csv(index=False).encode('utf-8')
        st.download_button("Download Filtered Report", csv, "surmac_drillthrough_report.csv", "text/csv")

    except Exception as e:
        st.error(f"Error processing files: {e}")
        st.info("Check if your Excel files have the expected 'Branch' and 'WO type' headers.")

else:
    st.info("Please upload the WIP and Labor Register files in the sidebar to start the analysis.")
