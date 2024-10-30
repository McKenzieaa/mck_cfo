import os
import pandas as pd
import streamlit as st

# File paths for data sources
# user_profile = os.path.expanduser("~")
# folder_path = os.path.join(user_profile, 'source', 'mck_setup', 'streamlit_dashboard', 'data')

rma_file_path = r"streamlit_dashboard/data/RMA.xlsx"
public_comps_file_path = r"streamlit_dashboard/data/Public Listed Companies US.xlsx"
# Helper functions to load and process data
def get_industries(file_path, sheet_name):
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        return df['Industry'].unique().tolist()
    except Exception as e:
        st.error(f"Error loading industries: {str(e)}")
        return []

def create_table(file_path, selected_industries):
    try:
        # Load sheets from the Excel file
        is_rma_df = pd.read_excel(file_path, sheet_name='IS - RMA')
        bs_rma_df = pd.read_excel(file_path, sheet_name='BS - RMA')

        # Clean and filter data
        is_rma_df['Industry'] = is_rma_df['Industry'].str.strip().str.lower()
        bs_rma_df['Industry'] = bs_rma_df['Industry'].str.strip().str.lower()
        selected_lower = [industry.strip().lower() for industry in selected_industries]

        # Filter based on selected industries
        is_filtered = is_rma_df[is_rma_df['Industry'].isin(selected_lower)][['MainLineItems', 'Value (in %)']]
        bs_filtered = bs_rma_df[bs_rma_df['Industry'].isin(selected_lower)][['MainLineItems', 'Value (in %)']]

        # Format values as percentages
        is_filtered['Value (in %)'] = (is_filtered['Value (in %)'] * 100).round(0).astype(str) + '%'
        bs_filtered['Value (in %)'] = (bs_filtered['Value (in %)'] * 100).round(0).astype(str) + '%'

        return is_filtered, bs_filtered
    except Exception as e:
        st.error(f"Error creating RMA tables: {str(e)}")
        return pd.DataFrame(), pd.DataFrame()

def load_public_comps_data():
    try:
        df = pd.read_excel(public_comps_file_path, sheet_name="FY 2023")
        df = df[df['Industry'].notna() & (df['Industry'] != "")]
        df_unpivoted = pd.melt(
            df,
            id_vars=["Name", "Country", "Industry", "Business Description", "SIC Code"],
            var_name="LineItems",
            value_name="Value"
        )
        df_unpivoted['LineItems'] = df_unpivoted['LineItems'].str.replace(" (in %)", "", regex=False)
        df_unpivoted['Value'] = pd.to_numeric(df_unpivoted['Value'].replace("-", 0), errors='coerce').fillna(0)

        industries = df_unpivoted['Industry'].unique().tolist()
        return industries, df_unpivoted
    except Exception as e:
        st.error(f"Error loading public comps data: {str(e)}")
        return [], pd.DataFrame()

# Load data for dropdowns and public comps
rma_industries = get_industries(rma_file_path, sheet_name="Industry Filter")
public_comps_industries, public_comps_data = load_public_comps_data()

# Layout for Benchmarking View
def get_benchmarking_layout():
    st.title("Benchmarking Dashboard")

    # RMA Benchmarking Tab
    st.subheader("RMA Benchmarking")
    selected_rma_industries = st.multiselect("Select Industry (RMA)", rma_industries)
    if selected_rma_industries:
        is_table, bs_table = create_table(rma_file_path, selected_rma_industries)
        if not is_table.empty and not bs_table.empty:
            col1, col2 = st.columns(2)
            with col1:
                st.write("### Income Statement")
                st.dataframe(is_table)
            with col2:
                st.write("### Balance Sheet")
                st.dataframe(bs_table)
        else:
            st.info("No matching data found for the selected industries.")

    # Public Comps Benchmarking Tab
    st.subheader("Public Comps Benchmarking")
    selected_public_comps = st.multiselect("Select Industry (Public Comps)", public_comps_industries)

    if selected_public_comps:
        filtered_df = public_comps_data[public_comps_data['Industry'].isin(selected_public_comps)]

        income_items = [
            "Revenue", "COGS", "Gross Profit", "EBITDA", "Operating Profit",
            "Other Expenses", "Operating Expenses", "Net Income"
        ]
        balance_sheet_items = [
            "Cash", "Accounts Receivables", "Inventories", "Other Current Assets",
            "Total Current Assets", "Fixed Assets", "PPE", "Total Assets",
            "Accounts Payable", "Short Term Debt", "Long Term Debt",
            "Other Current Liabilities", "Total Current Liabilities",
            "Other Liabilities", "Total Liabilities", "Net Worth",
            "Total Liabilities & Equity"
        ]

        # Income Statement DataFrame
        income_statement_df = (
            filtered_df[filtered_df['LineItems'].isin(income_items)]
            .groupby('LineItems')['Value']
            .mean()
            .reindex(income_items)  # Ensures the order matches 'income_items'
            .reset_index()
        )

        # Balance Sheet DataFrame
        balance_sheet_df = (
            filtered_df[filtered_df['LineItems'].isin(balance_sheet_items)]
            .groupby('LineItems')['Value']
            .mean()
            .reindex(balance_sheet_items)  # Ensures the order matches 'balance_sheet_items'
            .reset_index()
        )

        # Format values as percentages
        def format_percentage(value):
            return f"{value:.1f}%" if not pd.isnull(value) else ""

        col1, col2 = st.columns(2)

        with col1:
            st.write("### Income Statement")
            st.dataframe(
                income_statement_df.style.format({'Value'*100: format_percentage}),
                use_container_width=True
            )

        with col2:
            st.write("### Balance Sheet")
            st.dataframe(
                balance_sheet_df.style.format({'Value'*100: format_percentage}),
                use_container_width=True
            )
get_benchmarking_layout()