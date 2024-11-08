import os
import pandas as pd
import streamlit as st
from io import BytesIO
from pptx import Presentation
from pptx.util import Inches
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode


st.set_page_config(page_title="Transaction and Public Company Dashboard", layout="wide")

st.markdown(
    """
    <style>
    html, body, [class*="stMarkdown"] {
        font-size: 14px;
    }
    h1, h2, h3, h4, h5, h6 {
        font-size: 16px !important;
    }
    .ag-root-wrapper, .ag-theme-alpine {
        font-size: 13px !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

path_transaction = r'streamlit_dashboard/data/Updated - Precedent Transaction.xlsx'
path_public_comps = r"streamlit_dashboard/data/Public Listed Companies US.xlsx"
rma_file_path = r"streamlit_dashboard/data/RMA.xlsx"

@st.cache_data
def get_transactions_data():
    df = pd.read_excel(path_transaction, sheet_name="Final - Precedent Transactions")
    df['Announced Date'] = pd.to_datetime(df['Announced Date'], errors='coerce')
    df.dropna(subset=['Announced Date'], inplace=True)
    df['Year'] = df['Announced Date'].dt.year.astype(int)
    df['EV/Revenue'] = pd.to_numeric(df['EV/Revenue'], errors='coerce').fillna(0).round(1)
    df['EV/EBITDA'] = pd.to_numeric(df['EV/EBITDA'], errors='coerce').fillna(0).round(1)
    columns_to_display = {
        'Target': 'Company',
        'Geographic Locations': 'Location',
        'Year': 'Year',
        'Industry': 'Industry',
        'EV/Revenue': 'EV/Revenue',
        'EV/EBITDA': 'EV/EBITDA',
        'Business Description': 'Business Description'
    }
    return df[list(columns_to_display.keys())].rename(columns=columns_to_display)

@st.cache_data
def get_public_comps_data():
    df = pd.read_excel(path_public_comps, sheet_name="FY 2023")
    df['Enterprise Value (in $)'] = pd.to_numeric(df['Enterprise Value (in $)'], errors='coerce')
    df['Revenue (in $)'] = pd.to_numeric(df['Revenue (in $)'], errors='coerce').round(1)
    df['EBITDA (in $)'] = pd.to_numeric(df['EBITDA (in $)'], errors='coerce').round(1)
    df['EV/Revenue'] = df['Enterprise Value (in $)'] / df['Revenue (in $)']
    df['EV/EBITDA'] = df['Enterprise Value (in $)'] / df['EBITDA (in $)']
    df = df.dropna(subset=['Country', 'Industry', 'EV/Revenue', 'EV/EBITDA'])
    return df

def get_industries(file_path, sheet_name):
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        return df['Industry'].unique().tolist()
    except Exception as e:
        st.error(f"Error loading industries: {str(e)}")
        return []

def create_table(file_path, selected_industries):
    try:
        is_rma_df = pd.read_excel(file_path, sheet_name='IS - RMA')
        bs_rma_df = pd.read_excel(file_path, sheet_name='BS - RMA')
        is_rma_df['Industry'] = is_rma_df['Industry'].str.strip().str.lower()
        bs_rma_df['Industry'] = bs_rma_df['Industry'].str.strip().str.lower()
        selected_lower = [industry.strip().lower() for industry in selected_industries]

        is_filtered = is_rma_df[is_rma_df['Industry'].isin(selected_lower)][['MainLineItems', 'Value (in %)']]
        bs_filtered = bs_rma_df[bs_rma_df['Industry'].isin(selected_lower)][['MainLineItems', 'Value (in %)']]

        is_filtered['Value (in %)'] = (is_filtered['Value (in %)'] * 100).round(0).astype(str) + '%'
        bs_filtered['Value (in %)'] = (bs_filtered['Value (in %)'] * 100).round(0).astype(str) + '%'

        return is_filtered, bs_filtered
    except Exception as e:
        st.error(f"Error creating RMA tables: {str(e)}")
        return pd.DataFrame(), pd.DataFrame()

@st.cache_data
def load_public_comps_data():
    try:
        df = pd.read_excel(path_public_comps, sheet_name="FY 2023")
        df = df[df['Industry'].notna() & (df['Industry'] != "")]
        df_unpivoted = pd.melt(
            df,
            id_vars=["Name", "Country", "Industry", "Business Description", "SIC Code"],
            var_name="LineItems",
            value_name="Value"
        )
        df_unpivoted['LineItems'] = df_unpivoted['LineItems'].str.replace(" (in %)", "", regex=False)
        df_unpivoted['Value'] = pd.to_numeric(df_unpivoted['Value'].replace("-", 0), errors='coerce').fillna(0)
        df_unpivoted["Value"] = df_unpivoted['Value'] * 100
        industries = df_unpivoted['Industry'].unique().tolist()
        return industries, df_unpivoted
    except Exception as e:
        st.error(f"Error loading public comps data: {str(e)}")
        return [], pd.DataFrame()

def get_benchmarking_layout():
    # st.title("Benchmarking Dashboard")

    rma_is_table, rma_bs_table = pd.DataFrame(), pd.DataFrame()
    pc_is_table, pc_bs_table = pd.DataFrame(), pd.DataFrame()

    st.subheader("RMA Benchmarking")
    selected_rma_industries = st.multiselect("Select Industry (RMA)", rma_industries)
    if selected_rma_industries:
        rma_is_table, rma_bs_table = create_table(rma_file_path, selected_rma_industries)
        if not rma_is_table.empty and not rma_bs_table.empty:
            col1, col2 = st.columns(2)
            with col1:
                st.write("### Income Statement")
                st.dataframe(rma_is_table, use_container_width=True, hide_index=True)
            with col2:
                st.write("### Balance Sheet")
                st.dataframe(rma_bs_table, use_container_width=True, hide_index=True)

    st.subheader("Public Comps Benchmarking")
    selected_public_comps = st.multiselect("Select Industry (Public Comps)", public_comps_industries)
    if selected_public_comps:
        filtered_df = public_comps_data[public_comps_data['Industry'].isin(selected_public_comps)]

        income_items = ["Revenue", "COGS", "Gross Profit", "EBITDA", "Operating Profit", "Other Expenses", "Net Income"]
        balance_sheet_items = ["Cash", "Total Current Assets", "Fixed Assets", "Total Assets", "Total Liabilities", "Net Worth"]

        pc_is_table = (
            filtered_df[filtered_df['LineItems'].isin(income_items)]
            .groupby('LineItems')['Value']
            .mean()
            .reindex(income_items)
            .reset_index()
        )
        pc_bs_table = (
            filtered_df[filtered_df['LineItems'].isin(balance_sheet_items)]
            .groupby('LineItems')['Value']
            .mean()
            .reindex(balance_sheet_items)
            .reset_index()
        )

        def format_percentage(value):
            return f"{value:.1f}%" if not pd.isnull(value) else ""

        col1, col2 = st.columns(2)
        with col1:
            st.write("### Income Statement")
            st.dataframe(pc_is_table.style.format({'Value': format_percentage}), use_container_width=True, hide_index=True)
        with col2:
            st.write("### Balance Sheet")
            st.dataframe(pc_bs_table.style.format({'Value': format_percentage}), use_container_width=True, hide_index=True)

    # Return all four tables, even if they are empty
    return rma_is_table, rma_bs_table, pc_is_table, pc_bs_table

rma_industries = get_industries(rma_file_path, sheet_name="Industry Filter")
public_comps_industries, public_comps_data = load_public_comps_data()

def display_data(df, chart_func):
    gb = GridOptionsBuilder.from_dataframe(df)
    gb.configure_selection('multiple', use_checkbox=True)
    gb.configure_default_column(editable=False, filter=True, sortable=True, resizable=True)
    grid_options = gb.build()
    grid_response = AgGrid(
        df,
        gridOptions=grid_options,
        update_mode=GridUpdateMode.SELECTION_CHANGED,
        theme='alpine',
        fit_columns_on_grid_load=True,
        height=400,
        width='100%'
    )
    selected_rows = pd.DataFrame(grid_response['selected_rows'])
    if not selected_rows.empty:
        return chart_func(selected_rows)
    else:
        st.info("Select rows to visualize data.")
        # Ensure it always returns two values
        return None, None

def plot_transactions_charts(data):
    grouped_data = data.groupby('Year').agg(
        avg_ev_revenue=('EV/Revenue', 'mean'),
        avg_ev_ebitda=('EV/EBITDA', 'mean')
    ).reset_index()

    st.subheader("EV/Revenue")
    ev_revenue_chart_data = grouped_data[['Year', 'avg_ev_revenue']].set_index('Year')
    st.bar_chart(ev_revenue_chart_data)

    st.subheader("EV/EBITDA")
    ev_ebitda_chart_data = grouped_data[['Year', 'avg_ev_ebitda']].set_index('Year')
    st.bar_chart(ev_ebitda_chart_data)

    return ev_revenue_chart_data, ev_ebitda_chart_data

def plot_public_comps_charts(data):
    st.subheader("EV/Revenue Chart")
    ev_revenue_chart_data = data[['Name', 'EV/Revenue']].set_index('Name')
    st.bar_chart(ev_revenue_chart_data)

    st.subheader("EV/EBITDA Chart")
    ev_ebitda_chart_data = data[['Name', 'EV/EBITDA']].set_index('Name')
    st.bar_chart(ev_ebitda_chart_data)

    return ev_revenue_chart_data, ev_ebitda_chart_data

def export_to_pptx(ev_revenue_transactions, ev_ebitda_transactions, ev_revenue_public, ev_ebitda_public, rma_is_table, rma_bs_table, pc_is_table, pc_bs_table):
    prs = Presentation()
    slide_layout = prs.slide_layouts[5]  # Blank layout for charts/tables

    def add_plotly_chart_slide(prs, title, fig):
        """Adds a Plotly chart slide to the PowerPoint presentation."""
        slide = prs.slides.add_slide(slide_layout)
        slide.shapes.title.text = title

        # Save Plotly figure as an image
        img = BytesIO()
        fig.write_image(img, format='png', engine="kaleido", width=900, height=400)
        img.seek(0)
        slide.shapes.add_picture(img, Inches(0.5), Inches(1.5), width=Inches(9), height=Inches(4))

    def add_table_slide(prs, title, table_df):
        """Adds a table slide to the PowerPoint presentation."""
        slide = prs.slides.add_slide(slide_layout)
        slide.shapes.title.text = title
        rows, cols = table_df.shape
        table = slide.shapes.add_table(rows + 1, cols, Inches(0.5), Inches(1.5), Inches(9), Inches(3)).table

        # Add header
        for col_idx, col_name in enumerate(table_df.columns):
            table.cell(0, col_idx).text = str(col_name)

        # Add table data
        for row_idx, row_data in enumerate(table_df.values):
            for col_idx, value in enumerate(row_data):
                table.cell(row_idx + 1, col_idx).text = str(value)

    # Add each Plotly chart as a slide if it exists
    if ev_revenue_transactions is not None:
        add_plotly_chart_slide(prs, "Precedent Transactions EV/Revenue Chart", ev_revenue_transactions)

    if ev_ebitda_transactions is not None:
        add_plotly_chart_slide(prs, "Precedent Transactions EV/EBITDA Chart", ev_ebitda_transactions)

    if ev_revenue_public is not None:
        add_plotly_chart_slide(prs, "Public Companies EV/Revenue Chart", ev_revenue_public)

    if ev_ebitda_public is not None:
        add_plotly_chart_slide(prs, "Public Companies EV/EBITDA Chart", ev_ebitda_public)

    # Add tables if they are not empty
    if not rma_is_table.empty:
        add_table_slide(prs, "RMA Benchmarking - Income Statement", rma_is_table)

    if not pc_is_table.empty:
        add_table_slide(prs, "PC Benchmarking - Income Statement", pc_is_table)

    if not rma_bs_table.empty:
        add_table_slide(prs, "RMA Benchmarking - Balance Sheet", rma_bs_table)

    if not pc_bs_table.empty:
        add_table_slide(prs, "PC Benchmarking - Balance Sheet", pc_bs_table)

    # Save the PowerPoint presentation to a BytesIO buffer
    pptx_io = BytesIO()
    prs.save(pptx_io)
    pptx_io.seek(0)
    return pptx_io

with st.expander("Precedent Transactions", expanded=False):
    transactions_df = get_transactions_data()
    ev_revenue_transactions, ev_ebitda_transactions = display_data(transactions_df, plot_transactions_charts)

with st.expander("Public Companies", expanded=False):
    public_comps_df = get_public_comps_data()
    columns_to_display = ['Name', 'Country', 'Industry', 'EV/Revenue', 'EV/EBITDA', 'Business Description']
    filtered_df = public_comps_df[columns_to_display]
    ev_revenue_public, ev_ebitda_public = display_data(filtered_df, plot_public_comps_charts)

with st.expander("Benchmarking", expanded=False):
    rma_is_table, rma_bs_table, pc_is_table, pc_bs_table = get_benchmarking_layout()

if st.button("Export All Charts and Tables to PowerPoint"):
    ev_revenue_transactions = ev_revenue_transactions or pd.DataFrame()
    ev_ebitda_transactions = ev_ebitda_transactions or pd.DataFrame()
    ev_revenue_public = ev_revenue_public or pd.DataFrame()
    ev_ebitda_public = ev_ebitda_public or pd.DataFrame()
    rma_is_table = rma_is_table if rma_is_table is not None else pd.DataFrame()
    rma_bs_table = rma_bs_table if rma_bs_table is not None else pd.DataFrame()
    pc_is_table = pc_is_table if pc_is_table is not None else pd.DataFrame()
    pc_bs_table = pc_bs_table if pc_bs_table is not None else pd.DataFrame()

    pptx_file = export_to_pptx(
        ev_revenue_transactions, ev_ebitda_transactions, ev_revenue_public, ev_ebitda_public,
        rma_is_table, rma_bs_table, pc_is_table, pc_bs_table
    )
    
    st.download_button(
        label="Download PowerPoint",
        data=pptx_file,
        file_name="all_charts_tables.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )