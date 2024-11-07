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
    gb.configure_default_column(editable=True, filter=True, sortable=True, resizable=True)
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

def display_shared_data():
    st.title("Access Shared Data")

    if "labour_fig" in st.session_state:
        st.subheader("Labour Force & Unemployment Data")
        st.plotly_chart(st.session_state["labour_fig"])

    if "external_fig" in st.session_state:
        st.subheader("External Driver Indicators")
        st.plotly_chart(st.session_state["external_fig"])

    if "gdp_fig" in st.session_state:
        st.subheader("GDP by Industry")
        st.plotly_chart(st.session_state["gdp_fig"])

    if "cpi_ppi_fig" in st.session_state:
        st.subheader("CPI and PPI Comparison")
        st.plotly_chart(st.session_state["cpi_ppi_fig"])
    return labour_fig, external_fig, gdp_fig, cpi_ppi_fig

def export_to_pptx(ev_revenue_transactions, ev_ebitda_transactions, ev_revenue_public, ev_ebitda_public, rma_is_table, rma_bs_table, pc_is_table, pc_bs_table,labour_fig,external_fig,gdp_fig,cpi_ppi_fig):
    prs = Presentation()
    slide_layout = prs.slide_layouts[5]

    if ev_revenue_transactions is not None:
        slide1 = prs.slides.add_slide(slide_layout)
        title1 = slide1.shapes.title
        title1.text = "Precedent Transactions EV/Revenue Chart"
        img1 = BytesIO()
        ev_revenue_transactions.plot(kind='bar').get_figure().savefig(img1, format='png', bbox_inches='tight')
        img1.seek(0)
        slide1.shapes.add_picture(img1, Inches(0.5), Inches(1.5), width=Inches(9), height=Inches(3))

    if ev_ebitda_transactions is not None:
        slide2 = prs.slides.add_slide(slide_layout)
        title2 = slide2.shapes.title
        title2.text = "Precedent Transactions EV/EBITDA Chart"
        img2 = BytesIO()
        ev_ebitda_transactions.plot(kind='bar').get_figure().savefig(img2, format='png', bbox_inches='tight')
        img2.seek(0)
        slide2.shapes.add_picture(img2, Inches(0.5), Inches(1.5), width=Inches(9), height=Inches(3))

    if ev_revenue_public is not None:
        slide3 = prs.slides.add_slide(slide_layout)
        title3 = slide3.shapes.title
        title3.text = "Public Companies EV/Revenue Chart"
        img3 = BytesIO()
        ev_revenue_public.plot(kind='bar').get_figure().savefig(img3, format='png', bbox_inches='tight')
        img3.seek(0)
        slide3.shapes.add_picture(img3, Inches(0.5), Inches(1.5), width=Inches(9), height=Inches(3))

    if ev_ebitda_public is not None:
        slide4 = prs.slides.add_slide(slide_layout)
        title4 = slide4.shapes.title
        title4.text = "Public Companies EV/EBITDA Chart"
        img4 = BytesIO()
        ev_ebitda_public.plot(kind='bar').get_figure().savefig(img4, format='png', bbox_inches='tight')
        img4.seek(0)
        slide4.shapes.add_picture(img4, Inches(0.5), Inches(1.5), width=Inches(9), height=Inches(3))

    if not rma_is_table.empty:
        slide = prs.slides.add_slide(slide_layout)
        title = slide.shapes.title
        title.text = "RMA Benchmarking - Income Statement"
        table_shape = slide.shapes.add_table(rma_is_table.shape[0]+1, rma_is_table.shape[1], Inches(0.5), Inches(1.5), Inches(9), Inches(3)).table
        for col_idx, col_name in enumerate(rma_is_table.columns):
            table_shape.cell(0, col_idx).text = col_name
        for row_idx, row_data in enumerate(rma_is_table.values):
            for col_idx, value in enumerate(row_data):
                table_shape.cell(row_idx+1, col_idx).text = str(value)

    if not pc_is_table.empty:
        slide6 = prs.slides.add_slide(slide_layout)
        title6 = slide6.shapes.title
        title6.text = "PC Benchmarking - Income Statement"
        table_shape = slide6.shapes.add_table(pc_is_table.shape[0]+1, pc_is_table.shape[1], Inches(0.5), Inches(1.5), Inches(9), Inches(3)).table
        for col_idx, col_name in enumerate(pc_is_table.columns):
            table_shape.cell(0, col_idx).text = col_name
        for row_idx, row_data in enumerate(pc_is_table.values):
            for col_idx, value in enumerate(row_data):
                table_shape.cell(row_idx+1, col_idx).text = str(value)

    if not rma_bs_table.empty:
        slide6 = prs.slides.add_slide(slide_layout)
        title6 = slide6.shapes.title
        title6.text = "RMA Benchmarking - Balance Sheet"
        table_shape = slide6.shapes.add_table(rma_bs_table.shape[0]+1, rma_bs_table.shape[1], Inches(0.5), Inches(1.5), Inches(9), Inches(3)).table
        for col_idx, col_name in enumerate(rma_bs_table.columns):
            table_shape.cell(0, col_idx).text = col_name
        for row_idx, row_data in enumerate(rma_bs_table.values):
            for col_idx, value in enumerate(row_data):
                table_shape.cell(row_idx+1, col_idx).text = str(value)

    if not pc_bs_table.empty:
        slide6 = prs.slides.add_slide(slide_layout)
        title6 = slide6.shapes.title
        title6.text = "PC Benchmarking - Balance Statement"
        table_shape = slide6.shapes.add_table(pc_bs_table.shape[0]+1, pc_bs_table.shape[1], Inches(0.5), Inches(1.5), Inches(9), Inches(3)).table
        for col_idx, col_name in enumerate(pc_bs_table.columns):
            table_shape.cell(0, col_idx).text = col_name
        for row_idx, row_data in enumerate(pc_bs_table.values):
            for col_idx, value in enumerate(row_data):
                table_shape.cell(row_idx+1, col_idx).text = str(value)

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

with st.expander("US Indicator", expanded=False):
    labour_fig,external_fig,gdp_fig,cpi_ppi_fig = display_shared_data()
if st.button("Export All Charts and Tables to PowerPoint"):
    # Ensure None values are handled properly when exporting
    ev_revenue_transactions = ev_revenue_transactions or pd.DataFrame()
    ev_ebitda_transactions = ev_ebitda_transactions or pd.DataFrame()
    ev_revenue_public = ev_revenue_public or pd.DataFrame()
    ev_ebitda_public = ev_ebitda_public or pd.DataFrame()
    rma_is_table = rma_is_table if rma_is_table is not None else pd.DataFrame()
    rma_bs_table = rma_bs_table if rma_bs_table is not None else pd.DataFrame()
    pc_is_table = pc_is_table if pc_is_table is not None else pd.DataFrame()
    pc_bs_table = pc_bs_table if pc_bs_table is not None else pd.DataFrame()

    pptx_file = export_to_pptx(ev_revenue_transactions, ev_ebitda_transactions, ev_revenue_public, ev_ebitda_public, rma_is_table, rma_bs_table, pc_is_table, pc_bs_table,labour_fig,external_fig,gdp_fig,cpi_ppi_fig)
    st.download_button(
        label="Download PowerPoint",
        data=pptx_file,
        file_name="all_charts_tables.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )