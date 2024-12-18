import dask.dataframe as dd
import streamlit as st
import pandas as pd
import numpy as np
import s3fs  # For accessing S3 data
from io import BytesIO
from pptx import Presentation
from pptx.util import Inches
import plotly.express as px

storage_options = {
        'key': st.secrets["aws"]["AWS_ACCESS_KEY_ID"],
        'secret': st.secrets["aws"]["AWS_SECRET_ACCESS_KEY"],
        'client_kwargs': {'region_name': st.secrets["aws"]["AWS_DEFAULT_REGION"]}
}

# Define S3 file paths
s3_path_rma = "s3://documentsapi/industry_data/rma_data.parquet"
s3_path_public_comp = "s3://documentsapi/industry_data/Public Listed Companies US.xlsx"

# Load the RMA data from S3 with Dask
df_rma = dd.read_parquet(s3_path_rma, storage_options=storage_options)  # Set to True if the bucket is public
df_rma = df_rma.rename(columns={
    'ReportID': 'Report_ID',      
    'Line Items': 'LineItems',    
    'Value': 'Value',             
    'Percent': 'Percent' 
})
usecols = [
    "Name", "Industry", "Revenue (in %)", "COGS (in %)", "Gross Profit (in %)", "EBITDA (in %)",
    "Operating Profit (in %)", "Other Expenses (in %)", "Operating Expenses (in %)", "Net Income (in %)",
    "Cash (in %)", "Accounts Receivables (in %)", "Inventories (in %)", "Other Current Assets (in %)",
    "Total Current Assets (in %)", "Fixed Assets (in %)", "PPE (in %)", "Total Assets (in %)",
    "Accounts Payable (in %)", "Short Term Debt (in %)", "Long Term Debt (in %)", "Other Current Liabilities (in %)",
    "Total Current Liabilities (in %)", "Other Liabilities (in %)", "Total Liabilities (in %)",
    "Net Worth (in %)", "Total Liabilities & Equity (in %)"
]
# Load the public company data
df_public_comp = pd.read_excel(s3_path_public_comp, sheet_name="FY 2023", storage_options=storage_options,usecols=usecols, engine='openpyxl')
df_public_comp = df_public_comp.rename(columns=lambda x: x.replace(" (in %)", ""))

# Filter out any missing or non-string values in the Industry column for both datasets
industries_rma = df_rma[~df_rma['Industry'].isnull() & df_rma['Industry'].map(lambda x: isinstance(x, str))]['Industry'].compute().unique()
industries_public = df_public_comp[~df_public_comp['Industry'].isnull() & df_public_comp['Industry'].map(lambda x: isinstance(x, str))]['Industry'].unique()
industries = sorted(set(industries_rma).union(set(industries_public)))

# Streamlit app title
st.title("Benchmarking")

# Single dropdown for selecting an industry
selected_industry = st.selectbox("Select Industry", industries)

# Define Income Statement and Balance Sheet LineItems
income_statement_items = ["Revenue", "COGS", "Gross Profit", "EBITDA", "Operating Profit", "Other Expenses", "Operating Expenses","Profit Before Taxes", "Net Income"]
balance_sheet_items = ["Cash", "Accounts Receivables", "Inventories", "Other Current Assets", "Total Current Assets", "Fixed Assets","Intangibles", "PPE", "Total Assets", "Accounts Payable", "Short Term Debt", "Long Term Debt", "Other Current Liabilities", "Total Current Liabilities", "Other Liabilities", "Total Liabilities", "Net Worth", "Total Liabilities & Equity"]

# Filter and prepare data only if an industry is selected
if selected_industry:

    filtered_df_rma = df_rma[df_rma['Industry'] == selected_industry].compute()

    # Map "Assets" and "Liabilities & Equity" to "Balance Sheet" if applicable
    if 'Report_ID' in filtered_df_rma.columns:
        filtered_df_rma['Report_ID'] = filtered_df_rma['Report_ID'].replace({"Assets": "Balance Sheet", "Liabilities & Equity": "Balance Sheet"})

    # RMA Percent for Income Statement and Balance Sheet
    income_statement_df_rma = filtered_df_rma[filtered_df_rma['Report_ID'] == 'Income Statement'][['LineItems', 'Percent']].rename(columns={'Percent': 'RMA Percent'})
    balance_sheet_df_rma = filtered_df_rma[filtered_df_rma['Report_ID'] == 'Balance Sheet'][['LineItems', 'Percent']].rename(columns={'Percent': 'RMA Percent'})

    filtered_df_public = df_public_comp[df_public_comp['Industry'] == selected_industry]

    df_unpivoted = pd.melt(
        filtered_df_public,
        id_vars=["Name", "Industry"],
        var_name="LineItems",
        value_name="Value"
    )
    df_unpivoted['LineItems'] = df_unpivoted['LineItems'].str.replace(" (in %)", "", regex=False)
    df_unpivoted['Value'] = pd.to_numeric(df_unpivoted['Value'].replace("-", 0), errors='coerce').fillna(0) * 100
    df_unpivoted = df_unpivoted.groupby('LineItems')['Value'].mean().reset_index()
    df_unpivoted = df_unpivoted.rename(columns={'Value': 'Public Comp Percent'})
    df_unpivoted['Public Comp Percent'] = df_unpivoted['Public Comp Percent'].round(0).astype(int).astype(str) + '%'


    # Split Public Comps data into Income Statement and Balance Sheet based on LineItems
    income_statement_df_public = df_unpivoted[df_unpivoted['LineItems'].isin(income_statement_items)]
    balance_sheet_df_public = df_unpivoted[df_unpivoted['LineItems'].isin(balance_sheet_items)]

    # Merge RMA and Public Comps for Income Statement and Balance Sheet tables
    income_statement_df = pd.merge(
        pd.DataFrame({'LineItems': income_statement_items}),
        income_statement_df_rma,
        on='LineItems',
        how='left'
    ).merge(
        income_statement_df_public,
        on='LineItems',
        how='left'
    )

    balance_sheet_df = pd.merge(
        pd.DataFrame({'LineItems': balance_sheet_items}),
        balance_sheet_df_rma,
        on='LineItems',
        how='left'
    ).merge(
        balance_sheet_df_public,
        on='LineItems',
        how='left'
    )

if selected_industry:
    # Convert percentages to numeric for plotting
    income_statement_df['RMA Percent'] = pd.to_numeric(
        income_statement_df['RMA Percent'].str.replace('%', '', regex=False), errors='coerce'
    )
    income_statement_df['Public Comp Percent'] = pd.to_numeric(
        income_statement_df['Public Comp Percent'].str.replace('%', '', regex=False), errors='coerce'
    )

    balance_sheet_df['RMA Percent'] = pd.to_numeric(
        balance_sheet_df['RMA Percent'].str.replace('%', '', regex=False), errors='coerce'
    )
    balance_sheet_df['Public Comp Percent'] = pd.to_numeric(
        balance_sheet_df['Public Comp Percent'].str.replace('%', '', regex=False), errors='coerce'
    )

    # Income Statement Bar Chart
    income_fig = px.bar(
        income_statement_df,
        x="LineItems",
        y=["RMA Percent", "Public Comp Percent"],
        # labels={"value": "Percentage (%)", "LineItems": "Items"},
        barmode="group",
        text_auto=True
    )

    income_fig.update_layout(
        xaxis_tickangle=45,
        height=400,
        margin=dict(t=50, b=50, l=50, r=50),
        showlegend=True, 
        legend_title=None,
        legend=dict(
            x=0, 
            y=1,
            traceorder='normal',
            orientation='h'
        )
    )

    # Balance Sheet Bar Chart
    balance_fig = px.bar(
        balance_sheet_df,
        x="LineItems",
        y=["RMA Percent", "Public Comp Percent"],
        barmode="group",
        text_auto=True
    )

    balance_fig.update_layout(
        xaxis_tickangle=45,
        height=400,
        margin=dict(t=50, b=50, l=50, r=50),
        showlegend=True, 
        legend_title=None,
        legend=dict(
            x=0,
            y=1,
            traceorder='normal',
            orientation='h'
        )
    )

    st.write("Income Statement")
    st.dataframe(income_statement_df.fillna(np.nan), hide_index=True, use_container_width=True)

    st.write("Income Statement Bar Chart")
    st.plotly_chart(income_fig, use_container_width=True)

    st.write("Balance Sheet")
    st.dataframe(balance_sheet_df.fillna(np.nan), hide_index=True, use_container_width=True)

    st.write("Balance Sheet Bar Chart")
    st.plotly_chart(balance_fig, use_container_width=True)

# Function to create and download PowerPoint presentation
def create_ppt(income_df, balance_df):
    prs = Presentation()
    slide_layout = prs.slide_layouts[5]  # Title and Content layout

    # Income Statement slide with table
    slide = prs.slides.add_slide(slide_layout)
    title = slide.shapes.title
    title.text = " "

    # Define table dimensions
    rows = len(income_df) + 1  # +1 for the header row
    cols = 3  # LineItems, RMA Percent, Public Comp Percent

    # Add table to slide
    table = slide.shapes.add_table(rows, cols, Inches(0.5), Inches(1), Inches(9), Inches(0.5 * rows)).table

    # Set table headers
    table.cell(0, 0).text = "Income Statement"
    table.cell(0, 1).text = "RMA"
    table.cell(0, 2).text = "Public Comps"

    # Fill table rows with data
    for i, row in income_df.iterrows():
        table.cell(i + 1, 0).text = str(row['LineItems'])
        table.cell(i + 1, 1).text = str(row['RMA Percent'] if pd.notnull(row['RMA Percent']) else 'N/A')
        table.cell(i + 1, 2).text = str(row['Public Comp Percent'] if pd.notnull(row['Public Comp Percent']) else 'N/A')

    # Balance Sheet slide with table
    slide = prs.slides.add_slide(slide_layout)
    title = slide.shapes.title
    title.text = " "

    # Define table dimensions
    rows = len(balance_df) + 1  # +1 for the header row
    cols = 3  # LineItems, RMA Percent, Public Comp Percent

    # Add table to slide
    table = slide.shapes.add_table(rows, cols, Inches(0.5), Inches(1), Inches(9), Inches(0.5 * rows)).table

    # Set table headers
    table.cell(0, 0).text = "Balance Sheet"
    table.cell(0, 1).text = "RMA"
    table.cell(0, 2).text = "Public Comps"

    # Fill table rows with data
    for i, row in balance_df.iterrows():
        table.cell(i + 1, 0).text = str(row['LineItems'])
        table.cell(i + 1, 1).text = str(row['RMA Percent'] if pd.notnull(row['RMA Percent']) else 'N/A')
        table.cell(i + 1, 2).text = str(row['Public Comp Percent'] if pd.notnull(row['Public Comp Percent']) else 'N/A')

    # Save to a BytesIO object
    ppt_bytes = BytesIO()
    prs.save(ppt_bytes)
    ppt_bytes.seek(0)
    return ppt_bytes

# Button to download the PowerPoint file
ppt_bytes = create_ppt(income_statement_df, balance_sheet_df)
st.download_button(label="Download PowerPoint", data=ppt_bytes, file_name="Benchmarking_Report.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")