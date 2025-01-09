import streamlit as st
import pandas as pd
import numpy as np
import s3fs 
from io import BytesIO
from pptx import Presentation
from pptx.util import Inches
import plotly.express as px

storage_options = {
        'key': st.secrets["aws"]["AWS_ACCESS_KEY_ID"],
        'secret': st.secrets["aws"]["AWS_SECRET_ACCESS_KEY"],
        'client_kwargs': {'region_name': st.secrets["aws"]["AWS_DEFAULT_REGION"]}
}

# Load the data from S3
s3_path = "s3://documentsapi/industry_data/benchmarking.csv"
# st.write("Loading data from:", s3_path)

df = pd.read_csv(s3_path, storage_options=storage_options)

# Filter columns
df = df[["Industry", "LineItems", "File", "ReportID"]]

# # Convert to Pandas DataFrame for Streamlit visualization
# pandas_df = df.compute()

# Filter data for "Income Statement" and display table with bar chart
income_statement_df = df[df["ReportID"] == "Income Statement"]

st.header("Income Statement")
st.write("Table:")
st.dataframe(income_statement_df)

st.write("Bar Chart:")
fig_income = px.bar(
    income_statement_df, x="LineItems", y=["public_comps", "rma"],
    title="Income Statement: Public Comps vs RMA",
    barmode="group"
)
st.plotly_chart(fig_income)

# Filter data for "Assets", "Liabilities & Equity" and display table with bar chart
assets_liabilities_df = df[df["ReportID"].isin(["Assets", "Liabilities & Equity"])]

st.header("Assets & Liabilities")
st.write("Table:")
st.dataframe(assets_liabilities_df)

st.write("Bar Chart:")
fig_assets_liabilities = px.bar(
    assets_liabilities_df, x="LineItems", y=["public_comps", "rma"],
    color="ReportID",
    title="Assets & Liabilities: Public Comps vs RMA",
    barmode="group"
)
st.plotly_chart(fig_assets_liabilities)
