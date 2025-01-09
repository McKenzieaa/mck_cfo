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

s3_path = "s3://documentsapi/industry_data/benchmarking.csv"

df = pd.read_csv(s3_path, storage_options=storage_options)
df["Public Comps"] = df["File"] == "Public comps"
df["RMA"] = df["File"] == "RMA"

df = df[["Industry", "LineItems", "ReportID", "Public Comps", "RMA"]]

income_statement_df = df[df["ReportID"] == "Income Statement"]

st.header("Income Statement")
st.dataframe(income_statement_df)

fig_income = px.bar(
    income_statement_df, x="LineItems", y=["Public Comps", "RMA"],
    title="Income Statement: Public Comps vs RMA",
    barmode="group"
)
st.plotly_chart(fig_income)

assets_liabilities_df = df[df["ReportID"].isin(["Assets", "Liabilities & Equity"])]

st.header("Assets & Liabilities")
st.dataframe(assets_liabilities_df)

fig_assets_liabilities = px.bar(
    assets_liabilities_df, x="LineItems", y=["Public Comps", "RMA"],
    color="ReportID",
    title="Assets & Liabilities: Public Comps vs RMA",
    barmode="group"
)
st.plotly_chart(fig_assets_liabilities)
