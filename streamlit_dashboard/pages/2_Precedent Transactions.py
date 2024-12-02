from pyspark.sql import SparkSession
from pyspark.sql.functions import col
import streamlit as st
import plotly.express as px
import pandas as pd
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode
from pptx import Presentation
from pptx.util import Inches
from io import BytesIO
import os

st.set_page_config(page_title="Precedent Transactions", layout="wide")

# Initialize Spark session
spark = SparkSession.builder \
    .appName("Precedent Transactions") \
    .getOrCreate()

# Define S3 file path and configure access
s3_path = "s3://documentsapi/industry_data/precedent.parquet"
try:
    storage_options = {
        'key': st.secrets["aws"]["AWS_ACCESS_KEY_ID"],
        'secret': st.secrets["aws"]["AWS_SECRET_ACCESS_KEY"],
        'client_kwargs': {'region_name': st.secrets["aws"]["AWS_DEFAULT_REGION"]}
    }
except KeyError:
    st.error("AWS credentials are not configured correctly in Streamlit secrets.")
    st.stop()

# Read Parquet file into a PySpark DataFrame
try:
    df = spark.read.parquet(
        s3_path
    ).select('Year', 'Target', 'EV/Revenue', 'EV/EBITDA', 'Business Description', 'Industry', 'Location')
except Exception as e:
    st.error(f"Error loading data from S3: {e}")
    st.stop()

# Get unique values for Industry and Location filters
industries = [row['Industry'] for row in df.select("Industry").distinct().collect()]
locations = [row['Location'] for row in df.select("Location").distinct().collect()]

# Display multi-select filters at the top without default selections
col1, col2 = st.columns(2)
selected_industries = col1.multiselect("Select Industry", industries)
selected_locations = col2.multiselect("Select Location", locations)

# Filter data based on multi-selections using PySpark filter() and .isin()
if selected_industries and selected_locations:
    filtered_df = df.filter(
        (col("Industry").isin(selected_industries)) & (col("Location").isin(selected_locations))
    ).select('Target', 'Year', 'EV/Revenue', 'EV/EBITDA', 'Business Description')

    # Convert PySpark DataFrame to Pandas DataFrame for Streamlit and AgGrid compatibility
    filtered_df = filtered_df.toPandas()
    filtered_df['Year'] = filtered_df['Year'].astype(int)

    # Set up Ag-Grid for selection
    st.title("Precedent Transactions")
    gb = GridOptionsBuilder.from_dataframe(filtered_df)
    gb.configure_grid_options(rowModelType='infinite')
    gb.configure_selection(selection_mode="multiple", use_checkbox=True)
    gb.configure_pagination(paginationAutoPageSize=False, paginationPageSize=50)
    gb.configure_column(
        field="Target",
        tooltipField="Business Description",
        maxWidth=400
    )
    gb.configure_columns(["Business Description"], hide=False)    
    grid_options = gb.build()

    # Display Ag-Grid table
    grid_response = AgGrid(
        filtered_df,
        gridOptions=grid_options,
        update_mode=GridUpdateMode.SELECTION_CHANGED,
        height=400,
        width='100%',
        theme='streamlit'
    )

    selected_data = pd.DataFrame(grid_response['selected_rows'])

    if not selected_data.empty:

        avg_data = selected_data.groupby('Year')[['EV/Revenue', 'EV/EBITDA']].mean().reset_index()
        avg_data['Year'] = avg_data['Year'].astype(int)

        # Define colors
        color_ev_revenue = "#032649"
        color_ev_ebitda = "#032649"

        median_ev_revenue = avg_data['EV/Revenue'].median()
        median_ev_ebitda = avg_data['EV/EBITDA'].median()

        # Create the EV/Revenue chart
        fig1_precedent = px.bar(avg_data, x='Year', y='EV/Revenue', title="EV/Revenue", text='EV/Revenue')
        fig1_precedent.update_traces(marker_color=color_ev_revenue, texttemplate='%{text:.1f}'+'x', textposition='auto')
        fig1_precedent.update_layout(
            yaxis_title="EV/Revenue", xaxis_title=" ",
            yaxis=dict(showgrid=False),
            xaxis=dict(tickmode='linear', dtick=1),
        )
        st.plotly_chart(fig1_precedent)

        # Create the EV/EBITDA chart
        fig2_precedent = px.bar(avg_data, x='Year', y='EV/EBITDA', title="EV/EBITDA", text='EV/EBITDA')
        fig2_precedent.update_traces(marker_color=color_ev_ebitda, texttemplate='%{text:.1f}'+'x', textposition='auto')
        fig2_precedent.update_layout(
            yaxis_title="EV/EBITDA", xaxis_title=" ",
            yaxis=dict(showgrid=False),
            xaxis=dict(tickmode='linear', dtick=1),
        )
        st.plotly_chart(fig2_precedent)

else:
    st.write("Please select at least one Industry and Location to view data.")
