import dask.dataframe as dd
import streamlit as st
import plotly.express as px
import pandas as pd
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode
from pptx import Presentation
from pptx.util import Inches
from io import BytesIO
import s3fs  # For accessing S3 data

# Function to export charts to PowerPoint
def export_charts_to_ppt(slides_data):
    ppt = Presentation()
    slide_layout = ppt.slide_layouts[5]
    for slide_title, charts in slides_data:
        slide = ppt.slides.add_slide(slide_layout)
        slide.shapes.title.text = slide_title
        for i, chart in enumerate(charts):
            chart_image = BytesIO()
            chart.write_image(chart_image, format="png", width=800, height=300)
            chart_image.seek(0)
            slide.shapes.add_picture(chart_image, Inches(1), Inches(1 + i * 2.5), width=Inches(8))
    ppt_bytes = BytesIO()
    ppt.save(ppt_bytes)
    ppt_bytes.seek(0)
    return ppt_bytes

# Streamlit page configuration
st.set_page_config(page_title="Valuation Analysis", layout="wide")

# Define S3 file paths
precedent_path = "s3://documentsapi/industry_data/precedent.parquet"
public_comp_path = "s3://documentsapi/industry_data/public_comp_data.parquet"

# Streamlit secrets for AWS credentials
try:
    storage_options = {
        'key': st.secrets["aws"]["AWS_ACCESS_KEY_ID"],
        'secret': st.secrets["aws"]["AWS_SECRET_ACCESS_KEY"],
        'client_kwargs': {'region_name': st.secrets["aws"]["AWS_DEFAULT_REGION"]}
    }
except KeyError:
    st.error("AWS credentials are not configured correctly in Streamlit secrets.")
    st.stop()

# Load data for both Public Comps and Precedent Transactions
try:
    # Load Precedent Transactions Data
    precedent_df = dd.read_parquet(
        precedent_path,
        storage_options=storage_options,
        usecols=['Year', 'Target', 'EV/Revenue', 'EV/EBITDA', 'Business Description', 'Industry', 'Location'],
        dtype={'EV/Revenue': 'float64', 'EV/EBITDA': 'float64'}
    )

    # Load Public Comps Data
    public_comp_df = dd.read_parquet(
        public_comp_path,
        storage_options=storage_options,
        usecols=['Name', 'Country', 'Enterprise Value (in $)', 'Revenue (in $)', 'EBITDA (in $)', 'Business Description', 'Industry']
    ).rename(columns={
        'Name': 'Company',
        'Country': 'Location',
        'Enterprise Value (in $)': 'Enterprise Value',
        'Revenue (in $)': 'Revenue',
        'EBITDA (in $)': 'EBITDA'
    })

    # Ensure numeric conversion
    public_comp_df['Enterprise Value'] = dd.to_numeric(public_comp_df['Enterprise Value'], errors='coerce')
    public_comp_df['Revenue'] = dd.to_numeric(public_comp_df['Revenue'], errors='coerce')
    public_comp_df['EBITDA'] = dd.to_numeric(public_comp_df['EBITDA'], errors='coerce')

    # Drop rows with invalid data for division
    public_comp_df = public_comp_df.dropna(subset=['Enterprise Value', 'Revenue', 'EBITDA'])

    # Calculate EV/Revenue and EV/EBITDA
    public_comp_df['EV/Revenue'] = public_comp_df['Enterprise Value'] / public_comp_df['Revenue']
    public_comp_df['EV/EBITDA'] = public_comp_df['Enterprise Value'] / public_comp_df['EBITDA']

    # Get unique industries and locations from Public Comps
    public_industries = public_comp_df['Industry'].dropna().unique().compute().tolist()
    public_locations = public_comp_df['Location'].dropna().unique().compute().tolist()

    # Compute the DataFrame for use in Streamlit
    precedent_df = precedent_df.compute()
    public_comp_df = public_comp_df.compute()

except Exception as e:
    st.error(f"Error loading data from S3: {e}")
    st.stop()

# Accordion for Precedent Transactions
with st.expander("Precedent Transactions"):
    industries = precedent_df['Industry'].dropna().unique()
    locations = precedent_df['Location'].dropna().unique()
    col1, col2 = st.columns(2)
    selected_industries = col1.multiselect("Select Industry", industries, key="precedent_industries")
    selected_locations = col2.multiselect("Select Location", locations, key="precedent_locations")
    if selected_industries and selected_locations:
        filtered_precedent_df = precedent_df[
            precedent_df['Industry'].isin(selected_industries) & precedent_df['Location'].isin(selected_locations)
        ].compute()
        st.title("Precedent Transactions")
        AgGrid(filtered_precedent_df, height=400, width='100%')
        # Example chart
        fig1_precedent = px.bar(filtered_precedent_df, x="Year", y="EV/Revenue", title="EV/Revenue - Precedent Transactions")
        st.plotly_chart(fig1_precedent)

# Accordion for Public Comps
with st.expander("Public Comps"):
    industries = public_comp_df['Industry'].dropna().unique()
    locations = public_comp_df['Location'].dropna().unique()
    col1, col2 = st.columns(2)
    selected_industries = col1.multiselect("Select Industry", industries, key="public_industries")
    selected_locations = col2.multiselect("Select Location", locations, key="public_locations")
    if selected_industries and selected_locations:
        filtered_public_df = public_comp_df[
            public_comp_df['Industry'].isin(selected_industries) & public_comp_df['Location'].isin(selected_locations)
        ]
        st.title("Public Comps")
        AgGrid(filtered_public_df, height=400, width='100%')
        # Example chart
        fig1_public = px.bar(filtered_public_df, x="Company", y="EV/Revenue", title="EV/Revenue - Public Comps")
        st.plotly_chart(fig1_public)

# Button to export combined PowerPoint
if st.button("Export Combined PowerPoint"):
    slides_data = [
        ("Precedent Transactions", [fig1_precedent]),
        ("Public Comps", [fig1_public])
    ]
    ppt_bytes = export_charts_to_ppt(slides_data)
    st.download_button(
        label="Download PowerPoint",
        data=ppt_bytes,
        file_name="valuation_analysis.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )
