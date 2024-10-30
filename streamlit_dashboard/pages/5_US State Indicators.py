import streamlit as st
import pandas as pd
import requests
import zipfile
import io
from datetime import date
from pptx import Presentation
from pptx.util import Inches
from io import BytesIO
import seaborn as sns
import matplotlib.pyplot as plt

# Initialize data
today_date = date.today().strftime("%Y-%m-%d")
state_gdp_data = None  # Initialize state GDP data

# State Data IDs
states_data_id = {
    "Alabama": {"ur_id": "ALUR", "labour_id": "LBSSA01"},
    "Alaska": {"ur_id": "AKUR", "labour_id": "LBSSA02"},
    "Arizona": {"ur_id": "AZUR", "labour_id": "LBSSA04"},
    # Add remaining states...
    "Wyoming": {"ur_id": "WYUR", "labour_id": "LBSSA56"}
}

def download_csv(state_name, data_type):
    data_ids = states_data_id.get(state_name)
    if not data_ids:
        return None

    data_id = data_ids["ur_id"] if data_type == "unemployment" else data_ids["labour_id"]
    url = f"https://fred.stlouisfed.org/graph/fredgraph.csv?id={data_id}&cosd=1976-01-01&coed={today_date}"

    response = requests.get(url)
    if response.status_code == 200:
        csv_data = pd.read_csv(io.StringIO(response.content.decode("utf-8")))
        column_name = "Unemployment" if data_type == "unemployment" else "Labour Force"
        csv_data.rename(columns={csv_data.columns[1]: column_name}, inplace=True)
        csv_data['DATE'] = pd.to_datetime(csv_data['DATE'])
        return csv_data
    else:
        st.error(f"Error downloading {data_type} data for {state_name}.")
        return None

def load_state_gdp_data():
    """Download and preprocess state-level GDP data."""
    global state_gdp_data
    url = "https://apps.bea.gov/regional/zip/SAGDP.zip"
    response = requests.get(url)

    if response.status_code == 200:
        with zipfile.ZipFile(io.BytesIO(response.content)) as z:
            csv_file_name = next((name for name in z.namelist() 
                                  if name.startswith("SAGDP1__ALL_AREAS_") and name.endswith(".csv")), None)

            if csv_file_name:
                with z.open(csv_file_name) as f:
                    df = pd.read_csv(f, usecols=lambda col: col not in [
                        "GeoFIPS", "Region", "TableName", "LineCode", 
                        "IndustryClassification", "Unit"
                    ], dtype={"Description": str})
                df = df[df["Description"] == "Current-dollar GDP (millions of current dollars) "]
                df = df.melt(id_vars=["GeoName"], var_name="Year", value_name="Value")
                df.rename(columns={"GeoName": "State"}, inplace=True)
                df = df[df["Year"].str.isdigit()]
                df["Year"] = df["Year"].astype(int)
                state_gdp_data = df[df["State"] != "United States"]
            else:
                st.error("No CSV file found with the specified prefix.")
    else:
        st.error("Failed to download GDP data.")

load_state_gdp_data()

def plot_unemployment_labour_chart(state_name):
    unemployment_data = download_csv(state_name, "unemployment")
    labour_data = download_csv(state_name, "labour")

    if unemployment_data is not None and labour_data is not None:
        unemployment_data = unemployment_data[unemployment_data['DATE'].dt.year >= 2000]
        labour_data = labour_data[labour_data['DATE'].dt.year >= 2000]

        merged_data = pd.merge(unemployment_data, labour_data, on='DATE')

        # Use Streamlit's native line chart
        st.line_chart(merged_data.set_index('DATE'), use_container_width=True)
        return merged_data
    else:
        st.warning(f"No data available for {state_name}.")
        return None

# Function to plot GDP trends using Streamlit native charts
def plot_gdp_chart(state_name):
    global state_gdp_data  # Assume `state_gdp_data` is loaded elsewhere

    if state_gdp_data is not None:
        gdp_data = state_gdp_data[state_gdp_data["State"].str.lower() == state_name.lower()]
        gdp_data = gdp_data[gdp_data["Year"] >= 2000]

        if not gdp_data.empty:
            gdp_chart_data = gdp_data.set_index('Year')['Value']
            st.line_chart(gdp_chart_data, use_container_width=True)
            return gdp_chart_data
        else:
            st.warning(f"No GDP data available for {state_name}.")
            return None
    else:
        st.warning("State GDP data not loaded.")
        return None

# Function to export charts to PowerPoint
def export_to_pptx(labour_data, gdp_data):
    prs = Presentation()
    slide_layout = prs.slide_layouts[5]

    # Slide 1: Unemployment & Labour Force
    slide1 = prs.slides.add_slide(slide_layout)
    title1 = slide1.shapes.title
    title1.text = "Unemployment & Labour Force"
    img1 = BytesIO()
    labour_data.plot().get_figure().savefig(img1, format="png")
    img1.seek(0)
    slide1.shapes.add_picture(img1, Inches(1), Inches(1), width=Inches(10))

    # Slide 2: GDP
    slide2 = prs.slides.add_slide(slide_layout)
    title2 = slide2.shapes.title
    title2.text = "GDP"
    img2 = BytesIO()
    gdp_data.plot().get_figure().savefig(img2, format="png")
    img2.seek(0)
    slide2.shapes.add_picture(img2, Inches(1), Inches(1), width=Inches(10))

    pptx_io = BytesIO()
    prs.save(pptx_io)
    pptx_io.seek(0)
    return pptx_io

# Layout for the state indicators page
def get_state_indicators_layout():
    state_name = st.selectbox("Select State", list(states_data_id.keys()), index=0)

    st.write(f"### {state_name} - Unemployment & Labour Force")
    labour_data = plot_unemployment_labour_chart(state_name)


    st.write(f"### {state_name} - GDP Over Time")
    gdp_data = plot_gdp_chart(state_name)

    # Export to PowerPoint button
    if st.button("Export Charts to PowerPoint") and labour_data is not None and gdp_data is not None:
        pptx_file = export_to_pptx(labour_data, gdp_data)
        st.download_button(
            label="Download PowerPoint",
            data=pptx_file,
            file_name="state_indicators.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
get_state_indicators_layout()