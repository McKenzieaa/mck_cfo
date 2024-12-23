import streamlit as st
import pandas as pd
import requests
import zipfile
import io
from datetime import date
from pptx import Presentation
from pptx.util import Inches
from io import BytesIO
import plotly.graph_objs as go
import seaborn as sns
import matplotlib.pyplot as plt
import dask.dataframe as dd

# Initialize data
today_date = date.today().strftime("%Y-%m-%d")
state_gdp_data = None 

# State Data IDs
states_data_id = {
    "Alabama": {"ur_id": "ALUR", "labour_id": "LBSSA01"},
    "Alaska": {"ur_id": "AKUR", "labour_id": "LBSSA02"},
    "Arizona": {"ur_id": "AZUR", "labour_id": "LBSSA04"},
    "Arkansas": {"ur_id": "ARUR", "labour_id": "LBSSA05"},
    "California": {"ur_id": "CAUR", "labour_id": "LBSSA06"},
    "Colorado": {"ur_id": "COUR", "labour_id": "LBSSA08"},
    "Connecticut": {"ur_id": "CTUR", "labour_id": "LBSSA09"},
    "Delaware": {"ur_id": "DEUR", "labour_id": "LBSSA10"},
    "District of Columbia": {"ur_id": "DCUR", "labour_id": "LBSSA11"},
    "Florida": {"ur_id": "FLUR", "labour_id": "LBSSA12"},
    "Georgia": {"ur_id": "GAUR", "labour_id": "LBSSA13"},
    "Hawaii": {"ur_id": "HIUR", "labour_id": "LBSSA15"},
    "Idaho": {"ur_id": "IDUR", "labour_id": "LBSSA16"},
    "Illinois": {"ur_id": "ILUR", "labour_id": "LBSSA17"},
    "Indiana": {"ur_id": "INUR", "labour_id": "LBSSA18"},
    "Iowa": {"ur_id": "IAUR", "labour_id": "LBSSA19"},
    "Kansas": {"ur_id": "KSUR", "labour_id": "LBSSA20"},
    "Kentucky": {"ur_id": "KYURN", "labour_id": "LBSSA21"},
    "Louisiana": {"ur_id": "LAUR", "labour_id": "LBSSA22"},
    "Maine": {"ur_id": "MEUR", "labour_id": "LBSSA23"},
    "Maryland": {"ur_id": "MDUR", "labour_id": "LBSSA24"},
    "Massachusetts": {"ur_id": "MAUR", "labour_id": "LBSSA25"},
    "Michigan": {"ur_id": "MIUR", "labour_id": "LBSSA26"},
    "Minnesota": {"ur_id": "MNUR", "labour_id": "LBSSA27"},
    "Mississippi": {"ur_id": "MSUR", "labour_id": "LBSSA28"},
    "Missouri": {"ur_id": "MTUR", "labour_id": "LBSSA29"},
    "Montana": {"ur_id": "MTUR", "labour_id": "LBSSA30"},
    "Nebraska": {"ur_id": "NEUR", "labour_id": "LBSSA31"},
    "Nevada": {"ur_id": "NVUR", "labour_id": "LBSSA32"},
    "New Hampshire": {"ur_id": "NHUR", "labour_id": "LBSSA33"},
    "New Jersey": {"ur_id": "NJURN", "labour_id": "LBSSA34"},
    "New Mexico": {"ur_id": "NMUR", "labour_id": "LBSSA35"},
    "New York": {"ur_id": "NYUR", "labour_id": "LBSSA36"},
    "North Carolina": {"ur_id": "NCUR", "labour_id": "LBSSA37"},
    "North Dakota": {"ur_id": "NDUR", "labour_id": "LBSSA38"},
    "Ohio": {"ur_id": "OHUR", "labour_id": "LBSSA39"},
    "Oklahoma": {"ur_id": "OKUR", "labour_id": "LBSSA40"},
    "Oregon": {"ur_id": "ORUR", "labour_id": "LBSSA41"},
    "Pennsylvania": {"ur_id": "PAUR", "labour_id": "LBSSA42"},
    "Puerto Rico": {"ur_id": "PRUR", "labour_id": "LBSSA43"},
    "Rhode Island": {"ur_id": "RIUR", "labour_id": "LBSSA44"},
    "South Carolina": {"ur_id": "SCUR", "labour_id": "LBSSA45"},
    "South Dakota": {"ur_id": "SDUR", "labour_id": "LBSSA46"},
    "Tennessee": {"ur_id": "TNUR", "labour_id": "LBSSA47"},
    "Texas": {"ur_id": "TXUR", "labour_id": "LBSSA48"},
    "Utah": {"ur_id": "UTUR", "labour_id": "LBSSA49"},
    "Vermont": {"ur_id": "VTUR", "labour_id": "LBSSA50"},
    "Washington": {"ur_id": "WAUR", "labour_id": "LBSSA54"},
    "West Virginia": {"ur_id": "WVUR", "labour_id": "LBSSA53"},
    "Wisconsin": {"ur_id": "WIUR", "labour_id": "LBSSA55"},
    "Wyoming": {"ur_id": "WYUR", "labour_id": "LBSSA56"}
}

line_colors = {
    "unemployment": "#032649",  # Dark blue
    "labour_force": "#EB8928",  # Orange
    "gdp": "#032649",  # Dark blue
}

def download_csv(state_name, data_type):
    data_ids = states_data_id.get(state_name)
    if not data_ids:
        return None

    data_id = data_ids["ur_id"] if data_type == "unemployment" else data_ids["labour_id"]
    url = f"https://fred.stlouisfed.org/graph/fredgraph.csv?id={data_id}&cosd=1976-01-01"

    response = requests.get(url)
    if response.status_code == 200:
        csv_data = pd.read_csv(io.StringIO(response.content.decode("utf-8")))
        column_name = "Unemployment" if data_type == "unemployment" else "Labour Force"
        csv_data.rename(columns={csv_data.columns[1]: column_name}, inplace=True)
        csv_data['observation_date'] = pd.to_datetime(csv_data['observation_date'])
        return csv_data
    else:
        st.error(f"Error downloading {data_type} data for {state_name}.")
        return None

def load_state_gdp_data():
    """Download and preprocess state-level GDP data."""
    global state_gdp_data
    url = "https://apps.bea.gov/regional/zip/SAGDP.zip"
    
    try:
        response = requests.get(url)
        if response.status_code == 200:
            with zipfile.ZipFile(io.BytesIO(response.content)) as z:
                csv_file_name = next(
                    (name for name in z.namelist() 
                     if name.startswith("SAGDP1__ALL_AREAS_") and name.endswith(".csv")), 
                    None
                )
                if csv_file_name:
                    with z.open(csv_file_name) as f:
                        # Load the CSV and exclude unnecessary columns
                        df = pd.read_csv(
                            f, 
                            usecols=lambda col: col not in [
                                "GeoFIPS", "Region", "TableName", "LineCode", 
                                "IndustryClassification", "Unit"
                            ],
                            dtype={"Description": str}
                        )

                    df = df[df["Description"] == "Current-dollar GDP (millions of current dollars) "]
                    df = df.melt(id_vars=["GeoName"], var_name="Year", value_name="Value")
                    df.rename(columns={"GeoName": "State"}, inplace=True)
                    df = df[df["Year"].str.isdigit()]
                    df["Year"] = df["Year"].astype(int)
                    df["Value"] = pd.to_numeric(df["Value"], errors='coerce')
                    df.dropna(subset=["Value"], inplace=True)
                    state_gdp_data = df[df["State"] != "United States"]
                else:
                    st.error("No matching CSV file found in the downloaded ZIP.")
        else:
            st.error(f"Failed to download GDP data. Status code: {response.status_code}")

    except Exception as e:
        st.error(f"An error occurred: {e}")

load_state_gdp_data()

def plot_unemployment_labour_chart(state_name):
    unemployment_data = download_csv(state_name, "unemployment")
    labour_data = download_csv(state_name, "labour")

    if unemployment_data is not None and labour_data is not None:
        unemployment_data = unemployment_data.rename(columns={'observation_date': 'DATE'})
        unemployment_data = unemployment_data[unemployment_data['DATE'].dt.year >= 2000]
        labour_data = labour_data.rename(columns={'observation_date': 'DATE'})
        labour_data = labour_data[labour_data['DATE'].dt.year >= 2000]

        merged_data = pd.merge(unemployment_data, labour_data, on='DATE')


        fig = go.Figure()
        fig.add_trace(go.Scatter(x=merged_data['DATE'], y=merged_data['Unemployment'], mode='lines',line=dict(color=line_colors["unemployment"]), name="Unemployment"))
        fig.add_trace(go.Scatter(x=merged_data['DATE'], y=merged_data['Labour Force'], mode='lines',line=dict(color=line_colors["labour_force"]), name="Labour Force"))

        last_row = merged_data.iloc[-1]
        fig.add_annotation(
            x=last_row['DATE'], y=last_row['Unemployment'],
            text=f" {last_row['Unemployment']:.1f}"+"%", showarrow=True, arrowhead=1, ax=-40, ay=-40
        )
        fig.add_annotation(
            x=last_row['DATE'], y=last_row['Labour Force'],
            text=f" {last_row['Labour Force']:.1f}"+"%", showarrow=True, arrowhead=1, ax=-40, ay=40
        )

        fig.update_layout(
            title=f"Labour Force & Unemployment Rate - {state_name}",
            xaxis_title=" ",
            yaxis_title="Rate",
            template="plotly_white",
            legend=dict(
                x=0, y=1,  # Upper left corner
                xanchor='left', yanchor='top',
                title_text=None 
            )
        )

        st.plotly_chart(fig, use_container_width=True)
        return fig
    else:
        st.warning(f"No data available for {state_name}.")
        return None

def plot_gdp_chart(state_name):
    global state_gdp_data

    if state_gdp_data is not None:
        gdp_data = state_gdp_data[state_gdp_data["State"].str.lower() == state_name.lower()]
        gdp_data = gdp_data[gdp_data["Year"] >= 2000]

        if not gdp_data.empty:
            fig = go.Figure()
            fig.add_trace(go.Scatter(x=gdp_data['Year'], y=gdp_data['Value'], mode='lines',line=dict(color=line_colors["gdp"]), name=f"{state_name} GDP"))

            last_row = gdp_data.iloc[-1]
            fig.add_annotation(
                x=last_row['Year'], y=last_row['Value'],
                text=f" {last_row['Value']:.0f}", showarrow=True, arrowhead=1, ax=-40, ay=-40
            )

            fig.update_layout(
                title=(f"GDP - {state_name}"),
                xaxis_title=" ",
                yaxis_title="GDP (Millions of Dollars)",
                template="plotly_white"
            )
            st.plotly_chart(fig, use_container_width=True)
            return fig
        else:
            st.warning(f"No GDP data available for {state_name}.")
            return None
    else:
        st.warning("State GDP data not loaded.")
        return None

def export_to_pptx(labour_fig, gdp_fig):
    prs = Presentation()
    slide_layout = prs.slide_layouts[5]

    slide1 = prs.slides.add_slide(slide_layout)
    title1 = slide1.shapes.title
    title1.text = " "
    img1 = BytesIO()
    labour_fig.write_image(img1, format="png")
    img1.seek(0)
    slide1.shapes.add_picture(img1, Inches(4.7), Inches(0.30), width=Inches(5), height=Inches(3.5))

    # slide2 = prs.slides.add_slide(slide_layout)
    # title2 = slide2.shapes.title
    # title2.text = " "
    img2 = BytesIO()
    gdp_fig.write_image(img2, format="png")
    img2.seek(0)
    slide1.shapes.add_picture(img2, Inches(4.7), Inches(3.85), width=Inches(5), height=Inches(3.5))

    pptx_io = BytesIO()
    prs.save(pptx_io)
    pptx_io.seek(0)
    return pptx_io

def get_state_indicators_layout():
    st.title("US State Indicators")
    state_name = st.selectbox("Select State", list(states_data_id.keys()), index=0)

    st.subheader(f"{state_name} - Unemployment & Labour Force")
    labour_fig = plot_unemployment_labour_chart(state_name)

    st.subheader(f"{state_name} - GDP Over Time")
    gdp_fig = plot_gdp_chart(state_name)

    if st.button("Export Charts to PowerPoint") and labour_fig is not None and gdp_fig is not None:
        pptx_file = export_to_pptx(labour_fig, gdp_fig)
        st.download_button(
            label="Download PowerPoint",
            data=pptx_file,
            file_name="state_indicators.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )

get_state_indicators_layout()