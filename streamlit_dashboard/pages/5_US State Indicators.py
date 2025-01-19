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
        st.warning(f"No data available for {state_name}.")
        return None

    data_id = data_ids.get("ur_id") if data_type == "unemployment" else data_ids.get("labour_id")
    
    if not data_id:
        st.warning(f"No {data_type} ID available for {state_name}.")
        return None

    url = f"https://fred.stlouisfed.org/graph/fredgraph.csv?id={data_id}&cosd=1976-01-01"

    try:
        response = requests.get(url)
        response.raise_for_status()  # Raise an error for HTTP issues
        
        csv_data = pd.read_csv(io.StringIO(response.content.decode("utf-8")), on_bad_lines='skip')  # Skip bad lines

        if csv_data.empty:
            st.warning(f"No {data_type} data available for {state_name}.")
            return None

        column_name = "Unemployment" if data_type == "unemployment" else "Labour Force"
        csv_data.rename(columns={csv_data.columns[1]: column_name}, inplace=True)
        csv_data = csv_data.rename(columns={'observation_date': 'DATE'})

        # Handle date errors and missing values
        csv_data['DATE'] = pd.to_datetime(csv_data['DATE'], errors='coerce')
        csv_data[column_name] = pd.to_numeric(csv_data[column_name], errors='coerce')  # Convert to numeric safely

        # Drop rows where DATE is NaT
        csv_data.dropna(subset=['DATE'], inplace=True)
        csv_data.fillna(0, inplace=True)  # Fill remaining NaNs with 0

        return csv_data

    except requests.exceptions.RequestException as e:
        st.error(f"Error fetching data for {state_name}: {e}")
        return None
    except Exception as e:
        st.error(f"Unexpected error processing {data_type} data for {state_name}: {e}")
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
                        df = pd.read_csv( f, 
                            usecols=lambda col: col not in [ "GeoFIPS", "Region", "TableName", "LineCode", "IndustryClassification", "Unit" ],dtype={"Description": str})
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

def load_puerto_rico_gdp_data():
    """Download and preprocess Puerto Rico GDP data."""
    url = "https://fred.stlouisfed.org/graph/fredgraph.csv?bgcolor=%23e1e9f0&chart_type=line&drp=0&fo=open%20sans&graph_bgcolor=%23ffffff&height=450&mode=fred&recession_bars=off&txtcolor=%23444444&ts=12&tts=12&width=1140&nt=0&thu=0&trc=0&show_legend=yes&show_axis_titles=yes&show_tooltip=yes&id=NYGDPMKTPCDPRI&scale=left&cosd=1960-01-01&coed=2023-01-01&line_color=%234572a7&link_values=false&line_style=solid&mark_type=none&mw=3&lw=2&ost=-99999&oet=99999&mma=0&fml=a&fq=Annual&fam=avg&fgst=lin&fgsnd=2020-02-01&line_index=1&transformation=lin&vintage_date=2025-01-19&revision_date=2025-01-19&nd=1960-01-01"

    try:
        df = pd.read_csv(url)
        df.rename(columns={"observation_date": "Year", "NYGDPMKTPCDPRI": "Value"}, inplace=True)
        df["Year"] = pd.to_datetime(df["Year"]).dt.year
        df["State"] = "Puerto Rico"
        return df
    except Exception as e:
        print(f"An error occurred while loading Puerto Rico GDP data: {e}")
        return None
    
puerto_rico_data = load_puerto_rico_gdp_data()
load_state_gdp_data()
state_gdp = pd.concat([state_gdp_data, puerto_rico_data], ignore_index=True)
print(state_gdp.head())


def plot_unemployment_labour_chart(state_name):
    unemployment_data = download_csv(state_name, "unemployment")
    labour_data = download_csv(state_name, "labour")

    if unemployment_data is not None and labour_data is not None:
        unemployment_data = unemployment_data[unemployment_data['DATE'].dt.year >= 2000]
        labour_data = labour_data[labour_data['DATE'].dt.year >= 2000]

        merged_data = pd.merge(unemployment_data, labour_data, on='DATE')

        fig = go.Figure()
        fig.add_trace(go.Scatter(x=merged_data['DATE'], y=merged_data['Unemployment'], mode='lines',line=dict(color=line_colors["unemployment"]), name="Unemployment"))
        fig.add_trace(go.Scatter(x=merged_data['DATE'], y=merged_data['Labour Force'], mode='lines',line=dict(color=line_colors["labour_force"]), name="Labour Force"))

        last_row = merged_data.iloc[-1]
        fig.add_annotation(
            x=last_row['DATE'], y=last_row['Unemployment'],
            text=f"{last_row['Unemployment']:.1f}"+"%", showarrow=True, arrowhead=1, ax=-40, ay=-40
        )
        fig.add_annotation(
            x=last_row['DATE'], y=last_row['Labour Force'],
            text=f" {last_row['Labour Force']:.1f}"+"%", showarrow=True, arrowhead=1, ax=-40, ay=40
        )

        fig.update_layout(
            title="",
            xaxis_title=" ",
            yaxis_title="Rate",
            template="plotly_white",
            legend=dict(
                x=0.01,  # Center the legend horizontally
                y=-0.2,  # Move the legend below the chart
                xanchor='left',
                yanchor='top',
                title_text=None,
                orientation='h',  # Horizontal legend
                font=dict(size=10)
            ),
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            margin=dict(l=5, r=5, t=10, b=120),  # Increased bottom margin for legend space
            height=300,
            width=500,
            xaxis=dict(showgrid=False),
            yaxis=dict(showgrid=False)
        )

        st.plotly_chart(fig, use_container_width=True)
        return fig
    else:
        st.warning(f"No data available for {state_name}.")
        return None

def plot_gdp_chart(state_name):
    global state_gdp

    if state_gdp is not None:
        gdp_data = state_gdp[state_gdp["State"].str.lower() == state_name.lower()]
        gdp_data = gdp_data[gdp_data["Year"] >= 2000]

        if not gdp_data.empty:
            fig = go.Figure()
            fig.add_trace(go.Scatter(x=gdp_data['Year'], y=gdp_data['Value'], mode='lines',line=dict(color=line_colors["gdp"]), name=f"{state_name} GDP"))

            last_row = gdp_data.iloc[-1]
            value_in_millions = last_row['Value'] 
            formatted_value = f"{value_in_millions:.0f}"

            fig.add_annotation(
                x=last_row['Year'], y=last_row['Value'],
                text=f" {formatted_value}", showarrow=True, arrowhead=1, ax=-40, ay=-40
            )

            fig.update_layout(
                title=(""),
                xaxis_title=" ",
                yaxis=dict(
                    showgrid=False,
                    title='State GDP',
                    color="#595959",
                    tickfont=dict(color="#595959"),
                    side='left',
                    tickformat=',',
                ),

                template="plotly_white",
                legend=dict( x=0.01, y=0.01, xanchor='left', yanchor='bottom',title_text=None ),
                plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',margin=dict(l=2, r=2, t=30,b=50),height=300,width=500,xaxis=dict(showgrid=False, color="#595959",tickfont=dict(color="#595959")))

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