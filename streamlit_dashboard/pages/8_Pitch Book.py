import requests
import zipfile
import io
import dask.dataframe as dd
import plotly.graph_objs as go
import streamlit as st
import plotly.express as px
import pandas as pd
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode
from pptx import Presentation
from pptx.util import Inches
from io import BytesIO
import s3fs  # For accessing S3 data
from datetime import date

today = date.today().strftime("%Y-%m-%d")
state_gdp_data = None
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
st.set_page_config(page_title="Pitch Book", layout="wide")

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
    url = f"https://fred.stlouisfed.org/graph/fredgraph.csv?id={data_id}&cosd=1976-01-01&coed={today}"

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

load_state_gdp_data()

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
            text=f"Last: {last_row['Unemployment']:.1f}"+"%", showarrow=True, arrowhead=1, ax=-40, ay=-40
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
    public_industries = public_comp_df['Industry'].dropna().compute().unique().tolist()
    public_locations = public_comp_df['Location'].dropna().compute().unique().tolist()

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
        ]
        filtered_precedent_df = filtered_precedent_df[['Target', 'Year', 'EV/Revenue', 'EV/EBITDA','Business Description']]
        filtered_precedent_df['Year'] = filtered_precedent_df['Year'].astype(int)

        st.subheader("Precedent Transactions")
        gb = GridOptionsBuilder.from_dataframe(filtered_precedent_df)
        gb.configure_selection(selection_mode="multiple", use_checkbox=True)
        gb.configure_column(
            field="Target",
            tooltipField="Business Description",
            maxWidth=400
            )
        gb.configure_columns(["Business Description"], hide=False)    
        grid_options = gb.build()

        # Display Ag-Grid table
        grid_response = AgGrid(
            filtered_precedent_df,
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
            color_ev_revenue = "#032649"  # Default Plotly blue
            color_ev_ebitda = "#032649"   # Default Plotly red

            # Create the EV/Revenue chart with data labels
            fig1_precedent = px.bar(avg_data, x='Year', y='EV/Revenue', title="EV/Revenue", text='EV/Revenue')
            fig1_precedent.update_traces(marker_color=color_ev_revenue, texttemplate='%{text:.1f}'+'x', textposition='inside')
            fig1_precedent.update_layout(yaxis_title="EV/Revenue", xaxis_title=" ")

            # Display the EV/Revenue chart
            st.plotly_chart(fig1_precedent)

            # Create the EV/EBITDA chart with data labels
            fig2_precedent = px.bar(avg_data, x='Year', y='EV/EBITDA', title="EV/EBITDA", text='EV/EBITDA')
            fig2_precedent.update_traces(marker_color=color_ev_ebitda, texttemplate='%{text:.1f}'+ 'x', textposition='inside')
            fig2_precedent.update_layout(yaxis_title="EV/EBITDA", xaxis_title=" ")

            # Display the EV/EBITDA chart
            st.plotly_chart(fig2_precedent)

# Accordion for Public Comps
with st.expander("Public Comps"):
    col1, col2 = st.columns(2)
    selected_industries = col1.multiselect("Select Industry", public_industries, key="public_industries")
    selected_locations = col2.multiselect("Select Location", public_locations, key="public_locations")
    if selected_industries and selected_locations:
        filtered_df = public_comp_df[public_comp_df['Industry'].isin(selected_industries) & public_comp_df['Location'].isin(selected_locations)]
        filtered_df = filtered_df[['Company',  'EV/Revenue', 'EV/EBITDA', 'Business Description']]
        filtered_df['EV/Revenue'] = filtered_df['EV/Revenue'].round(1)
        filtered_df['EV/EBITDA'] = filtered_df['EV/EBITDA'].round(1)

        # Set up Ag-Grid for selection
        st.subheader("Public Listed Companies")
        gb = GridOptionsBuilder.from_dataframe(filtered_df)
        gb.configure_selection(selection_mode="multiple", use_checkbox=True)
        gb.configure_column(
            field="Company",
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
            avg_data = selected_data.groupby('Company')[['EV/Revenue', 'EV/EBITDA']].mean().reset_index()

            color_ev_revenue = "#032649"  # Default Plotly blue
            color_ev_ebitda = "#032649"   # Default Plotly red

            # Create the EV/Revenue chart with data labels
            fig1_public = px.bar(avg_data, x='Company', y='EV/Revenue', title="EV/Revenue", text='EV/Revenue')
            fig1_public.update_traces(marker_color=color_ev_revenue, texttemplate='%{text:.1f}'+'x', textposition='inside')
            fig1_public.update_layout(yaxis_title="EV/Revenue", xaxis_title=" ")

            # Display the EV/Revenue chart
            st.plotly_chart(fig1_public)

            # Create the EV/EBITDA chart with data labels
            fig2_public = px.bar(avg_data, x='Company', y='EV/EBITDA', title="EV/EBITDA", text='EV/EBITDA')
            fig2_public.update_traces(marker_color=color_ev_ebitda,texttemplate='%{text:.1f}'+'x', textposition='inside')
            fig2_public.update_layout(yaxis_title="EV/EBITDA", xaxis_title=" ")

            # Display the EV/EBITDA chart
            st.plotly_chart(fig2_public)

with st.expander("State Indicators"):
    state_name = st.selectbox("Select State", list(states_data_id.keys()), index=0)
    st.subheader(f"{state_name} - Unemployment & Labour Force")
    labour_fig = plot_unemployment_labour_chart(state_name)
    st.subheader(f"{state_name} - GDP Over Time")
    gdp_fig = plot_gdp_chart(state_name)

# Button to export combined PowerPoint
if st.button("Export Pitchbook"):
    slides_data = [
        ("Precedent Transactions", [fig1_precedent],[fig2_precedent]),
        ("Public Comps", [fig1_public],[fig2_public]),
        (f"{state_name} - State Indicators", [labour_fig, gdp_fig])
    ]
    ppt_bytes = export_charts_to_ppt(slides_data)
    st.download_button(
        label="Download PowerPoint",
        data=ppt_bytes,
        file_name=f"pitch_book{today}.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )
