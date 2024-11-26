import pandas as pd
import streamlit as st
import plotly.express as px
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode
from pptx import Presentation
from pptx.util import Inches
from io import BytesIO
from sqlalchemy import create_engine

# Streamlit app title and layout
st.set_page_config(page_title="Precedent Transactions", layout="wide")

# Database connection details from Streamlit secrets
try:
    db_user = st.secrets["mysql"]["user"]
    db_password = st.secrets["mysql"]["password"]
    db_host = st.secrets["mysql"]["host"]
    db_name = st.secrets["mysql"]["database"]
except KeyError:
    st.error("MySQL credentials are not configured correctly in Streamlit secrets.")
    st.stop()

# Connect to the MySQL database
try:
    engine = create_engine(f"mysql+pymysql://{db_user}:{db_password}@{db_host}/{db_name}")
except Exception as e:
    st.error(f"Failed to connect to the database: {e}")
    st.stop()

# Read data into a Pandas DataFrame
try:
    query = """
    SELECT id, Announced Date, Target, `EV/Revenue`, `EV/EBITDA`, `Business Description`, Industry, Location
    FROM precedent
    """
    df = pd.read_sql_query(query, con=engine)
except Exception as e:
    st.error(f"Error loading data from MySQL: {e}")
    st.stop()

df.columns = df.columns.str.strip()
df['Announced Date'] = pd.to_datetime(df['Announced Date'], errors='coerce')
df['Year'] = df['Announced Date'].dt.year


# Get unique values for Industry and Location filters
try:
    industries = df['Industry'].dropna().unique().tolist()
    locations = df['Location'].dropna().unique().tolist()
except Exception as e:
    st.error(f"Error computing unique filter values: {e}")
    st.stop()

# Display multi-select filters at the top without default selections
col1, col2 = st.columns(2)
selected_industries = col1.multiselect("Select Industry", industries)
selected_locations = col2.multiselect("Select Location", locations)

# Filter data based on multi-selections using .isin()
if selected_industries and selected_locations:
    filtered_df = df[
        (df['Industry'].isin(selected_industries)) &
        (df['Location'].isin(selected_locations))
    ][['Target', 'Year', 'EV/Revenue', 'EV/EBITDA', 'Business Description']]

    filtered_df['Year'] = filtered_df['Year'].astype(int)

    # Set up Ag-Grid for selection
    st.title("Precedent Transactions")
    gb = GridOptionsBuilder.from_dataframe(filtered_df)
    gb.configure_selection(selection_mode="multiple", use_checkbox=True)
    gb.configure_column("Target", tooltipField="Business Description", maxWidth=400)
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

        # Create the EV/Revenue chart
        fig1 = px.bar(avg_data, x='Year', y='EV/Revenue', title="EV/Revenue", text='EV/Revenue')
        fig1.update_traces(texttemplate='%{text:.1f}x', textposition='inside')
        fig1.update_layout(yaxis_title="EV/Revenue", xaxis_title=" ")

        # Display the EV/Revenue chart
        st.plotly_chart(fig1)

        # Create the EV/EBITDA chart
        fig2 = px.bar(avg_data, x='Year', y='EV/EBITDA', title="EV/EBITDA", text='EV/EBITDA')
        fig2.update_traces(texttemplate='%{text:.1f}x', textposition='inside')
        fig2.update_layout(yaxis_title="EV/EBITDA", xaxis_title=" ")

        # Display the EV/EBITDA chart
        st.plotly_chart(fig2)

        # Button to export charts to PowerPoint
        export_ppt = st.button("Export Charts to PowerPoint")

        if export_ppt:
            ppt = Presentation()

            # Add EV/Revenue chart slide
            slide_layout = ppt.slide_layouts[5]
            slide1 = ppt.slides.add_slide(slide_layout)
            title1 = slide1.shapes.title
            title1.text = "Precedent Transactions"

            fig1_image = BytesIO()
            fig1.write_image(fig1_image, format="png")
            fig1_image.seek(0)
            slide1.shapes.add_picture(fig1_image, Inches(1), Inches(1), width=Inches(8))

            # Add EV/EBITDA chart slide
            fig2_image = BytesIO()
            fig2.write_image(fig2_image, format="png")
            fig2_image.seek(0)
            slide1.shapes.add_picture(fig2_image, Inches(1), Inches(3.5), width=Inches(8))

            ppt_bytes = BytesIO()
            ppt.save(ppt_bytes)
            ppt_bytes.seek(0)

            st.download_button(
                label="Download PowerPoint",
                data=ppt_bytes,
                file_name="charts_presentation.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
else:
    st.write("Please select at least one Industry and Location to view data.")
