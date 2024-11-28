import dask.dataframe as dd
import streamlit as st
import plotly.express as px
import pandas as pd
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode
from pptx import Presentation
from pptx.util import Inches
from io import BytesIO
import os
import s3fs  # For accessing S3 data

# Define S3 file path
s3_path = "s3://documentsapi/industry_data/public_comp_data.parquet"

# Streamlit secrets can be accessed if credentials are provided there
try:
    storage_options = {
        'key': st.secrets["aws"]["AWS_ACCESS_KEY_ID"],
        'secret': st.secrets["aws"]["AWS_SECRET_ACCESS_KEY"],
        'client_kwargs': {'region_name': st.secrets["aws"]["AWS_DEFAULT_REGION"]}
    }
except KeyError:
    st.error("AWS credentials are not configured correctly in Streamlit secrets.")
    st.stop()

# Read Parquet file from S3 with Dask
try:
    df = dd.read_parquet(
        s3_path,
        storage_options=storage_options,
        usecols=['Name', 'Country', 'Enterprise Value (in $)', 'Revenue (in $)', 'EBITDA (in $)', 'Business Description', 'Industry']
    ).rename(columns={
        'Name': 'Company',
        'Country': 'Location',
        'Enterprise Value (in $)': 'Enterprise Value',
        'Revenue (in $)': 'Revenue',
        'EBITDA (in $)': 'EBITDA',
    })
    
    # Convert Dask DataFrame to Pandas DataFrame
    df = df.compute()

    # Convert columns to numeric
    df['Enterprise Value'] = pd.to_numeric(df['Enterprise Value'], errors='coerce')
    df['Revenue'] = pd.to_numeric(df['Revenue'], errors='coerce')
    df['EBITDA'] = pd.to_numeric(df['EBITDA'], errors='coerce')

    # Calculate EV/Revenue and EV/EBITDA
    df['EV/Revenue'] = df['Enterprise Value'] / df['Revenue']
    df['EV/EBITDA'] = df['Enterprise Value'] / df['EBITDA']

except Exception as e:
    st.error(f"Error loading data from S3: {e}")
    st.stop()
    
# Streamlit app title
st.set_page_config(page_title="Public Listed Companies Analysis", layout="wide")

# Get unique values for Industry and Location filters
industries = df['Industry'].dropna().unique()
locations = df['Location'].dropna().unique()

# Display multi-select filters at the top without default selections
col1, col2 = st.columns(2)
selected_industries = col1.multiselect("Select Industry", industries)
selected_locations = col2.multiselect("Select Location", locations)

# Filter data based on multi-selections using .isin()
if selected_industries and selected_locations:
    filtered_df = df[df['Industry'].isin(selected_industries) & df['Location'].isin(selected_locations)]
    filtered_df = filtered_df[['Company',  'EV/Revenue', 'EV/EBITDA', 'Business Description']]
    filtered_df['EV/Revenue'] = filtered_df['EV/Revenue'].round(1)
    filtered_df['EV/EBITDA'] = filtered_df['EV/EBITDA'].round(1)

    # Set up Ag-Grid for selection
    st.title("Public Listed Companies")
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
        fig1 = px.bar(avg_data, x='Company', y='EV/Revenue', title="EV/Revenue", text='EV/Revenue')
        fig1.update_traces(marker_color=color_ev_revenue, texttemplate='%{text:.1f}'+'x', textposition='auto',textfont=dict(size=10))
        fig1.update_layout(yaxis_title="EV/Revenue", xaxis_title=" ",bargap=0.4,bargroupgap=0.4,yaxis=dict(showgrid=False),xaxis=dict(tickangle=0,automargin=True))

        st.plotly_chart(fig1)

        # Create the EV/EBITDA chart with data labels
        fig2 = px.bar(avg_data, x='Company', y='EV/EBITDA', title="EV/EBITDA", text='EV/EBITDA')
        fig2.update_traces(marker_color=color_ev_ebitda,texttemplate='%{text:.1f}'+'x', textposition='auto',textfont=dict(size=10))
        fig2.update_layout(yaxis_title="EV/EBITDA", xaxis_title=" ",bargap=0.4,bargroupgap=0.4,yaxis=dict(showgrid=False),xaxis=dict(tickangle=0,automargin=True))

        st.plotly_chart(fig2)

        # Button to export charts to PowerPoint
        export_ppt = st.button("Export Charts to PowerPoint")

        if export_ppt:
            # Define the correct path to your PowerPoint template
            template_path = os.path.join(os.getcwd(), "streamlit_dashboard", "data", "main_template_pitch.pptx")
            
            # Check if the file exists before attempting to load
            if not os.path.exists(template_path):
                st.error(f"PowerPoint template not found at: {template_path}")
                st.stop()

            ppt = Presentation(template_path)
            slide1 = ppt.slides[11]  # You can change the index to 0 for the first slide, 1 for the second slide, etc.
            
            # If slide does not exist, you can choose to add a new one
            if slide1 is None:
                slide_layout = ppt.slide_layouts[5]  # If no slide exists, create a blank slide
                slide1 = ppt.slides.add_slide(slide_layout)

            # Remove title
            title1 = slide1.shapes.title
            # title1.text = ""  # Remove chart title
            
            # Save EV/Revenue chart to an image
            fig1_image = BytesIO()
            fig1.write_image(fig1_image, format="png", width=900, height=300)
            fig1_image.seek(0)
            slide1.shapes.add_picture(fig1_image, Inches(0.11), Inches(0.90), width=Inches(9), height=Inches(2.8))

            # Add EV/EBITDA chart to the same slide
            fig2_image = BytesIO()
            fig2.write_image(fig2_image, format="png", width=900, height=300)
            fig2_image.seek(0)
            slide1.shapes.add_picture(fig2_image, Inches(0.11), Inches(3.70), width=Inches(9), height=Inches(2.8))

            # Save PowerPoint to BytesIO object for download
            ppt_bytes = BytesIO()
            ppt.save(ppt_bytes)
            ppt_bytes.seek(0)

            # Provide download link for PowerPoint
            st.download_button(
                label="Download PowerPoint",
                data=ppt_bytes,
                file_name="public_comps.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )

else:
    st.write("Please select at least one Industry and Location to view data.")