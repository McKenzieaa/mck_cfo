import dask.dataframe as dd
import streamlit as st
import plotly.express as px
import pandas as pd
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode
from pptx import Presentation
from pptx.util import Inches
from io import BytesIO
import s3fs  # For accessing S3 data

# Streamlit app title and layout
st.set_page_config(page_title="Financial Analysis", layout="wide")

# Accordion for Public Comps and Precedent Transactions
with st.expander("Public Comps"):
    # Define S3 file path for Public Comps
    s3_path_public = "s3://documentsapi/industry_data/public_comp_data.parquet"

    try:
        storage_options = {
            'key': st.secrets["aws"]["AWS_ACCESS_KEY_ID"],
            'secret': st.secrets["aws"]["AWS_SECRET_ACCESS_KEY"],
            'client_kwargs': {'region_name': st.secrets["aws"]["AWS_DEFAULT_REGION"]}
        }
        df_public = dd.read_parquet(
            s3_path_public,
            storage_options=storage_options,
            usecols=['Name', 'Country', 'Enterprise Value (in $)', 'Revenue (in $)', 'EBITDA (in $)', 'Business Description', 'Industry']
        ).rename(columns={
            'Name': 'Company',
            'Country': 'Location',
            'Enterprise Value (in $)': 'Enterprise Value',
            'Revenue (in $)': 'Revenue',
            'EBITDA (in $)': 'EBITDA',
        }).compute()
        df_public['EV/Revenue'] = df_public['Enterprise Value'] / df_public['Revenue']
        df_public['EV/EBITDA'] = df_public['Enterprise Value'] / df_public['EBITDA']
    except Exception as e:
        st.error(f"Error loading Public Comps data from S3: {e}")
        st.stop()

    industries_public = df_public['Industry'].dropna().unique()
    locations_public = df_public['Location'].dropna().unique()

    col1, col2 = st.columns(2)
    selected_industries_public = col1.multiselect("Select Industry (Public Comps)", industries_public)
    selected_locations_public = col2.multiselect("Select Location (Public Comps)", locations_public)

    if selected_industries_public and selected_locations_public:
        filtered_df_public = df_public[
            (df_public['Industry'].isin(selected_industries_public)) &
            (df_public['Location'].isin(selected_locations_public))
        ]
        filtered_df_public = filtered_df_public[['Company', 'EV/Revenue', 'EV/EBITDA', 'Business Description']]
        filtered_df_public['EV/Revenue'] = filtered_df_public['EV/Revenue'].round(1)
        filtered_df_public['EV/EBITDA'] = filtered_df_public['EV/EBITDA'].round(1)

        gb = GridOptionsBuilder.from_dataframe(filtered_df_public)
        gb.configure_selection(selection_mode="multiple", use_checkbox=True)
        grid_options = gb.build()
        grid_response_public = AgGrid(filtered_df_public, gridOptions=grid_options, height=400, theme='streamlit')
        selected_data_public = pd.DataFrame(grid_response_public['selected_rows'])

        if not selected_data_public.empty:
            avg_data_public = selected_data_public.groupby('Company')[['EV/Revenue', 'EV/EBITDA']].mean().reset_index()
            fig1_public = px.bar(avg_data_public, x='Company', y='EV/Revenue', title="Public Comps - EV/Revenue")
            fig2_public = px.bar(avg_data_public, x='Company', y='EV/EBITDA', title="Public Comps - EV/EBITDA")

            st.plotly_chart(fig1_public)
            st.plotly_chart(fig2_public)

with st.expander("Precedent Transactions"):
    # Define S3 file path for Precedent Transactions
    s3_path_precedent = "s3://documentsapi/industry_data/precedent.parquet"

    try:
        df_precedent = dd.read_parquet(
            s3_path_precedent,
            storage_options=storage_options,
            usecols=['Year', 'Target', 'EV/Revenue', 'EV/EBITDA', 'Business Description', 'Industry', 'Location']
        ).compute()
    except Exception as e:
        st.error(f"Error loading Precedent Transactions data from S3: {e}")
        st.stop()

    industries_precedent = df_precedent['Industry'].unique()
    locations_precedent = df_precedent['Location'].unique()

    col3, col4 = st.columns(2)
    selected_industries_precedent = col3.multiselect("Select Industry (Precedent Transactions)", industries_precedent)
    selected_locations_precedent = col4.multiselect("Select Location (Precedent Transactions)", locations_precedent)

    if selected_industries_precedent and selected_locations_precedent:
        filtered_df_precedent = df_precedent[
            (df_precedent['Industry'].isin(selected_industries_precedent)) &
            (df_precedent['Location'].isin(selected_locations_precedent))
        ]
        filtered_df_precedent = filtered_df_precedent[['Target', 'Year', 'EV/Revenue', 'EV/EBITDA', 'Business Description']]

        gb = GridOptionsBuilder.from_dataframe(filtered_df_precedent)
        gb.configure_selection(selection_mode="multiple", use_checkbox=True)
        grid_options = gb.build()
        grid_response_precedent = AgGrid(filtered_df_precedent, gridOptions=grid_options, height=400, theme='streamlit')
        selected_data_precedent = pd.DataFrame(grid_response_precedent['selected_rows'])

        if not selected_data_precedent.empty:
            avg_data_precedent = selected_data_precedent.groupby('Year')[['EV/Revenue', 'EV/EBITDA']].mean().reset_index()
            fig1_precedent = px.bar(avg_data_precedent, x='Year', y='EV/Revenue', title="Precedent Transactions - EV/Revenue")
            fig2_precedent = px.bar(avg_data_precedent, x='Year', y='EV/EBITDA', title="Precedent Transactions - EV/EBITDA")

            st.plotly_chart(fig1_precedent)
            st.plotly_chart(fig2_precedent)

if st.button("Export All Charts to PowerPoint"):
    ppt = Presentation()
    slide_layout = ppt.slide_layouts[5]

    # Public Comps Slides
    if 'fig1_public' in locals() and 'fig2_public' in locals():
        slide_public = ppt.slides.add_slide(slide_layout)
        slide_public.shapes.title.text = "Public Comps"
        fig1_image_public = BytesIO()
        fig1_public.write_image(fig1_image_public, format="png")
        slide_public.shapes.add_picture(fig1_image_public, Inches(1), Inches(1), width=Inches(8))
        fig2_image_public = BytesIO()
        fig2_public.write_image(fig2_image_public, format="png")
        slide_public.shapes.add_picture(fig2_image_public, Inches(1), Inches(3.5), width=Inches(8))

    # Precedent Transactions Slides
    if 'fig1_precedent' in locals() and 'fig2_precedent' in locals():
        slide_precedent = ppt.slides.add_slide(slide_layout)
        slide_precedent.shapes.title.text = "Precedent Transactions"
        fig1_image_precedent = BytesIO()
        fig1_precedent.write_image(fig1_image_precedent, format="png")
        slide_precedent.shapes.add_picture(fig1_image_precedent, Inches(1), Inches(1), width=Inches(8))
        fig2_image_precedent = BytesIO()
        fig2_precedent.write_image(fig2_image_precedent, format="png")
        slide_precedent.shapes.add_picture(fig2_image_precedent, Inches(1), Inches(3.5), width=Inches(8))

    ppt_bytes = BytesIO()
    ppt.save(ppt_bytes)
    ppt_bytes.seek(0)

    st.download_button("Download PowerPoint", data=ppt_bytes, file_name="Financial_Analysis.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
