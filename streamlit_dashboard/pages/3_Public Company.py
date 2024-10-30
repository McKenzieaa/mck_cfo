# pages/public_comps_view.py
import os
import pandas as pd
import streamlit as st
from io import BytesIO
from pptx import Presentation
from pptx.util import Inches
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode

path_public_comps= r'streamlit_dashboard/data/Public Listed Companies US.xlsx'
@st.cache_data
def get_public_comps_data():
    """Load and process the public companies data."""
    df = pd.read_excel(path_public_comps, sheet_name="FY 2023")
    df['Enterprise Value (in $)'] = pd.to_numeric(df['Enterprise Value (in $)'], errors='coerce')
    df['Revenue (in $)'] = pd.to_numeric(df['Revenue (in $)'], errors='coerce').round(1)
    df['EBITDA (in $)'] = pd.to_numeric(df['EBITDA (in $)'], errors='coerce').round(1)
    df['EV/Revenue'] = df['Enterprise Value (in $)'] / df['Revenue (in $)']
    df['EV/EBITDA'] = df['Enterprise Value (in $)'] / df['EBITDA (in $)']
    df = df.dropna(subset=['Country', 'Industry', 'EV/Revenue', 'EV/EBITDA'])
    return df

def display_public_comps():
    """Render the Public Companies page layout."""
    st.subheader("Public Companies")

    public_comps_df = get_public_comps_data()
    columns_to_display = ['Name', 'Country', 'Industry', 'EV/Revenue', 'EV/EBITDA', 'Business Description']
    filtered_df = public_comps_df[columns_to_display]

    gb = GridOptionsBuilder.from_dataframe(filtered_df)
    gb.configure_selection('multiple', use_checkbox=True)
    gb.configure_default_column(editable=True, filter=True, sortable=True, resizable=True)
    grid_options = gb.build()
    grid_response = AgGrid(
        filtered_df,
        gridOptions=grid_options,
        update_mode=GridUpdateMode.SELECTION_CHANGED,
        theme='alpine',
        fit_columns_on_grid_load=True,
        height=500,
        width='100%'
    )
    selected_rows = pd.DataFrame(grid_response['selected_rows'])
    if not selected_rows.empty:
        ev_revenue_fig, ev_ebitda_fig = plot_public_comps_charts(selected_rows)
        export_chart_options(ev_revenue_fig, ev_ebitda_fig)
    else:
        st.info("Select companies to visualize their data.")   

def plot_public_comps_charts(data):
    """Plot EV/Revenue and EV/EBITDA charts using Streamlit native charts."""
    st.subheader("EV/Revenue Chart")
    ev_revenue_chart_data = data[['Name', 'EV/Revenue']].set_index('Name')
    st.bar_chart(ev_revenue_chart_data)

    st.subheader("EV/EBITDA Chart")
    ev_ebitda_chart_data = data[['Name', 'EV/EBITDA']].set_index('Name')
    st.bar_chart(ev_ebitda_chart_data)

    return ev_revenue_chart_data, ev_ebitda_chart_data

def export_chart_options(ev_revenue_data, ev_ebitda_data):
    """Provide options to export charts as PowerPoint."""
    st.subheader("Export Charts")

    if st.button("Export Charts to PowerPoint"):
        pptx_file = export_to_pptx(ev_revenue_data, ev_ebitda_data)
        st.download_button(
            label="Download PowerPoint",
            data=pptx_file,
            file_name="public_companies_charts.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )

def export_to_pptx(ev_revenue_data, ev_ebitda_data):
    """Export charts to a PowerPoint presentation."""
    prs = Presentation()
    slide_layout = prs.slide_layouts[5]

    # Slide for EV/Revenue Chart
    slide1 = prs.slides.add_slide(slide_layout)
    title1 = slide1.shapes.title
    title1.text = "EV/Revenue Chart"
    img1 = BytesIO()
    ev_revenue_data.plot(kind='bar').get_figure().savefig(img1, format='png', bbox_inches='tight')
    img1.seek(0)
    slide1.shapes.add_picture(img1, Inches(0.5), Inches(1.5), width=Inches(9), height=Inches(3))

    # Slide for EV/EBITDA Chart
    slide2 = prs.slides.add_slide(slide_layout)
    title2 = slide2.shapes.title
    title2.text = "EV/EBITDA Chart"
    img2 = BytesIO()
    ev_ebitda_data.plot(kind='bar').get_figure().savefig(img2, format='png', bbox_inches='tight')
    img2.seek(0)
    slide2.shapes.add_picture(img2, Inches(0.5), Inches(1.5), width=Inches(9), height=Inches(3))

    pptx_io = BytesIO()
    prs.save(pptx_io)
    pptx_io.seek(0)
    return pptx_io
display_public_comps()