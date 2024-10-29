import os
import pandas as pd
import streamlit as st
from io import BytesIO
from pptx import Presentation
from pptx.util import Inches
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode

# Path to the Excel file
path_transaction = os.path.abspath(r'streamlit_dashboard/data/Updated - Precedent Transaction.xlsx')

@st.cache_data
def get_transactions_data():
    """Load and preprocess the precedent transactions data."""
    df = pd.read_excel(path_transaction, sheet_name="Final - Precedent Transactions")
    df['Announced Date'] = pd.to_datetime(df['Announced Date'], errors='coerce')
    df.dropna(subset=['Announced Date'], inplace=True)
    df['Year'] = df['Announced Date'].dt.year.astype(int)
    df['EV/Revenue'] = pd.to_numeric(df['EV/Revenue'], errors='coerce').fillna(0).round(1)
    df['EV/EBITDA'] = pd.to_numeric(df['EV/EBITDA'], errors='coerce').fillna(0).round(1)
    columns_to_display = {
        'Target': 'Company',
        'Geographic Locations': 'Location',
        'Year': 'Year',
        'Industry': 'Industry',
        'EV/Revenue': 'EV/Revenue',
        'EV/EBITDA': 'EV/EBITDA',
        'Business Description': 'Business Description'
    }
    return df[list(columns_to_display.keys())].rename(columns=columns_to_display)

def display_transactions():
    """Render the Transactions page layout."""
    st.subheader("Precedent Transactions")

    transactions_df = get_transactions_data()

    # Configure AgGrid
    gb = GridOptionsBuilder.from_dataframe(transactions_df)
    gb.configure_selection('multiple', use_checkbox=True)
    gb.configure_default_column(editable=True, filter=True, sortable=True, resizable=True)
    grid_options = gb.build()

    grid_response = AgGrid(
        transactions_df,
        gridOptions=grid_options,
        update_mode=GridUpdateMode.SELECTION_CHANGED,
        theme='alpine',
        fit_columns_on_grid_load=True,
        height=500,
        width='100%'
    )

    selected_rows = pd.DataFrame(grid_response['selected_rows'])
    bar_width = st.sidebar.slider("Select Bar Width", min_value=0.1, max_value=0.9, value=0.5, step=0.1)

    if not selected_rows.empty:
        ev_revenue_data, ev_ebitda_data = plot_transactions_charts(selected_rows, bar_width)
        export_chart_options(ev_revenue_data, ev_ebitda_data)
    else:
        st.info("Select transactions to visualize their data.")

def plot_transactions_charts(data, bar_width):
    """Plot EV/Revenue and EV/EBITDA charts."""
    grouped_data = data.groupby('Year').agg(
        avg_ev_revenue=('EV/Revenue', 'mean'),
        avg_ev_ebitda=('EV/EBITDA', 'mean')
    ).reset_index()

    st.subheader("EV/Revenue")
    ev_revenue_chart_data = grouped_data[['Year', 'avg_ev_revenue']].set_index('Year')
    st.bar_chart(ev_revenue_chart_data)

    st.subheader("EV/EBITDA")
    ev_ebitda_chart_data = grouped_data[['Year', 'avg_ev_ebitda']].set_index('Year')
    st.bar_chart(ev_ebitda_chart_data)

    return ev_revenue_chart_data, ev_ebitda_chart_data

def export_chart_options(ev_revenue_data, ev_ebitda_data):
    """Export charts as PowerPoint."""
    st.subheader("Export Charts")

    if st.button("Export Charts to PowerPoint"):
        pptx_file = export_to_pptx(ev_revenue_data, ev_ebitda_data)
        st.download_button(
            label="Download PowerPoint",
            data=pptx_file,
            file_name="transactions_charts.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )

def export_to_pptx(ev_revenue_data, ev_ebitda_data):
    """Export charts to a PowerPoint presentation."""
    prs = Presentation()
    slide_layout = prs.slide_layouts[5]

    # EV/Revenue Slide
    slide1 = prs.slides.add_slide(slide_layout)
    title1 = slide1.shapes.title
    title1.text = "EV/Revenue Chart"
    img1 = BytesIO()
    ev_revenue_data.plot(kind='bar').get_figure().savefig(img1, format='png', bbox_inches='tight')
    img1.seek(0)
    slide1.shapes.add_picture(img1, Inches(0.5), Inches(1.5), width=Inches(9), height=Inches(3))

    # EV/EBITDA Slide
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
