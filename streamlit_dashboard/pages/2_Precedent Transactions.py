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

st.set_page_config(page_title="Precedent Transactions", layout="wide")

# Define S3 file path
s3_path = "s3://documentsapi/industry_data/precedent.parquet"
try:
    storage_options = {
        'key': st.secrets["aws"]["AWS_ACCESS_KEY_ID"],
        'secret': st.secrets["aws"]["AWS_SECRET_ACCESS_KEY"],
        'client_kwargs': {'region_name': st.secrets["aws"]["AWS_DEFAULT_REGION"]}
    }
except KeyError:
    st.error("AWS credentials are not configured correctly in Streamlit secrets.")
    st.stop()

try:
    df = dd.read_parquet(
        s3_path,
        storage_options=storage_options,
        usecols=['Year', 'Target', 'EV/Revenue', 'EV/EBITDA', 'Business Description', 'Industry', 'Location'],
        dtype={'EV/Revenue': 'float64', 'EV/EBITDA': 'float64'}
    )
except Exception as e:
    st.error(f"Error loading data from S3: {e}")
    st.stop()
    

# Get unique values for Industry and Location filters
industries = df['Industry'].unique().compute()
locations = df['Location'].unique().compute()

# Display multi-select filters at the top without default selections
col1, col2 = st.columns(2)
selected_industries = col1.multiselect("Select Industry", industries)
selected_locations = col2.multiselect("Select Location", locations)

# Filter data based on multi-selections using .isin()
if selected_industries and selected_locations:
    filtered_df = df[df['Industry'].isin(selected_industries) & df['Location'].isin(selected_locations)]
    filtered_df = filtered_df[['Target', 'Year', 'EV/Revenue', 'EV/EBITDA','Business Description']]
    filtered_df = filtered_df.compute()  # Convert to Pandas for easier manipulation in Streamlit
    filtered_df['Year'] = filtered_df['Year'].astype(int)

    # Set up Ag-Grid for selection
    st.title("Precedent Transactions")
    gb = GridOptionsBuilder.from_dataframe(filtered_df)
    gb.configure_grid_options(rowModelType='infinite')
    gb.configure_selection(selection_mode="multiple", use_checkbox=True)
    gb.configure_pagination(paginationAutoPageSize=False, paginationPageSize=50)
    gb.configure_column(
        field="Target",
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

        avg_data = selected_data.groupby('Year')[['EV/Revenue', 'EV/EBITDA']].mean().reset_index()
        avg_data['Year'] = avg_data['Year'].astype(int)

            # Define colors
        color_ev_revenue = "#032649"  # Default Plotly blue
        color_ev_ebitda = "#032649"   # Default Plotly red

        median_ev_revenue = avg_data['EV/Revenue'].median()
        median_ev_ebitda = avg_data['EV/EBITDA'].median()

            # Create the EV/Revenue chart with data labels
        fig1_precedent = px.bar(avg_data, x='Year', y='EV/Revenue', title="EV/Revenue", text='EV/Revenue')  # No title
        fig1_precedent.update_traces(marker_color=color_ev_revenue, texttemplate='%{text:.1f}'+'x', textposition='auto',textfont=dict(size=12))
        fig1_precedent.update_layout(yaxis_title="EV/Revenue", xaxis_title=" ", bargap=0.4, bargroupgap=0.4, yaxis=dict(showgrid=False),xaxis=dict(tickmode='linear', tick0=avg_data['Year'].min(), dtick=1), shapes=[dict(type='line', x0=avg_data['Year'].min(), x1=avg_data['Year'].max(), y0=median_ev_revenue, y1=median_ev_revenue, line=dict(color='#EB8928', dash='dot', width=2))], annotations=[dict(x=avg_data['Year'].max(), y=median_ev_revenue, xanchor='left', yanchor='bottom', text=f'Median: {median_ev_revenue:.1f}'+'x', showarrow=False, font=dict(size=12, color='gray'), bgcolor='white')],plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',margin=dict(l=0, r=0, t=0),width=900,height=300)

        st.plotly_chart(fig1_precedent)

            # Create the EV/EBITDA chart with data labels
        fig2_precedent = px.bar(avg_data, x='Year', y='EV/EBITDA', title="EV/EBITDA", text='EV/EBITDA')
        fig2_precedent.update_traces(marker_color=color_ev_ebitda, texttemplate='%{text:.1f}'+ 'x', textposition='auto',textfont=dict(size=12))
        fig2_precedent.update_layout(yaxis_title="EV/EBITDA", xaxis_title=" ", bargap=0.4, bargroupgap=0.4, yaxis=dict(showgrid=False),xaxis=dict(tickmode='linear', tick0=avg_data['Year'].min(), dtick=1), shapes=[dict(type='line', x0=avg_data['Year'].min(), x1=avg_data['Year'].max(), y0=median_ev_ebitda, y1=median_ev_ebitda, line=dict(color='#EB8928', dash='dot', width=2))], annotations=[dict(x=avg_data['Year'].max(), y=median_ev_ebitda, xanchor='left', yanchor='bottom', text=f'Median: {median_ev_ebitda:.1f}'+'x', showarrow=False, font=dict(size=12, color='gray'), bgcolor='white')],plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',margin=dict(l=0, r=0, t=0),width=900,height=300)
            
        st.plotly_chart(fig2_precedent)
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
            slide1 = ppt.slides[10]  # You can change the index to 0 for the first slide, 1 for the second slide, etc.
            
            # If slide does not exist, you can choose to add a new one
            if slide1 is None:
                slide_layout = ppt.slide_layouts[5]  # If no slide exists, create a blank slide
                slide1 = ppt.slides.add_slide(slide_layout)

            # Remove title
            title1 = slide1.shapes.title
            # title1.text = ""  # Remove chart title
            
            # Save EV/Revenue chart to an image
            fig1_image = BytesIO()
            fig1_precedent.write_image(fig1_image, format="png", width=900, height=300)
            fig1_image.seek(0)
            slide1.shapes.add_picture(fig1_image, Inches(0.11), Inches(0.90), width=Inches(9), height=Inches(2.8))

            # Add EV/EBITDA chart to the same slide
            fig2_image = BytesIO()
            fig2_precedent.write_image(fig2_image, format="png", width=900, height=300)
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
                file_name="precedent_transaction.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )

else:
    st.write("Please select at least one Industry and Location to view data.")