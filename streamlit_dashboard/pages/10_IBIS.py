import streamlit as st
import pandas as pd
import mysql.connector
from plotly.subplots import make_subplots
import plotly.graph_objects as go
from pptx import Presentation
from pptx.util import Inches
import io

st.set_page_config(page_title="IBIS-Industry Analysis", layout="wide")

# Function to get distinct industries
def get_industries():
    host = st.secrets["mysql"]["host"]
    user = st.secrets["mysql"]["user"]
    password = st.secrets["mysql"]["password"]
    database = st.secrets["mysql"]["database"]

    # Connect to the database
    connection = mysql.connector.connect(
        host=host,
        user=user,
        password=password,
        database=database
    )

    # Query to get distinct industries
    query = "SELECT DISTINCT Industry FROM ibis_report"
    df = pd.read_sql(query, connection)
    connection.close()
    return df

# Function to get data for the selected industry
def get_data(industry):
    host = st.secrets["mysql"]["host"]
    user = st.secrets["mysql"]["user"]
    password = st.secrets["mysql"]["password"]
    database = st.secrets["mysql"]["database"]

    # Connect to the database
    connection = mysql.connector.connect(
        host=host,
        user=user,
        password=password,
        database=database
    )

    # Query to get data for the selected industry
    query = f"SELECT * FROM ibis_report WHERE Industry = '{industry}'"
    df = pd.read_sql(query, connection)
    connection.close()
    return df

# Function to create separate charts for each category
def create_category_charts(df):
    category_charts = []

    bar_color = '#032649'
    line_color = '#EB8928'

    for category in df['Category'].unique():
        category_data = df[df['Category'] == category]
        
        # Calculate the change for the category
        category_data['Change'] = category_data['Value'].pct_change() * 100
        
        # Create a subplot with secondary y-axis
        fig = make_subplots(specs=[[{"secondary_y": True}]])

        # Add bar chart for 'Value' on the primary y-axis
        fig.add_trace(
            go.Bar(
                x=category_data['Year'],
                y=category_data['Value'],
                name='Value',
                marker_color=bar_color
            ),
            secondary_y=False
        )

        # Add line chart for 'Change' on the secondary y-axis
        fig.add_trace(
            go.Scatter(
                x=category_data['Year'],
                y=category_data['Change'],
                name='Change (%)',
                mode='lines+markers',
                line=dict(color=line_color)
            ),
            secondary_y=True
        )

        # Update axis titles
        fig.update_layout(
            title_text=f"{category} - Value vs Change",
            xaxis_title="Year",
            yaxis_title="Value",
        )

        # Set secondary y-axis title
        fig.update_yaxes(title_text="Value", secondary_y=False)
        fig.update_yaxes(title_text="Change (%)", secondary_y=True)

        category_charts.append(fig)

    return category_charts

def export_charts_to_ppt(charts, filename="charts.pptx"):
    prs = Presentation()

    for i, chart in enumerate(charts):
        slide = prs.slides.add_slide(prs.slide_layouts[5])  # Use a blank slide layout
        title = slide.shapes.title
        title.text = f"Chart {i + 1}"

        # Save the chart as an image in-memory
        image_stream = io.BytesIO()
        chart.write_image(image_stream, format='png')
        image_stream.seek(0)

        # Add the image to the slide
        slide.shapes.add_picture(image_stream, Inches(1), Inches(1), width=Inches(8), height=Inches(4.5))

    # Save the PowerPoint file to a buffer
    ppt_buffer = io.BytesIO()
    prs.save(ppt_buffer)
    ppt_buffer.seek(0)
    return ppt_buffer



# Streamlit interface
st.title("IBIS - Industry Analysis")

# Get the list of industries
df_industries = get_industries()
industry_options = df_industries["Industry"].tolist()

# Dropdown for industry selection
industry = st.selectbox("Select Industry", industry_options)

# Get the data for the selected industry
df = get_data(industry)

# Create separate charts for each category
if not df.empty:
    charts = create_category_charts(df)
    
    # Display all charts
    for chart in charts:
        st.plotly_chart(chart)
else:
    st.write("No data available for the selected industry.")
