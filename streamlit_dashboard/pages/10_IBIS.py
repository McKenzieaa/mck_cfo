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

def create_category_charts(df):
    category_charts = []

    bar_color = '#032649'
    line_color = '#EB8928'

    for category in df['Category'].unique():
        category_data = df[df['Category'] == category]
        
        # Calculate the change for the category
        category_data['Change'] = category_data['Value'].pct_change() * 100
        
        # Get the last value for each category
        last_value = category_data['Value'].iloc[-1]
        last_change = category_data['Change'].iloc[-1]

        # Create a subplot with secondary y-axis
        fig = make_subplots(specs=[[{"secondary_y": True}]])

        # Add bar chart for 'Value' on the primary y-axis
        fig.add_trace(
            go.Bar(
                x=category_data['Year'],
                y=category_data['Value'],
                name='Value',
                marker_color=bar_color,
                text=[f"{value}" if i == len(category_data) - 1 else "" for i, value in enumerate(category_data['Value'])],  # Show text only for the last value
                textposition="outside"  # Place text outside the bars
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
                line=dict(color=line_color),
                text=[f"{change:.1f}%" if i == len(category_data) - 1 else "" for i, change in enumerate(category_data['Change'])],  # Show text only for the last value
                textposition="top center"  # Place text above the last marker
            ),
            secondary_y=True
        )

        # Update axis titles
        fig.update_layout(
            # title_text=f"{category} - Value vs Change",
            xaxis_title="Year",
            yaxis_title="Value",
        )

        # Set secondary y-axis title
        fig.update_yaxes(title_text="Value (in bn$)", secondary_y=False)
        fig.update_yaxes(title_text="Change (%)", secondary_y=True)

        # Update the legend position (upper-left)
        fig.update_layout(
            legend=dict(
                x=0, 
                y=1, 
                xanchor='left', 
                yanchor='top'
            ),
            yaxis=dict(showgrid=False),
            margin=dict(l=50, r=50, t=50,b=50),height=400,width=600
        )

        category_charts.append(fig)

    return category_charts

# Function to export charts to PowerPoint
def export_charts_to_ppt(charts, filename="charts.pptx"):
    prs = Presentation()
    
    # Define chart positions (in Inches)
    positions = [
        {"left": Inches(1), "top": Inches(1), "width": Inches(6), "height": Inches(3)},  # Position for Chart 1
        {"left": Inches(1), "top": Inches(4.5), "width": Inches(6), "height": Inches(3)},  # Position for Chart 2
        {"left": Inches(7.5), "top": Inches(1), "width": Inches(6), "height": Inches(3)},  # Position for Chart 3
        {"left": Inches(7.5), "top": Inches(4.5), "width": Inches(6), "height": Inches(3)},  # Position for Chart 4
    ]

    for i, chart in enumerate(charts):
        slide = prs.slides.add_slide(prs.slide_layouts[5])  # Use a blank slide layout

        image_stream = io.BytesIO()
        chart.write_image(image_stream, format='png')
        image_stream.seek(0)

        pos = positions[i % len(positions)]

        slide.shapes.add_picture(
            image_stream,
            pos["left"],
            pos["top"],
            width=pos["width"],
            height=pos["height"]
        )

    ppt_buffer = io.BytesIO()
    prs.save(ppt_buffer)
    ppt_buffer.seek(0)
    return ppt_buffer

st.title("IBIS - Industry Analysis")

df_industries = get_industries()
industry_options = df_industries["Industry"].tolist()
industry = st.selectbox("Select Industry", industry_options)

# Get the data for the selected industry
df = get_data(industry)

# Create separate charts for each category
if not df.empty:
    charts = create_category_charts(df)
    
    # Display all charts
    for chart in charts:
        st.plotly_chart(chart)

    # Display debug message
    st.write("Charts are ready for export!")

    # Add "Export to PowerPoint" button
    if st.button("Export Charts to PowerPoint"):
        ppt_buffer = export_charts_to_ppt(charts)

        # Provide download link for PowerPoint
        st.download_button(
            label="Download PowerPoint",
            data=ppt_buffer,
            file_name="industry_charts.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
else:
    st.write("No data available for the selected industry.")
