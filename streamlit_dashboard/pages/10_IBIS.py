import streamlit as st
import pandas as pd
import mysql.connector
import plotly.express as px

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
    
    # Define the colors
    bar_color = '#032649'  # Dark blue for bars
    line_color = '#EB8928'  # Orange for lines
    
    for category in df['Category'].unique():
        category_data = df[df['Category'] == category]
        category_data['Change'] = df['Change'] 
        
        # Create a bar chart with a dynamic x-axis and custom color
        fig = px.bar(
            category_data, 
            x='Year', 
            y='Value', 
            color='Category',  # Still allow category distinction for color
            title=f"{category} - Value vs Change",
            labels={'Value': 'Value', 'Year': 'Year'},
            color_discrete_sequence=[bar_color]  # Set dark blue color for bars
        )
        
        # Add a line for the change percentage with the specified orange color
        fig.add_scatter(
            x=category_data['Year'], 
            y=category_data['Change'], 
            mode='lines', 
            name=f'{category} Change',
            line=dict(color=line_color)  # Set orange color for the line
        )
        
        # Ensure x-axis is dynamic (auto)
        fig.update_layout(
            xaxis=dict(automargin=True, title='Year'),  # Ensure x-axis labels adjust dynamically
            yaxis=dict(title='Value'),
            title=dict(x=0.5),  # Center-align the chart title
        )

        category_charts.append((category, fig))
    
    return category_charts

# Streamlit interface
st.title("Industry Data Visualization")

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
