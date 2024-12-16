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

# Function to create separate charts for each category
def create_category_charts(df):
    category_charts = []
    
    for category in df['Category'].unique():
        category_data = df[df['Category'] == category]
        
        # Calculate the change for the category
        category_data['Change'] = category_data['Value'].pct_change() * 100
        
        # Create a bar chart with a line showing the change for this category
        fig = px.bar(category_data, x='Year', y='Value', color='Category', title=f"{category} - Value vs Change",
                     labels={'Value': 'Value', 'Year': 'Year'})
        
        # Add a line for the change percentage for this category
        fig.add_scatter(x=category_data['Year'], y=category_data['Change'], mode='lines', name=f'{category} Change')

        category_charts.append(fig)
    
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
