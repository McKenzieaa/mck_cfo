import streamlit as st
import pandas as pd
import mysql.connector
import plotly.express as px

# Connect to the MySQL database using credentials from st.secrets
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

    query = f"SELECT * FROM ibis_report WHERE industry = '{industry}'"
    df = pd.read_sql(query, connection)
    connection.close()
    return df

# Function to create a bar chart with a line for change
def create_chart(df):
    # Calculate the change for each category
    df['Change'] = df.groupby('Category')['Value'].pct_change() * 100

    # Create a bar chart with a line showing the change
    fig = px.bar(df, x='Year', y='Value', color='Category', title="Industry Data",
                 labels={'Value': 'Value', 'Year': 'Year'})
    
    # Add a line for the change percentage
    for category in df['Category'].unique():
        category_data = df[df['Category'] == category]
        fig.add_scatter(x=category_data['Year'], y=category_data['Change'], mode='lines', name=f'{category} Change')

    return fig

# Streamlit interface
st.title("Industry Data Visualization")

# Dropdown for industry selection
industry_options = ["Industry 1", "Industry 2", "Industry 3"]  # Replace with actual industry names from your database
industry = st.selectbox("Select Industry", industry_options)

# Get the data based on the selected industry
df = get_data(industry)

# Create the chart
if not df.empty:
    fig = create_chart(df)
    st.plotly_chart(fig)
else:
    st.write("No data available for the selected industry.")
