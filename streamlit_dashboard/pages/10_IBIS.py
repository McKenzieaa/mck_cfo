import streamlit as st
import pandas as pd
import mysql.connector
import plotly.express as px

# Load MySQL credentials from Streamlit secrets
host = st.secrets["mysql"]["host"]
user = st.secrets["mysql"]["user"]
password = st.secrets["mysql"]["password"]
database = st.secrets["mysql"]["database"]

# Function to establish MySQL connection
def get_connection():
    try:
        return mysql.connector.connect(
            host=host,
            user=user,
            password=password,
            database=database
        )
    except mysql.connector.Error as e:
        st.error(f"Error connecting to MySQL: {e}")
        st.stop()

# Function to fetch data from the MySQL table
def fetch_data():
    connection = get_connection()
    query = "SELECT * FROM ibis_report"
    try:
        data = pd.read_sql(query, connection)
    finally:
        connection.close()
    return data

# Streamlit App
def main():
    st.title("Industry Category Analysis")

    # Fetch data
    data = fetch_data()

    if data.empty:
        st.warning("No data found in the 'ibis_report' table.")
        return

    # Multi-select dropdown for industries
    industries = data['Industry'].unique()
    selected_industries = st.multiselect("Select Industries", industries, default=industries)

    if not selected_industries:
        st.warning("Please select at least one industry.")
        return

    # Filter data based on selected industries
    filtered_data = data[data['Industry'].isin(selected_industries)]

    # Generate bar charts for each category
    categories = [col for col in data.columns if col not in ['Industry', 'Category']]
    for category in categories:
        st.subheader(f"Bar Chart for {category}")
        fig = px.bar(
            filtered_data,
            x="Category",
            y=category,
            color="Industry",
            barmode="group",
            title=f"{category} by Category and Industry"
        )
        st.plotly_chart(fig)

if __name__ == "__main__":
    main()
