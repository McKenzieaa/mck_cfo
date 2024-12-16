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

    # Debugging: Show filtered data structure
    st.write("Filtered Data Preview:", filtered_data.head())

    # Check if required columns exist and have correct types
    required_columns = ["Year", "Value", "Business"]
    if not all(col in filtered_data.columns for col in required_columns):
        st.error(f"Required columns {required_columns} are missing in the data.")
        return

    # Ensure 'Year' and 'Value' are of correct types
    try:
        filtered_data["Year"] = pd.to_datetime(filtered_data["Year"], errors="coerce").dt.year
        filtered_data["Value"] = pd.to_numeric(filtered_data["Value"], errors="coerce")
    except Exception as e:
        st.error(f"Error processing data types: {e}")
        return

    # Check for null values after conversion
    if filtered_data["Year"].isnull().any() or filtered_data["Value"].isnull().any():
        st.error("Null values found in 'Year' or 'Value' after type conversion. Please check your data.")
        return

    # Generate bar chart
    st.subheader("Bar Chart for Yearly Values by Business")
    fig = px.bar(
        filtered_data.dropna(subset=["Year", "Value"]),
        x="Year",
        y="Value",
        color="Business",
        barmode="group",
        title="Yearly Values by Business"
    )
    st.plotly_chart(fig)
