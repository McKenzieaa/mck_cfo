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

    # Multi-select dropdown for industries with default value "Soyabean Farming"
    industries = data['Industry'].unique()
    # Set "Soyabean Farming" as the default if it exists in the list
    default_industry = ["Soyabean Farming"] if "Soyabean Farming" in industries else []
    selected_industries = st.multiselect("Select Industries", industries, default=default_industry)

    if not selected_industries:
        st.warning("Please select at least one industry.")
        return

    # Filter data based on selected industries
    filtered_data = data[data['Industry'].isin(selected_industries)]

    # Debugging: Show filtered data structure
    st.write("Filtered Data Preview:", filtered_data.head())

    # Check for required columns
    required_columns = ["Year", "Value", "Category"]
    if not all(col in filtered_data.columns for col in required_columns):
        st.error(f"Required columns {required_columns} are missing in the data.")
        return

    # Ensure data types are correct
    try:
        filtered_data["Year"] = pd.to_datetime(filtered_data["Year"], errors="coerce").dt.year
        filtered_data["Year"] = filtered_data["Year"].astype('Int64')  # Ensure Year is an integer
        filtered_data["Value"] = pd.to_numeric(filtered_data["Value"], errors="coerce")
    except Exception as e:
        st.error(f"Error converting data types: {e}")
        return

    # Handle missing or null values
    if filtered_data[["Year", "Value"]].isnull().any().any():
        st.error("Null or invalid values detected in 'Year' or 'Value'. Please check your data.")
        st.write("Null Values Preview:", filtered_data[filtered_data[["Year", "Value"]].isnull()])
        return

    # Fill missing values with zeros (or another strategy)
    filtered_data.fillna({"Value": 0}, inplace=True)

    # Group data by Year and Category, then sum 'Value'
    grouped_data = filtered_data.groupby(["Year", "Category"], as_index=False).agg({"Value": "sum"})

    # Unique categories
    categories = grouped_data['Category'].unique()

    # Dropdown for category selection
    selected_category = st.selectbox("Select a Category", categories)

    if selected_category:
        # Filter data based on selected category
        category_data = grouped_data[grouped_data['Category'] == selected_category]

        # Generate bar chart based on selected category
        st.subheader(f"Bar Chart for Category: {selected_category}")
        try:
            fig = px.bar(
                category_data,
                x="Year",
                y="Value",
                title=f"Yearly Values for Category: {selected_category}",
                labels={"Value": "Total Value", "Year": "Year"},
                template="plotly_dark"  # Optional styling for charts
            )
            st.plotly_chart(fig)
        except ValueError as e:
            st.error(f"Error creating the bar chart for {selected_category}: {e}")
            st.write("Debugging Data:", category_data)

if __name__ == "__main__":
    main()
