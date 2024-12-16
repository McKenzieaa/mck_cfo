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

    # Ensure that the default industry exists in the list
    default_industry = "Soyabean Farming" if "Soyabean Farming" in industries else industries[0]
    selected_industries = st.multiselect("Select Industries", industries, default=[default_industry])

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
        filtered_data["Value"] = pd.to_numeric(filtered_data["Value"], errors="coerce")
    except Exception as e:
        st.error(f"Error converting data types: {e}")
        return

    # Handle missing or null values
    if filtered_data[["Year", "Value"]].isnull().any().any():
        st.error("Null or invalid values detected in 'Year' or 'Value'. Please check your data.")
        st.write("Null Values Preview:", filtered_data[filtered_data[["Year", "Value"]].isnull()])
        return

    # Dropdown for selecting industry
    categories = filtered_data['Category'].unique()

    for category in categories:
        st.subheader(f"Charts for Category: {category}")

        # Filter data based on selected category
        category_data = filtered_data[filtered_data['Category'] == category]

        # Group the data by Year and Category
        grouped_data = category_data.groupby(["Year", "Category"], as_index=False).agg({"Value": "mean"})

        # Show the grouped data
        st.write("Grouped Data:", grouped_data)

        # Bar chart for each year with proper 'Year' labels
        st.subheader(f"Bar Chart for Category: {category}")
        fig = px.bar(
            grouped_data,
            x="Year",
            y="Value",
            color="Category",
            title=f"Yearly Values for Category: {category}",
            labels={"Value": "Average Value", "Year": "Year"},
            template="plotly_dark"  # Optional styling for charts
        )
        st.plotly_chart(fig)

        # Calculate and display the Change in Value for each year
        grouped_data["Change"] = grouped_data["Value"]

        # Line chart for the change in values
        st.subheader(f"Change in Value for Category: {category}")
        line_chart = px.line(
            grouped_data,
            x="Year",
            y="Change",
            title=f"Change in Value for Category: {category}",
            labels={"Change": "Change in Value", "Year": "Year"},
            template="plotly_dark"
        )
        st.plotly_chart(line_chart)

if __name__ == "__main__":
    main()
