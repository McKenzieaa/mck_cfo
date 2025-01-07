import pandas as pd
import tensorflow as tf
import mysql.connector
import streamlit as st
import plotly.express as px
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode
from pptx import Presentation
from pptx.util import Inches
from io import BytesIO
import os

st.set_page_config(page_title="Precedent Transactions", layout="wide")

# MySQL database connection details
host = st.secrets["mysql"]["host"]
user = st.secrets["mysql"]["user"]
password = st.secrets["mysql"]["password"]
database = st.secrets["mysql"]["database"]

# Connect to the MySQL database
try:
    conn = mysql.connector.connect(
        host=host,
        user=user,
        password=password,
        database=database
    )
except mysql.connector.Error as e:
    st.error(f"Error connecting to MySQL: {e}")
    st.stop()

# Query to fetch the data from the MySQL table
query = """
SELECT 
    `Year`, `Target`, `EV/Revenue`, `EV/EBITDA`, `Business Description`, `Industry`, `Location`
FROM 
    precedent_table
"""

try:
    df = pd.read_sql(query, conn)
except Exception as e:
    st.error(f"Error loading data from MySQL: {e}")
    st.stop()

# Close the MySQL connection
conn.close()

# Debugging: Print column data types
st.write("Data Types before processing:")
st.write(df.dtypes)

# Ensure all columns are TensorFlow-compatible
df['Year'] = df['Year'].fillna(0).astype('int32')
df['EV/Revenue'] = df['EV/Revenue'].fillna(0).astype('float32')
df['EV/EBITDA'] = df['EV/EBITDA'].fillna(0).astype('float32')
df['Target'] = df['Target'].fillna("").astype('string')
df['Business Description'] = df['Business Description'].fillna("").astype('string')
df['Industry'] = df['Industry'].fillna("").astype('string')
df['Location'] = df['Location'].fillna("").astype('string')

# Debugging: Check if there are still problematic types
st.write("Data Types after processing:")
st.write(df.dtypes)

# Convert the DataFrame to a TensorFlow Dataset
try:
    dataset = tf.data.Dataset.from_tensor_slices(dict(df))
    dataset = dataset.batch(32).prefetch(tf.data.AUTOTUNE)
    st.success("TensorFlow Dataset created successfully!")
except Exception as e:
    st.error(f"Error creating TensorFlow Dataset: {e}")
    st.stop()

# Extract unique industries and locations
industries = df['Industry'].unique()
locations = df['Location'].unique()

# Sidebar for selecting industries and locations
col1, col2 = st.columns(2)
selected_industries = col1.multiselect("Select Industry", industries)
selected_locations = col2.multiselect("Select Location", locations)

if selected_industries and selected_locations:
    # Filter data
    filtered_data = df[
        (df['Industry'].isin(selected_industries)) &
        (df['Location'].isin(selected_locations))
    ]
    filtered_data['Year'] = filtered_data['Year'].astype(int)

    # Display filtered data in Ag-Grid
    st.subheader("Precedent Transactions")
    gb = GridOptionsBuilder.from_dataframe(filtered_data)
    gb.configure_selection(selection_mode="multiple", use_checkbox=True)
    grid_options = gb.build()

    grid_response = AgGrid(
        filtered_data,
        gridOptions=grid_options,
        update_mode=GridUpdateMode.SELECTION_CHANGED,
        height=400,
        width='100%',
        theme='streamlit'
    )

    selected_data = pd.DataFrame(grid_response['selected_rows'])
    if not selected_data.empty:
        # Group by Year and calculate averages
        avg_data = selected_data.groupby('Year')[['EV/Revenue', 'EV/EBITDA']].mean().reset_index()

        # Generate bar charts
        fig1 = px.bar(avg_data, x='Year', y='EV/Revenue', title="EV/Revenue")
        st.plotly_chart(fig1)

        fig2 = px.bar(avg_data, x='Year', y='EV/EBITDA', title="EV/EBITDA")
        st.plotly_chart(fig2)

        # Export to PowerPoint
        export_ppt = st.button("Export Charts to PowerPoint")
        if export_ppt:
            template_path = os.path.join(os.getcwd(), "streamlit_dashboard", "data", "main_template_pitch.pptx")

            if not os.path.exists(template_path):
                st.error(f"PowerPoint template not found at: {template_path}")
                st.stop()

            ppt = Presentation(template_path)
            slide_layout = ppt.slide_layouts[5]  
            slide1 = ppt.slides.add_slide(slide_layout)

            # Save EV/Revenue chart to an image
            fig1_image = BytesIO()
            fig1.write_image(fig1_image, format="png", width=900, height=300)
            fig1_image.seek(0)
            slide1.shapes.add_picture(fig1_image, Inches(0.11), Inches(0.90), width=Inches(9), height=Inches(2.8))

            # Save EV/EBITDA chart to an image
            fig2_image = BytesIO()
            fig2.write_image(fig2_image, format="png", width=900, height=300)
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
