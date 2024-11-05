import streamlit as st
import pandas as pd
import numpy as np
from pygwalker.api.streamlit import StreamlitRenderer

st.set_page_config(layout="wide")

# Load the data
path_public_comps = r'streamlit_dashboard/data/Public Listed Companies US.xlsx'
df = pd.read_excel(path_public_comps, sheet_name="FY 2023")

# Data processing
df['Enterprise Value (in $)'] = pd.to_numeric(df['Enterprise Value (in $)'], errors='coerce')
df['Revenue (in $)'] = pd.to_numeric(df['Revenue (in $)'], errors='coerce').round(1)
df['EBITDA (in $)'] = pd.to_numeric(df['EBITDA (in $)'], errors='coerce').round(1)
df['EV/Revenue'] = df['Enterprise Value (in $)'] / df['Revenue (in $)']
df['EV/EBITDA'] = df['Enterprise Value (in $)'] / df['EBITDA (in $)']

columns_to_display = ['Name', 'Country', 'Industry', 'EV/Revenue', 'EV/EBITDA', 'Business Description']
df = df[columns_to_display]
df.replace("-", np.nan, inplace=True)
df.dropna(subset=['Country', 'Industry', 'EV/Revenue', 'EV/EBITDA'], inplace=True)
df = df[(~df['EV/Revenue'].isin([np.inf, -np.inf])) & (~df['EV/EBITDA'].isin([np.inf, -np.inf]))]
df = df.dropna(subset=['Country', 'Industry', 'EV/Revenue', 'EV/EBITDA'])
df['EV/Revenue'] = df['EV/Revenue'].apply(lambda x: f"{x:.1f}x")
df['EV/EBITDA'] = df['EV/EBITDA'].apply(lambda x: f"{x:.1f}x")

# Display in Streamlit
st.header("Visualisation")

pyg_app = StreamlitRenderer(df)
 
pyg_app.explorer()
