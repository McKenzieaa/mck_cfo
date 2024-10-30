import streamlit as st
from streamlit_navigation_bar import st_navbar  # Import the navbar component

# Set up Streamlit page configuration
st.set_page_config(
    page_title="McK Analysis",
    page_icon="",  # Add an icon path if required
    layout="wide",
    initial_sidebar_state="collapsed"  # Start with sidebar collapsed
)

# Create the navbar with the pages you want to navigate between
page = st_navbar(
    ["Home", "Public Comps", "Transactions", "US Indicators", "State Indicators", 
     "Benchmarking", "Presentation"]
)

# Redirect to other pages based on navbar selection
if page == "Public Comps":
    st.switch_page("pages/public_comps_view.py")
elif page == "Transactions":
    st.switch_page("pages/transactions_view.py")
elif page == "US Indicators":
    st.switch_page("pages/us_indicators_view.py")
elif page == "State Indicators":
    st.switch_page("pages/us_state_indicators_view.py")
elif page == "Benchmarking":
    st.switch_page("pages/benchmarking_view.py")
elif page == "Presentation":
    st.switch_page("pages/presentation_view.py")

# Default content for the "Home" page
if page == "Home":
    st.title("Industry Analysis Dashboard")
    st.write(
        "Welcome to the **McKenzie Financial Analysis Dashboard**. "
        "Explore detailed insights into market trends, precedent transactions, "
        "benchmarking data, and economic indicators across industries."
    )

# Hide the sidebar toggle to maintain focus on the navbar
custom_css = """
    <style>
        [data-testid="collapsedControl"] {
            display: none;
        }
    </style>
"""
st.markdown(custom_css, unsafe_allow_html=True)
