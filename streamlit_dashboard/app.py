import streamlit as st
from pages.home import get_home_layout
from pages.public_comps_view import display_public_comps
from pages.transactions_view import display_transactions
from pages.us_indicators_view import get_us_indicators_layout
from pages.us_state_indicators_view import get_state_indicators_layout
from pages.benchmarking_view import get_benchmarking_layout
from pages.presentation_view import presentation_view

# Set up the Streamlit page configuration
st.set_page_config(
    page_title="McK Analysis",
    page_icon="",  # Add path to an icon if available
    layout="wide",
    initial_sidebar_state="expanded"
)

# Dictionary mapping page names to layout functions
PAGES = {
    "Home": (get_home_layout, None),
    "Public Comps": (display_public_comps, None),
    "Precedent Transactions": (display_transactions, None),
    "US Indicators": (get_us_indicators_layout, None),
    "State Indicators": (get_state_indicators_layout, None),
    "Benchmarking": (get_benchmarking_layout, None),
    "Presentation": (presentation_view, None),
}

def render_page(page_name):
    """Render the selected page layout."""
    page_function, _ = PAGES.get(page_name, (None, None))  # Get the function and ignore the second item
    if page_function:
        page_function()  # Call the page function
    else:
        st.error("404: Page not found")

def sidebar_navigation():
    """Generate a sidebar navigation with clickable buttons instead of a dropdown."""
    st.sidebar.title("Navigation")
    # Ensure only the relevant items appear in the sidebar
    selected_page = st.sidebar.radio("", options=list(PAGES.keys()), index=0)
    st.experimental_set_query_params(page=selected_page)  # Store selection in query params
    return selected_page

def get_current_page():
    """Retrieve the current page from the URL parameters."""
    query_params = st.experimental_get_query_params()
    # Default to 'Home' if no page is specified
    return query_params.get("page", ["Home"])[0]

def main():
    """Main entry point for the app.""" 
    current_page = sidebar_navigation()  # Create sidebar links and get the selected page
    render_page(current_page)  # Render the appropriate page layout

# Custom CSS to hide sidebar toggle and ensure a clean layout
custom_css = """
    <style>
        [data-testid="collapsedControl"] {
            display: none;
        }
        section[data-testid="stSidebar"] {
            min-width: 250px;
            max-width: 250px;
            height: 100vh;
        }
        section[data-testid="stSidebar"] div {
            padding-top: 2rem;
        }
    </style>
"""
st.markdown(custom_css, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
