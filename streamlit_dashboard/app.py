import streamlit as st
from pages.home import get_home_layout
from pages.public_comps_view import get_public_comps_layout
from pages.transactions_view import get_transaction_layout
# from pages.us_indicators_view import get_us_indicators_layout, load_data
from pages.us_state_indicators_view import get_state_indicators_layout
from pages.benchmarking_view import get_benchmarking_layout
from pages.presentation_view import presentation_view

# Set up the Streamlit page configuration
st.set_page_config(
    page_title="McK Analysis",
    page_icon="",  # Add path to an icon if available
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Dictionary mapping page names to layout functions
PAGES = {
    "Home": get_home_layout,
    "Public Comps": get_public_comps_layout,
    "Precedent Transactions": get_transaction_layout,
    # "US Indicators": lambda: (get_us_indicators_layout, load_data),
    "State Indicators": get_state_indicators_layout,
    "Benchmarking": get_benchmarking_layout,
    "Presentation": presentation_view,
}

def render_page(page_name):
    """Render the selected page layout."""
    page_function = PAGES.get(page_name)
    if page_function:
        page_function()
    else:
        st.error("404: Page not found")

def sidebar_navigation():
    """Generate sidebar navigation with a selectbox."""
    st.sidebar.title("Navigation")
    selected_page = st.sidebar.selectbox("Select a Page", options=PAGES.keys())
    # Store the selected page in the query parameters
    st.experimental_set_query_params(page=selected_page)
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

# Hide the sidebar toggle using custom CSS
hide_sidebar_style = """
    <style>
        [data-testid="collapsedControl"] {
            display: none;
        }
    </style>
"""
st.markdown(hide_sidebar_style, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
