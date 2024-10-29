import streamlit as st
from pages.home import get_home_layout
from pages.public_comps_view import get_public_comps_layout
from pages.transactions_view import get_transaction_layout
from pages.us_indicators_view import get_us_indicators_layout,load_data
from pages.us_state_indicators_view import get_state_indicators_layout
from pages.benchmarking_view import get_benchmarking_layout
from pages.presentation_view import presentation_view

st.set_page_config(
    page_title="McK Analysis",
    page_icon='',
    layout="wide",
    initial_sidebar_state="expanded"
)

def sidebar_navigation():
    """Generate navigation links for the sidebar."""
    st.sidebar.title("Navigation")
    page = st.sidebar.radio(
        "Go to",
        (
            "Home",
            "Public Comps",
            "Precedent Transactions",
            "US Indicators",
            "State Indicators",
            "Benchmarking",
            "Presentation",
            "Industry: Energy",
        ),
    )
    return page

page = sidebar_navigation()

def render_page(page):
    """Render the selected page layout."""
    if page == "Home":
        get_home_layout()
    elif page == "Public Comps":
        get_public_comps_layout()
    elif page == "Precedent Transactions":
        get_transaction_layout()
    elif page == "US Indicators":
        get_us_indicators_layout(), load_data()
    elif page == "State Indicators":
        get_state_indicators_layout()
    elif page == "Benchmarking":
        get_benchmarking_layout()
    elif page == "Presentation":
        presentation_view()
    # elif page == "Industry: Energy":
    #     get_market_size_layout()
    #     get_industry_driver_layout()
    else:
        st.error("404: Page not found")

def main():
    """Main entry point for the app."""
    render_page(page)

if __name__ == "__main__":
    main()
