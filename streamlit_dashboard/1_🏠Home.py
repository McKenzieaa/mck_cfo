import streamlit as st

# Ensure that page configuration is the first Streamlit command
st.set_page_config(page_title="Finance Insights", layout="centered", page_icon="💼")

# Custom CSS for styling
def local_css(file_name):
    with open(file_name) as f:
        st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)

# Load CSS
local_css("McKenzieaa/mck_cfo/streamlit_dashboard/assets/icons/style.css")

# Header Section
st.markdown("""
# 💼 Finance Insights  
**Your Partner in Data-Driven Financial Solutions**
""", unsafe_allow_html=True)

# Introductory Section
st.write(
    """
Welcome to Finance Insights, where we empower businesses with deep financial analysis. 
Our services provide actionable insights through data analytics, predictive modeling, and 
strategic advisory—helping you make well-informed financial decisions.
"""
)

# Key Services Section
st.markdown("### Our Expertise:")
col1, col2 = st.columns(2)

with col1:
    st.markdown("""
    - **Financial Forecasting**  
      Predict and plan for the future with confidence.
    - **Risk Analysis**  
      Minimize risks through smart analytics.
    - **Performance Benchmarking**  
      Compare your financial metrics with industry standards.
    """)

with col2:
    st.markdown("""
    - **Investment Strategies**  
      Get custom portfolio optimization.
    - **Cost Analysis**  
      Identify cost-saving opportunities.
    - **Profitability Analysis**  
      Maximize profit potential.
    """)

# Call to Action Section
st.markdown(
    """
    **Ready to elevate your financial strategies?**  
    [Contact us](mailto:info@financeinsights.com) today and discover how we can drive your success.
    """
)

# Footer Section
st.markdown("""
---
© 2024 Finance Insights | **Privacy Policy** | **Terms of Service**
""", unsafe_allow_html=True)
