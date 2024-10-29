import streamlit as st

def get_home_layout():
    # Set the title and introduction text
    st.title("📊 Industry Analysis Dashboard")
    st.write("""
    Welcome to the **McKenzie Financial Analysis Dashboard**. 
    Explore detailed insights into market trends, precedent transactions, benchmarking data, 
    and economic indicators across industries.
    """)

    # Create two columns for layout
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("💼 Public Companies")
        st.write("Analyze financial metrics and KPIs of public companies.")
        if st.button("Go to Public Comps"):
            st.experimental_set_query_params(page="Public Comps")

        st.subheader("🔄 Precedent Transactions")
        st.write("Explore historical transaction data and benchmarks.")
        if st.button("Go to Precedent Transactions"):
            st.experimental_set_query_params(page="Precedent Transactions")

    with col2:
        st.subheader("🌍 Economic Indicators")
        st.write("Review national and state-level economic indicators.")
        if st.button("Go to State Indicators"):
            st.experimental_set_query_params(page="State Indicators")

        st.subheader("📈 Benchmarking")
        st.write("Benchmark industry performance and key metrics.")
        if st.button("Go to Benchmarking"):
            st.experimental_set_query_params(page="Benchmarking")

    # Optional: Add presentation view link at the bottom
    st.markdown("---")
    st.subheader("📊 Presentation View")
    st.write("Summarize and visualize all insights in a presentation format.")
    if st.button("Go to Presentation"):
        st.experimental_set_query_params(page="Presentation")
