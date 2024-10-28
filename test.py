import plotly.graph_objs as go
import pandas as pd
import streamlit as st
import os

# # Ensure the 'assets' folder exists
# if not os.path.exists("assets"):
#     os.makedirs("assets")

def fig_to_image(fig, filename):
    """Save the Plotly figure as an image in the assets folder."""
    image_path = f"assets/{filename}.png"
    try:
        with st.spinner("Saving image..."):
            # Save the figure using the Kaleido engine
            fig.write_image(image_path, format="png", engine="kaleido")
        st.success(f"Image saved at: {image_path}")
        return image_path  # Return the path to the saved image
    except Exception as e:
        st.error(f"Failed to save image: {str(e)}")
        return None

def plot_external_driver(selected_indicators):
    """Plot external driver indicators and return the Plotly figure."""
    if len(selected_indicators) == 0:
        selected_indicators = ["World GDP"]  # Default indicator

    fig = go.Figure()

    # Plot selected external indicators
    for indicator in selected_indicators:
        indicator_data = external_driver_df[external_driver_df['Indicator'] == indicator]

        if '% Change' not in indicator_data.columns:
            raise ValueError(f"Expected '% Change' column not found in {indicator}")

        fig.add_trace(
            go.Scatter(
                x=indicator_data['Year'],
                y=indicator_data['% Change'],
                mode='lines',
                name=indicator
            )
        )

    # Update layout
    fig.update_layout(
        title='External Driver Indicators',
        xaxis=dict(title=''),
        yaxis=dict(title='Percent Change'),
        hovermode='x'
    )

    st.plotly_chart(fig)
    return fig

# Example DataFrame for demonstration (replace with actual DataFrame)
external_driver_df = pd.DataFrame({
    'Indicator': ['World GDP', 'Oil Prices', 'World GDP', 'Oil Prices'],
    'Year': [2000, 2000, 2001, 2001],
    '% Change': [3.1, 2.4, 2.9, 5.0]
})

# Get unique indicators and plot the figure
selected_indicators = external_driver_df["Indicator"].unique()
external_driver_fig = plot_external_driver(selected_indicators)

# Button to save the image
if st.button("Save Image"):
    image_path = fig_to_image(external_driver_fig, "external_driver")

    # # Optionally display the saved image if successful
    # if image_path:
    #     st.image(image_path, caption="External Driver Indicators", use_column_width=True)
