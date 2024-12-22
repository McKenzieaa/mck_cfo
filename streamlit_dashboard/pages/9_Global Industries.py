import streamlit as st
import pandas as pd
import requests
import plotly.express as px
import zipfile
import io
from io import BytesIO
from pptx import Presentation
from pptx.util import Inches
import plotly.io as pio
from io import StringIO
import os

st.set_page_config(page_title="Global Industry Analysis", layout="wide")
    
category_data = [
    ('22T', 'Utilities', 'QREV', 'QSS'),
    ('2211T', 'Electric Power Generation, Transmission and Distribution', 'QREV', 'QSS')
]

categories = {
    category.strip(): {"cat_code": cat_code, "data_type": data_type, "program": program}
    for cat_code, category, data_type, program in category_data
}

selected_categories = ['Electric Power Generation, Transmission and Distribution']

def get_market_size_data(selected_categories):
    """Fetch and prepare market size data based on selected categories."""
    all_data = pd.DataFrame()

    for category in selected_categories:
        category = category.strip()
        try:
            cat_data = categories[category]
        except KeyError:
            continue

        url = f"https://www.census.gov/econ_export/?format=xls&mode=report&default=false&errormode=Dep&charttype=&chartmode=&chartadjn=&submit=GET+DATA&program={cat_data['program']}&startYear=2014&endYear=2024&categories%5B0%5D={cat_data['cat_code']}&dataType={cat_data['data_type']}&geoLevel=US&adjusted=false&notAdjusted=true&errorData=false&vert=1"
        response = requests.get(url)
        if response.status_code != 200:
            continue

        marketsize_data = BytesIO(response.content)
        df = pd.read_excel(marketsize_data)
        df = df.iloc[6:].reset_index(drop=True)
        df.columns = df.iloc[0]
        df = df.drop(0).reset_index(drop=True)
        df = df.dropna(axis=1, how='all')
        df['Value'] = pd.to_numeric(df['Value'].replace("N", pd.NA), errors='coerce')
        df[['Quarter', 'Year']] = df['Period'].str.split('-', expand=True)
        df['Quarter'] = df['Quarter'].replace({
            '1st Quarter': 'Q1', '2nd Quarter': 'Q2',
            '3rd Quarter': 'Q3', '4th Quarter': 'Q4'
        })
        df['Quarter_Year'] = df['Quarter'] + ' ' + df['Year'].astype(str)
        df['Category'] = category
        all_data = pd.concat([all_data, df], ignore_index=True)

    if all_data.empty:
        st.write("No data available for the selected categories.")
        return None

    return all_data

# Fetch Energy Data
url = "https://www.eia.gov/totalenergy/data/monthly/Zip_Excel_Month_end/MER_2024_09.zip"
response = requests.get(url)
url2 = "https://nyc3.digitaloceanspaces.com/owid-public/data/energy/owid-energy-data.csv"

with zipfile.ZipFile(io.BytesIO(response.content)) as z:
    with z.open("Table 07.01.xlsx") as f1:
        df_electricity_end_use = pd.read_excel(f1, sheet_name="Annual Data", skiprows=8)
    with z.open("Table 09.08.xlsx") as f2:
        df_avg_price = pd.read_excel(f2, sheet_name="Annual Data", skiprows=8)

df_electricity_end_use.rename(columns={df_electricity_end_use.columns[0]: "Year"}, inplace=True)
df_avg_price.rename(columns={df_avg_price.columns[0]: "Year"}, inplace=True)

df_electricity_gen = pd.read_csv(url2)
df_renew_share = df_electricity_gen.dropna(subset=['renewables_share_elec'])

# Per Capita Electricty Data
df_electricity_gen = pd.read_csv(url2)
df_per_cap_elec_gen = df_electricity_gen.dropna(subset=['fossil_elec_per_capita', 'nuclear_elec_per_capita', 'renewables_elec_per_capita'])
df_per_cap_elec_gen = df_per_cap_elec_gen[df_per_cap_elec_gen['year'] == 2023]
# df_per_cap_elec_gen['total_elec_per_capita'] = (
#     df_per_cap_elec_gen['fossil_elec_per_capita'] + df_per_cap_elec_gen['nuclear_elec_per_capita'] + df_per_cap_elec_gen['renewables_elec_per_capita']
# )
# top_10_countries = df_per_cap_elec_gen.nlargest(10, 'total_elec_per_capita')
# df_per_cap_elec_gen = top_10_countries.melt(
#     id_vars=['country'],
#     value_vars=['fossil_elec_per_capita', 'nuclear_elec_per_capita', 'renewables_elec_per_capita'],
#     var_name='Energy Source',
#     value_name='Per Capita Generation'
# )

# df_per_cap_elec_gen['Energy Source'] = df_per_cap_elec_gen['Energy Source'].replace({
#     'fossil_elec_per_capita': 'Fossil',
#     'nuclear_elec_per_capita': 'Nuclear',
#     'renewables_elec_per_capita': 'Renewables'
# })
# df_per_cap_elec_gen_pivot = df_per_cap_elec_gen.pivot(index='country', columns='Energy Source', values='Per Capita Generation')

selected_countries = ['China', 'India', 'World', 'Japan','Brazil','France', 'United States']
df_per_cap_elec_gen = df_per_cap_elec_gen[df_per_cap_elec_gen['country'].isin(selected_countries)]

df_per_cap_elec_gen = df_per_cap_elec_gen.melt(
    id_vars=['country'],
    value_vars=['fossil_elec_per_capita', 'nuclear_elec_per_capita', 'renewables_elec_per_capita'],
    var_name='Energy Source',
    value_name='Per Capita Generation'
)

df_per_cap_elec_gen['Energy Source'] = df_per_cap_elec_gen['Energy Source'].replace({
    'fossil_elec_per_capita': 'Fossil',
    'nuclear_elec_per_capita': 'Nuclear',
    'renewables_elec_per_capita': 'Renewables'
})
# Energy Consumption
ene_cons = "https://www.eia.gov/totalenergy/data/browser/csv.php?tbl=T07.06"
df_ene_cons = pd.read_csv(ene_cons)
df_ene_cons['Description'] = df_ene_cons['Description'].str.split(',', n=1).str[1]
df_ene_cons = df_ene_cons[df_ene_cons['Description'].str.contains("Residential|Transportation|Industrial|Commercial", case=False, na=False)]
df_ene_cons['Year'] = df_ene_cons['YYYYMM'].astype(str).str[:4]
df_ene_cons = df_ene_cons[['Year', 'Description', 'Value']]
df_ene_cons['Value'] = df_ene_cons['Value'].round(1)
# df_ene_cons = df_ene_cons.groupby(['Year', 'Description'], as_index=False).sum()

# Share of electricity production from renewables
share_elec_prod = "https://ourworldindata.org/grapher/share-electricity-renewables.csv?v=1&csvType=full&useColumnShortNames=true"

try:
    response = requests.get(share_elec_prod)
    response.raise_for_status() 
    csv_data = StringIO(response.text)
    df_share_elec_prod = pd.read_csv(csv_data)
except requests.exceptions.RequestException as e:
    print(f"Failed to fetch data: {e}")
    exit()

df_share_elec_prod.rename(columns={'Entity': 'Countries'}, inplace=True)
filt_share_elec_prod = df_share_elec_prod[df_share_elec_prod['Countries'].isin(selected_countries)]

if 'Year' in filt_share_elec_prod.columns and 'renewable_share_of_electricity__pct' in filt_share_elec_prod.columns:

    filt_share_elec_prod = filt_share_elec_prod.dropna(subset=['Year', 'renewable_share_of_electricity__pct'])
    filt_share_elec_prod['Year'] = pd.to_numeric(filt_share_elec_prod['Year'], errors='coerce')
    filt_share_elec_prod['renewable_share_of_electricity__pct'] = pd.to_numeric(
        filt_share_elec_prod['renewable_share_of_electricity__pct'], errors='coerce'
    )
    filt_share_elec_prod = filt_share_elec_prod.dropna(subset=['Year', 'renewable_share_of_electricity__pct'])
else:
    raise ValueError("Required columns 'Year' or 'renewable_share_of_electricity__pct' are missing from the DataFrame.")


# ENERGY
st.markdown("<h2 style='font-weight: bold; font-size:24px;'>Energy</h2>", unsafe_allow_html=True)
with st.expander("", expanded=True): 
    market_data = get_market_size_data(selected_categories)
    if market_data is not None:
        yearly_data = market_data.groupby(['Year', 'Category'], as_index=False).agg({'Value': 'mean'})
        fig1 = px.bar(
            yearly_data,
            x='Year', y='Value', color='Category',
            title="Market Size",
            labels={'Value': 'Market-Size (in millions)', 'Year': ''},
            color_discrete_sequence=["#0068c9"]
        )

        fig1.update_layout(
            legend=dict(
                x=0,  # Position at the left
                y=1,  # Position at the top
                title="",
                xanchor='left', 
                yanchor='top',
                font=dict(size=8)  # Set font size to 8
            )
        )
        # st.plotly_chart(fig1)

    fig2 = px.line(
        df_electricity_end_use, x="Year", y=df_electricity_end_use.columns[1],
        title="Electricity End Use (Billion Kilowatthours)"
    )
    fig2.update_traces(line_color="#0068c9")
    st.plotly_chart(fig2)

    # Average Price of Electricity Chart
    fig3 = px.line(
        df_avg_price, x="Year", y=df_avg_price.columns[1],
        title="Average Price of Electricity (Cents per Kilowatthour)"
    )
    fig3.update_traces(line_color="#0068c9")
    # st.plotly_chart(fig3)

    # Electricity Generation Map
    # st.sidebar.header("Electricity Generation")
    selected_year =(2023) #st.sidebar.slider("Select Year", 2000, 2023, 2023)
    df_selected_year = df_electricity_gen[df_electricity_gen["year"] == selected_year]
    fig4 = px.choropleth(
        df_selected_year,
        locations='country',
        locationmode='country names',
        color='electricity_generation',
        title=f'Electricity Generation by Country ({selected_year})',
        labels={'electricity_generation': 'Electricity Generation (GWh)'},
        color_continuous_scale="Blues"  # Example color scale; you can choose others like "Plasma", "Blues", etc.
    )

    # Optional: Update layout to fine-tune the color bar
    fig4.update_layout(
        coloraxis_colorbar=dict(
            title="Electricity Generation (GWh)",
            tickvals=[df_selected_year['electricity_generation'].min(), df_selected_year['electricity_generation'].max()],
            ticks="outside"
        )
    )

    # st.plotly_chart(fig4)

    # Renewable Share of Electricity
    # st.sidebar.header("Renewable Share Selection")
    selected_countries = ["World"]
    # st.sidebar.multiselect('Select Countries', df_renew_share['country'].unique(), default=["World"] )
    if selected_countries:
        filtered_df = df_renew_share[df_renew_share["country"].isin(selected_countries)]
        fig5 = px.line(
            filtered_df,
            x="year", y="renewables_share_elec", color="country",
            title="Renewable Share of Electricity"
        )
        fig5.update_traces(line_color="#0068c9")
        # st.plotly_chart(fig5)

    # Solar Projects Coming Up Next 12 Months
    # st.sidebar.header("Map of Solar Projects Coming Up Next 12 Months")
    solar_url = "https://www.eia.gov/electricity/monthly/epm_table_grapher.php?t=table_6_05"
    solar_data = pd.read_html(solar_url)[1]
    st.dataframe(solar_data)

    # st.sidebar.header("Per Capita Electricity")
    fig6 = px.bar(
        df_per_cap_elec_gen,
        x='Per Capita Generation',
        y='country',
        color='Energy Source',
        orientation='h',  # Horizontal orientation
        title='Electricity Generation per Capita by Energy Source (Major Countries in 2023)',
        labels={'country': 'Country', 'Per Capita Generation': 'Percentage of Total Generation'},
        color_discrete_map={
            'Fossil': '#0068c9',        
            'Nuclear': '#FFA500',      
            'Renewables': '#1C798A'     
        }
    )
    fig6.update_layout(barmode='stack')
    fig6.update_xaxes(title_text="")
    # st.plotly_chart(fig6)
    # Additional Bar Chart (Fig 7)
 
    fig7 = px.bar(
        df_ene_cons,
        x='Year', y='Value', color='Description',
        title='Energy Source Distribution Over Years',
        labels={'Value': 'Energy Value', 'Year': 'Year'}
    )

    fig8 = px.line(
        filt_share_elec_prod,
        x='Year',
        y='renewable_share_of_electricity__pct',
        color='Countries',
        title='Renewables as a Percentage of Electricity Production',
        labels={
            'Year': 'Year',
            'renewable_share_of_electricity__pct': 'Renewables - % Electricity',
            'Countries': 'Countries'
        }
    )

    # Format y-axis as a percentage
    fig8.update_yaxes(tickformat=".1%")

    col1, col2 = st.columns(2)

    with col1:
        st.plotly_chart(fig1, use_container_width=True)
        st.plotly_chart(fig2, use_container_width=True)
        st.plotly_chart(fig3, use_container_width=True)
        st.plotly_chart(fig7, use_container_width=True)

    with col2:
        st.plotly_chart(fig4, use_container_width=True)
        st.plotly_chart(fig5, use_container_width=True)
        st.plotly_chart(fig6, use_container_width=True)
        st.plotly_chart(fig8, use_container_width=True)
    
    

    st.write("<h3 style='font-weight: bold; font-size:24px;'>Value Chain</h3>", unsafe_allow_html=True)
    st.image("https://www.energy-uk.org.uk/wp-content/uploads/2023/04/EUK-Different-parts-of-energy-market-diagram.webp", use_container_width=False)

st.markdown("<h2 style='font-weight: bold; font-size:24px;'>Agriculture</h2>", unsafe_allow_html=True)
with st.expander("", expanded=False): 
    st.write("Agriculture-related analysis and visualizations go here.")

st.markdown("<h2 style='font-weight: bold; font-size:24px;'>Technology</h2>", unsafe_allow_html=True)
with st.expander("", expanded=False):
    st.write("Technology-related analysis and visualizations go here.")

st.markdown("<h2 style='font-weight: bold; font-size:24px;'>Automobiles</h2>", unsafe_allow_html=True)
with st.expander("", expanded=False):
    st.write("Automobiles-related analysis and visualizations go here.")

def export_to_pptx(fig1, fig2, fig3, fig4, fig5, fig6, fig7, fig8, value_chain_image_path):
    prs = Presentation()
    slide_layout = prs.slide_layouts[5]

    def add_slide_with_chart(prs, fig, title_text):
        slide = prs.slides.add_slide(slide_layout)
        title = slide.shapes.title
        title.text = title_text
        img_stream = BytesIO()
        fig.write_image(img_stream, format="png", engine="kaleido")
        img_stream.seek(0)
        slide.shapes.add_picture(img_stream, Inches(1), Inches(1), width=Inches(8))

    # Add slides with charts
    add_slide_with_chart(prs, fig1, "Market Size - Yearly")
    add_slide_with_chart(prs, fig2, "Electricity End Use")
    
    # Add value chain image to slide 3
    slide = prs.slides.add_slide(slide_layout)
    slide.shapes.title.text = "Value Chain Analysis"
    
    # Check if image exists and read it into BytesIO
    if os.path.exists(value_chain_image_path):
        img_stream = BytesIO()
        with open(value_chain_image_path, "rb") as img_file:
            img_stream.write(img_file.read())
        img_stream.seek(0)
        slide.shapes.add_picture(img_stream, Inches(1), Inches(1), width=Inches(8))
    else:
        raise FileNotFoundError(f"The image at {value_chain_image_path} was not found. Please check the path.")

    # Add other slides with charts
    add_slide_with_chart(prs, fig3, "Average Price of Electricity")
    add_slide_with_chart(prs, fig4, f"Electricity Generation by Country ({selected_year})")
    add_slide_with_chart(prs, fig5, "Renewable Share of Electricity")
    add_slide_with_chart(prs, fig6, "Per Capita Electricity-2023")
    add_slide_with_chart(prs, fig7, "Energy Source Consumption")
    add_slide_with_chart(prs, fig8, "Share of electricity production from renewables")

    pptx_stream = BytesIO()
    prs.save(pptx_stream)
    pptx_stream.seek(0)
    return pptx_stream

def export_chart_options(fig1, fig2, fig3, fig4, fig5, fig6, fig7, fig8, value_chain_image_path):
    if st.button("Export Charts to PowerPoint"):
        try:
            pptx_file = export_to_pptx(fig1, fig2, fig3, fig4, fig5, fig6, fig7, fig8, value_chain_image_path)
            st.download_button(
                label="Download PowerPoint",
                data=pptx_file,
                file_name="Energy_Industry_Analysis_Report.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
        except FileNotFoundError as e:
            st.error(f"Error: {e}")

# Example of calling the function with corrected file path
value_chain_image_path = r"streamlit_dashboard\data\value_chain.png"  # Use raw string or correct the slashes

# Example call to the function (ensure fig1, fig2, etc. are defined)
export_chart_options(fig1, fig2, fig3, fig4, fig5, fig6, fig7, fig8, value_chain_image_path)