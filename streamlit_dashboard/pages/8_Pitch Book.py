import requests
import numpy as np
import zipfile
import io
import dask.dataframe as dd
import plotly.graph_objs as go
import streamlit as st
import plotly.express as px
import pandas as pd
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode
from pptx import Presentation
from pptx.util import Inches
from io import BytesIO
from plotly.subplots import make_subplots
import s3fs  # For accessing S3 data
import os
from datetime import date

today = date.today().strftime("%Y-%m-%d")
state_gdp_data = None
try:
    storage_options = {
        'key': st.secrets["aws"]["AWS_ACCESS_KEY_ID"],
        'secret': st.secrets["aws"]["AWS_SECRET_ACCESS_KEY"],
        'client_kwargs': {'region_name': st.secrets["aws"]["AWS_DEFAULT_REGION"]}
    }
except KeyError:
    st.error("AWS credentials are not configured correctly in Streamlit secrets.")
    st.stop()

def update_figure_slide(ppt, title, fig, slide_number, width, height, left, top):
    if fig is None:
        print(f"Skipping slide '{title}' because the figure is None.")
        return  # Skip if fig is None

    # Get the slide corresponding to the slide number (index starts at 0)
    slide = ppt.slides[slide_number]  # Adjust for 0-based index

    # # Set slide title (optionally adjust placement based on the layout)
    # title_shape = slide.shapes.title
    # title_shape.text = f"Slide {slide_number}: {title}"  # Add slide number to title

    # Save the figure image to a BytesIO object (no size, position parameters here)
    fig_image = BytesIO()
    fig.write_image(fig_image, format="png")  # Only pass the format here
    fig_image.seek(0)

    # Use Inches for size and position only in the add_picture() function
    slide.shapes.add_picture(fig_image, Inches(left), Inches(top), Inches(width), Inches(height))
    fig_image.close()

def export_all_to_pptx(labour_fig_us, external_fig, gdp_fig_us, cpi_ppi_fig_us, fig1_precedent, fig2_precedent, fig1_public, fig2_public, labour_fig, gdp_fig):
    # Load the custom template
    template_path = os.path.join(os.getcwd(), "streamlit_dashboard", "data", "main_template_pitch.pptx")
    ppt = Presentation(template_path)  # Load the template

    # Use the existing slides (slide_number corresponds to the slide index)
    update_figure_slide(ppt, "Precedent - EV/Revenue", fig1_precedent, slide_number=10, width=9, height=3, left=0.11, top=0.90)
    update_figure_slide(ppt, "Precedent - EV/EBITDA", fig2_precedent, slide_number=10, width=9, height=3, left=0.11, top=3.70)
    update_figure_slide(ppt, "Public Comps - EV/Revenue", fig1_public, slide_number=11, width=9, height=3, left=0.11, top=0.90)
    update_figure_slide(ppt, "Public Comps - EV/EBITDA", fig2_public, slide_number=11, width=9, height=3, left=0.11, top=3.70)
    update_figure_slide(ppt, "Labour Force & Unemployment", labour_fig_us, slide_number=5, width=5, height=2.50, left=0.08, top=1.3)
    update_figure_slide(ppt, "External Driver Indicators", external_fig, slide_number=7, width=4.50, height=3.75, left=5.20, top=1.3)
    update_figure_slide(ppt, "GDP by Industry", gdp_fig_us, slide_number=5, width=5.00, height=2.50, left=0.08, top=4.4)
    update_figure_slide(ppt, "CPI and PPI Comparison", cpi_ppi_fig_us, slide_number=5, width=4.55, height=2.50, left=5.10, top=1.3)
    update_figure_slide(ppt, f"Labour force Statitics {state_name}", labour_fig, slide_number=4, width=5, height=2.50, left=0.08, top=1.3)
    update_figure_slide(ppt, f"GDP - {state_name} ", gdp_fig, slide_number=4, width=5.00, height=2.50, left=0.08, top=4.4)

    # Save the PPT file to BytesIO and return the bytes
    ppt_bytes = BytesIO()
    ppt.save(ppt_bytes)
    ppt_bytes.seek(0)
    return ppt_bytes

# Streamlit page configuration
st.set_page_config(page_title="Pitch Book", layout="wide")

# Define URLs and Paths
country = "USA"
timefrom = 2010
url_lfs = f"https://rplumber.ilo.org/data/indicator/?id=EAP_DWAP_SEX_AGE_RT_M&ref_area={country}&sex=SEX_T&classif1=AGE_AGGREGATE_TOTAL&timefrom={timefrom}&type=label&format=.csv"
url_unemp = f"https://rplumber.ilo.org/data/indicator/?id=UNE_DEAP_SEX_AGE_RT_M&ref_area={country}&sex=SEX_T&classif1=AGE_AGGREGATE_TOTAL&timefrom={timefrom}&type=label&format=.csv"
url_pop = "https://fred.stlouisfed.org/graph/fredgraph.csv?bgcolor=%23e1e9f0&chart_type=line&drp=0&fo=open%20sans&graph_bgcolor=%23ffffff&height=450&mode=fred&recession_bars=on&txtcolor=%23444444&ts=12&tts=12&width=1140&nt=0&thu=0&trc=0&show_legend=yes&show_axis_titles=yes&show_tooltip=yes&id=POPTHM&scale=left&cosd=2009-12-29&coed=2024-08-01&line_color=%234572a7&link_values=false&line_style=solid&mark_type=none&mw=3&lw=2&ost=-99999&oet=99999&mma=0&fml=a&fq=Monthly&fam=avg&fgst=lin&fgsnd=2020-02-01&line_index=1&transformation=lin&vintage_date=2024-10-09&revision_date=2024-10-09&nd=1959-01-01"
url_gdp_us = "https://apps.bea.gov/industry/Release/XLS/GDPxInd/GrossOutput.xlsx"
xls = pd.ExcelFile(url_gdp_us)

    # Labour Force Participation Rate Data
df_lfs = pd.read_csv(url_lfs)
df_lfs = df_lfs.rename(columns={'ref_area.label': 'country', 'obs_value': 'labour_force_rate'})
df_lfs['time'] = df_lfs['time'].astype(str)
time_split = df_lfs['time'].str.split('M', expand=True)
df_lfs['year'] = pd.to_numeric(time_split[0], errors='coerce').astype('Int64')
df_lfs['month'] = pd.to_numeric(time_split[1], errors='coerce').astype('Int64')

# Unemployment Rate Data
df_unemp = pd.read_csv(url_unemp)
df_unemp = df_unemp.rename(columns={'ref_area.label': 'country', 'obs_value': 'unemployment_rate'})
df_unemp['time'] = df_unemp['time'].astype(str)
time_split_unemp = df_unemp['time'].str.split('M', expand=True)
df_unemp['year'] = pd.to_numeric(time_split_unemp[0], errors='coerce').astype('Int64')
df_unemp['month'] = pd.to_numeric(time_split_unemp[1], errors='coerce').astype('Int64')

# Population Data
df_pop = pd.read_csv(url_pop)
df_pop = df_pop.rename(columns={'DATE': 'date', 'POPTHM': 'population'})
df_pop['date'] = pd.to_datetime(df_pop['date'])
df_pop['year'] = df_pop['date'].dt.year
df_pop['month'] = df_pop['date'].dt.month

# External Driver Data
external_driver_path = r"streamlit_dashboard/data/business_enviornmental_profiles.csv"
external_driver_df = pd.read_csv(external_driver_path)
external_driver_df['Year'] = pd.to_numeric(external_driver_df['Year'], errors='coerce')
indicator_mapping = {indicator: {'label': indicator, 'value': indicator} for indicator in external_driver_df['Indicator'].unique()}
external_driver_df['Indicator_Options'] = external_driver_df['Indicator'].map(indicator_mapping)

# CPI Industry Mapping
industry_mapping = {
    'All items': 'CUSR0000SA0',
    'Energy': 'CUSR0000SA0E',
    'All items less food': 'CUSR0000SA0L1',
    'All items less food and shelter': 'CUSR0000SA0L12',
    'All items less food, shelter, and energy': 'CUSR0000SA0L12E',
    'All items less food, shelter, energy, and used cars and trucks': 'CUSR0000SA0L12E4',
    'All items less food and energy': 'CUSR0000SA0L1E',
    'All items less shelter': 'CUSR0000SA0L2',
    'All items less medical care': 'CUSR0000SA0L5',
    'All items less energy': 'CUSR0000SA0LE',
    'Apparel less footwear': 'CUSR0000SA311',
    'Apparel': 'CUSR0000SAA',
    'Mens and boys apparel': 'CUSR0000SAA1',
    'Womens and girls apparel': 'CUSR0000SAA2',
    'Commodities': 'CUSR0000SAC',
    'Energy commodities': 'CUSR0000SACE',
    'Commodities less food': 'CUSR0000SACL1',
    'Commodities less food and beverages': 'CUSR0000SACL11',
    'Commodities less food and energy commodities': 'CUSR0000SACL1E',
    'Commodities less food, energy, and used cars and trucks': 'CUSR0000SACL1E4',
    'Durables': 'CUSR0000SAD',
    'Education and communication': 'CUSR0000SAE',
    'Education': 'CUSR0000SAE1',
    'Communication': 'CUSR0000SAE2',
    'Information and information processing': 'CUSR0000SAE21',
    'Education and communication commodities': 'CUSR0000SAEC',
    'Education and communication services': 'CUSR0000SAES',
    'Food and beverages': 'CUSR0000SAF',
    'Food': 'CUSR0000SAF1',
    'Food at home': 'CUSR0000SAF11',
    'Cereals and bakery products': 'CUSR0000SAF111',
    'Meats, poultry, fish, and eggs': 'CUSR0000SAF112',
    'Meats, poultry, and fish': 'CUSR0000SAF1121',
    'Meats': 'CUSR0000SAF11211',
    'Fruits and vegetables': 'CUSR0000SAF113',
    'Fresh fruits and vegetables': 'CUSR0000SAF1131',
    'Nonalcoholic beverages and beverage materials': 'CUSR0000SAF114',
    'Other food at home': 'CUSR0000SAF115',
    'Alcoholic beverages': 'CUSR0000SAF116',
    'Other goods and services': 'CUSR0000SAG',
    'Personal care': 'CUSR0000SAG1',
    'Other goods': 'CUSR0000SAGC',
    'Housing': 'CUSR0000SAH',
    'Shelter': 'CUSR0000SAH1',
    'Fuels and utilities': 'CUSR0000SAH2',
    'Household energy': 'CUSR0000SAH21',
    'Household furnishings and operations': 'CUSR0000SAH3',
    'Household furnishings and supplies': 'CUSR0000SAH31',
    'Medical care': 'CUSR0000SAM',
    'Medical care commodities': 'CUSR0000SAM1',
    'Medical care services': 'CUSR0000SAM2',
    'Nondurables': 'CUSR0000SAN',
    'Domestically produced farm food': 'CUSR0000SAN1D',
    'Nondurables less food': 'CUSR0000SANL1',
    'Nondurables less food and beverages': 'CUSR0000SANL11',
    'Nondurables less food, beverages, and apparel': 'CUSR0000SANL113',
    'Nondurables less food and apparel': 'CUSR0000SANL13',
    'Recreation': 'CUSR0000SAR',
    'Recreation commodities': 'CUSR0000SARC',
    'Recreation services': 'CUSR0000SARS',
    'Services': 'CUSR0000SAS',
    'Utilities and public transportation': 'CUSR0000SAS24',
    'Rent of shelter': 'CUSR0000SAS2RS',
    'Other services': 'CUSR0000SAS367',
    'Transportation services': 'CUSR0000SAS4',
    'Services less rent of shelter': 'CUSR0000SASL2RS',
    'Services less medical care services': 'CUSR0000SASL5',
    'Services less energy services': 'CUSR0000SASLE',
    'Transportation': 'CUSR0000SAT',
    'Private transportation': 'CUSR0000SAT1',
    'Transportation commodities less motor fuel': 'CUSR0000SATCLTB',
    'Mens apparel': 'CUSR0000SEAA',
    'Mens suits, sport coats, and outerwear': 'CUSR0000SEAA01',
    'Mens underwear, nightwear, swimwear and accessories': 'CUSR0000SEAA02',
    'Mens shirts and sweaters': 'CUSR0000SEAA03',
    'Mens pants and shorts': 'CUSR0000SEAA04',
    'Boys apparel': 'CUSR0000SEAB',
    'Womens apparel': 'CUSR0000SEAC',
    'Womens outerwear': 'CUSR0000SEAC01',
    'Womens dresses': 'CUSR0000SEAC02',
    'Womens suits and separates': 'CUSR0000SEAC03',
    'Womens underwear, nightwear, swimwear and accessories': 'CUSR0000SEAC04',
    'Girls apparel': 'CUSR0000SEAD',
    'Footwear': 'CUSR0000SEAE',
    'Mens footwear': 'CUSR0000SEAE01',
    'Boys and girls footwear': 'CUSR0000SEAE02',
    'Womens footwear': 'CUSR0000SEAE03',
    'Infants and toddlers apparel': 'CUSR0000SEAF',
    'Jewelry and watches': 'CUSR0000SEAG',
    'Watches': 'CUSR0000SEAG01',
    'Jewelry': 'CUSR0000SEAG02',
    'Educational books and supplies': 'CUSR0000SEEA',
    'Tuition, other school fees, and childcare': 'CUSR0000SEEB',
    'College tuition and fees': 'CUSR0000SEEB01',
    'Elementary and high school tuition and fees': 'CUSR0000SEEB02',
    'Day care and preschool': 'CUSR0000SEEB03',
    'Technical and business school tuition and fees': 'CUSR0000SEEB04',
    'Postage and delivery services': 'CUSR0000SEEC',
    'Postage': 'CUSR0000SEEC01',
    'Delivery services': 'CUSR0000SEEC02',
    'Information technology, hardware and services': 'CUSR0000SEEE',
    'Computers, peripherals, and smart home assistants': 'CUSR0000SEEE01',
    'Internet services and electronic information providers': 'CUSR0000SEEE03',
    'Telephone hardware, calculators, and other consumer information items': 'CUSR0000SEEE04',
    'Information technology commodities': 'CUSR0000SEEEC',
    'Cereals and cereal products': 'CUSR0000SEFA',
    'Flour and prepared flour mixes': 'CUSR0000SEFA01',
    'Breakfast cereal': 'CUSR0000SEFA02',
    'Rice, pasta, cornmeal': 'CUSR0000SEFA03',
    'Bakery products': 'CUSR0000SEFB',
    'Bread': 'CUSR0000SEFB01',
    'Fresh biscuits, rolls, muffins': 'CUSR0000SEFB02',
    'Cakes, cupcakes, and cookies': 'CUSR0000SEFB03',
    'Other bakery products': 'CUSR0000SEFB04',
    'Beef and veal': 'CUSR0000SEFC',
    'Uncooked ground beef': 'CUSR0000SEFC01',
    'Uncooked beef roasts': 'CUSR0000SEFC02',
    'Uncooked beef steaks': 'CUSR0000SEFC03',
    'Pork': 'CUSR0000SEFD',
    'Bacon, breakfast sausage, and related products': 'CUSR0000SEFD01',
    'Ham': 'CUSR0000SEFD02',
    'Pork chops': 'CUSR0000SEFD03',
    'Other pork including roasts, steaks, and ribs': 'CUSR0000SEFD04',
    'Other meats': 'CUSR0000SEFE',
    'Poultry': 'CUSR0000SEFF',
    'Chicken': 'CUSR0000SEFF01',
    'Other uncooked poultry including turkey': 'CUSR0000SEFF02',
    'Fish and seafood': 'CUSR0000SEFG',
    'Fresh fish and seafood': 'CUSR0000SEFG01',
    'Processed fish and seafood': 'CUSR0000SEFG02',
    'Eggs': 'CUSR0000SEFH',
    'Dairy and related products': 'CUSR0000SEFJ',
    'Milk': 'CUSR0000SEFJ01',
    'Cheese and related products': 'CUSR0000SEFJ02',
    'Ice cream and related products': 'CUSR0000SEFJ03',
    'Other dairy and related products': 'CUSR0000SEFJ04',
    'Fresh fruits': 'CUSR0000SEFK',
    'Apples': 'CUSR0000SEFK01',
    'Bananas': 'CUSR0000SEFK02',
    'Citrus fruits': 'CUSR0000SEFK03',
    'Other fresh fruits': 'CUSR0000SEFK04',
    'Fresh vegetables': 'CUSR0000SEFL',
    'Potatoes': 'CUSR0000SEFL01',
    'Lettuce': 'CUSR0000SEFL02',
    'Tomatoes': 'CUSR0000SEFL03',
    'Other fresh vegetables': 'CUSR0000SEFL04',
    'Processed fruits and vegetables': 'CUSR0000SEFM',
    'Canned fruits and vegetables': 'CUSR0000SEFM01',
    'Frozen fruits and vegetables': 'CUSR0000SEFM02',
    'Other processed fruits and vegetables including dried': 'CUSR0000SEFM03',
    'Juices and nonalcoholic drinks': 'CUSR0000SEFN',
    'Carbonated drinks': 'CUSR0000SEFN01',
    'Nonfrozen noncarbonated juices and drinks': 'CUSR0000SEFN03',
    'Beverage materials including coffee and tea': 'CUSR0000SEFP',
    'Coffee': 'CUSR0000SEFP01',
    'Other beverage materials including tea': 'CUSR0000SEFP02',
    'Sugar and sweets': 'CUSR0000SEFR',
    'Sugar and sugar substitutes': 'CUSR0000SEFR01',
    'Candy and chewing gum': 'CUSR0000SEFR02',
    'Other sweets': 'CUSR0000SEFR03',
    'Fats and oils': 'CUSR0000SEFS',
    'Butter and margarine': 'CUSR0000SEFS01',
    'Salad dressing': 'CUSR0000SEFS02',
    'Other fats and oils including peanut butter': 'CUSR0000SEFS03',
    'Other foods': 'CUSR0000SEFT',
    'Soups': 'CUSR0000SEFT01',
    'Frozen and freeze dried prepared foods': 'CUSR0000SEFT02',
    'Snacks': 'CUSR0000SEFT03',
    'Spices, seasonings, condiments, sauces': 'CUSR0000SEFT04',
    'Other miscellaneous foods': 'CUSR0000SEFT06',
    'Food away from home': 'CUSR0000SEFV',
    'Full service meals and snacks': 'CUSR0000SEFV01',
    'Food at employee sites and schools': 'CUSR0000SEFV03',
    'Other food away from home': 'CUSR0000SEFV05',
    'Alcoholic beverages at home': 'CUSR0000SEFW',
    'Beer, ale, and other malt beverages at home': 'CUSR0000SEFW01',
    'Distilled spirits at home': 'CUSR0000SEFW02',
    'Wine at home': 'CUSR0000SEFW03',
    'Alcoholic beverages away from home': 'CUSR0000SEFX',
    'Tobacco and smoking products': 'CUSR0000SEGA',
    'Cigarettes': 'CUSR0000SEGA01',
    'Miscellaneous personal services': 'CUSR0000SEGD',
    'Legal services': 'CUSR0000SEGD01',
    'Funeral expenses': 'CUSR0000SEGD02',
    'Laundry and dry cleaning services': 'CUSR0000SEGD03',
    'Financial services': 'CUSR0000SEGD05',
    'Miscellaneous personal goods': 'CUSR0000SEGE',
    'Rent of primary residence': 'CUSR0000SEHA',
    'Lodging away from home': 'CUSR0000SEHB',
    'Housing at school, excluding board': 'CUSR0000SEHB01',
    'Other lodging away from home including hotels and motels': 'CUSR0000SEHB02',
    'Owners equivalent rent of residences': 'CUSR0000SEHC',
    'Owners  equivalent rent of primary residence': 'CUSR0000SEHC01',
    'Fuel oil and other fuels': 'CUSR0000SEHE',
    'Fuel oil': 'CUSR0000SEHE01',
    'Propane, kerosene, and firewood': 'CUSR0000SEHE02',
    'Energy services': 'CUSR0000SEHF',
    'Electricity': 'CUSR0000SEHF01',
    'Utility (piped) gas service': 'CUSR0000SEHF02',
    'Water and sewer and trash collection services': 'CUSR0000SEHG',
    'Water and sewerage maintenance': 'CUSR0000SEHG01',
    'Garbage and trash collection': 'CUSR0000SEHG02',
    'Window and floor coverings and other linens': 'CUSR0000SEHH',
    'Window coverings': 'CUSR0000SEHH02',
    'Other linens': 'CUSR0000SEHH03',
    'Furniture and bedding': 'CUSR0000SEHJ',
    'Other furniture': 'CUSR0000SEHJ03',
    'Appliances': 'CUSR0000SEHK',
    'Major appliances': 'CUSR0000SEHK01',
    'Other appliances': 'CUSR0000SEHK02',
    'Other household equipment and furnishings': 'CUSR0000SEHL',
    'Indoor plants and flowers': 'CUSR0000SEHL02',
    'Nonelectric cookware and tableware': 'CUSR0000SEHL04',
    'Tools, hardware, outdoor equipment and supplies': 'CUSR0000SEHM',
    'Tools, hardware and supplies': 'CUSR0000SEHM01',
    'Outdoor equipment and supplies': 'CUSR0000SEHM02',
    'Housekeeping supplies': 'CUSR0000SEHN',
    'Household cleaning products': 'CUSR0000SEHN01',
    'Miscellaneous household products': 'CUSR0000SEHN03',
    'Moving, storage, freight expense': 'CUSR0000SEHP03',
    'Professional services': 'CUSR0000SEMC',
    'Physicians services': 'CUSR0000SEMC01',
    'Dental services': 'CUSR0000SEMC02',
    'Eyeglasses and eye care': 'CUSR0000SEMC03',
    'Services by other medical professionals': 'CUSR0000SEMC04',
    'Hospital and related services': 'CUSR0000SEMD',
    'Hospital services': 'CUSR0000SEMD01',
    'Nursing homes and adult day services': 'CUSR0000SEMD02',
    'Medicinal drugs': 'CUSR0000SEMF',
    'Prescription drugs': 'CUSR0000SEMF01',
    'Nonprescription drugs': 'CUSR0000SEMF02',
    'Video and audio': 'CUSR0000SERA',
    'Televisions': 'CUSR0000SERA01',
    'Cable, satellite, and live streaming television service': 'CUSR0000SERA02',
    'Other video equipment': 'CUSR0000SERA03',
    'Audio equipment': 'CUSR0000SERA05',
    'Video and audio products': 'CUSR0000SERAC',
    'Video and audio services': 'CUSR0000SERAS',
    'Pets, pet products and services': 'CUSR0000SERB',
    'Pets and pet products': 'CUSR0000SERB01',
    'Pet services including veterinary': 'CUSR0000SERB02',
    'Sporting goods': 'CUSR0000SERC',
    'Sports vehicles including bicycles': 'CUSR0000SERC01',
    'Sports equipment': 'CUSR0000SERC02',
    'Photography': 'CUSR0000SERD',
    'Photographic equipment and supplies': 'CUSR0000SERD01',
    'Other recreational goods': 'CUSR0000SERE',
    'Toys': 'CUSR0000SERE01',
    'Sewing machines, fabric and supplies': 'CUSR0000SERE02',
    'Music instruments and accessories': 'CUSR0000SERE03',
    'Other recreation services': 'CUSR0000SERF',
    'Club membership for shopping clubs, fraternal, or other organizations, or participant sports fees': 'CUSR0000SERF01',
    'Admissions': 'CUSR0000SERF02',
    'Fees for lessons or instructions': 'CUSR0000SERF03',
    'Recreational reading materials': 'CUSR0000SERG',
    'New and used motor vehicles': 'CUSR0000SETA',
    'New vehicles': 'CUSR0000SETA01',
    'Used cars and trucks': 'CUSR0000SETA02',
    'Leased cars and trucks': 'CUSR0000SETA03',
    'Car and truck rental': 'CUSR0000SETA04',
    'Motor fuel': 'CUSR0000SETB',
    'Gasoline (all types)': 'CUSR0000SETB01',
    'Other motor fuels': 'CUSR0000SETB02',
    'Motor vehicle parts and equipment': 'CUSR0000SETC',
    'Tires': 'CUSR0000SETC01',
    'Motor vehicle maintenance and repair': 'CUSR0000SETD',
    'Motor vehicle repair': 'CUSR0000SETD03',
    'Motor vehicle insurance': 'CUSR0000SETE',
    'Parking and other fees': 'CUSR0000SETF03',
    'Public transportation': 'CUSR0000SETG',
    'Airline fares': 'CUSR0000SETG01',
    'Other intercity transportation': 'CUSR0000SETG02',
    'Cookies': 'CUSR0000SS02042',
    'Crackers, bread, and cracker products': 'CUSR0000SS0206A',
    'Frozen and refrigerated bakery products, pies, tarts, turnovers': 'CUSR0000SS0206B',
    'Bacon and related products': 'CUSR0000SS04011',
    'Breakfast sausage and related products': 'CUSR0000SS04012',
    'Ham, excluding canned': 'CUSR0000SS04031',
    'Frankfurters': 'CUSR0000SS05011',
    'Lunchmeats': 'CUSR0000SS0501A',
    'Fresh whole chicken': 'CUSR0000SS06011',
    'Shelf stable fish and seafood': 'CUSR0000SS07011',
    'Frozen fish and seafood': 'CUSR0000SS07021',
    'Fresh whole milk': 'CUSR0000SS09011',
    'Fresh milk other than whole': 'CUSR0000SS09021',
    'Butter': 'CUSR0000SS10011',
    'Oranges, including tangerines': 'CUSR0000SS11031',
    'Canned fruits': 'CUSR0000SS13031',
    'Frozen vegetables': 'CUSR0000SS14011',
    'Canned vegetables': 'CUSR0000SS14021',
    'Margarine': 'CUSR0000SS16011',
    'Roasted coffee': 'CUSR0000SS17031',
    'Salt and other seasonings and spices': 'CUSR0000SS18041',
    'Sauces and gravies': 'CUSR0000SS18043',
    'Other condiments': 'CUSR0000SS1804B',
    'Prepared salads': 'CUSR0000SS18064',
    'Whiskey at home': 'CUSR0000SS20021',
    'Distilled spirits, excluding whiskey, at home': 'CUSR0000SS20022',
    'Distilled spirits away from home': 'CUSR0000SS20053',
    'Laundry equipment': 'CUSR0000SS30021',
    'Stationery, stationery supplies, gift wrap': 'CUSR0000SS33032',
    'New cars': 'CUSR0000SS45011',
    'New cars and trucks': 'CUSR0000SS4501A',
    'New trucks': 'CUSR0000SS45021',
    'New motorcycles': 'CUSR0000SS45031',
    'Gasoline, unleaded regular': 'CUSR0000SS47014',
    'Gasoline, unleaded midgrade': 'CUSR0000SS47015',
    'Gasoline, unleaded premium': 'CUSR0000SS47016',
    'Parking fees and tolls': 'CUSR0000SS52051',
    'Intercity train fare': 'CUSR0000SS53022',
    'Ship fare': 'CUSR0000SS53023',
    'Inpatient hospital services': 'CUSR0000SS5702',
    'Outpatient hospital services': 'CUSR0000SS5703',
    'Toys, games, hobbies and playground equipment': 'CUSR0000SS61011',
    'Photographic equipment': 'CUSR0000SS61023',
    'Pet food': 'CUSR0000SS61031',
    'Purchase of pets, pet supplies, accessories': 'CUSR0000SS61032',
    'Admission to movies, theaters, and concerts': 'CUSR0000SS62031',
    'Admission to sporting events': 'CUSR0000SS62032',
    'Pet services': 'CUSR0000SS62053',
    'Veterinarian services': 'CUSR0000SS62054',
    'Tax return preparation and other accounting fees': 'CUSR0000SS68023',
    'Food at elementary and secondary schools': 'CUSR0000SSFV031A'
    }

file_path = r"streamlit_dashboard/data/CPI_industry.txt"
ppi_file_path = r"streamlit_dashboard/data/PPI.txt"
# Load CPI data
df = pd.read_csv(file_path, delimiter=',').dropna().reset_index(drop=True)
df_unpivoted = df.melt(id_vars=["Series ID"], var_name="Month & Year", value_name="Value")
df_unpivoted = df_unpivoted[df_unpivoted["Value"].str.strip() != ""]
df_unpivoted["Series ID"] = df_unpivoted["Series ID"].astype(str)
df_unpivoted["Value"] = pd.to_numeric(df_unpivoted["Value"], errors='coerce')
df_unpivoted["Month & Year"] = pd.to_datetime(df_unpivoted["Month & Year"], format='%b %Y', errors='coerce')
df_cleaned = df_unpivoted.dropna(subset=["Series ID", "Month & Year", "Value"])
all_items_data = df_cleaned[df_cleaned['Series ID'] == 'CUSR0000SA0']
all_items_data = all_items_data[all_items_data['Month & Year'] >= '2010-01-01']
# Function to fetch CPI data for the selected industry

# Load and clean PPI data
df_ppi = pd.read_csv(ppi_file_path, delimiter=',').dropna().reset_index(drop=True)
df_ppi_unpivoted = df_ppi.melt(id_vars=["Year"], var_name="Month", value_name="Value")
df_ppi_unpivoted["Month & Year"] = pd.to_datetime(df_ppi_unpivoted["Month"] + " " + df_ppi_unpivoted["Year"].astype(str),format='%b %Y', errors='coerce')
df_ppi_unpivoted['Value'] = pd.to_numeric(df_ppi_unpivoted['Value'], errors='coerce')
df_ppi_unpivoted = df_ppi_unpivoted.dropna(subset=['Month & Year', 'Value'])
df_ppi_unpivoted = df_ppi_unpivoted[df_ppi_unpivoted["Month & Year"] >= '2010-01-01']

# Clean and reshape GDP data
df_gdp_us = pd.read_excel(xls, sheet_name="TGO105-A")
df_gdp_us = df_gdp_us.iloc[6:].reset_index(drop=True)
df_gdp_us.columns = df_gdp_us.iloc[0]
df_gdp_us = df_gdp_us.drop(0).reset_index(drop=True)
df_gdp_us = df_gdp_us.drop(columns=["Line"])
df_gdp_us = df_gdp_us.drop(df_gdp_us.columns[1], axis=1)
df_gdp_us = df_gdp_us.rename(columns={df_gdp_us.columns[df_gdp_us.isna().any()].tolist()[0]: 'Industry'})
df_gdp_us["Industry"] = df_gdp_us["Industry"].replace("    All industries", "GDP")
df_gdp_us["Industry"] = df_gdp_us["Industry"].str.replace("  ", "")
df_gdp_unpivoted = df_gdp_us.melt(id_vars=["Industry"], var_name="Year", value_name="Value")
df_gdp_unpivoted["Year"] = df_gdp_unpivoted["Year"].astype(int)
df_gdp_unpivoted["Value"] = pd.to_numeric(df_gdp_unpivoted["Value"], errors='coerce')
df_gdp_unpivoted = df_gdp_unpivoted.dropna(subset=["Value"])

# Clean and reshape GDP Percent Change data
df_pct_change = pd.read_excel(xls, sheet_name="TGO101-A")
df_pct_change = df_pct_change.iloc[6:].reset_index(drop=True)
df_pct_change.columns = df_pct_change.iloc[0]
df_pct_change = df_pct_change.drop(0).reset_index(drop=True)
df_pct_change = df_pct_change.drop(columns=["Line"])
df_pct_change = df_pct_change.drop(df_pct_change.columns[1], axis=1)
df_pct_change = df_pct_change.rename(columns={df_pct_change.columns[df_pct_change.isna().any()].tolist()[0]: 'Industry'})
df_pct_change["Industry"] = df_pct_change["Industry"].replace("    All industries", "GDP")
df_pct_change["Industry"] = df_pct_change["Industry"].str.replace("  ", "")
df_pct_unpivoted = df_pct_change.melt(id_vars=["Industry"], var_name="Year", value_name="Percent Change")
df_pct_unpivoted["Year"] = df_pct_unpivoted["Year"].astype(int)
df_pct_unpivoted["Percent Change"] = pd.to_numeric(df_pct_unpivoted["Percent Change"], errors='coerce')
df_pct_unpivoted = df_pct_unpivoted.dropna(subset=["Percent Change"])

df_combined = pd.merge(
    df_gdp_unpivoted,
    df_pct_unpivoted,
    on=["Industry", "Year"],
    how="inner"
   )

    # Filter GDP data
df_gdp_filtered = df_combined[df_combined['Industry'] == 'GDP']

# Create a list of industries excluding GDP for the dropdown
industry_options = df_combined['Industry'].unique().tolist()
industry_options.remove('GDP')

def fetch_cpi_data(series_id, df_cleaned):
    selected_data = df_cleaned[df_cleaned['Series ID'] == series_id]
    selected_data = selected_data[selected_data['Month & Year'] >= '2010-01-01']
    return selected_data[['Month & Year', 'Value']].rename(columns={'Month & Year': 'date', 'Value': 'value'})

def plot_labour_unemployment():
    # Merge unemployment and labour force data
    merged = pd.merge(df_lfs, df_unemp, on=["year", "month", "country"], how='inner')
    merged = pd.merge(merged, df_pop, on=["year", "month"], how='inner')

    fig = go.Figure()

    # Plot population as an area chart on the primary y-axis
    min_population = merged['population'].min()
    fig.add_trace(go.Scatter(
        x=pd.to_datetime(merged[['year', 'month']].assign(day=1)),
        y=merged['population'],
        fill='tozeroy',  # Area chart
        fillcolor='#032649', 
        name='Population',
        mode='none',
        line=dict(color='#032649'),
        yaxis='y1'
    ))

    # Plot unemployment rate on the secondary y-axis
    fig.add_trace(go.Scatter(
        x=pd.to_datetime(merged[['year', 'month']].assign(day=1)),
        y=merged['unemployment_rate'],
        name='Unemployment Rate',
        mode='lines',
        line=dict(color='#EB8928'),
        yaxis='y2'
    ))

    # Plot labour force participation rate on the secondary y-axis
    fig.add_trace(go.Scatter(
        x=pd.to_datetime(merged[['year', 'month']].assign(day=1)),
        y=merged['labour_force_rate'],
        name='Labour Force Participation Rate',
        mode='lines',
        line=dict(color='#595959'),
        yaxis='y2'
    ))

    fig.update_layout(
        title='',
        xaxis=dict(showgrid=False, showticklabels=True),  # No title
        yaxis=dict(
            title='Population',
            side='left',
            range=[merged['population'].min(), merged['population'].max() * 1.1],
            showgrid=False
        ),
        yaxis2=dict(
            title='Rate (%)',
            overlaying='y',  # Overlay on the primary y-axis
            side='right',
            showgrid=False
        ),
        legend=dict(
            orientation="h",x=0.01, y=0.99, bgcolor='rgba(255, 255, 255, 0.6)', font=dict(size=8)
        ),
        hovermode='x unified',  # Unified hover mode
        template='plotly_white',
        margin=dict(l=0, r=0, t=0, b=0)
    )
    st.plotly_chart(fig, use_container_width=True)
    return fig

def plot_external_driver(selected_indicators):

    colors = ['#032649', '#EB8928', '#595959', '#A5A5A5', '#1C798A']

    if not selected_indicators:
        selected_indicators = ["World GDP"]

    fig = go.Figure()

    for i, indicator in enumerate(selected_indicators):
        indicator_data = external_driver_df[external_driver_df['Indicator'] == indicator]

        if '% Change' not in indicator_data.columns:
            raise ValueError(f"Expected '% Change' column not found in {indicator}")

        # Cycle through colors if there are more than 5 indicators
        color = colors[i % len(colors)]  # Use modulus to cycle through the colors

        # Ensure the color is a valid string (in case of any unexpected value)
        if isinstance(color, str) and color.startswith('#') and len(color) == 7:
            fig.add_trace(
                go.Scatter(
                    x=indicator_data['Year'],
                    y=indicator_data['% Change'],
                    mode='lines',
                    name=indicator,
                    line=dict(color=color),  # Apply the color dynamically
                )
            )
        else:
            raise ValueError(f"Invalid color value: {color} for indicator: {indicator}")

    fig.update_layout(
        title='',
        xaxis=dict(
            showgrid=False,    # Remove gridlines
            showticklabels=True,
            showline=False,    # Remove the x-axis line
        ),
        yaxis=dict(
            title='Percent Change',
            showgrid=False,    # Remove gridlines
            showline=False,    # Remove the y-axis line
        ),
        hovermode='x',
        legend=dict(
            x=0, y=1, 
            orientation='h',
            xanchor='left', 
            yanchor='top', 
            traceorder='normal',
            font=dict(size=10),
            bgcolor='rgba(255, 255, 255, 0)', 
            bordercolor='rgba(255, 255, 255, 0)', 
            borderwidth=0 
        ),
        plot_bgcolor='rgba(0,0,0,0)',  # Transparent background for plot area
        paper_bgcolor='rgba(0,0,0,0)', # Transparent background for the whole figure
        margin=dict(l=0, r=0, t=0, b=0)  # Remove padding
    )

    st.plotly_chart(fig, use_container_width=True)
    return fig

def plot_cpi_ppi(selected_series_id):
    """
    Plot CPI and PPI data on a single chart for comparison.
    """
    fig = go.Figure()

    # Fetch and plot the selected CPI industry data
    cpi_data = fetch_cpi_data(selected_series_id, df_cleaned)
    if not cpi_data.empty:
        fig.add_trace(
            go.Scatter(
                x=cpi_data['date'],
                y=cpi_data['value'],
                mode='lines',
                name='CPI by Industry',
                line=dict(color='#032649')
            )
        )
    else:
        st.warning(f"No data available for the selected CPI series: {selected_series_id}")

    # Plot CPI-US All Items data
    if not all_items_data.empty:
        fig.add_trace(
            go.Scatter(
                x=all_items_data['Month & Year'],
                y=all_items_data['Value'],
                mode='lines',
                name='CPI-US',
                line=dict(color='#EB8928', dash='solid')
            )
        )
    else:
        st.warning("No CPI-US All Items data available to display.")

    # Plot aggregated PPI data
    if not df_ppi_unpivoted.empty:
        df_ppi_aggregated = df_ppi_unpivoted.groupby('Month & Year', as_index=False).agg({'Value': 'mean'})
        fig.add_trace(
            go.Scatter(
                x=df_ppi_aggregated['Month & Year'],
                y=df_ppi_aggregated['Value'],
                mode='lines',
                name='PPI-US',
                line=dict(color='#1C798A')
            )
        )
    else:
        st.warning("No PPI data available to display.")

    # Configure the layout of the chart
    fig.update_layout(
        title='',
        xaxis=dict(showgrid=False, showticklabels=True),
        yaxis=dict(title='Value', showgrid=False),
        legend=dict(orientation="h",x=0.01, y=0.99, bgcolor='rgba(255, 255, 255, 0.6)', font=dict(size=8)),hovermode='x unified',plot_bgcolor='rgba(0,0,0,0)',paper_bgcolor='rgba(0,0,0,0)',margin=dict(t=0, b=0, l=0, r=0) )
    st.plotly_chart(fig, use_container_width=True)#, key=key)
    return fig

def plot_gdp_and_industry(selected_industry=None):
    fig = make_subplots(specs=[[{"secondary_y": True}]])

    # 1. Add GDP Value Line (Primary Axis)
    fig.add_trace(
        go.Scatter(
            x=df_gdp_filtered['Year'],
            y=df_gdp_filtered['Value'],
            mode='lines',
            name='GDP - Value',
            fill='tozeroy',  # Create area chart by filling to the x-axis
            fillcolor='#032649', #'rgba(235, 137, 40, 0.6)', 
            line=dict(color='#032649', width=2),
            marker=dict(size=6)
        ),
        secondary_y=False
    )

    # 2. Add GDP Percent Change Line (Secondary Axis)
    fig.add_trace(
        go.Scatter(
            x=df_gdp_filtered['Year'],
            y=df_gdp_filtered['Percent Change'],
            mode='lines',
            name='GDP - Percent Change',
            line=dict(color='#A5A5A5', width=2, dash='solid'),
            marker=dict(size=6)
        ),
        secondary_y=True
    )

    # Check if an industry is selected
    if selected_industry:
        df_industry_filtered = df_combined[df_combined['Industry'] == selected_industry]

        # 3. Add Selected Industry Value Line (Primary Axis)
        fig.add_trace(
            go.Scatter(
                x=df_industry_filtered['Year'],
                y=df_industry_filtered['Value'],
                mode='none',
                name=f'GDP Industry - Value',
                fill='tozeroy',  # Area chart
                fillcolor='#EB8928', 
                line=dict(color='#EB8928', width=2),
                marker=dict(size=6)
            ),
            secondary_y=False
        )

        # 4. Add Selected Industry Percent Change Line (Secondary Axis)
        fig.add_trace(
            go.Scatter(
                x=df_industry_filtered['Year'],
                y=df_industry_filtered['Percent Change'],
                mode='lines',
                name=f'GDP Industry - % Change',
                line=dict(color='#1C798A', width=2, dash='solid'),
                marker=dict(size=6)
            ),
            secondary_y=True
        )

    # Update layout
    fig.update_layout(
        title='',
        xaxis_title='',
        yaxis_title='Value',
        yaxis2_title='Percent Change',
        legend=dict(orientation="h",x=0.01, y=0.99, bgcolor='rgba(255, 255, 255, 0.6)',font=dict(size=8)),template='plotly_white',plot_bgcolor='rgba(0,0,0,0)',paper_bgcolor='rgba(0,0,0,0)',margin=dict(t=0, b=0, l=0, r=0))

    st.plotly_chart(fig, use_container_width=True)
    return fig

states_data_id = {
    "Alabama": {"ur_id": "ALUR", "labour_id": "LBSSA01"},
    "Alaska": {"ur_id": "AKUR", "labour_id": "LBSSA02"},
    "Arizona": {"ur_id": "AZUR", "labour_id": "LBSSA04"},
    "Arkansas": {"ur_id": "ARUR", "labour_id": "LBSSA05"},
    "California": {"ur_id": "CAUR", "labour_id": "LBSSA06"},
    "Colorado": {"ur_id": "COUR", "labour_id": "LBSSA08"},
    "Connecticut": {"ur_id": "CTUR", "labour_id": "LBSSA09"},
    "Delaware": {"ur_id": "DEUR", "labour_id": "LBSSA10"},
    "District of Columbia": {"ur_id": "DCUR", "labour_id": "LBSSA11"},
    "Florida": {"ur_id": "FLUR", "labour_id": "LBSSA12"},
    "Georgia": {"ur_id": "GAUR", "labour_id": "LBSSA13"},
    "Hawaii": {"ur_id": "HIUR", "labour_id": "LBSSA15"},
    "Idaho": {"ur_id": "IDUR", "labour_id": "LBSSA16"},
    "Illinois": {"ur_id": "ILUR", "labour_id": "LBSSA17"},
    "Indiana": {"ur_id": "INUR", "labour_id": "LBSSA18"},
    "Iowa": {"ur_id": "IAUR", "labour_id": "LBSSA19"},
    "Kansas": {"ur_id": "KSUR", "labour_id": "LBSSA20"},
    "Kentucky": {"ur_id": "KYURN", "labour_id": "LBSSA21"},
    "Louisiana": {"ur_id": "LAUR", "labour_id": "LBSSA22"},
    "Maine": {"ur_id": "MEUR", "labour_id": "LBSSA23"},
    "Maryland": {"ur_id": "MDUR", "labour_id": "LBSSA24"},
    "Massachusetts": {"ur_id": "MAUR", "labour_id": "LBSSA25"},
    "Michigan": {"ur_id": "MIUR", "labour_id": "LBSSA26"},
    "Minnesota": {"ur_id": "MNUR", "labour_id": "LBSSA27"},
    "Mississippi": {"ur_id": "MSUR", "labour_id": "LBSSA28"},
    "Missouri": {"ur_id": "MTUR", "labour_id": "LBSSA29"},
    "Montana": {"ur_id": "MTUR", "labour_id": "LBSSA30"},
    "Nebraska": {"ur_id": "NEUR", "labour_id": "LBSSA31"},
    "Nevada": {"ur_id": "NVUR", "labour_id": "LBSSA32"},
    "New Hampshire": {"ur_id": "NHUR", "labour_id": "LBSSA33"},
    "New Jersey": {"ur_id": "NJURN", "labour_id": "LBSSA34"},
    "New Mexico": {"ur_id": "NMUR", "labour_id": "LBSSA35"},
    "New York": {"ur_id": "NYUR", "labour_id": "LBSSA36"},
    "North Carolina": {"ur_id": "NCUR", "labour_id": "LBSSA37"},
    "North Dakota": {"ur_id": "NDUR", "labour_id": "LBSSA38"},
    "Ohio": {"ur_id": "OHUR", "labour_id": "LBSSA39"},
    "Oklahoma": {"ur_id": "OKUR", "labour_id": "LBSSA40"},
    "Oregon": {"ur_id": "ORUR", "labour_id": "LBSSA41"},
    "Pennsylvania": {"ur_id": "PAUR", "labour_id": "LBSSA42"},
    "Puerto Rico": {"ur_id": "PRUR", "labour_id": "LBSSA43"},
    "Rhode Island": {"ur_id": "RIUR", "labour_id": "LBSSA44"},
    "South Carolina": {"ur_id": "SCUR", "labour_id": "LBSSA45"},
    "South Dakota": {"ur_id": "SDUR", "labour_id": "LBSSA46"},
    "Tennessee": {"ur_id": "TNUR", "labour_id": "LBSSA47"},
    "Texas": {"ur_id": "TXUR", "labour_id": "LBSSA48"},
    "Utah": {"ur_id": "UTUR", "labour_id": "LBSSA49"},
    "Vermont": {"ur_id": "VTUR", "labour_id": "LBSSA50"},
    "Washington": {"ur_id": "WAUR", "labour_id": "LBSSA54"},
    "West Virginia": {"ur_id": "WVUR", "labour_id": "LBSSA53"},
    "Wisconsin": {"ur_id": "WIUR", "labour_id": "LBSSA55"},
    "Wyoming": {"ur_id": "WYUR", "labour_id": "LBSSA56"}
}
line_colors = {
    "unemployment": "#032649",  # Dark blue
    "labour_force": "#EB8928",  # Orange
    "gdp": "#032649",  # Dark blue
}

def download_csv(state_name, data_type):
    data_ids = states_data_id.get(state_name)
    if not data_ids:
        return None

    data_id = data_ids["ur_id"] if data_type == "unemployment" else data_ids["labour_id"]
    url = f"https://fred.stlouisfed.org/graph/fredgraph.csv?id={data_id}&cosd=1976-01-01&coed={today}"

    response = requests.get(url)
    if response.status_code == 200:
        csv_data = pd.read_csv(io.StringIO(response.content.decode("utf-8")))
        column_name = "Unemployment" if data_type == "unemployment" else "Labour Force"
        csv_data.rename(columns={csv_data.columns[1]: column_name}, inplace=True)
        csv_data['DATE'] = pd.to_datetime(csv_data['DATE'])
        return csv_data
    else:
        st.error(f"Error downloading {data_type} data for {state_name}.")
        return None

def load_state_gdp_data():
    """Download and preprocess state-level GDP data."""
    global state_gdp_data
    url = "https://apps.bea.gov/regional/zip/SAGDP.zip"
    
    try:
        response = requests.get(url)
        if response.status_code == 200:
            with zipfile.ZipFile(io.BytesIO(response.content)) as z:
                csv_file_name = next(
                    (name for name in z.namelist() 
                     if name.startswith("SAGDP1__ALL_AREAS_") and name.endswith(".csv")), 
                    None
                )
                if csv_file_name:
                    with z.open(csv_file_name) as f:
                        df = pd.read_csv( f, 
                            usecols=lambda col: col not in [ "GeoFIPS", "Region", "TableName", "LineCode", "IndustryClassification", "Unit" ],dtype={"Description": str})
                    df = df[df["Description"] == "Current-dollar GDP (millions of current dollars) "]
                    df = df.melt(id_vars=["GeoName"], var_name="Year", value_name="Value")
                    df.rename(columns={"GeoName": "State"}, inplace=True)
                    df = df[df["Year"].str.isdigit()]
                    df["Year"] = df["Year"].astype(int)
                    df["Value"] = pd.to_numeric(df["Value"], errors='coerce')
                    df.dropna(subset=["Value"], inplace=True)
                    state_gdp_data = df[df["State"] != "United States"]
                else:
                    st.error("No matching CSV file found in the downloaded ZIP.")
        else:
            st.error(f"Failed to download GDP data. Status code: {response.status_code}")

    except Exception as e:
        st.error(f"An error occurred: {e}")

load_state_gdp_data()

def plot_unemployment_labour_chart(state_name):
    unemployment_data = download_csv(state_name, "unemployment")
    labour_data = download_csv(state_name, "labour")

    if unemployment_data is not None and labour_data is not None:
        unemployment_data = unemployment_data[unemployment_data['DATE'].dt.year >= 2000]
        labour_data = labour_data[labour_data['DATE'].dt.year >= 2000]

        merged_data = pd.merge(unemployment_data, labour_data, on='DATE')


        fig = go.Figure()
        fig.add_trace(go.Scatter(x=merged_data['DATE'], y=merged_data['Unemployment'], mode='lines',line=dict(color=line_colors["unemployment"]), name="Unemployment"))
        fig.add_trace(go.Scatter(x=merged_data['DATE'], y=merged_data['Labour Force'], mode='lines',line=dict(color=line_colors["labour_force"]), name="Labour Force"))

        last_row = merged_data.iloc[-1]
        fig.add_annotation(
            x=last_row['DATE'], y=last_row['Unemployment'],
            text=f"{last_row['Unemployment']:.1f}"+"%", showarrow=True, arrowhead=1, ax=-40, ay=-40
        )
        fig.add_annotation(
            x=last_row['DATE'], y=last_row['Labour Force'],
            text=f" {last_row['Labour Force']:.1f}"+"%", showarrow=True, arrowhead=1, ax=-40, ay=40
        )

        fig.update_layout(
            title=f"Labour Force & Unemployment Rate - {state_name}",
            xaxis_title=" ",
            yaxis_title="Rate",
            template="plotly_white",
            legend=dict( x=0, y=1, xanchor='left', yanchor='top',title_text=None ),plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)')

        st.plotly_chart(fig, use_container_width=True)
        return fig
    else:
        st.warning(f"No data available for {state_name}.")
        return None

def plot_gdp_chart(state_name):
    global state_gdp_data

    if state_gdp_data is not None:
        gdp_data = state_gdp_data[state_gdp_data["State"].str.lower() == state_name.lower()]
        gdp_data = gdp_data[gdp_data["Year"] >= 2000]

        if not gdp_data.empty:
            fig = go.Figure()
            fig.add_trace(go.Scatter(x=gdp_data['Year'], y=gdp_data['Value'], mode='lines',line=dict(color=line_colors["gdp"]), name=f"{state_name} GDP"))

            last_row = gdp_data.iloc[-1]
            fig.add_annotation(
                x=last_row['Year'], y=last_row['Value'],
                text=f" {last_row['Value']:.0f}", showarrow=True, arrowhead=1, ax=-40, ay=-40
            )

            fig.update_layout(
                title=(f"GDP - {state_name}"),
                xaxis_title=" ",
                yaxis_title="GDP (Millions of Dollars)",
                template="plotly_white",
                plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)')

            st.plotly_chart(fig, use_container_width=True)
            return fig
        else:
            st.warning(f"No GDP data available for {state_name}.")
            return None
    else:
        st.warning("State GDP data not loaded.")
        return None
# Define S3 file paths
precedent_path = "s3://documentsapi/industry_data/precedent.parquet"
public_comp_path = "s3://documentsapi/industry_data/public_comp_data.parquet"
s3_path_rma = "s3://documentsapi/industry_data/rma_data.parquet"
s3_path_public_comp = "s3://documentsapi/industry_data/Public Listed Companies US.xlsx"

df_rma = dd.read_parquet(s3_path_rma, storage_options=storage_options)  # Set to True if the bucket is public
df_rma = df_rma.rename(columns={
    'ReportID': 'Report_ID',      
    'Line Items': 'LineItems',    
    'Value': 'Value',             
    'Percent': 'Percent' 
})
usecols = [
    "Name", "Industry", "Revenue (in %)", "COGS (in %)", "Gross Profit (in %)", "EBITDA (in %)",
    "Operating Profit (in %)", "Other Expenses (in %)", "Operating Expenses (in %)", "Net Income (in %)",
    "Cash (in %)", "Accounts Receivables (in %)", "Inventories (in %)", "Other Current Assets (in %)",
    "Total Current Assets (in %)", "Fixed Assets (in %)", "PPE (in %)", "Total Assets (in %)",
    "Accounts Payable (in %)", "Short Term Debt (in %)", "Long Term Debt (in %)", "Other Current Liabilities (in %)",
    "Total Current Liabilities (in %)", "Other Liabilities (in %)", "Total Liabilities (in %)",
    "Net Worth (in %)", "Total Liabilities & Equity (in %)"
]
# Load the public company data
df_public_comp = pd.read_excel(s3_path_public_comp, sheet_name="FY 2023", storage_options=storage_options,usecols=usecols, engine='openpyxl')
df_public_comp = df_public_comp.rename(columns=lambda x: x.replace(" (in %)", ""))

# Load data for both Public Comps and Precedent Transactions
try:
    # Load Precedent Transactions Data
    precedent_df = dd.read_parquet(
        precedent_path,
        storage_options=storage_options,
        usecols=['Year', 'Target', 'EV/Revenue', 'EV/EBITDA', 'Business Description', 'Industry', 'Location'],
        dtype={'EV/Revenue': 'float64', 'EV/EBITDA': 'float64'}
    )

    # Load Public Comps Data
    public_comp_df = dd.read_parquet(
        public_comp_path,
        storage_options=storage_options,
        usecols=['Name', 'Country', 'Enterprise Value (in $)', 'Revenue (in $)', 'EBITDA (in $)', 'Business Description', 'Industry']
    ).rename(columns={
        'Name': 'Company',
        'Country': 'Location',
        'Enterprise Value (in $)': 'Enterprise Value',
        'Revenue (in $)': 'Revenue',
        'EBITDA (in $)': 'EBITDA'
    })

    # Ensure numeric conversion
    public_comp_df['Enterprise Value'] = dd.to_numeric(public_comp_df['Enterprise Value'], errors='coerce')
    public_comp_df['Revenue'] = dd.to_numeric(public_comp_df['Revenue'], errors='coerce')
    public_comp_df['EBITDA'] = dd.to_numeric(public_comp_df['EBITDA'], errors='coerce')

    # Drop rows with invalid data for division
    public_comp_df = public_comp_df.dropna(subset=['Enterprise Value', 'Revenue', 'EBITDA'])

    # Calculate EV/Revenue and EV/EBITDA
    public_comp_df['EV/Revenue'] = public_comp_df['Enterprise Value'] / public_comp_df['Revenue']
    public_comp_df['EV/EBITDA'] = public_comp_df['Enterprise Value'] / public_comp_df['EBITDA']

    # Get unique industries and locations from Public Comps
    public_industries = public_comp_df['Industry'].dropna().compute().unique().tolist()
    public_locations = public_comp_df['Location'].dropna().compute().unique().tolist()

    # Compute the DataFrame for use in Streamlit
    precedent_df = precedent_df.compute()
    public_comp_df = public_comp_df.compute()

except Exception as e:
    st.error(f"Error loading data from S3: {e}")
    st.stop()

# Accordion for Precedent Transactions
with st.expander("Precedent Transactions"):
    industries = precedent_df['Industry'].dropna().unique()
    locations = precedent_df['Location'].dropna().unique()
    col1, col2 = st.columns(2)
    selected_industries = col1.multiselect("Select Industry", industries, key="precedent_industries")
    selected_locations = col2.multiselect("Select Location", locations, key="precedent_locations")
    if selected_industries and selected_locations:
        filtered_precedent_df = precedent_df[
            precedent_df['Industry'].isin(selected_industries) & precedent_df['Location'].isin(selected_locations)
        ]
        filtered_precedent_df = filtered_precedent_df[['Target', 'Year', 'EV/Revenue', 'EV/EBITDA','Business Description']]
        filtered_precedent_df['Year'] = filtered_precedent_df['Year'].astype(int)

        st.subheader("Precedent Transactions")
        gb = GridOptionsBuilder.from_dataframe(filtered_precedent_df)
        gb.configure_selection(selection_mode="multiple", use_checkbox=True)
        gb.configure_column(
            field="Target",
            tooltipField="Business Description",
            maxWidth=400
            )
        gb.configure_columns(["Business Description"], hide=False)    
        grid_options = gb.build()

        # Display Ag-Grid table
        grid_response = AgGrid(
            filtered_precedent_df,
            gridOptions=grid_options,
            update_mode=GridUpdateMode.SELECTION_CHANGED,
            height=400,
            width='100%',
            theme='streamlit'
        )
        selected_data = pd.DataFrame(grid_response['selected_rows'])
        if not selected_data.empty:

            avg_data = selected_data.groupby('Year')[['EV/Revenue', 'EV/EBITDA']].mean().reset_index()
            avg_data['Year'] = avg_data['Year'].astype(int)

            # Define colors
            color_ev_revenue = "#032649"  # Default Plotly blue
            color_ev_ebitda = "#032649"   # Default Plotly red

            median_ev_revenue = avg_data['EV/Revenue'].median()
            median_ev_ebitda = avg_data['EV/EBITDA'].median()

            # Create the EV/Revenue chart with data labels
            fig1_precedent = px.bar(avg_data, x='Year', y='EV/Revenue', title="", text='EV/Revenue')  # No title
            fig1_precedent.update_traces(marker_color=color_ev_revenue, texttemplate='%{text:.1f}'+'x', textposition='auto',textfont=dict(size=10))
            fig1_precedent.update_layout(yaxis_title="EV/Revenue", xaxis_title=" ", bargap=0.4, bargroupgap=0.4, yaxis=dict(showgrid=False), shapes=[dict(type='line', x0=avg_data['Year'].min(), x1=avg_data['Year'].max(), y0=median_ev_revenue, y1=median_ev_revenue, line=dict(color='#EB8928', dash='dot', width=2))], annotations=[dict(x=avg_data['Year'].max(), y=median_ev_revenue, xanchor='left', yanchor='bottom', text=f'Median: {median_ev_revenue:.1f}'+'x', showarrow=False, font=dict(size=12, color='gray'), bgcolor='white')],plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',margin=dict(l=0, r=0, t=0, b=0))

            st.plotly_chart(fig1_precedent)

            # Create the EV/EBITDA chart with data labels
            fig2_precedent = px.bar(avg_data, x='Year', y='EV/EBITDA', title="", text='EV/EBITDA')
            fig2_precedent.update_traces(marker_color=color_ev_ebitda, texttemplate='%{text:.1f}'+ 'x', textposition='auto',textfont=dict(size=10))
            fig2_precedent.update_layout(yaxis_title="EV/EBITDA", xaxis_title=" ", bargap=0.4, bargroupgap=0.4, yaxis=dict(showgrid=False), shapes=[dict(type='line', x0=avg_data['Year'].min(), x1=avg_data['Year'].max(), y0=median_ev_ebitda, y1=median_ev_ebitda, line=dict(color='#EB8928', dash='dot', width=2))], annotations=[dict(x=avg_data['Year'].max(), y=median_ev_ebitda, xanchor='left', yanchor='bottom', text=f'Median: {median_ev_ebitda:.1f}'+'x', showarrow=False, font=dict(size=12, color='gray'), bgcolor='white')],plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',margin=dict(l=0, r=0, t=0, b=0))
            
            st.plotly_chart(fig2_precedent)

# Accordion for Public Comps
with st.expander("Public Comps"):
    col1, col2 = st.columns(2)
    selected_industries = col1.multiselect("Select Industry", public_industries, key="public_industries")
    selected_locations = col2.multiselect("Select Location", public_locations, key="public_locations")
    if selected_industries and selected_locations:
        filtered_df = public_comp_df[public_comp_df['Industry'].isin(selected_industries) & public_comp_df['Location'].isin(selected_locations)]
        filtered_df = filtered_df[['Company',  'EV/Revenue', 'EV/EBITDA', 'Business Description']]
        filtered_df['EV/Revenue'] = filtered_df['EV/Revenue'].round(1)
        filtered_df['EV/EBITDA'] = filtered_df['EV/EBITDA'].round(1)

        # Set up Ag-Grid for selection
        st.subheader("Public Listed Companies")
        gb = GridOptionsBuilder.from_dataframe(filtered_df)
        gb.configure_selection(selection_mode="multiple", use_checkbox=True)
        gb.configure_column(
            field="Company",
            tooltipField="Business Description",
            maxWidth=400
        )
        gb.configure_columns(["Business Description"], hide=False)    
        grid_options = gb.build()

        # Display Ag-Grid table
        grid_response = AgGrid(
            filtered_df,
            gridOptions=grid_options,
            update_mode=GridUpdateMode.SELECTION_CHANGED,
            height=400,
            width='100%',
            theme='streamlit'
        )

        selected_data = pd.DataFrame(grid_response['selected_rows'])

        if not selected_data.empty:
            avg_data = selected_data.groupby('Company')[['EV/Revenue', 'EV/EBITDA']].mean().reset_index()

            color_ev_revenue = "#032649"  # Default Plotly blue
            color_ev_ebitda = "#032649"   # Default Plotly red

            median_ev_revenue = avg_data['EV/Revenue'].median()
            median_ev_ebitda = avg_data['EV/EBITDA'].median()

            # Create the EV/Revenue chart with data labels
            fig1_public = px.bar(avg_data, x='Company', y='EV/Revenue', title="EV/Revenue", text='EV/Revenue')
            fig1_public.update_traces(marker_color=color_ev_revenue, texttemplate='%{text:.1f}'+'x', textposition='auto',textfont=dict(size=10))
            fig1_public.update_layout(yaxis_title="EV/Revenue", xaxis_title=" ",bargap=0.4,bargroupgap=0.4,yaxis=dict(showgrid=False),xaxis=dict(tickangle=0,automargin="height+width"))
            fig1_public.add_shape(type="line",x0=-0.5, x1=len(avg_data['Company']) - 0.5,  y0=median_ev_revenue, y1=median_ev_revenue,line=dict(color="#EB8928", width=2, dash="dot"),  xref="x", yref="y")
            fig1_public.add_annotation(x=len(avg_data['Company']) - 1, y=median_ev_revenue + 0.2, text=f"Median: {median_ev_revenue:.1f}x",showarrow=False, font=dict(size=10, color="gray"), xanchor="left",bgcolor='white',plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)')

            st.plotly_chart(fig1_public)

            # Create the EV/EBITDA chart with data labels
            fig2_public = px.bar(avg_data, x='Company', y='EV/EBITDA', title="EV/EBITDA", text='EV/EBITDA')
            fig2_public.update_traces(marker_color=color_ev_ebitda,texttemplate='%{text:.1f}'+'x', textposition='auto',textfont=dict(size=10))
            fig2_public.update_layout(yaxis_title="EV/EBITDA", xaxis_title=" ",bargap=0.4,bargroupgap=0.4,yaxis=dict(showgrid=False),xaxis=dict(tickangle=0,automargin="height+width"))
            fig2_public.add_shape(type="line",x0=-0.5, x1=len(avg_data['Company']) - 0.5,  y0=median_ev_ebitda, y1=median_ev_ebitda,line=dict(color="#EB8928", width=2, dash="dot"),  xref="x", yref="y")
            fig2_public.add_annotation(x=len(avg_data['Company']) - 1, y=median_ev_ebitda + 0.2, text=f"Median: {median_ev_ebitda:.1f}x",showarrow=False, font=dict(size=10, color="gray"), xanchor="left",bgcolor='white',plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)')
            
            st.plotly_chart(fig2_public)

with st.expander("US Indicators"):
    st.subheader("US Indicators")

    # Labour Force & Unemployment Data
    st.subheader("Labour Force & Unemployment")
    labour_fig_us = plot_labour_unemployment()

    # External Driver Indicators
    st.subheader("External Driver Indicators")
    selected_indicators = st.multiselect(
        "Select External Indicators",
        options=external_driver_df["Indicator"].unique(),
        default=["World GDP"],
        key="external_indicators_multiselect"
    )
    external_fig = plot_external_driver(selected_indicators)

    # GDP by Industry
    st.subheader("GDP by Industry")
    selected_gdp_industry = st.selectbox(
        "Select Industry",
        options=df_combined["Industry"].unique(),
        index=0,
        key="gdp_industry_selectbox"
    )
    gdp_fig_us = plot_gdp_and_industry(selected_gdp_industry)

    # CPI and PPI Comparison
    st.subheader("CPI & PPI")
    selected_cpi_series = st.selectbox(
        "Select CPI Industry",
        list(industry_mapping.keys()),
        index=1,
        key="cpi_series_selectbox"
    )
    selected_series_id = industry_mapping[selected_cpi_series]
    cpi_ppi_fig_us = plot_cpi_ppi(selected_series_id)

with st.expander("State Indicators"):
    st.subheader("State Indicators - US")
    state_name = st.selectbox("Select State", list(states_data_id.keys()), index=0)
    st.subheader(f"{state_name} - Unemployment & Labour Force")
    labour_fig = plot_unemployment_labour_chart(state_name)
    st.subheader(f"{state_name} - GDP Over Time")
    gdp_fig = plot_gdp_chart(state_name)

with st.expander("Benchmarking"):
    st.subheader("Benchmarking")
    # Create dropdown and process data for RMA and public comps
    industries_rma = df_rma[~df_rma['Industry'].isnull() & df_rma['Industry'].map(lambda x: isinstance(x, str))]['Industry'].compute().unique()
    industries_public = df_public_comp[~df_public_comp['Industry'].isnull() & df_public_comp['Industry'].map(lambda x: isinstance(x, str))]['Industry'].unique()
    industries = sorted(set(industries_rma).union(set(industries_public)))
    
    income_statement_items = ["Revenue", "COGS", "Gross Profit", "EBITDA", "Operating Profit", "Other Expenses", "Operating Expenses","Profit Before Taxes", "Net Income"]
    balance_sheet_items = ["Cash", "Accounts Receivables", "Inventories", "Other Current Assets", "Total Current Assets", "Fixed Assets","Intangibles", "PPE", "Total Assets", "Accounts Payable", "Short Term Debt", "Long Term Debt", "Other Current Liabilities", "Total Current Liabilities", "Other Liabilities", "Total Liabilities", "Net Worth", "Total Liabilities & Equity"]
    
    selected_industry = st.selectbox("Select Industry", industries)
    if selected_industry:

        filtered_df_rma = df_rma[df_rma['Industry'] == selected_industry].compute()

        if 'Report_ID' in filtered_df_rma.columns:
            filtered_df_rma['Report_ID'] = filtered_df_rma['Report_ID'].replace({"Assets": "Balance Sheet", "Liabilities & Equity": "Balance Sheet"})

        income_statement_df_rma = filtered_df_rma[filtered_df_rma['Report_ID'] == 'Income Statement'][['LineItems', 'Percent']].rename(columns={'Percent': 'RMA Percent'})
        balance_sheet_df_rma = filtered_df_rma[filtered_df_rma['Report_ID'] == 'Balance Sheet'][['LineItems', 'Percent']].rename(columns={'Percent': 'RMA Percent'})

        filtered_df_public = df_public_comp[df_public_comp['Industry'] == selected_industry]

        df_unpivoted = pd.melt(
            filtered_df_public,
            id_vars=["Name", "Industry"],
            var_name="LineItems",
            value_name="Value"
        )
        df_unpivoted['LineItems'] = df_unpivoted['LineItems'].str.replace(" (in %)", "", regex=False)
        df_unpivoted['Value'] = pd.to_numeric(df_unpivoted['Value'].replace("-", 0), errors='coerce').fillna(0) * 100
        df_unpivoted = df_unpivoted.groupby('LineItems')['Value'].mean().reset_index()
        df_unpivoted = df_unpivoted.rename(columns={'Value': 'Public Comp Percent'})
        df_unpivoted['Public Comp Percent'] = df_unpivoted['Public Comp Percent'].round(0).astype(int).astype(str) + '%'

        income_statement_df_public = df_unpivoted[df_unpivoted['LineItems'].isin(income_statement_items)]
        balance_sheet_df_public = df_unpivoted[df_unpivoted['LineItems'].isin(balance_sheet_items)]

        income_statement_df = pd.merge(
            pd.DataFrame({'LineItems': income_statement_items}),
            income_statement_df_rma,
            on='LineItems',
            how='left'
        ).merge(
            income_statement_df_public,
            on='LineItems',
            how='left'
        )

        balance_sheet_df = pd.merge(
            pd.DataFrame({'LineItems': balance_sheet_items}),
            balance_sheet_df_rma,
            on='LineItems',
            how='left'
        ).merge(
            balance_sheet_df_public,
            on='LineItems',
            how='left'
        )

        st.write("Income Statement")
        st.dataframe(income_statement_df.fillna(np.nan), hide_index=True, use_container_width=True)

        st.write("Balance Sheet")
        st.dataframe(balance_sheet_df.fillna(np.nan), hide_index=True, use_container_width=True)

if st.button("Export Charts to PowerPoint", key="export_button"):
    try:
        # Export the charts to PowerPoint using the export_all_to_pptx function
        pptx_file = export_all_to_pptx(
            labour_fig_us, external_fig, gdp_fig_us, cpi_ppi_fig_us, 
            fig1_precedent, fig2_precedent, fig1_public, fig2_public, 
            labour_fig, gdp_fig
        )

        # Create a download button for the user to download the PowerPoint file
        st.download_button(
            label="Download PowerPoint",  # The label for the button
            data=pptx_file,  # The PowerPoint file content (must be in byte format)
            file_name=f"Pitch_Book_{date.today().strftime('%Y-%m-%d')}.pptx",  # The default filename for the download
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"  # MIME type for PowerPoint
        )

    except Exception as e:
        st.error(f"Error during PowerPoint export: {e}")
