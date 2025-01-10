import os
import pandas as pd
import streamlit as st
from pptx import Presentation
from pptx.util import Inches
from io import BytesIO
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from PIL import Image

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
df_pop = df_pop.rename(columns={'observation_date': 'date', 'POPTHM': 'population'})
df_pop['date'] = pd.to_datetime(df_pop['date'])
df_pop['year'] = df_pop['date'].dt.year
df_pop['month'] = df_pop['date'].dt.month
df_pop['population'] = df_pop['population']/1000

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
df_gdp_unpivoted["Value"] = pd.to_numeric(df_gdp_unpivoted["Value"], errors='coerce')*1000
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
    title='',  # No title
    xaxis=dict(
        showgrid=False,
        showticklabels=True,
        color="#595959",
        tickfont=dict(color="#595959"),
        tickangle=0,  # Rotate x-axis labels to avoid overlap
        automargin=True  # Automatically adjust margins for better spacing
    ),
    yaxis=dict(
        showgrid=False,
        title='Population',
        color="#595959",
        tickfont=dict(color="#595959"),
        side='left',
        range=[merged['population'].min(), merged['population'].max() * 1.1],
        tickformat=',',
    ),
    yaxis2=dict(
        title='Rate (%)',
        overlaying='y',
        side='right'
    ),
    legend=dict(
        orientation="h",
        x=0.01,
        y=-0.15,  # Move legend below the plot to avoid overlap
        bgcolor='rgba(255, 255, 255, 0.6)',
        font=dict(size=10)
    ),
    hovermode='x unified',
    template='plotly_white',
    plot_bgcolor='rgba(0,0,0,0)',  # Transparent plot background
    paper_bgcolor='rgba(0,0,0,0)',  # Transparent paper background
    height=300,  # Increased height for better spacing
    width=500,  # Adjusted width for better visualization
    margin=dict(b=80, t=30,l=10, r=10)  # Add more bottom margin for x-axis labels
)
    st.plotly_chart(fig, use_container_width=True)
    return fig

def plot_external_driver(selected_indicators):
    colors = ['#032649', '#1C798A', '#EB8928', '#595959', '#A5A5A5']

    if not selected_indicators:
        selected_indicators = ["World GDP"]

    fig = go.Figure()

    for i, indicator in enumerate(selected_indicators):
        indicator_data = external_driver_df[external_driver_df['Indicator'] == indicator]

        if '% Change' not in indicator_data.columns:
            raise ValueError(f"Expected '% Change' column not found in {indicator}")

        color = colors[i % len(colors)]

        if isinstance(color, str) and color.startswith('#') and len(color) == 7:
            fig.add_trace(go.Scatter(
                x=indicator_data['Year'],
                y=indicator_data['% Change'],
                mode='lines',
                name=indicator,
                line=dict(color=color),
            ))
        else:
            raise ValueError(f"Invalid color value: {color} for indicator: {indicator}")

    fig.update_layout(
        title=' ',
        xaxis=dict(
            showgrid=False,
            showticklabels=True,
            color="#595959",  # X-axis label and line color
            tickfont=dict(color="#595959"),  # X-axis tick labels color
        ),
        yaxis=dict(
            title='',
            showgrid=False,
            color="#595959",  # Y-axis label and line color
            tickfont=dict(color="#595959"),  # Y-axis tick labels color
        ),
        hovermode='x',
        legend=dict(
            x=0.1,
            y=-0.3,  # Place legend below the x-axis
            orientation='h',  # Horizontal legend
            xanchor='center',
            yanchor='top',
            traceorder='normal',
            font=dict(size=10, color="#595959"),  # Legend text color
            bgcolor='rgba(255, 255, 255, 0)',
        ),
        plot_bgcolor='rgba(0,0,0,0)', 
        paper_bgcolor='rgba(0,0,0,0)',
        height=400,  # Chart height
        width=500,  # Chart width
        margin=dict(b=200, t=50, l=30, r=25),  # Increased bottom margin
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
        xaxis=dict(
            showgrid=False,
            showticklabels=True,
            color="#595959",  # X-axis label and line color
            tickfont=dict(color="#595959"),  # X-axis tick labels color
        ),
        yaxis=dict(
            title='Value',
            showgrid=False,
            color="#595959",  # Y-axis label and line color
            tickfont=dict(color="#595959"),  # Y-axis tick labels color
        ),
        legend=dict(
            orientation="h",
            x=0.01,  # Align the legend horizontally
            y=-0.2,  # Place it below the chart
            xanchor='left',
            yanchor='top',
            bgcolor='rgba(255, 255, 255, 0.6)',
            font=dict(size=10, color="#595959"),  # Legend text color
        ),
        hovermode='x unified',
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        height=300,
        width=500,
        margin=dict(b=60, t=20),  # Increased bottom margin for space
    )

    st.plotly_chart(fig, use_container_width=True)
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
            marker=dict(size=10)
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
            marker=dict(size=10)
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
                marker=dict(size=10)
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
                marker=dict(size=10)
            ),
            secondary_y=True
        )

    # Update layout
    fig.update_layout(
        title='',
        xaxis_title='',
        yaxis_title='Value',
        yaxis2_title='Percent Change',
        xaxis=dict(
            showgrid=False,
            showticklabels=True,
            color="#595959",  # X-axis label and line color
            tickfont=dict(color="#595959"),  # X-axis tick labels color
        ),
        yaxis=dict(
            title='',
            showgrid=False,
            color="#595959",  # Y-axis label and line color
            tickfont=dict(color="#595959"),  # Y-axis tick labels color
            tickformat=',',
        ),
        legend=dict(
            orientation="h",
            x=0.01,  # Center the legend horizontally
            y=-0.15,  # Place it below the chart
            xanchor='left',  # Center alignment
            yanchor='top',  # Align to top of the legend box
            bgcolor='rgba(255, 255, 255, 0.6)',
            font=dict(size=10),
            traceorder='normal'
        ),
        template='plotly_white',
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        height=450,
        width=700,
        margin=dict(b=120, t=80,l=10, r=10),  # Increased bottom margin for space
    )
    st.plotly_chart(fig, use_container_width=True)
    return fig

# Function to export charts to PowerPoint
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

def export_all_to_pptx(labour_fig, external_fig, gdp_fig, cpi_ppi_fig):
    # Load the custom template
    template_path = os.path.join(os.getcwd(), "streamlit_dashboard", "data", "main_template_pitch.pptx")
    ppt = Presentation(template_path)  # Load the template

    # Use the existing slides (slide_number corresponds to the slide index)
    update_figure_slide(ppt, "Labour Force & Unemployment Data", labour_fig, slide_number=5, width=5, height=2.50, left=0.08, top=1.3)
    update_figure_slide(ppt, "External Driver Indicators", external_fig, slide_number=7, width=4.50, height=3.75, left=5.20, top=1.3)
    update_figure_slide(ppt, "GDP by Industry", gdp_fig, slide_number=5, width=5.00, height=2.50, left=0.08, top=4.4)
    update_figure_slide(ppt, "CPI and PPI Comparison", cpi_ppi_fig, slide_number=5, width=4.55, height=2.50, left=5.10, top=1.3)

    # Save the PPT file to BytesIO and return the bytes
    ppt_bytes = BytesIO()
    ppt.save(ppt_bytes)
    ppt_bytes.seek(0)
    return ppt_bytes

def get_us_indicators_layout():
    """Render the full dashboard layout and export data directly without session state."""
    st.title("US Indicators Dashboard")

    # Labour Force & Unemployment Data
    st.subheader("Labour Force & Unemployment")
    labour_fig = plot_labour_unemployment()

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
    gdp_fig = plot_gdp_and_industry(selected_gdp_industry)

    # CPI and PPI Comparison
    st.subheader("CPI & PPI")
    selected_cpi_series = st.selectbox(
        "Select CPI Industry", 
        list(industry_mapping.keys()), 
        index=1, 
        key="cpi_series_selectbox"
    )
    selected_series_id = industry_mapping[selected_cpi_series]
    cpi_ppi_fig = plot_cpi_ppi(selected_series_id)

    if st.button("Export Charts to PowerPoint", key="export_button"):
        # Export the charts to PowerPoint using the export_all_to_pptx function
        pptx_file = export_all_to_pptx(labour_fig, external_fig, gdp_fig, cpi_ppi_fig)
        
        # Create a download button for the user to download the PowerPoint file
        st.download_button(
            label="Download PowerPoint",  # The label for the button
            data=pptx_file,  # The PowerPoint file content
            file_name="US_indicators.pptx",  # The default filename for the download
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"  # MIME type for PowerPoint
        )

get_us_indicators_layout()
