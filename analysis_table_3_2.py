import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import openpyxl
import re

# Read the demand data
df_demand = pd.read_excel('./CM_Data_Explorer May 2024 (2).xlsx', sheet_name='3.2 Cleantech demand by mineral')

# Helper function to clean sheet names
def clean_sheet_name(name):
    clean_name = re.sub(r'[\\/*?:\[\]]', '', str(name))
    clean_name = clean_name.replace('(', '').replace(')', '')
    clean_name = clean_name.strip()[:31]
    return clean_name

def clean_demand_data(df):
    # Remove any completely empty rows and columns
    df = df.dropna(how='all').dropna(axis=1, how='all')
    df = df.reset_index(drop=True)
    
    print("First few rows before cleaning:")
    print(df.head())
    
    # Find the year row (first row with 2023)
    year_row = None
    for idx, row in df.iterrows():
        if 2023 in row.values:
            year_row = idx
            break
    
    if year_row is None:
        raise ValueError("Could not find year row with 2023")
    
    # Get years
    year_row_data = df.iloc[year_row]
    years = [year for year in year_row_data if isinstance(year, (int, float)) and not pd.isna(year)]
    years = sorted(list(set(years)))  # Get unique years in order
    
    # Get data rows (exclude header rows)
    data_rows = df.iloc[year_row + 1:].copy()
    data_rows = data_rows.dropna(how='all')  # Remove any completely empty rows
    data_rows = data_rows.reset_index(drop=True)
    
    return data_rows, year_row_data, years

def analyze_total_demand(data_rows, year_row_data, years):
    scenarios = {}
    current_scenario = None
    
    print("\nAnalyzing total demand...")
    
    # Find where each year appears for each scenario
    year_columns = {}
    scenario_ranges = {
        'Stated Policies': (1, 6),      # Columns for 2023, 2030-2050
        'Announced Pledges': (7, 11),   # Columns for 2030-2050
        'Net Zero': (12, 16)            # Columns for 2030-2050
    }
    
    # Define metals in order
    metals = [
        "Boron",
        "Cadmium",
        "Chromium",
        "Copper",
        "Cobalt",
        "Gallium",
        "Germanium",
        "Battery-grade graphite",
        "Hafnium",
        "Indium",
        "Iridium",
        "Lead",
        "Lithium",
        "Magnesium",
        "Manganese",
        "Molybdenum",
        "Nickel",
        "Niobium",
        "PGMs (other than iridium)",
        "Selenium",
        "Silicon",
        "Silver",
        "Tantalum",
        "Tellurium",
        "Tin",
        "Titanium",
        "Tungsten",
        "Vanadium",
        "Zinc",
        "Zirconium",
        "Arsenic",
        "Neodymium",
        "Dysprosium",
        "Praseodymium",
        "Terbium",
        "Yttrium",
        "Lanthanum"
    ]
    
    # Create DataFrames for each scenario
    for scenario, (start, end) in scenario_ranges.items():
        scenario_data = []
        
        for idx, row in data_rows.iterrows():
            metal = row.iloc[0]  # First column contains metal names
            
            if pd.notna(metal) and isinstance(metal, str) and not metal.startswith("Total"):
                data = {'Metal': metal}
                
                # Get data for years in this scenario's range
                col_indices = range(start, end + 1)
                scenario_years = years[len(years)-len(col_indices):] if start > 1 else years[:len(col_indices)]
                
                for year, col_idx in zip(scenario_years, col_indices):
                    value = row.iloc[col_idx]
                    if pd.notna(value):
                        data[str(year)] = value
                
                if len(data) > 1:  # More than just the Metal column
                    scenario_data.append(data)
                    print(f"Added data for {metal} in {scenario}")
        
        if scenario_data:
            scenarios[scenario] = pd.DataFrame(scenario_data)
    
    return scenarios

def save_to_excel(scenario_data):
    if not scenario_data:
        raise ValueError("No data to save!")
    
    with pd.ExcelWriter('3.2 Cleantech demand by minerals.xlsx', engine='openpyxl', mode='w') as writer:
        for scenario, df in scenario_data.items():
            if not df.empty:
                # Clean up the data
                numeric_cols = df.select_dtypes(include=[np.number]).columns
                if not numeric_cols.empty:
                    df[numeric_cols] = df[numeric_cols].round(3)
                
                df = df.replace([np.inf, -np.inf], np.nan)
                
                # Save to Excel
                sheet_name = clean_sheet_name(scenario)
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                
                # Auto-adjust column widths
                worksheet = writer.sheets[sheet_name]
                for idx, col in enumerate(df.columns):
                    max_length = max(
                        df[col].astype(str).apply(len).max(),
                        len(str(col))
                    ) + 2
                    worksheet.column_dimensions[openpyxl.utils.get_column_letter(idx + 1)].width = max_length

# Run the analysis
data_rows, year_row_data, years = clean_demand_data(df_demand)
scenario_data = analyze_total_demand(data_rows, year_row_data, years)

try:
    save_to_excel(scenario_data)
    print("\nAnalysis complete! Data saved to 'total_demand_scenarios.xlsx'")
except Exception as e:
    print(f"\nError saving Excel file: {e}")
