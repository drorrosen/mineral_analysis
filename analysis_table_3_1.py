import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import openpyxl
import re

# Read the demand data
df_demand = pd.read_excel('./CM_Data_Explorer May 2024 (2).xlsx', sheet_name='3.1 Cleantech demand by tech')

# Helper function to clean sheet names
def clean_sheet_name(name):
    clean_name = re.sub(r'[\\/*?:\[\]]', '', str(name))
    clean_name = clean_name.replace('(', '').replace(')', '')
    clean_name = clean_name.strip()[:31]
    return clean_name

def clean_demand_scenario_data(df):
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

def analyze_demand_scenarios(data_rows, year_row_data, years):
    scenarios = {}
    current_metal = None
    current_scenario = None
    scenario_data = []
    
    print("\nAnalyzing demand scenarios...")
    
    # Find where each year appears for each scenario
    year_columns = {}
    scenario_ranges = {
        'Stated Policies': (1, 6),      # Columns for 2023, 2030-2050
        'Announced Pledges': (7, 11),   # Columns for 2030-2050
        'Net Zero': (12, 16)            # Columns for 2030-2050
    }
    
    # Define exact metal names to match in correct order
    metals = [
        "Chromium",
        "Copper",
        "Cobalt",
        "Battery-grade graphite",
        "Lithium",
        "Manganese",
        "Molybdenum",
        "Nickel",
        "PGMs",
        "Silicon",
        "Silver",
        "Zinc",
        "Neodymium"
    ]
    
    for year in years:
        year_columns[year] = [i for i, val in enumerate(year_row_data) if val == year]
        print(f"Year {year} appears in columns: {year_columns[year]}")
    
    for idx, row in data_rows.iterrows():
        category = row.iloc[0]  # First column contains categories/countries
        
        if pd.notna(category) and isinstance(category, str):
            # Check if this is a new metal section
            is_new_metal = False
            for metal in metals:
                if metal in category and not category.startswith("Total"):
                    current_metal = metal
                    is_new_metal = True
                    break
            
            if is_new_metal:
                # Save previous metal's data if it exists
                if scenario_data:
                    if current_metal not in scenarios:
                        scenarios[current_metal] = {}
                    for scenario, (start, end) in scenario_ranges.items():
                        scenario_df = create_scenario_df(scenario_data, years, start, end)
                        if not scenario_df.empty:
                            scenarios[current_metal][scenario] = scenario_df
                    scenario_data = []
                
                print(f"\nStarting new metal: {current_metal}")
                
            elif category != current_metal and not any(x in category for x in ["Total", "Notes:", "Base case"]):
                # This is a technology/sector row
                data = {'Technology': category}
                
                # Get data for all columns
                for col_idx in range(len(row)):
                    if col_idx > 0:  # Skip the first column (category name)
                        value = row.iloc[col_idx]
                        if pd.notna(value):
                            data[f'col_{col_idx}'] = value
                
                if len(data) > 1:  # More than just the Technology column
                    scenario_data.append(data)
                    print(f"Added data for {category} under {current_metal}")
    
    # Process the last metal
    if current_metal and scenario_data:
        if current_metal not in scenarios:
            scenarios[current_metal] = {}
        for scenario, (start, end) in scenario_ranges.items():
            scenario_df = create_scenario_df(scenario_data, years, start, end)
            if not scenario_df.empty:
                scenarios[current_metal][scenario] = scenario_df
    
    return scenarios

def create_scenario_df(data, years, start_col, end_col):
    """Create a DataFrame for a specific scenario using column ranges."""
    if not data:
        return pd.DataFrame()
    
    # Create base DataFrame
    df = pd.DataFrame(data)
    
    # Create new DataFrame with Technology and year columns
    result_data = []
    for _, row in df.iterrows():
        new_row = {'Technology': row['Technology']}
        
        # Map columns to years based on the scenario's range
        col_indices = range(start_col, end_col + 1)
        scenario_years = years[len(years)-len(col_indices):] if start_col > 1 else years[:len(col_indices)]
        
        for year, col_idx in zip(scenario_years, col_indices):
            col_name = f'col_{col_idx}'
            if col_name in row:
                new_row[str(year)] = row[col_name]
        
        result_data.append(new_row)
    
    result_df = pd.DataFrame(result_data)
    return result_df if len(result_df.columns) > 1 else pd.DataFrame()

def save_to_excel(scenario_data):
    if not scenario_data:
        raise ValueError("No data to save!")
    
    # Define metal order
    metal_order = [
        "Chromium",
        "Copper",
        "Cobalt",
        "Battery-grade graphite",
        "Lithium",
        "Manganese",
        "Molybdenum",
        "Nickel",
        "PGMs",
        "Silicon",
        "Silver",
        "Zinc",
        "Neodymium"
    ]
    
    with pd.ExcelWriter('demand_scenarios.xlsx', engine='openpyxl', mode='w') as writer:
        # Process metals in the correct order
        for metal in metal_order:
            if metal in scenario_data:
                scenarios = scenario_data[metal]
                for scenario, df in scenarios.items():
                    if not df.empty:
                        sheet_name = clean_sheet_name(f"{metal}_{scenario}")
                        
                        # Clean up the data
                        numeric_cols = df.select_dtypes(include=[np.number]).columns
                        if not numeric_cols.empty:
                            df[numeric_cols] = df[numeric_cols].round(3)
                        
                        df = df.replace([np.inf, -np.inf], np.nan)
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                        
                        # Auto-adjust column widths
                        worksheet = writer.sheets[sheet_name]
                        for idx, col in enumerate(df.columns):
                            max_length = max(
                                df[col].astype(str).apply(len).max(),
                                len(str(col))
                            ) + 2
                            worksheet.column_dimensions[openpyxl.utils.get_column_letter(idx + 1)].width = max_length

    print("\nCreated scenarios file with metals in correct order:")
    print("- demand_scenarios.xlsx")

# Run the analysis
data_rows, year_row_data, years = clean_demand_scenario_data(df_demand)
scenario_data = analyze_demand_scenarios(data_rows, year_row_data, years)

try:
    save_to_excel(scenario_data)
    print("\nAnalysis complete! Data saved to 'demand_scenarios.xlsx'")
except Exception as e:
    print(f"\nError saving Excel file: {e}")

