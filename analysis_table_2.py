import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import openpyxl
import re

# Read the supply data
df_supply = pd.read_excel('./CM_Data_Explorer May 2024 (2).xlsx', sheet_name='2 Total supply for key minerals')

# Helper function to clean sheet names
def clean_sheet_name(name):
    clean_name = re.sub(r'[\\/*?:\[\]]', '', str(name))
    clean_name = clean_name.replace('(', '').replace(')', '')
    clean_name = clean_name.strip()[:31]
    return clean_name

def clean_mineral_supply_data(df):
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

def analyze_mineral_supply(data_rows, year_row_data, years):
    metals = {}
    current_metal = None
    mining_data = []
    refining_data = []
    
    print("\nAnalyzing supply data...")
    
    # First, let's understand the column structure
    print("\nColumn structure:")
    for col in range(len(year_row_data)):
        if pd.notna(year_row_data.iloc[col]):
            print(f"Column {col}: {year_row_data.iloc[col]}")
    
    # Find where each year appears
    year_columns = {}
    for year in years:
        year_columns[year] = [i for i, val in enumerate(year_row_data) if val == year]
        print(f"Year {year} appears in columns: {year_columns[year]}")
    
    in_refining_section = False
    mining_countries = []
    refining_countries = []
    
    for idx, row in data_rows.iterrows():
        category = row.iloc[0]  # First column contains categories/countries
        
        if pd.notna(category) and isinstance(category, str):
            if " - Mining" in category:
                # Save previous metal's data if it exists
                if current_metal and (mining_data or refining_data):
                    print(f"\nFor {current_metal}:")
                    print("Mining countries:", mining_countries)
                    print("Refining countries:", refining_countries)
                    metals[current_metal] = process_metal_data(mining_data, refining_data, years, year_row_data)
                
                current_metal = category.split(" - ")[0].strip()
                print(f"\nStarting new metal: {current_metal}")
                mining_data = []
                refining_data = []
                mining_countries = []
                refining_countries = []
                in_refining_section = False
                
            elif " - Refining" in category:
                in_refining_section = True
                print(f"\nStarting refining section for {current_metal}")
                
            elif category != current_metal and not any(x in category for x in ["- Mining", "- Refining", "Notes:", "Base case"]):
                # This is a country row
                if not in_refining_section:
                    mining_countries.append(category)
                else:
                    refining_countries.append(category)
                
                data_mining = {'Country': category}
                data_refining = {'Country': category}
                
                # Get both mining and refining data for each year
                for year in years:
                    cols = year_columns[year]
                    if len(cols) >= 2:  # We have both mining and refining data
                        mining_val = row.iloc[cols[0]]  # First occurrence is mining
                        refining_val = row.iloc[cols[1]]  # Second occurrence is refining
                        if pd.notna(mining_val):
                            data_mining[f'Mining_{int(year)}'] = mining_val
                        if pd.notna(refining_val):
                            data_refining[f'Refining_{int(year)}'] = refining_val
                
                # Only add rows that have data
                if len(data_mining) > 1:  # More than just the Country column
                    mining_data.append(data_mining)
                    print(f"Added mining data for {category}")
                if len(data_refining) > 1:  # More than just the Country column
                    refining_data.append(data_refining)
                    print(f"Added refining data for {category}")
    
    # Process the last metal
    if current_metal and (mining_data or refining_data):
        print(f"\nFor {current_metal}:")
        print("Mining countries:", mining_countries)
        print("Refining countries:", refining_countries)
        metals[current_metal] = process_metal_data(mining_data, refining_data, years, year_row_data)
    
    return metals

def process_metal_data(mining_data, refining_data, years, year_row_data):
    print(f"\nProcessing metal data:")
    print(f"Mining data entries: {len(mining_data)}")
    print(f"Refining data entries: {len(refining_data)}")
    
    # Convert to DataFrames
    mining_df = pd.DataFrame(mining_data) if mining_data else pd.DataFrame(columns=['Country'])
    refining_df = pd.DataFrame(refining_data) if refining_data else pd.DataFrame(columns=['Country'])
    
    print("\nMining DataFrame:")
    print(mining_df.head())
    print("\nRefining DataFrame:")
    print(refining_df.head())
    
    # Get countries from mining data to maintain order
    mining_countries = mining_df['Country'].tolist() if not mining_df.empty else []
    
    # Get additional countries from refining data
    refining_countries = refining_df['Country'].tolist() if not refining_df.empty else []
    additional_countries = [c for c in refining_countries if c not in mining_countries]
    
    # Combine countries maintaining mining order and adding new refining countries at the end
    all_countries = mining_countries + additional_countries
    
    # Create final dataframe with all countries
    final_data = []
    for country in all_countries:
        row_data = {'Country': country}
        
        # Add mining data
        if country in mining_countries:
            mining_row = mining_df[mining_df['Country'] == country]
            if not mining_row.empty:
                for year in years:
                    col = f'Mining_{int(year)}'
                    if col in mining_row.columns:
                        row_data[col] = mining_row[col].iloc[0]
        
        # Add refining data
        if country in refining_countries:
            refining_row = refining_df[refining_df['Country'] == country]
            if not refining_row.empty:
                for year in years:
                    col = f'Refining_{int(year)}'
                    if col in refining_row.columns:
                        row_data[col] = refining_row[col].iloc[0]
        
        final_data.append(row_data)
    
    result_df = pd.DataFrame(final_data)
    print("\nFinal DataFrame shape:", result_df.shape)
    print("Final DataFrame columns:", result_df.columns.tolist())
    return result_df

def save_to_excel(metal_data):
    if not metal_data:
        raise ValueError("No data to save!")
        
    # Create separate mining and refining DataFrames for each metal
    mining_tables = {}
    refining_tables = {}
    
    for metal, df in metal_data.items():
        if not df.empty:
            # Split columns into mining and refining
            mining_cols = ['Country'] + [col for col in df.columns if 'Mining' in col]
            refining_cols = ['Country'] + [col for col in df.columns if 'Refining' in col]
            
            # Create separate tables
            mining_table = df[mining_cols].copy()
            refining_table = df[refining_cols].copy()
            
            # Remove rows with all NaN values except Country
            mining_table = mining_table.dropna(subset=[col for col in mining_cols if col != 'Country'], how='all')
            refining_table = refining_table.dropna(subset=[col for col in refining_cols if col != 'Country'], how='all')
            
            mining_tables[metal] = mining_table
            refining_tables[metal] = refining_table
    
    # Save to Excel with separate sheets for mining and refining
    with pd.ExcelWriter('mineral_supply_mining.xlsx', engine='openpyxl', mode='w') as writer:
        for metal, df in mining_tables.items():
            if not df.empty:
                sheet_name = clean_sheet_name(metal)
                
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
    
    with pd.ExcelWriter('mineral_supply_refining.xlsx', engine='openpyxl', mode='w') as writer:
        for metal, df in refining_tables.items():
            if not df.empty:
                sheet_name = clean_sheet_name(metal)
                
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

    print("\nCreated separate files for mining and refining data:")
    print("- mineral_supply_mining.xlsx")
    print("- mineral_supply_refining.xlsx")

# Run the analysis
data_rows, year_row_data, years = clean_mineral_supply_data(df_supply)
metal_data = analyze_mineral_supply(data_rows, year_row_data, years)

try:
    save_to_excel(metal_data)
    print("\nAnalysis complete! Data saved to 'mineral_supply_mining.xlsx' and 'mineral_supply_refining.xlsx'")
except Exception as e:
    print(f"\nError saving Excel files: {e}")
