import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import openpyxl
import re

# Read the data
df2 = pd.read_excel('./CM_Data_Explorer May 2024 (2).xlsx', sheet_name='1 Total demand for key minerals')

# Helper function to clean sheet names - moved to top
def clean_sheet_name(name):
    # Remove invalid characters and limit length
    clean_name = re.sub(r'[\\/*?:\[\]]', '', str(name))  # Remove invalid Excel characters
    clean_name = clean_name.replace('(', '').replace(')', '')  # Remove parentheses
    clean_name = clean_name.strip()[:31]  # Excel sheet names limited to 31 chars
    return clean_name

def clean_mineral_demand_data(df):
    # Remove any completely empty rows and columns
    df = df.dropna(how='all').dropna(axis=1, how='all')
    df = df.reset_index(drop=True)
    
    print("First few rows before cleaning:")
    print(df.head())
    
    # Find the scenario row and year row
    scenario_row = df[df.iloc[:, 2].str.contains('scenario', na=False)].index[0]
    year_row = df[df.iloc[:, 1] == 2023].index[0]
    
    # Get the row with scenarios and years
    scenario_row_data = df.iloc[scenario_row]
    year_row_data = df.iloc[year_row]
    
    # Initialize scenario data
    scenarios = []
    current_scenario = None
    scenario_columns = {}
    
    # First, get the 2023 data which is shared across scenarios
    base_year_col = df.columns[df.iloc[year_row] == 2023].tolist()[0]
    
    # Get data rows (exclude header rows)
    data_rows = df.iloc[year_row + 1:].copy()
    data_rows = data_rows.dropna(how='all')  # Remove any completely empty rows
    data_rows = data_rows.reset_index(drop=True)
    
    # Get base year data aligned with cleaned data rows
    base_year_data = data_rows.iloc[:, df.columns.get_loc(base_year_col)]
    
    # Create new dataframe with the correct structure
    df_clean = pd.DataFrame()
    df_clean['Category'] = data_rows.iloc[:, 0].reset_index(drop=True)
    
    # Map scenarios to their columns
    for col in range(1, len(df.columns)):
        cell_value = scenario_row_data.iloc[col]
        if pd.notna(cell_value) and 'scenario' in str(cell_value).lower():
            current_scenario = cell_value
            scenarios.append(current_scenario)
            scenario_columns[current_scenario] = []
        if current_scenario and pd.notna(year_row_data.iloc[col]):
            scenario_columns[current_scenario].append(col)
    
    # Add data for each scenario
    for scenario in scenarios:
        # Add 2023 data first
        df_clean[f"{scenario}_2023"] = base_year_data
        
        # Add other years
        cols = scenario_columns[scenario]
        years = [year_row_data.iloc[col] for col in cols]
        
        # Get data for this scenario
        scenario_data = data_rows.iloc[:, cols].copy()
        scenario_data.columns = [f"{scenario}_{int(year)}" for year in years]
        
        # Add to clean dataframe
        df_clean = pd.concat([df_clean, scenario_data], axis=1)
    
    print("\nFirst few rows after cleaning:")
    print(df_clean.head())
    print("\nColumns in cleaned data:")
    print(df_clean.columns.tolist())
    print("\nShape of cleaned data:", df_clean.shape)
    
    return df_clean, scenarios

def analyze_mineral_data(df, scenarios):
    minerals = {}
    current_mineral = None
    current_data = []
    
    for idx, row in df.iterrows():
        category = row['Category']
        
        # Check if this is a mineral header row (no numeric values and not empty)
        if pd.notna(category) and not any(pd.to_numeric(row.iloc[1:], errors='coerce').notna()):
            # Save previous mineral's data if it exists
            if current_mineral and current_data:
                minerals[current_mineral] = pd.DataFrame(current_data)
                current_data = []
            current_mineral = category
        # If this is a data row and we have a current mineral
        elif current_mineral and any(pd.to_numeric(row.iloc[1:], errors='coerce').notna()):
            current_data.append(row)
    
    # Add the last mineral's data
    if current_mineral and current_data:
        minerals[current_mineral] = pd.DataFrame(current_data)
    
    return minerals

# Save organized data to Excel
def save_to_excel(mineral_data, scenarios):
    with pd.ExcelWriter('organized_mineral_demand.xlsx', engine='openpyxl', mode='w') as writer:
        # First save a default sheet to ensure we have at least one visible sheet
        pd.DataFrame(['Mineral Analysis Results']).to_excel(writer, sheet_name='Overview', index=False)
        
        # Create summary for each scenario
        for scenario in scenarios:
            summary_data = []
            for mineral, df in mineral_data.items():
                try:
                    if f'{scenario}_2023' in df.columns and f'{scenario}_2050' in df.columns:
                        total_2023 = pd.to_numeric(df[f'{scenario}_2023'], errors='coerce').sum()
                        total_2050 = pd.to_numeric(df[f'{scenario}_2050'], errors='coerce').sum()
                        
                        if pd.notna(total_2023) and total_2023 != 0:
                            growth = ((total_2050 - total_2023) / total_2023 * 100)
                            cagr = (((total_2050/total_2023)**(1/27)) - 1) * 100
                        else:
                            growth = np.nan
                            cagr = np.nan
                        
                        summary_data.append({
                            'Mineral': mineral,
                            'Total_2023': total_2023,
                            'Total_2050': total_2050,
                            'Growth_Rate_%': growth,
                            'CAGR_%': cagr
                        })
                except Exception as e:
                    print(f"Could not calculate summary for {mineral} in {scenario}: {e}")
            
            if summary_data:  # Only create summary if we have data
                summary_df = pd.DataFrame(summary_data)
                if not summary_df.empty:
                    # Handle potential missing columns
                    for col in ['Growth_Rate_%', 'CAGR_%']:
                        if col not in summary_df.columns:
                            summary_df[col] = np.nan
                    
                    summary_df = summary_df.dropna(subset=['Growth_Rate_%'])
                    if not summary_df.empty:
                        summary_df = summary_df.sort_values('Growth_Rate_%', ascending=False)
                        
                        # Format numbers
                        for col in ['Total_2023', 'Total_2050']:
                            summary_df[col] = summary_df[col].round(2)
                        for col in ['Growth_Rate_%', 'CAGR_%']:
                            summary_df[col] = summary_df[col].round(1)
                        
                        # Save summary sheet
                        sheet_name = clean_sheet_name(f'Summary_{scenario}')
                        summary_df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        # Save individual mineral sheets
        for mineral, df in mineral_data.items():
            if not df.empty:
                sheet_name = clean_sheet_name(mineral)
                numeric_cols = df.select_dtypes(include=[np.number]).columns
                if not numeric_cols.empty:
                    df[numeric_cols] = df[numeric_cols].round(3)
                df.to_excel(writer, sheet_name=sheet_name, index=False)

# Run the analysis
df2_clean, scenarios = clean_mineral_demand_data(df2)
mineral_data = analyze_mineral_data(df2_clean, scenarios)

try:
    save_to_excel(mineral_data, scenarios)
    print("\nAnalysis complete! Data saved to 'organized_mineral_demand.xlsx'")
except Exception as e:
    print(f"\nError saving Excel file: {e}")

