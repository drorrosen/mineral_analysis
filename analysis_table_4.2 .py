import pandas as pd
import numpy as np
import openpyxl
import re

def clean_sheet_name(name):
    """Clean sheet name for Excel compatibility"""
    clean_name = re.sub(r'[\\/*?:\[\]]', '', str(name))
    clean_name = clean_name.replace('(', '').replace(')', '')
    clean_name = clean_name.strip()[:31]
    return clean_name

def clean_wind_data(df):
    """Clean and structure the wind data"""
    # Remove empty rows and columns
    df = df.dropna(how='all').dropna(axis=1, how='all')
    df = df.reset_index(drop=True)
    
    print("\nFirst few rows of raw data:")
    print(df.head())
    
    # Define scenarios and sections
    scenarios = [
        'Stated Policies scenario',
        'Announced Pledges scenario',
        'Net Zero Emissions by 2050 scenario'
    ]
    
    sections = ['Base case', 'Constrained rare earth elements supply']
    
    # Materials to capture (in order as they appear in the data)
    materials = [
        'Boron', 'Chromium', 'Copper', 'Manganese', 'Molybdenum',
        'Nickel', 'Zinc', 'Neodymium', 'Dysprosium', 'Praseodymium',
        'Terbium', 'Total wind'
    ]
    
    # Process each section
    all_data = []
    current_section = None
    
    for idx, row in df.iterrows():
        first_col = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ''
        
        # Check if this is a section header
        if first_col in sections:
            current_section = first_col
            continue
        
        # Skip rows until we find first section
        if current_section is None:
            continue
            
        material = first_col
        if material in materials:
            row_data = {
                'Section': current_section,
                'Material': material
            }
            
            # Get 2023 value (base year) - same for all scenarios
            row_data['2023'] = row.iloc[1]
            
            # Process Stated Policies scenario (columns 2-6)
            for i, year in enumerate([2030, 2035, 2040, 2045, 2050]):
                row_data[f'Stated Policies_{year}'] = row.iloc[2+i]
            
            # Process Announced Pledges scenario (columns 7-11)
            for i, year in enumerate([2030, 2035, 2040, 2045, 2050]):
                row_data[f'Announced Pledges_{year}'] = row.iloc[7+i]
            
            # Process Net Zero scenario (columns 12-16)
            for i, year in enumerate([2030, 2035, 2040, 2045, 2050]):
                row_data[f'Net Zero_{year}'] = row.iloc[12+i]
            
            all_data.append(row_data)
    
    # Create DataFrame
    result_df = pd.DataFrame(all_data)
    
    print("\nProcessed data:")
    print(result_df.head())
    print(f"Total rows: {len(result_df)}")
    
    return result_df, scenarios

def save_to_excel(df, scenarios):
    """Save organized data to Excel"""
    with pd.ExcelWriter('4_2_wind_scenarios.xlsx', engine='openpyxl', mode='w') as writer:
        # Save overview sheet
        pd.DataFrame(['Wind Technology Analysis by Scenario']).to_excel(writer, sheet_name='Overview', index=False)
        
        # Save main data
        df.to_excel(writer, sheet_name='Wind_Data', index=False)
        
        # Create section-specific sheets
        for section in df['Section'].unique():
            section_df = df[df['Section'] == section].copy()
            sheet_name = clean_sheet_name(section)
            section_df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            # Auto-adjust column widths
            worksheet = writer.sheets[sheet_name]
            for idx, col in enumerate(section_df.columns):
                max_length = max(
                    section_df[col].astype(str).apply(len).max(),
                    len(str(col))
                ) + 2
                worksheet.column_dimensions[openpyxl.utils.get_column_letter(idx + 1)].width = max_length
        
        # Create scenario-specific views
        for scenario in scenarios:
            # Get columns for this scenario
            scenario_cols = ['Section', 'Material', '2023'] + [col for col in df.columns if scenario in col]
            scenario_df = df[scenario_cols].copy()
            
            # Clean up the dataframe
            numeric_cols = scenario_df.select_dtypes(include=[np.number]).columns
            if not numeric_cols.empty:
                scenario_df[numeric_cols] = scenario_df[numeric_cols].round(3)
            
            # Create sheet name
            sheet_name = clean_sheet_name(scenario)
            scenario_df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            # Auto-adjust column widths
            worksheet = writer.sheets[sheet_name]
            for idx, col in enumerate(scenario_df.columns):
                max_length = max(
                    scenario_df[col].astype(str).apply(len).max(),
                    len(str(col))
                ) + 2
                worksheet.column_dimensions[openpyxl.utils.get_column_letter(idx + 1)].width = max_length

def main():
    # Read the data
    print("\nReading wind data...")
    df = pd.read_excel('CM_Data_Explorer May 2024 (2).xlsx', sheet_name='4.2 Wind')
    
    # Clean and analyze data
    print("\nCleaning and analyzing data...")
    result_df, scenarios = clean_wind_data(df)
    
    # Save results
    try:
        save_to_excel(result_df, scenarios)
        print("\nAnalysis complete! Data saved to '4_2_wind_scenarios.xlsx'")
    except Exception as e:
        print(f"\nError saving Excel file: {e}")

if __name__ == "__main__":
    main()
