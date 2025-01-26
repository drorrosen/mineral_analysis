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

def clean_ev_data(df):
    """Clean and structure the EV data"""
    # Remove empty rows and columns
    df = df.dropna(how='all').dropna(axis=1, how='all')
    df = df.reset_index(drop=True)
    
    print("First few rows before cleaning:")
    print(df.head())
    
    # Define scenarios
    scenarios = [
        'Stated Policies scenario',
        'Announced Pledges scenario',
        'Net Zero Emissions by 2050 scenario'
    ]
    
    # Known technology sections
    tech_sections = {
        'Base case': None,
        'High material prices': None,
        'Wider use of silicon-rich anodes': None,
        'Faster uptake of solid state batteries': None,
        'Lower battery sizes': None,
        'Limited battery size reduction': None
    }
    
    # Find start indices for each section
    for idx, row in df.iterrows():
        if pd.notna(row.iloc[0]):
            for tech in tech_sections.keys():
                if tech in str(row.iloc[0]):
                    tech_sections[tech] = idx
    
    # Convert to list of tuples (start, end)
    section_indices = []
    tech_names = list(tech_sections.keys())
    for i in range(len(tech_names)):
        start = tech_sections[tech_names[i]]
        end = tech_sections[tech_names[i + 1]] if i < len(tech_names) - 1 else len(df)
        if start is not None:
            section_indices.append((tech_names[i], start, end))
    
    print("\nFound sections:", section_indices)
    
    # Materials to capture
    materials = [
        'Copper', 'Cobalt', 'Battery-grade graphite', 'Lithium',
        'Manganese', 'Nickel', 'Silicon', 'Neodymium',
        'Dysprosium', 'Praseodymium', 'Terbium', 'Total EV'
    ]
    
    # Process each section
    technologies = {}
    
    for tech_name, start_idx, end_idx in section_indices:
        print(f"\nProcessing section: {tech_name}")
        section_df = df.iloc[start_idx:end_idx].copy()
        
        # Get data rows
        data_rows = []
        current_material = None
        
        for idx, row in section_df.iterrows():
            material = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ''
            
            if material in materials:
                current_material = material
                row_data = {'Material': material}
                
                # Base year (2023)
                row_data['2023'] = row.iloc[1]
                
                # Stated Policies scenario (columns 2-6)
                for i, year in enumerate([2030, 2035, 2040, 2045, 2050]):
                    row_data[f'Stated Policies_{year}'] = row.iloc[2+i]
                
                # Announced Pledges scenario (columns 7-11)
                for i, year in enumerate([2030, 2035, 2040, 2045, 2050]):
                    row_data[f'Announced Pledges_{year}'] = row.iloc[7+i]
                
                # Net Zero scenario (columns 12-16)
                for i, year in enumerate([2030, 2035, 2040, 2045, 2050]):
                    row_data[f'Net Zero_{year}'] = row.iloc[12+i]
                
                data_rows.append(row_data)
        
        if data_rows:
            technologies[tech_name] = pd.DataFrame(data_rows)
            print(f"Processed {len(data_rows)} materials for {tech_name}")
    
    print("\nFinal results:")
    for tech, df in technologies.items():
        print(f"\n{tech}:")
        print(df.head())
        print(f"Total materials: {len(df)}")
    
    return technologies, scenarios

def save_to_excel(tech_data, scenarios):
    """Save organized data to Excel with separate sheets for each technology-scenario combination"""
    with pd.ExcelWriter('4_3_ev_scenarios.xlsx', engine='openpyxl', mode='w') as writer:
        # Save overview sheet
        pd.DataFrame(['EV Technology Analysis by Scenario']).to_excel(writer, sheet_name='Overview', index=False)
        
        # Save data for each technology
        for tech, df in tech_data.items():
            if not df.empty:
                # Create separate sheets for each scenario
                for scenario in scenarios:
                    # Get columns for this scenario
                    scenario_cols = ['Material'] + ['2023'] + [col for col in df.columns if scenario in col]
                    scenario_df = df[scenario_cols].copy()
                    
                    # Clean up the dataframe
                    numeric_cols = scenario_df.select_dtypes(include=[np.number]).columns
                    if not numeric_cols.empty:
                        scenario_df[numeric_cols] = scenario_df[numeric_cols].round(3)
                    
                    # Create sheet name
                    sheet_name = clean_sheet_name(f"{tech}_{scenario}")
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
    print("\nReading EV data...")
    df = pd.read_excel('CM_Data_Explorer May 2024 (2).xlsx', sheet_name='4.3 EV')
    
    # Clean and analyze data
    print("\nCleaning and analyzing data...")
    tech_data, scenarios = clean_ev_data(df)
    
    # Save results
    try:
        save_to_excel(tech_data, scenarios)
        print("\nAnalysis complete! Data saved to '4_3_ev_scenarios.xlsx'")
    except Exception as e:
        print(f"\nError saving Excel file: {e}")

if __name__ == "__main__":
    main() 