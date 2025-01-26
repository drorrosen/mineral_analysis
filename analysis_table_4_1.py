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

def clean_solar_data(df):
    """Clean and structure the solar PV data"""
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
        'Comeback of high Cd-Te technology': None,
        'Wider adoption of Ga-As technology': None,
        'Wider adoption of perovskite solar cells': None
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
        'Cadmium', 'Copper', 'Gallium', 'Germanium', 'Indium', 'Lead',
        'Molybdenum', 'Nickel', 'Selenium', 'Silicon', 'Silver',
        'Tellurium', 'Tin', 'Zinc', 'Arsenic'
    ]
    
    # Process each section
    technologies = {}
    
    for tech_name, start_idx, end_idx in section_indices:
        print(f"\nProcessing section: {tech_name}")
        section_df = df.iloc[start_idx:end_idx].copy()
        
        # Find year row in this section
        year_row_idx = None
        for idx, row in section_df.iterrows():
            if pd.notna(row.iloc[1]) and row.iloc[1] == 2023:
                year_row_idx = idx - start_idx
                break
        
        if year_row_idx is None:
            print(f"No year row found in section {tech_name}")
            continue
        
        year_row = section_df.iloc[year_row_idx]
        
        # Map columns to years
        year_cols = {}
        for col in range(len(section_df.columns)):
            if pd.notna(year_row[col]) and isinstance(year_row[col], (int, float)):
                year_cols[int(year_row[col])] = col
        
        print(f"Year columns: {year_cols}")
        
        # Get data rows
        data_rows = []
        for idx in range(year_row_idx + 1, len(section_df)):
            row = section_df.iloc[idx]
            material = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ''
            
            if material in materials:
                row_data = {'Category': material}
                
                # Get 2023 value (base year)
                base_value = row.iloc[year_cols[2023]]
                
                # Add data for each scenario
                for scenario in scenarios:
                    # Add 2023 value (same for all scenarios)
                    row_data[f"{scenario}_2023"] = base_value
                    
                    # Add other years
                    for year in [2030, 2035, 2040, 2045, 2050]:
                        col_idx = year_cols[year]
                        row_data[f"{scenario}_{year}"] = row.iloc[col_idx]
                
                data_rows.append(row_data)
        
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
    with pd.ExcelWriter('4_1_solar_pv_scenarios.xlsx', engine='openpyxl', mode='w') as writer:
        # Save overview sheet
        pd.DataFrame(['Solar PV Technology Analysis by Scenario']).to_excel(writer, sheet_name='Overview', index=False)
        
        # Save data for each technology
        for tech, df in tech_data.items():
            if not df.empty:
                # Create separate sheets for each scenario
                for scenario in scenarios:
                    # Get columns for this scenario
                    scenario_cols = ['Category'] + [col for col in df.columns if scenario in col]
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
    print("\nReading solar PV data...")
    df = pd.read_excel('CM_Data_Explorer May 2024 (2).xlsx', sheet_name='4.1 Solar PV')
    
    # Clean and analyze data
    print("\nCleaning and analyzing data...")
    tech_data, scenarios = clean_solar_data(df)
    
    # Save results
    try:
        save_to_excel(tech_data, scenarios)
        print("\nAnalysis complete! Data saved to '4_1_solar_pv_scenarios.xlsx'")
    except Exception as e:
        print(f"\nError saving Excel file: {e}")

if __name__ == "__main__":
    main()