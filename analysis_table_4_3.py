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

def process_section_data(df, start_idx, end_idx):
    """Process data for a specific section"""
    section_df = df.iloc[start_idx:end_idx].copy()
    data_rows = []
    
    for _, row in section_df.iterrows():
        material = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ''
        if pd.notna(material) and material not in ['Base case', 'High material prices', 
            'Wider use of silicon-rich anodes', 'Faster uptake of solid state batteries',
            'Lower battery sizes', 'Limited battery size reduction']:
            
            row_data = {'Material': material}
            
            # Get values for all columns
            for col in range(len(row)):
                if pd.notna(row.iloc[col]):
                    row_data[df.columns[col]] = row.iloc[col]
            
            data_rows.append(row_data)
    
    return pd.DataFrame(data_rows) if data_rows else None

def clean_ev_data(df):
    """Clean and structure the EV data"""
    # Remove empty rows and columns
    df = df.dropna(how='all').dropna(axis=1, how='all')
    df = df.reset_index(drop=True)
    
    print("\nFirst few rows of raw data:")
    print(df.head())
    
    # Define sections
    sections = {
        'Base case': None,
        'High material prices': None,
        'Wider use of silicon-rich anodes': None,
        'Faster uptake of solid state batteries': None,
        'Lower battery sizes': None,
        'Limited battery size reduction': None
    }
    
    # Find section start indices
    for idx, row in df.iterrows():
        if pd.notna(row.iloc[0]):
            for section in sections.keys():
                if section in str(row.iloc[0]):
                    sections[section] = idx
    
    print("\nFound section indices:", sections)
    
    # Process each section
    section_data = {}
    section_names = list(sections.keys())
    
    for i, section in enumerate(section_names):
        start_idx = sections[section]
        if start_idx is not None:
            end_idx = sections[section_names[i + 1]] if i < len(section_names) - 1 else len(df)
            section_df = process_section_data(df, start_idx, end_idx)
            if section_df is not None:
                section_data[section] = section_df
                print(f"\nProcessed {section}:")
                print(f"Shape: {section_df.shape}")
                print("First few rows:")
                print(section_df.head())
    
    return section_data

def save_to_excel(section_data):
    """Save processed data to Excel"""
    with pd.ExcelWriter('4_3_ev_scenarios.xlsx', engine='openpyxl', mode='w') as writer:
        # Save overview sheet
        pd.DataFrame(['EV Battery Technology Analysis']).to_excel(writer, sheet_name='Overview', index=False)
        
        # Save each section to a separate sheet
        for section, df in section_data.items():
            sheet_name = clean_sheet_name(section)
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            # Auto-adjust column widths
            worksheet = writer.sheets[sheet_name]
            for idx, col in enumerate(df.columns):
                max_length = max(
                    df[col].astype(str).apply(len).max(),
                    len(str(col))
                ) + 2
                worksheet.column_dimensions[openpyxl.utils.get_column_letter(idx + 1)].width = max_length

def main():
    # Read the data
    print("\nReading EV data...")
    df = pd.read_excel('CM_Data_Explorer May 2024 (2).xlsx', sheet_name='4.3 EV')
    
    # Clean and analyze data
    print("\nCleaning and analyzing data...")
    section_data = clean_ev_data(df)
    
    # Save results
    try:
        save_to_excel(section_data)
        print("\nAnalysis complete! Data saved to '4_3_ev_scenarios.xlsx'")
    except Exception as e:
        print(f"\nError saving Excel file: {e}")

if __name__ == "__main__":
    main()