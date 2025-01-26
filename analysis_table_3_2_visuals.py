import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import numpy as np
import os
from plotly.subplots import make_subplots

# Define consistent color scheme
SCENARIO_COLORS = {
    'Stated Policies': '#1f77b4',      # Blue
    'Announced Pledges': '#ff7f0e',    # Orange
    'Net Zero': '#2ca02c'              # Green
}

def load_data(filename='3.2 Cleantech demand by mineral.xlsx'):
    """Load data from Excel file"""
    # First, print all available sheets
    excel_file = pd.ExcelFile(filename)
    print("\nAvailable sheets in the Excel file:")
    print(excel_file.sheet_names)
    
    sheets = ['Stated Policies', 'Announced Pledges', 'Net Zero']
    data = {}
    
    for sheet in sheets:
        print(f"\n{'='*50}")
        print(f"Processing sheet: {sheet}")
        print(f"{'='*50}")
        
        df = pd.read_excel(filename, sheet_name=sheet)
        
        print("\nFirst few rows of raw data:")
        print(df.head())
        print("\nDataFrame info:")
        print(df.info())
        
        # Handle different scenarios
        if sheet == 'Stated Policies':
            years = ['2023', '2030', '2035', '2040', '2045', '2050']
        else:
            years = ['2030', '2035', '2040', '2045', '2050']
        
        # Clean column names (remove .0 from year columns)
        df.columns = [str(col).replace('.0', '') for col in df.columns]
        print("\nCleaned columns:", df.columns.tolist())
        
        data[sheet] = df
        print(f"\nFinal shape of {sheet} DataFrame:", df.shape)
    
    return data

def calculate_growth_rate(start_value, end_value):
    """Calculate growth rate handling zero values"""
    if start_value == 0:
        if end_value == 0:
            return 0
        else:
            # Use a small base value instead of 0
            start_value = 0.001
    return ((end_value / start_value) - 1) * 100

def create_top_metals_analysis(data):
    """Create trend analysis for top 5 metals in each scenario"""
    colors = px.colors.qualitative.Set2
    
    for scenario, df in data.items():
        base_year = '2023' if '2023' in df.columns else '2030'
        
        # Calculate growth rates handling zero values
        df_sorted = pd.DataFrame({
            'Metal': df['Metal'],
            'Growth': [calculate_growth_rate(row[base_year], row['2050']) for _, row in df.iterrows()]
        }).sort_values('Growth', ascending=False)
        
        # Get top 5 and bottom 5 metals
        top_5 = df_sorted.head(5)
        bottom_5 = df_sorted.tail(5)
        
        # Create and save growing metals figure
        fig_growing = go.Figure()
        
        for idx, (_, row) in enumerate(top_5.iterrows()):
            metal = row['Metal']
            growth = row['Growth']
            values = df[df['Metal'] == metal].iloc[0]
            
            fig_growing.add_trace(
                go.Scatter(
                    x=df.columns[1:],  # Skip 'Metal' column
                    y=values[1:],      # Skip 'Metal' value
                    name=f"{metal} (+{growth:.1f}%)",
                    mode='lines+markers',
                    line=dict(color=colors[idx], width=3),
                    marker=dict(size=8)
                )
            )
        
        fig_growing.update_layout(
            height=700,
            width=1200,
            title=dict(
                text=f"<b>Top 5 Growing Metals - {scenario}</b>",
                x=0.5,
                font=dict(size=20)
            ),
            xaxis=dict(
                title="Year",
                gridcolor='lightgrey',
                title_font=dict(size=14)
            ),
            yaxis=dict(
                title="Demand (kt)",
                gridcolor='lightgrey',
                title_font=dict(size=14)
            ),
            plot_bgcolor='white',
            paper_bgcolor='white',
            showlegend=True,
            legend=dict(
                yanchor="top",
                y=0.99,
                xanchor="left",
                x=1.02,
                font=dict(size=12)
            )
        )
        
        fig_growing.write_html(f'figure_3_2/top_growing_metals_{scenario.lower().replace(" ", "_")}.html')
        
        # Create and save declining metals figure
        fig_declining = go.Figure()
        
        for idx, (_, row) in enumerate(bottom_5.iterrows()):
            metal = row['Metal']
            growth = row['Growth']
            values = df[df['Metal'] == metal].iloc[0]
            
            fig_declining.add_trace(
                go.Scatter(
                    x=df.columns[1:],  # Skip 'Metal' column
                    y=values[1:],      # Skip 'Metal' value
                    name=f"{metal} ({growth:.1f}%)",
                    mode='lines+markers',
                    line=dict(color=colors[idx], width=3),
                    marker=dict(size=8)
                )
            )
        
        fig_declining.update_layout(
            height=700,
            width=1200,
            title=dict(
                text=f"<b>Top 5 Declining Metals - {scenario}</b>",
                x=0.5,
                font=dict(size=20)
            ),
            xaxis=dict(
                title="Year",
                gridcolor='lightgrey',
                title_font=dict(size=14)
            ),
            yaxis=dict(
                title="Demand (kt)",
                gridcolor='lightgrey',
                title_font=dict(size=14)
            ),
            plot_bgcolor='white',
            paper_bgcolor='white',
            showlegend=True,
            legend=dict(
                yanchor="top",
                y=0.99,
                xanchor="left",
                x=1.02,
                font=dict(size=12)
            )
        )
        
        fig_declining.write_html(f'figure_3_2/top_declining_metals_{scenario.lower().replace(" ", "_")}.html')

def create_statistics_table(data):
    """Create statistical summary tables for each scenario"""
    for scenario, df in data.items():
        # Calculate statistics
        stats_df = pd.DataFrame()
        stats_df['Metal'] = df['Metal']
        
        # Get base year (2023 for Stated Policies, 2030 for others)
        base_year = '2023' if '2023' in df.columns else '2030'
        years = 27 if base_year == '2023' else 20
        
        # Calculate growth rates handling zero values
        growth_rates = []
        cagr_rates = []
        
        for _, row in df.iterrows():
            start_val = row[base_year]
            end_val = row['2050']
            growth = calculate_growth_rate(start_val, end_val)
            growth_rates.append(growth)
            
            # Calculate CAGR
            if start_val == 0:
                start_val = 0.001
            if end_val == 0:
                cagr = -100  # Complete decline
            else:
                cagr = ((end_val/start_val) ** (1/years) - 1) * 100
            cagr_rates.append(cagr)
        
        stats_df[f'Growth {base_year}-2050 (%)'] = growth_rates
        stats_df['CAGR (%)'] = cagr_rates
        stats_df[f'{base_year} Value'] = df[base_year]
        stats_df['2050 Value'] = df['2050']
        
        # Sort by total growth
        stats_df = stats_df.sort_values(f'Growth {base_year}-2050 (%)', ascending=False)
        
        # Create table visualization
        fig = go.Figure(data=[go.Table(
            header=dict(
                values=list(stats_df.columns),
                fill_color='paleturquoise',
                align='left',
                font=dict(size=12)
            ),
            cells=dict(
                values=[stats_df[col] for col in stats_df.columns],
                fill_color='lavender',
                align='left',
                format=[
                    None,       # Metal
                    '.1f',     # Growth
                    '.1f',     # CAGR
                    '.1f',     # Base year value
                    '.1f'      # 2050 value
                ]
            )
        )])
        
        fig.update_layout(
            title=dict(
                text=f"Statistical Summary - {scenario}",
                x=0.5,
                font=dict(size=16)
            ),
            width=1200,
            height=800
        )
        
        fig.write_html(f'figure_3_2/statistics_{scenario.lower().replace(" ", "_")}.html')

def main():
    # Create figure directory if it doesn't exist
    if not os.path.exists('figure_3_2'):
        os.makedirs('figure_3_2')
    
    # Load data
    data = load_data()
    
    # Create visualizations
    create_top_metals_analysis(data)
    create_statistics_table(data)
    
    print("\nAnalysis complete! Created visualizations in 'figure_3_2' directory:")
    print("1. Top metals analysis for each scenario")
    print("2. Statistical summary tables")

if __name__ == "__main__":
    main()
