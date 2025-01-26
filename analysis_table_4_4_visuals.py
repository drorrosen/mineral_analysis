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

def create_dataframe():
    """Create dataframes from the grid battery storage data"""
    # Define the data
    data = {
        'Base case': [39.60, 2.53, 86.62, 9.22, 2.66, 12.11, 0.24, 0.00],
        'Material': ['Copper', 'Cobalt', 'Battery-grade graphite', 'Lithium', 
                    'Manganese', 'Nickel', 'Silicon', 'Vanadium']
    }
    
    # Years for each scenario
    years = ['2030', '2035', '2040', '2045', '2050']
    
    # Stated Policies scenario data
    stated_policies = [
        [175.75, 260.28, 380.68, 454.32, 544.67],  # Copper
        [4.13, 3.27, 0.00, 0.00, 0.00],            # Cobalt
        [303.46, 410.26, 540.58, 598.63, 677.00],  # Battery-grade graphite
        [34.81, 45.35, 59.39, 68.29, 76.64],       # Lithium
        [3.85, 3.04, 0.00, 0.00, 0.00],            # Manganese
        [17.57, 14.82, 0.00, 0.00, 0.00],          # Nickel
        [3.23, 5.79, 10.28, 14.61, 17.65],         # Silicon
        [21.96, 159.92, 323.65, 388.13, 469.11]    # Vanadium
    ]
    
    # Announced Pledges scenario data
    announced_pledges = [
        [220.85, 327.40, 509.25, 589.01, 701.53],  # Copper
        [5.20, 4.11, 0.00, 0.00, 0.00],            # Cobalt
        [381.32, 516.05, 723.16, 776.11, 871.96],  # Battery-grade graphite
        [43.74, 57.05, 79.44, 88.54, 98.72],       # Lithium
        [4.84, 3.83, 0.00, 0.00, 0.00],            # Manganese
        [22.08, 18.65, 0.00, 0.00, 0.00],          # Nickel
        [4.06, 7.29, 13.75, 18.94, 22.74],         # Silicon
        [27.59, 201.16, 432.96, 503.20, 604.20]    # Vanadium
    ]
    
    # Net Zero scenario data
    net_zero = [
        [281.99, 443.25, 654.39, 745.11, 897.64],  # Copper
        [6.63, 5.56, 0.00, 0.00, 0.00],            # Cobalt
        [486.90, 698.66, 929.26, 981.79, 1115.71], # Battery-grade graphite
        [55.85, 77.24, 102.09, 112.00, 126.31],    # Lithium
        [6.18, 5.18, 0.00, 0.00, 0.00],            # Manganese
        [28.19, 25.24, 0.00, 0.00, 0.00],          # Nickel
        [5.19, 9.87, 17.67, 23.95, 29.09],         # Silicon
        [35.23, 272.33, 556.35, 635.55, 773.10]    # Vanadium
    ]
    
    # Create the main dataframe
    df = pd.DataFrame(data)
    
    # Add scenario data
    for i, year in enumerate(years):
        df[f'Stated Policies_{year}'] = [row[i] for row in stated_policies]
        df[f'Announced Pledges_{year}'] = [row[i] for row in announced_pledges]
        df[f'Net Zero_{year}'] = [row[i] for row in net_zero]
    
    return df

def create_mineral_trends(df):
    """Create trend analysis for each material"""
    materials = df['Material'].unique()
    years = ['2023'] + ['2030', '2035', '2040', '2045', '2050']
    scenarios = ['Stated Policies', 'Announced Pledges', 'Net Zero']
    
    for material in materials:
        fig = go.Figure()
        material_data = df[df['Material'] == material].iloc[0]
        
        for scenario in scenarios:
            values = [material_data['Base case']]  # Start with 2023 value
            for year in years[1:]:
                values.append(material_data[f'{scenario}_{year}'])
            
            fig.add_trace(go.Scatter(
                x=years,
                y=values,
                name=scenario,
                line=dict(color=SCENARIO_COLORS[scenario], width=3),
                mode='lines+markers'
            ))
        
        fig.update_layout(
            title=f"{material} Demand Trends",
            xaxis_title="Year",
            yaxis_title="Demand (kt)",
            height=600,
            width=1000,
            template="plotly_white",
            hovermode="x unified"
        )
        
        # Save the figure
        fig.write_html(f'figure_4_4/{material.lower().replace(" ", "_")}_trends.html')

def create_scenario_comparison(df):
    """Create heatmap comparing scenarios in 2050"""
    materials = df['Material'].unique()
    scenarios = ['Stated Policies', 'Announced Pledges', 'Net Zero']
    
    values_2050 = []
    for material in materials:
        material_data = df[df['Material'] == material].iloc[0]
        values_2050.append([material_data[f'{scenario}_2050'] for scenario in scenarios])
    
    fig = go.Figure(data=go.Heatmap(
        z=values_2050,
        x=scenarios,
        y=materials,
        colorscale='RdYlBu',
        text=np.round(values_2050, 1),
        texttemplate='%{text}',
        textfont={"size": 10},
        colorbar_title="Demand (kt)"
    ))
    
    fig.update_layout(
        title="2050 Demand Comparison Across Scenarios",
        height=800,
        width=1000,
        template="plotly_white"
    )
    
    fig.write_html('figure_4_4/scenario_comparison_2050.html')

def create_growth_analysis(df):
    """Create growth rate analysis"""
    materials = df['Material'].unique()
    scenarios = ['Stated Policies', 'Announced Pledges', 'Net Zero']
    
    growth_data = []
    for material in materials:
        material_data = df[df['Material'] == material].iloc[0]
        base_value = material_data['Base case']
        
        for scenario in scenarios:
            if base_value != 0:
                growth = ((material_data[f'{scenario}_2050'] - base_value) / base_value) * 100
                growth_data.append({
                    'Material': material,
                    'Scenario': scenario,
                    'Growth': growth
                })
    
    growth_df = pd.DataFrame(growth_data)
    
    fig = px.bar(growth_df, x='Material', y='Growth', color='Scenario',
                 title='Growth Rate (2023-2050)',
                 labels={'Growth': 'Growth Rate (%)', 'Material': 'Material'},
                 color_discrete_map=SCENARIO_COLORS,
                 barmode='group')
    
    fig.update_layout(
        height=600,
        width=1200,
        template="plotly_white",
        xaxis_tickangle=-45
    )
    
    fig.write_html('figure_4_4/growth_rates.html')

def create_total_demand_analysis(df):
    """Create analysis of total demand over time"""
    years = ['2023'] + ['2030', '2035', '2040', '2045', '2050']
    scenarios = ['Stated Policies', 'Announced Pledges', 'Net Zero']
    
    # Calculate total demand for each year and scenario
    totals = {'2023': df['Base case'].sum()}
    for year in years[1:]:
        for scenario in scenarios:
            totals[f'{scenario}_{year}'] = df[f'{scenario}_{year}'].sum()
    
    fig = go.Figure()
    
    # Add trace for each scenario
    for scenario in scenarios:
        values = [totals['2023']]  # Start with 2023 value
        for year in years[1:]:
            values.append(totals[f'{scenario}_{year}'])
        
        fig.add_trace(go.Scatter(
            x=years,
            y=values,
            name=scenario,
            line=dict(color=SCENARIO_COLORS[scenario], width=3),
            mode='lines+markers'
        ))
    
    fig.update_layout(
        title="Total Mineral Demand for Grid Battery Storage",
        xaxis_title="Year",
        yaxis_title="Total Demand (kt)",
        height=600,
        width=1000,
        template="plotly_white",
        hovermode="x unified"
    )
    
    fig.write_html('figure_4_4/total_demand_trends.html')

def main():
    # Create figure directory if it doesn't exist
    if not os.path.exists('figure_4_4'):
        os.makedirs('figure_4_4')
    
    # Create dataframe
    df = create_dataframe()
    
    # Create visualizations
    create_mineral_trends(df)
    create_scenario_comparison(df)
    create_growth_analysis(df)
    create_total_demand_analysis(df)
    
    print("\nAnalysis complete! Created visualizations in 'figure_4_4' directory:")
    print("1. Individual mineral trend analysis")
    print("2. 2050 scenario comparison heatmap")
    print("3. Growth rate analysis")
    print("4. Total demand analysis")

if __name__ == "__main__":
    main()
