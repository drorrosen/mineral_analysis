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
    """Create dataframes from the hydrogen technologies data"""
    # Define the data
    data = {
        'Material': ['Copper', 'Cobalt', 'Iridium', 'Nickel', 'PGMs (other than iridium)', 
                    'Zirconium', 'Yttrium', 'Total hydrogen technologies'],
        'Base case': [0.0, 0.0, 0.0, 1.0, 0.0, 0.1, 0.0, 1.1]
    }
    
    # Years for each scenario
    years = ['2030', '2035', '2040', '2045', '2050']
    
    # Stated Policies scenario data
    stated_policies = [
        [0.0, 0.0, 0.0, 0.0, 0.0],      # Copper
        [0.0, 0.0, 0.0, 0.0, 0.0],      # Cobalt
        [0.0, 0.0, 0.0, 0.0, 0.0],      # Iridium
        [4.7, 5.2, 4.1, 3.9, 5.4],      # Nickel
        [0.0, 0.0, 0.0, 0.0, 0.0],      # PGMs
        [0.6, 0.7, 0.6, 0.5, 0.7],      # Zirconium
        [0.0, 0.0, 0.0, 0.0, 0.0],      # Yttrium
        [5.4, 5.9, 4.7, 4.5, 6.1]       # Total
    ]
    
    # Announced Pledges scenario data
    announced_pledges = [
        [0.1, 0.1, 0.1, 0.1, 0.2],      # Copper
        [0.0, 0.0, 0.0, 0.0, 0.1],      # Cobalt
        [0.0, 0.0, 0.0, 0.0, 0.0],      # Iridium
        [32.8, 32.3, 32.4, 45.3, 64.0],  # Nickel
        [0.0, 0.0, 0.1, 0.1, 0.1],      # PGMs
        [4.5, 4.4, 4.4, 6.2, 8.7],      # Zirconium
        [0.0, 0.0, 0.0, 0.0, 0.0],      # Yttrium
        [37.4, 36.9, 37.0, 51.7, 73.0]   # Total
    ]
    
    # Net Zero scenario data
    net_zero = [
        [0.2, 0.2, 0.2, 0.1, 0.2],      # Copper
        [0.1, 0.1, 0.1, 0.1, 0.1],      # Cobalt
        [0.0, 0.0, 0.0, 0.0, 0.0],      # Iridium
        [74.0, 80.1, 68.3, 58.5, 75.4],  # Nickel
        [0.0, 0.1, 0.1, 0.1, 0.1],      # PGMs
        [10.1, 10.9, 9.3, 8.0, 10.3],    # Zirconium
        [0.0, 0.0, 0.0, 0.0, 0.0],      # Yttrium
        [84.4, 91.4, 77.9, 66.8, 86.1]   # Total
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
            hovermode="x unified",
            legend=dict(
                yanchor="top",
                y=0.99,
                xanchor="left",
                x=1.02
            )
        )
        
        # Save the figure
        safe_material = material.lower().replace(" ", "_").replace("(", "").replace(")", "")
        fig.write_html(f'figure_4_6/{safe_material}_trends.html')

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
        text=np.round(values_2050, 2),
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
    
    fig.write_html('figure_4_6/scenario_comparison_2050.html')

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
            elif material_data[f'{scenario}_2050'] > 0:
                # Handle cases where base is 0 but there is growth
                growth_data.append({
                    'Material': material,
                    'Scenario': scenario,
                    'Growth': float('inf')
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
    
    fig.write_html('figure_4_6/growth_rates.html')

def create_material_comparison(df):
    """Create comparison analysis between key materials"""
    # Select main materials (excluding total)
    materials = [m for m in df['Material'].unique() if 'Total' not in m]
    years = ['2023'] + ['2030', '2035', '2040', '2045', '2050']
    scenarios = ['Stated Policies', 'Announced Pledges', 'Net Zero']
    
    # Create subplots for each scenario
    fig = make_subplots(
        rows=1, cols=3, 
        subplot_titles=scenarios,
        shared_yaxes=True
    )
    
    # Define color scale for materials
    material_colors = px.colors.qualitative.Set3[:len(materials)]
    color_map = dict(zip(materials, material_colors))
    
    for i, scenario in enumerate(scenarios, 1):
        for material in materials:
            material_data = df[df['Material'] == material].iloc[0]
            
            values = [material_data['Base case']]  # Start with 2023 value
            for year in years[1:]:
                values.append(material_data[f'{scenario}_{year}'])
            
            fig.add_trace(
                go.Scatter(
                    x=years,
                    y=values,
                    name=material,
                    line=dict(
                        color=color_map[material],
                        width=3
                    ),
                    showlegend=(i == 1)  # Show legend only for first scenario
                ),
                row=1, col=i
            )
    
    fig.update_layout(
        title="Material Comparison Across Scenarios",
        height=600,
        width=1500,
        template="plotly_white",
        hovermode="x unified",
        legend=dict(
            yanchor="top",
            y=0.99,
            xanchor="left",
            x=1.02
        )
    )
    
    # Update x and y axis labels
    fig.update_xaxes(title_text="Year")
    fig.update_yaxes(title_text="Demand (kt)", col=1)
    
    fig.write_html('figure_4_6/material_comparison.html')

def main():
    # Create figure directory if it doesn't exist
    if not os.path.exists('figure_4_6'):
        os.makedirs('figure_4_6')
    
    # Create dataframe
    df = create_dataframe()
    
    # Create visualizations
    create_mineral_trends(df)
    create_scenario_comparison(df)
    create_growth_analysis(df)
    create_material_comparison(df)
    
    print("\nAnalysis complete! Created visualizations in 'figure_4_6' directory:")
    print("1. Individual mineral trend analysis")
    print("2. 2050 scenario comparison heatmap")
    print("3. Growth rate analysis")
    print("4. Material comparison analysis")

if __name__ == "__main__":
    main()
