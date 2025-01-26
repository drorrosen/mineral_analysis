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
    """Create dataframes from the electricity networks data"""
    # Define the data
    data = {
        'Base case': [4143.5, 4118.4],
        'Technology': ['Base case', 'Wider direct current (DC) technology development'],
        'Material': ['Copper', 'Copper']
    }
    
    # Years for each scenario
    years = ['2030', '2035', '2040', '2045', '2050']
    
    # Stated Policies scenario data
    stated_policies = [
        [6104.0, 6168.2, 6125.2, 6312.1, 6084.4],  # Base case
        [5954.34, 5923.00, 5790.91, 5868.96, 5565.44]  # DC technology
    ]
    
    # Announced Pledges scenario data
    announced_pledges = [
        [6807.7, 7542.0, 8186.3, 8471.3, 8096.0],  # Base case
        [6645.61, 7251.26, 7745.24, 7878.29, 7406.93]  # DC technology
    ]
    
    # Net Zero scenario data
    net_zero = [
        [7939.7, 9793.4, 10510.1, 10437.8, 9245.5],  # Base case
        [7755.7, 9422.8, 9934.3, 9685.9, 8447.6]  # DC technology
    ]
    
    # Create the main dataframe
    df = pd.DataFrame(data)
    
    # Add scenario data
    for i, year in enumerate(years):
        df[f'Stated Policies_{year}'] = [row[i] for row in stated_policies]
        df[f'Announced Pledges_{year}'] = [row[i] for row in announced_pledges]
        df[f'Net Zero_{year}'] = [row[i] for row in net_zero]
    
    return df

def create_technology_trends(df):
    """Create trend analysis for each technology"""
    technologies = df['Technology'].unique()
    years = ['2023'] + ['2030', '2035', '2040', '2045', '2050']
    scenarios = ['Stated Policies', 'Announced Pledges', 'Net Zero']
    
    # Define consistent colors for scenarios
    SCENARIO_COLORS = {
        'Stated Policies': '#1f77b4',      # Blue
        'Announced Pledges': '#ff7f0e',    # Orange
        'Net Zero': '#2ca02c'              # Green
    }
    
    for technology in technologies:
        fig = go.Figure()
        tech_data = df[df['Technology'] == technology].iloc[0]
        
        for scenario in scenarios:
            values = [tech_data['Base case']]  # Start with 2023 value
            for year in years[1:]:
                values.append(tech_data[f'{scenario}_{year}'])
            
            fig.add_trace(go.Scatter(
                x=years,
                y=values,
                name=scenario,
                line=dict(color=SCENARIO_COLORS[scenario], width=3),
                mode='lines+markers'
            ))
        
        fig.update_layout(
            title=f"Copper Demand Trends - {technology}",
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
        safe_tech_name = technology.lower().replace(" ", "_").replace("(", "").replace(")", "")
        fig.write_html(f'figure_4_5/{safe_tech_name}_trends.html')

def create_scenario_comparison(df):
    """Create heatmap comparing scenarios in 2050"""
    technologies = df['Technology'].unique()
    scenarios = ['Stated Policies', 'Announced Pledges', 'Net Zero']
    
    values_2050 = []
    for technology in technologies:
        tech_data = df[df['Technology'] == technology].iloc[0]
        values_2050.append([tech_data[f'{scenario}_2050'] for scenario in scenarios])
    
    fig = go.Figure(data=go.Heatmap(
        z=values_2050,
        x=scenarios,
        y=technologies,
        colorscale='RdYlBu',
        text=np.round(values_2050, 1),
        texttemplate='%{text}',
        textfont={"size": 10},
        colorbar_title="Demand (kt)"
    ))
    
    fig.update_layout(
        title="2050 Demand Comparison Across Scenarios",
        height=600,
        width=1000,
        template="plotly_white"
    )
    
    fig.write_html('figure_4_5/scenario_comparison_2050.html')

def create_growth_analysis(df):
    """Create growth rate analysis"""
    technologies = df['Technology'].unique()
    scenarios = ['Stated Policies', 'Announced Pledges', 'Net Zero']
    
    growth_data = []
    for technology in technologies:
        tech_data = df[df['Technology'] == technology].iloc[0]
        base_value = tech_data['Base case']
        
        for scenario in scenarios:
            growth = ((tech_data[f'{scenario}_2050'] - base_value) / base_value) * 100
            growth_data.append({
                'Technology': technology,
                'Scenario': scenario,
                'Growth': growth
            })
    
    growth_df = pd.DataFrame(growth_data)
    
    fig = px.bar(growth_df, x='Technology', y='Growth', color='Scenario',
                 title='Growth Rate (2023-2050)',
                 labels={'Growth': 'Growth Rate (%)', 'Technology': 'Technology'},
                 color_discrete_map=SCENARIO_COLORS,
                 barmode='group')
    
    fig.update_layout(
        height=600,
        width=1200,
        template="plotly_white",
        xaxis_tickangle=-45
    )
    
    fig.write_html('figure_4_5/growth_rates.html')

def create_technology_comparison(df):
    """Create comparison analysis between technologies"""
    technologies = df['Technology'].unique()
    years = ['2023'] + ['2030', '2035', '2040', '2045', '2050']
    scenarios = ['Stated Policies', 'Announced Pledges', 'Net Zero']
    
    # Create subplots for each scenario
    fig = make_subplots(
        rows=1, cols=3, 
        subplot_titles=scenarios,
        shared_yaxes=True
    )
    
    # Define consistent colors for technologies
    TECH_COLORS = {
        'Base case': '#1f77b4',  # Blue
        'Wider direct current (DC) technology development': '#ff7f0e'  # Orange
    }
    
    for i, scenario in enumerate(scenarios, 1):
        for j, technology in enumerate(technologies):
            tech_data = df[df['Technology'] == technology].iloc[0]
            
            values = [tech_data['Base case']]  # Start with 2023 value
            for year in years[1:]:
                values.append(tech_data[f'{scenario}_{year}'])
            
            fig.add_trace(
                go.Scatter(
                    x=years,
                    y=values,
                    name=f"{technology}",
                    line=dict(
                        width=3,
                        color=TECH_COLORS[technology],
                        dash='dash' if j == 1 else None  # Add dash for DC technology
                    ),
                    showlegend=(i == 1)  # Show legend only for first scenario
                ),
                row=1, col=i
            )
    
    fig.update_layout(
        title="Technology Comparison Across Scenarios",
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
    
    fig.write_html('figure_4_5/technology_comparison.html')

def main():
    # Create figure directory if it doesn't exist
    if not os.path.exists('figure_4_5'):
        os.makedirs('figure_4_5')
    
    # Create dataframe
    df = create_dataframe()
    
    # Create visualizations
    create_technology_trends(df)
    create_scenario_comparison(df)
    create_growth_analysis(df)
    create_technology_comparison(df)
    
    print("\nAnalysis complete! Created visualizations in 'figure_4_5' directory:")
    print("1. Individual technology trend analysis")
    print("2. 2050 scenario comparison heatmap")
    print("3. Growth rate analysis")
    print("4. Technology comparison analysis")

if __name__ == "__main__":
    main()
