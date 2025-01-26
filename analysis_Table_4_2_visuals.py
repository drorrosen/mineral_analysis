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

def load_data(filename='4_2_wind_scenarios.xlsx'):
    """Load data from Excel file and print basic information"""
    # First, print all available sheets
    excel_file = pd.ExcelFile(filename)
    print("\nAvailable sheets in the Excel file:")
    print(excel_file.sheet_names)
    
    # Load all sheets into a dictionary
    data = {}
    for sheet in excel_file.sheet_names:
        if sheet != 'Overview':  # Skip the overview sheet
            df = pd.read_excel(filename, sheet_name=sheet)
            data[sheet] = df
    
    return data

def create_material_trends(data, material):
    """Create trend analysis for a specific material across scenarios, separate figures for each case"""
    base_case = data['Base case']
    constrained = data['Constrained rare earth elements']
    
    # Filter for the specific material
    base_material = base_case[base_case['Material'] == material]
    constrained_material = constrained[constrained['Material'] == material]
    
    # Get all years (2023 plus years from scenario columns)
    years = ['2023']
    for col in base_material.columns:
        if any(scenario in col for scenario in ['Stated Policies_', 'Announced Pledges_', 'Net Zero_']):
            year = col.split('_')[-1]
            if year not in years:
                years.append(year)
    years.sort()  # Sort years chronologically
    
    scenarios = ['Stated Policies', 'Announced Pledges', 'Net Zero']
    
    # Create Base Case figure
    fig_base = go.Figure()
    
    for scenario in scenarios:
        scenario_cols = ['2023'] + [f"{scenario}_{year}" for year in years[1:]]
        fig_base.add_trace(go.Scatter(
            x=years,
            y=base_material[scenario_cols].values[0],
            name=scenario,
            line=dict(color=SCENARIO_COLORS[scenario], width=3),
            mode='lines+markers'
        ))
    
    fig_base.update_layout(
        title=f"{material} Demand Trends - Base Case",
        xaxis_title="Year",
        yaxis_title="Demand (kt)",
        height=600,
        width=1000,
        template="plotly_white",
        hovermode="x unified",
        showlegend=True,
        legend=dict(
            yanchor="top",
            y=0.99,
            xanchor="left",
            x=1.02
        )
    )
    
    fig_base.write_html(f'figure_4_2/{material.lower().replace(" ", "_")}_base_case_trends.html')
    
    # Create Constrained Case figure
    fig_constrained = go.Figure()
    
    for scenario in scenarios:
        scenario_cols = ['2023'] + [f"{scenario}_{year}" for year in years[1:]]
        fig_constrained.add_trace(go.Scatter(
            x=years,
            y=constrained_material[scenario_cols].values[0],
            name=scenario,
            line=dict(color=SCENARIO_COLORS[scenario], width=3),
            mode='lines+markers'
        ))
    
    fig_constrained.update_layout(
        title=f"{material} Demand Trends - Constrained Supply Case",
        xaxis_title="Year",
        yaxis_title="Demand (kt)",
        height=600,
        width=1000,
        template="plotly_white",
        hovermode="x unified",
        showlegend=True,
        legend=dict(
            yanchor="top",
            y=0.99,
            xanchor="left",
            x=1.02
        )
    )
    
    fig_constrained.write_html(f'figure_4_2/{material.lower().replace(" ", "_")}_constrained_case_trends.html')

def create_comparison_heatmap(data):
    """Create heatmap comparing base case vs constrained supply in 2050"""
    base_case = data['Base case']
    constrained = data['Constrained rare earth elements']
    
    # Calculate percentage differences for 2050
    scenarios = ['Stated Policies', 'Announced Pledges', 'Net Zero']
    materials = base_case['Material'].unique()
    
    diff_data = []
    for material in materials:
        base_row = base_case[base_case['Material'] == material]
        const_row = constrained[constrained['Material'] == material]
        
        row_data = {'Material': material}
        for scenario in scenarios:
            base_val = base_row[f'{scenario}_2050'].values[0]
            const_val = const_row[f'{scenario}_2050'].values[0]
            pct_diff = ((const_val - base_val) / base_val) * 100
            row_data[scenario] = pct_diff
        diff_data.append(row_data)
    
    diff_df = pd.DataFrame(diff_data)
    
    # Create heatmap
    fig = go.Figure(data=go.Heatmap(
        z=diff_df[scenarios].values,
        x=scenarios,
        y=diff_df['Material'],
        colorscale='RdBu',
        zmid=0,
        text=np.round(diff_df[scenarios].values, 1),
        texttemplate='%{text}%',
        textfont={"size": 10},
        colorbar_title="% Difference<br>(Constrained vs Base)"
    ))
    
    fig.update_layout(
        title="Impact of Supply Constraints in 2050<br>(% difference from base case)",
        height=800,
        width=1000,
        template="plotly_white"
    )
    
    fig.write_html('figure_4_2/supply_constraint_impact.html')

def create_growth_analysis(data):
    """Create growth rate analysis for both cases"""
    base_case = data['Base case']
    constrained = data['Constrained rare earth elements']
    
    scenarios = ['Stated Policies', 'Announced Pledges', 'Net Zero']
    materials = base_case['Material'].unique()
    
    growth_data = []
    for material in materials:
        base_row = base_case[base_case['Material'] == material]
        const_row = constrained[constrained['Material'] == material]
        
        row_data = {'Material': material}
        for scenario in scenarios:
            # Base case growth
            base_2023 = base_row['2023'].values[0]
            base_2050 = base_row[f'{scenario}_2050'].values[0]
            base_growth = ((base_2050 / base_2023) - 1) * 100
            row_data[f'{scenario} (Base)'] = base_growth
            
            # Constrained case growth
            const_2023 = const_row['2023'].values[0]
            const_2050 = const_row[f'{scenario}_2050'].values[0]
            const_growth = ((const_2050 / const_2023) - 1) * 100
            row_data[f'{scenario} (Constrained)'] = const_growth
        
        growth_data.append(row_data)
    
    growth_df = pd.DataFrame(growth_data)
    
    # Create bar chart
    fig = go.Figure()
    
    for scenario in scenarios:
        # Base case bars (solid)
        fig.add_trace(go.Bar(
            name=f'{scenario} (Base)',
            x=growth_df['Material'],
            y=growth_df[f'{scenario} (Base)'],
            marker_color=SCENARIO_COLORS[scenario]
        ))
        
        # Constrained case bars (with pattern)
        fig.add_trace(go.Bar(
            name=f'{scenario} (Constrained)',
            x=growth_df['Material'],
            y=growth_df[f'{scenario} (Constrained)'],
            marker=dict(
                color=SCENARIO_COLORS[scenario],
                pattern=dict(
                    shape="/",  # Valid pattern shape for diagonal lines
                    size=10,    # Size of the pattern
                    solidity=0.5  # Opacity of the pattern
                )
            )
        ))
    
    fig.update_layout(
        title="Growth Rates 2023-2050 by Scenario and Case",
        xaxis_title="Material",
        yaxis_title="Growth Rate (%)",
        barmode='group',
        height=600,
        width=1200,
        template="plotly_white",
        xaxis_tickangle=-45,
        legend=dict(
            yanchor="top",
            y=0.99,
            xanchor="left",
            x=1.02
        )
    )
    
    fig.write_html('figure_4_2/growth_rates.html')

def create_statistical_summary(data):
    """Create statistical summary tables for both cases"""
    for case, df in [('Base case', data['Base case']), 
                    ('Constrained', data['Constrained rare earth elements'])]:
        
        stats_data = []
        for _, row in df.iterrows():
            material = row['Material']
            base_value = row['2023']
            
            for scenario in ['Stated Policies', 'Announced Pledges', 'Net Zero']:
                final_value = row[f'{scenario}_2050']
                total_growth = ((final_value / base_value) - 1) * 100
                cagr = (np.power(final_value / base_value, 1/27) - 1) * 100
                
                stats_data.append({
                    'Material': material,
                    'Scenario': scenario,
                    '2023 Value': base_value,
                    '2050 Value': final_value,
                    'Total Growth (%)': total_growth,
                    'CAGR (%)': cagr
                })
        
        stats_df = pd.DataFrame(stats_data)
        
        fig = go.Figure(data=[go.Table(
            header=dict(
                values=list(stats_df.columns),
                fill_color='paleturquoise',
                align='left'
            ),
            cells=dict(
                values=[stats_df[col] for col in stats_df.columns],
                fill_color='lavender',
                align='left',
                format=[None, None, '.2f', '.2f', '.1f', '.1f']
            )
        )])
        
        fig.update_layout(
            title=f"Statistical Summary - {case}",
            height=800,
            width=1200
        )
        
        fig.write_html(f'figure_4_2/statistics_{case.lower().replace(" ", "_")}.html')

def main():
    # Create figure directory if it doesn't exist
    if not os.path.exists('figure_4_2'):
        os.makedirs('figure_4_2')
    
    # Load data
    data = load_data()
    
    # Create visualizations
    # Material-specific trends
    materials = data['Base case']['Material'].unique()
    for material in materials:
        create_material_trends(data, material)
    
    # Overall analysis
    create_comparison_heatmap(data)
    create_growth_analysis(data)
    create_statistical_summary(data)
    
    print("\nAnalysis complete! Created visualizations in 'figure_4_2' directory:")
    print("1. Individual material trend analysis")
    print("2. Supply constraint impact heatmap")
    print("3. Growth rate analysis")
    print("4. Statistical summaries")

if __name__ == "__main__":
    main()
