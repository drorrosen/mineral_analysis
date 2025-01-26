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

# Define consistent material colors (using qualitative colors)
MATERIAL_COLORS = px.colors.qualitative.Set3  # Will be assigned to materials in order

def get_material_color_dict(materials):
    """Create consistent color mapping for materials"""
    return {material: MATERIAL_COLORS[i % len(MATERIAL_COLORS)] 
            for i, material in enumerate(materials)}

def verify_2023_values():
    """
    Verify that 2023 values are the same across scenarios for each technology and material
    """
    file_path = '4.1 solar_pv_scenarios.xlsx'
    try:
        excel_file = pd.ExcelFile(file_path)
    except FileNotFoundError:
        print(f"\nError: Could not find {file_path}")
        print("Please make sure the file exists in the current directory.")
        raise

    # Group data by technology and material
    tech_groups = {}
    for sheet_name in excel_file.sheet_names:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        tech = df['Technology'].iloc[0]  # Get technology name from first row
        
        if tech not in tech_groups:
            tech_groups[tech] = []
        tech_groups[tech].append((sheet_name, df))
    
    print("\nVerifying 2023 values across scenarios:")
    inconsistencies = []
    
    for tech, sheet_dfs in tech_groups.items():
        print(f"\n{tech}:")
        
        # Get unique materials
        materials = sheet_dfs[0][1]['Material'].unique()
        
        for material in materials:
            values_2023 = {}
            for sheet_name, df in sheet_dfs:
                material_scenarios = df[df['Material'] == material]
                for _, row in material_scenarios.iterrows():
                    scenario = row['Scenario']
                    value_2023 = row['2023.0']
                    if not pd.isna(value_2023):  # Only store non-NaN values
                        values_2023[f"{sheet_name}_{scenario}"] = value_2023
            
            # Check if all non-NaN 2023 values are the same
            unique_values = set(v for v in values_2023.values() if not pd.isna(v))
            if len(unique_values) > 1:
                print(f"  Warning: Inconsistent 2023 values for {material}:")
                for scenario, value in values_2023.items():
                    print(f"    {scenario}: {value}")
                inconsistencies.append((tech, material, values_2023))
            elif unique_values:  # Only print if we have any non-NaN values
                print(f"  {material}: {list(unique_values)[0]} (consistent)")
    
    return tech_groups, inconsistencies

def load_organized_data():
    """
    Load and organize data from the Excel file
    """
    file_path = '4.1 solar_pv_scenarios.xlsx'
    try:
        excel_file = pd.ExcelFile(file_path)
    except FileNotFoundError:
        print(f"\nError: Could not find {file_path}")
        print("Please make sure the file exists in the current directory.")
        raise

    # Load data for each technology
    organized_data = {}
    
    for sheet_name in excel_file.sheet_names:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        tech = df['Technology'].iloc[0]
        
        if tech not in organized_data:
            organized_data[tech] = {
                'scenarios': {},
                '2023': {}
            }
        
        # Get unique materials
        materials = df['Material'].unique()
        
        # Store data for each scenario
        for scenario in df['Scenario'].unique():
            scenario_data = df[df['Scenario'] == scenario]
            
            # Create dictionary for this scenario if it doesn't exist
            if scenario not in organized_data[tech]['scenarios']:
                organized_data[tech]['scenarios'][scenario] = {}
            
            # Store data for each material
            for material in materials:
                material_data = scenario_data[scenario_data['Material'] == material].iloc[0]
                
                # Store 2023 value if not already stored and not NaN
                if material not in organized_data[tech]['2023'] and not pd.isna(material_data['2023.0']):
                    organized_data[tech]['2023'][material] = material_data['2023.0']
                
                # Store other years' data
                year_data = {}
                for year in ['2030', '2035.0', '2040.0', '2045.0', '2050.0']:
                    year_data[year.replace('.0', '')] = material_data[year]
                
                organized_data[tech]['scenarios'][scenario][material] = year_data
    
    return organized_data

def inspect_excel_file():
    """
    Print detailed information about the Excel file structure and content
    """
    file_path = '4.1 solar_pv_scenarios.xlsx'
    try:
        excel_file = pd.ExcelFile(file_path)
        print("\nExcel File Structure:")
        print("=" * 80)
        print(f"Found {len(excel_file.sheet_names)} sheets:")
        
        for sheet_name in excel_file.sheet_names:
            print(f"\n\nSheet: '{sheet_name}'")
            print("-" * 80)
            
            # Read the sheet
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            
            # Print basic info
            print(f"Shape: {df.shape}")
            print("\nColumns:")
            for col in df.columns:
                print(f"  - {col}")
            
            print("\nFirst few rows:")
            print(df.head().to_string())
            print("\n" + "=" * 80)
        
        return excel_file.sheet_names
        
    except FileNotFoundError:
        print(f"\nError: Could not find {file_path}")
        print("Please make sure the file exists in the current directory.")
        raise

def fix_2023_values():
    """
    Fix 2023 values to be consistent across scenarios for each technology and material
    """
    file_path = '4.1 solar_pv_scenarios.xlsx'
    try:
        excel_file = pd.ExcelFile(file_path)
    except FileNotFoundError:
        print(f"\nError: Could not find {file_path}")
        print("Please make sure the file exists in the current directory.")
        raise

    # Skip 'Overview' sheet
    sheet_names = [s for s in excel_file.sheet_names if s != 'Overview']
    
    # Group sheets by technology
    tech_groups = {}
    for sheet in sheet_names:
        tech = sheet.split('_Stated')[0]
        if tech not in tech_groups:
            tech_groups[tech] = []
        tech_groups[tech].append(sheet)
    
    # Process each technology group
    fixed_data = {}
    for tech, sheets in tech_groups.items():
        print(f"\nProcessing {tech}:")
        
        # Get all dataframes for this technology
        dfs = {sheet: pd.read_excel(file_path, sheet_name=sheet) for sheet in sheets}
        
        # Get the correct 2023 values from Stated Policies scenario
        stated_policies_sheet = [s for s in sheets if 'Stated Policies' in s][0]
        base_df = dfs[stated_policies_sheet]
        correct_2023_values = {}
        
        col_2023 = [col for col in base_df.columns if '2023' in col][0]
        for _, row in base_df.iterrows():
            material = row['Category']
            correct_2023_values[material] = row[col_2023]
        
        print(f"\nCorrect 2023 values for {tech}:")
        for material, value in correct_2023_values.items():
            print(f"  {material}: {value}")
        
        # Fix 2023 values in all scenarios
        fixed_dfs = {}
        for sheet, df in dfs.items():
            df_copy = df.copy()
            col_2023 = [col for col in df_copy.columns if '2023' in col][0]
            
            # Update 2023 values
            for material, value in correct_2023_values.items():
                df_copy.loc[df_copy['Category'] == material, col_2023] = value
            
            fixed_dfs[sheet] = df_copy
        
        fixed_data[tech] = fixed_dfs
    
    # Save fixed data to new Excel file
    output_file = '4.1 solar_pv_scenarios_fixed.xlsx'
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Copy overview sheet
        pd.read_excel(file_path, sheet_name='Overview').to_excel(writer, sheet_name='Overview', index=False)
        
        # Save fixed sheets
        for tech, dfs in fixed_data.items():
            for sheet_name, df in dfs.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
    
    print(f"\nFixed data saved to {output_file}")
    return fixed_data

def create_mineral_plots(data):
    """Create line plots for each material showing trends across scenarios"""
    for tech, tech_data in data.items():
        # Get all materials and scenarios
        materials = tech_data['2023'].keys()
        scenarios = tech_data['scenarios'].keys()
        years = ['2023', '2030', '2035', '2040', '2045', '2050']
        
        for material in materials:
            fig = go.Figure()
            
            # Add trace for each scenario
            for scenario in scenarios:
                # Get values for all years
                values = []
                for year in years:
                    if year == '2023':
                        values.append(tech_data['2023'][material])
                    else:
                        values.append(tech_data['scenarios'][scenario][material][year])
                
                fig.add_trace(go.Scatter(
                    x=years,
                    y=values,
                    name=scenario,
                    mode='lines+markers',
                    line=dict(color=SCENARIO_COLORS[scenario]),
                    hovertemplate="Year: %{x}<br>Value: %{y:.2f} kt<extra></extra>"
                ))
            
            fig.update_layout(
                title=f"{tech} - {material} Demand by Scenario",
                xaxis_title="Year",
                yaxis_title="Demand (kt)",
                hovermode='x unified',
                showlegend=True,
                template='plotly_white'
            )
            
            # Save the figure
            fig.write_html(f'figure_4_1/{tech}_{material}_demand.html'.replace(' ', '_'))

def create_scenario_comparison_heatmap(data):
    """Create heatmaps comparing 2050 values across scenarios"""
    for tech, tech_data in data.items():
        # Get materials and scenarios
        materials = list(tech_data['2023'].keys())
        scenarios = list(tech_data['scenarios'].keys())
        
        # Create matrix of 2050 values
        values = []
        for material in materials:
            row = []
            for scenario in scenarios:
                row.append(tech_data['scenarios'][scenario][material]['2050'])
            values.append(row)
        
        fig = go.Figure(data=go.Heatmap(
            z=values,
            x=scenarios,
            y=materials,
            colorscale='RdYlBu_r',
            text=values,
            texttemplate='%{text:.2f}',
            textfont={"size": 10},
            hoverongaps=False,
            hovertemplate='Material: %{y}<br>Scenario: %{x}<br>Value: %{z:.2f} kt<extra></extra>'
        ))
        
        fig.update_layout(
            title=f"{tech} - 2050 Demand Comparison",
            xaxis_title="Scenario",
            yaxis_title="Material",
            template='plotly_white'
        )
        
        fig.write_html(f'figure_4_1/{tech}_2050_comparison.html'.replace(' ', '_'))

def create_growth_analysis(data):
    """Create visualization of growth rates from 2023 to 2050"""
    for tech, tech_data in data.items():
        materials = list(tech_data['2023'].keys())
        scenarios = list(tech_data['scenarios'].keys())
        
        growth_data = []
        for material in materials:
            base_value = tech_data['2023'][material]
            for scenario in scenarios:
                if base_value != 0:
                    growth = ((tech_data['scenarios'][scenario][material]['2050'] - base_value) / base_value) * 100
                else:
                    growth = float('inf') if tech_data['scenarios'][scenario][material]['2050'] > 0 else 0
                
                growth_data.append({
                    'Material': material,
                    'Scenario': scenario,
                    'Growth': growth,
                    'Base Value': base_value,
                    '2050 Value': tech_data['scenarios'][scenario][material]['2050']
                })
        
        df = pd.DataFrame(growth_data)
        
        fig = px.bar(df, x='Material', y='Growth', color='Scenario', barmode='group',
                    title=f"{tech} - Growth Rate (2023-2050)",
                    labels={'Growth': 'Growth Rate (%)', 'Material': 'Material'},
                    hover_data=['Base Value', '2050 Value'])
        
        fig.update_layout(template='plotly_white')
        fig.write_html(f'figure_4_1/{tech}_growth_rates.html'.replace(' ', '_'))

def create_metal_comparison_plots(data):
    """Create comparison plots across all materials"""
    for tech, tech_data in data.items():
        # Compare 2050 values across materials
        materials = list(tech_data['2023'].keys())
        scenarios = list(tech_data['scenarios'].keys())
        
        fig = go.Figure()
        
        for scenario in scenarios:
            values_2050 = [tech_data['scenarios'][scenario][m]['2050'] for m in materials]
            
            fig.add_trace(go.Bar(
                name=scenario,
                x=materials,
                y=values_2050,
                hovertemplate="Material: %{x}<br>2050 Demand: %{y:.2f} kt<extra></extra>"
            ))
        
        fig.update_layout(
            title=f"{tech} - 2050 Demand Comparison Across Materials",
            xaxis_title="Material",
            yaxis_title="2050 Demand (kt)",
            barmode='group',
            template='plotly_white'
        )
        
        fig.write_html(f'figure_4_1/{tech}_material_comparison.html'.replace(' ', '_'))

def create_proportion_plots(data):
    """Create improved proportion plots for each technology"""
    for tech, tech_data in data.items():
        materials = list(tech_data['2023'].keys())
        scenarios = list(tech_data['scenarios'].keys())
        years = ['2023', '2030', '2035', '2040', '2045', '2050']
        material_colors = get_material_color_dict(materials)
        
        for scenario in scenarios:
            # Prepare data
            plot_data = []
            for year in years:
                year_data = {}
                total = 0
                
                # Get values for each material
                for material in materials:
                    if year == '2023':
                        value = tech_data['2023'][material]
                    else:
                        value = tech_data['scenarios'][scenario][material][year]
                    year_data[material] = value
                    total += value
                
                # Calculate percentages and store both absolute and percentage values
                if total > 0:
                    for material, value in year_data.items():
                        plot_data.append({
                            'Year': year,
                            'Material': material,
                            'Value (kt)': value,
                            'Percentage': (value / total) * 100
                        })
            
            df = pd.DataFrame(plot_data)
            
            # Create subplots: one for absolute values, one for percentages
            fig = make_subplots(
                rows=2, cols=1,
                subplot_titles=(f"Absolute Demand (kt)", "Material Share (%)"),
                vertical_spacing=0.15,
                row_heights=[0.6, 0.4]
            )
            
            # Add stacked bars for absolute values
            for material in materials:
                material_data = df[df['Material'] == material]
                fig.add_trace(
                    go.Bar(
                        name=material,
                        x=material_data['Year'],
                        y=material_data['Value (kt)'],
                        marker_color=material_colors[material],
                        hovertemplate="Year: %{x}<br>" +
                                    "Material: " + material + "<br>" +
                                    "Value: %{y:.2f} kt<br>" +
                                    "Share: %{customdata:.1f}%<extra></extra>",
                        customdata=material_data['Percentage']
                    ),
                    row=1, col=1
                )
            
            # Add stacked bars for percentages
            for material in materials:
                material_data = df[df['Material'] == material]
                fig.add_trace(
                    go.Bar(
                        name=material,
                        x=material_data['Year'],
                        y=material_data['Percentage'],
                        marker_color=material_colors[material],
                        hovertemplate="Year: %{x}<br>" +
                                    "Material: " + material + "<br>" +
                                    "Share: %{y:.1f}%<br>" +
                                    "Value: %{customdata:.2f} kt<extra></extra>",
                        customdata=material_data['Value (kt)'],
                        showlegend=False
                    ),
                    row=2, col=1
                )
            
            # Update layout
            fig.update_layout(
                title=dict(
                    text=f"{tech} - Material Composition ({scenario})",
                    x=0.5,
                    font=dict(size=16)
                ),
                barmode='stack',
                template='plotly_white',
                height=900,
                width=1200,
                showlegend=True,
                legend=dict(
                    yanchor="top",
                    y=0.99,
                    xanchor="left",
                    x=1.02
                )
            )
            
            # Update axes
            fig.update_xaxes(title_text="Year", row=2, col=1)
            fig.update_yaxes(title_text="Demand (kt)", row=1, col=1)
            fig.update_yaxes(title_text="Share (%)", row=2, col=1)
            
            # Save figure
            fig.write_html(f'figure_4_1/{tech}_{scenario}_composition.html'.replace(' ', '_'))

def create_statistics_table(data):
    """Create streamlined statistics tables for each technology"""
    for tech, tech_data in data.items():
        materials = list(tech_data['2023'].keys())
        scenarios = list(tech_data['scenarios'].keys())
        
        # Prepare statistics data
        stats_data = []
        for material in materials:
            material_stats = {
                'Material': material,
                'Base Value (2023) kt': round(tech_data['2023'][material], 2),
            }
            
            # Calculate statistics across scenarios
            for scenario in scenarios:
                scenario_data = tech_data['scenarios'][scenario][material]
                
                # 2050 value
                material_stats[f'{scenario} (2050) kt'] = round(scenario_data['2050'], 2)
                
                # Calculate growth rate
                base_value = tech_data['2023'][material]
                if base_value != 0:
                    growth = ((scenario_data['2050'] - base_value) / base_value) * 100
                    material_stats[f'{scenario} Growth (%)'] = round(growth, 1)
                
                # Calculate max value and its year
                all_values = [scenario_data[year] for year in ['2030', '2035', '2040', '2045', '2050']]
                max_value = max(all_values)
                max_year = ['2030', '2035', '2040', '2045', '2050'][all_values.index(max_value)]
                if max_value > scenario_data['2050']:  # Only show peak if it's not at 2050
                    material_stats[f'{scenario} Peak'] = f"{round(max_value, 2)} kt ({max_year})"
                else:
                    material_stats[f'{scenario} Peak'] = "At 2050"
            
            stats_data.append(material_stats)
        
        # Create DataFrame
        df_stats = pd.DataFrame(stats_data)
        
        # Create table visualization
        fig = go.Figure(data=[go.Table(
            header=dict(
                values=[col.replace('Announced Pledges', 'AP').replace('Stated Policies', 'SP')
                       .replace('Net Zero', 'NZ') for col in df_stats.columns],
                fill_color='paleturquoise',
                align='left',
                font=dict(size=12, color='black')
            ),
            cells=dict(
                values=[df_stats[col] for col in df_stats.columns],
                fill_color='lavender',
                align='left',
                font=dict(size=11),
                format=[None] + ['.2f' if 'kt' in col else '.1f' if '%' in col else None 
                               for col in df_stats.columns[1:]]
            )
        )])
        
        # Update layout
        fig.update_layout(
            title=dict(
                text=f"{tech} - Key Material Statistics",
                font=dict(size=16, color='black')
            ),
            width=1200,
            height=max(400, len(materials) * 30 + 100),
            margin=dict(t=50, l=20, r=20, b=20)
        )
        
        # Save figure
        fig.write_html(f'figure_4_1/{tech}_statistics_table.html'.replace(' ', '_'))
        
        # Create summary statistics
        summary_stats = []
        for scenario in scenarios:
            scenario_summary = {
                'Scenario': scenario,
                'Top Material by 2050': '',
                '2050 Demand (kt)': 0,
                'Highest Growth': '',
                'Growth Rate (%)': 0
            }
            
            for material in materials:
                base_value = tech_data['2023'][material]
                final_value = tech_data['scenarios'][scenario][material]['2050']
                
                # Find material with highest 2050 demand
                if final_value > scenario_summary['2050 Demand (kt)']:
                    scenario_summary['Top Material by 2050'] = material
                    scenario_summary['2050 Demand (kt)'] = final_value
                
                # Find material with highest growth
                if base_value != 0:
                    growth = ((final_value - base_value) / base_value) * 100
                    if growth > scenario_summary['Growth Rate (%)']:
                        scenario_summary['Highest Growth'] = material
                        scenario_summary['Growth Rate (%)'] = growth
            
            summary_stats.append(scenario_summary)
        
        # Create summary table
        df_summary = pd.DataFrame(summary_stats)
        
        fig_summary = go.Figure(data=[go.Table(
            header=dict(
                values=list(df_summary.columns),
                fill_color='paleturquoise',
                align='left',
                font=dict(size=12, color='black')
            ),
            cells=dict(
                values=[df_summary[col] for col in df_summary.columns],
                fill_color='lavender',
                align='left',
                font=dict(size=11),
                format=[None, None, '.2f', None, '.1f']
            )
        )])
        
        fig_summary.update_layout(
            title=dict(
                text=f"{tech} - Key Findings",
                font=dict(size=16, color='black')
            ),
            width=1000,
            height=200,
            margin=dict(t=50, l=20, r=20, b=20)
        )
        
        fig_summary.write_html(f'figure_4_1/{tech}_key_findings.html'.replace(' ', '_'))

def create_aggregate_plots(data):
    """Create aggregate plots combining all related visualizations"""
    
    # 1. Aggregate Mineral Trends
    for tech, tech_data in data.items():
        materials = list(tech_data['2023'].keys())
        scenarios = list(tech_data['scenarios'].keys())
        years = ['2023', '2030', '2035', '2040', '2045', '2050']
        
        # Calculate number of rows and columns for subplots
        n_materials = len(materials)
        n_cols = 3  # We'll use 3 columns
        n_rows = (n_materials + 2) // 3  # Round up division
        
        # Create subplots
        fig = make_subplots(
            rows=n_rows, cols=n_cols,
            subplot_titles=[f"{material} Demand" for material in materials],
            vertical_spacing=0.08,
            horizontal_spacing=0.05
        )
        
        # Add traces for each material
        for i, material in enumerate(materials):
            row = (i // n_cols) + 1
            col = (i % n_cols) + 1
            
            for scenario in scenarios:
                values = []
                for year in years:
                    if year == '2023':
                        values.append(tech_data['2023'][material])
                    else:
                        values.append(tech_data['scenarios'][scenario][material][year])
                
                fig.add_trace(
                    go.Scatter(
                        x=years,
                        y=values,
                        name=scenario,
                        mode='lines+markers',
                        line=dict(color=SCENARIO_COLORS[scenario]),
                        showlegend=True if i == 0 else False,  # Show legend only for first material
                        hovertemplate=f"{material}<br>Year: %{{x}}<br>Value: %{{y:.2f}} kt<extra></extra>"
                    ),
                    row=row, col=col
                )
        
        # Update layout
        fig.update_layout(
            title=dict(
                text=f"{tech} - All Materials Demand Trends",
                x=0.5,
                font=dict(size=20)
            ),
            height=300 * n_rows,
            width=1500,
            template='plotly_white',
            showlegend=True,
            legend=dict(
                yanchor="top",
                y=0.99,
                xanchor="left",
                x=1.02
            )
        )
        
        # Update all subplot axes
        for i in range(1, n_materials + 1):
            row = (i - 1) // n_cols + 1
            col = (i - 1) % n_cols + 1
            fig.update_xaxes(title_text="Year", row=row, col=col)
            fig.update_yaxes(title_text="Demand (kt)", row=row, col=col)
        
        # Save figure
        fig.write_html(f'figure_4_1/{tech}_aggregate_trends.html'.replace(' ', '_'))
    
    # 2. Aggregate Scenario Comparison
    for tech, tech_data in data.items():
        fig = make_subplots(
            rows=2, cols=2,
            subplot_titles=[
                "2050 Demand Comparison",
                "Growth Rates",
                "Material Composition",
                "Key Statistics"
            ],
            specs=[
                [{"type": "heatmap"}, {"type": "bar"}],
                [{"type": "bar"}, {"type": "table"}]
            ],
            vertical_spacing=0.15,
            horizontal_spacing=0.1
        )
        
        # Add heatmap (2050 comparison)
        materials = list(tech_data['2023'].keys())
        scenarios = list(tech_data['scenarios'].keys())
        values = [[tech_data['scenarios'][s][m]['2050'] for s in scenarios] for m in materials]
        
        fig.add_trace(
            go.Heatmap(
                z=values,
                x=scenarios,
                y=materials,
                colorscale='RdYlBu_r',
                text=[[f"{v:.2f}" for v in row] for row in values],
                texttemplate="%{text}",
                textfont={"size": 10},
                hovertemplate='Material: %{y}<br>Scenario: %{x}<br>Value: %{z:.2f} kt<extra></extra>'
            ),
            row=1, col=1
        )
        
        # Add growth rates
        growth_data = []
        for material in materials:
            base_value = tech_data['2023'][material]
            for scenario in scenarios:
                if base_value != 0:
                    growth = ((tech_data['scenarios'][scenario][material]['2050'] - base_value) / base_value) * 100
                    growth_data.append({
                        'Material': material,
                        'Scenario': scenario,
                        'Growth': growth
                    })
        
        df_growth = pd.DataFrame(growth_data)
        for scenario in scenarios:
            scenario_data = df_growth[df_growth['Scenario'] == scenario]
            fig.add_trace(
                go.Bar(
                    name=scenario,
                    x=scenario_data['Material'],
                    y=scenario_data['Growth'],
                    marker_color=SCENARIO_COLORS[scenario],
                    hovertemplate="Material: %{x}<br>Growth: %{y:.1f}%<extra></extra>"
                ),
                row=1, col=2
            )
        
        # Add composition for 2050
        composition_2050 = []
        for scenario in scenarios:
            values_2050 = [tech_data['scenarios'][scenario][m]['2050'] for m in materials]
            total = sum(values_2050)
            shares = [v/total * 100 for v in values_2050]
            
            fig.add_trace(
                go.Bar(
                    name=scenario,
                    x=materials,
                    y=shares,
                    marker_color=SCENARIO_COLORS[scenario],
                    hovertemplate="Material: %{x}<br>Share: %{y:.1f}%<extra></extra>"
                ),
                row=2, col=1
            )
        
        # Add key statistics table
        stats_summary = pd.DataFrame([{
            'Metric': 'Total 2050 Demand (kt)',
            **{s: sum(tech_data['scenarios'][s][m]['2050'] for m in materials) for s in scenarios}
        }, {
            'Metric': 'Average Growth (%)',
            **{s: np.mean([((tech_data['scenarios'][s][m]['2050']/tech_data['2023'][m] - 1) * 100) 
                          for m in materials if tech_data['2023'][m] != 0]) for s in scenarios}
        }])
        
        fig.add_trace(
            go.Table(
                header=dict(
                    values=['Metric'] + list(scenarios),
                    fill_color='paleturquoise',
                    align='left'
                ),
                cells=dict(
                    values=[stats_summary[col] for col in stats_summary.columns],
                    fill_color='lavender',
                    align='left',
                    format=[None] + ['.2f'] * len(scenarios)
                )
            ),
            row=2, col=2
        )
        
        # Update layout
        fig.update_layout(
            title=dict(
                text=f"{tech} - Comprehensive Analysis",
                x=0.5,
                font=dict(size=20)
            ),
            height=1200,
            width=1800,
            template='plotly_white',
            showlegend=True,
            barmode='group'
        )
        
        # Update axes labels
        fig.update_xaxes(title_text="Material", row=1, col=2)
        fig.update_yaxes(title_text="Growth Rate (%)", row=1, col=2)
        fig.update_xaxes(title_text="Material", row=2, col=1)
        fig.update_yaxes(title_text="Share (%)", row=2, col=1)
        
        # Save figure
        fig.write_html(f'figure_4_1/{tech}_comprehensive_analysis.html'.replace(' ', '_'))

def main():
    # First inspect the Excel file structure
    print("\nInspecting Excel file structure...")
    sheet_names = inspect_excel_file()
    
    # Then proceed with verification and analysis
    print("\nVerifying 2023 values...")
    tech_groups, inconsistencies = verify_2023_values()
    
    if inconsistencies:
        print("\nFound inconsistencies in 2023 values. Fixing them...")
        fixed_data = fix_2023_values()
    else:
        print("\nNo inconsistencies found in 2023 values.")
    
    # Create figures directory if it doesn't exist
    if not os.path.exists('figure_4_1'):
        os.makedirs('figure_4_1')
    
    # Load organized data (using fixed data if there were inconsistencies)
    data = load_organized_data()
    
    # Create visualizations
    create_mineral_plots(data)
    create_scenario_comparison_heatmap(data)
    create_growth_analysis(data)
    create_metal_comparison_plots(data)
    create_proportion_plots(data)
    create_statistics_table(data)
    create_aggregate_plots(data)
    
    print("\nAnalysis complete! Created visualizations in 'figure_4_1' directory:")
    print("1. Individual mineral comparisons")
    print("2. 2050 comparison heatmaps")
    print("3. Growth rate analysis")
    print("4. Mineral demand comparison")
    print("5. Proportion plots")
    print("6. Statistics tables")
    print("7. Aggregate plots")

if __name__ == "__main__":
    main()

    