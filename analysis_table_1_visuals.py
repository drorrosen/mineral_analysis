import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import numpy as np
import os
import matplotlib.pyplot as plt

# Define consistent color scheme
SCENARIO_COLORS = {
    'Stated Policies': '#1f77b4',                    # Strong blue
    'Announced Pledges': '#d62728',                  # Strong red
    'Net Zero': '#2ca02c'                            # Strong green
}

# Define consistent material colors (using qualitative colors)
MATERIAL_COLORS = px.colors.qualitative.Set3

# And let's add a new constant for mineral colors
MINERAL_COLORS = {
    'Copper': '#ff7f0e',                            # Orange
    'Lithium': '#9467bd',                           # Purple
    'Nickel': '#8c564b',                            # Brown
    'Cobalt': '#e377c2',                            # Pink
    'Graphite all grades natural and': '#7f7f7f',   # Gray
    'Magnet rare earth elements': '#bcbd22'         # Olive
}

def get_material_color_dict(materials):
    """Create consistent color mapping for materials"""
    return {material: MATERIAL_COLORS[i % len(MATERIAL_COLORS)] 
            for i, material in enumerate(materials)}

def load_data(filename='1_organized_mineral_demand.xlsx'):
    """
    Load and inspect data from the Excel file
    """
    print(f"\nReading data from {filename}")
    
    try:
        # Read all sheets first
        excel_file = pd.ExcelFile(filename)
        all_sheets = excel_file.sheet_names
        print(f"\nAll sheets in file: {all_sheets}")
        
        print("\nINSPECTING ALL SHEETS:")
        print("=" * 80)
        
        # Print contents of all sheets first
        for sheet in all_sheets:
            df = pd.read_excel(filename, sheet_name=sheet)
            print(f"\nSheet: {sheet}")
            print("-" * 80)
            print(f"Shape: {df.shape}")
            print("\nColumns:", df.columns.tolist())
            print("\nFirst 10 rows:")
            print(df.head(10))
            print("=" * 80)
        
        # Now filter and load only mineral sheets
        mineral_sheets = [sheet for sheet in all_sheets 
                         if not (sheet.startswith('Overview') or sheet.startswith('Summary'))]
        
        dfs = {}
        for sheet in mineral_sheets:
            dfs[sheet] = pd.read_excel(filename, sheet_name=sheet)
        
        if not dfs:
            print("\nWarning: No mineral sheets found in the Excel file!")
        else:
            print(f"\nFound {len(dfs)} mineral sheets: {list(dfs.keys())}")
        
        return dfs
        
    except FileNotFoundError:
        print(f"\nError: Could not find {filename}")
        print("Please run analysis_table_1.py first to generate this file.")
        raise
    except Exception as e:
        print(f"\nError reading Excel file: {str(e)}")
        print(f"Error type: {type(e)}")
        print(f"Error details: {str(e)}")
        raise

def create_mineral_plots(dfs):
    """Create individual trend plots for each mineral, with three scenarios side by side"""
    # Skip Overview and Summary sheets
    mineral_sheets = [sheet for sheet in dfs.keys() 
                     if not (sheet.startswith('Overview') or sheet.startswith('Summary'))]
    
    # Define scenarios
    scenarios = [
        'Stated Policies scenario',
        'Announced Pledges scenario',
        'Net Zero Emissions by 2050 scenario'
    ]
    
    for mineral in mineral_sheets:
        df = dfs[mineral]
        # Skip share rows
        df_filtered = df[~df['Category'].str.contains('share|Share|Total', case=False, na=False)]
        
        # Create subplots - one for each scenario
        fig = make_subplots(rows=1, cols=3,
                           subplot_titles=[s.replace(' scenario', '') for s in scenarios],
                           horizontal_spacing=0.1)
        
        # Define colors for categories
        categories = df_filtered['Category'].unique()
        colors = px.colors.qualitative.Set3[:len(categories)]
        
        # Find max y value across all scenarios for consistent y-axis
        max_y = 0
        for scenario in scenarios:
            scenario_cols = [col for col in df.columns if scenario in col]
            for category in categories:
                category_data = df_filtered[df_filtered['Category'] == category]
                values = category_data[scenario_cols].values[0]
                max_y = max(max_y, max(values))
        
        # Add some padding to max_y (5% extra space)
        max_y = max_y * 1.05
        
        for col_idx, scenario in enumerate(scenarios, 1):
            scenario_cols = [col for col in df.columns if scenario in col]
            years = [int(col.split('_')[-1]) for col in scenario_cols]
            
            for cat_idx, category in enumerate(categories):
                category_data = df_filtered[df_filtered['Category'] == category]
                values = category_data[scenario_cols].values[0]
                
                fig.add_trace(
                    go.Scatter(
                        x=years,
                        y=values,
                        name=category,
                        line=dict(color=colors[cat_idx], width=2),
                        mode='lines+markers',
                        marker=dict(size=8),
                        showlegend=(col_idx == 1),  # Only show legend for first subplot
                        hovertemplate="Year: %{x}<br>Demand: %{y:.1f} kt<extra></extra>"
                    ),
                    row=1, col=col_idx
                )
        
        # Update layout
        fig.update_layout(
            title=dict(
                text=f"{mineral} Demand by Category - All Scenarios",
                x=0.5,
                font=dict(size=24)
            ),
            height=600,
            width=1800,
            template='plotly_white',
            showlegend=True,
            legend=dict(
                yanchor="top",
                y=0.99,
                xanchor="left",
                x=1.02,
                bgcolor='rgba(255,255,255,0.8)'
            ),
            margin=dict(r=250)
        )
        
        # Update all x and y axes with consistent range
        for i in range(1, 4):
            fig.update_xaxes(title_text="Year", row=1, col=i, gridcolor='rgba(0,0,0,0.1)')
            fig.update_yaxes(title_text="Demand (kt)" if i == 1 else None,
                           row=1, col=i, 
                           gridcolor='rgba(0,0,0,0.1)',
                           range=[0, max_y])  # Set consistent y-axis range
        
        fig.write_html(f'figures/{mineral.lower().replace(" ", "_")}_trends.html')

def create_proportion_plots_mpl(dfs):
    """Create proportion plots using matplotlib"""
    # Skip Overview and Summary sheets
    mineral_sheets = [sheet for sheet in dfs.keys() 
                     if not (sheet.startswith('Overview') or sheet.startswith('Summary'))]
    
    # Define scenarios
    scenarios = [
        'Stated Policies scenario',
        'Announced Pledges scenario',
        'Net Zero Emissions by 2050 scenario'
    ]
    
    for mineral in mineral_sheets:
        df = dfs[mineral]
        # Skip share rows
        df_filtered = df[~df['Category'].str.contains('share|Share', case=False, na=False)]
        
        for scenario in scenarios:
            # Get relevant columns for this scenario
            columns = ['Category'] + [col for col in df.columns if scenario in col]
            
            # Get proportions for each year
            proportions_data = []
            for col in columns[1:]:  # Skip Category column
                year_data = df_filtered[['Category', col]].copy()
                total = year_data[col].sum()
                if total > 0:  # Avoid division by zero
                    year_data['Proportion'] = year_data[col] / total
                else:
                    year_data['Proportion'] = 0
                proportions_data.append(year_data['Proportion'].values)
            
            # Create stacked bar chart
            fig, ax = plt.subplots(figsize=(12, 8))
            
            # Get categories for legend
            categories = df_filtered['Category'].unique()
            
            # Create bars
            bars = ax.barh(range(len(columns[1:])), 
                          [1] * len(columns[1:]), 
                          label=categories[0])
            
            left = np.zeros(len(columns[1:]))
            all_bars = []  # Store all bar containers
            
            for i, category in enumerate(categories):
                values = [data[i] for data in proportions_data]
                bars = ax.barh(range(len(columns[1:])), values, 
                             left=left, label=category)
                left += values
                all_bars.append(bars)
            
            # Adding text labels on all bars - only top 3 per year
            for year_idx in range(len(columns[1:])):
                # Get all values for this year
                year_values = []
                for bars in all_bars:
                    bar = bars.patches[year_idx]
                    width = bar.get_width()
                    if width > 0:
                        year_values.append({
                            'width': width,
                            'x_pos': bar.get_x() + width / 2,
                            'y_pos': bar.get_y() + bar.get_height() / 2,
                            'bar': bar
                        })
                
                # Sort by width and get top 3
                top_values = sorted(year_values, key=lambda x: x['width'], reverse=True)[:3]
                
                # Add labels for top 3
                for value in top_values:
                    label = f'{value["width"]:.2%}'  # Format as percentage
                    ax.text(value['x_pos'], value['y_pos'],
                           label,
                           va='center',
                           ha='center',
                           fontsize=9,
                           color='black',
                           fontweight='bold',
                           bbox=dict(facecolor='white',
                                   alpha=0.7,
                                   edgecolor='none',
                                   pad=1))
            
            # Customize plot
            ax.set_title(f'{mineral} - {scenario}\nProportion of Clean Technology Demand by Category', 
                        fontsize=15, pad=20)
            ax.set_xlabel('Proportion of Total Clean Technologies', fontsize=12)
            ax.set_ylabel('Year', fontsize=12)
            
            # Before setting ticklabels, set the ticks
            ax.set_yticks(range(len(columns[1:])))
            ax.set_yticklabels([col.split('_')[-1] for col in columns[1:]])
            
            # Add grid
            ax.grid(True, linestyle='--', alpha=0.7, color='grey')
            
            # Move legend outside
            ax.legend(title='Category', 
                     loc='center left', 
                     bbox_to_anchor=(1.0, 0.5), 
                     fontsize=10)
            
            # Adjust layout to prevent label cutoff
            plt.tight_layout()
            
            # Create safe filename by removing special characters and spaces
            safe_mineral = "".join(c for c in mineral.lower() if c.isalnum() or c == '_')
            safe_scenario = "".join(c for c in scenario.lower() if c.isalnum() or c == '_')
            
            # Save figure with more width for legend
            plt.savefig(f'figures/{safe_mineral}_{safe_scenario}_proportions.png',
                       bbox_inches='tight', 
                       dpi=300,
                       facecolor='white',
                       edgecolor='none')
            plt.close()

def create_statistics_tables(dfs):
    """Create comprehensive statistics tables for each mineral"""
    # Skip Overview and Summary sheets
    mineral_sheets = [sheet for sheet in dfs.keys() 
                     if not (sheet.startswith('Overview') or sheet.startswith('Summary'))]
    
    scenarios = [
        'Stated Policies scenario',
        'Announced Pledges scenario',
        'Net Zero Emissions by 2050 scenario'
    ]
    
    for mineral in mineral_sheets:
        df = dfs[mineral]
        # Skip share rows but keep total
        df_filtered = df[~df['Category'].str.contains('share|Share', case=False, na=False)]
        
        # Prepare statistics data
        stats_data = []
        for category in df_filtered['Category'].unique():
            category_data = {
                'Category': category,
                'Base Value (2023) kt': df_filtered[df_filtered['Category'] == category][f'{scenarios[0]}_2023'].values[0]
            }
            
            # Add 2050 values and growth rates for each scenario
            for scenario in scenarios:
                value_2023 = df_filtered[df_filtered['Category'] == category][f'{scenario}_2023'].values[0]
                value_2050 = df_filtered[df_filtered['Category'] == category][f'{scenario}_2050'].values[0]
                
                scenario_name = scenario.replace(' scenario', '')
                category_data[f'{scenario_name} 2050 (kt)'] = value_2050
                
                if value_2023 != 0:
                    growth = ((value_2050 / value_2023) - 1) * 100
                    category_data[f'{scenario_name} Growth (%)'] = round(growth, 1)
                else:
                    category_data[f'{scenario_name} Growth (%)'] = float('inf') if value_2050 > 0 else 0
                
                # Calculate peak value and year
                scenario_cols = [col for col in df.columns if scenario in col]
                years = [int(col.split('_')[-1]) for col in scenario_cols]
                values = df_filtered[df_filtered['Category'] == category][scenario_cols].values[0]
                max_value = max(values)
                max_year = years[list(values).index(max_value)]
                
                if max_value > value_2050:
                    category_data[f'{scenario_name} Peak'] = f"{round(max_value, 2)} kt ({max_year})"
                else:
                    category_data[f'{scenario_name} Peak'] = "At 2050"
            
            stats_data.append(category_data)
        
        # Create DataFrame
        df_stats = pd.DataFrame(stats_data)
        
        # Create table visualization
        fig = go.Figure(data=[go.Table(
            header=dict(
                values=list(df_stats.columns),
                fill_color='paleturquoise',
                align='left',
                font=dict(size=12, color='black')
            ),
            cells=dict(
                values=[df_stats[col] for col in df_stats.columns],
                fill_color='lavender',
                align='left',
                font=dict(size=11)
            )
        )])
        
        fig.update_layout(
            title=dict(
                text=f"{mineral} - Comprehensive Statistics",
                font=dict(size=16, color='black')
            ),
            width=1500,
            height=max(400, len(df_stats) * 30 + 100),
            margin=dict(t=50, l=20, r=20, b=20)
        )
        
        fig.write_html(f'figures/{mineral.lower().replace(" ", "_")}_statistics.html')

def create_cross_mineral_comparisons(dfs):
    """Create comparison visualizations across all minerals for each scenario"""
    scenarios = [
        'Stated Policies scenario',
        'Announced Pledges scenario',
        'Net Zero Emissions by 2050 scenario'
    ]
    
    # Map full scenario names to short versions
    scenario_map = {
        'Stated Policies scenario': 'stated_policies',
        'Announced Pledges scenario': 'announced_pledges',
        'Net Zero Emissions by 2050 scenario': 'net_zero'
    }
    
    for scenario in scenarios:
        # Prepare data for comparisons
        total_demands = []
        growth_rates = []
        
        for mineral in dfs.keys():
            if mineral.startswith('Overview') or mineral.startswith('Summary'):
                continue
                
            df = dfs[mineral]
            total_row = df[df['Category'] == 'Total clean technologies']
            
            if not total_row.empty:
                value_2023 = total_row[f'{scenario}_2023'].values[0]
                value_2050 = total_row[f'{scenario}_2050'].values[0]
                
                total_demands.append({
                    'Mineral': mineral,
                    'Demand_2023': value_2023,
                    'Demand_2050': value_2050
                })
                
                if value_2023 != 0:
                    growth = ((value_2050 - value_2023) / value_2023) * 100
                else:
                    growth = float('inf') if value_2050 > 0 else 0
                
                growth_rates.append({
                    'Mineral': mineral,
                    'Growth': growth
                })
        
        # Create total demand comparison plot
        fig_demand = go.Figure()
        df_demand = pd.DataFrame(total_demands)
        
        fig_demand.add_trace(go.Bar(
            name='2023',
            x=df_demand['Mineral'],
            y=df_demand['Demand_2023'],
            marker_color='lightblue'
        ))
        
        fig_demand.add_trace(go.Bar(
            name='2050',
            x=df_demand['Mineral'],
            y=df_demand['Demand_2050'],
            marker_color='darkblue'
        ))
        
        fig_demand.update_layout(
            title=f"Total Demand Comparison - {scenario.replace(' scenario', '')}",
            barmode='group',
            yaxis_title="Demand (kt)",
            height=500,
            width=800,
            template='plotly_white'
        )
        
        # Save total demand comparison using mapped scenario name
        scenario_name = scenario_map[scenario]
        fig_demand.write_html(f'figures/total_demand_comparison_{scenario_name}.html')
        
        # Create growth rate comparison plot
        fig_growth = go.Figure()
        df_growth = pd.DataFrame(growth_rates).sort_values('Growth', ascending=True)
        
        fig_growth.add_trace(go.Bar(
            x=df_growth['Mineral'],
            y=df_growth['Growth'],
            marker_color=['red' if x < 0 else 'green' for x in df_growth['Growth']],
            text=[f"{x:.1f}%" for x in df_growth['Growth']],  # Add percentage text
            textposition='auto',  # Automatically position text
            hovertemplate="Mineral: %{x}<br>Growth: %{text}<extra></extra>"  # Custom hover text
        ))
        
        fig_growth.update_layout(
            title=f"Growth Rate Comparison (2023-2050) - {scenario.replace(' scenario', '')}",
            yaxis_title="Growth Rate (%)",
            height=500,
            width=800,
            template='plotly_white',
            yaxis=dict(
                tickformat=',.1f%',  # Format y-axis ticks as percentages
                ticksuffix='%'  # Add % to tick labels
            ),
            uniformtext=dict(
                mode='hide',  # Hide text that doesn't fit
                minsize=8  # Minimum text size
            )
        )
        
        # Save growth rate comparison using mapped scenario name
        fig_growth.write_html(f'figures/growth_comparison_{scenario_name}.html')

def main():
    # Create figures directory if it doesn't exist
    if not os.path.exists('figures'):
        os.makedirs('figures')
    
    # Load and inspect all data
    dfs = load_data()
    
    # Check if we have any data to process
    if not dfs:
        print("\nNo data to process. Please check if the input file contains mineral sheets.")
        return
    
    try:
        # Create individual mineral visualizations
        create_mineral_plots(dfs)
        create_proportion_plots_mpl(dfs)
        create_statistics_tables(dfs)
        create_cross_mineral_comparisons(dfs)
        
        print("\nAnalysis complete! Created interactive visualizations in 'figures' directory:")
        print("1. Individual mineral trend plots (all scenarios) (.html)")
        print("2. Individual mineral proportion plots (.png)")
        print("3. Individual mineral statistics tables (.html)")
        print("4. Cross-mineral comparison plots (.html)")
    
    except Exception as e:
        print(f"\nError during visualization creation: {str(e)}")
        print("Please check if the input file structure matches the expected format.")
        raise

if __name__ == "__main__":
    main() 