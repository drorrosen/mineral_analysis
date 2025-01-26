import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import numpy as np
import os
from plotly.subplots import make_subplots
import matplotlib.pyplot as plt

# Define consistent color scheme
SCENARIO_COLORS = {
    'Stated Policies': '#1f77b4',      # Blue
    'Announced Pledges': '#ff7f0e',    # Orange
    'Net Zero': '#2ca02c'              # Green
}

# Define consistent material colors (using qualitative colors)
MATERIAL_COLORS = px.colors.qualitative.Set3

def get_material_color_dict(materials):
    """Create consistent color mapping for materials"""
    return {material: MATERIAL_COLORS[i % len(MATERIAL_COLORS)] 
            for i, material in enumerate(materials)}

def load_data():
    """
    Load and inspect data from both mining and refining Excel files
    """
    files = {
        'mining': '2_mineral_supply_mining.xlsx',
        'refining': '2_mineral_supply_refining.xlsx'
    }
    
    all_data = {}
    
    for data_type, filename in files.items():
        print(f"\nReading {data_type} data from {filename}")
        
        try:
            # Read all sheets
            dfs = pd.read_excel(filename, sheet_name=None)
            
            # Print inspection info
            print(f"\nDATA INSPECTION - {data_type.upper()}")
            for sheet_name, df in dfs.items():
                print(f"\nSheet: {sheet_name}")
                print(f"Shape: {df.shape}")
                print("\nColumns:", df.columns.tolist())
                print("\nFirst few rows:")
                print(df.head())
            
            all_data[data_type] = dfs
            
        except FileNotFoundError:
            print(f"\nError: Could not find {filename}")
            print(f"Please run analysis_table_2.py first to generate {data_type} file.")
            raise
    
    return all_data

def create_mining_refining_comparison(data, mineral):
    """Create comparison plots between mining and refining for a specific mineral"""
    mining_df = data['mining'][mineral]
    refining_df = data['refining'][mineral]
    
    years = ['2023', '2030', '2035', '2040']
    
    fig = go.Figure()
    
    # Better color palette for more distinction
    colors = px.colors.qualitative.Dark24
    
    # Add mining data for each country
    for i, country in enumerate(mining_df['Country'].unique()):
        country_data = mining_df[mining_df['Country'] == country]
        values = [country_data[f'Mining_{year}'].values[0] for year in years]
        
        fig.add_trace(go.Scatter(
            x=years,
            y=values,
            name=f"{country} (Mining)",
            mode='lines+markers',
            line=dict(color=colors[i % len(colors)], dash='solid', width=2),
            marker=dict(size=8, line=dict(width=2, color='white')),
            hovertemplate="Year: %{x}<br>Mining: %{y:.1f} kt<extra></extra>"
        ))
    
    # Add refining data for each country
    for i, country in enumerate(refining_df['Country'].unique()):
        country_data = refining_df[refining_df['Country'] == country]
        values = [country_data[f'Refining_{year}'].values[0] for year in years]
        
        fig.add_trace(go.Scatter(
            x=years,
            y=values,
            name=f"{country} (Refining)",
            mode='lines+markers',
            line=dict(color=colors[i % len(colors)], dash='dot', width=2),
            marker=dict(size=8, line=dict(width=2, color='white')),
            hovertemplate="Year: %{x}<br>Refining: %{y:.1f} kt<extra></extra>"
        ))
    
    fig.update_layout(
        title=dict(
            text=f"{mineral} - Mining vs Refining by Country",
            x=0.5,
            font=dict(size=24)
        ),
        xaxis=dict(
            title="Year",
            gridcolor='rgba(0,0,0,0.1)',
            showgrid=True
        ),
        yaxis=dict(
            title="Production (kt)",
            gridcolor='rgba(0,0,0,0.1)',
            showgrid=True
        ),
        plot_bgcolor='white',
        paper_bgcolor='white',
        hovermode='x unified',
        showlegend=True,
        legend=dict(
            yanchor="top",
            y=0.99,
            xanchor="left",
            x=1.05,
            bgcolor='rgba(255,255,255,0.8)'
        ),
        width=1200,
        height=800,
        margin=dict(r=300)
    )
    
    fig.write_html(f'figure_2/{mineral.lower().replace(" ", "_")}_mining_refining_comparison.html')

def create_trend_plots(data, mineral):
    """Create separate trend plots for mining and refining"""
    years = ['2023', '2030', '2035', '2040']
    activities = ['mining', 'refining']
    
    for activity in activities:
        df = data[activity][mineral]
        fig = go.Figure()
        
        # Better color palette for more distinction
        colors = px.colors.qualitative.Dark24
        
        for i, country in enumerate(df['Country'].unique()):
            country_data = df[df['Country'] == country]
            values = [country_data[f'{activity.capitalize()}_{year}'].values[0] for year in years]
            
            fig.add_trace(go.Scatter(
                x=years,
                y=values,
                name=country,
                mode='lines+markers',
                line=dict(color=colors[i % len(colors)], width=2),
                marker=dict(size=8, line=dict(width=2, color='white')),
                hovertemplate=f"Year: %{{x}}<br>{activity.capitalize()}: %{{y:.1f}} kt<extra></extra>"
            ))
        
        fig.update_layout(
            title=dict(
                text=f"{mineral} - {activity.capitalize()} by Country",
                x=0.5,
                font=dict(size=24)
            ),
            xaxis=dict(
                title="Year",
                gridcolor='rgba(0,0,0,0.1)',
                showgrid=True
            ),
            yaxis=dict(
                title=f"{activity.capitalize()} Production (kt)",
                gridcolor='rgba(0,0,0,0.1)',
                showgrid=True
            ),
            plot_bgcolor='white',
            paper_bgcolor='white',
            hovermode='x unified',
            showlegend=True,
            legend=dict(
                yanchor="top",
                y=0.99,
                xanchor="left",
                x=1.05,
                bgcolor='rgba(255,255,255,0.8)'
            ),
            width=1200,
            height=800,
            margin=dict(r=300)
        )
        
        fig.write_html(f'figure_2/{mineral.lower().replace(" ", "_")}_{activity}_trend.html')

def get_country_colors():
    """Create a consistent color mapping for countries using professional colors"""
    # Professional color palette
    colors = [
        '#1f77b4',  # Steel Blue
        '#ff7f0e',  # Dark Orange
        '#2ca02c',  # Forest Green
        '#d62728',  # Crimson
        '#9467bd',  # Purple
        '#8c564b',  # Brown
        '#e377c2',  # Pink
        '#7f7f7f',  # Gray
        '#bcbd22',  # Olive
        '#17becf',  # Cyan
        '#aec7e8',  # Light Blue
        '#ffbb78',  # Light Orange
        '#98df8a',  # Light Green
        '#ff9896',  # Light Red
        '#c5b0d5',  # Light Purple
        '#c49c94',  # Light Brown
        '#f7b6d2',  # Light Pink
        '#c7c7c7',  # Light Gray
        '#dbdb8d',  # Light Olive
        '#9edae5'   # Light Cyan
    ]
    
    # Common countries that appear in the data
    main_countries = [
        'China',
        'Chile', 
        'Democratic Republic of Congo',
        'Peru',
        'Russia',
        'United States',
        'Indonesia',
        'Australia',
        'Japan',
        'Finland',
        'Canada',
        'India',
        'Brazil',
        'Mexico',
        'Argentina',
        'Myanmar',
        'Philippines',
        'New Caledonia',
        'Rest of world',
        'Others'
    ]
    
    # Create color mapping
    return {country: colors[i] for i, country in enumerate(main_countries)}

def create_country_dominance_analysis(data):
    """Create pie charts showing country dominance in mining and refining for 2023 and 2040"""
    activities = ['mining', 'refining']
    minerals = list(data['mining'].keys())
    years = ['2023', '2040']
    
    # Get consistent country colors
    country_colors = get_country_colors()
    
    for activity in activities:
        for year in years:
            # Create main pie chart figure
            fig = make_subplots(
                rows=2, cols=3,
                subplot_titles=minerals,
                specs=[[{'type':'domain'}, {'type':'domain'}, {'type':'domain'}],
                      [{'type':'domain'}, {'type':'domain'}, {'type':'domain'}]]
            )
            
            # Create top 3 shares figure
            fig_top3 = make_subplots(
                rows=2, cols=3,
                subplot_titles=minerals,
                specs=[[{'type':'domain'}, {'type':'domain'}, {'type':'domain'}],
                      [{'type':'domain'}, {'type':'domain'}, {'type':'domain'}]]
            )
            
            row = 1
            col = 1
            
            for mineral in minerals:
                df = data[activity][mineral]
                
                # Get total from the 'Total' row if it exists, otherwise sum all values
                total_row = df[df['Country'] == 'Total']
                total = total_row[f'{activity.capitalize()}_{year}'].values[0] if not total_row.empty else df[f'{activity.capitalize()}_{year}'].sum()
                
                # Filter out 'Total' row but keep 'Rest of world'
                df_filtered = df[df['Country'] != 'Total']
                values = df_filtered[f'{activity.capitalize()}_{year}']
                labels = df_filtered['Country']
                
                # Get colors for these countries
                colors = [country_colors.get(country, '#808080') for country in labels]
                
                # Main pie chart (absolute values)
                fig.add_trace(
                    go.Pie(
                        values=values,
                        labels=labels,
                        name=mineral,
                        title=mineral,
                        marker=dict(colors=colors),
                        hovertemplate="Country: %{label}<br>Production: %{value:.1f} kt<br>Share: %{percent:.1%}<extra></extra>"
                    ),
                    row=row, col=col
                )
                
                # Calculate and sort shares for top 3
                df_shares = pd.DataFrame({
                    'Country': labels,
                    'Value': values,
                    'Share': values / total
                })
                
                # Filter out 'Rest of world' for top 3 calculation
                df_shares_no_row = df_shares[df_shares['Country'] != 'Rest of world']
                top3_data = df_shares_no_row.nlargest(3, 'Share')
                
                # Calculate others (including 'Rest of world')
                others_share = 1 - top3_data['Share'].sum()
                
                # Add "Others" category
                top3_data = pd.concat([
                    top3_data,
                    pd.DataFrame({'Country': ['Others'], 'Share': [others_share]})
                ])
                
                # Get colors for top 3 plus Others
                top3_colors = [country_colors.get(country, '#808080') for country in top3_data['Country']]
                
                # Top 3 shares pie chart
                fig_top3.add_trace(
                    go.Pie(
                        values=top3_data['Share'],
                        labels=top3_data['Country'],
                        name=mineral,
                        title=mineral,
                        marker=dict(colors=top3_colors),
                        hovertemplate="Country: %{label}<br>Share: %{percent:.1%}<extra></extra>"
                    ),
                    row=row, col=col
                )
                
                col += 1
                if col > 3:
                    col = 1
                    row += 1
            
            # Update main figure layout
            fig.update_layout(
                title=dict(
                    text=f"Country Dominance in {activity.capitalize()} ({year})",
                    x=0.5,
                    font=dict(size=24)
                ),
                showlegend=True,
                width=1500,
                height=1000,
                paper_bgcolor='white',
                plot_bgcolor='white'
            )
            
            # Update top 3 figure layout
            fig_top3.update_layout(
                title=dict(
                    text=f"Top 3 Country Shares in {activity.capitalize()} ({year})",
                    x=0.5,
                    font=dict(size=24)
                ),
                showlegend=True,
                width=1500,
                height=1000,
                paper_bgcolor='white',
                plot_bgcolor='white'
            )
            
            # Save both figures
            fig.write_html(f'figure_2/country_dominance_{activity}_{year}.html')
            fig_top3.write_html(f'figure_2/top3_shares_{activity}_{year}.html')

def create_mining_refining_ratio(data, year='2023'):
    """Create analysis of mining to refining ratio by country"""
    minerals = list(data['mining'].keys())
    
    for mineral in minerals:
        mining_df = data['mining'][mineral]
        refining_df = data['refining'][mineral]
        
        # Get all unique countries
        countries = pd.concat([
            mining_df['Country'],
            refining_df['Country']
        ]).unique()
        
        ratio_data = []
        for country in countries:
            mining_val = mining_df[mining_df['Country'] == country][f'Mining_{year}'].values
            mining_value = mining_val[0] if len(mining_val) > 0 else 0
            
            refining_val = refining_df[refining_df['Country'] == country][f'Refining_{year}'].values
            refining_value = refining_val[0] if len(refining_val) > 0 else 0
            
            if mining_value > 0 or refining_value > 0:
                ratio_data.append({
                    'Country': country,
                    'Mining': mining_value,
                    'Refining': refining_value,
                    'Ratio': mining_value/refining_value if refining_value > 0 else np.inf
                })
        
        df = pd.DataFrame(ratio_data)
        
        # Create figure
        fig = go.Figure()
        
        # Add bars for mining and refining
        fig.add_trace(go.Bar(
            name='Mining',
            x=df['Country'],
            y=df['Mining'],
            marker_color='#1f77b4',
            opacity=0.7,
            hovertemplate="Country: %{x}<br>Mining: %{y:.1f} kt<extra></extra>"
        ))
        
        fig.add_trace(go.Bar(
            name='Refining',
            x=df['Country'],
            y=df['Refining'],
            marker_color='#ff7f0e',
            opacity=0.7,
            hovertemplate="Country: %{x}<br>Refining: %{y:.1f} kt<extra></extra>"
        ))
        
        fig.update_layout(
            title=dict(
                text=f"{mineral} - Mining vs Refining by Country ({year})",
                x=0.5,
                font=dict(size=24)
            ),
            xaxis=dict(
                title="Country",
                gridcolor='rgba(0,0,0,0.1)',
                showgrid=True,
                tickangle=45
            ),
            yaxis=dict(
                title="Production (kt)",
                gridcolor='rgba(0,0,0,0.1)',
                showgrid=True
            ),
            plot_bgcolor='white',
            paper_bgcolor='white',
            barmode='group',
            showlegend=True,
            legend=dict(
                yanchor="top",
                y=0.99,
                xanchor="left",
                x=1.05,
                bgcolor='rgba(255,255,255,0.8)'
            ),
            width=1200,
            height=800,
            margin=dict(r=300, b=100)
        )
        
        fig.write_html(f'figure_2/{mineral.lower().replace(" ", "_")}_mining_refining_ratio_{year}.html')

def create_statistics_table(data):
    """Create statistical tables for mining and refining data"""
    for data_type in ['mining', 'refining']:
        for mineral in data[data_type].keys():
            df = data[data_type][mineral]
            years = ['2023', '2030', '2035', '2040']
            
            # Prepare statistics data
            stats_data = []
            for country in df['Country'].unique():
                if country == 'Total':
                    continue
                    
                country_data = {
                    'Country': country,
                    'Base Value (2023) kt': df[df['Country'] == country][f'{data_type.capitalize()}_2023'].values[0],
                    'Final Value (2040) kt': df[df['Country'] == country][f'{data_type.capitalize()}_2040'].values[0],
                }
                
                # Calculate growth rate
                base_value = country_data['Base Value (2023) kt']
                final_value = country_data['Final Value (2040) kt']
                if base_value != 0:
                    growth = ((final_value - base_value) / base_value) * 100
                    country_data['Growth Rate (%)'] = round(growth, 1)
                else:
                    country_data['Growth Rate (%)'] = float('inf') if final_value > 0 else 0
                
                # Find peak value and year
                values = [df[df['Country'] == country][f'{data_type.capitalize()}_{year}'].values[0] for year in years]
                max_value = max(values)
                max_year = years[values.index(max_value)]
                if max_value > final_value:
                    country_data['Peak Production'] = f"{round(max_value, 2)} kt ({max_year})"
                else:
                    country_data['Peak Production'] = "At 2040"
                
                # Calculate market share
                total_2040 = df[df['Country'] == 'Total'][f'{data_type.capitalize()}_2040'].values[0]
                country_data['Market Share 2040 (%)'] = round((final_value / total_2040) * 100, 1)
                
                stats_data.append(country_data)
            
            # Create DataFrame and sort by Final Value
            df_stats = pd.DataFrame(stats_data)
            df_stats = df_stats.sort_values('Final Value (2040) kt', ascending=False)
            
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
                    text=f"{mineral} - {data_type.capitalize()} Statistics by Country",
                    font=dict(size=16, color='black')
                ),
                width=1200,
                height=max(400, len(df_stats) * 30 + 100),
                margin=dict(t=50, l=20, r=20, b=20)
            )
            
            fig.write_html(f'figure_2/{mineral.lower().replace(" ", "_")}_{data_type}_statistics.html')

def create_summary_statistics(data):
    """Create summary statistics tables for mining and refining overview"""
    for activity in ['mining', 'refining']:
        # Collect summary data across all minerals
        summary_data = []
        
        for mineral in data[activity].keys():
            df = data[activity][mineral]
            countries = df[df['Country'] != 'Total']['Country'].unique()
            
            for country in countries:
                country_data = df[df['Country'] == country]
                value_2023 = country_data[f'{activity.capitalize()}_2023'].values[0]
                value_2040 = country_data[f'{activity.capitalize()}_2040'].values[0]
                
                # Calculate growth
                if value_2023 != 0:
                    growth = ((value_2040 - value_2023) / value_2023) * 100
                else:
                    growth = float('inf') if value_2040 > 0 else 0
                
                # Calculate CAGR
                if value_2023 > 0 and value_2040 > 0:
                    cagr = (((value_2040/value_2023)**(1/17)) - 1) * 100  # 17 years from 2023 to 2040
                else:
                    cagr = None
                
                summary_data.append({
                    'Country': country,
                    'Mineral': mineral,
                    'Production 2023 (kt)': value_2023,
                    'Production 2040 (kt)': value_2040,
                    'Total Growth (%)': round(growth, 1),
                    'CAGR (%)': round(cagr, 1) if cagr is not None else None,
                    'Share 2023 (%)': None,  # Will calculate after aggregating
                    'Share 2040 (%)': None   # Will calculate after aggregating
                })
        
        # Convert to DataFrame
        df_summary = pd.DataFrame(summary_data)
        
        # Calculate total production for each year
        total_2023 = df_summary['Production 2023 (kt)'].sum()
        total_2040 = df_summary['Production 2040 (kt)'].sum()
        
        # Calculate shares
        df_summary['Share 2023 (%)'] = (df_summary['Production 2023 (kt)'] / total_2023 * 100).round(1)
        df_summary['Share 2040 (%)'] = (df_summary['Production 2040 (kt)'] / total_2040 * 100).round(1)
        
        # Sort by 2040 production
        df_summary = df_summary.sort_values('Production 2040 (kt)', ascending=False)
        
        # Create table visualization
        fig = go.Figure(data=[go.Table(
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
                format=[
                    None,  # Country
                    None,  # Mineral
                    '.1f',  # Production 2023
                    '.1f',  # Production 2040
                    '.1f',  # Total Growth
                    '.1f',  # CAGR
                    '.1f',  # Share 2023
                    '.1f'   # Share 2040
                ]
            )
        )])
        
        fig.update_layout(
            title=dict(
                text=f"Summary Statistics - {activity.title()} (2023-2040)",
                font=dict(size=16, color='black')
            ),
            width=1200,
            height=max(400, len(df_summary) * 30 + 100),
            margin=dict(t=50, l=20, r=20, b=20)
        )
        
        fig.write_html(f'figure_2/summary_statistics_{activity}.html')

def create_aggregate_trends(data):
    """Create aggregate trend plots combining all countries for each mineral"""
    for data_type in ['mining', 'refining']:
        for mineral in data[data_type].keys():
            df = data[data_type][mineral]
            years = ['2023', '2030', '2035', '2040']
            
            # Get top 5 countries by 2040 production
            top_countries = df[df['Country'] != 'Total'].nlargest(5, f'{data_type.capitalize()}_2040')['Country'].tolist()
            other_countries = [c for c in df['Country'].unique() if c not in top_countries and c != 'Total']
            
            fig = make_subplots(
                rows=2, cols=2,
                subplot_titles=[
                    "Production Trends - Top 5 Countries",
                    "Market Share Evolution",
                    "Year-over-Year Growth Rates",
                    "Regional Distribution 2040"
                ],
                specs=[
                    [{"type": "scatter"}, {"type": "bar"}],
                    [{"type": "bar"}, {"type": "pie"}]
                ],
                vertical_spacing=0.2,
                horizontal_spacing=0.15
            )
            
            # 1. Production Trends
            colors = px.colors.qualitative.Set3
            for i, country in enumerate(top_countries):
                values = [df[df['Country'] == country][f'{data_type.capitalize()}_{year}'].values[0] for year in years]
                fig.add_trace(
                    go.Scatter(
                        x=years,
                        y=values,
                        name=country,
                        mode='lines+markers',
                        line=dict(color=colors[i]),
                        hovertemplate="Year: %{x}<br>Production: %{y:.1f} kt<extra></extra>",
                        legendgroup="group1",
                        showlegend=True,
                        legendgrouptitle_text="Production Trends"
                    ),
                    row=1, col=1
                )
            
            # 2. Market Share Evolution
            total_values = [df[df['Country'] == 'Total'][f'{data_type.capitalize()}_{year}'].values[0] for year in years]
            for i, country in enumerate(top_countries):
                values = [df[df['Country'] == country][f'{data_type.capitalize()}_{year}'].values[0] for year in years]
                shares = [v/t * 100 for v, t in zip(values, total_values)]
                fig.add_trace(
                    go.Bar(
                        x=years,
                        y=shares,
                        name=country,
                        marker_color=colors[i],
                        hovertemplate="Year: %{x}<br>Share: %{y:.1f}%<extra></extra>",
                        legendgroup="group2",
                        showlegend=True,
                        legendgrouptitle_text="Market Share"
                    ),
                    row=1, col=2
                )
            
            # 3. Year-over-Year Growth Rates
            for i, country in enumerate(top_countries):
                values = [df[df['Country'] == country][f'{data_type.capitalize()}_{year}'].values[0] for year in years]
                growth_rates = [(values[i]/values[i-1] - 1) * 100 for i in range(1, len(values))]
                fig.add_trace(
                    go.Bar(
                        x=years[1:],
                        y=growth_rates,
                        name=country,
                        marker_color=colors[i],
                        hovertemplate="Year: %{x}<br>Growth: %{y:.1f}%<extra></extra>",
                        legendgroup="group3",
                        showlegend=True,
                        legendgrouptitle_text="Growth Rates"
                    ),
                    row=2, col=1
                )
            
            # 4. Regional Distribution 2040
            values_2040 = [df[df['Country'] == country][f'{data_type.capitalize()}_2040'].values[0] 
                          for country in top_countries + ['Rest of world']]
            fig.add_trace(
                go.Pie(
                    labels=top_countries + ['Rest of world'],
                    values=values_2040,
                    marker=dict(colors=colors),
                    hovertemplate="Country: %{label}<br>Share: %{percent}<br>Value: %{value:.1f} kt<extra></extra>",
                    legendgroup="group4",
                    showlegend=True,
                    legendgrouptitle_text="Regional Distribution"
                ),
                row=2, col=2
            )
            
            # Update layout
            fig.update_layout(
                title=dict(
                    text=f"{mineral} - {data_type.capitalize()} Comprehensive Analysis",
                    x=0.5,
                    font=dict(size=20)
                ),
                height=1200,
                width=1800,
                template='plotly_white',
                showlegend=True,
                # Update legend layout
                legend=dict(
                    tracegroupgap=30,
                    yanchor="middle",
                    y=0.5,
                    xanchor="right",
                    x=1.15,
                    bgcolor='rgba(255,255,255,0.8)',
                    bordercolor='rgba(0,0,0,0.2)',
                    borderwidth=1
                ),
                barmode='group'
            )
            
            # Update axes labels
            fig.update_xaxes(title_text="Year", row=1, col=1)
            fig.update_yaxes(title_text="Production (kt)", row=1, col=1)
            fig.update_xaxes(title_text="Year", row=1, col=2)
            fig.update_yaxes(title_text="Market Share (%)", row=1, col=2)
            fig.update_xaxes(title_text="Year", row=2, col=1)
            fig.update_yaxes(title_text="Growth Rate (%)", row=2, col=1)
            
            fig.write_html(f'figure_2/{mineral.lower().replace(" ", "_")}_{data_type}_aggregate_trends.html')

def create_proportion_plots(data):
    """Create stacked proportion plots showing country distribution over time"""
    for data_type in ['mining', 'refining']:
        for mineral in data[data_type].keys():
            df = data[data_type][mineral]
            years = ['2023', '2030', '2035', '2040']
            
            # Get all countries except 'Total' and sort by 2040 value
            countries = df[df['Country'] != 'Total'].sort_values(
                f'{data_type.capitalize()}_2040', 
                ascending=False
            )['Country'].tolist()
            
            # Create figure
            fig = go.Figure()
            
            # Calculate proportions for each country
            for country in countries:
                values = []
                for year in years:
                    total = df[df['Country'] == 'Total'][f'{data_type.capitalize()}_{year}'].values[0]
                    value = df[df['Country'] == country][f'{data_type.capitalize()}_{year}'].values[0]
                    values.append(value / total)
                
                fig.add_trace(go.Bar(
                    name=country,
                    x=years,
                    y=values,
                    text=[f'{v:.1%}' for v in values],
                    textposition='inside',
                    hovertemplate="Year: %{x}<br>Country: " + country + "<br>Share: %{y:.1%}<extra></extra>"
                ))
            
            fig.update_layout(
                title=dict(
                    text=f"{mineral} - {data_type.capitalize()} Country Distribution",
                    x=0.5,
                    font=dict(size=20)
                ),
                xaxis=dict(
                    title="Year",
                    gridcolor='rgba(0,0,0,0.1)'
                ),
                yaxis=dict(
                    title="Share of Total Production",
                    gridcolor='rgba(0,0,0,0.1)',
                    tickformat='.0%'
                ),
                barmode='stack',
                plot_bgcolor='white',
                paper_bgcolor='white',
                showlegend=True,
                legend=dict(
                    title="Country",
                    yanchor="top",
                    y=0.99,
                    xanchor="left",
                    x=1.05,
                    bgcolor='rgba(255,255,255,0.8)'
                ),
                width=1200,
                height=800,
                margin=dict(r=300, b=100)
            )
            
            fig.write_html(f'figure_2/{mineral.lower().replace(" ", "_")}_{data_type}_proportions.html')

def create_proportion_plots_mpl(data):
    """Create stacked proportion plots using matplotlib"""
    for data_type in ['mining', 'refining']:
        for mineral in data[data_type].keys():
            df = data[data_type][mineral]
            years = ['2023', '2030', '2035', '2040']
            
            # Get top countries and sort by 2040 value
            countries = df[df['Country'] != 'Total'].sort_values(
                f'{data_type.capitalize()}_2040', 
                ascending=False
            )['Country'].tolist()[:8]  # Limit to top 8 countries for readability
            
            # Create figure
            fig, ax = plt.subplots(figsize=(12, 8))
            
            # Prepare data
            proportions = []
            for country in countries:
                values = []
                for year in years:
                    total = df[df['Country'] == 'Total'][f'{data_type.capitalize()}_{year}'].values[0]
                    value = df[df['Country'] == country][f'{data_type.capitalize()}_{year}'].values[0]
                    values.append(value / total)
                proportions.append(values)
            
            # Create stacked bars
            bottom = np.zeros(len(years))
            for i, country_data in enumerate(proportions):
                bars = ax.bar(years, country_data, bottom=bottom, label=countries[i])
                bottom += country_data
                
                # Add percentage labels
                for j, rect in enumerate(bars):
                    height = rect.get_height()
                    if height > 0.05:  # Only show labels for segments > 5%
                        ax.text(rect.get_x() + rect.get_width()/2.,
                               rect.get_y() + height/2.,
                               f'{height:.0%}',
                               ha='center', va='center',
                               color='white', fontweight='bold')
            
            # Customize plot
            ax.set_title(f'{mineral} - {data_type.capitalize()}\nCountry Distribution', 
                        fontsize=15, pad=20)
            ax.set_xlabel('Year', fontsize=12)
            ax.set_ylabel('Share of Total Production', fontsize=12)
            
            # Format y-axis as percentage
            ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda y, _: '{:.0%}'.format(y)))
            
            # Add grid
            ax.grid(True, linestyle='--', alpha=0.7, color='grey')
            
            # Move legend outside
            ax.legend(title='Country', loc='center left', 
                     bbox_to_anchor=(1.0, 0.5), fontsize=10)
            
            # Adjust layout
            plt.tight_layout()
            
            # Save figure
            plt.savefig(f'figure_2/{mineral.lower().replace(" ", "_")}_{data_type}_proportions.png',
                       bbox_inches='tight', dpi=300)
            plt.close()

def create_detailed_growth_analysis(data):
    """Create detailed growth analysis for each mineral and activity type"""
    growth_data = []
    
    for data_type in ['mining', 'refining']:
        for mineral in data[data_type].keys():
            df = data[data_type][mineral]
            
            for country in df['Country'].unique():
                if country == 'Total':
                    continue
                
                # Get 2023 and 2040 values
                value_2023 = df[df['Country'] == country][f'{data_type.capitalize()}_2023'].values[0]
                value_2040 = df[df['Country'] == country][f'{data_type.capitalize()}_2040'].values[0]
                
                # Calculate growth rate
                if value_2023 != 0:
                    growth = ((value_2040 - value_2023) / value_2023) * 100
                else:
                    growth = float('inf') if value_2040 > 0 else 0
                
                growth_data.append({
                    'Mineral': mineral,
                    'Activity': data_type.capitalize(),
                    'Country': country,
                    'Growth_Rate': growth,
                    'Value_2023': value_2023,
                    'Value_2040': value_2040
                })
    
    growth_df = pd.DataFrame(growth_data)
    
    # Create scatter plot
    fig = go.Figure()
    
    # Different colors for mining and refining
    colors = {'Mining': '#1f77b4', 'Refining': '#ff7f0e'}
    
    for activity in ['Mining', 'Refining']:
        activity_data = growth_df[growth_df['Activity'] == activity].sort_values('Growth_Rate')
        
        fig.add_trace(go.Scatter(
            x=activity_data['Growth_Rate'],
            y=[f"{row['Mineral']} - {row['Country']}" for _, row in activity_data.iterrows()],
            name=activity,
            mode='markers',
            marker=dict(
                size=12,
                color=colors[activity],
                line=dict(width=2, color='white')
            ),
            hovertemplate=(
                "Mineral: %{customdata[0]}<br>" +
                "Country: %{customdata[1]}<br>" +
                "2023: %{customdata[2]:.1f} kt<br>" +
                "2040: %{customdata[3]:.1f} kt<br>" +
                "Growth: %{x:.1f}%<extra></extra>"
            ),
            customdata=activity_data[['Mineral', 'Country', 'Value_2023', 'Value_2040']]
        ))
    
    fig.update_layout(
        title=dict(
            text='Growth Rate (2023-2040) by Mineral, Country and Activity',
            x=0.5,
            font=dict(size=20)
        ),
        xaxis=dict(
            title="Growth Rate (%)",
            gridcolor='rgba(0,0,0,0.1)',
            showgrid=True,
            zeroline=True,
            zerolinecolor='black',
            zerolinewidth=1
        ),
        yaxis=dict(
            title=None,
            gridcolor='rgba(0,0,0,0.1)',
            showgrid=True
        ),
        plot_bgcolor='white',
        paper_bgcolor='white',
        showlegend=True,
        legend=dict(
            title="Activity",
            yanchor="top",
            y=0.99,
            xanchor="left",
            x=1.05,
            bgcolor='rgba(255,255,255,0.8)'
        ),
        width=1200,
        height=max(800, len(growth_df) * 20),
        margin=dict(r=300, l=300)
    )
    
    fig.write_html('figure_2/growth_analysis.html')

def main():
    # Create figure_2 directory if it doesn't exist
    if not os.path.exists('figure_2'):
        os.makedirs('figure_2')
    
    # Load and inspect data
    data = load_data()
    
    # Create visualizations
    for mineral in data['mining'].keys():
        create_trend_plots(data, mineral)
        create_mining_refining_ratio(data, '2023')
    
    create_country_dominance_analysis(data)
    
    # Add new visualization calls
    create_statistics_table(data)
    create_summary_statistics(data)
    create_aggregate_trends(data)
    create_proportion_plots(data)
    create_proportion_plots_mpl(data)
    create_detailed_growth_analysis(data)
    
    print("\nAnalysis complete! Created visualizations in 'figure_2' directory:")
    print("1. Separate mining and refining trends for each mineral")
    print("2. Country dominance pie charts for 2023 and 2040")
    print("3. Mining to Refining ratio analysis by country")
    print("4. Statistical tables")
    print("5. Aggregate trend analysis")
    print("6. Country proportion analysis (interactive HTML and static PNG)")
    print("7. Detailed growth analysis")

if __name__ == "__main__":
    main()
