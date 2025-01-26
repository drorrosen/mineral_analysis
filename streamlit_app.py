import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import os
import numpy as np
from PIL import Image
import hashlib

# Import visualization modules
from analysis_table_1_visuals import *
from analysis_table_2_visuals import *
#from analysis_table_3_1_visuals import *
from analysis_table_4_1_visuals import *
from analysis_table_4_4_visuals import *
from analysis_table_4_5_visuals import *
from analysis_table_4_6_visuals import *

# Set page config
st.set_page_config(
    page_title="Mineral Demand Analysis",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
st.markdown("""
    <style>
    .main {
        background-color: #f5f5f5;
    }
    .stMarkdown {
        font-family: 'Helvetica Neue', sans-serif;
    }
    .stTitle {
        font-weight: bold;
        color: #2c3e50;
    }
    .stSubheader {
        color: #34495e;
    }
    .stText {
        font-size: 16px;
        line-height: 1.6;
    }
    .sidebar .sidebar-content {
        background-color: #2c3e50;
    }
    </style>
""", unsafe_allow_html=True)

# Initialize session state for login
if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False

def check_password():
    """Returns `True` if the user had the correct password."""
    
    def password_entered():
        """Checks whether a password entered by the user is correct."""
        if st.session_state["username"].strip() == "mineral_analysis" and st.session_state["password"].strip() == "z1x2c3v4b5":
            st.session_state.logged_in = True
            del st.session_state["password"]  # Don't store password
            del st.session_state["username"]  # Don't store username
        else:   
            st.session_state.logged_in = False

    if not st.session_state.logged_in:
        st.title("🔒 Login Required")
        
        # Center the login form
        col1, col2, col3 = st.columns([1,2,1])
        with col2:
            st.text_input("Username", key="username")
            st.text_input("Password", type="password", key="password")
            st.button("Log In", on_click=password_entered)
        return False
    
    return True

# Show content only if logged in
if check_password():
    # Sidebar navigation
    st.sidebar.title("Navigation")
    page = st.sidebar.radio(
        "Select a Page",
        ["Introduction", 
         "Table 1: Total Demand for Key Minerals", 
         "Table 2: Total Supply for Key Minerals",
         "Table 3.2: Cleantech Demand by Mineral",
         "Table 4.1: Solar PV",
         "Table 4.2: Wind",
         "Table 4.4: Grid Battery Storage",
         "Table 4.5: Electricity Networks",
         "Table 4.6: Hydrogen Technologies"]
    )

    if page == "Introduction":
        # Header
        st.title("⚡ Clean Energy Transition Mineral Analysis")
        
        # Introduction text
        st.markdown("""
        ### About This Dashboard
        This interactive dashboard provides comprehensive analysis of mineral demand 
        and supply in the context of clean energy transition. The analysis covers:
        
        ### Tables Analyzed
        - **Table 1**: Total Demand for Key Minerals
        - **Table 2**: Total Supply for Key Minerals
        - **Table 3.2**: Cleantech Demand by Mineral
        - **Table 4.1**: Solar PV Technology Analysis
        - **Table 4.2**: Wind Technology Analysis
        - **Table 4.4**: Grid Battery Storage Analysis
        - **Table 4.5**: Electricity Networks Analysis
        
        ### Key Features
        - Scenario-based analysis (Stated Policies, Announced Pledges, Net Zero)
        - Detailed statistics and trends
        - Interactive visualizations
        - Cross-mineral comparisons
        
        Use the sidebar navigation to explore each table's analysis.
        """)

    elif page == "Table 1: Total Demand for Key Minerals":
        st.title("Mineral Demand Analysis")
        
        # Get list of minerals and add General View
        minerals = ['General View', 'Copper', 'Cobalt', 'Lithium', 'Nickel', 
                   'Magnet rare earth elements', 'Graphite all grades natural and']
        
        # Sidebar filters
        st.sidebar.subheader("Filters")
        selected_mineral = st.sidebar.selectbox(
            "Select Mineral",
            minerals
        )
        
        if selected_mineral == "General View":
            st.subheader("Cross-Mineral Comparison")
            
            # Display comparison for each scenario
            scenarios = ["Stated Policies", "Announced Pledges", "Net Zero"]
            
            for scenario in scenarios:
                st.markdown(f"### {scenario} Scenario")
                
                # Create two columns for different metrics
                col1, col2 = st.columns(2)
                
                with col1:
                    try:
                        # Total demand comparison
                        with open(f'figures_total_demand_comparison_{scenario.lower().replace(" ", "_")}.html', 'r', encoding='utf-8') as f:
                            st.components.v1.html(f.read(), height=500)
                    except FileNotFoundError:
                        st.error(f"Total demand comparison not found for {scenario}")
                
                with col2:
                    try:
                        # Growth rate comparison
                        with open(f'figures_growth_comparison_{scenario.lower().replace(" ", "_")}.html', 'r', encoding='utf-8') as f:
                            st.components.v1.html(f.read(), height=500)
                    except FileNotFoundError:
                        st.error(f"Growth comparison not found for {scenario}")
                
                st.markdown("---")
        
        else:
            # Create safe filename version of mineral name
            safe_mineral = selected_mineral.lower().replace(" ", "_")
            
            # Create special format for proportion files
            proportion_mineral = "".join(c for c in selected_mineral.lower() if c.isalnum() or c == '_')
            
            # Display statistics first
            st.subheader(f"{selected_mineral} - Key Statistics")
            try:
                with open(f'figures_{safe_mineral}_statistics.html', 'r', encoding='utf-8') as f:
                    st.components.v1.html(f.read(), height=400)
            except FileNotFoundError:
                st.error(f"Statistics visualization not found for {selected_mineral}")
            
            # Add some spacing
            st.markdown("---")
            
            # Display trends
            st.subheader(f"{selected_mineral} - Demand Trends by Scenario")
            try:
                with open(f'figures_{safe_mineral}_trends.html', 'r', encoding='utf-8') as f:
                    st.components.v1.html(f.read(), height=700)
            except FileNotFoundError:
                st.error(f"Trends visualization not found for {selected_mineral}")
            
            # Add some spacing
            st.markdown("---")
            
            # Display proportions
            st.subheader(f"{selected_mineral} - Demand Proportions by Scenario")
            
            try:
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    st.image(f'figures/figures_{proportion_mineral}_statedpoliciesscenario_proportions.png', 
                            caption="Stated Policies Scenario", 
                            use_container_width=True)
                
                with col2:
                    st.image(f'figures/figures_{proportion_mineral}_announcedpledgesscenario_proportions.png', 
                            caption="Announced Pledges Scenario", 
                            use_container_width=True)
                
                with col3:
                    st.image(f'figures/figures_{proportion_mineral}_netzeroemissionsby2050scenario_proportions.png', 
                            caption="Net Zero Scenario", 
                            use_container_width=True)
            except FileNotFoundError as e:
                st.error(f"Proportion visualizations not found for {selected_mineral}. Error: {str(e)}")
            
            # Add download section at the bottom
            st.markdown("---")
            st.subheader("Download Visualizations")
            
            col1, col2 = st.columns(2)
            
            with col1:
                try:
                    with open(f'figures_{safe_mineral}_trends.html', 'r', encoding='utf-8') as f:
                        st.download_button(
                            label="Download Trends Analysis",
                            data=f.read(),
                            file_name=f"{safe_mineral}_trends.html",
                            mime="text/html"
                        )
                except FileNotFoundError:
                    st.error("Trends file not available for download")
            
            with col2:
                try:
                    with open(f'figures_{safe_mineral}_statistics.html', 'r', encoding='utf-8') as f:
                        st.download_button(
                            label="Download Statistics",
                            data=f.read(),
                            file_name=f"{safe_mineral}_statistics.html",
                            mime="text/html"
                        )
                except FileNotFoundError:
                    st.error("Statistics file not available for download")

    elif page == "Table 2: Total Supply for Key Minerals":
        st.title("Production Analysis")
        
        # Get list of minerals and add General Views
        minerals = ['General View - Mining', 'General View - Refining', 
                   'Copper', 'Cobalt', 'Lithium', 'Nickel', 
                   'Magnet rare earth elements', 'Graphite']
        
        # Sidebar filters
        st.sidebar.subheader("Filters")
        selected_mineral = st.sidebar.selectbox(
            "Select Mineral/View",
            minerals
        )
        
        if "General View" in selected_mineral:
            activity = "mining" if "Mining" in selected_mineral else "refining"
            st.subheader(f"Cross-Mineral {activity.title()} Analysis")
            
            # Display summary statistics table first
            try:
                with open(f'figure_2_summary_statistics_{activity}.html', 'r', encoding='utf-8') as f:
                    st.components.v1.html(f.read(), height=400, scrolling=True)
            except FileNotFoundError:
                st.error(f"Summary statistics not found for {activity}")
            
            st.markdown("\n\n")
            
            # Display country dominance 2023
            st.subheader("Country Distribution 2023")
            try:
                with open(f'figure_2_country_dominance_{activity}_2023.html', 'r', encoding='utf-8') as f:
                    st.components.v1.html(f.read(), height=800, scrolling=True)
            except FileNotFoundError:
                st.error(f"2023 dominance visualization not found for {activity}")
            
            st.markdown("\n\n")
            
            # Display country dominance 2040
            st.subheader("Country Distribution 2040")
            try:
                with open(f'figure_2_country_dominance_{activity}_2040.html', 'r', encoding='utf-8') as f:
                    st.components.v1.html(f.read(), height=800, scrolling=True)
            except FileNotFoundError:
                st.error(f"2040 dominance visualization not found for {activity}")
        
        else:
            # Create safe filename version of mineral name
            safe_mineral = selected_mineral.lower().replace(" ", "_")
            
            # Display statistics tables first
            st.subheader("Production Statistics")
            
            # Mining statistics
            try:
                with open(f'figure_2_mining_statistics.html', 'r', encoding='utf-8') as f:
                    st.components.v1.html(f.read(), height=600, scrolling=True)
            except FileNotFoundError:
                st.error(f"Mining statistics not found for {selected_mineral}")
            
            st.markdown("\n\n")
            
            # Refining statistics
            try:
                with open(f'figure_2_refining_statistics.html', 'r', encoding='utf-8') as f:
                    st.components.v1.html(f.read(), height=600, scrolling=True)
            except FileNotFoundError:
                st.error(f"Refining statistics not found for {selected_mineral}")
            
            st.markdown("\n\n")
            
            # Display mining trend
            st.subheader("Mining Production Trend")
            try:
                with open(f'figure_2_mining_trend.html', 'r', encoding='utf-8') as f:
                    st.components.v1.html(f.read(), height=800, scrolling=True)
            except FileNotFoundError:
                st.error(f"Mining trend visualization not found for {selected_mineral}")
            
            st.markdown("\n\n")
            
            # Display refining trend
            st.subheader("Refining Production Trend")
            try:
                with open(f'figure_2_refining_trend.html', 'r', encoding='utf-8') as f:
                    st.components.v1.html(f.read(), height=800, scrolling=True)
            except FileNotFoundError:
                st.error(f"Refining trend visualization not found for {selected_mineral}")

    elif page == "Table 3.2: Cleantech Demand by Mineral":
        st.title("Cleantech Demand by Mineral")
        
        # Get list of scenarios
        scenarios = ['Stated Policies', 'Announced Pledges', 'Net Zero']
        
        # Sidebar filters
        st.sidebar.subheader("Filters")
        selected_scenario = st.sidebar.selectbox(
            "Select Scenario",
            scenarios
        )
        
        # Add scenario description
        scenario_descriptions = {
            'Stated Policies': "Base scenario reflecting current policies and trends",
            'Announced Pledges': "Scenario based on announced climate commitments",
            'Net Zero': "Scenario targeting net zero emissions by 2050"
        }
        st.info(scenario_descriptions[selected_scenario])
        
        # Create safe filename version
        safe_scenario = selected_scenario.lower().replace(" ", "_")
        
        # Add key insights section
        st.markdown("### Key Insights")
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("**Top Growing Metals:**")
            try:
                df = pd.read_excel('3.2 Cleantech demand by mineral.xlsx', sheet_name=selected_scenario)
                if '2023' in df.columns:
                    base_year = '2023'
                else:
                    base_year = '2030'
                    df.columns = [str(col).replace('.0', '') for col in df.columns]
                
                # Calculate growth rates
                growth_rates = ((df['2050'] / df[base_year]) - 1) * 100
                df_sorted = pd.DataFrame({
                    'Metal': df['Metal'],
                    'Growth': growth_rates
                }).sort_values('Growth', ascending=False)
                
                top_3 = df_sorted.head(3)['Metal'].tolist()
                for metal in top_3:
                    st.write(f"• {metal}")
            except Exception as e:
                st.error(f"Could not load top metals data: {str(e)}")
        
        with col2:
            st.markdown("**Declining Metals:**")
            try:
                # Use the same DataFrame from above
                bottom_3 = df_sorted.tail(3)['Metal'].tolist()
                for metal in bottom_3:
                    st.write(f"• {metal}")
            except Exception as e:
                st.error(f"Could not load declining metals data: {str(e)}")
        
        # Display statistics table
        try:
            with open(f'figure_3_2_statistics_{safe_scenario}.html', 'r', encoding='utf-8') as f:
                st.components.v1.html(f.read(), height=800, scrolling=True)
        except FileNotFoundError:
            st.error(f"Statistics not found for {selected_scenario}")
        
        st.markdown("\n\n")
        
        # Display top metals analysis
        st.subheader("Growing Metals Analysis")
        try:
            with open(f'figure_3_2_top_growing_metals_{safe_scenario}.html', 'r', encoding='utf-8') as f:
                st.components.v1.html(f.read(), height=700, scrolling=True)
        except FileNotFoundError:
            st.error(f"Growing metals analysis not found for {selected_scenario}")

        st.markdown("\n\n")

        st.subheader("Declining Metals Analysis")
        try:
            with open(f'figure_3_2_top_declining_metals_{safe_scenario}.html', 'r', encoding='utf-8') as f:
                st.components.v1.html(f.read(), height=700, scrolling=True)
        except FileNotFoundError:
            st.error(f"Declining metals analysis not found for {selected_scenario}")

    elif page == "Table 4.1: Solar PV":
        st.title("Solar PV Technology Analysis")
        
        # Get list of technologies
        technologies = [
            'Base case', 
            'Comeback of high Cd-Te technology',
            'Wider adoption of Ga-As technology',
            'Wider adoption of perovskite solar cells'
        ]
        
        # Sidebar filters
        st.sidebar.subheader("Filters")
        selected_tech = st.sidebar.selectbox(
            "Select Technology",
            technologies
        )
        
        # Create safe filename version of technology name
        safe_tech = selected_tech.replace(" ", "_")
        
        # Display statistics table
        try:
            with open(f'figure_4_1_{safe_tech}_statistics_table.html', 'r', encoding='utf-8') as f:
                st.components.v1.html(f.read(), height=600, scrolling=True)
        except FileNotFoundError:
            st.error(f"Statistics not found for {selected_tech}")
        
        st.markdown("\n\n")
        
        # Display key findings
        try:
            with open(f'figure_4_1_{safe_tech}_key_findings.html', 'r', encoding='utf-8') as f:
                st.components.v1.html(f.read(), height=400, scrolling=True)
        except FileNotFoundError:
            st.error(f"Key findings not found for {selected_tech}")
        
        st.markdown("\n\n")
        
        # Display aggregate trends
        try:
            with open(f'figure_4_1_{safe_tech}_aggregate_trends.html', 'r', encoding='utf-8') as f:
                st.components.v1.html(f.read(), height=800, scrolling=True)
        except FileNotFoundError:
            st.error(f"Aggregate trends not found for {selected_tech}")
        
        st.markdown("\n\n")
        
        # Display 2050 comparison heatmap
        try:
            with open(f'figure_4_1_{safe_tech}_2050_comparison.html', 'r', encoding='utf-8') as f:
                st.components.v1.html(f.read(), height=600, scrolling=True)
        except FileNotFoundError:
            st.error(f"2050 comparison not found for {selected_tech}")
        
        st.markdown("\n\n")
        
        # Display growth rates
        try:
            with open(f'figure_4_1_{safe_tech}_growth_rates.html', 'r', encoding='utf-8') as f:
                st.components.v1.html(f.read(), height=600, scrolling=True)
        except FileNotFoundError:
            st.error(f"Growth rates not found for {selected_tech}")

    elif page == "Table 4.2: Wind":
        st.title("Wind Technology Analysis")
        
        # Get list of materials and add General View
        materials = ['General View'] + sorted([
            'Boron', 'Chromium', 'Copper', 'Manganese', 'Molybdenum',
            'Nickel', 'Zinc', 'Neodymium', 'Dysprosium', 'Praseodymium',
            'Terbium', 'Total wind'
        ])
        
        # Sidebar filters
        st.sidebar.subheader("Filters")
        selected_material = st.sidebar.selectbox(
            "Select Material",
            materials
        )
        
        if selected_material == "General View":
            # Display supply constraint impact heatmap
            st.subheader("Supply Constraint Impact Analysis")
            try:
                with open(f'figure_4_2_supply_constraint_impact.html', 'r', encoding='utf-8') as f:
                    st.components.v1.html(f.read(), height=800)
            except FileNotFoundError:
                st.error("Supply constraint impact visualization not found")
            
            st.markdown("---")
            
            # Display growth rates
            st.subheader("Growth Rate Analysis")
            try:
                with open(f'figure_4_2_growth_rates.html', 'r', encoding='utf-8') as f:
                    st.components.v1.html(f.read(), height=600)
            except FileNotFoundError:
                st.error("Growth rates visualization not found")
            
            st.markdown("---")
            
            # Display statistical summaries
            col1, col2 = st.columns(2)
            
            with col1:
                st.subheader("Base Case Statistics")
                try:
                    with open(f'figure_4_2_statistics_base_case.html', 'r', encoding='utf-8') as f:
                        st.components.v1.html(f.read(), height=800)
                except FileNotFoundError:
                    st.error("Base case statistics not found")
            
            with col2:
                st.subheader("Constrained Case Statistics")
                try:
                    with open(f'figure_4_2_statistics_constrained.html', 'r', encoding='utf-8') as f:
                        st.components.v1.html(f.read(), height=800)
                except FileNotFoundError:
                    st.error("Constrained case statistics not found")
        
        else:
            # Create safe filename version of material name
            safe_material = selected_material.lower().replace(" ", "_")
            
            # Display base case trends
            st.subheader(f"{selected_material} - Base Case Scenarios")
            try:
                with open(f'figure_4_2_base_case_trends.html', 'r', encoding='utf-8') as f:
                    st.components.v1.html(f.read(), height=600)
            except FileNotFoundError:
                st.error(f"Base case trends not found for {selected_material}")
            
            st.markdown("---")
            
            # Display constrained case trends
            st.subheader(f"{selected_material} - Constrained Supply Scenarios")
            try:
                with open(f'figure_4_2_constrained_case_trends.html', 'r', encoding='utf-8') as f:
                    st.components.v1.html(f.read(), height=600)
            except FileNotFoundError:
                st.error(f"Constrained case trends not found for {selected_material}")
            
            # Add download section
            st.markdown("---")
            st.subheader("Download Visualizations")
            
            col1, col2 = st.columns(2)
            
            with col1:
                try:
                    with open(f'figure_4_2_base_case_trends.html', 'r', encoding='utf-8') as f:
                        st.download_button(
                            label="Download Base Case Analysis",
                            data=f.read(),
                            file_name=f"{safe_material}_base_case_trends.html",
                            mime="text/html"
                        )
                except FileNotFoundError:
                    st.error("Base case file not available for download")
            
            with col2:
                try:
                    with open(f'figure_4_2_constrained_case_trends.html', 'r', encoding='utf-8') as f:
                        st.download_button(
                            label="Download Constrained Case Analysis",
                            data=f.read(),
                            file_name=f"{safe_material}_constrained_case_trends.html",
                            mime="text/html"
                        )
                except FileNotFoundError:
                    st.error("Constrained case file not available for download")

    elif page == "Table 4.4: Grid Battery Storage":
        st.title("Grid Battery Storage Analysis")
        
        # Get list of materials and add General View
        materials = ['General View'] + [
            'Copper', 'Cobalt', 'Battery-grade graphite', 'Lithium',
            'Manganese', 'Nickel', 'Silicon', 'Vanadium'
        ]
        
        # Sidebar filters
        st.sidebar.subheader("Filters")
        selected_material = st.sidebar.selectbox(
            "Select Material",
            materials
        )
        
        if selected_material == "General View":
            # Display total demand trends
            st.subheader("Total Mineral Demand Trends")
            try:
                with open(f'figure_4_4_total_demand_trends.html', 'r', encoding='utf-8') as f:
                    st.components.v1.html(f.read(), height=600)
            except FileNotFoundError:
                st.error("Total demand trends visualization not found")
            
            st.markdown("---")
            
            # Display scenario comparison heatmap
            st.subheader("2050 Scenario Comparison")
            try:
                with open(f'figure_4_4_scenario_comparison_2050.html', 'r', encoding='utf-8') as f:
                    st.components.v1.html(f.read(), height=800)
            except FileNotFoundError:
                st.error("Scenario comparison visualization not found")
            
            st.markdown("---")
            
            # Display growth rates
            st.subheader("Growth Rate Analysis (2023-2050)")
            try:
                with open(f'figure_4_4_growth_rates.html', 'r', encoding='utf-8') as f:
                    st.components.v1.html(f.read(), height=600)
            except FileNotFoundError:
                st.error("Growth rates visualization not found")
            
            # Add key insights section
            st.markdown("---")
            st.subheader("Key Insights")
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("""
                **Highest Demand Materials (2050):**
                - Battery-grade graphite
                - Copper
                - Vanadium
                """)
            
            with col2:
                st.markdown("""
                **Key Trends:**
                - Most materials show significant growth across scenarios
                - Some materials (Cobalt, Manganese, Nickel) show decline after 2035
                - Net Zero scenario shows highest demand across materials
                """)
        
        else:
            # Create safe filename version of material name
            safe_material = selected_material.lower().replace(" ", "_")
            
            # Display individual material trends
            st.subheader(f"{selected_material} Demand Trends")
            try:
                with open(f'figure_4_4_{safe_material}_trends.html', 'r', encoding='utf-8') as f:
                    st.components.v1.html(f.read(), height=600)
            except FileNotFoundError:
                st.error(f"Trends visualization not found for {selected_material}")
            
            st.markdown("---")
            
            # Add material-specific insights
            st.subheader("Material Insights")
            
            # Calculate growth rates and key statistics for the selected material
            df = create_dataframe()
            material_data = df[df['Material'] == selected_material].iloc[0]
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("**2023 Baseline:**")
                st.write(f"{material_data['Base case']:.2f} kt")
                
                st.markdown("**2050 Projections:**")
                st.write(f"- Stated Policies: {material_data['Stated Policies_2050']:.2f} kt")
                st.write(f"- Announced Pledges: {material_data['Announced Pledges_2050']:.2f} kt")
                st.write(f"- Net Zero: {material_data['Net Zero_2050']:.2f} kt")
            
            with col2:
                st.markdown("**Growth Rates (2023-2050):**")
                base = material_data['Base case']
                if base != 0:
                    for scenario in ['Stated Policies', 'Announced Pledges', 'Net Zero']:
                        growth = ((material_data[f'{scenario}_2050'] - base) / base) * 100
                else:
                    st.write("Growth rates not available (zero base value)")
            
            # Add download section
            st.markdown("---")
            st.subheader("Download Visualization")
            
            try:
                with open(f'figure_4_4_{safe_material}_trends.html', 'r', encoding='utf-8') as f:
                    st.download_button(
                        label="Download Trends Analysis",
                        data=f.read(),
                        file_name=f"{safe_material}_trends.html",
                        mime="text/html"
                    )
            except FileNotFoundError:
                st.error("Trends file not available for download")

    elif page == "Table 4.5: Electricity Networks":
        st.title("Electricity Networks Analysis")
        
        # Get list of views
        views = ['Overview', 'Base case', 'Wider direct current (DC) technology development']
        
        # Sidebar filters
        st.sidebar.subheader("Filters")
        selected_view = st.sidebar.selectbox(
            "Select View",
            views
        )
        
        if selected_view == 'Overview':
            # Display technology comparison
            st.subheader("Technology Comparison Across Scenarios")
            try:
                with open(f'figure_4_5_technology_comparison.html', 'r', encoding='utf-8') as f:
                    st.components.v1.html(f.read(), height=600)
            except FileNotFoundError:
                st.error("Technology comparison visualization not found")
            
            st.markdown("---")
            
            # Display scenario comparison heatmap
            st.subheader("2050 Scenario Comparison")
            try:
                with open(f'figure_4_5_scenario_comparison_2050.html', 'r', encoding='utf-8') as f:
                    st.components.v1.html(f.read(), height=600)
            except FileNotFoundError:
                st.error("Scenario comparison visualization not found")
            
            st.markdown("---")
            
            # Display growth rates
            st.subheader("Growth Rate Analysis (2023-2050)")
            try:
                with open(f'figure_4_5_growth_rates.html', 'r', encoding='utf-8') as f:
                    st.components.v1.html(f.read(), height=600)
            except FileNotFoundError:
                st.error("Growth rates visualization not found")
            
            # Add key insights section
            st.markdown("---")
            st.subheader("Key Insights")
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("""
                **Technology Comparison:**
                - Base case shows higher copper demand across scenarios
                - DC technology development could reduce copper demand
                - Both technologies show significant growth from 2023
                """)
            
            with col2:
                st.markdown("""
                **Scenario Impact:**
                - Net Zero scenario shows highest demand
                - Announced Pledges scenario shows moderate demand
                - Stated Policies scenario shows lowest demand growth
                """)
        
        else:
            # Create safe filename version
            safe_view = selected_view.lower().replace(" ", "_").replace("(", "").replace(")", "")
            
            if selected_view == 'Base case':
                filename = 'figure_4_5_base_case_trends.html'
            else:
                filename = 'figure_4_5_wider_direct_current_dc_technology_development_trends.html'
            
            # Display individual technology trends
            st.subheader(f"{selected_view} - Copper Demand Trends")
            try:
                with open(f'figure_4_5_{filename}', 'r', encoding='utf-8') as f:
                    st.components.v1.html(f.read(), height=600)
            except FileNotFoundError:
                st.error(f"Trends visualization not found for {selected_view}")
            
            st.markdown("---")
            
            # Add technology-specific insights
            st.subheader("Technology Insights")
            
            # Calculate growth rates and key statistics
            df = create_dataframe()
            try:
                tech_data = df[df['Technology'] == selected_view].iloc[0]
                
                col1, col2 = st.columns(2)
                
                with col1:
                    st.markdown("**2023 Baseline:**")
                    st.write(f"{tech_data['Base case']:.2f} kt")
                    
                    st.markdown("**2050 Projections:**")
                    st.write(f"- Stated Policies: {tech_data['Stated Policies_2050']:.2f} kt")
                    st.write(f"- Announced Pledges: {tech_data['Announced Pledges_2050']:.2f} kt")
                    st.write(f"- Net Zero: {tech_data['Net Zero_2050']:.2f} kt")
                
                with col2:
                    st.markdown("**Growth Rates (2023-2050):**")
                    base = tech_data['Base case']
                    for scenario in ['Stated Policies', 'Announced Pledges', 'Net Zero']:
                        growth = ((tech_data[f'{scenario}_2050'] - base) / base) * 100
                        st.write(f"- {scenario}: {growth:.1f}%")
            
            except IndexError:
                st.error("Error: Could not find technology data. Please check the selected view.")
            
            # Add download section
            st.markdown("---")
            st.subheader("Download Visualization")
            
            try:
                with open(f'figure_4_5_{filename}', 'r', encoding='utf-8') as f:
                    st.download_button(
                        label="Download Trends Analysis",
                        data=f.read(),
                        file_name=filename,
                        mime="text/html"
                    )
            except FileNotFoundError:
                st.error("Trends file not available for download")

    elif page == "Table 4.6: Hydrogen Technologies":
        st.title("Hydrogen Technologies Analysis")
        
        # Get list of materials and add General View
        materials = ['Overview'] + [
            'Copper', 'Cobalt', 'Iridium', 'Nickel', 
            'PGMs (other than iridium)', 'Zirconium', 'Yttrium',
            'Total hydrogen technologies'
        ]
        
        # Sidebar filters
        st.sidebar.subheader("Filters")
        selected_material = st.sidebar.selectbox(
            "Select Material",
            materials
        )
        
        if selected_material == "Overview":
            # Display material comparison
            st.subheader("Material Comparison Across Scenarios")
            try:
                with open(f'figure_4_6_material_comparison.html', 'r', encoding='utf-8') as f:
                    st.components.v1.html(f.read(), height=600)
            except FileNotFoundError:
                st.error("Material comparison visualization not found")
            
            st.markdown("---")
            
            # Display scenario comparison heatmap
            st.subheader("2050 Scenario Comparison")
            try:
                with open(f'figure_4_6_scenario_comparison_2050.html', 'r', encoding='utf-8') as f:
                    st.components.v1.html(f.read(), height=800)
            except FileNotFoundError:
                st.error("Scenario comparison visualization not found")
            
            st.markdown("---")
            
            # Display growth rates
            st.subheader("Growth Rate Analysis (2023-2050)")
            try:
                with open(f'figure_4_6_growth_rates.html', 'r', encoding='utf-8') as f:
                    st.components.v1.html(f.read(), height=600)
            except FileNotFoundError:
                st.error("Growth rates visualization not found")
            
            # Add key insights section
            st.markdown("---")
            st.subheader("Key Insights")
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("""
                **Key Materials:**
                - Nickel shows highest demand across scenarios
                - Zirconium shows moderate demand growth
                - Most other materials show minimal demand
                """)
            
            with col2:
                st.markdown("""
                **Scenario Impact:**
                - Net Zero scenario shows significantly higher demand
                - Announced Pledges shows moderate increase
                - Stated Policies shows minimal growth
                """)
        
        else:
            # Create safe filename version of material name
            safe_material = selected_material.lower().replace(" ", "_").replace("(", "").replace(")", "")
            filename = f'figure_4_6_{safe_material}_trends.html'
            
            # Display individual material trends
            st.subheader(f"{selected_material} Demand Trends")
            try:
                with open(filename, 'r', encoding='utf-8') as f:
                    st.components.v1.html(f.read(), height=600)
            except FileNotFoundError:
                st.error(f"Trends visualization not found for {selected_material}")
            
            st.markdown("---")
            
            # Add material-specific insights
            st.subheader("Material Insights")
            
            # Calculate growth rates and key statistics
            df = create_dataframe()
            material_data = df[df['Material'] == selected_material].iloc[0]
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("**2023 Baseline:**")
                st.write(f"{material_data['Base case']:.2f} kt")
                
                st.markdown("**2050 Projections:**")
                st.write(f"- Stated Policies: {material_data['Stated Policies_2050']:.2f} kt")
                st.write(f"- Announced Pledges: {material_data['Announced Pledges_2050']:.2f} kt")
                st.write(f"- Net Zero: {material_data['Net Zero_2050']:.2f} kt")
            
            with col2:
                st.markdown("**Growth Rates (2023-2050):**")
                base = material_data['Base case']
                if base != 0:
                    for scenario in ['Stated Policies', 'Announced Pledges', 'Net Zero']:
                        growth = ((material_data[f'{scenario}_2050'] - base) / base) * 100
                        st.write(f"- {scenario}: {growth:.1f}%")
                else:
                    if any(material_data[f'{scenario}_2050'] > 0 for scenario in ['Stated Policies', 'Announced Pledges', 'Net Zero']):
                        st.write("Growth from zero baseline - infinite growth rate")
                    else:
                        st.write("No demand in baseline or projections")
            
            # Add download section
            st.markdown("---")
            st.subheader("Download Visualization")
            
            try:
                with open(filename, 'r', encoding='utf-8') as f:
                    st.download_button(
                        label="Download Trends Analysis",
                        data=f.read(),
                        file_name=f"figure_4_6_{safe_material}_trends.html",
                        mime="text/html"
                    )
            except FileNotFoundError:
                st.error("Trends file not available for download")
