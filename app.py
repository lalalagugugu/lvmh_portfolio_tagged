import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
import os

# Page configuration
st.set_page_config(
    page_title="LVMH Portfolio Analysis",
    page_icon="🍾",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better styling
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f4e79;
        text-align: center;
        margin-bottom: 2rem;
    }
    .kpi-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 1rem;
        border-radius: 10px;
        color: white;
        text-align: center;
        margin: 0.5rem 0;
    }
    .kpi-value {
        font-size: 2rem;
        font-weight: bold;
    }
    .kpi-label {
        font-size: 0.9rem;
        opacity: 0.9;
    }
    .category-box {
        border: 2px solid #e0e0e0;
        border-radius: 8px;
        padding: 1rem;
        margin: 0.5rem 0;
        background-color: #f8f9fa;
    }
    .category-title {
        font-weight: bold;
        color: #1f4e79;
        margin-bottom: 0.5rem;
    }
</style>
""", unsafe_allow_html=True)

@st.cache_data
def load_data():
    """Load and process all Excel files"""
    data = {}
    
    # Load mentions data for all available years (prioritize verified files)
    years = ['2019', '2020', '2021', '2022', '2023', '2024', '2025H1']
    for year in years:
        # Try verified mentions first, fall back to original
        # Priority order: standardized verified > standardized > verified > original
        standardized_verified_file = f'lvmh_{year}FY_maison_mentions_standardized_verified.xlsx' if year != '2025H1' else f'lvmh_{year}_maison_mentions_standardized_verified.xlsx'
        standardized_file = f'lvmh_{year}FY_maison_mentions_standardized.xlsx' if year != '2025H1' else f'lvmh_{year}_maison_mentions_standardized.xlsx'
        verified_mentions_file = f'lvmh_{year}FY_maison_mentions_verified.xlsx' if year != '2025H1' else f'lvmh_{year}_maison_mentions_verified.xlsx'
        original_mentions_file = f'lvmh_{year}FY_maison_mentions.xlsx' if year != '2025H1' else f'lvmh_{year}_maison_mentions.xlsx'
        details_file = f'lvmh_{year}FY_maison_details_standardized.xlsx' if year != '2025H1' else f'lvmh_{year}_maison_details_standardized.xlsx'
        
        # Load mentions data (prefer standardized verified)
        mentions_file = None
        if os.path.exists(standardized_verified_file):
            mentions_file = standardized_verified_file
        elif os.path.exists(standardized_file):
            mentions_file = standardized_file
        elif os.path.exists(verified_mentions_file):
            mentions_file = verified_mentions_file
        else:
            mentions_file = original_mentions_file
        if os.path.exists(mentions_file):
            df = pd.read_excel(mentions_file)
            # Data should already be standardized from source
            data[f'mentions_{year}'] = df
        if os.path.exists(details_file):
            df = pd.read_excel(details_file)
            # Data should already be standardized from source  
            data[f'details_{year}'] = df
    
    return data

def calculate_total_mentions(df):
    """Calculate total mentions for each Maison"""
    categories = ['Product', 'Place', 'Partnership', 'ESG', 'Performance', 
                  'Digital', 'Pricing', 'Promotion', 'People', 'Awards']
    df['Total_Mentions'] = df[categories].sum(axis=1)
    return df

def create_cross_year_ranking(data):
    """Create ranking table with mentions and rankings for all years"""
    years = ['2019', '2020', '2021', '2022', '2023', '2024', '2025H1']
    
    # Get all unique Maisons across all years
    all_maisons = set()
    for year in years:
        year_key = f'mentions_{year}'
        if year_key in data:
            year_df = data[year_key].copy()
            all_maisons.update(year_df['Maison'].unique())
    
    # Prepare cross-year data
    cross_year_data = []
    
    for maison in sorted(all_maisons):
        row_data = {'Maison': maison}
        
        for year in years:
            year_key = f'mentions_{year}'
            if year_key in data:
                year_df = data[year_key].copy()
                year_df['Maison'] = year_df['Maison'].str.replace('Bvlgari', 'Bulgari', regex=False)
                maison_data = year_df[year_df['Maison'] == maison]
                
                if not maison_data.empty:
                    total_mentions = maison_data.iloc[0]['Total_Mentions']
                    row_data[f'{year}_mentions'] = total_mentions
                else:
                    row_data[f'{year}_mentions'] = 0
        
        # Calculate rankings for each year
        for year in years:
            year_key = f'mentions_{year}'
            if year_key in data:
                year_df = data[year_key].copy()
                year_df['Maison'] = year_df['Maison'].str.replace('Bvlgari', 'Bulgari', regex=False)
                year_df['Rank'] = year_df['Total_Mentions'].rank(method='min', ascending=False).astype(int)
                
                maison_rank = year_df[year_df['Maison'] == maison]
                if not maison_rank.empty:
                    row_data[f'{year}_rank'] = maison_rank.iloc[0]['Rank']
                else:
                    row_data[f'{year}_rank'] = None
            
        cross_year_data.append(row_data)
    
    return pd.DataFrame(cross_year_data)

def create_total_mentions_chart(data):
    """Create stacked bar chart showing total mentions by category across all years"""
    years = ['2019', '2020', '2021', '2022', '2023', '2024', '2025H1']
    categories = ['Product', 'Place', 'Partnership', 'ESG', 'Performance', 
                  'Digital', 'Pricing', 'Promotion', 'People', 'Awards']
    
    chart_data = []
    for year in years:
        year_key = f'mentions_{year}'
        if year_key in data:
            year_df = data[year_key]
            for category in categories:
                category_total = year_df[category].sum()
                chart_data.append({
                    'Year': year,
                    'Category': category,
                    'Mentions': category_total
                })
    
    if chart_data:
        df_chart = pd.DataFrame(chart_data)
        
        fig = px.bar(
            df_chart,
            x='Year',
            y='Mentions',
            color='Category',
            title='Total Mentions by Year and Category',
            color_discrete_sequence=px.colors.qualitative.Set3
        )
        
        # Ensure x-axis treats years as categories
        fig.update_xaxes(type='category')
        
        fig.update_layout(
            height=400,
            showlegend=True,
            xaxis_tickangle=-45
        )
        
        return fig
    return None

def get_previous_year(selected_year):
    """Get the previous year for comparison"""
    year_mapping = {
        '2025H1': '2024',
        '2024': '2023',
        '2023': '2022',
        '2022': '2021',
        '2021': '2020',
        '2020': '2019',
        '2019': None
    }
    return year_mapping.get(selected_year)

def get_kpis(data, selected_year):
    """Calculate KPIs for the selected year"""
    year_key = f'mentions_{selected_year}'
    prev_year = get_previous_year(selected_year)
    prev_year_key = f'mentions_{prev_year}' if prev_year else None
    
    if year_key not in data:
        return None, None, None, None
    
    current_df = calculate_total_mentions(data[year_key].copy())
    current_df = current_df.sort_values('Total_Mentions', ascending=False)
    
    # Most mentioned Maison
    most_mentioned = current_df.iloc[0]
    
    # Most mentioned from previous year
    prev_most_mentioned = None
    if prev_year_key and prev_year_key in data:
        prev_df = calculate_total_mentions(data[prev_year_key].copy())
        prev_most_mentioned = prev_df.sort_values('Total_Mentions', ascending=False).iloc[0]
    
    # Most mentioned in each category with tie handling
    categories = ['Product', 'Place', 'Partnership', 'ESG', 'Performance', 
                  'Digital', 'Pricing', 'Promotion', 'People', 'Awards']
    category_leaders = {}
    prev_category_leaders = {}
    
    for category in categories:
        # Current year category leaders (handle ties)
        max_mentions = current_df[category].max()
        if max_mentions > 1:  # Only show winners if max > 1
            leaders = current_df[current_df[category] == max_mentions]
            category_leaders[category] = leaders
        else:
            category_leaders[category] = pd.DataFrame()  # Empty for no winners
        
        # Previous year category leaders
        if prev_year_key and prev_year_key in data:
            prev_df = data[prev_year_key]
            prev_max_mentions = prev_df[category].max()
            if prev_max_mentions > 1:
                prev_leaders = prev_df[prev_df[category] == prev_max_mentions]
                prev_category_leaders[category] = prev_leaders
            else:
                prev_category_leaders[category] = pd.DataFrame()
    
    return most_mentioned, prev_most_mentioned, category_leaders, prev_category_leaders

def create_stacked_bar_chart(data, selected_year, top_n=10):
    """Create stacked bar chart for top N Maisons clustered by Maison with hierarchical x-axis"""
    year_key = f'mentions_{selected_year}'
    prev_year = get_previous_year(selected_year)
    prev_year_key = f'mentions_{prev_year}' if prev_year else None
    
    if year_key not in data:
        return None
    
    current_df = calculate_total_mentions(data[year_key].copy())
    top_maisons = current_df.nlargest(top_n, 'Total_Mentions')['Maison'].tolist()
    
    categories = ['Product', 'Place', 'Partnership', 'ESG', 'Performance', 
                  'Digital', 'Pricing', 'Promotion', 'People', 'Awards']
    
    # Prepare data for both years with Maison-Year combination
    chart_data = []
    
    # Create ordered list for x-axis: Maison (prev_year), Maison (selected_year) pairs
    ordered_labels = []
    maison_groups = []
    
    for maison in top_maisons:
        maison_bars = []
        
        # Add previous year first, then current year for each Maison
        if prev_year_key and prev_year_key in data and maison in data[prev_year_key]['Maison'].values:
            ordered_labels.append(f"{maison} ({prev_year})")
            maison_bars.append(f"{maison} ({prev_year})")
            prev_row = data[prev_year_key][data[prev_year_key]['Maison'] == maison].iloc[0]
            for category in categories:
                chart_data.append({
                    'Maison_Year': f"{maison} ({prev_year})",
                    'Maison': maison,
                    'Category': category,
                    'Mentions': prev_row[category],
                    'Year': prev_year
                })
        
        # Add current year
        if maison in current_df['Maison'].values:
            ordered_labels.append(f"{maison} ({selected_year})")
            maison_bars.append(f"{maison} ({selected_year})")
            row = current_df[current_df['Maison'] == maison].iloc[0]
            for category in categories:
                chart_data.append({
                    'Maison_Year': f"{maison} ({selected_year})",
                    'Maison': maison,
                    'Category': category,
                    'Mentions': row[category],
                    'Year': selected_year
                })
        
        if maison_bars:
            maison_groups.append({'name': maison, 'bars': maison_bars})
    
    df_chart = pd.DataFrame(chart_data)
    
    if df_chart.empty:
        return None
    
    # Create stacked bar chart with bars clustered by Maison
    fig = px.bar(
        df_chart, 
        x='Maison_Year', 
        y='Mentions', 
        color='Category',
        title=f'Top {top_n} Maisons: Category Evolution ({prev_year} vs {selected_year})',
        color_discrete_sequence=px.colors.qualitative.Set3
    )
    
    # Set the x-axis order to show Maison pairs (prev_year, selected_year)
    fig.update_xaxes(categoryorder='array', categoryarray=ordered_labels)
    
    # Create hierarchical x-axis with custom spacing
    fig.update_layout(
        height=600,
        showlegend=True,
        xaxis_tickangle=-45,
        xaxis=dict(
            tickmode='array',
            tickvals=list(range(len(ordered_labels))),
            ticktext=ordered_labels,
            # Add spacing between Maison groups
            tick0=0,
            dtick=1,
            # Custom spacing - this will be handled by the data structure
        ),
        # Add annotations for Maison group labels
        annotations=[]
    )
    
    # Add Maison group labels above the chart
    group_annotations = []
    current_pos = 0
    
    for group in maison_groups:
        group_size = len(group['bars'])
        group_center = current_pos + (group_size - 1) / 2
        
        group_annotations.append({
            'x': group_center,
            'y': 1.02,
            'xref': 'x',
            'yref': 'paper',
            'text': group['name'],
            'showarrow': False,
            'font': {'size': 14, 'color': 'black'},
            'xanchor': 'center'
        })
        
        current_pos += group_size
    
    fig.update_layout(annotations=group_annotations)
    
    # Adjust bar spacing for better visual grouping
    fig.update_traces(
        width=0.8,  # Make bars slightly narrower
        marker_line_width=0.5,
        marker_line_color='white'
    )
    
    return fig

def display_category_view(data, selected_category, selected_year, unique_maisons):
    """Display category-specific view with total counts by year and activities table"""
    st.header(f"📊 {selected_category} Activities Analysis")
    
    # Get years data
    years = ['2019', '2020', '2021', '2022', '2023', '2024', '2025H1']
    
    # 1. Total Counts by Year Chart
    st.subheader(f"📈 Total {selected_category} Mentions by Year")
    
    year_totals = []
    for year in years:
        year_key = f'mentions_{year}'
        if year_key in data:
            year_df = data[year_key]
            if selected_category in year_df.columns:
                total = year_df[selected_category].sum()
                year_totals.append({'Year': year, 'Total': total})
    
    if year_totals:
        df_totals = pd.DataFrame(year_totals)
        fig_totals = px.bar(
            df_totals,
            x='Year',
            y='Total',
            title=f'Total {selected_category} Mentions by Year',
            color='Total',
            color_continuous_scale='viridis'
        )
        fig_totals.update_xaxes(type='category')
        fig_totals.update_layout(height=400)
        st.plotly_chart(fig_totals, use_container_width=True)
    
    # 2. Activities by Year by Maison Table
    st.subheader(f"📋 {selected_category} Activities by Year by Maison")
    
    # Collect activities data
    activities_data = {}
    
    for year in years:
        year_key = f'details_{year}'
        if year_key in data:
            details_df = data[year_key]
            
            # Find category columns for this year
            category_columns = [col for col in details_df.columns if col.startswith(selected_category)]
            
            if category_columns:
                for _, row in details_df.iterrows():
                    maison = row['Maison']
                    activities = []
                    
                    for col in category_columns:
                        if col in row.index and pd.notna(row[col]) and str(row[col]).strip():
                            activities.append(str(row[col]).strip())
                    
                    if activities:
                        if maison not in activities_data:
                            activities_data[maison] = {}
                        activities_data[maison][year] = activities
    
    # Create table data
    if activities_data:
        # Get all years that have data
        all_years_with_data = set()
        for maison_data in activities_data.values():
            all_years_with_data.update(maison_data.keys())
        all_years_with_data = sorted(list(all_years_with_data))
        
        # Create table
        table_data = []
        for maison in sorted(activities_data.keys()):
            row_data = {'Maison': maison}
            for year in all_years_with_data:
                if year in activities_data[maison]:
                    # Join multiple activities with line breaks
                    activities = activities_data[maison][year]
                    numbered_activities = [f"{i+1}. {activity}" for i, activity in enumerate(activities)]
                    row_data[year] = "<br>".join(numbered_activities)
                else:
                    row_data[year] = ""
            table_data.append(row_data)
        
        if table_data:
            df_table = pd.DataFrame(table_data)
            
            # Display as HTML table for better formatting
            html_table = "<table style='width:100%; border-collapse: collapse;'>"
            html_table += "<thead><tr style='background-color: #f0f2f6;'>"
            html_table += f"<th style='border: 1px solid #ddd; padding: 8px;'>Maison</th>"
            for year in all_years_with_data:
                html_table += f"<th style='border: 1px solid #ddd; padding: 8px; text-align: center;'>{year}</th>"
            html_table += "</tr></thead><tbody>"
            
            for _, row in df_table.iterrows():
                html_table += "<tr>"
                html_table += f"<td style='border: 1px solid #ddd; padding: 8px; font-weight: bold;'>{row['Maison']}</td>"
                for year in all_years_with_data:
                    cell_content = row[year] if row[year] != "" else ""
                    html_table += f"<td style='border: 1px solid #ddd; padding: 8px; vertical-align: top;'>{cell_content}</td>"
                html_table += "</tr>"
            
            html_table += "</tbody></table>"
            st.markdown(html_table, unsafe_allow_html=True)
        else:
            st.info(f"No {selected_category} activities found for any Maison in the available data.")
    else:
        st.info(f"No {selected_category} activities found in the data.")

def display_maison_details(data, selected_maison, selected_year):
    """Display detailed view for a specific Maison with multi-year data"""
    st.subheader(f"📊 {selected_maison} - Multi-Year Analysis")
    
    # Get available years for this Maison
    available_years = []
    for year in ['2019', '2020', '2021', '2022', '2023', '2024', '2025H1']:
        year_key = f'mentions_{year}'
        if year_key in data:
            mentions_df = data[year_key]
            if selected_maison in mentions_df['Maison'].values:
                available_years.append(year)
    
    if not available_years:
        st.error(f"No data found for {selected_maison}")
        return
    
    categories = ['Product', 'Place', 'Partnership', 'ESG', 'Performance', 
                  'Digital', 'Pricing', 'Promotion', 'People', 'Awards']
    
    # Prepare multi-year data for chart
    chart_data = []
    for year in available_years:
        year_key = f'mentions_{year}'
        mentions_df = data[year_key]
        maison_data = mentions_df[mentions_df['Maison'] == selected_maison]
        
        if not maison_data.empty:
            maison_row = maison_data.iloc[0]
            for category in categories:
                chart_data.append({
                    'Year': year,
                    'Category': category,
                    'Mentions': maison_row[category]
                })
    
    if chart_data:
        df_chart = pd.DataFrame(chart_data)
        
        # Create stacked bar chart showing evolution with categorical x-axis
        fig = px.bar(
            df_chart,
            x='Year',
            y='Mentions',
            color='Category',
            title=f"{selected_maison} - Category Evolution Over Time",
            color_discrete_sequence=px.colors.qualitative.Set3
        )
        
        # Ensure x-axis treats years as categories
        fig.update_xaxes(type='category')
        fig.update_layout(
            height=500,
            showlegend=True,
            xaxis_tickangle=-45
        )
        
        st.plotly_chart(fig, use_container_width=True)
    
    # Create activity table by year and category
    st.subheader("📝 Activities by Year and Category")
    
    # Prepare activity data
    activity_data = {}
    
    for year in available_years:
        details_key = f'details_{year}'
        if details_key in data:
            details_df = data[details_key]
            maison_details = details_df[details_df['Maison'] == selected_maison]
            
            if not maison_details.empty:
                detail_row = maison_details.iloc[0]
                
                # Group details by category
                detail_categories = {
                    'Product': ['Product_1', 'Product_2', 'Product_3', 'Product_4', 'Product_5'],
                    'Place': ['Place_1', 'Place_2', 'Place_3', 'Place_4', 'Place_5'],
                    'Partnership': ['Partnership_1', 'Partnership_2', 'Partnership_3'],
                    'ESG': ['ESG_1', 'ESG_2'],
                    'Performance': ['Performance_1'],
                    'Digital': ['Digital_1', 'Digital_2'],
                    'Pricing': ['Pricing_1'],
                    'Promotion': ['Promotion_1', 'Promotion_2', 'Promotion_3', 'Promotion_4'],
                    'People': ['People_1'],
                    'Awards': ['Awards_1']
                }
                
                for category, columns in detail_categories.items():
                    activities = []
                    for col in columns:
                        if col in detail_row.index and pd.notna(detail_row[col]) and str(detail_row[col]).strip():
                            activities.append(str(detail_row[col]).strip())
                    
                    if activities:
                        if category not in activity_data:
                            activity_data[category] = {}
                        activity_data[category][year] = activities
    
    # Create table
    if activity_data:
        # Create DataFrame for the table
        table_data = []
        for category in categories:
            row = {'Category': category}
            for year in available_years:
                if category in activity_data and year in activity_data[category]:
                    activities = activity_data[category][year]
                    # Number the activities
                    numbered_activities = [f"{i+1}. {activity}" for i, activity in enumerate(activities)]
                    row[year] = " | ".join(numbered_activities)
                else:
                    row[year] = ""
            table_data.append(row)
        
        df_table = pd.DataFrame(table_data)
        st.dataframe(df_table, use_container_width=True, hide_index=True)
    else:
        st.info("No detailed activity data available for this Maison.")

def main():
    # Header
    st.markdown('<h1 class="main-header">🍾 LVMH Portfolio Analysis Dashboard</h1>', unsafe_allow_html=True)
    
    # Load data
    data = load_data()
    
    # Check if verified data is being used
    verified_years = []
    years = ['2019', '2020', '2021', '2022', '2023', '2024', '2025H1']
    for year in years:
        verified_file = f'lvmh_{year}FY_maison_mentions_verified.xlsx' if year != '2025H1' else f'lvmh_{year}_maison_mentions_verified.xlsx'
        if os.path.exists(verified_file) and f'mentions_{year}' in data:
            verified_years.append(year)
    
    if verified_years:
        st.success(f"✅ Using verified mention counts for: {', '.join(verified_years)}")
    else:
        st.info("ℹ️ Using original mention counts")
    
    # Sidebar filters
    st.sidebar.header("🔍 Filters")
    
    # Available years
    available_years = []
    for key in data.keys():
        if key.startswith('mentions_'):
            year = key.split('_')[1]
            available_years.append(year)
    
    if not available_years:
        st.error("No data files found. Please ensure Excel files are in the correct format.")
        return
    
    selected_year = st.sidebar.selectbox("Fiscal Year", available_years, index=len(available_years)-1)
    
    # Available Maisons (unique across all years) - using standardized data
    unique_maisons = set()
    for year in ['2019', '2020', '2021', '2022', '2023', '2024', '2025H1']:
        year_key = f'mentions_{year}'
        if year_key in data:
            year_df = data[year_key].copy()
            # Names should already be standardized in the new files, but apply minimal cleanup if needed
            year_df['Maison'] = year_df['Maison'].str.strip()
            unique_maisons.update(year_df['Maison'].unique())
    
    unique_maisons = sorted(list(unique_maisons))
    selected_maison_mode = st.sidebar.radio("View Mode", ["Overview", "Single Maison", "Multiple Maisons"])
    
    if selected_maison_mode == "Single Maison":
        selected_maison = st.sidebar.selectbox("Select Maison", unique_maisons)
        selected_maisons = []  # Initialize empty list
    elif selected_maison_mode == "Multiple Maisons":
        selected_maisons = st.sidebar.multiselect("Select Maisons", unique_maisons, default=[unique_maisons[0]] if unique_maisons else [])
        selected_maison = "All"  # Initialize to avoid error
    else:
        selected_maison = "All"
        selected_maisons = []  # Initialize empty list
    
    # Available categories
    categories = ['Product', 'Place', 'Partnership', 'ESG', 'Performance', 
                  'Digital', 'Pricing', 'Promotion', 'People', 'Awards']
    selected_category = st.sidebar.selectbox("Activity Category", ["All"] + categories)
    
    # Main content
    if selected_category != "All":
        # Category-specific view
        display_category_view(data, selected_category, selected_year, unique_maisons)
    elif selected_maison_mode == "Overview" or (selected_maison_mode != "Multiple Maisons" and selected_maison == "All"):
        # KPI Dashboard
        st.header("📊 Key Performance Indicators")
        
        most_mentioned, prev_most_mentioned, category_leaders, prev_category_leaders = get_kpis(data, selected_year)
        
        if most_mentioned is not None:
            # KPI Cards with merged information
            col1, col2, col3 = st.columns(3)
            
            with col1:
                # Most Mentioned Maison with previous year comparison
                prev_text = ""
                if prev_most_mentioned is not None:
                    prev_text = f"<div style='font-size: 0.7rem; opacity: 0.8; margin-top: 0.3rem;'>Previous: {prev_most_mentioned['Maison']} ({prev_most_mentioned['Total_Mentions']})</div>"
                
                st.markdown(f"""
                <div class="kpi-card">
                    <div class="kpi-value">{most_mentioned['Total_Mentions']}</div>
                    <div class="kpi-label">Most Mentioned Maison</div>
                    <div style="font-size: 0.8rem; margin-top: 0.5rem;">{most_mentioned['Maison']}</div>
                    {prev_text}
                </div>
                """, unsafe_allow_html=True)
            
            with col2:
                # Total Mentions with previous year comparison
                total_mentions = data[year_key]['Total_Mentions'].sum()
                
                prev_total_mentions = ""
                if prev_most_mentioned is not None and get_previous_year(selected_year):
                    prev_year_key = f'mentions_{get_previous_year(selected_year)}'
                    if prev_year_key in data:
                        prev_total_mentions = data[prev_year_key]['Total_Mentions'].sum()
                        prev_text = f"<div style='font-size: 0.7rem; opacity: 0.8; margin-top: 0.3rem;'>Previous: {prev_total_mentions}</div>"
                    else:
                        prev_text = ""
                else:
                    prev_text = ""
                
                st.markdown(f"""
                <div class="kpi-card">
                    <div class="kpi-value">{total_mentions}</div>
                    <div class="kpi-label">Total Mentions</div>
                    <div style="font-size: 0.8rem; margin-top: 0.5rem;">All Categories</div>
                    {prev_text}
                </div>
                """, unsafe_allow_html=True)
            
            with col3:
                # Total Maisons
                total_maisons = len(data[year_key])
                
                prev_maisons = ""
                if prev_most_mentioned is not None and get_previous_year(selected_year):
                    prev_year_key = f'mentions_{get_previous_year(selected_year)}'
                    if prev_year_key in data:
                        prev_maisons = len(data[prev_year_key])
                        prev_text = f"<div style='font-size: 0.7rem; opacity: 0.8; margin-top: 0.3rem;'>Previous: {prev_maisons}</div>"
                    else:
                        prev_text = ""
                else:
                    prev_text = ""
                
                st.markdown(f"""
                <div class="kpi-card">
                    <div class="kpi-value">{total_maisons}</div>
                    <div class="kpi-label">Total Maisons</div>
                    <div style="font-size: 0.8rem; margin-top: 0.5rem;">Analyzed</div>
                    {prev_text}
                </div>
                """, unsafe_allow_html=True)
            
            # Total Mentions by Year chart
            st.subheader("📈 Total Mentions by Year")
            
            total_chart = create_total_mentions_chart(data)
            if total_chart:
                st.plotly_chart(total_chart, use_container_width=True)
            
            # Category Leaders Table - only show categories with winners
            st.subheader("🏆 Category Leaders")
            
            leaders_data = []
            for category, leaders_df in category_leaders.items():
                if not leaders_df.empty:
                    # Handle ties - combine all tied winners into one cell
                    tied_maisons = leaders_df['Maison'].tolist()
                    mentions_count = leaders_df[category].iloc[0]  # All tied leaders have same count
                    total_mentions = leaders_df['Total_Mentions'].iloc[0]  # Use first one's total
                    
                    leaders_data.append({
                        'Category': category,
                        'Maison': ', '.join(tied_maisons),
                        'Mentions': mentions_count,
                        'Total Mentions': total_mentions
                    })
                # Skip categories with no winners (max mentions <= 1)
            
            if leaders_data:
                leaders_df = pd.DataFrame(leaders_data)
                leaders_df = leaders_df.sort_values(['Mentions', 'Category'], ascending=[False, True])
                
                # Add previous year comparison if available
                if prev_category_leaders:
                    prev_data = []
                    for category, prev_leaders_df in prev_category_leaders.items():
                        if not prev_leaders_df.empty:
                            tied_maisons = prev_leaders_df['Maison'].tolist()
                            mentions_count = prev_leaders_df[category].iloc[0]
                            prev_data.append({
                                'Category': category,
                                'Previous Maison': ', '.join(tied_maisons),
                                'Previous Mentions': mentions_count
                            })
                    
                    if prev_data:
                        prev_df = pd.DataFrame(prev_data)
                        leaders_df = leaders_df.merge(prev_df, on='Category', how='left')
                
                st.dataframe(leaders_df, use_container_width=True)
            
            # Stacked Bar Chart
            st.subheader("📈 Top 10 Maisons: Category Evolution")
            
            fig = create_stacked_bar_chart(data, selected_year, 10)
            if fig:
                st.plotly_chart(fig, use_container_width=True)
            
            # Cross-year ranking table
            st.subheader("📊 Cross-Year Maison Ranking")
            
            cross_year_df = create_cross_year_ranking(data)
            
            # Show ranking explanation
            st.info("**Ranking**: Standard ranking where tied Maisons share the same rank and subsequent ranks are skipped.")
            
            st.dataframe(cross_year_df, use_container_width=True)
        
    elif selected_maison_mode == "Single Maison":
        # Individual Maison Details
        display_maison_details(data, selected_maison, selected_year)
    
    elif selected_maison_mode == "Multiple Maisons":
        # Multiple Maisons comparison
        if selected_maisons:
            st.subheader("📊 Multiple Maisons Comparison")
            
            # Create separate stacked bar charts for each Maison
            years = ['2019', '2020', '2021', '2022', '2023', '2024', '2025H1']
            categories = ['Product', 'Place', 'Partnership', 'ESG', 'Performance', 
                          'Digital', 'Pricing', 'Promotion', 'People', 'Awards']
            
            # Layout: Side-by-side charts for first 2 Maisons
            if len(selected_maisons) >= 1:
                col1, col2 = st.columns(2)
                
                chart_maisons = selected_maisons[:2]  # Only show first 2 for charts
                
                for i, maison in enumerate(chart_maisons):
                    with col1 if i == 0 else col2:
                        st.subheader(f"📊 {maison}")
                        
                        # Prepare chart data for this Maison
                        chart_data = []
                        
                        for year in years:
                            year_key = f'mentions_{year}'
                            if year_key in data:
                                maison_data = data[year_key][data[year_key]['Maison'] == maison]
                                if not maison_data.empty:
                                    maison_row = maison_data.iloc[0]
                                    for category in categories:
                                        chart_data.append({
                                            'Year': year,
                                            'Category': category,
                                            'Mentions': maison_row[category]
                                        })
                        
                        if chart_data:
                            df_chart = pd.DataFrame(chart_data)
                            
                            # Create stacked bar chart for this Maison
                            fig = px.bar(
                                df_chart,
                                x='Year',
                                y='Mentions',
                                color='Category',
                                title=f'{maison}: Multi-Year Trend',
                                color_discrete_sequence=px.colors.qualitative.Set3
                            )
                            
                            # Ensure categorical x-axis
                            fig.update_xaxes(type='category')
                            fig.update_layout(
                                height=400,
                                showlegend=True,
                                xaxis_tickangle=-45
                            )
                            
                            st.plotly_chart(fig, use_container_width=True)
            
            # Detailed activities table for selected fiscal year
            st.subheader(f"📋 Detailed Activities ({selected_year})")
            
            # Create comparison table for selected Maisons
            year_details_key = f'details_{selected_year}'
            if year_details_key in data:
                # Prepare comparison table data
                comparison_data = []
                
                for category in categories:
                    category_row = {'Category': category}
                    
                    for maison in selected_maisons:
                        maison_details_data = data[year_details_key][data[year_details_key]['Maison'] == maison]
                        
                        if not maison_details_data.empty:
                            details_row = maison_details_data.iloc[0]
                            activities_in_category = []
                            
                            # Collect activities for this category and maison
                            for col in details_row.index:
                                if col.startswith(f'{category}_') and pd.notna(details_row[col]) and str(details_row[col]).strip():
                                    activity_text = str(details_row[col]).strip()
                                    if activity_text != '0' and activity_text.lower() != 'nan':
                                        activities_in_category.append(activity_text)
                            
                            # Format activities with numbering
                            if activities_in_category:
                                numbered_activities = []
                                for idx, activity in enumerate(activities_in_category, 1):
                                    numbered_activities.append(f"{idx}. {activity}")
                                category_row[maison] = '\\n'.join(numbered_activities)
                            else:
                                category_row[maison] = 'No activities'
                        else:
                            category_row[maison] = 'No data'
                    
                    comparison_data.append(category_row)
                
                if comparison_data:
                    comparison_df = pd.DataFrame(comparison_data)
                    comparison_df = comparison_df.set_index('Category')
                    
                    # Display with numbered activities visible as separate lines
                    st.markdown("**Activity Descriptions (numbered)**")
                    st.dataframe(comparison_df, use_container_width=True)
                    
                    # Note about numbered entries
                    st.info("💡 **Note**: Multiple activities in the same category are numbered and displayed on separate lines within the same cell.")
                else:
                    st.info("No activity data found.")
            else:
                st.info(f"No details data found for {selected_year}.")
        else:
            st.info("Please select at least one Maison for comparison.")
    
    # Footer
    st.markdown("---")
    st.markdown("""
    <div style="text-align: center; color: #666; font-size: 0.9rem;">
        LVMH Portfolio Analysis Dashboard | Data from {selected_year} Financial Reports
    </div>
    """.format(selected_year=selected_year), unsafe_allow_html=True)

if __name__ == "__main__":
    main()
