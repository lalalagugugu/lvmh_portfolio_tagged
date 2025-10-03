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
    years = ['2022', '2023', '2024', '2025H1']
    for year in years:
        # Try verified mentions first, fall back to original
        verified_mentions_file = f'lvmh_{year}FY_maison_mentions_verified.xlsx' if year != '2025H1' else f'lvmh_{year}_maison_mentions_verified.xlsx'
        original_mentions_file = f'lvmh_{year}FY_maison_mentions.xlsx' if year != '2025H1' else f'lvmh_{year}_maison_mentions.xlsx'
        details_file = f'lvmh_{year}FY_maison_details.xlsx' if year != '2025H1' else f'lvmh_{year}_maison_details.xlsx'
        
        # Load mentions data (prefer verified)
        mentions_file = verified_mentions_file if os.path.exists(verified_mentions_file) else original_mentions_file
        if os.path.exists(mentions_file):
            df = pd.read_excel(mentions_file)
            # Standardize Bulgari spelling
            df['Maison'] = df['Maison'].str.replace('Bvlgari', 'Bulgari', regex=False)
            data[f'mentions_{year}'] = df
        if os.path.exists(details_file):
            df = pd.read_excel(details_file)
            # Standardize Bulgari spelling
            df['Maison'] = df['Maison'].str.replace('Bvlgari', 'Bulgari', regex=False)
            data[f'details_{year}'] = df
    
    return data

def calculate_total_mentions(df):
    """Calculate total mentions for each Maison"""
    categories = ['Product', 'Place', 'Partnership', 'ESG', 'Performance', 
                  'Digital', 'Pricing', 'Promotion', 'People', 'Awards']
    df['Total_Mentions'] = df[categories].sum(axis=1)
    return df

def get_previous_year(selected_year):
    """Get the previous year for comparison"""
    year_mapping = {
        '2025H1': '2024',
        '2024': '2023',
        '2023': '2022',
        '2022': None
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

def display_maison_details(data, selected_maison, selected_year):
    """Display detailed view for a specific Maison with multi-year data"""
    st.subheader(f"📊 {selected_maison} - Multi-Year Analysis")
    
    # Get available years for this Maison
    available_years = []
    for year in ['2022', '2023', '2024', '2025H1']:
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
                        if pd.notna(detail_row[col]) and str(detail_row[col]).strip():
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
    years = ['2022', '2023', '2024', '2025H1']
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
    
    # Available Maisons
    year_key = f'mentions_{selected_year}'
    if year_key in data:
        available_maisons = sorted(data[year_key]['Maison'].unique())
        selected_maison = st.sidebar.selectbox("Maison (for detailed view)", ["All"] + available_maisons)
    else:
        selected_maison = "All"
    
    # Available categories
    categories = ['Product', 'Place', 'Partnership', 'ESG', 'Performance', 
                  'Digital', 'Pricing', 'Promotion', 'People', 'Awards']
    selected_category = st.sidebar.selectbox("Activity Category", ["All"] + categories)
    
    # Main content
    if selected_maison == "All":
        # KPI Dashboard
        st.header("📊 Key Performance Indicators")
        
        most_mentioned, prev_most_mentioned, category_leaders, prev_category_leaders = get_kpis(data, selected_year)
        
        if most_mentioned is not None:
            # Simplified KPI Cards
            col1, col2 = st.columns(2)
            
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
                # Total Maisons with previous year comparison
                total_maisons = len(data[year_key])
                prev_total = ""
                if prev_most_mentioned is not None and get_previous_year(selected_year):
                    prev_year_key = f'mentions_{get_previous_year(selected_year)}'
                    if prev_year_key in data:
                        prev_total = len(data[prev_year_key])
                        prev_text = f"<div style='font-size: 0.7rem; opacity: 0.8; margin-top: 0.3rem;'>Previous: {prev_total}</div>"
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
            
            # Overall ranking
            st.subheader("📊 Overall Maison Ranking")
            
            current_df = calculate_total_mentions(data[year_key].copy())
            ranking_df = current_df[['Maison', 'Total_Mentions'] + categories].sort_values('Total_Mentions', ascending=False)
            
            st.dataframe(ranking_df, use_container_width=True)
        
    else:
        # Individual Maison Details
        display_maison_details(data, selected_maison, selected_year)
    
    # Footer
    st.markdown("---")
    st.markdown("""
    <div style="text-align: center; color: #666; font-size: 0.9rem;">
        LVMH Portfolio Analysis Dashboard | Data from {selected_year} Financial Reports
    </div>
    """.format(selected_year=selected_year), unsafe_allow_html=True)

if __name__ == "__main__":
    main()
