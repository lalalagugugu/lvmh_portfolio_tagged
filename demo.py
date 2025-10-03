#!/usr/bin/env python3
"""
Demo script to showcase key features of the LVMH Portfolio Analysis Dashboard
"""

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import os

def load_data():
    """Load all Excel files"""
    data = {}
    
    for year in ['2023', '2024']:
        mentions_file = f'lvmh_{year}FY_maison_mentions.xlsx'
        details_file = f'lvmh_{year}FY_maison_details.xlsx'
        
        if os.path.exists(mentions_file):
            data[f'mentions_{year}'] = pd.read_excel(mentions_file)
        if os.path.exists(details_file):
            data[f'details_{year}'] = pd.read_excel(details_file)
    
    return data

def calculate_total_mentions(df):
    """Calculate total mentions for each Maison"""
    categories = ['Product', 'Place', 'Partnership', 'ESG', 'Performance', 
                  'Digital', 'Pricing', 'Promotion', 'People', 'Awards']
    df['Total_Mentions'] = df[categories].sum(axis=1)
    return df

def demo_kpis():
    """Demonstrate KPI calculations"""
    print("🍾 LVMH Portfolio Analysis - Key Performance Indicators")
    print("=" * 60)
    
    data = load_data()
    
    for year in ['2023', '2024']:
        year_key = f'mentions_{year}'
        if year_key not in data:
            continue
            
        print(f"\n📊 {year} Fiscal Year Analysis:")
        print("-" * 40)
        
        df = calculate_total_mentions(data[year_key].copy())
        df = df.sort_values('Total_Mentions', ascending=False)
        
        # Most mentioned Maison
        most_mentioned = df.iloc[0]
        print(f"🏆 Most Mentioned Maison: {most_mentioned['Maison']} ({most_mentioned['Total_Mentions']} mentions)")
        
        # Top 5 Maisons
        print(f"\n📈 Top 5 Maisons:")
        for i, (_, row) in enumerate(df.head(5).iterrows(), 1):
            print(f"  {i}. {row['Maison']}: {row['Total_Mentions']} mentions")
        
        # Category leaders
        categories = ['Product', 'Place', 'Partnership', 'ESG', 'Performance', 
                      'Digital', 'Pricing', 'Promotion', 'People', 'Awards']
        
        print(f"\n🎯 Category Leaders:")
        for category in categories:
            leader = df.loc[df[category].idxmax()]
            if leader[category] > 0:
                print(f"  {category}: {leader['Maison']} ({leader[category]} mentions)")

def demo_evolution():
    """Demonstrate year-over-year evolution"""
    print("\n\n📈 Year-over-Year Evolution Analysis")
    print("=" * 60)
    
    data = load_data()
    
    if 'mentions_2023' not in data or 'mentions_2024' not in data:
        print("❌ Both 2023 and 2024 data required for evolution analysis")
        return
    
    df_2023 = calculate_total_mentions(data['mentions_2023'].copy())
    df_2024 = calculate_total_mentions(data['mentions_2024'].copy())
    
    # Merge data for comparison
    comparison = df_2023[['Maison', 'Total_Mentions']].merge(
        df_2024[['Maison', 'Total_Mentions']], 
        on='Maison', 
        suffixes=('_2023', '_2024')
    )
    
    # Calculate change
    comparison['Change'] = comparison['Total_Mentions_2024'] - comparison['Total_Mentions_2023']
    comparison['Change_Pct'] = (comparison['Change'] / comparison['Total_Mentions_2023'] * 100).round(1)
    
    # Top performers
    top_improvers = comparison.nlargest(5, 'Change')
    print("\n🚀 Top 5 Improvers (2023 → 2024):")
    for _, row in top_improvers.iterrows():
        if row['Change'] > 0:
            print(f"  {row['Maison']}: {row['Total_Mentions_2023']} → {row['Total_Mentions_2024']} (+{row['Change']}, +{row['Change_Pct']}%)")
    
    # Biggest decliners
    top_decliners = comparison.nsmallest(5, 'Change')
    print("\n📉 Biggest Decliners (2023 → 2024):")
    for _, row in top_decliners.iterrows():
        if row['Change'] < 0:
            print(f"  {row['Maison']}: {row['Total_Mentions_2023']} → {row['Total_Mentions_2024']} ({row['Change']}, {row['Change_Pct']}%)")

def demo_category_analysis():
    """Demonstrate category-specific analysis"""
    print("\n\n🎯 Category-Specific Analysis")
    print("=" * 60)
    
    data = load_data()
    
    if 'mentions_2024' not in data:
        print("❌ 2024 data required for category analysis")
        return
    
    df = data['mentions_2024']
    categories = ['Product', 'Place', 'Partnership', 'ESG', 'Performance', 
                  'Digital', 'Pricing', 'Promotion', 'People', 'Awards']
    
    print("\n📊 Category Distribution (2024):")
    for category in categories:
        total_mentions = df[category].sum()
        active_maisons = (df[category] > 0).sum()
        avg_mentions = (df[category].sum() / len(df)).round(1)
        
        print(f"  {category}:")
        print(f"    Total mentions: {total_mentions}")
        print(f"    Active Maisons: {active_maisons}/{len(df)}")
        print(f"    Average per Maison: {avg_mentions}")
        
        # Top performer in category
        if total_mentions > 0:
            top_maison = df.loc[df[category].idxmax()]
            print(f"    Top performer: {top_maison['Maison']} ({top_maison[category]} mentions)")
        print()

def main():
    """Run all demo functions"""
    print("🍾 LVMH Portfolio Analysis Dashboard - Demo")
    print("=" * 60)
    print("This demo showcases the key features of the web application.")
    print("The actual dashboard provides interactive filtering and visualizations.")
    print()
    
    try:
        demo_kpis()
        demo_evolution()
        demo_category_analysis()
        
        print("\n" + "=" * 60)
        print("✅ Demo completed successfully!")
        print("\n🌐 To access the interactive dashboard:")
        print("   1. Run: streamlit run app.py")
        print("   2. Open: http://localhost:8501")
        print("\n📋 Features available in the web dashboard:")
        print("   • Interactive KPI dashboard")
        print("   • Stacked bar charts for top 10 Maisons")
        print("   • Individual Maison detailed analysis")
        print("   • Filtering by year, Maison, and category")
        print("   • Scalable for additional years of data")
        
    except Exception as e:
        print(f"❌ Error during demo: {str(e)}")

if __name__ == "__main__":
    main()
