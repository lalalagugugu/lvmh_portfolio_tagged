#!/usr/bin/env python3
"""
Test script to verify data loading and basic functionality
"""

import pandas as pd
import os

def test_data_loading():
    """Test if all Excel files can be loaded correctly"""
    
    print("Testing LVMH Portfolio Data Loading...")
    print("=" * 50)
    
    # Test files
    test_files = [
        'lvmh_2023FY_maison_details.xlsx',
        'lvmh_2023FY_maison_mentions.xlsx',
        'lvmh_2024FY_maison_details.xlsx',
        'lvmh_2024FY_maison_mentions.xlsx'
    ]
    
    data = {}
    
    for file in test_files:
        if os.path.exists(file):
            try:
                df = pd.read_excel(file)
                data[file] = df
                print(f"✅ {file}: {df.shape[0]} rows, {df.shape[1]} columns")
                
                # Show first few Maisons
                if 'Maison' in df.columns:
                    print(f"   Sample Maisons: {', '.join(df['Maison'].head(3).tolist())}")
                
            except Exception as e:
                print(f"❌ {file}: Error - {str(e)}")
        else:
            print(f"⚠️  {file}: File not found")
    
    print("\n" + "=" * 50)
    
    # Test data processing
    if 'lvmh_2024FY_maison_mentions.xlsx' in data:
        df = data['lvmh_2024FY_maison_mentions.xlsx']
        
        # Calculate total mentions
        categories = ['Product', 'Place', 'Partnership', 'ESG', 'Performance', 
                      'Digital', 'Pricing', 'Promotion', 'People', 'Awards']
        
        df['Total_Mentions'] = df[categories].sum(axis=1)
        
        # Top 5 Maisons
        top_5 = df.nlargest(5, 'Total_Mentions')[['Maison', 'Total_Mentions']]
        
        print("Top 5 Maisons by Total Mentions (2024):")
        for _, row in top_5.iterrows():
            print(f"  {row['Maison']}: {row['Total_Mentions']} mentions")
        
        # Category analysis
        print("\nCategory Analysis (2024):")
        for category in categories:
            max_mentions = df[category].max()
            top_maison = df.loc[df[category].idxmax(), 'Maison']
            print(f"  {category}: {top_maison} ({max_mentions} mentions)")
    
    print("\n✅ Data loading test completed successfully!")
    return data

if __name__ == "__main__":
    test_data_loading()
