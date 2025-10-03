#!/usr/bin/env python3
"""
Compare original and verified mention counts
"""

import pandas as pd
import os

def compare_mentions(year):
    """Compare original vs verified mentions for a year"""
    
    print(f"=== COMPARISON FOR {year} ===")
    
    # File paths
    original_file = f'lvmh_{year}FY_maison_mentions.xlsx' if year != '2025H1' else f'lvmh_{year}_maison_mentions.xlsx'
    verified_file = f'lvmh_{year}FY_maison_mentions_verified.xlsx' if year != '2025H1' else f'lvmh_{year}_maison_mentions_verified.xlsx'
    
    if not os.path.exists(original_file) or not os.path.exists(verified_file):
        print("Files not found")
        return
    
    # Load data
    original_df = pd.read_excel(original_file)
    verified_df = pd.read_excel(verified_file)
    
    # Standardize Bulgari spelling
    original_df['Maison'] = original_df['Maison'].str.replace('Bvlgari', 'Bulgari', regex=False)
    verified_df['Maison'] = verified_df['Maison'].str.replace('Bvlgari', 'Bulgari', regex=False)
    
    categories = ['Product', 'Place', 'Partnership', 'ESG', 'Performance', 
                  'Digital', 'Pricing', 'Promotion', 'People', 'Awards']
    
    # Merge dataframes
    comparison = original_df[['Maison'] + categories].merge(
        verified_df[['Maison'] + categories], 
        on='Maison', 
        suffixes=('_original', '_verified')
    )
    
    # Calculate differences
    differences = []
    for _, row in comparison.iterrows():
        maison = row['Maison']
        for category in categories:
            orig_count = row[f'{category}_original']
            verif_count = row[f'{category}_verified']
            if orig_count != verif_count:
                differences.append({
                    'Maison': maison,
                    'Category': category,
                    'Original': orig_count,
                    'Verified': verif_count,
                    'Difference': verif_count - orig_count
                })
    
    if differences:
        diff_df = pd.DataFrame(differences)
        print(f"Found {len(differences)} differences:")
        print(diff_df.to_string(index=False))
        
        # Summary statistics
        print(f"\nSummary:")
        print(f"Total differences: {len(differences)}")
        print(f"Maisons affected: {diff_df['Maison'].nunique()}")
        print(f"Categories affected: {diff_df['Category'].nunique()}")
        
        # Show net changes
        net_changes = diff_df.groupby('Category')['Difference'].sum().sort_values(ascending=False)
        print(f"\nNet changes by category:")
        for category, change in net_changes.items():
            print(f"  {category}: {change:+d}")
    else:
        print("No differences found - data is consistent!")
    
    print()

def main():
    """Compare all years"""
    
    years = ['2022', '2023', '2024', '2025H1']
    
    print("=== ORIGINAL vs VERIFIED MENTION COUNTS COMPARISON ===")
    print()
    
    for year in years:
        compare_mentions(year)

if __name__ == "__main__":
    main()
