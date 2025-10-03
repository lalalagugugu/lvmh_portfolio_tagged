#!/usr/bin/env python3
"""
Generate verified mention counts from details workbooks
"""

import pandas as pd
import os

def count_mentions_in_details(details_row, category):
    """Count non-null entries for a specific category in details row"""
    
    # Define column mappings for each category
    category_columns = {
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
    
    if category not in category_columns:
        return 0
    
    columns = category_columns[category]
    count = 0
    
    for col in columns:
        if col in details_row.index:
            value = details_row[col]
            if pd.notna(value) and str(value).strip():
                count += 1
    
    return count

def generate_verified_mentions(year):
    """Generate verified mentions workbook for a specific year"""
    
    # File paths
    details_file = f'lvmh_{year}FY_maison_details.xlsx' if year != '2025H1' else f'lvmh_{year}_maison_details.xlsx'
    output_file = f'lvmh_{year}FY_maison_mentions_verified.xlsx' if year != '2025H1' else f'lvmh_{year}_maison_mentions_verified.xlsx'
    
    if not os.path.exists(details_file):
        print(f"Details file not found: {details_file}")
        return
    
    print(f"Processing {year}...")
    
    # Load details data
    details_df = pd.read_excel(details_file)
    
    # Standardize Bulgari spelling
    details_df['Maison'] = details_df['Maison'].str.replace('Bvlgari', 'Bulgari', regex=False)
    
    # Categories to count
    categories = ['Product', 'Place', 'Partnership', 'ESG', 'Performance', 
                  'Digital', 'Pricing', 'Promotion', 'People', 'Awards']
    
    # Create verified mentions dataframe
    verified_data = []
    
    for _, row in details_df.iterrows():
        maison = row['Maison']
        year_value = row['Year']
        
        # Count mentions for each category
        mention_counts = {}
        for category in categories:
            mention_counts[category] = count_mentions_in_details(row, category)
        
        # Calculate total mentions
        total_mentions = sum(mention_counts.values())
        
        # Create row data
        row_data = {
            'Maison': maison,
            'Year': year_value,
            **mention_counts,
            'Total_Mentions': total_mentions
        }
        
        verified_data.append(row_data)
    
    # Create DataFrame
    verified_df = pd.DataFrame(verified_data)
    
    # Sort by total mentions (descending)
    verified_df = verified_df.sort_values('Total_Mentions', ascending=False)
    
    # Save to Excel
    verified_df.to_excel(output_file, index=False)
    
    print(f"Generated: {output_file}")
    print(f"Total Maisons: {len(verified_df)}")
    print(f"Total mentions across all categories: {verified_df['Total_Mentions'].sum()}")
    
    # Show top 5 Maisons
    print("Top 5 Maisons by total mentions:")
    for _, row in verified_df.head(5).iterrows():
        print(f"  {row['Maison']}: {row['Total_Mentions']} mentions")
    
    return verified_df

def main():
    """Generate verified mentions for all years"""
    
    years = ['2019', '2022', '2023', '2024', '2025H1']
    
    print("=== GENERATING VERIFIED MENTIONS WORKBOOKS ===")
    print()
    
    for year in years:
        try:
            generate_verified_mentions(year)
            print()
        except Exception as e:
            print(f"Error processing {year}: {str(e)}")
            print()
    
    print("=== VERIFICATION COMPLETE ===")
    print("Generated files:")
    for year in years:
        output_file = f'lvmh_{year}FY_maison_mentions_verified.xlsx' if year != '2025H1' else f'lvmh_{year}_maison_mentions_verified.xlsx'
        if os.path.exists(output_file):
            print(f"  ✅ {output_file}")
        else:
            print(f"  ❌ {output_file} (not generated)")

if __name__ == "__main__":
    main()
