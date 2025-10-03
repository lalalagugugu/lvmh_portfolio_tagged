#!/usr/bin/env python3
"""
Generate standardized verified mention files based on activity details
"""

import pandas as pd
import os
import re

def clean_citations(text):
    """Remove citation marks from text"""
    if pd.isna(text) or not isinstance(text, str):
        return text
    
    # Remove citation patterns like 【879963710027701†L3310-L3314】
    text = re.sub(r'【[^】]*】', '', text)
    
    # Clean up extra spaces
    text = re.sub(r'\s+', ' ', text).strip()
    
    return text if text else None

def count_activities_for_maison(details_row):
    """Count non-empty activities for each category for a Maison"""
    categories = ['Product', 'Place', 'Partnership', 'ESG', 'Performance', 
                  'Digital', 'Pricing', 'Promotion', 'People', 'Awards']
    
    counts = {}
    for category in categories:
        count = 0
        # Look for columns that start with this category name followed by underscore
        for col in details_row.index:
            if col.startswith(f'{category}_') and pd.notna(details_row[col]) and str(details_row[col]).strip():
                activity_text = str(details_row[col]).strip()
                activity_text = clean_citations(activity_text)
                # Count only non-empty activities (excluding '0', 'nan', empty strings)
                if activity_text and activity_text != '0' and activity_text.lower() not in ['nan', '']:
                    count += 1
        counts[category] = count
    
    # Calculate total mentions
    counts['Total_Mentions'] = sum(counts.values())
    
    return counts

def generate_standardized_verified(year_str):
    """Generate standardized verified mentions for a specific year"""
    
    # Define file names - check for standardized files first
    if year_str == '2025H1':
        details_file = f'lvmh_{year_str}_maison_details_standardized.xlsx'
        output_file = f'lvmh_{year_str}_maison_mentions_standardized_verified.xlsx'
    else:
        details_file = f'lvmh_{year_str}FY_maison_details_standardized.xlsx'
        output_file = f'lvmh_{year_str}FY_maison_mentions_standardized_verified.xlsx'
    
    print(f"Processing {year_str}...")
    
    if not os.path.exists(details_file):
        print(f"  Details file not found: {details_file}")
        return
    
    # Read details file
    df_details = pd.read_excel(details_file)
    print(f"  Loaded details file: {len(df_details)} rows")
    
    # Process each Maison
    verified_data = []
    
    for index, row in df_details.iterrows():
        maison = row['Maison']
        year_value = row['Year'] if 'Year' in row else year_str
        
        # Count activities for this Maison
        counts = count_activities_for_maison(row)
        
        # Create row for verified mentions
        verified_row = {
            'Maison': maison,
            'Year': year_value
        }
        
        # Add category counts
        for category in ['Product', 'Place', 'Partnership', 'ESG', 'Performance', 
                         'Digital', 'Pricing', 'Promotion', 'People', 'Awards']:
            verified_row[category] = counts[category]
        
        verified_row['Total_Mentions'] = counts['Total_Mentions']
        
        verified_data.append(verified_row)
    
    # Create DataFrame
    verified_df = pd.DataFrame(verified_data)
    
    # Sort by Total_Mentions descending
    verified_df = verified_df.sort_values('Total_Mentions', ascending=False).reset_index(drop=True)
    
    print(f"  Generated verified file: {output_file}")
    print(f"  Total Maisons: {len(verified_df)}")
    print(f"  Total mentions across all categories: {verified_df['Total_Mentions'].sum()}")
    
    # Show top 5 Maisons
    print(f"  Top 5 Maisons by total mentions:")
    for i, row in verified_df.head(5).iterrows():
        print(f"    {row['Maison']}: {row['Total_Mentions']} mentions")
    
    # Save verified file
    verified_df.to_excel(output_file, index=False)
    print(f"  ✅ Saved: {output_file}")
    
    return verified_df

def main():
    """Generate standardized verified mentions for all years"""
    
    years = ['2019', '2022', '2023', '2024', '2025H1']
    
    print("=== GENERATING STANDARDIZED VERIFIED MENTIONS WORKBOOKS ===")
    print()
    
    generated_files = []
    
    for year in years:
        try:
            df = generate_standardized_verified(year)
            if df is not None:
                filename = f'lvmh_{year}FY_maison_mentions_standardized_verified.xlsx' if year != '2025H1' else f'lvmh_{year}_maison_mentions_standardized_verified.xlsx'
                generated_files.append(filename)
            print()
        except Exception as e:
            print(f"  ❌ Error processing {year}: {str(e)}")
            print()
    
    print("=== VERIFICATION COMPLETE ===")
    print(f"Generated files:")
    for filename in generated_files:
        print(f"  ✅ {filename}")

if __name__ == "__main__":
    main()
