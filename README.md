# LVMH Portfolio Analysis Dashboard

A comprehensive web application for analyzing LVMH Maison activities and mentions across different categories and fiscal years.

## Features

### 📊 Key Performance Indicators (KPIs)
- **Most Mentioned Maison**: The Maison with the highest total mentions in the selected year
- **Previous Year Leader**: The most mentioned Maison from the previous year for comparison
- **Category Leaders**: The most mentioned Maison in each activity category
- **Total Maisons**: Count of all analyzed Maisons

### 📈 Visualizations
- **Stacked Bar Charts**: Top 10 Maisons showing category mentions comparison between years
- **Category Breakdown**: Individual Maison analysis with detailed category mentions
- **Overall Ranking**: Complete ranking of all Maisons by total mentions

### 🔍 Filtering System
- **Fiscal Year**: Select between available years (2023, 2024, and future years)
- **Maison**: Choose specific Maison for detailed analysis or "All" for overview
- **Activity Category**: Filter by specific categories (Product, Place, Partnership, ESG, etc.)

### 📝 Detailed Analysis
When a specific Maison is selected, the dashboard provides:
- Category-wise mention breakdown
- Detailed activity descriptions
- Performance metrics and trends

## Data Structure

The application expects Excel files with the following naming convention:
- `lvmh_YYYYFY_maison_details.xlsx`: Detailed activities for each Maison
- `lvmh_YYYYFY_maison_mentions.xlsx`: Mention counts by category

### Activity Categories
- **Product**: Product launches and innovations
- **Place**: Store openings and location-based activities
- **Partnership**: Collaborations and partnerships
- **ESG**: Environmental, Social, and Governance initiatives
- **Performance**: Business performance mentions
- **Digital**: Digital innovations and online activities
- **Pricing**: Pricing strategies and changes
- **Promotion**: Marketing and promotional activities
- **People**: Leadership changes and personnel moves
- **Awards**: Recognition and awards received

## Installation & Usage

### Prerequisites
- Python 3.7+
- pip package manager

### Setup
1. Install required dependencies:
```bash
pip install -r requirements.txt
```

2. Ensure Excel files are in the project directory with correct naming format

3. Run the application:
```bash
streamlit run app.py
```

4. Open your browser and navigate to the provided local URL (typically `http://localhost:8501`)

## Scalability

The application is designed to easily accommodate additional years of data:
1. Add new Excel files following the naming convention
2. The application will automatically detect and include new years in the filter options
3. All visualizations and KPIs will automatically update to include new data

## File Structure
```
├── app.py                    # Main Streamlit application
├── requirements.txt          # Python dependencies
├── README.md                # This file
├── lvmh_2023FY_maison_details.xlsx
├── lvmh_2023FY_maison_mentions.xlsx
├── lvmh_2024FY_maison_details.xlsx
└── lvmh_2024FY_maison_mentions.xlsx
```

## Technology Stack
- **Streamlit**: Web application framework
- **Pandas**: Data manipulation and analysis
- **Plotly**: Interactive visualizations
- **OpenPyXL**: Excel file processing

## Future Enhancements
- Export functionality for reports
- Advanced filtering options
- Trend analysis across multiple years
- Comparative analysis tools
- Data validation and error handling
