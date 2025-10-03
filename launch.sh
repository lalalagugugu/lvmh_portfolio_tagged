#!/bin/bash

# LVMH Portfolio Analysis Dashboard Launcher
# This script starts the Streamlit web application

echo "🍾 LVMH Portfolio Analysis Dashboard"
echo "===================================="
echo ""

# Check if Python is available
if ! command -v python &> /dev/null; then
    echo "❌ Python is not installed or not in PATH"
    exit 1
fi

# Check if required files exist
if [ ! -f "app.py" ]; then
    echo "❌ app.py not found. Please run this script from the project directory."
    exit 1
fi

# Check if Excel files exist
excel_files=(
    "lvmh_2023FY_maison_details.xlsx"
    "lvmh_2023FY_maison_mentions.xlsx"
    "lvmh_2024FY_maison_details.xlsx"
    "lvmh_2024FY_maison_mentions.xlsx"
)

missing_files=()
for file in "${excel_files[@]}"; do
    if [ ! -f "$file" ]; then
        missing_files+=("$file")
    fi
done

if [ ${#missing_files[@]} -gt 0 ]; then
    echo "⚠️  Warning: The following Excel files are missing:"
    for file in "${missing_files[@]}"; do
        echo "   - $file"
    done
    echo ""
    echo "The application will still run, but some data may not be available."
    echo ""
fi

# Install requirements if needed
echo "📦 Checking dependencies..."
pip install -r requirements.txt --quiet

# Start the application
echo "🚀 Starting LVMH Portfolio Analysis Dashboard..."
echo ""
echo "The application will open in your default web browser."
echo "If it doesn't open automatically, go to: http://localhost:8501"
echo ""
echo "Press Ctrl+C to stop the application."
echo ""

# Start Streamlit
streamlit run app.py --server.port 8501 --server.headless false
