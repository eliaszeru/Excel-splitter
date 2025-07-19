"""
Create sample Excel data for testing the Excel Splitter application
"""
import pandas as pd
import numpy as np
from datetime import datetime, timedelta

def create_sample_data():
    """Create sample Excel data with various columns for testing"""
    
    # Sample data for fashion industry (Ralph Lauren context)
    np.random.seed(42)  # For reproducible results
    
    # Generate 100 sample records
    n_records = 100
    
    # Product categories
    categories = ['Shirts', 'Pants', 'Dresses', 'Shoes', 'Accessories', 'Outerwear']
    
    # Seasons
    seasons = ['Spring', 'Summer', 'Fall', 'Winter']
    
    # Genders
    genders = ['Men', 'Women', 'Unisex']
    
    # Colors
    colors = ['Black', 'White', 'Blue', 'Red', 'Green', 'Brown', 'Gray', 'Pink']
    
    # Sizes
    sizes = ['XS', 'S', 'M', 'L', 'XL', 'XXL']
    
    # Regions
    regions = ['North America', 'Europe', 'Asia', 'South America', 'Africa']
    
    # Price ranges
    price_ranges = ['Budget', 'Mid-range', 'Premium', 'Luxury']
    
    # Generate sample data
    data = {
        'Product_ID': [f'PROD{i:03d}' for i in range(1, n_records + 1)],
        'Product_Name': [f'Product {i}' for i in range(1, n_records + 1)],
        'Category': np.random.choice(categories, n_records),
        'Season': np.random.choice(seasons, n_records),
        'Gender': np.random.choice(genders, n_records),
        'Color': np.random.choice(colors, n_records),
        'Size': np.random.choice(sizes, n_records),
        'Region': np.random.choice(regions, n_records),
        'Price_Range': np.random.choice(price_ranges, n_records),
        'Price': np.random.uniform(25, 500, n_records).round(2),
        'Stock_Quantity': np.random.randint(0, 100, n_records),
        'Active': np.random.choice(['Yes', 'No'], n_records),
        'Launch_Date': [(datetime.now() - timedelta(days=np.random.randint(0, 365))).strftime('%Y-%m-%d') for _ in range(n_records)]
    }
    
    # Create DataFrame
    df = pd.DataFrame(data)
    
    # Save to Excel file
    filename = 'sample_data.xlsx'
    df.to_excel(filename, index=False, sheet_name='Products')
    
    print(f"✅ Sample data created: {filename}")
    print(f"📊 Total records: {len(df)}")
    print(f"📋 Columns: {list(df.columns)}")
    print("\n📈 Sample data preview:")
    print(df.head())
    
    print("\n🎯 Example rules you can test:")
    print("1. Single Column: Category = 'Shirts'")
    print("2. AND Logic: Gender = 'Men' AND Season = 'Winter'")
    print("3. OR Logic: Color = 'Black' OR Color = 'White'")
    print("4. AND Logic: Price_Range = 'Premium' AND Region = 'North America'")
    
    return filename

if __name__ == "__main__":
    create_sample_data() 