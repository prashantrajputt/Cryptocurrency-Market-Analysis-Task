import requests
import pandas as pd
import time
from datetime import datetime
import xlwings as xw
import os

def fetch_crypto_data():
    """
    Fetch top 50 cryptocurrency data from CoinGecko API
    """
    try:
        url = "https://api.coingecko.com/api/v3/coins/markets"
        params = {
            "vs_currency": "usd",
            "order": "market_cap_desc",
            "per_page": 50,
            "page": 1,
            "sparkline": False
        }
        response = requests.get(url, params=params)
        response.raise_for_status()
        return response.json()
    except requests.RequestException as e:
        print(f"Error fetching data: {e}")
        return None
def process_crypto_data(data):
    """
    Process raw API data into a pandas DataFrame
    """
    df = pd.DataFrame(data)
    df = df[[
        'name', 'symbol', 'current_price', 'market_cap',
        'total_volume', 'price_change_percentage_24h'
    ]]
    df.columns = [
        'Name', 'Symbol', 'Price (USD)', 'Market Cap',
        '24h Volume', '24h Change %'
    ]
    return df
def analyze_data(df):
    """
    Perform analysis on cryptocurrency data
    """
    analysis = {
        'top_5_by_market_cap': df.head(),
        'average_price': df['Price (USD)'].mean(),
        'highest_24h_change': df.nlargest(1, '24h Change %'),
        'lowest_24h_change': df.nsmallest(1, '24h Change %')
    }
    return analysis
def update_excel(df, analysis, wb_path):
    """
    Update Excel workbook with latest data and analysis
    """
    try:
        wb = xw.Book(wb_path)
        
        # Update main data sheet
        sheet_data = wb.sheets['CryptoData']
        sheet_data.range('A1').value = df
        sheet_data.range('A1').expand().autofit()
        
        # Update analysis sheet
        sheet_analysis = wb.sheets['Analysis']
        sheet_analysis.range('A1').value = "Top 5 by Market Cap"
        sheet_analysis.range('A2').value = analysis['top_5_by_market_cap']
        
        sheet_analysis.range('A8').value = "Average Price (USD)"
        sheet_analysis.range('B8').value = analysis['average_price']
        
        sheet_analysis.range('A10').value = "Highest 24h Change"
        sheet_analysis.range('A11').value = analysis['highest_24h_change']
        
        sheet_analysis.range('A13').value = "Lowest 24h Change"
        sheet_analysis.range('A14').value = analysis['lowest_24h_change']
        
        # Add timestamp
        sheet_data.range('I1').value = f"Last Updated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        
        wb.save()
    except Exception as e:
        print(f"Error updating Excel: {e}")
def generate_report(analysis, report_path):
    """
    Generate analysis report in markdown format
    """
    report = f"""# Cryptocurrency Market Analysis Report
Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

## Top 5 Cryptocurrencies by Market Cap
{analysis['top_5_by_market_cap'].to_markdown()}

## Market Overview
- Average Price: ${analysis['average_price']:.2f}
- Highest 24h Change: {analysis['highest_24h_change']['Name'].values[0]} ({analysis['highest_24h_change']['24h Change %'].values[0]:.2f}%)
- Lowest 24h Change: {analysis['lowest_24h_change']['Name'].values[0]} ({analysis['lowest_24h_change']['24h Change %'].values[0]:.2f}%)
"""
    with open(report_path, 'w') as f:
        f.write(report)            
def main():
    wb_path = 'crypto_analysis.xlsx'
    report_path = 'crypto_analysis_report.md'
    update_interval = 300  # 5 minutes in seconds
    
    # Create Excel workbook if it doesn't exist
    if not os.path.exists(wb_path):
        wb = xw.Book()
        wb.sheets.add('CryptoData')
        wb.sheets.add('Analysis')
        wb.save(wb_path)
    
    while True:
        print(f"Fetching data at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        
        # Fetch and process data
        data = fetch_crypto_data()
        if data:
            df = process_crypto_data(data)
            analysis = analyze_data(df)
            
            # Update Excel and generate report
            update_excel(df, analysis, wb_path)
            generate_report(analysis, report_path)
            
            print("Data updated successfully")
        
        time.sleep(update_interval)

if __name__ == "__main__":
    main()        