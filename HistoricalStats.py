import http.client
import json
import pandas as pd
from openpyxl import load_workbook, Workbook
import os
from dotenv import load_dotenv  # Import dotenv to load environment variables

# Load environment variables from .env file
load_dotenv()

# Get API Key from environment variables
api_key = os.getenv("API_KEY")

# Ensure the API key is set
if not api_key:
    raise ValueError("Error: API_KEY is missing! Please set it in the .env file.")

# Get the script's directory dynamically
script_dir = os.path.dirname(os.path.abspath(__file__))

# Define input and output folder paths
input_folder = os.path.join(script_dir, "Inputfiles")
output_folder = os.path.join(script_dir, "Outputfiles")

# Ensure the output folder exists
os.makedirs(output_folder, exist_ok=True)

# Construct file paths
excel_output = os.path.join(output_folder, "HistoricalStatsUnpivoted.xlsx")

# API details
API_HOST = "stock.indianapi.in"

# List of required statistics
STAT_TYPES = [
    "quarter_results", "yoy_results", "balancesheet", "cashflow", 
    "ratios", "shareholding_pattern_quarterly", "shareholding_pattern_yearly"
]

def fetch_historical_stats(symbol, stat_type):
    """Fetch historical stats for a given stock symbol and stat type."""
    conn = http.client.HTTPSConnection(API_HOST)
    headers = {'X-Api-Key': api_key}
    url = f"/historical_stats?stock_name={symbol}&stats={stat_type}"
    
    conn.request("GET", url, headers=headers)
    res = conn.getresponse()
    data = res.read()
    conn.close()
    
    try:
        return json.loads(data.decode("utf-8"))
    except json.JSONDecodeError:
        print(f"Error decoding JSON for {symbol} - {stat_type}")
        return None


def process_data(symbol, stat_type, data):
    """Transform API response data into a structured format for Excel."""
    records = []

    if not data:
        print(f"No data available for {symbol} - {stat_type}")
        return records

    for attribute, values in data.items():
        for date, value in values.items():
            records.append([symbol, date, attribute, value])

    return records


def save_to_excel(data, excel_path, sheet_name):
    """Save structured data to an Excel file after deleting old data."""
    df = pd.DataFrame(data, columns=["Symbol", "Date", "Attribute", "Value"])
    
    try:
        # Remove old data before writing new data
        if os.path.exists(excel_path):
            os.remove(excel_path)
            print(f"Deleted old file: {excel_path}")

        # Create a new workbook and write data
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)

        print(f"Data saved successfully in {excel_path} -> {sheet_name}")

    except Exception as e:
        print(f"Error saving to Excel: {e}")


def process_symbols(symbols, excel_path, sheet_name):
    """Fetch and save historical stats for multiple symbols."""
    all_data = []

    for symbol in symbols:
        print(f"Processing {symbol}...")
        for stat_type in STAT_TYPES:
            print(f"  Fetching {stat_type} data...")
            data = fetch_historical_stats(symbol, stat_type)
            all_data.extend(process_data(symbol, stat_type, data))

    if all_data:
        save_to_excel(all_data, excel_path, sheet_name)
    else:
        print("No data to save.")


# Example usage
symbols = ["COCHINSHIP", "DIXON", "DEEPAKFERT"]  # Replace with your stock symbols
sheet_name = "Historical Data"

# Run the script
process_symbols(symbols, excel_output, sheet_name)
