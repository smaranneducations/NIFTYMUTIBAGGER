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
table_path = os.path.join(input_folder, "StockSymbols.xlsx")
excel_output = os.path.join(output_folder, "StockData.xlsx")

def fetch_stock_data(symbol, period="5yr", filter="default"):
    """Fetch stock data from the API."""
    conn = http.client.HTTPSConnection("stock.indianapi.in")
    headers = {'X-Api-Key': api_key}
    url = f"/historical_data?stock_name={symbol}&period={period}&filter={filter}"
    conn.request("GET", url, headers=headers)
    res = conn.getresponse()
    data = res.read()
    conn.close()
    return json.loads(data.decode("utf-8"))

def save_to_excel_consolidated(data, path, sheet_name, symbol):
    """Append stock data to Excel while adding a Symbol column."""
    
    if 'datasets' not in data or not data['datasets']:  # Check if datasets exist and are not empty
        print(f"No data found for {symbol}. Skipping...")
        return  # Skip to next symbol if no data

    consolidated_df = pd.DataFrame()
    
    for dataset in data['datasets']:
        if dataset['values']:  # Check if values exist and are not empty
            example_entry = dataset['values'][0]
            columns = ['Date'] + [f'{dataset["label"]} {i}' for i in range(1, len(example_entry))]
            df = pd.DataFrame(dataset['values'], columns=columns)
            
            if consolidated_df.empty:
                consolidated_df = df
            else:
                consolidated_df = pd.merge(consolidated_df, df, on='Date', how='outer')
    
    consolidated_df['Symbol'] = symbol  # Add the Symbol column

    try:
        # Remove old file before writing new data
        if os.path.exists(path):
            os.remove(path)
            print(f"Deleted old file: {path}")

        # Create a new workbook and write data
        with pd.ExcelWriter(path, engine='openpyxl') as writer:
            consolidated_df.to_excel(writer, sheet_name=sheet_name, index=False)

        print(f"Data for {symbol} has been written to {sheet_name} in Excel.")

    except Exception as e:
        print(f"An error occurred while saving to Excel for {symbol}: {e}")


def get_stock_data_to_excel(symbol, period, filter, excel_path, sheet_name):
    """Retrieve stock data and append to Excel."""
    data = fetch_stock_data(symbol, period, filter)  # Fetch data first
    if data:  # Ensure data is valid before proceeding
        save_to_excel_consolidated(data, excel_path, sheet_name, symbol)
    else:
        print(f"Failed to fetch data for {symbol}. Check API key and parameters.")


def process_all_symbols(table_path, excel_output, sheet_name, period="5yr", filter="default", max_symbols=200):
    """Read symbols from 'AllIndices' table and fetch stock data for each, allowing period and filter customization."""
    
    if not os.path.exists(table_path):
        print(f"Error: The file {table_path} does not exist.")
        return

    df = pd.read_excel(table_path, sheet_name="AllIndices")  # Load table

    if 'Symbol' not in df.columns:
        print("Error: 'Symbol' column not found in 'AllIndices' table.")
        return

    symbols = df['Symbol'].dropna().unique()[:max_symbols]  # Get unique symbols (limit to 200)

    for symbol in symbols:
        print(f"Processing: {symbol} with period={period} and filter={filter}")
        get_stock_data_to_excel(symbol, period, filter, excel_output, sheet_name)


# Run with default parameters
process_all_symbols(table_path, excel_output, "Sheet1", period="5yr", filter="default", max_symbols=200)
