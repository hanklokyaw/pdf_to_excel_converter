# grab daily trade information
import pandas as pd
import requests
from datetime import datetime
from FinMind.data import DataLoader
import os
import camelot

# Get today's date
today_date = datetime.today().strftime('%Y%m%d')

# Send a GET request to the API
response = requests.get("https://www.twse.com.tw/exchangeReport/STOCK_DAY_ALL?response=json")

# Parse the JSON response
json_data = response.json()

def fetch_data():

    # Extract relevant data from the JSON
    data_list = []
    for entry in json_data['data']:
        # Convert comma-separated values to numerical types
        entry[2] = int(entry[2].replace(',', ''))  # Convert 成交股數 to int
        entry[3] = int(entry[3].replace(',', ''))  # Convert 成交金額 to int
        
        # Remove commas and convert strings to float
        entry[4] = float(entry[4].replace(',', ''))  # Convert 開盤價 to float
        entry[5] = float(entry[5].replace(',', ''))  # Convert 最高價 to float
        entry[6] = float(entry[6].replace(',', ''))  # Convert 最低價 to float
        entry[7] = float(entry[7].replace(',', ''))  # Convert 收盤價 to float
        entry[9] = int(entry[9].replace(',', ''))  # Convert 成交筆數 to int
        
        # Create a dictionary with the extracted data
        data_dict = {
            "Ticker": entry[0],
            "Name": entry[1],
            "Trading_Volume": entry[2],
            "Trading_Money": entry[3],
            "Open": entry[4],
            "High": entry[5],
            "Low": entry[6],
            "Close": entry[7],
            "Spread": entry[8],
            "Trading_Turnover": entry[9]
        }
        
        # Append the dictionary to the list
        data_list.append(data_dict)

    # Create a DataFrame from the list of dictionaries
    twse_df = pd.DataFrame(data_list)

    # Convert 'Spread' column to float while retaining negative signs
    twse_df['Spread'] = twse_df['Spread'].str.replace('[X]', '', regex=True)  # Remove 'X'
    twse_df['Spread'] = twse_df['Spread'].str.replace('[+]', '', regex=True)  # Remove '+'
    twse_df['Spread'] = twse_df['Spread'].astype(float)

    # Display the DataFrame
    print(twse_df)

    # Save DataFrame to Excel file
    file_name = f"TWSE_update_{today_date}.xlsx"
    twse_df.to_excel(file_name, index=False)

    print(f"File saved as: {file_name}")
    
    
def download_historical_prices(ticker, date):

    try:
        dl = DataLoader()
        stock_data = dl.taiwan_stock_daily(stock_id=ticker, start_date=date)
        stock_data.to_excel(f"{ticker}_{date}.xlsx", index=False)
        print(f"Saved {ticker}_{date}.xlsx")
        return stock_data
    except:
        print(f"Failed to save {ticker}_{date}.xlsx")
        


def convert_pdf_folder_to_excel(source_folder, destination_folder):
    # Ensure destination folder exists
    if not os.path.exists(destination_folder):
        os.makedirs(destination_folder)

    # Loop through all files in the source folder
    for filename in os.listdir(source_folder):
        if filename.endswith('.pdf'):
            source_file = os.path.join(source_folder, filename)
            dest_file = os.path.join(destination_folder, filename.replace('.pdf', '.xlsx'))

            # Read PDF file using camelot
            tables = camelot.read_pdf(source_file, pages='all', flavor='stream')

            # Save each table to an Excel file
            with pd.ExcelWriter(dest_file) as writer:
                for i, table in enumerate(tables):
                    table.df.to_excel(writer, sheet_name=f"Sheet_{i}", index=False)
            print(f"Converted {filename} to {dest_file}")


def convert_pdf_to_excel(source_file, page='all'):
    # Get the base name and directory of the source file
    base_name = os.path.basename(source_file)
    directory = os.path.dirname(source_file)
    
    # Replace .pdf with .xlsx
    excel_file = os.path.join(directory, base_name.replace('.pdf', '.xlsx'))
    
    # Read PDF file using camelot
    tables = camelot.read_pdf(source_file, pages=page, flavor='stream')
    
    # Save each table to an Excel file
    with pd.ExcelWriter(excel_file) as writer:
        for i, table in enumerate(tables):
            table.df.to_excel(writer, sheet_name=f"Sheet_{i}", index=False)
    
    print(f"Converted {source_file} to {excel_file}")

