import yfinance as yf
import pandas as pd
from datetime import datetime

def fetch_stock_data(symbol):
    stock = yf.Ticker(symbol)
    df = stock.history(period="1y")  # Fetch 1 year of data
    if not df.empty:
        df = df[['Open', 'High', 'Low', 'Close', 'Volume']]
        df.index = df.index.tz_localize(None)  # Remove timezone information
        return df
    else:
        print(f"No data found for {symbol}")
        return None

def main():
    # List of symbols for Borsa Istanbul stocks
    symbols = ['TCELL.IS', 'KCHOL.IS', 'ASELS.IS']  # Yahoo Finance symbols for BIST stocks
    all_data = {}

    for symbol in symbols:
        print(f"Fetching data for {symbol}...")
        stock_data = fetch_stock_data(symbol)
        if stock_data is not None:
            all_data[symbol] = stock_data

    # Save the data to an Excel file
    if all_data:
        with pd.ExcelWriter(f'stock_data_{datetime.now().strftime("%Y%m%d")}.xlsx') as writer:
            for symbol, data in all_data.items():
                data.to_excel(writer, sheet_name=symbol)
        print("Data successfully fetched and saved.")
    else:
        print("No data fetched, Excel file not created.")

if __name__ == "__main__":
    main()
