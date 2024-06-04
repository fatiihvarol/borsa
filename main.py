import requests
import pandas as pd
from datetime import datetime
import time

# Alpha Vantage API anahtarını buraya ekleyin
API_KEY = 'N5XTM8CU5RCTU8JV'
BASE_URL = 'https://www.alphavantage.co/query'

def fetch_stock_data(symbol):
    params = {
        'function': 'TIME_SERIES_DAILY',
        'symbol': symbol,
        'apikey': API_KEY,
        'outputsize': 'compact'
    }
    response = requests.get(BASE_URL, params=params)
    data = response.json()
    if 'Time Series (Daily)' in data:
        df = pd.DataFrame.from_dict(data['Time Series (Daily)'], orient='index')
        df = df.rename(columns={
            '1. open': 'Open',
            '2. high': 'High',
            '3. low': 'Low',
            '4. close': 'Close',
            '5. volume': 'Volume'
        })
        df.index = pd.to_datetime(df.index)
        df = df.sort_index()
        return df
    else:
        print(f"Error fetching data for {symbol}: {data}")
        return None

def main():
    # İzlemek istediğiniz hisse senetlerinin sembollerini buraya ekleyin
    symbols = ['AAPL', 'GOOGL', 'MSFT']
    all_data = {}

    for symbol in symbols:
        print(f"Fetching data for {symbol}...")
        stock_data = fetch_stock_data(symbol)
        if stock_data is not None:
            all_data[symbol] = stock_data
        time.sleep(3)  # Alpha Vantage ücretsiz hesaplar için saniyede 5 istek limitine sahiptir

    # Verileri birleştirip kaydetmek
    with pd.ExcelWriter(f'stock_data_{datetime.now().strftime("%Y%m%d")}.xlsx') as writer:
        for symbol, data in all_data.items():
            data.to_excel(writer, sheet_name=symbol)
    
    print("Veriler başarıyla çekildi ve kaydedildi.")

if __name__ == "__main__":
    main()
