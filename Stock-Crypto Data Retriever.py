import yfinance as yf
import pandas as pd
from datetime import datetime, timedelta

def is_valid_ticker(ticker):
    try:
        df = yf.Ticker(ticker).history(period="1d")
        return not df.empty
    except:
        return False

def get_data(ticker):
    df = yf.Ticker(ticker).history(period="max")
    df.index = df.index.tz_localize(None)
    return df

def save_to_excel(ticker, data):
    filename = f"{ticker}.xlsx"

    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        # Sheet 1: Full data
        data.to_excel(writer, sheet_name="Full History")

        # Sheet 2: Last 5 years
        five_years_ago = datetime.now() - timedelta(days=5*365)
        data_5y = data[data.index >= pd.to_datetime(five_years_ago)]

        if not data_5y.empty:
            data_5y.to_excel(writer, sheet_name="Last 5 Years")
        else:
            pd.DataFrame({"Message": ["Less than 5 years of data available."]}).to_excel(writer, sheet_name="Last 5 Years")

    print(f"File Saved: {filename}")


def main():
    print("Yahoo Finance Ticker Downloader to Excel")

    while True:
        try:
            count = int(input("How many stocks/cryptos do you want to check? "))
            if count <= 0:
                print("Enter a number greater than 0.")
                continue
            break
        except ValueError:
            print("Invalid input. Please enter an integer.")

    for i in range(count):
        while True:
            ticker = input(f"Enter ticker #{i+1} (e.g., AAPL or BTC-USD): ").upper().strip()
            if is_valid_ticker(ticker):
                print(f"Valid ticker: {ticker}. Downloading data...\n")
                data = get_data(ticker)
                save_to_excel(ticker, data)
                break
            else:
                print(f"Invalid ticker: {ticker}. Please try again.\n")

    print("All files processed.")

if __name__ == "__main__":
    main()


