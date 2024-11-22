import requests
import pandas as pd
from openpyxl import Workbook
import time

# Fetch Live Data
def fetch_crypto_data():
    url = "https://api.coingecko.com/api/v3/coins/markets"
    params = {"vs_currency": "usd", "order": "market_cap_desc", "per_page": 50, "page": 1}
    response = requests.get(url, params=params)
    data = response.json()
    return pd.DataFrame(data)

# Analyze Data
def analyze_data(df):
    top_5 = df.nlargest(5, 'market_cap')
    avg_price = df['current_price'].mean()
    highest_change = df.nlargest(1, 'price_change_percentage_24h')
    lowest_change = df.nsmallest(1, 'price_change_percentage_24h')
    return top_5, avg_price, highest_change, lowest_change

# Update Excel
def update_excel(df, filename="C:/Users/visha/Documents/crypto_datas.xlsx"):
    with pd.ExcelWriter(filename, mode="w", engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Live Data")

# Main Loop
if __name__ == "__main__":
    while True:
        data = fetch_crypto_data()
        top_5, avg_price, high, low = analyze_data(data)
        
        print("Top 5 Cryptocurrencies:\n", top_5)
        print("Average Price:", avg_price)
        print("Highest Change:\n", high)
        print("Lowest Change:\n", low)
        
        update_excel(data)
        time.sleep(300)  # Updates every 5 minutes
