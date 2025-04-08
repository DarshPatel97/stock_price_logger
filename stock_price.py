import pandas as pd
import yfinance as yf
import matplotlib.pyplot as plt
import seaborn as sns

# Step 1: Fetch Walmart stock data since 01/01/2000
ticker = "WMT"  # Walmart's stock symbol
start_date = "2000-01-01"
stock_data = yf.download(ticker, start=start_date)

# Step 2: Select necessary columns (Date, Open, Close)
df = stock_data[['Open', 'Close']].copy()
df.reset_index(inplace=True)  # Reset index to make 'Date' a column
df['Date'] = df['Date'].dt.strftime('%m/%d/%Y')  # Convert to MM/DD/YYYY format

# Step 3: Compute Daily Change (% change from the previous day)
df['Daily Change'] = df['Close'].pct_change() * 100  # Percentage change

# Step 4: Compute Absolute Change (% change since 01/01/2000)
initial_price = df.loc[0, 'Close']  # Closing price on 01/01/2000
df['Absolute Change'] = ((df['Close'] - initial_price) / initial_price) * 100

# Step 5: Save data to an Excel file
file_name = "walmart_stock.xlsx"
df.to_excel(file_name, index=True)
print(f"✅ Stock data saved to {file_name}")