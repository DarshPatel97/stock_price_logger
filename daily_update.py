import pandas as pd
import yfinance as yf
import datetime
import time
import pytz

# Define the Excel file name
file_name = "walmart_stock.xlsx"

# Get the current date in EST timezone (market time zone)
est = pytz.timezone('US/Eastern')
current_time_est = datetime.datetime.now(est)
current_date = current_time_est.strftime('%m/%d/%Y')

# Function to get today's stock data and update the file
def update_stock_data():
    ticker = "WMT"  # Walmart's stock symbol

    # Fetch stock data for today (open and close prices)
    stock_data = yf.download(ticker, start=current_date, end=current_date, interval='1m')

    # Extract the Open and Close prices
    open_price = stock_data.iloc[0]['Open']  # Market open (first row)
    close_price = stock_data.iloc[-1]['Close']  # Market close (last row)
    
    # Read the existing data from the Excel file
    try:
        df = pd.read_excel(file_name)
    except FileNotFoundError:
        # If file doesn't exist, create a new one with basic columns
        df = pd.DataFrame(columns=["Date", "Open", "Close", "Daily Change", "Absolute Change"])

    # Append the new Open and Close prices for today
    new_row = {
        "Date": current_date,
        "Open": open_price,
        "Close": close_price,
        "Daily Change": None,  # This will be updated later
        "Absolute Change": None  # This will be updated later
    }
    df = df.append(new_row, ignore_index=True)

    # Save back to Excel
    df.to_excel(file_name, index=False)

    print(f"Stock data updated for {current_date}. Open: {open_price}, Close: {close_price}")

# Check if the current time is around market open or close
def market_open():
    # Define market open time (9:30 AM EST)
    open_time = datetime.datetime(current_time_est.year, current_time_est.month, current_time_est.day, 9, 30, tzinfo=est)
    return current_time_est >= open_time and current_time_est < (open_time + datetime.timedelta(minutes=5))

def market_close():
    # Define market close time (4:00 PM EST)
    close_time = datetime.datetime(current_time_est.year, current_time_est.month, current_time_est.day, 16, 0, tzinfo=est)
    return current_time_est >= close_time and current_time_est < (close_time + datetime.timedelta(minutes=5))

# Main loop: Check if it's market open or close time and update
while True:
    if market_open():
        update_stock_data()  # Update stock data for market open
        time.sleep(60)  # Wait for a minute before checking again
    elif market_close():
        update_stock_data()  # Update stock data for market close
        time.sleep(60)  # Wait for a minute before checking again
    else:
        time.sleep(60)  # Check every minute for market open/close
