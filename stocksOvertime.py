import yfinance as yf
import datetime
import pandas as pd
import matplotlib.pyplot as plt

# ----------------------------
# Step 1: Define Date Range
# ----------------------------
start_date_str = "2025-01-27"
end_date_str   = "2025-02-15"

# "Today" is February 15, 2025
today = datetime.date(2025, 2, 15)

# Convert the assignment's end date to a date object
end_date = datetime.datetime.strptime(end_date_str, "%Y-%m-%d").date()

# Use today's date if it comes before the assignment's end date
if today < end_date:
    effective_end_date = today
else:
    effective_end_date = end_date

# IMPORTANT: yfinance’s 'end' parameter is exclusive.
# Add one day to include data for the effective_end_date.
effective_end_date_str = (effective_end_date + datetime.timedelta(days=1)).strftime("%Y-%m-%d")

# ----------------------------
# Step 2: Define the Stock Tickers
# ----------------------------
tickers = ["NVDA", "SNPS", "LOPE"]

# ----------------------------
# Step 3: Download the Historical Data
# ----------------------------
print(f"Fetching data from {start_date_str} to {effective_end_date} for tickers: {tickers}")
data = yf.download(tickers, start=start_date_str, end=effective_end_date_str)

# Display the retrieved data
print("\nHistorical Stock Data:")
print(data)

# ----------------------------
# Step 4: Plot the Closing Prices and Save to an Image File
# ----------------------------
if 'Close' in data.columns.get_level_values(0):
    # Extract the closing prices
    close_data = data['Close']
    
    # Create the plot
    ax = close_data.plot(figsize=(10, 6), title="Closing Prices for NVDA, SNPS, LOPE")
    plt.xlabel("Date")
    plt.ylabel("Price (USD)")
    plt.legend(tickers)
    plt.tight_layout()
    
    # Save the plot as a PNG file instead of displaying it
    plot_filename = "closing_prices.png"
    plt.savefig(plot_filename)
    plt.close()  # Close the figure to avoid blocking further execution
    
    print(f"Plot saved as '{plot_filename}'")
    
    # ----------------------------
    # Step 4.1: Calculate Money Gain
    # ----------------------------
    # Get the first and last available closing prices
    first_prices = close_data.iloc[0]
    last_prices  = close_data.iloc[-1]
    
    # Define the number of shares held for each ticker
    shares = {"NVDA": 26, "SNPS": 6, "LOPE": 20}
    
    money_gain = {}
    for ticker in tickers:
        if ticker in close_data.columns:
            gain = (last_prices[ticker] - first_prices[ticker]) * shares[ticker]
            money_gain[ticker] = gain
            print(f"Money gain for {ticker} shares: ${gain:.2f}")
    
    total_gain = sum(money_gain.values())
    print(f"Total money gain: ${total_gain:.2f}")
    
    # Create a DataFrame for money gains
    money_gain_df = pd.DataFrame(list(money_gain.items()), columns=["Ticker", "Money Gain"])
    total_df = pd.DataFrame([["Total", total_gain]], columns=["Ticker", "Money Gain"])
    money_gain_df = pd.concat([money_gain_df, total_df], ignore_index=True)
    
else:
    print("No 'Close' price data found in the downloaded dataset.")

# ----------------------------
# Step 5: Write the Data to an Excel File
# ----------------------------
excel_file = "stock_data_overtime.xlsx"

with pd.ExcelWriter(excel_file, engine="openpyxl") as writer:
    # Write the entire dataset to a sheet named "HistoricalData"
    data.to_excel(writer, sheet_name="HistoricalData")
    
    # Write only the closing prices to another sheet if available
    if 'Close' in data.columns.get_level_values(0):
        close_data.to_excel(writer, sheet_name="ClosingPrices")
        # Write the money gain calculations to a new sheet "MoneyGain"
        money_gain_df.to_excel(writer, sheet_name="MoneyGain", index=False)

print(f"Data has been successfully written to '{excel_file}'")
