import yfinance as yf
import datetime
import pandas as pd
import matplotlib.pyplot as plt

# ----------------------------
# Step 1: Define Date Range
# ----------------------------

# Assignment-specified start and end dates (as strings)
start_date_str = "2025-01-27"
end_date_str   = "2025-02-16"

# "Today" is February 11, 2025
today = datetime.date(2025, 2, 11)

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
else:
    print("No 'Close' price data found in the downloaded dataset.")

# ----------------------------
# Step 5: Write the Data to an Excel File
# ----------------------------
excel_file = "stock_data.xlsx"

with pd.ExcelWriter(excel_file, engine="openpyxl") as writer:
    # Write the entire dataset to a sheet named "HistoricalData"
    data.to_excel(writer, sheet_name="HistoricalData")
    
    # Optionally, write only the closing prices to another sheet
    if 'Close' in data.columns.get_level_values(0):
        close_data.to_excel(writer, sheet_name="ClosingPrices")

print(f"Data has been successfully written to '{excel_file}'")
