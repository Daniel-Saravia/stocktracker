import yfinance as yf
import datetime
import pandas as pd

# ----------------------------
# Step 1: Define Date Range
# ----------------------------
start_date_str = "2025-01-27"
end_date_str   = "2025-02-15"
today = datetime.date(2025, 2, 15)

# Convert the assignment's end date to a date object
end_date = datetime.datetime.strptime(end_date_str, "%Y-%m-%d").date()

# Use today's date if it comes before the assignment's end date
effective_end_date = today if today < end_date else end_date

# IMPORTANT: yfinance’s 'end' parameter is exclusive.
# Add one day to include data for the effective_end_date.
effective_end_date_str = (effective_end_date + datetime.timedelta(days=1)).strftime("%Y-%m-%d")

# ----------------------------
# Step 2: Define Stock Tickers and Shares
# ----------------------------
tickers = ["NVDA", "SNPS", "LOPE"]
shares = {"NVDA": 26, "SNPS": 6, "LOPE": 20}

# ----------------------------
# Step 3: Download the Historical Data
# ----------------------------
print(f"Fetching data from {start_date_str} to {effective_end_date} for tickers: {tickers}")
data = yf.download(tickers, start=start_date_str, end=effective_end_date_str)

# ----------------------------
# Step 4: Calculate Money Gain and Prepare a Summary Table
# ----------------------------
if 'Close' in data.columns.get_level_values(0):
    close_data = data['Close']
    first_prices = close_data.iloc[0]
    last_prices  = close_data.iloc[-1]

    summary_data = []
    total_gain = 0

    for ticker in tickers:
        if ticker in close_data.columns:
            gain = (last_prices[ticker] - first_prices[ticker]) * shares[ticker]
            total_gain += gain
            summary_data.append({
                "Ticker": ticker,
                "Shares": shares[ticker],
                "First Price": first_prices[ticker],
                "Last Price": last_prices[ticker],
                "Money Gain": gain
            })

    # Append a total row
    summary_data.append({
        "Ticker": "Total",
        "Shares": "",
        "First Price": "",
        "Last Price": "",
        "Money Gain": total_gain
    })

    summary_df = pd.DataFrame(summary_data)
else:
    print("No 'Close' price data found.")
    summary_df = pd.DataFrame()

# ----------------------------
# Step 5: Write the Summary and Prices to an Excel File
# ----------------------------
excel_file = "stock_summary.xlsx"

with pd.ExcelWriter(excel_file, engine="openpyxl") as writer:
    # Write the summary table to a sheet named "MoneyGain"
    summary_df.to_excel(writer, index=False, sheet_name="MoneyGain")
    
    # Write the detailed closing prices over the time period to "ClosingPrices"
    if 'Close' in data.columns.get_level_values(0):
        close_data.to_excel(writer, sheet_name="ClosingPrices")

print(f"Summary and detailed closing prices written to '{excel_file}'")
