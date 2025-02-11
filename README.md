# Stock Data Tracker

This Python script downloads and processes historical stock data for the tickers **NVDA**, **SNPS**, and **LOPE**. The data is retrieved for the period from **January 27, 2025** to **February 11, 2025** (using today's date when the assignment end date is in the future). The script then:

- Fetches historical data using the **yfinance** library.
- Generates a plot of the closing prices with **matplotlib** and saves it as a PNG image.
- Exports the full dataset (and optionally, the closing prices) to an Excel workbook using **pandas** and **openpyxl**.

## Features

- **Dynamic Date Handling:**  
  Uses the assignment-specified start date (2025-01-27) and an effective end date (today’s date, 2025-02-11) when today's date is before the specified end date (2025-02-16).

- **Data Retrieval:**  
  Downloads historical stock data for multiple tickers.

- **Data Visualization:**  
  Creates and saves a plot of closing prices to avoid blocking the script execution.

- **Excel Export:**  
  Writes the complete dataset to an Excel file (`stock_data.xlsx`) with multiple sheets:
  - A sheet for all historical data.
  - A separate sheet for closing prices.

## Requirements

- **Python 3.x**
- **yfinance**
- **pandas**
- **matplotlib**
- **openpyxl**

## Installation

1. **Clone or Download the Repository:**  
   Download the project files to your local machine.

2. **Install the Required Packages:**  
   Open a terminal or command prompt and run:
   ```bash
   pip install yfinance pandas matplotlib openpyxl
   ```

## Usage

### Run the Script:
From the terminal in the project directory, run:

```bash
python3 stocks.py
```

### Output Files:
- The stock data will be saved to an Excel file named `stock_data.xlsx`.
- The closing price plot will be saved as `closing_prices.png`.

### Console Output:
The script will print the downloaded data to the console and notify you when the plot and Excel file have been saved successfully.

## Code Overview

### Date Range Setup:
The script sets the start date as `2025-01-27` and calculates the effective end date as `2025-02-11` (today’s date) when today's date is earlier than the specified assignment end date (`2025-02-16`). An extra day is added to ensure that data for the effective end date is included.

### Data Download:
Uses `yfinance.download()` to retrieve historical stock data for **NVDA, SNPS, and LOPE**.

### Plotting:
Generates a plot of the closing prices using `matplotlib`, then saves it to `closing_prices.png` to avoid blocking the script with an interactive window.

### Excel Export:
Uses `pandas.ExcelWriter` with the `openpyxl` engine to write the full dataset to an Excel file, including a separate sheet for closing prices.

## Troubleshooting

### Slow Execution or Hanging:
If the script appears slow or hangs, ensure that you are not waiting on an interactive plot window. This version of the script saves the plot to a file instead of displaying it interactively.

### Missing Dependencies:
Verify that all required packages are installed. Use the provided `pip install` command if necessary.

### Data Issues:
Ensure that the specified date range is valid and that there is available stock data for the given period.

## License
This project is licensed under the MIT License.

## Contact
If you have any questions or need further assistance, please contact [Your Name or Email].
