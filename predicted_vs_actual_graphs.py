import yfinance as yf
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime

# eric: name your VBA file as PLTR.xlsm
stock_name = "NVDA"
# yfinance uses "^" to represent index
index_true = 0
year = 2024
# Download stock daily price data for year
start_date = datetime(year, 1, 1)
end_date = datetime(year, 12, 31)
stock = yf.Ticker("^"*index_true + stock_name)
stock_data = stock.history(start=start_date, end=end_date)

# Load Excel file
excel_file = pd.ExcelFile(stock_name + ".xlsm")
sheet_name = excel_file.sheet_names[0]
df = excel_file.parse(sheet_name)

# Extract date and predicted price columns
date_col = df[df.columns[ord("P") - ord("A")]]
predicted_price_col = df[df.columns[ord("Q") - ord("A")]]

plt.figure(figsize=(12, 6))
plt.plot(stock_data.index, stock_data["Close"], label="Actual Price")
plt.plot(date_col, predicted_price_col, label="Predicted Price")
plt.xlabel("Date")
plt.ylabel("Price")
plt.title(f"{stock_name} {year} Daily Price Comparison")
plt.legend()
plt.show()

# Calculate percentage change
actual_pct_change = stock_data["Close"].pct_change()
predicted_pct_change = predicted_price_col.pct_change()

# Plot the line graph
plt.figure(figsize=(12, 6))
plt.plot(stock_data.index, actual_pct_change, label="Actual Price % Change")
plt.plot(date_col, predicted_pct_change, label="Predicted Price % Change")
plt.xlabel("Date")
plt.ylabel("Percentage Change")
plt.title(f"{stock_name} {year} Daily Price Percentage Change Comparison")
plt.legend()
plt.show()

# Create a new column with 1 if both predicted and actual percentage changes are positive, 0 otherwise
combined_pct_change = pd.DataFrame(
    {
        "Date": actual_pct_change.index,
        "Actual": actual_pct_change.reset_index(drop=True),
        "Predicted": predicted_pct_change[: len(actual_pct_change)].reset_index(drop=True),
    }
)
combined_pct_change["Value"] = (
    (combined_pct_change["Actual"] > 0) & (combined_pct_change["Predicted"] > 0)
    | (combined_pct_change["Actual"] < 0) & (combined_pct_change["Predicted"] < 0)
).astype(int)

# Plot the line graph
plt.figure(figsize=(12, 6))
# plt.plot(combined_pct_change['Date'], combined_pct_change['Value'], label='Correct Prediction')
plt.scatter(
    combined_pct_change["Date"],
    combined_pct_change["Value"],
    label="Correct Prediction",
    s=10,
)
# plt.bar(combined_pct_change['Date'], combined_pct_change['Value'], label='Correct Prediction')
plt.xlabel("Date")
plt.ylabel("Value (1 or 0)")
plt.title(
    f"{stock_name} {year} Successful Prediction of Sign of Percentage Change ({round(combined_pct_change['Value'].sum() / len(combined_pct_change) * 100, 2)}%)"
)
plt.legend()
plt.show()

rolling_days = 30
# Calculate the percentage of correct predictions over time
percentage_correct = combined_pct_change["Value"].rolling(window=rolling_days).mean() * 100

# Plot the line graph
plt.figure(figsize=(12, 6))
plt.plot(
    percentage_correct.index,
    percentage_correct,
    label="Percentage of Correct Predictions",
)
plt.xlabel("Date")
plt.ylabel("Percentage (%)")
plt.title(f"{stock_name} {year} Percentage of Correct Predictions (over the last {rolling_days} days)")
plt.legend()
plt.show()

