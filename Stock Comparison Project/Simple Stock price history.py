import pandas as pd
import yfinance as yf
import matplotlib.pyplot as plt
import openpyxl
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.drawing.image import Image
from openpyxl.styles import Font, Alignment, NamedStyle

# Function to fetch stock data
def fetch_stock_data(ticker, start_date, end_date):
    stock = yf.download(ticker, start=start_date, end=end_date)
    return stock

# Define stock parameters
ticker = "AAPL"  # Example: Apple Inc.
start_date = "2014-01-01"
end_date = "2024-01-01"

# Fetch stock data
stock_data = fetch_stock_data(ticker, start_date, end_date)
market_data = fetch_stock_data("^GSPC", start_date, end_date)

# Flatten MultiIndex columns if they exist
if isinstance(stock_data.columns, pd.MultiIndex):
    stock_data.columns = [col[0] for col in stock_data.columns]
if isinstance(market_data.columns, pd.MultiIndex):
    market_data.columns = [col[0] for col in market_data.columns]

# Display available columns
print("Available columns:", stock_data.columns)

# Ensure we use the correct column name
price_column = "Adj Close" if "Adj Close" in stock_data.columns else "Close"

# Display the first five rows
print(stock_data.head())

# Calculate Daily Returns
stock_data["Daily Return"] = stock_data[price_column].pct_change()
market_data["Daily Return"] = market_data[price_column].pct_change()

# Compute Moving Averages
stock_data["50-Day MA"] = stock_data[price_column].rolling(window=50).mean()
stock_data["200-Day MA"] = stock_data[price_column].rolling(window=200).mean()

# Calculate Volatility (30-day rolling standard deviation of daily returns)
stock_data["Volatility"] = stock_data["Daily Return"].rolling(window=30).std()

# Compute RSI (Relative Strength Index)
window_length = 14
delta = stock_data[price_column].diff()
gain = (delta.where(delta > 0, 0)).rolling(window=window_length).mean()
loss = (-delta.where(delta < 0, 0)).rolling(window=window_length).mean()
rs = gain / loss
stock_data["RSI"] = 100 - (100 / (1 + rs))

# Compute Bollinger Bands
rolling_mean = stock_data[price_column].rolling(window=20).mean()
rolling_std = stock_data[price_column].rolling(window=20).std()
stock_data["Upper Band"] = rolling_mean + (rolling_std * 2)
stock_data["Lower Band"] = rolling_mean - (rolling_std * 2)

# Compute Exponential Moving Averages
stock_data["20-Day EMA"] = stock_data[price_column].ewm(span=20, adjust=False).mean()
stock_data["50-Day EMA"] = stock_data[price_column].ewm(span=50, adjust=False).mean()

# Compute Maximum Drawdown
rolling_max = stock_data[price_column].cummax()
drawdown = (stock_data[price_column] - rolling_max) / rolling_max
stock_data["Max Drawdown"] = drawdown.rolling(window=252).min()

# Compute Beta
cov_matrix = stock_data["Daily Return"].cov(market_data["Daily Return"])
market_var = market_data["Daily Return"].var()
stock_data["Beta"] = cov_matrix / market_var

# Compute Sharpe Ratio
risk_free_rate = 0.02
stock_data["Excess Return"] = stock_data["Daily Return"] - (risk_free_rate / 252)
sharpe_ratio = stock_data["Excess Return"].mean() / stock_data["Excess Return"].std()
print(f"Sharpe Ratio: {sharpe_ratio}")

# Turn off interactive mode to prevent graphs from displaying
plt.ioff()

# Generate and save plots
charts = {
    "stock_price_moving_avg.png": stock_data[[price_column, "50-Day MA", "200-Day MA"]],
    "daily_returns_histogram.png": stock_data["Daily Return"],
    "rolling_volatility.png": stock_data["Volatility"]
}

for filename, data in charts.items():
    plt.figure(figsize=(12,6))
    if isinstance(data, pd.DataFrame):
        data.plot()
    else:
        plt.hist(data.dropna(), bins=50, alpha=0.75, color='blue')
    plt.title(filename.replace(".png", "").replace("_", " ").title())
    plt.savefig(filename)
    plt.close()

# Export Data to Excel
excel_filename = f"{ticker}_Financial_Analysis.xlsx"
writer = pd.ExcelWriter(excel_filename, engine="openpyxl")

# Save stock data and calculations to different sheets
stock_data.to_excel(writer, sheet_name="Stock Data")
analysis_columns = ["Daily Return", "50-Day MA", "200-Day MA", "Volatility", "RSI", "Upper Band", "Lower Band", "20-Day EMA", "50-Day EMA", "Max Drawdown", "Beta"]
stock_data[analysis_columns].to_excel(writer, sheet_name="Analysis")

# Save the workbook first so we can modify it
writer.close()
wb = openpyxl.load_workbook(excel_filename)
chart_sheet = wb.create_sheet(title="Charts")

# List of saved image files
chart_images = ["stock_price_moving_avg.png", "daily_returns_histogram.png", "rolling_volatility.png"]

for i, img_filename in enumerate(chart_images):
    img = Image(img_filename)
    chart_sheet.add_image(img, f"A{i*20 + 1}")  # Position images vertically

# Save final formatted Excel file
wb.save(excel_filename)

print(f"Excel file '{excel_filename}' successfully updated with graphs and formatted analysis!")
