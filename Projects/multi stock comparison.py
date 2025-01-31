import pandas as pd
import yfinance as yf
import matplotlib.pyplot as plt
import openpyxl
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.drawing.image import Image
from openpyxl.styles import Font, Alignment, NamedStyle

# Define stock parameters
tickers = ["AAPL", "TSLA", "GOOGL"]  # List of stocks to analyze
start_date = "2014-01-01"
end_date = "2024-01-01"

# Function to fetch stock data
def fetch_stock_data(ticker, start_date, end_date):
    stock = yf.download(ticker, start=start_date, end=end_date)
    return stock

# Dictionary to store stock data for comparison
comparison_data = {}

for ticker in tickers:
    try:
        print(f"Processing {ticker}...")
        
        # Fetch stock data
        stock_data = fetch_stock_data(ticker, start_date, end_date)
        market_data = fetch_stock_data("^GSPC", start_date, end_date)
        
        if stock_data.empty:
            print(f"⚠️ No data found for {ticker}. Skipping...")
            continue  # Skip to the next stock

        # Flatten MultiIndex columns if they exist
        if isinstance(stock_data.columns, pd.MultiIndex):
            stock_data.columns = [col[0] for col in stock_data.columns]
        if isinstance(market_data.columns, pd.MultiIndex):
            market_data.columns = [col[0] for col in market_data.columns]

        # Ensure we use the correct column name
        price_column = "Adj Close" if "Adj Close" in stock_data.columns else "Close"

        # Calculate Daily Returns
        stock_data["Daily Return"] = stock_data[price_column].pct_change()
        market_data["Daily Return"] = market_data[price_column].pct_change()

        # Compute Moving Averages
        stock_data["50-Day MA"] = stock_data[price_column].rolling(window=50).mean()
        stock_data["200-Day MA"] = stock_data[price_column].rolling(window=200).mean()

        # Calculate Volatility (30-day rolling standard deviation of daily returns)
        stock_data["Volatility"] = stock_data["Daily Return"].rolling(window=30).std()

        # Store stock data for comparison later
        comparison_data[ticker] = stock_data[[price_column, "Daily Return", "50-Day MA", "200-Day MA", "Volatility"]]

        # Export Data to Excel (individual stock analysis)
        excel_filename = f"{ticker}_Financial_Analysis.xlsx"
        writer = pd.ExcelWriter(excel_filename, engine="openpyxl")

        # Save stock data and calculations to different sheets
        stock_data.to_excel(writer, sheet_name="Stock Data")
        analysis_columns = ["Daily Return", "50-Day MA", "200-Day MA", "Volatility"]
        stock_data[analysis_columns].to_excel(writer, sheet_name="Analysis")

        # Save the workbook
        writer.close()
        print(f"Excel file '{excel_filename}' successfully created for {ticker}!")
    
    except Exception as e:
        print(f"❌ Error processing {ticker}: {e}")
        continue

# Generate Multi-Stock Comparison Excel File
comparison_filename = "Multi_Stock_Comparison.xlsx"
with pd.ExcelWriter(comparison_filename, engine="openpyxl") as writer:
    for ticker, data in comparison_data.items():
        data.to_excel(writer, sheet_name=ticker)  # Each stock gets its own sheet

# Generate Comparison Graphs
plt.figure(figsize=(12,6))
for ticker in comparison_data:
    plt.plot(comparison_data[ticker].index, comparison_data[ticker]["Close"], label=ticker)
plt.legend()
plt.title("Stock Price Comparison")
plt.xlabel("Date")
plt.ylabel("Stock Price")
plt.savefig("stock_price_comparison.png")
plt.close()

plt.figure(figsize=(12,6))
for ticker in comparison_data:
    plt.plot(comparison_data[ticker].index, comparison_data[ticker]["Volatility"], label=ticker)
plt.legend()
plt.title("Stock Volatility Comparison")
plt.xlabel("Date")
plt.ylabel("Volatility")
plt.savefig("volatility_comparison.png")
plt.close()

# Embed Comparison Graphs into the Multi-Stock Excel file
wb = openpyxl.load_workbook(comparison_filename)
chart_sheet = wb.create_sheet(title="Comparison Charts")
chart_images = ["stock_price_comparison.png", "volatility_comparison.png"]

for i, img_filename in enumerate(chart_images):
    img = Image(img_filename)
    chart_sheet.add_image(img, f"A{i*20 + 1}")  # Position images vertically

# Save final comparison Excel file
wb.save(comparison_filename)
print(f"Excel file '{comparison_filename}' successfully created with comparison graphs!")
