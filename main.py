from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import os
import time as t
from datetime import datetime, time
from webdriver_manager.chrome import ChromeDriverManager
import pytz
import warnings
warnings.simplefilter(action='ignore', category=UserWarning)

import pandas as pd
URL_TEMPLATE = "https://digital.fidelity.com/search/main?q={}&ccsource=ss"


class DriverManager:
    @staticmethod
    def create_driver():
        # Define driver options
        chrome_options = Options()
        chrome_options.add_argument("--disable-search-engine-choice-screen")

        # Run chrome in background without pop-up window
        chrome_options.add_argument("--headless=new")
        chrome_options.add_argument("--window-size=1920,1080")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument(
            "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
        )

        # Set download path to current directory
        download_path = os.getcwd()
        prefs = {'download.default_directory': download_path}
        chrome_options.add_experimental_option('prefs', prefs)

        return webdriver.Chrome(
            service=Service(ChromeDriverManager().install()),
            options=chrome_options
        )


class StockScraper:
    def __init__(self):
        self.driver = DriverManager.create_driver()

    # Get the stock price from the URL
    def get_price(self,stock_name):
        price = None
        price_text = "N/A"
        timestamp = datetime.now()
        url = URL_TEMPLATE.format(stock_name)

        try:

            self.driver.get(url)
            WebDriverWait(self.driver, 15).until(
                lambda d: d.execute_script("return document.readyState") == "complete"
            )
            element = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, "price-font-weight"))
            )
            price_text = element.text
            price = float(price_text.replace('$', '').replace(',', ''))

            print(f"Stock data for {stock_name} is {price_text} at {timestamp}")
        except Exception as e:
            print(f"Error fetching price for {stock_name}: {e}")
        return price, price_text, timestamp

    # Close the driver
    def close(self):
        self.driver.quit()


class ExcelLogger:

    # Check if the market is open
    @staticmethod
    def is_market_open():
        market_operating_timezone = pytz.timezone("US/Eastern")
        now_et = datetime.now(market_operating_timezone).time()
        market_open = time(9, 30)
        market_close = time(16, 00)
        return market_open <= now_et <= market_close

    # Save the data to the excel sheet
    @staticmethod
    def save(stock_name, price_float, price_text, timestamp, file_path):
        """Save the stock price data to an Excel file with timestamp"""
        # Create a new dataframe with this data point
        note = "Market Open" if ExcelLogger.is_market_open() else "Market Closed or Stale Data"
        df = pd.DataFrame({
            'Date': [timestamp.strftime('%Y-%m-%d')],
            'Time': [timestamp.strftime('%H:%M:%S')],
            'Price Text': [price_text],
            'Price Float': [price_float],
            'Note': [note]
        })

        try:
            if os.path.exists(file_path):
                with pd.ExcelWriter(file_path, engine="openpyxl", mode="a", if_sheet_exists="overlay") as writer:
                    try:
                        existing = pd.read_excel(file_path, sheet_name=stock_name)
                        df = pd.concat([existing, df], ignore_index=True)
                    except ValueError:
                        pass  # Sheet doesn't exist, will be created
                    df.to_excel(writer, sheet_name=stock_name, index=False)
                    print(f"{stock_name} data saved to {file_path}")
            else:
                with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
                    df.to_excel(writer, sheet_name=stock_name, index=False)
                print(f"{stock_name} data saved to {file_path}")
        except Exception as e:
            print(f"Error saving data for {stock_name}: {e}")


def main():

    # main execution logic
    EXCEL_FILE = "stock_data.xlsx"
    stocks = ["tesla", "apple", "nvidia", "Manchester", "google", "nike"]

    for stock in stocks:
        scraper = StockScraper()
        price, price_text, timestamp = scraper.get_price(stock)
        ExcelLogger.save(stock, price, price_text, timestamp, EXCEL_FILE)
        scraper.close()

    print("✅ All stock data retrieved and saved.")

# Create an instance and run
if __name__ == "__main__":
  main()
