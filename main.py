from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import os
import time
from datetime import datetime
from webdriver_manager.chrome import ChromeDriverManager
import warnings
warnings.simplefilter(action='ignore', category=UserWarning)

# Install required packages
import subprocess
import sys
import importlib


def install_package(package):
    print(f"Installing {package}...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", package])
    print(f"{package} installed successfully")


# Check and install required packages
required_packages = ["pandas", "openpyxl"]
for package in required_packages:
    try:
        importlib.import_module(package)
        print(f"{package} is already installed")
    except ImportError:
        install_package(package)

# Import pandas after ensuring it's installed
import pandas as pd

URL_TEMPLATE = "https://digital.fidelity.com/search/main?q={}&ccsource=ss"
EXCEL_FILE = "stock_data.xlsx"

stocks =["tesla", "apple", "nvidia","Manchester", "google", "nike" ]


class WebAutomation:

    def get_data(self,stock_name):
        price = None
        price_text = "N/A"
        timestamp = datetime.now()
        url = URL_TEMPLATE.format(stock_name)

        # Define driver options
        chrome_options = Options()
        chrome_options.add_argument("--disable-search-engine-choice-screen")

        # Set download path to current directory
        download_path = os.getcwd()
        prefs = {'download.default_directory': download_path}
        chrome_options.add_experimental_option('prefs', prefs)

        # Initialize the driver with ChromeDriverManager to handle driver version
        driver = webdriver.Chrome(
            service=Service(ChromeDriverManager().install()),
            options=chrome_options
        )

        try:
            # Navigate to URL

            driver.get(url)

            # Add a small delay to ensure page loads
            time.sleep(3)

            # Wait for price element to load (timeout after 10 seconds)
            wait = WebDriverWait(driver, 10)
            price_element = wait.until(
                EC.presence_of_element_located((By.CLASS_NAME, "price-font-weight"))
            )

            # Extract the price
            price_text = price_element.text
            print(f"{stock_name.title()} stock price: {price_text} captured at {timestamp}")

            # Convert price text to float (removing $ and commas)
            if price_text:
                clean_price = price_text.replace('$', '').replace(',', '')
                try:
                    price = float(clean_price)
                except ValueError:
                    print(f"Could not convert price '{price_text}' to float")

            # Save to Excel
            self.save_to_excel(stock_name,price, price_text, timestamp)
            print(f"Data for {stock_name} successfully saved in {EXCEL_FILE}")

        except Exception as e:
            print(f"An error occurred retrieving {stock_name.title()} stock price: {e}")


        finally:
            # Always close the driver when done
            driver.quit()


    def save_to_excel(self, stock_name, price_float, price_text, timestamp):
        """Save the stock price data to an Excel file with timestamp"""
        # Create a new dataframe with this data point
        new_data = pd.DataFrame({
            'Date': [timestamp.strftime('%Y-%m-%d')],
            'Time': [timestamp.strftime('%H:%M:%S')],
            'Price Text': [price_text],
            'Price Float': [price_float]
        })

        try:
            # Check if file exists
            if os.path.exists(EXCEL_FILE):
                with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl", mode="a", if_sheet_exists="overlay") as writer:
                    try:
                        # Read existing data
                        existing_data = pd.read_excel(EXCEL_FILE, sheet_name=stock_name)
                        # Append new data
                        updated_data = pd.concat([existing_data, new_data], ignore_index=True)
                    except ValueError:
                        # Create sheet if it doesn't exist yet
                        updated_data = new_data

                    updated_data.to_excel(writer, sheet_name=stock_name, index=False)

            else:
                with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
                    new_data.to_excel(writer, sheet_name=stock_name, index=False)

        except Exception as e:
            print(f"Error saving {stock_name} data: {e}")


# Create an instance and run
if __name__ == "__main__":
    bot = WebAutomation()
    for stock in stocks:
        bot.get_data(stock)

    print("Process Completed!")
