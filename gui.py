import io

import requests

from main import StockScraper, ExcelLogger
import sys
import os
import pandas as pd
import plotly.graph_objs as go
from PyQt5.QtWidgets import (QApplication, QMainWindow, QVBoxLayout,
                             QHBoxLayout, QPushButton, QWidget,
                             QComboBox, QLabel, QMessageBox, QProgressDialog)
from PyQt5.QtWebEngineWidgets import QWebEngineView
from PyQt5.QtCore import QUrl, QThread, pyqtSignal, Qt
from PyQt5.QtGui import QPixmap
from datetime import datetime
import pytz


stocks = ["Tesla", "Apple", "Nvidia", "Google", "Nike", "Manchester"]
file = "stock_data.xlsx"

# Company logos - simplified mapping
company_logos = {
    "Tesla": "https://logo.clearbit.com/tesla.com",
    "Apple": "https://logo.clearbit.com/apple.com",
    "Nvidia": "https://logo.clearbit.com/nvidia.com",
    "Google": "https://logo.clearbit.com/google.com",
    "Nike": "https://logo.clearbit.com/nike.com",
    "Manchester": "https://logo.clearbit.com/manutd.com"}


class ScrapeAllWorker(QThread):
    """Worker Thread to handle scraping all stocks"""
    progress_signal = pyqtSignal(str)
    finished_signal = pyqtSignal

    def run(self):
        for stock in stocks:
            self.progress_signal.emit(f"Collecting {stock.upper()} stock data...")
            try:
                scraper=StockScraper()
                price, price_text, timestamp = scraper.get_price(stock)
                scraper.close()
                ExcelLogger.save(stock, price, price_text, timestamp, file)
                self.progress_signal.emit(f"Saved {stock.upper()} - Price: {price_text}")
            except Exception as e:
                self.progress_signal.emit(f"Error with {stock}: {str(e)}")

        self.progress_signal.emit("All stocks successfully collected!")
        self.finished_signal.emit()


class StockDataVisualizer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Stock Price Tracker")
        self.setGeometry(100,100,1200,800)

        # Main widget and layout
        main_widget = QWidget()
        main_layout = QVBoxLayout()
        main_widget.setLayout(main_layout)
        self.setCentralWidget(main_widget)

        # Top control Layout
        control_layout = QHBoxLayout()
        main_layout.addLayout(control_layout)

        # Stock selection dropdown
        self.stock_selector = QComboBox()
        self.stock_selector.addItems(stocks)
        control_layout.addWidget(QLabel("Select Stock:"))
        control_layout.addWidget(self.stock_selector)

        # Fetch button
        fetch_button = QPushButton("Fetch and Plot Stock Price")
        fetch_button.clicked.connect(self.fetch_and_plot_stock)
        control_layout.addWidget(fetch_button)

        # Scrape all button
        scrape_all_button = QPushButton("Collect data for all stocks")
        scrape_all_button.clicked.connect(self.scrape_all_stocks)
        control_layout.addWidget(scrape_all_button)

        # Close button
        close_button = QPushButton("Close App")
        close_button.clicked.connect(self.close)
        close_button.setStyleSheet("background-color: #ff6b6b;")
        control_layout.addWidget(close_button)

        # Logo display
        logo_layout = QHBoxLayout()
        main_layout.addLayout(logo_layout)
        self.logo_label = QLabel("Company Logo")
        self.logo_label.setAlignment(Qt.AlignCenter)
        self.logo_label.setMinimumHeight(50)
        logo_layout.addWidget(self.logo_label)

        # Plotly WebEngine View for rendering graphs
        self.web_view = QWebEngineView()
        main_layout.addWidget(self.web_view)

        # Status Message
        self.status_label = QLabel("Ready")
        self.status_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(self.status_label)

        # Initialize with first stock
        self.update_logo(stocks[0])

    def update_logo(self, stock_name):
        """Update the company logo displayed"""
        try:
            url = company_logos[stock_name]
            response = requests.get(url)
            if response.status_code == 200:
                logo_data = io.BytesIO(response._content)
                pixmap = QPixmap()
                pixmap.loadFromData(logo_data.getvalue())
                scaled_pixmap = pixmap.scaled(100,100, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                self.logo_label.setPixmap(scaled_pixmap)
                return

            # If logo not found or error occurs, clear logo
            self.logo_label.setText(f"{stock_name.upper()} Logo")
        except Exception as e:
            print(f"Error loading logo: {e}")

    def fetch_and_plot_stock(self):
        stock_name = self.stock_selector.currentText()
        self.status_label.setText(f"Collecting {stock_name.upper()} data...")
        # Update UI
        QApplication.processEvents()

        # Update the logo
        self.update_logo(stock_name)

        # Fetch stock data
        try:
            scraper = StockScraper()
            price, price_text, timestamp = scraper.get_price(stock_name)
            scraper.close()

            # Save to Excel
            EXCEL_FILE = file
            ExcelLogger.save(stock_name,price,price_text,timestamp,EXCEL_FILE)

            # Create visualisation
            self.visualize_stock_data(stock_name, EXCEL_FILE)
            self.status_label.setText(f"{stock_name.upper()} data collected and plotted successfully! Price: {price_text}")

        except Exception as e:
            self.status_label.setText(f"Error: {str(e)}")
            QMessageBox.critical(self, "Error", f"Error collecting data: {str(e)}")

    def scrape_all_stocks(self):
        """Handle Scraping all stocks with progress updates"""

        # Create a progress dialog
        progress = QProgressDialog("Scraping stocks...", "Cancel", 0,0,self)
        progress.setWindowTitle("Scraping Progress")
        progress.setWindowModality(Qt.WindowModal)
        progress.setMinimumDuration(0)
        progress.setValue(0)
        progress.setAutoClose(False)
        progress.show()

        # Create and start worker thread
        self.worker = ScrapeAllWorker()

        # Connect signals
        self.worker.progress_signal.connect(progress.setLabelText)
        self.worker.progress_signal.connect(self.status_label.setText)
        self.worker.finished_signal.connect(progress.close)

        # Start the worker
        self.worker.start()

        # Show confirmation when finished
        self.worker.finished_signal.connect(
            lambda: QMessageBox.information(self, "Complete", "All stocks scraped successfully!"))


    def visualize_stock_data(self, stock_name, file_path):
        try:
            # Read the specific stock sheet
            df= pd.read_excel(file_path, sheet_name=stock_name)

            # Convert 'Date' and 'Time' columns to datetime
            df['DateTime'] = pd.to_datetime(df['Date'] + ' ' + df['Time'])

            # Sort by datetime to ensure chronological order
            df = df.sort_values('DateTime')

            # Create Plotly figure
            fig = go.Figure(data=[go.Scatter(
                x=df['DateTime'],
                y=df['Price Float'],
                mode='lines+markers',
                name=stock_name.upper()
            )])

            # Customize Layout
            fig.update_layout(
                title = f"{stock_name.upper()} Stock Price Over Time",
                xaxis_title = "Date and Time",
                yaxis_title = 'Stock Price ($)',
                template = 'plotly_white'
            )

            # Save to HTML and Load in WebView
            fig.write_html("stock_plot.html")
            plot_path = os.path.abspath("stock_plot.html")

            # Convert file path to QUrl
            plot_url = QUrl.fromLocalFile(plot_path)
            self.web_view.load(plot_url)

        except Exception as e:
            print(f'Error visualizing data: {e}')



def main():
    app = QApplication(sys.argv)
    main_window = StockDataVisualizer()
    main_window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()