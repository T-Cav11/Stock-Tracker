from main import StockScraper, ExcelLogger
import sys
import os
import pandas as pd
import plotly.graph_objs as go
from PyQt5.QtWidgets import (QApplication, QMainWindow, QVBoxLayout,
                             QHBoxLayout, QPushButton, QWidget,
                             QComboBox, QLabel)
from PyQt5.QtWebEngineWidgets import QWebEngineView
from PyQt5.QtCore import QUrl, QThread
from datetime import datetime
import pytz


stocks = ["tesla", "apple", "nvidia", "google", "nike", "Manchester"]
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

        # Plotly WebEngine View for rendering graphs
        self.web_view = QWebEngineView()
        main_layout.addWidget(self.web_view)

    def fetch_and_plot_stock(self):
        stock_name = self.stock_selector.currentText()

        # Fetch stock data
        scraper = StockScraper()
        price, price_text, timestamp = scraper.get_price(stock_name)
        scraper.close()

        # Save to Excel
        EXCEL_FILE = file
        ExcelLogger.save(stock_name,price,price_text,timestamp,EXCEL_FILE)

        # Create visualisation
        self.visualize_stock_data(stock_name, EXCEL_FILE)

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

            # Conver file path to QUrl
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