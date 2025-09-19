import tkinter as tk
from tkinter import ttk, messagebox
import sqlite3
import logging
import os
import requests
from datetime import datetime

# Setup logging
USER_DATA_DIR = os.path.expanduser("~/PortfolioTracker")
if not os.path.exists(USER_DATA_DIR):
    os.makedirs(USER_DATA_DIR)
logging.basicConfig(
    filename=os.path.join(USER_DATA_DIR, 'portfolio.log'),
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger()

class PortfolioTrackerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Portfolio Tracker")
        self.root.geometry("1000x600")
        self.api_key = "YOUR_ALPHA_VANTAGE_API_KEY"  # Replace with your key

        # Styling
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TButton", font=("Helvetica", 10))
        style.configure("TLabel", font=("Helvetica", 10))

        # Create main frame
        self.main_frame = ttk.Frame(self.root)
        self.main_frame.pack(expand=True, fill="both", padx=5, pady=5)
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(0, weight=1)

        # Create treeview
        self.tree = ttk.Treeview(
            self.main_frame,
            columns=("Symbol", "Name", "Price", "Shares", "Value"),
            show="headings"
        )
        self.tree.heading("Symbol", text="Symbol")
        self.tree.heading("Name", text="Company Name")
        self.tree.heading("Price", text="Price ($)")
        self.tree.heading("Shares", text="Shares Owned")
        self.tree.heading("Value", text="Total Value ($)")
        self.tree.column("Symbol", width=100, anchor="center")
        self.tree.column("Name", width=200, anchor="w")
        self.tree.column("Price", width=100, anchor="center")
        self.tree.column("Shares", width=100, anchor="center")
        self.tree.column("Value", width=100, anchor="center")
        self.tree.grid(row=0, column=0, sticky="nsew")

        # Scrollbar
        scrollbar = ttk.Scrollbar(self.main_frame, orient="vertical", command=self.tree.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=scrollbar.set)

        # Add stock entry
        self.entry_frame = ttk.Frame(self.main_frame)
        self.entry_frame.grid(row=1, column=0, sticky="ew")
        ttk.Label(self.entry_frame, text="Symbol:").pack(side="left")
        self.symbol_entry = ttk.Entry(self.entry_frame, width=10)
        self.symbol_entry.pack(side="left", padx=5)
        ttk.Label(self.entry_frame, text="Shares:").pack(side="left")
        self.shares_entry = ttk.Entry(self.entry_frame, width=10)
        self.shares_entry.pack(side="left", padx=5)
        ttk.Button(self.entry_frame, text="Add Stock", command=self.add_stock).pack(side="left", padx=5)
        ttk.Button(self.entry_frame, text="Refresh Prices", command=self.refresh_prices).pack(side="left", padx=5)

        # Initialize database and load portfolio
        self.init_db()
        self.load_portfolio()

    def init_db(self):
        """Initialize SQLite database for portfolio."""
        conn = sqlite3.connect(os.path.join(USER_DATA_DIR, "portfolio.db"))
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS portfolio (
                symbol TEXT PRIMARY KEY,
                company_name TEXT,
                shares INTEGER,
                price REAL,
                last_updated TEXT
            )
        """)
        conn.commit()
        conn.close()

    def load_portfolio(self):
        """Load portfolio from database into treeview."""
        conn = sqlite3.connect(os.path.join(USER_DATA_DIR, "portfolio.db"))
        cursor = conn.cursor()
        cursor.execute("SELECT symbol, company_name, shares, price FROM portfolio")
        rows = cursor.fetchall()
        conn.close()

        for symbol, name, shares, price in rows:
            value = price * shares if price and shares else 0
            self.tree.insert("", "end", values=(symbol, name, f"${price:.2f}", shares, f"${value:.2f}"))
        logger.info(f"Loaded {len(rows)} stocks from portfolio")

    def add_stock(self):
        """Add a stock to the portfolio."""
        symbol = self.symbol_entry.get().strip().upper()
        shares = self.shares_entry.get().strip()
        if not symbol or not shares.isdigit():
            messagebox.showerror("Error", "Enter a valid symbol and number of shares")
            logger.error("Invalid input: Symbol and shares required")
            return

        shares = int(shares)
        price, name = self.fetch_stock_data(symbol)
        if price is None or name is None:
            messagebox.showerror("Error", f"Failed to fetch data for {symbol}")
            logger.error(f"Failed to fetch data for {symbol}")
            return

        # Save to database
        conn = sqlite3.connect(os.path.join(USER_DATA_DIR, "portfolio.db"))
        cursor = conn.cursor()
        cursor.execute("""
            INSERT OR REPLACE INTO portfolio (symbol, company_name, shares, price, last_updated)
            VALUES (?, ?, ?, ?, ?)
        """, (symbol, name, shares, price, datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
        conn.commit()
        conn.close()

        # Update treeview
        value = price * shares
        self.tree.insert("", "end", values=(symbol, name, f"${price:.2f}", shares, f"${value:.2f}"))
        logger.info(f"Added {symbol} with {shares} shares at ${price:.2f}")
        self.symbol_entry.delete(0, tk.END)
        self.shares_entry.delete(0, tk.END)

    def refresh_prices(self):
        """Refresh stock prices in the portfolio."""
        conn = sqlite3.connect(os.path.join(USER_DATA_DIR, "portfolio.db"))
        cursor = conn.cursor()
        cursor.execute("SELECT symbol, shares FROM portfolio")
        rows = cursor.fetchall()

        # Clear treeview
        for item in self.tree.get_children():
            self.tree.delete(item)

        # Update prices
        for symbol, shares in rows:
            price, name = self.fetch_stock_data(symbol)
            if price is None or name is None:
                logger.error(f"Failed to refresh data for {symbol}")
                continue
            cursor.execute("""
                UPDATE portfolio SET price = ?, company_name = ?, last_updated = ?
                WHERE symbol = ?
            """, (price, name, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), symbol))
            value = price * shares
            self.tree.insert("", "end", values=(symbol, name, f"${price:.2f}", shares, f"${value:.2f}"))
            logger.info(f"Refreshed {symbol}: ${price:.2f}")

        conn.commit()
        conn.close()
        messagebox.showinfo("Success", "Portfolio prices refreshed")

    def fetch_stock_data(self, symbol):
        """Fetch stock price and name from Alpha Vantage."""
        try:
            url = f"https://www.alphavantage.co/query?function=GLOBAL_QUOTE&symbol={symbol}&apikey={self.api_key}"
            response = requests.get(url)
            response.raise_for_status()
            data = response.json()
            if "Global Quote" in data and "05. price" in data["Global Quote"]:
                price = float(data["Global Quote"]["05. price"])
                name = data.get("Global Quote", {}).get("01. symbol", symbol)  # Fallback to symbol
                return price, name
            return None, None
        except Exception as e:
            logger.error(f"Error fetching data for {symbol}: {str(e)}")
            return None, None

if __name__ == "__main__":
    root = tk.Tk()
    app = PortfolioTrackerApp(root)
    root.mainloop()