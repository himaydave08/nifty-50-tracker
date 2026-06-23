import os
import sys
import argparse
import datetime
import time
import requests
import io
import pandas as pd

# --- CONFIGURATION ---
TIMEZONE_IST = "Asia/Kolkata"

# NSE Trading Holidays for 2026 (Excludes Weekends)
NSE_HOLIDAYS_2026 = {
    "2026-01-26",  # Republic Day
    "2026-03-03",  # Holi
    "2026-03-26",  # Shri Ram Navami
    "2026-03-31",  # Shri Mahavir Jayanti
    "2026-04-03",  # Good Friday
    "2026-04-14",  # Dr. Baba Saheb Ambedkar Jayanti
    "2026-05-01",  # Maharashtra Day
    "2026-05-28",  # Bakri Id
    "2026-06-26",  # Muharram
    "2026-09-14",  # Ganesh Chaturthi
    "2026-10-02",  # Mahatma Gandhi Jayanti
    "2026-10-20",  # Dussehra
    "2026-11-10",  # Diwali-Balipratipada
    "2026-11-24",  # Sri Guru Nanak Dev
    "2026-12-25",  # Christmas
}

# Checkpoint Configurations (IST Time)
# Maps checkpoint name to its corresponding yfinance candle and field
CHECKPOINT_CONFIG = [
    # (checkpoint_name, candle_time_str, price_field)
    ("09:15", "09:15:00", "Open"),
    ("09:45", "09:30:00", "Close"),
    ("10:15", "10:00:00", "Close"),
    ("10:45", "10:30:00", "Close"),
    ("11:15", "11:00:00", "Close"),
    ("11:45", "11:30:00", "Close"),
    ("12:15", "12:00:00", "Close"),
    ("12:45", "12:30:00", "Close"),
    ("13:15", "13:00:00", "Close"),
    ("13:45", "13:30:00", "Close"),
    ("14:15", "14:00:00", "Close"),
    ("14:45", "14:30:00", "Close"),
    ("15:15", "15:00:00", "Close"),
    ("15:30", "15:15:00", "Close"),
]

# Checkpoint Names
CHECKPOINT_NAMES = [c[0] for c in CHECKPOINT_CONFIG]

# Helper: Convert checkpoint name to 1-based column index
# Symbol is Col 1, Company Name is Col 2, 09:15 is Col 3, etc.
def checkpoint_col_idx(checkpoint_name):
    try:
        idx = CHECKPOINT_NAMES.index(checkpoint_name)
        return 3 + idx
    except ValueError:
        return -1

# Fallback Nifty 50 List in case official download fails
FALLBACK_NIFTY50 = [
    ("ADANIENT.NS", "Adani Enterprises Ltd."),
    ("ADANIPORTS.NS", "Adani Ports and Special Economic Zone Ltd."),
    ("APOLLOHOSP.NS", "Apollo Hospitals Enterprise Ltd."),
    ("ASIANPAINT.NS", "Asian Paints Ltd."),
    ("AXISBANK.NS", "Axis Bank Ltd."),
    ("BAJAJ-AUTO.NS", "Bajaj Auto Ltd."),
    ("BAJFINANCE.NS", "Bajaj Finance Ltd."),
    ("BAJAJFINSV.NS", "Bajaj Finserv Ltd."),
    ("BEL.NS", "Bharat Electronics Ltd."),
    ("BPCL.NS", "Bharat Petroleum Corporation Ltd."),
    ("BHARTIARTL.NS", "Bharti Airtel Ltd."),
    ("BRITANNIA.NS", "Britannia Industries Ltd."),
    ("CIPLA.NS", "Cipla Ltd."),
    ("COALINDIA.NS", "Coal India Ltd."),
    ("DIVISLAB.NS", "Divi's Laboratories Ltd."),
    ("DRREDDY.NS", "Dr. Reddy's Laboratories Ltd."),
    ("EICHERMOT.NS", "Eicher Motors Ltd."),
    ("GRASIM.NS", "Grasim Industries Ltd."),
    ("HCLTECH.NS", "HCL Technologies Ltd."),
    ("HDFCBANK.NS", "HDFC Bank Ltd."),
    ("HDFCLIFE.NS", "HDFC Life Insurance Company Ltd."),
    ("HEROMOTOCO.NS", "Hero MotoCorp Ltd."),
    ("HINDALCO.NS", "Hindalco Industries Ltd."),
    ("HINDUNILVR.NS", "Hindustan Unilever Ltd."),
    ("ICICIBANK.NS", "ICICI Bank Ltd."),
    ("ITC.NS", "ITC Ltd."),
    ("INDUSINDBK.NS", "IndusInd Bank Ltd."),
    ("INFY.NS", "Infosys Ltd."),
    ("JSWSTEEL.NS", "JSW Steel Ltd."),
    ("KOTAKBANK.NS", "Kotak Mahindra Bank Ltd."),
    ("LT.NS", "Larsen & Toubro Ltd."),
    ("LTM.NS", "LTM Limited"),
    ("M&M.NS", "Mahindra & Mahindra Ltd."),
    ("MARUTI.NS", "Maruti Suzuki India Ltd."),
    ("NTPC.NS", "NTPC Ltd."),
    ("NESTLEIND.NS", "Nestle India Ltd."),
    ("ONGC.NS", "Oil & Natural Gas Corporation Ltd."),
    ("POWERGRID.NS", "Power Grid Corporation of India Ltd."),
    ("RELIANCE.NS", "Reliance Industries Ltd."),
    ("SBILIFE.NS", "SBI Life Insurance Company Ltd."),
    ("SHRIRAMFIN.NS", "Shriram Finance Ltd."),
    ("SBIN.NS", "State Bank of India"),
    ("SUNPHARMA.NS", "Sun Pharmaceutical Industries Ltd."),
    ("TCS.NS", "Tata Consultancy Services Ltd."),
    ("TATACONSUM.NS", "Tata Consumer Products Ltd."),
    ("TMCV.NS", "Tata Motors Ltd. (CV)"),
    ("TATASTEEL.NS", "Tata Steel Ltd."),
    ("TECHM.NS", "Tech Mahindra Ltd."),
    ("TITAN.NS", "Titan Company Ltd."),
    ("ULTRACEMCO.NS", "UltraTech Cement Ltd."),
    ("WIPRO.NS", "Wipro Ltd."),
    ("TRENT.NS", "Trent Ltd.")
]

# Ensure exactly 50 unique items in fallback (filtering duplicates if any)
seen_tickers = set()
UNIQUE_FALLBACK = []
for ticker, name in FALLBACK_NIFTY50:
    if ticker not in seen_tickers:
        seen_tickers.add(ticker)
        UNIQUE_FALLBACK.append((ticker, name))
UNIQUE_FALLBACK = UNIQUE_FALLBACK[:50]

# --- CORE FUNCTIONS ---

def is_trading_day(date_obj):
    """Checks if a date is a trading day (no weekends, no NSE holidays)."""
    # 5 is Saturday, 6 is Sunday
    if date_obj.weekday() in (5, 6):
        return False
    date_str = date_obj.strftime("%Y-%m-%d")
    if date_str in NSE_HOLIDAYS_2026:
        return False
    return True

def fetch_nifty50_constituents():
    """Downloads current Nifty 50 constituents from NSE website. Falls back on failure."""
    url = "https://archives.nseindia.com/content/indices/ind_nifty50list.csv"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    }
    try:
        print("Attempting to fetch NIFTY 50 constituents from NSE...")
        response = requests.get(url, headers=headers, timeout=10)
        if response.status_code == 200:
            df = pd.read_csv(io.StringIO(response.content.decode('utf-8')))
            constituents = []
            for _, row in df.iterrows():
                symbol = str(row['Symbol']).strip()
                name = str(row['Company Name']).strip()
                if symbol and name:
                    constituents.append((f"{symbol}.NS", name))
            if len(constituents) >= 45: # Verify we got a reasonable number
                print(f"Successfully fetched {len(constituents)} constituents from NSE.")
                return constituents[:50]
    except Exception as e:
        print(f"Warning: Failed to download constituents list: {e}")
    
    print("Using pre-configured fallback NIFTY 50 stock list.")
    return UNIQUE_FALLBACK

CONFIG_FILE = "config.json"

def load_config():
    """Loads Google Sheet Web App URL from config.json."""
    if not os.path.exists(CONFIG_FILE):
        print(f"Error: Configuration file '{CONFIG_FILE}' not found.")
        print("Please create a 'config.json' file in the project folder with this format:")
        print('{\n  "google_sheet_url": "https://script.google.com/macros/s/.../exec"\n}')
        sys.exit(1)
    try:
        with open(CONFIG_FILE, "r") as f:
            import json
            cfg = json.load(f)
            url = cfg.get("google_sheet_url")
            if not url or "macros/s/" not in url:
                print("Error: Invalid or missing 'google_sheet_url' in config.json.")
                sys.exit(1)
            return url
    except Exception as e:
        print(f"Error reading config.json: {e}")
        sys.exit(1)

def read_stock_list_from_sheet(google_sheet_url):
    """Retrieves stock symbols and company names from Google Sheet config."""
    try:
        print("Retrieving NIFTY 50 stock list from Google Sheet...")
        r = requests.get(google_sheet_url, timeout=15)
        if r.status_code == 200:
            res = r.json()
            if res.get("status") == "success":
                stocks = [(s["ticker"], s["name"]) for s in res.get("stocks", [])]
                if stocks:
                    print(f"Successfully retrieved {len(stocks)} stocks from Google Sheet.")
                    return stocks
    except Exception as e:
        print(f"Warning: Failed to fetch stock list from Google Sheets: {e}")
    
    print("Using pre-configured fallback stock list.")
    return UNIQUE_FALLBACK

def get_ticker_series(df, ticker, field):
    """Helper to retrieve series from MultiIndex df with swapped levels."""
    if not isinstance(df.columns, pd.MultiIndex):
        if field in df.columns:
            return df[field]
        return None
    try:
        return df[(ticker, field)]
    except KeyError:
        pass
    try:
        return df[(field, ticker)]
    except KeyError:
        pass
    return None

def find_price_robustly(df, ticker, target_dt, field, tolerance_minutes=10):
    """Finds price in a DataFrame close to the target datetime using robust indexing."""
    series = get_ticker_series(df, ticker, field)
    if series is None or series.empty:
        return None
        
    deltas = [abs(idx - target_dt) for idx in series.index]
    min_delta = min(deltas) if deltas else datetime.timedelta(days=1)
    if min_delta <= datetime.timedelta(minutes=tolerance_minutes):
        closest_idx = series.index[deltas.index(min_delta)]
        price = series.loc[closest_idx]
        if pd.notna(price):
            return float(price)
    return None

def fetch_and_upload_prices(google_sheet_url, target_date, tickers):
    """Downloads yfinance data and POSTs it to the Google Sheet Web App."""
    import pytz
    tz = pytz.timezone(TIMEZONE_IST)
    
    start_date_str = target_date.strftime("%Y-%m-%d")
    end_date_str = (target_date + datetime.timedelta(days=1)).strftime("%Y-%m-%d")
    
    print(f"Fetching intraday 15m data from yfinance for {start_date_str}...")
    import yfinance as yf
    
    session = requests.Session()
    session.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    })
    
    try:
        df = yf.download(tickers, start=start_date_str, end=end_date_str, interval="15m", group_by="ticker", session=session)
    except Exception as e:
        print(f"Error downloading data: {e}")
        return False
        
    if df.empty:
        print(f"No price data available for {start_date_str}. (Market might not have opened, or it's a holiday/weekend).")
        return False
        
    print(f"Successfully downloaded yfinance data. Rows: {len(df)}")
    
    # Extract prices for the payload
    payload = {
        "date": target_date.strftime("%d-%b-%Y"),
        "data": {}
    }
    
    cell_updates = 0
    for ticker in tickers:
        stock_prices = {}
        for checkpoint_name, candle_time_str, field in CHECKPOINT_CONFIG:
            target_dt_str = f"{start_date_str} {candle_time_str}"
            target_dt = pd.to_datetime(target_dt_str).tz_localize(tz)
            
            price = find_price_robustly(df, ticker, target_dt, field)
            if price is not None:
                stock_prices[checkpoint_name] = price
                cell_updates += 1
        if stock_prices:
            payload["data"][ticker] = stock_prices
            
    if cell_updates == 0:
        print("No new checkpoint prices were found to write.")
        return False
        
    print(f"Uploading {cell_updates} price values to Google Sheets...")
    try:
        r = requests.post(google_sheet_url, json=payload, headers={"Content-Type": "application/json"}, timeout=30)
        if r.status_code == 200:
            res = r.json()
            if res.get("status") == "success":
                print(f"Success! Google Sheet updated. Cells changed: {res.get('updatedCells', 0)}")
                return True
            else:
                print(f"Error from Google Web App: {res.get('message')}")
        else:
            print(f"HTTP Error posting to Google Sheets (Status: {r.status_code}): {r.text}")
    except Exception as e:
        print(f"Failed to connect to Google Sheets Web App: {e}")
        
    return False

# --- DAEMON LOOP IMPLEMENTATION ---

def get_next_checkpoint_time(current_time):
    """Given current local datetime (IST), returns the next checkpoint time and name."""
    for checkpoint_name, _, _ in CHECKPOINT_CONFIG:
        h, m = map(int, checkpoint_name.split(":"))
        checkpoint_dt = current_time.replace(hour=h, minute=m, second=0, microsecond=0)
        if checkpoint_dt > current_time:
            return checkpoint_dt, checkpoint_name
    return None, None

def run_tracker(target_date, force_daemon=False):
    """Main tracker function. Connects to Google Sheet Web App."""
    import pytz
    tz = pytz.timezone(TIMEZONE_IST)
    
    print(f"Checking trading day status for: {target_date.strftime('%Y-%m-%d')}")
    if not is_trading_day(target_date):
        print(f"Date {target_date.strftime('%Y-%m-%d')} is not an NSE trading day. Skipping execution.")
        return
        
    print("Loading Google Sheets configuration...")
    google_sheet_url = load_config()
    
    stocks = read_stock_list_from_sheet(google_sheet_url)
    tickers = [s[0] for s in stocks]
    
    # 1. Backfill all checkpoints passed so far today
    print("Executing catch-up/backfill for passed checkpoints...")
    fetch_and_upload_prices(google_sheet_url, target_date, tickers)
    
    # 2. Check if we should enter daemon mode
    now_ist = datetime.datetime.now(tz)
    market_close_ist = now_ist.replace(hour=15, minute=30, second=0, microsecond=0)
    
    is_today = (now_ist.date() == target_date)
    should_loop = is_today and (now_ist < market_close_ist + datetime.timedelta(minutes=5))
    
    if force_daemon:
        should_loop = True
        
    if not should_loop:
        print("Market is closed or we are running historical backfill. Tracker execution finished.")
        return
        
    print("\n--- Entering Real-Time Daemon Mode ---")
    print(f"Current local time: {now_ist.strftime('%I:%M:%S %p')} IST")
    
    while True:
        now_ist = datetime.datetime.now(tz)
        if now_ist >= market_close_ist + datetime.timedelta(minutes=5):
            print("Market has closed and post-close buffer completed. Exiting daemon mode.")
            break
            
        next_dt, next_name = get_next_checkpoint_time(now_ist)
        if next_dt is None:
            print("No more checkpoints remaining for today. Exiting daemon mode.")
            break
            
        sleep_secs = (next_dt - now_ist).total_seconds()
        sleep_secs += 30  # 30-sec buffer for yfinance lag
        
        print(f"Next checkpoint: {next_name} at {next_dt.strftime('%I:%M %p')}.")
        print(f"Sleeping for {int(sleep_secs // 60)}m {int(sleep_secs % 60)}s...")
        
        try:
            time.sleep(sleep_secs)
        except KeyboardInterrupt:
            print("\nDaemon loop interrupted by user. Exiting.")
            break
            
        print(f"\nWaking up at {datetime.datetime.now(tz).strftime('%I:%M:%S %p')} IST...")
        print(f"Triggering update for checkpoint: {next_name}")
        
        retries = 5
        success = False
        for attempt in range(1, retries + 1):
            print(f"Upload attempt {attempt}/{retries}...")
            success = fetch_and_upload_prices(google_sheet_url, target_date, tickers)
            if success:
                break
            print("Waiting 30 seconds before retry...")
            time.sleep(30)
            
        if not success:
            print(f"Warning: Failed to upload price for checkpoint {next_name} after {retries} attempts.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="NIFTY 50 Live Price Tracker (Excel)")
    parser.add_argument("--date", help="Target date in DD-MMM-YYYY format (e.g. 19-Jun-2026)")
    parser.add_argument("--daemon", action="store_true", help="Force daemon mode looping")
    
    args = parser.parse_args()
    
    import pytz
    tz = pytz.timezone(TIMEZONE_IST)
    
    if args.date:
        try:
            target_date = datetime.datetime.strptime(args.date, "%d-%b-%Y").date()
        except ValueError:
            print("Error: Date must be in DD-MMM-YYYY format (e.g., 19-Jun-2026)")
            sys.exit(1)
    else:
        target_date = datetime.datetime.now(tz).date()
        
    run_tracker(target_date, force_daemon=args.daemon)
