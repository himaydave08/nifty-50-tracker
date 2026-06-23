# NIFTY 50 Live Price Tracker (Google Sheets Dashboard)

An automated serverless price tracker that logs prices for all 50 constituent stocks of the NIFTY 50 index throughout the trading day, directly onto a **Google Sheet**. The system automatically generates daily tabs, applies styles, dynamic formulas, conditional formatting, and sends email alerts.

Because the data is synced in the cloud, you can open the **Google Sheets mobile app** on your phone to track stock movements, verdicts, and receive email alerts live from anywhere!

---

## Features

- **Mobile Accessibility**: Open the Google Sheet on your smartphone or tablet to view live updates checkpoint-by-checkpoint.
- **Constituents Auto-Update**: Pulls the current Nifty 50 list and stores it in the `Stock List` configuration tab. You can add, edit, or remove tickers in the `Stock List` tab, and the script will automatically update its tracking list.
- **Weekend & Holiday Skipping**: Excludes weekends and NSE-scheduled trading holidays for 2026, exiting immediately to conserve resources.
- **Free Email Alerts (Gmail Integration)**:
  - **Instant Alerts**: Sends an email notification to your phone the *moment* a stock's verdict transitions from `Hold` to `Review` (turns RED) during market hours.
  - **End-of-Day Summary**: Sends a complete wrap-up report at 3:30 PM IST with final statistics (total Holds, Reviews, and a checklist of flagged stocks) and a link to the spreadsheet.
- **Fully Dynamic Spreadsheet Formulas**:
  - The stock **Verdict** (`Hold` / `Review`) is calculated using live Google Sheets formulas. If any checkpoint price falls below the previous checkpoint, it triggers `Review`; otherwise, it remains `Hold`.
  - The bottom **Daily Summary Block** uses dynamic formulas (`COUNTIF`, `COUNTA`) to count total Hold vs. Review signals.
- **Interactive Conditional Formatting**: Highlights cell backgrounds dynamically (Green = up, Red = down, Grey = neutral). If you manually override a price, cell backgrounds and verdicts update instantly.

---

## Google Sheet Structure

1. **`Stock List` Sheet**: A single source of truth containing stock symbols and names. The script reads tickers from this tab.
   - **Columns A & B**: Symbol & Company Name.
   - **Cell D1**: Header `Alert Email`.
   - **Cell D2**: Target recipient email address for alerts (defaults to your Google account email, but can be customized to any address).
2. **`DD-MMM-YYYY` Sheets**: Generated automatically for each trading day (e.g., `19-Jun-2026`). Columns include:
   - **Symbol & Company Name**
   - **9:15 AM (Open Price)**
   - **Checkpoints**: `09:45`, `10:15`, `10:45`, `11:15`, `11:45`, `12:15`, `12:45`, `13:15`, `13:45`, `14:15`, `14:45`, `15:15`
   - **15:30 PM (Close Price)**
   - **Greens**: Number of green checkpoints today (price increased).
   - **Reds**: Number of red checkpoints today (price decreased).
   - **Verdict**: Dynamic `Hold` (if Greens > Reds) / `Review` (if Reds >= Greens) / `No Data` (if empty).

---

## Setup Options

Choose one of the following methods to automate the tracker to run every trading day:

### Option A: 100% Serverless Cloud Tracking (Recommended - Easiest & PC-Off)
This method runs completely inside Google Sheets. You don't need to install Python, run any command lines, or set up any external servers.

1. **Paste Apps Script**:
   - Open a Google Sheet.
   - Click **Extensions** -> **Apps Script**.
   - Delete any default code, and copy-paste the entire contents of [apps_script.js](file:///c:/Users/Himay/OneDrive/Desktop/share%20market%20project/apps_script.js).
   - Click the **Save** icon (disk symbol).
2. **Set Up the Time Trigger**:
   - In the Apps Script editor, click the **Triggers** icon (clock symbol) on the left sidebar.
   - Click **+ Add Trigger** (bottom right).
   - Configure the trigger:
     - **Choose which function to run**: Select **`fetchNiftyPrices`**.
     - **Choose which deployment should run**: Select **Head**.
     - **Select event source**: Select **Time-driven**.
     - **Select type of time based trigger**: Select **Minutes timer**.
     - **Select minute interval**: Select **Every 30 minutes**.
     - Click **Save**.
   - *Grant authorizations: Click your account -> click **Advanced** -> click **Go to Untitled project (unsafe)** -> click **Allow**.*

*That's it! Google will now fetch Nifty prices and update your sheet automatically every 30 minutes during market hours, completely in the cloud!*

---

### Option B: Cloud Hosting on PythonAnywhere (Laptop-Off)
If you want to use the Python script instead of Apps Script triggers:

1. **Deploy Web App**: Paste [apps_script.js](file:///c:/Users/Himay/OneDrive/Desktop/share%20market%20project/apps_script.js) into your Apps Script, click **Deploy** -> **New deployment** -> **Web app**, set access to **Anyone**, and deploy. Copy the Web App URL.
2. **Configure config.json**: In the project folder, create `config.json`:
   ```json
   {
     "google_sheet_url": "YOUR_WEB_APP_URL"
   }
   ```
3. **Upload Files**: Create a free account on **[pythonanywhere.com](https://www.pythonanywhere.com)** and upload `nifty_tracker.py` and `config.json`.
4. **Install Packages**: In their Bash console, run:
   ```bash
   pip install --user requests yfinance pandas pytz
   ```
5. **Schedule Task**: In the **Tasks** tab, add a task at **`03:45` UTC** (9:15 AM IST) running:
   ```bash
   python /home/yourusername/nifty_tracker.py
   ```

---

### Option C: Local Automation (Windows Task Scheduler)
To run the script on your local computer daily:

1. Follow steps 1-2 in Option B.
2. Install packages: `pip install requests yfinance pandas pytz`
3. Open **Task Scheduler** on Windows, click **Create Basic Task...**, set trigger to **Daily** at **9:15 AM**, and configure action to start a program:
   - **Program/script**: `python`
   - **Add arguments**: `nifty_tracker.py`
   - **Start in**: Absolute path to this project folder.
