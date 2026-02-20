import os
import gspread
import requests
import re
from datetime import datetime
from loguru import logger
from dotenv import load_dotenv
from src.models.trade_model import Trade

load_dotenv()

class SheetsLedger:
    """
    Manages the Google Sheets ledger to record trades.
    """
    def __init__(self):
        self.credentials_path = os.getenv("GOOGLE_SHEETS_KEY_FILE")
        self.spreadsheet_name = os.getenv("SPREADSHEET_NAME")
        self.spreadsheet_id = os.getenv("SPREADSHEET_ID")
        self.worksheet_name = os.getenv("WORKSHEET_NAME", "Trades")
        self.client = None
        self.sheet = None
        self.existing_order_ids = set() # Cache to prevent duplicates
        
        self._validate_config()

    def _validate_config(self):
        """Ensures required Sheets variables are present."""
        missing = []
        if not self.credentials_path: missing.append("GOOGLE_SHEETS_KEY_FILE")
        # Need either Name or ID
        if not self.spreadsheet_name and not self.spreadsheet_id:
            missing.append("SPREADSHEET_NAME or SPREADSHEET_ID")
        
        if missing:
            logger.critical(f"Missing required Sheets variables: {', '.join(missing)}")
            import sys
            sys.exit(1)

    def connect(self):
        """
        Authenticates with Google Sheets API with detailed diagnostics.
        """
        if not os.path.exists(self.credentials_path):
            logger.critical(f"Google Sheets key file not found at: {self.credentials_path}")
            raise FileNotFoundError(f"Key file missing: {self.credentials_path}")

        try:
            self.client = gspread.service_account(filename=self.credentials_path)
            
            try:
                if self.spreadsheet_id:
                    logger.debug(f"Opening spreadsheet by ID: {self.spreadsheet_id}")
                    spreadsheet = self.client.open_by_key(self.spreadsheet_id)
                else:
                    logger.debug(f"Opening spreadsheet by Name: {self.spreadsheet_name}")
                    spreadsheet = self.client.open(self.spreadsheet_name)
            except gspread.exceptions.SpreadsheetNotFound:
                raise
            except (PermissionError, gspread.exceptions.APIError) as e:
                logger.critical("PERMISSION DENIED or API DISABLED.")
                logger.info("1. Ensure Google Sheets API & Google Drive API are ENABLED in Google Cloud Console.")
                logger.info("2. Link: https://console.cloud.google.com/apis/library/sheets.googleapis.com")
                logger.error(f"Detail: {e}")
                raise

            # Tab Renaming Strategy
            target_sheet_name = os.getenv("WORKSHEET_NAME", "Buy_Sell_Status")
            
            # Try to get the target sheet directly
            try:
                self.sheet = spreadsheet.worksheet(target_sheet_name)
            except gspread.exceptions.WorksheetNotFound:
                # If not found, check if "Trades" or "Sheet1" exists to rename
                renamed = False
                for old_name in ["Trades", "Sheet1"]:
                    try:
                        old_sheet = spreadsheet.worksheet(old_name)
                        old_sheet.update_title(target_sheet_name)
                        self.sheet = old_sheet
                        logger.info(f"Renamed worksheet '{old_name}' to '{target_sheet_name}'")
                        renamed = True
                        break
                    except gspread.exceptions.WorksheetNotFound:
                        continue
                
                if not renamed:
                     # Create new if nothing found
                    self.sheet = spreadsheet.add_worksheet(target_sheet_name, rows=1000, cols=10)
                    logger.info(f"Created new worksheet '{target_sheet_name}'")

            logger.info(f"Connected to Google Sheet: {spreadsheet.title} | Worksheet: {self.sheet.title}")
            self._ensure_headers()
            self._load_existing_ids() # Load IDs for duplicate check
                
        except Exception as e:
            logger.error(f"Google Sheets Connection Error ({type(e).__name__}): {e}")
            raise

    def _ensure_headers(self):
        """
        Checks if the first row is a header row. If missing or looks like data, inserts headers.
        """
        try:
            expected_headers = [
                "Trade Time", "Stock Code", "Stock Name", "Action", 
                "Shares", "Unit Price", "Total Amount", "Order ID", "Broker"
            ]
            
            existing_headers = self.sheet.row_values(1)
            
            if not existing_headers:
                self.sheet.append_row(expected_headers)
                return

            if existing_headers[0] != "Trade Time":
                self.sheet.insert_row(expected_headers, index=1)

        except Exception as e:
            logger.error(f"Failed to ensure headers: {e}")

    def _load_existing_ids(self):
        """Loads all Order IDs from the sheet to prevent duplicates."""
        try:
            # Assuming Order ID is in column 8 (index 7) based on expected_headers
            all_ids = self.sheet.col_values(8)
            # Skip header
            self.existing_order_ids = set(val.strip() for val in all_ids[1:] if val.strip())
            logger.info(f"Loaded {len(self.existing_order_ids)} existing Order IDs.")
        except Exception as e:
            logger.warning(f"Could not load existing Order IDs: {e}")

    def append_trade(self, trade: Trade):
        if not self.sheet: self.connect()
        
        # Duplicate Check
        if trade.order_id and trade.order_id.strip() in self.existing_order_ids:
            logger.warning(f"Duplicate trade detected, skipping: {trade.order_id}")
            return False

        try:
            row = [
                trade.date.strftime("%Y-%m-%d %H:%M:%S"),
                trade.symbol,
                trade.stock_name or "",
                trade.side,
                trade.quantity,
                trade.price,
                trade.total_amount,
                trade.order_id,
                trade.broker
            ]
            self.sheet.append_row(row)
            if trade.order_id:
                self.existing_order_ids.add(trade.order_id.strip())
            logger.success(f"Recorded: {trade.symbol} {trade.stock_name} {trade.side}")
            return True
        except Exception as e:
            logger.error(f"Error appending trade to sheet: {e}")
            raise

    def batch_append_trades(self, trades: list[Trade]):
        """Efficiently appends multiple trades in one go."""
        if not self.sheet: self.connect()
        
        rows_to_add = []
        for trade in trades:
            if trade.order_id and trade.order_id.strip() in self.existing_order_ids:
                continue
            
            rows_to_add.append([
                trade.date.strftime("%Y-%m-%d %H:%M:%S"),
                trade.symbol,
                trade.stock_name or "",
                trade.side,
                trade.quantity,
                trade.price,
                trade.total_amount,
                trade.order_id,
                trade.broker
            ])
            if trade.order_id:
                self.existing_order_ids.add(trade.order_id.strip())

        if not rows_to_add:
            return 0

        try:
            self.sheet.append_rows(rows_to_add)
            logger.success(f"Batch recorded {len(rows_to_add)} trades.")
            return len(rows_to_add)
        except Exception as e:
            logger.error(f"Batch append failed: {e}")
            raise

    def _get_yahoo_price(self, code: str) -> float:
        """Lightweight Yahoo Finance scraper using query1 API to avoid pandas/yfinance dependencies."""
        for suffix in [".TW", ".TWO"]:
            ticker = f"{code}{suffix}"
            try:
                url = f"https://query1.finance.yahoo.com/v8/finance/chart/{ticker}"
                headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'}
                response = requests.get(url, headers=headers, timeout=10)
                if response.status_code == 200:
                    data = response.json()
                    # Check if result exists and has meta
                    if data.get('chart', {}).get('result'):
                        price = data['chart']['result'][0]['meta'].get('regularMarketPrice')
                        if price and price > 0:
                            return float(price)
            except Exception as e:
                logger.debug(f"Yahoo check failed for {ticker}: {e}")
                continue
        return 0.0

    def update_portfolio_summary(self):
        """
        Expert AI Manager: Generates a 'Morgan-style' Total Wealth Executive Dashboard.
        Features:
        1. REAL-TIME Pricing: Uses =GOOGLEFINANCE for live market updates.
        2. TOTAL WEALTH: Tracks Realized P/L + Unrealized P/L (Current Value).
        3. FINANCIAL KPIs: Total Invested, Current Value, ROI %, Total Profit.
        """
        if not self.sheet: self.connect()
        logger.info("Designing Executive Total Wealth Dashboard (Option B)...")

        try:
            rows = self.sheet.get_all_values()
            if not rows or len(rows) < 2: return

            headers = rows[0]
            def find_col(possible_names, default_idx):
                for name in possible_names:
                    if name in headers: return headers.index(name)
                return default_idx

            idx_time = find_col(["Trade Time", "Date"], 0)
            idx_code = find_col(["Stock Code", "Stock Symbol", "Symbol"], 1)
            idx_name = find_col(["Stock Name", "Name"], 2)
            idx_action = find_col(["Action", "Side"], 3)
            idx_qty = find_col(["Shares", "Quantity"], 4)
            idx_price = find_col(["Unit Price", "Price"], 5)
            idx_amt = find_col(["Total Amount", "Amount"], 6)

            portfolio = {} 
            total_realized_pl = 0.0

            for row in rows[1:]:
                if len(row) <= max(idx_code, idx_action, idx_qty, idx_amt): continue
                
                code = row[idx_code].strip()
                name = row[idx_name].strip()
                action = row[idx_action].strip().upper()
                
                try:
                    qty = float(str(row[idx_qty]).replace(",", ""))
                    amt = float(str(row[idx_amt]).replace(",", ""))
                except Exception: continue

                if code not in portfolio:
                    portfolio[code] = {
                        "name": name, 
                        "buy_qty": 0.0, "buy_amt": 0.0, 
                        "sell_qty": 0.0, "sell_amt": 0.0,
                        "realized_pl": 0.0
                    }
                
                d = portfolio[code]
                if name: d["name"] = name

                if "BUY" in action or "買" in action:
                    d["buy_qty"] += qty
                    d["buy_amt"] += amt
                elif "SELL" in action or "賣" in action:
                    # Calculate realized P/L for this sale using current Avg Cost
                    avg_buy_before = d["buy_amt"] / d["buy_qty"] if d["buy_qty"] > 0 else 0
                    current_sale_price = amt / qty if qty > 0 else 0
                    profit = (current_sale_price - avg_buy_before) * qty
                    d["realized_pl"] += profit
                    total_realized_pl += profit
                    
                    d["sell_qty"] += qty
                    d["sell_amt"] += amt

            # --- Fetch Yahoo Finance Prices (Lightweight) ---
            # Fetch for anything we intend to show: active holdings OR significant realized trades
            codes = [
                c for c in portfolio 
                if (portfolio[c]["buy_qty"] - portfolio[c]["sell_qty"]) != 0 
                or portfolio[c]["realized_pl"] != 0
            ]
            prices = {}
            if codes:
                logger.info(f"Fetching Yahoo prices for {len(codes)} assets shown in Dashboard...")
                for code in codes:
                    prices[code] = self._get_yahoo_price(code)

            # Prepare Dashboard Structure
            summary_rows = []
            for code in sorted(portfolio.keys()):
                d = portfolio[code]
                net_qty = d["buy_qty"] - d["sell_qty"]
                avg_buy = d["buy_amt"] / d["buy_qty"] if d["buy_qty"] > 0 else 0
                live_price = prices.get(code, 0)
                
                # Show active holdings or significant realized trades
                if net_qty > 0 or d["realized_pl"] != 0:
                    # Formula for P/L based on static live_price
                    # Columns: Code, Name, Qty, Avg Buy, Live Price, Unrealized P/L (Formula), P/L %, Realized
                    row_data = [
                        code, 
                        d["name"], 
                        net_qty, 
                        avg_buy,
                        live_price,
                        f'=IF(C{{row}}>0, (E{{row}}-D{{row}})*C{{row}}, 0)', # Unrealized P/L
                        f'=IF(AND(E{{row}}>0, D{{row}}>0), (E{{row}}-D{{row}})/D{{row}}, 0)', # P/L %
                        d["realized_pl"]
                    ]
                    summary_rows.append(row_data)

            # Create/Reset Dashboard Sheet
            spreadsheet = self.sheet.spreadsheet
            dash_name = "Executive_Dashboard"
            try:
                ws = spreadsheet.worksheet(dash_name)
                spreadsheet.del_worksheet(ws)
            except gspread.exceptions.WorksheetNotFound:
                pass

            dash_sheet = spreadsheet.add_worksheet(dash_name, 100, 20)
            
            # --- HEADER & KPIs ---
            dash_sheet.update("A1", [["--- MORGAN TOTAL WEALTH EXECUTIVE DASHBOARD ---"]])
            
            # KPI Names
            kpi_labels = [["Total Realized P/L", "Unrealized P/L (Live)", "Total Wealth Change", "ROI %"]]
            dash_sheet.update("A3:D3", kpi_labels)
            
            # KPI Formulas (Dynamically referencing the table below)
            # Find end row of data
            last_data_row = 10 + len(summary_rows)
            kpi_formulas = [[
                f"=SUM(H11:H{last_data_row})", # Total Realized
                f"=SUM(F11:F{last_data_row})", # Unrealized
                "=A4+B4",                      # Total Wealth Change
                "=IFERROR(C4/SUMPRODUCT(C11:C{lr}, D11:D{lr}), 0)".replace("{lr}", str(last_data_row)) # ROI
            ]]
            dash_sheet.update("A4:D4", kpi_formulas, value_input_option='USER_ENTERED')

            # --- HOLDINGS TABLE ---
            table_headers = [["Code", "Asset Name", "Qty", "Avg Buy", "Live Price", "Unrealized P/L", "P/L %", "Realized P/L"]]
            dash_sheet.update("A10:H10", table_headers)
            
            if summary_rows:
                # Resolve formulas with correct row index
                final_table = []
                for i, r in enumerate(summary_rows):
                    row_idx = 11 + i
                    processed_row = [str(cell).replace("{row}", str(row_idx)) for cell in r]
                    final_table.append(processed_row)
                
                dash_sheet.update("A11", final_table, value_input_option='USER_ENTERED')

            # --- FORMATTING (The "Wall Street" Look) ---
            twd_format = {"numberFormat": {"type": "CURRENCY", "pattern": '"NT$"#,##0.00'}}
            
            # KPI Colors & Formats
            dash_sheet.format("A4:C4", {**twd_format, "textFormat": {"bold": True, "fontSize": 12}})
            dash_sheet.format("D4", {"numberFormat": {"type": "PERCENT", "pattern": "0.00%"}, "textFormat": {"bold": True}})
            
            # Table Header
            dash_sheet.format("A10:H10", {"textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}}, "backgroundColor": {"red": 0.2, "green": 0.2, "blue": 0.2}})
            
            # Table Data Formats (Currency for money, Percent for yield)
            if summary_rows:
                dash_sheet.format(f"D11:F{last_data_row}", twd_format)
                dash_sheet.format(f"H11:H{last_data_row}", twd_format)
                dash_sheet.format(f"G11:G{last_data_row}", {"numberFormat": {"type": "PERCENT", "pattern": "0.00%"}})
            
            # Conditional Formatting for P/L (Green for Profit, Red for Loss)
            requests = [
                {
                    "addConditionalFormatRule": {
                        "rule": {
                            "ranges": [{"sheetId": dash_sheet.id, "startRowIndex": 10, "endRowIndex": last_data_row, "startColumnIndex": 5, "endColumnIndex": 7}],
                            "booleanRule": {
                                "condition": {"type": "NUMBER_GREATER", "values": [{"userEnteredValue": "0"}]},
                                "format": {"textFormat": {"foregroundColor": {"red": 0, "green": 0.6, "blue": 0}}}
                            }
                        }, "index": 0
                    }
                },
                {
                    "addConditionalFormatRule": {
                        "rule": {
                            "ranges": [{"sheetId": dash_sheet.id, "startRowIndex": 10, "endRowIndex": last_data_row, "startColumnIndex": 5, "endColumnIndex": 7}],
                            "booleanRule": {
                                "condition": {"type": "NUMBER_LESS", "values": [{"userEnteredValue": "0"}]},
                                "format": {"textFormat": {"foregroundColor": {"red": 0.8, "green": 0, "blue": 0}}}
                            }
                        }, "index": 1
                    }
                }
            ]
            
            # Charts
            sheet_id = dash_sheet.id
            chart_req = {
                "addChart": {
                    "chart": {
                        "spec": {
                            "title": "Portfolio Value Distribution",
                            "pieChart": {
                                "legendPosition": "RIGHT_LEGEND",
                                "domain": {"sourceRange": {"sources": [{"sheetId": sheet_id, "startRowIndex": 10, "endRowIndex": last_data_row, "startColumnIndex": 1, "endColumnIndex": 2}]}},
                                "series": {"sourceRange": {"sources": [{"sheetId": sheet_id, "startRowIndex": 10, "endRowIndex": last_data_row, "startColumnIndex": 5, "endColumnIndex": 6}]}}
                            }
                        },
                        "position": {"overlayPosition": {"anchorCell": {"sheetId": sheet_id, "rowIndex": 1, "columnIndex": 10}, "widthPixels": 400, "heightPixels": 250}}
                    }
                }
            }
            requests.append(chart_req)
            
            spreadsheet.batch_update({"requests": requests})
            logger.success("Morgan Total Wealth Dashboard is now LIVE with real-time pricing!")

        except Exception as e:
            logger.error(f"Failed to generate Total Wealth Dashboard: {e}")
            import traceback
            logger.debug(traceback.format_exc())
