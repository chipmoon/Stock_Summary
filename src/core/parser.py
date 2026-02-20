import re
from datetime import datetime
from loguru import logger
from bs4 import BeautifulSoup
from src.models.trade_model import Trade

class TradeParser:
    """
    Expert AI Logic to extract trade details from various email formats.
    Supports: Traditional Regex, Fubon Securities (富邦證券) HTML reports.
    """
    
    # Traditional Regex Patterns
    REGEX_PATTERNS = {
        "symbol": r"(?:Symbol|Ticker|Stock):\s*([A-Z0-9]+)",
        "side": r"(?:Action|Side|Type):\s*(BUY|SELL|Buy|Sell)",
        "quantity": r"(?:Quantity|Qty|Shares):\s*([\d,.]+)",
        "price": r"(?:Price|Cost):\s*\$?([\d,.]+)",
        "order_id": r"(?:Order ID|Ref):\s*([A-Z0-9-]+)"
    }

    @classmethod
    def parse(cls, subject: str, body: str) -> list[Trade]:
        """
        Orchestrates parsing by detecting the report type.
        Returns a list of Trade objects because one report might contain multiple trades.
        """
        if "富邦證券" in subject:
            return cls._parse_fubon_report(subject, body)
        
        # Fallback to single-trade regex parser
        trade = cls._parse_regex(body)
        return [trade] if trade else []

    @classmethod
    def _parse_fubon_report(cls, subject: str, html_body: str) -> list[Trade]:
        """Parses Fubon Securities Taiwan HTML table report with high resilience."""
        trades = []
        try:
            # 1. Extract Date from Subject (e.g., 2026年1月19日)
            date_match = re.search(r"(\d{4})年(\d{1,2})月(\d{1,2})日", subject)
            if not date_match:
                logger.error(f"Could not extract date from Fubon subject: {subject}")
                return []
            
            report_date = f"{date_match.group(1)}-{date_match.group(2).zfill(2)}-{date_match.group(3).zfill(2)}"
            
            # 2. Parse HTML Table
            soup = BeautifulSoup(html_body, "html.parser")
            rows = soup.find_all("tr")
            
            for row in rows:
                cols = row.find_all(["td", "th"])
                if len(cols) < 7:
                    continue
                
                texts = [c.get_text(strip=True) for c in cols]
                
                # Filter out header row or notification rows
                if any(kw in texts[0] for kw in ["股票名稱", "以上資料", "重要提示"]):
                    continue
                
                # Column mapping:
                # 0: 股票名稱 (e.g., 2344華邦電)
                # 1: 交易類別 (e.g., 現賣)
                # 2: 成交股數 (e.g., 50)
                # 3: 成交單價 (117.00)
                # 5: 委託書編號
                # 6: 成交時間 (09:40:05)

                # ADVANCED VALIDATION: Ensure quantity and price look like numbers
                qty_raw = texts[2].replace(",", "")
                price_raw = texts[3].replace(",", "")
                
                if not re.match(r"^\d+\.?\d*$", qty_raw) or not re.match(r"^\d+\.?\d*$", price_raw):
                    logger.debug(f"Skipping non-trade row: {texts[0]} | {qty_raw} | {price_raw}")
                    continue

                try:
                    symbol_full = texts[0]  # e.g., "2344華邦電"
                    # Match code (digits) and rest as name
                    match = re.match(r"(\d+)(.*)", symbol_full)
                    if match:
                        symbol = match.group(1)
                        stock_name = match.group(2).strip()
                    else:
                        symbol = symbol_full
                        stock_name = ""
                    
                    side = "SELL" if "賣" in texts[1] else "BUY"
                    qty = float(qty_raw)
                    price = float(price_raw)
                    order_id = texts[5]
                    trade_time = texts[6]
                    
                    # Combine Date and Time
                    full_ts = datetime.strptime(f"{report_date} {trade_time}", "%Y-%m-%d %H:%M:%S")
                    
                    trades.append(Trade(
                        symbol=symbol,
                        stock_name=stock_name,
                        side=side,
                        quantity=qty,
                        price=price,
                        date=full_ts,
                        order_id=order_id,
                        broker="Fubon Securities (富邦證券)"
                    ))
                except (ValueError, IndexError) as e:
                    logger.warning(f"Skipping malformed row data: {e} | Content: {texts}")
                    continue
            
            logger.info(f"Fubon Parser: Extracted {len(trades)} trades from report.")
            return trades

        except Exception as e:
            logger.error(f"Fubon Global Parsing Error: {e}")
            return []

    @classmethod
    def _parse_regex(cls, text: str) -> Trade | None:
        """Original regex-based parser for single-trade emails."""
        try:
            extracted = {}
            for key, pattern in cls.REGEX_PATTERNS.items():
                match = re.search(pattern, text, re.IGNORECASE)
                if match:
                    extracted[key] = match.group(1).strip()
            
            if not all(k in extracted for k in ["symbol", "side", "quantity", "price"]):
                return None

            return Trade(
                symbol=extracted["symbol"].upper(),
                side=extracted["side"].upper(),
                quantity=float(extracted["quantity"].replace(",", "")),
                price=float(extracted["price"].replace(",", "")),
                order_id=extracted.get("order_id")
            )
        except Exception:
            return None
