import time
import os
import sys
from loguru import logger
from dotenv import load_dotenv
from src.integrations.email_scout import EmailScout
from src.integrations.sheets_ledger import SheetsLedger
from src.core.parser import TradeParser

load_dotenv()

def run_diagnostics(scout: EmailScout, ledger: SheetsLedger, dry_run: bool):
    """Verifies that all integrations can connect before starting the loop."""
    logger.info("Running system diagnostics...")
    
    # Check Email
    if not scout.test_connection():
        logger.critical("Email authentication failed. Please check .env and use an App Password.")
        sys.exit(1)
        
    # Check Sheets (if not dry run)
    if not dry_run:
        try:
            ledger.connect()
            logger.success("Google Sheets connection successful!")
        except Exception as e:
            logger.critical(f"Google Sheets connection failed: {e}")
            sys.exit(1)
    
    logger.success("All systems green. Starting automation loop.")

def main():
    import argparse
    parser = argparse.ArgumentParser(description="Trade Ledger Bot (Automation Interface)")
    parser.add_argument("--days", type=int, default=2, help="Number of days back to scan (default: 2)")
    parser.add_argument("--since", type=str, help="Start date in DD-Mon-YYYY format (e.g. 01-Jan-2026)")
    parser.add_argument("--dry-run", action="store_true", help="Scan without writing to Google Sheets")
    parser.add_argument("--force-sync", action="store_true", help="Scan only unread regardless of date (Original behavior)")
    args = parser.parse_args()

    logger.info("Starting Trade Ledger Bot (Jarvis)...")
    
    try:
        scout = EmailScout()
        ledger = SheetsLedger()
        
        # Priority: CLI argument > .env > Default(True)
        env_dry_run = os.getenv("DRY_RUN", "True").lower() == "true"
        dry_run = args.dry_run if args.dry_run else env_dry_run

        if dry_run:
            logger.warning("DRY RUN MODE ENABLED: No data will be written to Google Sheets.")
        
        # Diagnostic Check & Connection
        run_diagnostics(scout, ledger, dry_run)

        # Sync Logic Choice
        if args.force_sync:
            logger.info("Syncing only UNREAD emails (Legacy mode)...")
            email_iter = scout.fetch_new_emails()
        else:
            logger.info(f"Syncing emails from the last {args.days} days (Catch-up mode)...")
            email_iter = scout.fetch_recent_trades(days=args.days, since_date=args.since)

        # 3. Synchronize Records
        trades_recorded = 0
        emails_processed = 0
        
        # Luôn cập nhật giờ quét mới làm bằng chứng bot đã chạy (ngay cả khi chưa tìm thấy mail)
        if not dry_run:
            ledger.update_sync_timestamp()

        for subject, email_body in email_iter:
            emails_processed += 1
            trades = TradeParser.parse(subject, email_body)
            
            if trades:
                for trade in trades:
                    logger.success(f"Extracted: {trade.symbol} {trade.side} @ {trade.price} on {trade.date.date()}")
                    
                    if dry_run:
                        logger.info(f"[DRY RUN] Would record: {trade.model_dump()}")
                    else:
                        if ledger.append_trade(trade):
                            trades_recorded += 1
            else:
                logger.warning(f"Failed to parse any trade from email: {subject}")

        # Update Portfolio Summary (always if not dry run to refresh prices and sync time)
        if not dry_run:
            ledger.update_portfolio_summary()

        # Final Report
        logger.info("=" * 30)
        logger.info("Automation Task Finished")
        logger.info(f"Emails Scanned: {emails_processed}")
        logger.info(f"Trades Recorded: {trades_recorded}")
        if emails_processed == 0:
            logger.info("No trade emails found in the specified range.")
            
        logger.info("=" * 30)
        sys.exit(0)

    except KeyboardInterrupt:
        logger.info("Bot stopped by user. Goodbye!")
        sys.exit(0)
    except Exception as e:
        logger.critical(f"Unexpected fatal error: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
