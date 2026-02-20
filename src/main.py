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
    logger.info("Starting Trade Ledger Bot (Jarvis)...")
    
    try:
        scout = EmailScout()
        ledger = SheetsLedger()
        
        dry_run = os.getenv("DRY_RUN", "True").lower() == "true"

        if dry_run:
            logger.warning("DRY RUN MODE ENABLED: No data will be written to Google Sheets.")
        
        # Diagnostic Check & Connection
        run_diagnostics(scout, ledger, dry_run)

        # Single-Run Logic
        logger.info("Scanning for new trade emails...")
        emails_processed = 0
        trades_recorded = 0
        
        for subject, email_body in scout.fetch_new_emails():
            emails_processed += 1
            trades = TradeParser.parse(subject, email_body)
            
            if trades:
                for trade in trades:
                    logger.success(f"Extracted: {trade.symbol} {trade.side} @ {trade.price}")
                    
                    if dry_run:
                        logger.info(f"[DRY RUN] Would record: {trade.model_dump()}")
                    else:
                        ledger.append_trade(trade)
                        trades_recorded += 1
            else:
                logger.warning(f"Failed to parse any trade from email: {subject}")

        # Update Portfolio Summary (only if not dry run)
        if not dry_run and trades_recorded >= 0:
            ledger.update_portfolio_summary()

        # Final Report
        logger.info("=" * 30)
        logger.info("Automation Task Finished")
        logger.info(f"Emails Scanned: {emails_processed}")
        logger.info(f"Trades Recorded: {trades_recorded}")
        if emails_processed == 0:
            logger.info("No unread trade emails found. Tip: Mark your Fubon email as 'Unread' in Gmail to process it.")
            
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
