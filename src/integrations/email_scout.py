import os
import sys
from imap_tools import MailBox, AND, errors
from loguru import logger
from dotenv import load_dotenv

load_dotenv()

class EmailScout:
    """
    Monitors email inbox for new trade confirmations with robust error handling.
    """
    def __init__(self):
        self.host = os.getenv("EMAIL_HOST")
        self.user = os.getenv("EMAIL_USER")
        self.password = os.getenv("EMAIL_PASS")
        self.folder = os.getenv("EMAIL_FOLDER", "INBOX")
        
        # Gmail optimization: Archived emails often move from INBOX to "All Mail"
        if "gmail.com" in self.host.lower() and self.folder == "INBOX":
            logger.info("Gmail detected. If historical data is missing, consider changing EMAIL_FOLDER to '\"[Gmail]/All Mail\"' in .env")
        
        self._validate_config()

    def _validate_config(self):
        """Ensures all required environment variables are present."""
        config_map = {
            "EMAIL_HOST": self.host,
            "EMAIL_USER": self.user,
            "EMAIL_PASS": self.password
        }
        missing = [k for k, v in config_map.items() if not v]
        
        if missing:
            logger.critical(f"Missing required environment variables: {', '.join(missing)}")
            sys.exit(1)
        
        logger.debug(f"Config validated for user: {self.user} on {self.host}")

    def test_connection(self) -> bool:
        """
        Attempts a login to verify credentials. 
        Returns True if successful, raises error otherwise.
        """
        try:
            with MailBox(self.host).login(self.user, self.password, self.folder):
                logger.success("IMAP connection test successful!")
                return True
        except errors.MailboxLoginError as e:
            logger.error(f"AUTHENTICATION FAILED: Check your user/host and ensure you are using an APP PASSWORD. Detail: {e}")
            return False
        except Exception as e:
            logger.error(f"IMAP Error during test: {type(e).__name__}: {e}")
            return False

    def fetch_emails(self, criteria_list: list[str] = None, only_unread: bool = True, since_date: str = None):
        """
        Yields (subject, body) for emails matching any criteria.
        since_date expected format: 'DD-Mon-YYYY' e.g. '01-Jan-2019'
        """
        from datetime import datetime
        if criteria_list is None:
            criteria_list = ["Trade Confirmation", "富邦證券"]

        search_params = {}
        
        # GitHub Action optimization: If a specific date is given, we scan ALL (read/unread) 
        # to ensure no trade is missed even if the user manually opened the mail.
        if only_unread and not since_date:
            search_params["seen"] = False
        
        if since_date:
            try:
                date_obj = datetime.strptime(since_date, "%d-%b-%Y").date()
                search_params["date_gte"] = date_obj
                logger.info(f"Targeting emails from: {since_date} onwards...")
            except Exception as e:
                logger.error(f"Invalid date format for since_date: {since_date}. Use DD-Mon-YYYY. Error: {e}")

        # Build search filter using imap_tools AND builder
        search_filter = AND(**search_params) if search_params else AND(all=True)

        try:
            with MailBox(self.host).login(self.user, self.password, self.folder) as mailbox:
                logger.debug(f"Scanning emails (Filter: {search_params})...")
                
                # Fetch messages (reverse=True for recent first)
                for msg in mailbox.fetch(search_filter, reverse=True):
                    subject = msg.subject or ""
                    match_found = any(c.lower() in subject.lower() for c in criteria_list)
                    
                    if match_found:
                        logger.info(f"Matched trade report: {subject} ({msg.date.date()})")
                        yield subject, (msg.html or msg.text)
        except Exception as e:
            logger.error(f"Error during email fetch: {type(e).__name__}: {e}")

    # New method for optimized daily sync
    def fetch_recent_trades(self, days: int = 2, since_date: str = None):
        """Fetches trades within a window to ensure no gaps."""
        from datetime import datetime, timedelta
        
        if not since_date:
            # Default to 'Today - N days' to cover gaps
            since_date = (datetime.now() - timedelta(days=days)).strftime("%d-%b-%Y")
        
        return self.fetch_emails(since_date=since_date, only_unread=False)

    # Alias for backward compatibility
    def fetch_new_emails(self, criteria_list: list[str] = None):
        return self.fetch_emails(criteria_list=criteria_list, only_unread=True)
