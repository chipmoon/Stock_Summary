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
        if only_unread:
            search_params["seen"] = False
        
        if since_date:
            # imap_tools requires datetime.date object for date criteria
            try:
                date_obj = datetime.strptime(since_date, "%d-%b-%Y").date()
                search_params["date_gte"] = date_obj
            except Exception as e:
                logger.error(f"Invalid date format for since_date: {since_date}. Use DD-Mon-YYYY. Error: {e}")

        # Build search filter using imap_tools AND builder
        search_filter = AND(**search_params) if search_params else AND(all=True)

        try:
            with MailBox(self.host).login(self.user, self.password, self.folder) as mailbox:
                logger.debug(f"Scanning emails in {self.folder} (Filter: {search_params})...")
                
                # Fetch messages (reverse=True is faster for finding recent ones)
                limit = 1000 if since_date else 50
                
                for msg in mailbox.fetch(search_filter, limit=limit, reverse=True):
                    subject = msg.subject or ""
                    
                    # Check if any criteria matches the subject
                    match_found = any(c.lower() in subject.lower() for c in criteria_list)
                    
                    if match_found:
                        logger.info(f"Matched trade report: {subject}")
                        yield subject, (msg.html or msg.text)
        except Exception as e:
            logger.error(f"Error during email fetch: {type(e).__name__}: {e}")

    # Alias for backward compatibility in main.py
    def fetch_new_emails(self, criteria_list: list[str] = None):
        return self.fetch_emails(criteria_list=criteria_list, only_unread=True)
