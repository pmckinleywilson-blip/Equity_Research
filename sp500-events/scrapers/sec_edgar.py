"""Detection Tier: SEC EDGAR 8-K filing scraper.

Searches EFTS (full-text search) for 8-K filings containing:
- Item 2.02: Results of Operations and Financial Condition
- Item 7.01: Regulation FD Disclosure
- Item 8.01: Other Events

These signal earnings releases, conference calls, and ad-hoc events.
"""
import logging
import re
import sys
from datetime import datetime, timedelta
from pathlib import Path

import httpx

sys.path.insert(0, str(Path(__file__).parent.parent / "backend"))

from config import get_settings

logger = logging.getLogger(__name__)
settings = get_settings()

EFTS_URL = "https://efts.sec.gov/LATEST/search-index"
FILING_URL = "https://www.sec.gov/Archives/edgar/data"


def search_8k_filings(start_date: str = None, end_date: str = None) -> list[dict]:
    """Search SEC EDGAR for recent 8-K filings with earnings-related items.

    Uses the EFTS full-text search API.
    Rate limit: 10 requests/second. User-Agent must include contact email.
    """
    if start_date is None:
        start_date = (datetime.utcnow() - timedelta(days=1)).strftime("%Y-%m-%d")
    if end_date is None:
        end_date = datetime.utcnow().strftime("%Y-%m-%d")

    headers = {
        "User-Agent": settings.sec_edgar_user_agent,
        "Accept": "application/json",
    }

    # Search for 8-K filings with key item numbers
    search_queries = [
        '"Item 2.02"',  # Results of Operations
        '"Item 7.01"',  # Regulation FD
        '"Item 8.01"',  # Other Events
    ]

    all_filings = []

    for query in search_queries:
        try:
            params = {
                "q": query,
                "dateRange": "custom",
                "startdt": start_date,
                "enddt": end_date,
                "forms": "8-K",
            }

            resp = httpx.get(
                "https://efts.sec.gov/LATEST/search-index",
                params=params,
                headers=headers,
                timeout=30,
            )
            resp.raise_for_status()
            data = resp.json()

            hits = data.get("hits", {}).get("hits", [])
            for hit in hits:
                source = hit.get("_source", {})
                filing = {
                    "cik": source.get("entity_id", ""),
                    "entity_name": source.get("entity_name", ""),
                    "filed_date": source.get("file_date", ""),
                    "form_type": source.get("form_type", ""),
                    "file_url": source.get("file_url", ""),
                    "query_match": query,
                }
                all_filings.append(filing)

        except Exception as e:
            logger.error("EDGAR EFTS search failed for %s: %s", query, e)

    logger.info("EDGAR: found %d 8-K filings from %s to %s", len(all_filings), start_date, end_date)
    return all_filings


def extract_event_details_from_8k(filing_url: str) -> dict:
    """Parse an 8-K filing to extract event date, webcast URL, dial-in number.

    This does a best-effort extraction from the filing text.
    """
    headers = {
        "User-Agent": settings.sec_edgar_user_agent,
    }

    try:
        resp = httpx.get(filing_url, headers=headers, timeout=30)
        resp.raise_for_status()
        text = resp.text
    except Exception as e:
        logger.error("Failed to fetch 8-K filing %s: %s", filing_url, e)
        return {}

    details = {}

    # Extract webcast URL patterns
    webcast_patterns = [
        r'https?://[^\s<"]+(?:webcast|event|q4inc|notified|cision)[^\s<"]*',
        r'https?://[^\s<"]+(?:earnings|investor|ir\.)[^\s<"]*(?:webcast|call|event)[^\s<"]*',
    ]
    for pattern in webcast_patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            details["webcast_url"] = match.group(0).rstrip(".")
            break

    # Extract phone number patterns
    phone_pattern = r'(?:dial[- ]?in|telephone|call[- ]?in)[:\s]*\(?(\+?1?[-.\s]?\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4})'
    match = re.search(phone_pattern, text, re.IGNORECASE)
    if match:
        details["phone_number"] = match.group(1).strip()

    # Extract passcode
    passcode_pattern = r'(?:passcode|access code|conference id|pin)[:\s]*(\d{4,10})'
    match = re.search(passcode_pattern, text, re.IGNORECASE)
    if match:
        details["phone_passcode"] = match.group(1)

    # Extract date/time of the call
    # Look for patterns like "January 15, 2026 at 4:30 PM"
    datetime_pattern = r'(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},?\s+\d{4}\s+(?:at\s+)?\d{1,2}:\d{2}\s*(?:AM|PM|a\.m\.|p\.m\.)'
    match = re.search(datetime_pattern, text, re.IGNORECASE)
    if match:
        details["datetime_text"] = match.group(0)

    return details


def process_8k_filings(filings: list[dict], db=None) -> dict:
    """Process 8-K filings and store as tentative events or enrich existing ones."""
    from database import SessionLocal, init_db
    from models import Event, Company

    close_session = False
    if db is None:
        init_db()
        db = SessionLocal()
        close_session = True

    try:
        enriched = 0
        for filing in filings:
            cik = filing.get("cik", "").lstrip("0")
            if not cik:
                continue

            # Look up company by CIK
            company = db.query(Company).filter(Company.cik == cik.zfill(10)).first()
            if not company:
                # Try without zero-padding
                company = db.query(Company).filter(Company.cik == cik).first()
            if not company:
                continue

            # Try to extract details from the filing
            if filing.get("file_url"):
                details = extract_event_details_from_8k(filing["file_url"])
                if details:
                    # Find matching tentative event for this company
                    filed = filing.get("filed_date", "")
                    if filed:
                        from datetime import date as date_type
                        filed_date = datetime.strptime(filed, "%Y-%m-%d").date()
                        # Look for events within 30 days of filing
                        existing = (
                            db.query(Event)
                            .filter(
                                Event.ticker == company.ticker,
                                Event.event_type == "earnings",
                                Event.event_date >= filed_date,
                                Event.event_date <= filed_date + timedelta(days=30),
                            )
                            .first()
                        )
                        if existing and not existing.ir_verified:
                            if details.get("webcast_url") and not existing.webcast_url:
                                existing.webcast_url = details["webcast_url"]
                            if details.get("phone_number") and not existing.phone_number:
                                existing.phone_number = details["phone_number"]
                            if details.get("phone_passcode") and not existing.phone_passcode:
                                existing.phone_passcode = details["phone_passcode"]
                            existing.source_url = filing.get("file_url")
                            existing.updated_at = datetime.utcnow()
                            enriched += 1

        db.commit()
        logger.info("EDGAR: enriched %d events with filing details", enriched)
        return {"enriched": enriched}

    finally:
        if close_session:
            db.close()


def run_edgar_detection(days_back: int = 2):
    """Run SEC EDGAR 8-K detection."""
    start = (datetime.utcnow() - timedelta(days=days_back)).strftime("%Y-%m-%d")
    end = datetime.utcnow().strftime("%Y-%m-%d")

    filings = search_8k_filings(start, end)
    result = process_8k_filings(filings)
    return result


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    result = run_edgar_detection()
    print(f"EDGAR detection complete: {result}")
