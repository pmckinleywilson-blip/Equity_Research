"""Verification Tier: Company IR page scraper — SOURCE OF TRUTH.

Scrapes investor relations event pages to:
1. Verify detected events (tentative → confirmed)
2. Extract webcast URLs, dial-in numbers, event details
3. Discover events not found by detection sources

This is the authoritative source. If an event is found on the IR page,
it is marked as ir_verified=TRUE and status='confirmed'.
"""
import logging
import re
import sys
from datetime import datetime, timedelta, date
from pathlib import Path
from typing import Optional

import httpx
from bs4 import BeautifulSoup

sys.path.insert(0, str(Path(__file__).parent.parent / "backend"))

from database import SessionLocal, init_db
from models import Event, Company

logger = logging.getLogger(__name__)

# Common IR page platform patterns
IR_PLATFORMS = {
    "q4": {
        "event_selector": ".module-event, .event-item, [class*='event']",
        "date_selector": ".event-date, .date, time",
        "title_selector": ".event-title, .title, h3, h4",
        "link_selector": "a[href*='webcast'], a[href*='event'], a.btn",
    },
    "notified": {
        "event_selector": ".nir-widget--event, .event-row, [data-event]",
        "date_selector": ".event-date, .date",
        "title_selector": ".event-title, .title",
        "link_selector": "a[href*='notified'], a[href*='webcast']",
    },
    "generic": {
        "event_selector": "[class*='event'], [class*='calendar'], tr, .row",
        "date_selector": "[class*='date'], time, td:first-child",
        "title_selector": "[class*='title'], h3, h4, td:nth-child(2)",
        "link_selector": "a[href*='webcast'], a[href*='event'], a[href*='register']",
    },
}


def scrape_ir_page(ir_url: str, scraper_type: str = "generic") -> list[dict]:
    """Scrape a company's IR events page for upcoming events using LLM extraction.

    Fetches the page, extracts text, sends to Gemini for structured extraction.
    No regex parsing — the LLM handles all text interpretation.
    """
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "Accept": "text/html,application/xhtml+xml",
    }

    try:
        resp = httpx.get(ir_url, headers=headers, timeout=30, follow_redirects=True)
        resp.raise_for_status()
    except Exception as e:
        logger.warning("Failed to fetch IR page %s: %s", ir_url, e)
        return []

    soup = BeautifulSoup(resp.text, "lxml")

    # Extract page text
    body = soup.select_one("main") or soup.select_one("body")
    if not body:
        return []

    text = body.get_text(separator="\n")
    if len(text.strip()) < 100:
        return []

    # Use LLM to extract events
    try:
        from llm_extractor import extract_event_from_text
        result = extract_event_from_text(text, "")
        if result and result.get("event_date"):
            result["source"] = "ir_page"
            result["source_url"] = ir_url
            return [result]
    except ImportError:
        logger.error("llm_extractor not available for IR page scraping")
    except Exception as e:
        logger.error("LLM IR extraction failed for %s: %s", ir_url, e)

    return []


def _parse_event_element(elem, platform: dict, base_url: str) -> Optional[dict]:
    """Parse a single event element from an IR page."""
    event = {}

    # Extract title
    title_elem = elem.select_one(platform["title_selector"])
    if title_elem:
        event["title"] = title_elem.get_text(strip=True)

    # Extract date
    date_elem = elem.select_one(platform["date_selector"])
    if date_elem:
        date_text = date_elem.get_text(strip=True)
        parsed_date = _parse_date_text(date_text)
        if parsed_date:
            event["event_date"] = parsed_date

    # Also try datetime attribute on <time> elements
    time_elem = elem.find("time")
    if time_elem and time_elem.get("datetime"):
        try:
            dt = datetime.fromisoformat(time_elem["datetime"].replace("Z", "+00:00"))
            event["event_date"] = dt.date()
            event["event_time"] = dt.time()
        except ValueError:
            pass

    # Extract webcast link
    link_elem = elem.select_one(platform["link_selector"])
    if link_elem and link_elem.get("href"):
        href = link_elem["href"]
        if not href.startswith("http"):
            from urllib.parse import urljoin
            href = urljoin(base_url, href)
        event["webcast_url"] = href

    # Extract phone number from text
    text = elem.get_text()
    phone_match = re.search(r'(?:1[-.])?(?:\(\d{3}\)|\d{3})[-.\s]?\d{3}[-.\s]?\d{4}', text)
    if phone_match:
        event["phone_number"] = phone_match.group(0)

    passcode_match = re.search(r'(?:passcode|access code|pin|id)[:\s]*(\d{4,10})', text, re.I)
    if passcode_match:
        event["phone_passcode"] = passcode_match.group(1)

    # Determine event type from title
    title = event.get("title", "").lower()
    if any(kw in title for kw in ["earnings", "quarterly", "results", "q1", "q2", "q3", "q4"]):
        event["event_type"] = "earnings"
    elif any(kw in title for kw in ["investor day", "analyst day", "capital markets"]):
        event["event_type"] = "investor_day"
    elif any(kw in title for kw in ["conference", "summit", "presentation"]):
        event["event_type"] = "conference"
    else:
        event["event_type"] = "ad_hoc"

    return event if event.get("event_date") else None


def _parse_date_text(text: str) -> Optional[date]:
    """Parse various date formats from IR page text."""
    text = text.strip()

    # Common date formats
    formats = [
        "%B %d, %Y",       # January 15, 2026
        "%b %d, %Y",       # Jan 15, 2026
        "%m/%d/%Y",         # 01/15/2026
        "%Y-%m-%d",         # 2026-01-15
        "%d %B %Y",         # 15 January 2026
        "%d %b %Y",         # 15 Jan 2026
        "%B %d %Y",         # January 15 2026
    ]

    for fmt in formats:
        try:
            return datetime.strptime(text[:len(datetime.now().strftime(fmt)) + 5], fmt).date()
        except (ValueError, IndexError):
            continue

    # Try regex extraction
    match = re.search(
        r'(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)\s+\d{1,2},?\s+\d{4}',
        text,
        re.I,
    )
    if match:
        for fmt in ["%B %d, %Y", "%b %d, %Y", "%B %d %Y", "%b %d %Y"]:
            try:
                return datetime.strptime(match.group(0).replace(",", ", "), fmt).date()
            except ValueError:
                continue

    return None


def _extract_jsonld_events(soup: BeautifulSoup) -> list[dict]:
    """Extract events from Schema.org JSON-LD structured data."""
    import json
    events = []

    for script in soup.find_all("script", type="application/ld+json"):
        try:
            data = json.loads(script.string)
            if isinstance(data, list):
                for item in data:
                    if item.get("@type") == "Event":
                        events.append(_jsonld_to_event(item))
            elif isinstance(data, dict) and data.get("@type") == "Event":
                events.append(_jsonld_to_event(data))
        except (json.JSONDecodeError, AttributeError):
            continue

    return [e for e in events if e]


def _jsonld_to_event(data: dict) -> Optional[dict]:
    """Convert a Schema.org Event JSON-LD object to our event format."""
    try:
        event = {
            "title": data.get("name", ""),
            "description": data.get("description", ""),
            "event_type": "ad_hoc",
        }

        # Parse date
        start = data.get("startDate", "")
        if start:
            try:
                dt = datetime.fromisoformat(start.replace("Z", "+00:00"))
                event["event_date"] = dt.date()
                event["event_time"] = dt.time()
            except ValueError:
                return None

        # URL
        if data.get("url"):
            event["webcast_url"] = data["url"]

        # Location might be a URL for virtual events
        location = data.get("location", {})
        if isinstance(location, dict) and location.get("url"):
            event["webcast_url"] = location["url"]

        return event
    except Exception:
        return None


def verify_events_on_ir(tickers: list[str] = None, db=None):
    """Verify detected events against company IR pages.

    For each company with a known IR URL:
    1. Scrape the IR page
    2. Match scraped events against DB events
    3. Update matching events: ir_verified=TRUE, status='confirmed', add details
    """
    close_session = False
    if db is None:
        init_db()
        db = SessionLocal()
        close_session = True

    try:
        query = db.query(Company).filter(Company.ir_events_url.isnot(None))
        if tickers:
            query = query.filter(Company.ticker.in_(tickers))

        companies = query.all()
        verified = 0
        new_discovered = 0

        for company in companies:
            try:
                ir_events = scrape_ir_page(
                    company.ir_events_url,
                    company.ir_scraper_type or "generic",
                )
            except Exception as e:
                logger.error("IR scrape failed for %s: %s", company.ticker, e)
                continue

            for ir_event in ir_events:
                event_date = ir_event.get("event_date")
                if not event_date or event_date < datetime.utcnow().date():
                    continue

                # Try to match with existing DB event
                existing = (
                    db.query(Event)
                    .filter(
                        Event.ticker == company.ticker,
                        Event.event_date == event_date,
                    )
                    .first()
                )

                if existing:
                    # Verify and enrich
                    existing.ir_verified = True
                    existing.status = "confirmed"
                    if ir_event.get("title"):
                        existing.title = ir_event["title"]
                    if ir_event.get("webcast_url"):
                        existing.webcast_url = ir_event["webcast_url"]
                    if ir_event.get("phone_number"):
                        existing.phone_number = ir_event["phone_number"]
                    if ir_event.get("phone_passcode"):
                        existing.phone_passcode = ir_event["phone_passcode"]
                    if ir_event.get("event_time"):
                        existing.event_time = ir_event["event_time"]
                    existing.source = "ir_page"
                    existing.source_url = company.ir_events_url
                    existing.updated_at = datetime.utcnow()
                    verified += 1
                else:
                    # New event discovered directly from IR page
                    new_event = Event(
                        ticker=company.ticker,
                        company_name=company.company_name,
                        event_type=ir_event.get("event_type", "ad_hoc"),
                        event_date=event_date,
                        event_time=ir_event.get("event_time"),
                        title=ir_event.get("title"),
                        webcast_url=ir_event.get("webcast_url"),
                        phone_number=ir_event.get("phone_number"),
                        phone_passcode=ir_event.get("phone_passcode"),
                        source="ir_page",
                        source_url=company.ir_events_url,
                        ir_verified=True,
                        status="confirmed",
                    )
                    db.add(new_event)
                    new_discovered += 1

            company.last_scraped = datetime.utcnow()

        db.commit()
        result = {"verified": verified, "new_discovered": new_discovered, "companies_checked": len(companies)}
        logger.info("IR verification: %s", result)
        return result

    finally:
        if close_session:
            db.close()


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    result = verify_events_on_ir()
    print(f"IR verification complete: {result}")
