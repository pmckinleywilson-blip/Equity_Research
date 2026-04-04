"""Detection Tier: Earnings calendar scrapers (Nasdaq + Finnhub + Alpha Vantage).

These are DETECTION sources — they find potential event dates.
Events are stored as 'tentative' until verified on the company IR page.
"""
import logging
import sys
from datetime import datetime, date, timedelta
from pathlib import Path

import httpx

sys.path.insert(0, str(Path(__file__).parent.parent / "backend"))

from database import SessionLocal, init_db
from models import Event, Company
from config import get_settings

logger = logging.getLogger(__name__)
settings = get_settings()


def scrape_nasdaq_calendar(date_str: str = None) -> list[dict]:
    """Fetch earnings calendar from Nasdaq API.

    Free, no key needed. Returns dates for all US-listed companies.
    """
    if date_str is None:
        date_str = datetime.utcnow().strftime("%Y-%m-%d")

    url = "https://api.nasdaq.com/api/calendar/earnings"
    params = {"date": date_str}
    headers = {
        "User-Agent": "Mozilla/5.0 (compatible; SP500Events/1.0)",
        "Accept": "application/json",
    }

    try:
        resp = httpx.get(url, params=params, headers=headers, timeout=30)
        resp.raise_for_status()
        data = resp.json()
    except Exception as e:
        logger.error("Nasdaq calendar request failed: %s", e)
        return []

    rows = data.get("data", {}).get("rows", [])
    events = []
    for row in rows:
        ticker = row.get("symbol", "").strip()
        name = row.get("name", "").strip()
        time_str = row.get("time", "")
        fiscal_end = row.get("fiscalQuarterEnding", "")

        if not ticker or not name:
            continue

        # Parse time: "time-not-supplied", "time-pre-market", "time-after-hours", or HH:MM
        event_time = None
        if ":" in time_str:
            try:
                event_time = datetime.strptime(time_str, "%H:%M").time()
            except ValueError:
                pass

        events.append({
            "ticker": ticker,
            "company_name": name,
            "event_type": "earnings",
            "event_date": date_str,
            "event_time": event_time,
            "fiscal_quarter": fiscal_end,
            "source": "nasdaq",
        })

    logger.info("Nasdaq: found %d earnings events for %s", len(events), date_str)
    return events


def scrape_finnhub_calendar(from_date: str = None, to_date: str = None) -> list[dict]:
    """Fetch earnings calendar from Finnhub API.

    Free tier: 60 calls/min. Returns dates + EPS estimates.
    """
    if not settings.finnhub_api_key:
        logger.warning("No Finnhub API key configured — skipping")
        return []

    if from_date is None:
        from_date = datetime.utcnow().strftime("%Y-%m-%d")
    if to_date is None:
        to_date = (datetime.utcnow() + timedelta(days=30)).strftime("%Y-%m-%d")

    url = "https://finnhub.io/api/v1/calendar/earnings"
    params = {
        "from": from_date,
        "to": to_date,
        "token": settings.finnhub_api_key,
    }

    try:
        resp = httpx.get(url, params=params, timeout=30)
        resp.raise_for_status()
        data = resp.json()
    except Exception as e:
        logger.error("Finnhub calendar request failed: %s", e)
        return []

    earnings = data.get("earningsCalendar", [])
    events = []
    for item in earnings:
        ticker = item.get("symbol", "").strip()
        event_date = item.get("date", "")
        quarter = item.get("quarter")
        year = item.get("year")

        if not ticker or not event_date:
            continue

        fiscal_q = f"Q{quarter} {year}" if quarter and year else None

        # Finnhub uses "bmo" (before market open) / "amc" (after market close)
        hour_str = item.get("hour", "")
        event_time = None
        if hour_str == "bmo":
            from datetime import time
            event_time = time(8, 0)  # Approximate
        elif hour_str == "amc":
            from datetime import time
            event_time = time(16, 30)  # Approximate

        events.append({
            "ticker": ticker,
            "company_name": ticker,  # Finnhub doesn't return company name
            "event_type": "earnings",
            "event_date": event_date,
            "event_time": event_time,
            "fiscal_quarter": fiscal_q,
            "source": "finnhub",
        })

    logger.info("Finnhub: found %d earnings events for %s to %s", len(events), from_date, to_date)
    return events


def store_detected_events(events: list[dict], db=None):
    """Store detected events in the database as 'tentative'.

    Only inserts new events. Does NOT overwrite IR-verified events.
    """
    close_session = False
    if db is None:
        init_db()
        db = SessionLocal()
        close_session = True

    try:
        # Get all known tickers
        known_tickers = {r[0] for r in db.query(Company.ticker).all()}

        inserted = 0
        skipped = 0
        for ev in events:
            ticker = ev["ticker"].upper()
            if ticker not in known_tickers:
                skipped += 1
                continue

            event_date = ev["event_date"]
            if isinstance(event_date, str):
                event_date = datetime.strptime(event_date, "%Y-%m-%d").date()

            # Check if event already exists
            existing = (
                db.query(Event)
                .filter(
                    Event.ticker == ticker,
                    Event.event_date == event_date,
                    Event.event_type == ev["event_type"],
                )
                .first()
            )

            if existing:
                # Don't overwrite IR-verified data
                if existing.ir_verified:
                    skipped += 1
                    continue
                # Update tentative event with new data (keep existing if better)
                if ev.get("event_time") and not existing.event_time:
                    existing.event_time = ev["event_time"]
                if ev.get("fiscal_quarter") and not existing.fiscal_quarter:
                    existing.fiscal_quarter = ev["fiscal_quarter"]
                existing.updated_at = datetime.utcnow()
                skipped += 1
            else:
                # Look up company name
                company = db.query(Company).filter(Company.ticker == ticker).first()
                company_name = company.company_name if company else ev.get("company_name", ticker)

                new_event = Event(
                    ticker=ticker,
                    company_name=company_name,
                    event_type=ev["event_type"],
                    event_date=event_date,
                    event_time=ev.get("event_time"),
                    fiscal_quarter=ev.get("fiscal_quarter"),
                    source=ev["source"],
                    ir_verified=False,
                    status="tentative",
                )
                db.add(new_event)
                inserted += 1

        db.commit()
        logger.info("Stored events: %d inserted, %d skipped", inserted, skipped)
        return {"inserted": inserted, "skipped": skipped}

    finally:
        if close_session:
            db.close()


def run_detection(days_ahead: int = 90):
    """Run all detection sources and store results."""
    init_db()
    db = SessionLocal()

    all_events = []

    # Nasdaq: scan day by day for next N days
    today = datetime.utcnow().date()
    for i in range(days_ahead):
        d = today + timedelta(days=i)
        if d.weekday() >= 5:  # Skip weekends
            continue
        try:
            events = scrape_nasdaq_calendar(d.strftime("%Y-%m-%d"))
            all_events.extend(events)
        except Exception as e:
            logger.error("Nasdaq scrape failed for %s: %s", d, e)

    # Finnhub: one call for the whole range
    try:
        to_date = (today + timedelta(days=days_ahead)).strftime("%Y-%m-%d")
        events = scrape_finnhub_calendar(today.strftime("%Y-%m-%d"), to_date)
        all_events.extend(events)
    except Exception as e:
        logger.error("Finnhub scrape failed: %s", e)

    result = store_detected_events(all_events, db)
    db.close()
    return result


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    result = run_detection(days_ahead=30)
    print(f"Detection complete: {result}")
