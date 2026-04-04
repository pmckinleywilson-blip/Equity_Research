"""Unified event store — the single point of entry for all event data.

ALL sources (wire services, Nasdaq, IR pages, historical patterns) must
go through this module to store or update events. This ensures quality
gates are applied consistently and conflict resolution follows the
source hierarchy.

Usage:
    from event_store import upsert_event

    result = upsert_event(db, {
        "ticker": "AAPL",
        "event_type": "earnings",
        "event_date": date(2026, 4, 30),
        "event_time": "17:00:00",
        "webcast_url": "https://investor.apple.com",
        "source": "businesswire",
        "source_url": "https://www.businesswire.com/...",
    })
    # result: "inserted", "updated", "skipped", or "rejected"
"""
import logging
from datetime import date, datetime, time, timedelta
from typing import Optional

from sqlalchemy.orm import Session

from models import Event, Company
from quality_gates import (
    validate_event, determine_status, should_update, merge_event_data,
    clean_webcast_url, clean_time, fix_weekend_date,
    SOURCE_PRIORITY,
)

logger = logging.getLogger(__name__)


def upsert_event(db: Session, event_data: dict) -> str:
    """Insert or update an event, applying all quality gates.

    Returns: "inserted", "updated", "skipped", or "rejected:<reason>"
    """
    # Step 1: Clean input data
    event_data = _clean_input(event_data)

    # Step 2: Validate against quality gates
    errors = validate_event(event_data)
    if errors:
        ticker = event_data.get("ticker", "???")
        logger.debug("Rejected %s: %s", ticker, errors)
        return f"rejected:{'; '.join(errors)}"

    ticker = event_data["ticker"]
    event_date = event_data["event_date"]
    event_type = event_data["event_type"]
    source = event_data["source"]

    # Step 3: Check ticker exists in company database
    company = db.query(Company).filter(Company.ticker == ticker).first()
    if not company:
        return "rejected:ticker not in company database"

    # Step 4: Look for existing event — exact date match
    existing = (
        db.query(Event)
        .filter(
            Event.ticker == ticker,
            Event.event_date == event_date,
            Event.event_type == event_type,
        )
        .first()
    )

    if existing:
        return _update_existing(db, existing, event_data)

    # Step 5: Look for nearby event (within 14 days) from lower-priority source
    # This handles date corrections (e.g., wire says Apr 22, Nasdaq said Apr 28)
    nearby = (
        db.query(Event)
        .filter(
            Event.ticker == ticker,
            Event.event_type == event_type,
            Event.event_date >= event_date - timedelta(days=14),
            Event.event_date <= event_date + timedelta(days=14),
        )
        .first()
    )

    if nearby:
        if should_update(nearby.source, source):
            return _update_existing(db, nearby, event_data, correct_date=True)
        else:
            # Lower priority source has a different date — skip
            return "skipped"

    # Step 6: No existing event — insert new
    return _insert_new(db, event_data, company)


def _clean_input(data: dict) -> dict:
    """Clean and normalize input data."""
    cleaned = dict(data)

    # Normalize ticker
    if cleaned.get("ticker"):
        cleaned["ticker"] = str(cleaned["ticker"]).strip().upper()

    # Parse date if string
    if isinstance(cleaned.get("event_date"), str):
        try:
            cleaned["event_date"] = datetime.strptime(cleaned["event_date"], "%Y-%m-%d").date()
        except ValueError:
            pass

    # Fix weekend dates
    if isinstance(cleaned.get("event_date"), date) and cleaned.get("event_type") == "earnings":
        cleaned["event_date"] = fix_weekend_date(cleaned["event_date"])

    # Clean time
    if cleaned.get("event_time"):
        cleaned["event_time"] = clean_time(cleaned["event_time"])

    # Clean webcast URL
    if cleaned.get("webcast_url"):
        cleaned["webcast_url"] = clean_webcast_url(cleaned["webcast_url"])

    # Determine status
    cleaned["status"] = determine_status(cleaned)

    return cleaned


def _update_existing(db: Session, existing: Event, new_data: dict, correct_date: bool = False) -> str:
    """Update an existing event with new data from a (possibly higher-priority) source."""

    # Build dicts for merge
    existing_dict = {
        "ticker": existing.ticker,
        "event_type": existing.event_type,
        "event_date": existing.event_date,
        "event_time": existing.event_time.strftime("%H:%M:%S") if existing.event_time else None,
        "webcast_url": existing.webcast_url,
        "phone_number": existing.phone_number,
        "phone_passcode": existing.phone_passcode,
        "title": existing.title,
        "fiscal_quarter": existing.fiscal_quarter,
        "source": existing.source,
        "source_url": existing.source_url,
    }

    merged = merge_event_data(existing_dict, new_data)

    # Apply merged data to the existing record
    changed = False

    if correct_date and merged["event_date"] != existing.event_date:
        existing.event_date = merged["event_date"]
        changed = True

    if merged.get("event_time") and merged["event_time"] != (existing.event_time.strftime("%H:%M:%S") if existing.event_time else None):
        if isinstance(merged["event_time"], str):
            existing.event_time = datetime.strptime(merged["event_time"], "%H:%M:%S").time()
        elif isinstance(merged["event_time"], time):
            existing.event_time = merged["event_time"]
        changed = True

    for field in ["webcast_url", "phone_number", "phone_passcode", "title", "fiscal_quarter"]:
        new_val = merged.get(field)
        old_val = getattr(existing, field)
        if new_val and new_val != old_val:
            setattr(existing, field, new_val)
            changed = True

    # Update source if higher priority provided data
    if should_update(existing.source, new_data["source"]):
        existing.source = merged["source"]
        if merged.get("source_url"):
            existing.source_url = merged["source_url"]

    # Recalculate status
    new_status = determine_status({
        "event_time": existing.event_time,
        "webcast_url": existing.webcast_url,
        "source": existing.source,
    })
    if new_status != existing.status:
        existing.status = new_status
        existing.ir_verified = new_status == "confirmed"
        changed = True

    existing.updated_at = datetime.utcnow()

    if changed:
        return "updated"
    return "skipped"


def _insert_new(db: Session, event_data: dict, company: Company) -> str:
    """Insert a new event."""
    event_time = None
    if event_data.get("event_time"):
        if isinstance(event_data["event_time"], str):
            event_time = datetime.strptime(event_data["event_time"], "%H:%M:%S").time()
        elif isinstance(event_data["event_time"], time):
            event_time = event_data["event_time"]

    status = determine_status(event_data)

    new_event = Event(
        ticker=event_data["ticker"],
        company_name=company.company_name,
        event_type=event_data["event_type"],
        event_date=event_data["event_date"],
        event_time=event_time,
        timezone=event_data.get("timezone", "America/New_York"),
        title=event_data.get("title", "")[:500] if event_data.get("title") else None,
        description=event_data.get("description"),
        webcast_url=event_data.get("webcast_url"),
        phone_number=event_data.get("phone_number"),
        phone_passcode=event_data.get("phone_passcode"),
        fiscal_quarter=event_data.get("fiscal_quarter"),
        source=event_data["source"],
        source_url=event_data.get("source_url"),
        ir_verified=status == "confirmed",
        status=status,
    )
    db.add(new_event)
    return "inserted"


def bulk_upsert(db: Session, events: list[dict]) -> dict:
    """Upsert multiple events, returning counts.

    Returns: {"inserted": N, "updated": N, "skipped": N, "rejected": N}
    """
    counts = {"inserted": 0, "updated": 0, "skipped": 0, "rejected": 0}

    for event_data in events:
        result = upsert_event(db, event_data)
        if result.startswith("rejected"):
            counts["rejected"] += 1
        else:
            counts[result] += 1

    db.commit()
    logger.info("Bulk upsert: %s", counts)
    return counts
