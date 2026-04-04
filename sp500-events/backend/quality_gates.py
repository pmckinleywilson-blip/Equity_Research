"""Quality gates — validate all event data before it enters the database.

Every event must pass ALL gates before storage. No exceptions.
Bad data is rejected with a reason, not silently fixed.
"""
from datetime import date, time, timedelta
from typing import Optional
import re
import logging

logger = logging.getLogger(__name__)

# Source authority ranking (higher = more trusted)
SOURCE_PRIORITY = {
    "prnewswire": 5,
    "businesswire": 5,
    "globenewswire": 5,
    "ir_page": 4,
    "nasdaq": 2,
    "historical_pattern": 1,
    "estimated": 0,
    "manual": 3,
}

VALID_EVENT_TYPES = {"earnings", "investor_day", "conference", "ad_hoc"}
VALID_STATUSES = {"confirmed", "tentative", "estimated", "pending", "postponed", "cancelled"}
VALID_SOURCES = set(SOURCE_PRIORITY.keys())

# PR page domains that should NEVER be stored as webcast URLs
PR_PAGE_DOMAINS = {"prnewswire.com", "businesswire.com", "globenewswire.com"}


class ValidationError(Exception):
    """Raised when an event fails quality gates."""
    pass


def validate_event(event_data: dict) -> list[str]:
    """Validate an event dict against all quality gates.

    Returns a list of error messages. Empty list = all gates passed.
    """
    errors = []

    # Gate 1: Required fields
    for field in ("ticker", "event_date", "event_type", "source"):
        if not event_data.get(field):
            errors.append(f"Missing required field: {field}")

    if errors:
        return errors  # Can't check further without required fields

    ticker = event_data["ticker"]
    event_date = event_data["event_date"]

    # Gate 2: Ticker format
    if not re.match(r'^[A-Z0-9.\-]{1,10}$', str(ticker)):
        errors.append(f"Invalid ticker format: {ticker}")

    # Gate 3: Event type must be valid
    if event_data["event_type"] not in VALID_EVENT_TYPES:
        errors.append(f"Invalid event type: {event_data['event_type']}")

    # Gate 4: Source must be recognized
    if event_data["source"] not in VALID_SOURCES:
        errors.append(f"Invalid source: {event_data['source']}")

    # Gate 5: Date must be a weekday (for earnings)
    if isinstance(event_date, date):
        if event_data["event_type"] == "earnings" and event_date.weekday() >= 5:
            errors.append(f"Earnings on weekend: {event_date} ({event_date.strftime('%A')})")

    # Gate 6: Date must be within reasonable range
    if isinstance(event_date, date):
        today = date.today()
        if event_date < today:
            errors.append(f"Date is in the past: {event_date}")
        if event_date > today + timedelta(days=180):
            errors.append(f"Date is more than 6 months away: {event_date}")

    # Gate 7: Time must be reasonable (6 AM - 10 PM ET)
    event_time = event_data.get("event_time")
    if event_time:
        if isinstance(event_time, time):
            if event_time.hour < 6 or event_time.hour > 22:
                errors.append(f"Unreasonable time: {event_time}")
        elif isinstance(event_time, str):
            try:
                parts = event_time.split(":")
                hour = int(parts[0])
                if hour < 6 or hour > 22:
                    errors.append(f"Unreasonable time: {event_time}")
            except (ValueError, IndexError):
                errors.append(f"Invalid time format: {event_time}")

    # Gate 8: Webcast URL must not be a PR page URL
    webcast_url = event_data.get("webcast_url")
    if webcast_url:
        for domain in PR_PAGE_DOMAINS:
            if domain in webcast_url:
                errors.append(f"Webcast URL is a PR page: {webcast_url[:80]}")
                break

    # Gate 9: Webcast URL must look like a URL
    if webcast_url and not webcast_url.startswith("http"):
        errors.append(f"Webcast URL missing protocol: {webcast_url[:80]}")

    return errors


def determine_status(event_data: dict) -> str:
    """Determine the correct status based on available data.

    CONFIRMED = has date + time + webcast URL (all three)
    TENTATIVE = has date, missing time or webcast
    ESTIMATED = source is estimated or historical_pattern
    """
    has_time = bool(event_data.get("event_time"))
    has_webcast = bool(event_data.get("webcast_url"))
    source = event_data.get("source", "")

    if source in ("estimated", "historical_pattern"):
        return "estimated"
    if has_time and has_webcast:
        return "confirmed"
    return "tentative"


def should_update(existing_source: str, new_source: str) -> bool:
    """Determine if a new source should overwrite an existing one.

    Higher priority source always wins.
    Equal priority: new data updates (more recent).
    """
    existing_priority = SOURCE_PRIORITY.get(existing_source, 0)
    new_priority = SOURCE_PRIORITY.get(new_source, 0)
    return new_priority >= existing_priority


def merge_event_data(existing: dict, new: dict) -> dict:
    """Merge new event data into existing, respecting source priority.

    Rules:
    - Higher priority source overwrites lower priority fields
    - Never overwrite populated fields with None/empty from lower source
    - Date from higher source always wins
    - Details (time, webcast, phone) from any source fill gaps
    """
    existing_priority = SOURCE_PRIORITY.get(existing.get("source", ""), 0)
    new_priority = SOURCE_PRIORITY.get(new.get("source", ""), 0)

    merged = dict(existing)

    # Date: higher priority wins, or same priority uses new (more recent)
    if new_priority >= existing_priority and new.get("event_date"):
        merged["event_date"] = new["event_date"]
        merged["source"] = new["source"]
        if new.get("source_url"):
            merged["source_url"] = new["source_url"]

    # Detail fields: fill gaps from any source, overwrite from higher priority
    detail_fields = ["event_time", "webcast_url", "phone_number", "phone_passcode",
                     "title", "fiscal_quarter", "description"]

    for field in detail_fields:
        new_value = new.get(field)
        existing_value = merged.get(field)

        if new_value:
            if not existing_value:
                # Fill gap
                merged[field] = new_value
            elif new_priority > existing_priority:
                # Higher priority overwrites
                merged[field] = new_value

    # Recalculate status
    merged["status"] = determine_status(merged)
    merged["ir_verified"] = merged["status"] == "confirmed"

    return merged


def clean_webcast_url(url: str) -> Optional[str]:
    """Clean and validate a webcast URL. Returns None if invalid."""
    if not url:
        return None

    url = url.strip().rstrip(".,;)")

    # Reject PR page URLs
    for domain in PR_PAGE_DOMAINS:
        if domain in url:
            return None

    # Must start with http
    if not url.startswith("http"):
        return None

    # Must be a reasonable length
    if len(url) > 1000:
        return None

    return url


def clean_time(time_str: str) -> Optional[str]:
    """Clean and validate a time string. Returns HH:MM:SS format or None."""
    if not time_str:
        return None

    try:
        if isinstance(time_str, time):
            return time_str.strftime("%H:%M:%S")

        # Handle HH:MM:SS
        parts = time_str.split(":")
        hour = int(parts[0])
        minute = int(parts[1]) if len(parts) > 1 else 0

        if hour < 0 or hour > 23 or minute < 0 or minute > 59:
            return None
        if hour < 6 or hour > 22:
            return None

        return f"{hour:02d}:{minute:02d}:00"
    except (ValueError, IndexError):
        return None


def fix_weekend_date(d: date) -> date:
    """Shift a weekend date to the nearest weekday (Monday)."""
    if d.weekday() == 5:  # Saturday -> Monday
        return d + timedelta(days=2)
    elif d.weekday() == 6:  # Sunday -> Monday
        return d + timedelta(days=1)
    return d
