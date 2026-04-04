"""Generate .ics calendar files for events.

Supports:
- Single event .ics download (METHOD:PUBLISH)
- Bulk .ics download with multiple events (METHOD:PUBLISH)
- Subscribable feed .ics (METHOD:PUBLISH, dynamically generated)
- Direct calendar invite .ics (METHOD:REQUEST, one ATTENDEE per file)
"""
from datetime import datetime, timedelta
from icalendar import Calendar, Event as ICalEvent, vText
from models import Event


def _build_vevent(event: Event) -> ICalEvent:
    """Build a VEVENT component from a database Event."""
    vevent = ICalEvent()
    vevent.add("summary", event.title or f"{event.ticker} {event.event_type.replace('_', ' ').title()}")
    vevent.add("uid", f"sp500events-{event.id}@sp500events.com")

    # Date/time
    if event.event_time:
        from zoneinfo import ZoneInfo
        tz = ZoneInfo(event.timezone or "America/New_York")
        dt_start = datetime.combine(event.event_date, event.event_time).replace(tzinfo=tz)
        dt_end = dt_start + timedelta(hours=1)
        vevent.add("dtstart", dt_start)
        vevent.add("dtend", dt_end)
    else:
        vevent.add("dtstart", event.event_date)
        vevent.add("dtend", event.event_date + timedelta(days=1))

    # Description with full join details
    desc_parts = []
    if event.webcast_url:
        desc_parts.append(f"Join Webcast: {event.webcast_url}")
    if event.phone_number:
        line = f"Dial-in: {event.phone_number}"
        if event.phone_passcode:
            line += f"\nPasscode: {event.phone_passcode}"
        desc_parts.append(line)
    if event.replay_url:
        desc_parts.append(f"Replay: {event.replay_url}")
    if event.fiscal_quarter:
        desc_parts.append(f"Period: {event.fiscal_quarter}")
    desc_parts.append(f"\nSource: sp500events.com | Status: {event.status}")
    if event.description:
        desc_parts.append(f"\n{event.description}")

    vevent.add("description", "\n".join(desc_parts))

    # URL (clickable in calendar apps)
    if event.webcast_url:
        vevent.add("url", event.webcast_url)
        vevent.add("location", event.webcast_url)

    vevent.add("status", "CONFIRMED" if event.status == "confirmed" else "TENTATIVE")
    vevent.add("dtstamp", datetime.utcnow())
    vevent.add("created", event.created_at or datetime.utcnow())
    vevent.add("last-modified", event.updated_at or datetime.utcnow())

    return vevent


def generate_single_ics(event: Event) -> bytes:
    """Generate a downloadable .ics for a single event."""
    cal = Calendar()
    cal.add("prodid", "-//SP500Events//EN")
    cal.add("version", "2.0")
    cal.add("method", "PUBLISH")
    cal.add("x-wr-calname", "SP500 Events")
    cal.add_component(_build_vevent(event))
    return cal.to_ical()


def generate_bulk_ics(events: list[Event]) -> bytes:
    """Generate a single .ics file containing multiple events."""
    cal = Calendar()
    cal.add("prodid", "-//SP500Events//EN")
    cal.add("version", "2.0")
    cal.add("method", "PUBLISH")
    cal.add("x-wr-calname", "SP500 Events")

    for event in events:
        cal.add_component(_build_vevent(event))

    return cal.to_ical()


def generate_feed_ics(events: list[Event]) -> bytes:
    """Generate a subscribable .ics feed (same as bulk, but intended for subscription)."""
    cal = Calendar()
    cal.add("prodid", "-//SP500Events//EN")
    cal.add("version", "2.0")
    cal.add("method", "PUBLISH")
    cal.add("x-wr-calname", "SP500 Events Watchlist")
    cal.add("x-wr-timezone", "America/New_York")
    # Refresh interval hint for calendar apps (4 hours)
    cal.add("refresh-interval;value=duration", "PT4H")
    cal.add("x-published-ttl", "PT4H")

    for event in events:
        cal.add_component(_build_vevent(event))

    return cal.to_ical()


def generate_invite_ics(event: Event, attendee_email: str) -> bytes:
    """Generate a METHOD:REQUEST .ics for direct calendar invite.

    This creates a proper meeting invite that auto-populates the recipient's calendar.
    Each invite has exactly one ATTENDEE for privacy.
    """
    cal = Calendar()
    cal.add("prodid", "-//SP500Events//EN")
    cal.add("version", "2.0")
    cal.add("method", "REQUEST")

    vevent = _build_vevent(event)

    # Add organizer and single attendee
    from icalendar import vCalAddress
    organizer = vCalAddress(f"mailto:invites@sp500events.com")
    organizer.params["cn"] = vText("SP500 Events")
    vevent.add("organizer", organizer)

    attendee = vCalAddress(f"mailto:{attendee_email}")
    attendee.params["rsvp"] = vText("FALSE")
    attendee.params["partstat"] = vText("NEEDS-ACTION")
    vevent.add("attendee", attendee)

    # Sequence number for updates
    vevent.add("sequence", 0)

    cal.add_component(vevent)
    return cal.to_ical()


def generate_gmail_url(event: Event) -> str:
    """Generate a Google Calendar 'add event' URL for one-click add."""
    from urllib.parse import quote

    title = event.title or f"{event.ticker} {event.event_type.replace('_', ' ').title()}"

    if event.event_time:
        from zoneinfo import ZoneInfo
        tz = ZoneInfo(event.timezone or "America/New_York")
        dt_start = datetime.combine(event.event_date, event.event_time).replace(tzinfo=tz)
        dt_end = dt_start + timedelta(hours=1)
        dates = f"{dt_start.strftime('%Y%m%dT%H%M%S')}/{dt_end.strftime('%Y%m%dT%H%M%S')}"
        ctz = event.timezone or "America/New_York"
    else:
        dates = f"{event.event_date.strftime('%Y%m%d')}/{(event.event_date + timedelta(days=1)).strftime('%Y%m%d')}"
        ctz = ""

    details_parts = []
    if event.webcast_url:
        details_parts.append(f"Webcast: {event.webcast_url}")
    if event.phone_number:
        details_parts.append(f"Dial-in: {event.phone_number}")
        if event.phone_passcode:
            details_parts.append(f"Passcode: {event.phone_passcode}")
    details = quote("\n".join(details_parts))

    location = quote(event.webcast_url or "")

    url = f"https://calendar.google.com/calendar/render?action=TEMPLATE&text={quote(title)}&dates={dates}&details={details}&location={location}"
    if ctz:
        url += f"&ctz={quote(ctz)}"
    return url
