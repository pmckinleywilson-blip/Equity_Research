"""Watchlist auto-invite notification service.

After each scraper run, finds newly confirmed events and sends
direct calendar invites to all matching subscribers.
"""
import logging
from datetime import datetime
from sqlalchemy.orm import Session
from sqlalchemy import and_, not_

from models import Event, WatchlistSubscription, NotificationLog
from services.ics_generator import generate_invite_ics
from services.email_service import send_calendar_invite
from security import generate_unsubscribe_token
from config import get_settings

logger = logging.getLogger(__name__)
settings = get_settings()


async def notify_new_events(db: Session) -> dict:
    """Find newly confirmed events and send invites to matching subscribers.

    Returns: {"events_processed": N, "invites_sent": N, "errors": N}
    """
    # Find confirmed events that haven't been notified yet
    new_events = (
        db.query(Event)
        .filter(
            Event.ir_verified == True,
            Event.notified_at.is_(None),
            Event.event_date >= datetime.utcnow().date(),
        )
        .all()
    )

    if not new_events:
        return {"events_processed": 0, "invites_sent": 0, "errors": 0}

    # Get all active subscriptions
    subscriptions = (
        db.query(WatchlistSubscription)
        .filter(WatchlistSubscription.is_active == True)
        .all()
    )

    invites_sent = 0
    errors = 0

    for event in new_events:
        matching_subs = [
            sub for sub in subscriptions
            if event.ticker in (sub.tickers or [])
        ]

        for sub in matching_subs:
            # Check if already notified
            existing = (
                db.query(NotificationLog)
                .filter(
                    NotificationLog.subscription_id == sub.id,
                    NotificationLog.event_id == event.id,
                )
                .first()
            )
            if existing:
                continue

            # Generate and send invite
            ics_content = generate_invite_ics(event, sub.email)
            unsubscribe_token = generate_unsubscribe_token(sub.email)
            unsubscribe_url = f"{settings.site_url}/api/v1/unsubscribe?token={unsubscribe_token}"

            title = event.title or f"{event.ticker} {event.event_type.replace('_', ' ').title()}"
            subject = f"Calendar Invite: {title} - {event.event_date.strftime('%b %d, %Y')}"

            body_parts = [f"<strong>{title}</strong> has been confirmed."]
            if event.webcast_url:
                body_parts.append(f'<br>Webcast: <a href="{event.webcast_url}">{event.webcast_url}</a>')
            if event.phone_number:
                body_parts.append(f"<br>Dial-in: {event.phone_number}")
                if event.phone_passcode:
                    body_parts.append(f" (Passcode: {event.phone_passcode})")
            body_parts.append(f"<br>Date: {event.event_date.strftime('%B %d, %Y')}")
            if event.event_time:
                body_parts.append(f" at {event.event_time.strftime('%I:%M %p')} {event.timezone}")

            success = await send_calendar_invite(
                to_email=sub.email,
                subject=subject,
                body_text="".join(body_parts),
                ics_content=ics_content,
                unsubscribe_url=unsubscribe_url,
            )

            if success:
                # Log notification
                log_entry = NotificationLog(
                    subscription_id=sub.id,
                    event_id=event.id,
                )
                db.add(log_entry)
                invites_sent += 1
            else:
                errors += 1

        # Mark event as notified
        event.notified_at = datetime.utcnow()

    db.commit()

    result = {
        "events_processed": len(new_events),
        "invites_sent": invites_sent,
        "errors": errors,
    }
    logger.info("Notification run: %s", result)
    return result
