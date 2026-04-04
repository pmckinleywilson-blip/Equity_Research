"""Calendar feed router — .ics downloads and subscribable feeds."""
from datetime import datetime
from typing import Optional

from fastapi import APIRouter, Depends, Query, HTTPException
from fastapi.responses import Response
from sqlalchemy.orm import Session

from database import get_db
from models import Event, WatchlistSubscription
from services.ics_generator import generate_bulk_ics, generate_feed_ics, generate_single_ics, generate_gmail_url

router = APIRouter(prefix="/api/v1", tags=["calendar"])


@router.get("/calendar.ics")
def download_calendar(
    tickers: Optional[str] = Query(None, description="Comma-separated tickers"),
    confirmed_only: bool = Query(False),
    db: Session = Depends(get_db),
):
    """Download .ics calendar file with upcoming events.

    Can also be used as a subscribable URL — paste into Outlook/Gmail
    calendar settings to auto-update.
    """
    query = db.query(Event).filter(Event.event_date >= datetime.utcnow().date())

    if tickers:
        ticker_list = [t.strip().upper() for t in tickers.split(",") if t.strip()]
        query = query.filter(Event.ticker.in_(ticker_list))
    if confirmed_only:
        query = query.filter(Event.ir_verified == True)

    events = query.order_by(Event.event_date.asc()).limit(1000).all()
    ics_content = generate_bulk_ics(events)

    return Response(
        content=ics_content,
        media_type="text/calendar",
        headers={
            "Content-Disposition": "attachment; filename=sp500-events.ics",
        },
    )


@router.get("/calendar/{event_id}.ics")
def download_single_event_ics(
    event_id: int,
    db: Session = Depends(get_db),
):
    """Download .ics for a single event (Outlook one-click add)."""
    event = db.query(Event).filter(Event.id == event_id).first()
    if not event:
        raise HTTPException(status_code=404, detail="Event not found")

    ics_content = generate_single_ics(event)
    title = (event.title or f"{event.ticker}_{event.event_type}").replace(" ", "_")

    return Response(
        content=ics_content,
        media_type="text/calendar",
        headers={
            "Content-Disposition": f"attachment; filename={title}.ics",
        },
    )


@router.get("/calendar/{event_id}/gmail")
def get_gmail_url(
    event_id: int,
    db: Session = Depends(get_db),
):
    """Get Google Calendar 'add event' URL for one-click add from Gmail."""
    event = db.query(Event).filter(Event.id == event_id).first()
    if not event:
        raise HTTPException(status_code=404, detail="Event not found")

    url = generate_gmail_url(event)
    return {"gmail_url": url}


@router.get("/feed/{token}.ics")
def subscribable_feed(
    token: str,
    db: Session = Depends(get_db),
):
    """Subscribable personal calendar feed.

    Paste this URL into Outlook or Gmail calendar settings:
    - Outlook: File → Account Settings → Internet Calendars → paste URL
    - Gmail: Settings → Add calendar → From URL → paste URL

    Calendar app polls automatically. New events appear on refresh.
    """
    sub = (
        db.query(WatchlistSubscription)
        .filter(
            WatchlistSubscription.feed_token == token,
            WatchlistSubscription.is_active == True,
        )
        .first()
    )
    if not sub:
        raise HTTPException(status_code=404, detail="Feed not found or inactive")

    tickers = sub.tickers or []
    events = (
        db.query(Event)
        .filter(
            Event.ticker.in_(tickers),
            Event.event_date >= datetime.utcnow().date(),
        )
        .order_by(Event.event_date.asc())
        .all()
    )

    ics_content = generate_feed_ics(events)

    return Response(
        content=ics_content,
        media_type="text/calendar",
        headers={
            "Cache-Control": "max-age=3600",  # Cache for 1 hour
        },
    )


@router.post("/calendar/bulk.ics")
async def download_bulk_ics(
    event_ids: list[int],
    db: Session = Depends(get_db),
):
    """Download a single .ics file containing multiple events (bulk add to Outlook)."""
    if len(event_ids) > 500:
        raise HTTPException(status_code=400, detail="Maximum 500 events per bulk download")

    events = db.query(Event).filter(Event.id.in_(event_ids)).order_by(Event.event_date.asc()).all()
    if not events:
        raise HTTPException(status_code=404, detail="No events found")

    ics_content = generate_bulk_ics(events)

    return Response(
        content=ics_content,
        media_type="text/calendar",
        headers={
            "Content-Disposition": "attachment; filename=sp500-events-bulk.ics",
        },
    )
