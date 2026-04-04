"""Watchlist upload and subscription router."""
from datetime import datetime
from typing import Optional

from fastapi import APIRouter, Depends, UploadFile, File, Form, Query, HTTPException
from sqlalchemy.orm import Session

from database import get_db
from models import Event, Company, WatchlistSubscription
from schemas import SubscribeRequest, SubscribeResponse, WatchlistUploadResponse
from security import generate_feed_token, verify_unsubscribe_token
from services.csv_processor import process_watchlist_csv
from config import get_settings

router = APIRouter(prefix="/api/v1", tags=["watchlist"])
settings = get_settings()


@router.post("/watchlist", response_model=WatchlistUploadResponse)
async def upload_watchlist(
    csv_file: UploadFile = File(...),
    email: Optional[str] = Form(None),
    calendar_type: Optional[str] = Form("outlook"),
    db: Session = Depends(get_db),
):
    """Upload CSV watchlist to filter events and optionally subscribe for auto-invites.

    CSV must have a 'ticker' (or 'symbol') column. Other columns are ignored.
    If email is provided, creates a subscription for automatic calendar invites.
    """
    content = await csv_file.read()
    result = await process_watchlist_csv(content)

    if result["errors"]:
        raise HTTPException(status_code=400, detail=result["errors"])

    tickers = result["tickers"]
    if not tickers:
        raise HTTPException(status_code=400, detail="No valid tickers found in CSV")

    # Check which tickers we know about
    known = db.query(Company.ticker).filter(Company.ticker.in_(tickers)).all()
    known_tickers = {r[0] for r in known}
    unknown_tickers = [t for t in tickers if t not in known_tickers]
    found_tickers = [t for t in tickers if t in known_tickers]

    # Count confirmed vs pending events
    today = datetime.utcnow().date()
    confirmed = (
        db.query(Event)
        .filter(Event.ticker.in_(found_tickers), Event.ir_verified == True, Event.event_date >= today)
        .count()
    )
    pending = (
        db.query(Event)
        .filter(Event.ticker.in_(found_tickers), Event.ir_verified == False, Event.event_date >= today)
        .count()
    )

    feed_url = None
    message = f"Found {len(found_tickers)} tickers with {confirmed} confirmed events."

    # Create subscription if email provided
    if email:
        sub = db.query(WatchlistSubscription).filter(WatchlistSubscription.email == email.lower()).first()
        if sub:
            sub.tickers = found_tickers
            sub.calendar_type = calendar_type or "outlook"
            sub.is_active = True
        else:
            sub = WatchlistSubscription(
                email=email.lower(),
                tickers=found_tickers,
                calendar_type=calendar_type or "outlook",
                feed_token=generate_feed_token(),
                is_active=True,
            )
            db.add(sub)
        db.commit()
        db.refresh(sub)

        feed_url = f"{settings.site_url}/api/v1/feed/{sub.feed_token}.ics"
        message += " Subscribed for auto-invites. You can also paste the feed URL into your calendar app."

    return WatchlistUploadResponse(
        tickers_found=found_tickers,
        tickers_unknown=unknown_tickers,
        events_confirmed=confirmed,
        events_pending=pending,
        feed_url=feed_url,
        message=message,
    )


@router.post("/subscribe", response_model=SubscribeResponse)
def subscribe(
    req: SubscribeRequest,
    db: Session = Depends(get_db),
):
    """Subscribe to auto-invites for a list of tickers.

    Returns a subscribable .ics feed URL and sends calendar invites
    for any events already confirmed.
    """
    # Validate tickers exist
    known = db.query(Company.ticker).filter(Company.ticker.in_(req.tickers)).all()
    known_tickers = [r[0] for r in known]

    if not known_tickers:
        raise HTTPException(status_code=400, detail="None of the provided tickers are in our database")

    # Upsert subscription
    sub = db.query(WatchlistSubscription).filter(WatchlistSubscription.email == req.email).first()
    if sub:
        sub.tickers = known_tickers
        sub.calendar_type = req.calendar_type
        sub.is_active = True
    else:
        sub = WatchlistSubscription(
            email=req.email,
            tickers=known_tickers,
            calendar_type=req.calendar_type,
            feed_token=generate_feed_token(),
            is_active=True,
        )
        db.add(sub)

    db.commit()
    db.refresh(sub)

    # Count events
    today = datetime.utcnow().date()
    confirmed = (
        db.query(Event)
        .filter(Event.ticker.in_(known_tickers), Event.ir_verified == True, Event.event_date >= today)
        .count()
    )
    pending = len(known_tickers) - confirmed  # Approximate: tickers without confirmed events

    feed_url = f"{settings.site_url}/api/v1/feed/{sub.feed_token}.ics"

    return SubscribeResponse(
        feed_url=feed_url,
        events_confirmed=confirmed,
        events_pending=max(0, pending),
        message=f"Subscribed! {confirmed} events ready now. You'll receive calendar invites as more are confirmed.",
    )


@router.get("/unsubscribe")
def unsubscribe(
    token: str = Query(...),
    db: Session = Depends(get_db),
):
    """One-click unsubscribe via signed JWT token (included in every email)."""
    email = verify_unsubscribe_token(token)
    if not email:
        raise HTTPException(status_code=400, detail="Invalid or expired unsubscribe link")

    sub = db.query(WatchlistSubscription).filter(WatchlistSubscription.email == email).first()
    if sub:
        sub.is_active = False
        db.commit()

    return {"message": "Successfully unsubscribed. You will no longer receive calendar invites."}
