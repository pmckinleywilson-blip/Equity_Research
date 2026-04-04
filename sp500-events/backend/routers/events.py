"""Events API router — the primary machine-readable interface."""
from datetime import date, datetime
from typing import Optional

from fastapi import APIRouter, Depends, Query, HTTPException
from fastapi.responses import Response
from sqlalchemy.orm import Session
from sqlalchemy import func

from database import get_db
from models import Event, Company
from schemas import EventResponse, EventListResponse, CompanyResponse, CompanyListResponse, HealthResponse

router = APIRouter(prefix="/api/v1", tags=["events"])


@router.get("/events", response_model=EventListResponse)
def list_events(
    ticker: Optional[str] = Query(None, description="Comma-separated tickers: AAPL,MSFT"),
    q: Optional[str] = Query(None, description="Search by ticker or company name"),
    date_from: Optional[date] = Query(None, description="Start date (YYYY-MM-DD)"),
    date_to: Optional[date] = Query(None, description="End date (YYYY-MM-DD)"),
    type: Optional[str] = Query(None, description="Event type: earnings, investor_day, conference, ad_hoc"),
    sector: Optional[str] = Query(None, description="Sector filter"),
    index: Optional[str] = Query(None, description="Index: sp500, russell3000"),
    confirmed_only: bool = Query(False, description="Only IR-verified events"),
    page: int = Query(1, ge=1),
    per_page: int = Query(100, ge=1, le=5000),
    format: Optional[str] = Query(None, description="Output format: json (default), rss"),
    db: Session = Depends(get_db),
):
    """List upcoming events with filters. Primary endpoint for AI agents."""
    query = db.query(Event).filter(Event.event_date >= datetime.utcnow().date())

    if ticker:
        tickers = [t.strip().upper() for t in ticker.split(",") if t.strip()]
        query = query.filter(Event.ticker.in_(tickers))

    # Search by ticker OR company name
    if q:
        search = q.strip()
        query = query.filter(
            (Event.ticker.ilike(f"%{search}%")) |
            (Event.company_name.ilike(f"%{search}%"))
        )

    if date_from:
        query = query.filter(Event.event_date >= date_from)
    if date_to:
        query = query.filter(Event.event_date <= date_to)
    if type:
        query = query.filter(Event.event_type == type)
    if confirmed_only:
        query = query.filter(Event.ir_verified == True)
    if sector:
        query = query.join(Company, Event.ticker == Company.ticker).filter(
            Company.sector.ilike(f"%{sector}%")
        )
    if index:
        if index == "sp500":
            query = query.join(Company, Event.ticker == Company.ticker).filter(
                Company.market_cap_tier == "sp500"
            )

    total = query.count()
    events = (
        query.order_by(Event.event_date.asc(), Event.ticker.asc())
        .offset((page - 1) * per_page)
        .limit(per_page)
        .all()
    )

    if format == "rss":
        return _events_to_rss(events)

    event_responses = [EventResponse.model_validate(e) for e in events]

    # If ticker was searched but no events found, check if company exists
    # and return a placeholder so users know the company is tracked
    if ticker and total == 0:
        tickers = [t.strip().upper() for t in ticker.split(",") if t.strip()]
        for t in tickers:
            company = db.query(Company).filter(Company.ticker == t).first()
            if company:
                event_responses.append(EventResponse(
                    id=0,
                    ticker=company.ticker,
                    company_name=company.company_name,
                    event_type="earnings",
                    event_date=date.today(),
                    event_time=None,
                    timezone="America/New_York",
                    title="No events announced yet",
                    description=None,
                    webcast_url=None,
                    phone_number=None,
                    phone_passcode=None,
                    replay_url=None,
                    fiscal_quarter=None,
                    source="none",
                    source_url=None,
                    ir_verified=False,
                    status="pending",
                    created_at=datetime.utcnow(),
                    updated_at=datetime.utcnow(),
                ))
                total += 1

    pages = (total + per_page - 1) // per_page
    return EventListResponse(
        events=event_responses,
        total=total,
        page=page,
        per_page=per_page,
        pages=pages,
    )


@router.get("/events/{ticker}", response_model=list[EventResponse])
def get_events_by_ticker(
    ticker: str,
    db: Session = Depends(get_db),
):
    """Get all upcoming events for a specific ticker."""
    events = (
        db.query(Event)
        .filter(
            Event.ticker == ticker.upper(),
            Event.event_date >= datetime.utcnow().date(),
        )
        .order_by(Event.event_date.asc())
        .all()
    )
    return [EventResponse.model_validate(e) for e in events]


@router.get("/companies", response_model=CompanyListResponse)
def list_companies(
    index: Optional[str] = Query(None, description="Filter: sp500"),
    sector: Optional[str] = Query(None),
    db: Session = Depends(get_db),
):
    """List all tracked companies (Russell 3000 + S&P 500)."""
    query = db.query(Company)
    if index == "sp500":
        query = query.filter(Company.market_cap_tier == "sp500")
    if sector:
        query = query.filter(Company.sector.ilike(f"%{sector}%"))

    companies = query.order_by(Company.ticker.asc()).all()
    return CompanyListResponse(
        companies=[CompanyResponse.model_validate(c) for c in companies],
        total=len(companies),
    )


@router.get("/health", response_model=HealthResponse)
def health_check(db: Session = Depends(get_db)):
    """Health check with data freshness metrics."""
    total_events = db.query(func.count(Event.id)).scalar() or 0
    confirmed_events = db.query(func.count(Event.id)).filter(Event.ir_verified == True).scalar() or 0
    total_companies = db.query(func.count(Company.ticker)).scalar() or 0
    last_scrape = db.query(func.max(Company.last_scraped)).scalar()

    return HealthResponse(
        status="ok",
        total_events=total_events,
        confirmed_events=confirmed_events,
        total_companies=total_companies,
        last_scrape=last_scrape,
    )


def _events_to_rss(events: list[Event]) -> Response:
    """Convert events to RSS XML feed."""
    items = []
    for e in events:
        title = e.title or f"{e.ticker} {e.event_type.replace('_', ' ').title()}"
        desc_parts = [f"Date: {e.event_date}"]
        if e.event_time:
            desc_parts.append(f"Time: {e.event_time} {e.timezone}")
        if e.webcast_url:
            desc_parts.append(f"Webcast: {e.webcast_url}")
        if e.phone_number:
            desc_parts.append(f"Dial-in: {e.phone_number}")

        items.append(f"""
        <item>
            <title>{_xml_escape(title)}</title>
            <description>{_xml_escape(chr(10).join(desc_parts))}</description>
            <pubDate>{e.created_at.strftime('%a, %d %b %Y %H:%M:%S +0000') if e.created_at else ''}</pubDate>
            <guid>sp500events-{e.id}</guid>
        </item>""")

    rss = f"""<?xml version="1.0" encoding="UTF-8"?>
<rss version="2.0">
<channel>
    <title>SP500 Events Calendar</title>
    <link>https://sp500events.com</link>
    <description>S&amp;P 500 and Russell 3000 earnings dates and corporate events</description>
    {"".join(items)}
</channel>
</rss>"""

    return Response(content=rss, media_type="application/rss+xml")


def _xml_escape(text: str) -> str:
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&apos;")
    )
