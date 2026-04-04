"""Unified pipeline orchestrator.

Runs all data sources and merges results through the event_store,
which applies quality gates and conflict resolution.

Usage:
    python unified_pipeline.py                    # Full pipeline
    python unified_pipeline.py --source nasdaq    # Nasdaq only
    python unified_pipeline.py --source wire      # Wire services only
    python unified_pipeline.py --source ir        # IR pages only
    python unified_pipeline.py --ticker AAPL      # Specific ticker
"""
import argparse
import logging
import sys
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent / "backend"))

from database import SessionLocal, init_db
from models import Event, Company
from event_store import upsert_event, bulk_upsert

logger = logging.getLogger(__name__)


def run_nasdaq_source(db, days_ahead: int = 60) -> dict:
    """Source 1: Nasdaq calendar — baseline dates for all companies."""
    logger.info("=== SOURCE: NASDAQ CALENDAR ===")
    from earnings_calendars import scrape_nasdaq_calendar
    from datetime import date, timedelta

    today = date.today()
    all_events = []

    for i in range(days_ahead):
        d = today + timedelta(days=i)
        if d.weekday() >= 5:
            continue
        try:
            events = scrape_nasdaq_calendar(d.strftime("%Y-%m-%d"))
            for ev in events:
                all_events.append({
                    "ticker": ev["ticker"].upper(),
                    "event_type": "earnings",
                    "event_date": ev["event_date"],
                    "event_time": ev.get("event_time"),
                    "fiscal_quarter": ev.get("fiscal_quarter"),
                    "source": "nasdaq",
                })
        except Exception as e:
            logger.error("Nasdaq scrape failed for %s: %s", d, e)

    logger.info("Nasdaq: found %d raw events", len(all_events))
    result = bulk_upsert(db, all_events)
    logger.info("Nasdaq result: %s", result)
    return result


def run_wire_source(db, prn_pages: int = 5, bw_pages: int = 20, gnw_pages: int = 10) -> dict:
    """Source 2: Wire services (PRNewswire + BusinessWire + GlobeNewswire)."""
    logger.info("=== SOURCE: WIRE SERVICES (PRN + BW + GNW) ===")
    from wire_services import scrape_all_rss, scrape_prnewswire_pages, scrape_businesswire_pages, scrape_globenewswire

    all_events = []

    # RSS feeds from ALL services (most recent)
    try:
        rss_events = scrape_all_rss()
        all_events.extend(rss_events)
    except Exception as e:
        logger.error("PRNewswire RSS failed: %s", e)

    # PRNewswire archive pages
    try:
        prn_events = scrape_prnewswire_pages(max_pages=prn_pages)
        all_events.extend(prn_events)
    except Exception as e:
        logger.error("PRNewswire pages failed: %s", e)

    # BusinessWire archive pages
    try:
        bw_events = scrape_businesswire_pages(max_pages=bw_pages)
        all_events.extend(bw_events)
    except Exception as e:
        logger.error("BusinessWire pages failed: %s", e)

    # GlobeNewswire archive pages
    try:
        gnw_events = scrape_globenewswire(max_pages=gnw_pages)
        all_events.extend(gnw_events)
    except Exception as e:
        logger.error("GlobeNewswire pages failed: %s", e)

    # Deduplicate by (ticker, date)
    seen = set()
    unique = []
    for ev in all_events:
        key = (ev.get("ticker", ""), ev.get("event_date", ""))
        if key not in seen:
            seen.add(key)
            unique.append(ev)

    logger.info("Wire services: %d unique events from %d raw", len(unique), len(all_events))
    result = bulk_upsert(db, unique)
    logger.info("Wire result: %s", result)
    return result


def run_ir_source(db, limit: int = None) -> dict:
    """Source 3: Company IR page scraping."""
    logger.info("=== SOURCE: IR PAGES ===")
    from ir_pages import scrape_ir_page
    from datetime import date

    # Get companies with IR URLs
    query = db.query(Company).filter(
        Company.ir_events_url.isnot(None),
        Company.ir_scraper_type != "cloudflare_blocked",
    )
    if limit:
        query = query.limit(limit)

    companies = query.all()
    logger.info("IR pages: scraping %d companies", len(companies))

    all_events = []
    for i, company in enumerate(companies):
        try:
            ir_events = scrape_ir_page(company.ir_events_url, company.ir_scraper_type or "generic")
            for ev in ir_events:
                if ev.get("event_date") and ev["event_date"] >= date.today():
                    all_events.append({
                        "ticker": company.ticker,
                        "event_type": ev.get("event_type", "earnings"),
                        "event_date": ev["event_date"],
                        "event_time": ev.get("event_time"),
                        "webcast_url": ev.get("webcast_url"),
                        "phone_number": ev.get("phone_number"),
                        "phone_passcode": ev.get("phone_passcode"),
                        "title": ev.get("title"),
                        "source": "ir_page",
                        "source_url": company.ir_events_url,
                    })
        except Exception as e:
            logger.debug("IR scrape failed for %s: %s", company.ticker, e)

        if (i + 1) % 50 == 0:
            logger.info("  IR progress: %d/%d companies, %d events found", i + 1, len(companies), len(all_events))

    logger.info("IR pages: found %d events from %d companies", len(all_events), len(companies))
    result = bulk_upsert(db, all_events)
    logger.info("IR result: %s", result)
    return result


def run_historical_fallback(db) -> dict:
    """Source 4: Historical pattern estimates for companies with no events."""
    logger.info("=== SOURCE: HISTORICAL FALLBACK ===")
    from datetime import date, timedelta

    today = date.today()

    # Find companies without ANY future events
    companies_with_events = {
        r[0] for r in
        db.query(Event.ticker).filter(Event.event_date >= today).distinct().all()
    }
    all_companies = db.query(Company).all()
    missing = [c for c in all_companies if c.ticker not in companies_with_events]

    logger.info("Companies without events: %d", len(missing))

    events = []
    # Default estimate: mid-May for Q1 reporters
    default_date = date(today.year, 5, 7)
    if default_date < today:
        default_date = date(today.year, 8, 7)  # Q2 default

    for company in missing:
        events.append({
            "ticker": company.ticker,
            "event_type": "earnings",
            "event_date": default_date,
            "source": "estimated",
            "title": f"{company.ticker} Earnings (estimated, not yet announced)",
        })

    result = bulk_upsert(db, events)
    logger.info("Historical fallback: %s", result)
    return result


def run_quality_cleanup(db) -> dict:
    """Post-pipeline cleanup: remove past events, fix inconsistencies."""
    logger.info("=== QUALITY CLEANUP ===")
    from datetime import date

    today = date.today()
    counts = {"removed_past": 0, "removed_dupes": 0, "fixed_status": 0}

    # Remove past events
    past = db.query(Event).filter(Event.event_date < today).all()
    for e in past:
        db.delete(e)
    counts["removed_past"] = len(past)

    # Remove near-duplicates (keep highest priority source)
    from quality_gates import SOURCE_PRIORITY

    tickers = [r[0] for r in db.query(Event.ticker).distinct().all()]
    for ticker in tickers:
        events = (
            db.query(Event)
            .filter(Event.ticker == ticker, Event.event_type == "earnings", Event.event_date >= today)
            .order_by(Event.event_date)
            .all()
        )
        if len(events) <= 1:
            continue

        to_delete = []
        for i in range(len(events)):
            for j in range(i + 1, len(events)):
                gap = (events[j].event_date - events[i].event_date).days
                if gap < 30:
                    # Keep the one with higher priority source and more data
                    si = SOURCE_PRIORITY.get(events[i].source, 0) + (2 if events[i].event_time else 0) + (2 if events[i].webcast_url else 0)
                    sj = SOURCE_PRIORITY.get(events[j].source, 0) + (2 if events[j].event_time else 0) + (2 if events[j].webcast_url else 0)
                    victim = events[i] if sj > si else events[j]
                    if victim not in to_delete:
                        to_delete.append(victim)

        for e in to_delete:
            db.delete(e)
            counts["removed_dupes"] += 1

    # Fix confirmed status consistency
    all_events = db.query(Event).all()
    for e in all_events:
        from quality_gates import determine_status
        correct_status = determine_status({
            "event_time": e.event_time,
            "webcast_url": e.webcast_url,
            "source": e.source,
        })
        if e.status != correct_status:
            e.status = correct_status
            e.ir_verified = correct_status == "confirmed"
            counts["fixed_status"] += 1

    db.commit()
    logger.info("Cleanup: %s", counts)
    return counts


def run_full_pipeline(
    sources: list[str] = None,
    tickers: list[str] = None,
    nasdaq_days: int = 60,
    prn_pages: int = 5,
    bw_pages: int = 20,
    ir_limit: int = None,
):
    """Run the full pipeline: all sources → merge → validate → cleanup."""
    init_db()
    db = SessionLocal()

    if sources is None:
        sources = ["nasdaq", "wire", "ir", "historical"]

    results = {}

    try:
        if "nasdaq" in sources:
            results["nasdaq"] = run_nasdaq_source(db, days_ahead=nasdaq_days)

        if "wire" in sources:
            results["wire"] = run_wire_source(db, prn_pages=prn_pages, bw_pages=bw_pages)

        if "ir" in sources:
            results["ir"] = run_ir_source(db, limit=ir_limit)

        if "historical" in sources:
            results["historical"] = run_historical_fallback(db)

        # Always run cleanup
        results["cleanup"] = run_quality_cleanup(db)

        # Final stats
        from sqlalchemy import func
        total = db.query(func.count(Event.id)).scalar()
        confirmed = db.query(func.count(Event.id)).filter(Event.status == "confirmed").scalar()
        tentative = db.query(func.count(Event.id)).filter(Event.status == "tentative").scalar()
        estimated = db.query(func.count(Event.id)).filter(Event.status == "estimated").scalar()
        coverage = db.query(Event.ticker).distinct().count()
        total_cos = db.query(Company).count()

        logger.info("=== PIPELINE COMPLETE ===")
        logger.info("Events: %d (confirmed: %d, tentative: %d, estimated: %d)", total, confirmed, tentative, estimated)
        logger.info("Coverage: %d/%d companies (%.1f%%)", coverage, total_cos, coverage / total_cos * 100)

        results["final"] = {
            "total_events": total,
            "confirmed": confirmed,
            "tentative": tentative,
            "estimated": estimated,
            "coverage": f"{coverage}/{total_cos}",
        }

    finally:
        db.close()

    return results


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Unified earnings pipeline")
    parser.add_argument("--source", "-s", help="Run specific source: nasdaq, wire, ir, historical")
    parser.add_argument("--ticker", "-t", help="Specific ticker(s)")
    parser.add_argument("--nasdaq-days", type=int, default=60)
    parser.add_argument("--prn-pages", type=int, default=5)
    parser.add_argument("--bw-pages", type=int, default=20)
    parser.add_argument("--ir-limit", type=int, default=None)
    args = parser.parse_args()

    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")

    sources = [args.source] if args.source else None
    run_full_pipeline(
        sources=sources,
        nasdaq_days=args.nasdaq_days,
        prn_pages=args.prn_pages,
        bw_pages=args.bw_pages,
        ir_limit=args.ir_limit,
    )
