"""Scraping pipeline orchestrator.

Coordinates the detect → verify on IR → store → notify flow:
1. Run detection sources (Nasdaq, Finnhub, EDGAR, press releases)
2. For each detected event, verify on the company IR page
3. Store/update events in the database
4. Trigger notifications for newly confirmed events
"""
import argparse
import asyncio
import logging
import sys
from datetime import datetime, timedelta
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent / "backend"))

from database import SessionLocal, init_db
from models import Event, Company

logger = logging.getLogger(__name__)


def run_full_pipeline(
    days_ahead: int = 90,
    run_detection: bool = True,
    run_verification: bool = True,
    run_notifications: bool = True,
    tickers: list[str] = None,
):
    """Run the complete scraping pipeline."""
    init_db()
    results = {}

    # Step 1: Detection
    if run_detection:
        logger.info("=== Step 1: Detection ===")
        from earnings_calendars import run_detection as detect
        results["detection"] = detect(days_ahead=days_ahead)

        # Also run EDGAR detection
        try:
            from sec_edgar import run_edgar_detection
            results["edgar"] = run_edgar_detection(days_back=2)
        except Exception as e:
            logger.error("EDGAR detection failed: %s", e)
            results["edgar"] = {"error": str(e)}

    # Step 2: IR Page Verification
    if run_verification:
        logger.info("=== Step 2: IR Verification ===")
        from ir_pages import verify_events_on_ir

        db = SessionLocal()

        if tickers:
            # Verify specific tickers (e.g., triggered by new detection)
            results["ir_verification"] = verify_events_on_ir(tickers=tickers, db=db)
        else:
            # Verify by priority tier
            today = datetime.utcnow().date()
            day_of_week = today.weekday()

            # S&P 500: every day
            sp500 = [
                c.ticker for c in
                db.query(Company)
                .filter(Company.market_cap_tier == "sp500", Company.ir_events_url.isnot(None))
                .all()
            ]
            if sp500:
                results["ir_sp500"] = verify_events_on_ir(tickers=sp500, db=db)

            # Mid-caps: every 2 days (Mon, Wed, Fri)
            if day_of_week in (0, 2, 4):
                mid_caps = [
                    c.ticker for c in
                    db.query(Company)
                    .filter(Company.market_cap_tier == "mid", Company.ir_events_url.isnot(None))
                    .all()
                ]
                if mid_caps:
                    results["ir_mid"] = verify_events_on_ir(tickers=mid_caps, db=db)

            # Small-caps: every 3 days (Mon, Thu)
            if day_of_week in (0, 3):
                small_caps = [
                    c.ticker for c in
                    db.query(Company)
                    .filter(Company.market_cap_tier == "small", Company.ir_events_url.isnot(None))
                    .all()
                ]
                if small_caps:
                    results["ir_small"] = verify_events_on_ir(tickers=small_caps, db=db)

            # Priority check: any ticker that was just detected but not yet verified
            unverified = [
                r[0] for r in
                db.query(Event.ticker)
                .filter(
                    Event.ir_verified == False,
                    Event.created_at >= datetime.utcnow() - timedelta(hours=6),
                )
                .distinct()
                .all()
            ]
            if unverified:
                logger.info("Priority IR check for %d newly detected tickers", len(unverified))
                results["ir_priority"] = verify_events_on_ir(tickers=unverified, db=db)

        db.close()

    # Step 3: Notifications
    if run_notifications:
        logger.info("=== Step 3: Notifications ===")
        from services.notifier import notify_new_events
        db = SessionLocal()
        results["notifications"] = asyncio.run(notify_new_events(db))
        db.close()

    logger.info("=== Pipeline Complete ===")
    for key, value in results.items():
        logger.info("  %s: %s", key, value)

    return results


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="SP500 Events scraping pipeline")
    parser.add_argument("--ticker", "-t", help="Run for specific ticker(s), comma-separated")
    parser.add_argument("--days", "-d", type=int, default=90, help="Days ahead to scan")
    parser.add_argument("--detect-only", action="store_true", help="Only run detection")
    parser.add_argument("--verify-only", action="store_true", help="Only run IR verification")
    parser.add_argument("--notify-only", action="store_true", help="Only run notifications")
    args = parser.parse_args()

    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(name)s %(levelname)s %(message)s")

    tickers = [t.strip().upper() for t in args.ticker.split(",")] if args.ticker else None

    run_full_pipeline(
        days_ahead=args.days,
        run_detection=not args.verify_only and not args.notify_only,
        run_verification=not args.detect_only and not args.notify_only,
        run_notifications=not args.detect_only and not args.verify_only,
        tickers=tickers,
    )
