"""Cross-validation: merge data from multiple sources and set confidence.

Rules:
- IR page is the source of truth (ir_verified=TRUE → confirmed)
- Multiple detection sources increase confidence but don't confirm
- Flags conflicts for manual review
"""
import logging
import sys
from datetime import datetime, timedelta
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent / "backend"))

from database import SessionLocal, init_db
from models import Event, Company

logger = logging.getLogger(__name__)


def run_cross_validation(db=None):
    """Run cross-validation checks on all events."""
    close_session = False
    if db is None:
        init_db()
        db = SessionLocal()
        close_session = True

    try:
        results = {
            "companies_without_events": [],
            "stale_tentative": 0,
            "total_confirmed": 0,
            "total_tentative": 0,
        }

        today = datetime.utcnow().date()

        # Check: every company should have at least one upcoming event
        # (during earnings season, most will have announced dates)
        companies = db.query(Company).all()
        for company in companies:
            upcoming = (
                db.query(Event)
                .filter(Event.ticker == company.ticker, Event.event_date >= today)
                .count()
            )
            if upcoming == 0:
                results["companies_without_events"].append(company.ticker)

        # Count confirmed vs tentative
        results["total_confirmed"] = (
            db.query(Event)
            .filter(Event.ir_verified == True, Event.event_date >= today)
            .count()
        )
        results["total_tentative"] = (
            db.query(Event)
            .filter(Event.ir_verified == False, Event.event_date >= today)
            .count()
        )

        # Flag stale tentative events (detected >7 days ago, still not verified)
        stale_cutoff = datetime.utcnow() - timedelta(days=7)
        stale = (
            db.query(Event)
            .filter(
                Event.ir_verified == False,
                Event.created_at < stale_cutoff,
                Event.event_date >= today,
            )
            .all()
        )
        results["stale_tentative"] = len(stale)

        # Log summary
        total_companies = len(companies)
        gap_count = len(results["companies_without_events"])
        logger.info(
            "Cross-validation: %d/%d companies have events. "
            "%d confirmed, %d tentative, %d stale tentative.",
            total_companies - gap_count,
            total_companies,
            results["total_confirmed"],
            results["total_tentative"],
            results["stale_tentative"],
        )

        if gap_count > 0:
            logger.warning(
                "Companies without upcoming events (%d): %s",
                gap_count,
                results["companies_without_events"][:20],
            )

        return results

    finally:
        if close_session:
            db.close()


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    results = run_cross_validation()
    print(f"Validation results: {results}")
