"""Comprehensive data quality validation for ALL companies.

This is not a spot-check. This validates every single company and event
in the database against quality rules that a financial tool must meet.

Run: python -m pytest test_data_quality.py -v
Or:  python test_data_quality.py (for summary report)
"""
import pytest
import sys
from datetime import date, datetime, timedelta
from pathlib import Path
from collections import Counter

sys.path.insert(0, str(Path(__file__).parent.parent / "backend"))

from database import SessionLocal, init_db
from models import Event, Company
from sqlalchemy import func


@pytest.fixture(scope="module")
def db():
    init_db()
    session = SessionLocal()
    yield session
    session.close()


# ══════════════════════════════════════════════════════════════════
# RULE 1: Every company must have at least one event
# ══════════════════════════════════════════════════════════════════

class TestCompleteness:
    def test_every_company_has_event(self, db):
        """Every company in the database must have at least one event."""
        companies = {c.ticker for c in db.query(Company).all()}
        events = {r[0] for r in db.query(Event.ticker).distinct().all()}
        missing = companies - events
        assert len(missing) == 0, (
            f"{len(missing)} companies have no events: {sorted(missing)[:20]}"
        )

    def test_sp500_fully_covered(self, db):
        """Every S&P 500 company must have at least one event."""
        sp500 = {c.ticker for c in db.query(Company).filter(Company.market_cap_tier == "sp500").all()}
        events = {r[0] for r in db.query(Event.ticker).distinct().all()}
        missing = sp500 - events
        assert len(missing) == 0, (
            f"{len(missing)} S&P 500 companies have no events: {sorted(missing)}"
        )

    def test_minimum_company_count(self, db):
        """Must have at least 2500 companies (Russell 3000 target)."""
        count = db.query(Company).count()
        assert count >= 2500, f"Only {count} companies"

    def test_minimum_event_count(self, db):
        """Must have at least 2000 events during earnings season."""
        count = db.query(Event).count()
        assert count >= 2000, f"Only {count} events"


# ══════════════════════════════════════════════════════════════════
# RULE 2: No duplicate events
# ══════════════════════════════════════════════════════════════════

class TestNoDuplicates:
    def test_no_duplicate_earnings_same_date(self, db):
        """No company should have two earnings events on the same date."""
        dupes = (
            db.query(Event.ticker, Event.event_date, func.count(Event.id))
            .filter(Event.event_type == "earnings")
            .group_by(Event.ticker, Event.event_date)
            .having(func.count(Event.id) > 1)
            .all()
        )
        assert len(dupes) == 0, (
            f"{len(dupes)} duplicate earnings events: "
            f"{[(t, str(d), c) for t, d, c in dupes[:10]]}"
        )

    def test_no_multiple_earnings_same_quarter(self, db):
        """No company should have two earnings events within 30 days of each other."""
        today = date.today()
        tickers = db.query(Event.ticker).distinct().all()
        problems = []
        for (ticker,) in tickers:
            events = (
                db.query(Event)
                .filter(
                    Event.ticker == ticker,
                    Event.event_type == "earnings",
                    Event.event_date >= today,
                )
                .order_by(Event.event_date)
                .all()
            )
            for i in range(len(events) - 1):
                gap = (events[i + 1].event_date - events[i].event_date).days
                if gap < 30:
                    problems.append(
                        f"{ticker}: {events[i].event_date} and {events[i+1].event_date} ({gap}d apart)"
                    )
        assert len(problems) == 0, (
            f"{len(problems)} companies have earnings too close together:\n"
            + "\n".join(problems[:20])
        )


# ══════════════════════════════════════════════════════════════════
# RULE 3: No stale or past events
# ══════════════════════════════════════════════════════════════════

class TestNoStaleData:
    def test_no_past_events(self, db):
        """All events should be in the future (or today)."""
        yesterday = date.today() - timedelta(days=1)
        past = db.query(Event).filter(Event.event_date < yesterday).count()
        assert past == 0, f"{past} events are in the past"

    def test_events_within_reasonable_range(self, db):
        """No events should be more than 6 months in the future."""
        max_date = date.today() + timedelta(days=180)
        far_future = db.query(Event).filter(Event.event_date > max_date).count()
        assert far_future == 0, (
            f"{far_future} events are more than 6 months away"
        )


# ══════════════════════════════════════════════════════════════════
# RULE 4: Confirmed events must have complete data
# ══════════════════════════════════════════════════════════════════

class TestConfirmedQuality:
    def test_confirmed_have_time(self, db):
        """Every confirmed event must have a time."""
        bad = (
            db.query(Event)
            .filter(Event.status == "confirmed", Event.event_time.is_(None))
            .all()
        )
        tickers = [e.ticker for e in bad]
        assert len(bad) == 0, (
            f"{len(bad)} confirmed events have no time: {tickers[:20]}"
        )

    def test_confirmed_have_webcast(self, db):
        """Every confirmed event must have a webcast URL."""
        bad = (
            db.query(Event)
            .filter(Event.status == "confirmed", Event.webcast_url.is_(None))
            .all()
        )
        tickers = [e.ticker for e in bad]
        assert len(bad) == 0, (
            f"{len(bad)} confirmed events have no webcast URL: {tickers[:20]}"
        )

    def test_confirmed_webcast_urls_are_valid(self, db):
        """Webcast URLs should look like actual URLs, not PR page URLs."""
        bad = []
        confirmed = db.query(Event).filter(
            Event.status == "confirmed",
            Event.webcast_url.isnot(None),
        ).all()
        for e in confirmed:
            url = e.webcast_url
            # PR page URLs are not webcasts
            if "prnewswire.com" in url or "businesswire.com" in url:
                bad.append(f"{e.ticker}: {url[:80]}")
        assert len(bad) == 0, (
            f"{len(bad)} confirmed events have PR page URLs as webcast (not actual webcast):\n"
            + "\n".join(bad[:10])
        )

    def test_confirmed_times_are_reasonable(self, db):
        """Conference call times should be between 6 AM and 10 PM ET."""
        from datetime import time as t
        bad = []
        confirmed = db.query(Event).filter(
            Event.status == "confirmed",
            Event.event_time.isnot(None),
        ).all()
        for e in confirmed:
            hour = e.event_time.hour
            if hour < 6 or hour > 22:
                bad.append(f"{e.ticker}: {e.event_time}")
        assert len(bad) == 0, (
            f"{len(bad)} confirmed events have unreasonable times: {bad[:10]}"
        )


# ══════════════════════════════════════════════════════════════════
# RULE 5: Data consistency
# ══════════════════════════════════════════════════════════════════

class TestDataConsistency:
    def test_all_tickers_in_companies_table(self, db):
        """Every event ticker must exist in the companies table."""
        event_tickers = {r[0] for r in db.query(Event.ticker).distinct().all()}
        company_tickers = {r[0] for r in db.query(Company.ticker).all()}
        orphans = event_tickers - company_tickers
        assert len(orphans) == 0, (
            f"{len(orphans)} event tickers not in companies table: {sorted(orphans)[:20]}"
        )

    def test_event_types_are_valid(self, db):
        """All event types must be from the allowed set."""
        valid_types = {"earnings", "investor_day", "conference", "ad_hoc"}
        types = {r[0] for r in db.query(Event.event_type).distinct().all()}
        invalid = types - valid_types
        assert len(invalid) == 0, f"Invalid event types: {invalid}"

    def test_statuses_are_valid(self, db):
        """All statuses must be from the allowed set."""
        valid = {"confirmed", "tentative", "estimated", "pending", "postponed", "cancelled"}
        statuses = {r[0] for r in db.query(Event.status).distinct().all()}
        invalid = statuses - valid
        assert len(invalid) == 0, f"Invalid statuses: {invalid}"

    def test_sources_are_valid(self, db):
        """All sources must be recognized."""
        valid = {
            "nasdaq", "prnewswire", "businesswire", "globenewswire",
            "ir_page", "historical_pattern", "estimated", "manual",
        }
        sources = {r[0] for r in db.query(Event.source).distinct().all()}
        invalid = sources - valid
        assert len(invalid) == 0, f"Invalid sources: {invalid}"

    def test_dates_are_weekdays(self, db):
        """Earnings events should be on weekdays (Mon-Fri)."""
        bad = []
        events = db.query(Event).filter(Event.event_type == "earnings").all()
        for e in events:
            if e.event_date.weekday() >= 5:  # Saturday=5, Sunday=6
                bad.append(f"{e.ticker}: {e.event_date} ({e.event_date.strftime('%A')})")
        assert len(bad) == 0, (
            f"{len(bad)} earnings events on weekends: {bad[:10]}"
        )


# ══════════════════════════════════════════════════════════════════
# RULE 6: Wire service coverage quality
# ══════════════════════════════════════════════════════════════════

class TestWireCoverage:
    def test_sp500_wire_coverage_above_threshold(self, db):
        """S&P 500 wire coverage should grow as season progresses.

        Thresholds (tighten these weekly during earnings season):
        - Early season (now): >= 20%
        - Mid season: >= 50%
        - Late season: >= 70%
        """
        sp500 = {c.ticker for c in db.query(Company).filter(Company.market_cap_tier == "sp500").all()}
        wire_sources = ("prnewswire", "businesswire")
        wire_sp500 = {
            r[0] for r in
            db.query(Event.ticker)
            .filter(Event.ticker.in_(sp500), Event.source.in_(wire_sources))
            .distinct()
            .all()
        }
        pct = len(wire_sp500) / len(sp500) * 100
        assert pct >= 20, (
            f"Only {pct:.0f}% of S&P 500 has wire data ({len(wire_sp500)}/{len(sp500)}). "
            f"Expected at least 20%. Run deeper scrape."
        )

    def test_confirmed_percentage_reasonable(self, db):
        """Confirmed rate should grow as wire PRs are processed.

        Thresholds (tighten weekly):
        - Early season (now): >= 3%
        - Mid season: >= 15%
        - Late season: >= 30%
        """
        total = db.query(Event).count()
        confirmed = db.query(Event).filter(Event.status == "confirmed").count()
        pct = confirmed / total * 100
        assert pct >= 3, (
            f"Only {pct:.1f}% confirmed ({confirmed}/{total}). "
            f"Expected at least 3%. Check extraction quality."
        )


# ══════════════════════════════════════════════════════════════════
# RULE 7: Cross-source date validation
# ══════════════════════════════════════════════════════════════════

class TestCrossValidation:
    def test_no_conflicting_dates_for_same_company(self, db):
        """If a company has multiple events of the same type, dates should not conflict."""
        today = date.today()
        problems = []
        tickers = [r[0] for r in db.query(Event.ticker).distinct().all()]
        for ticker in tickers:
            events = (
                db.query(Event)
                .filter(
                    Event.ticker == ticker,
                    Event.event_type == "earnings",
                    Event.event_date >= today,
                )
                .all()
            )
            if len(events) > 1:
                dates = [e.event_date for e in events]
                # Check if any two dates are very close (likely duplicates)
                for i in range(len(dates)):
                    for j in range(i + 1, len(dates)):
                        gap = abs((dates[i] - dates[j]).days)
                        if gap < 14:
                            problems.append(f"{ticker}: {dates[i]} vs {dates[j]}")
        assert len(problems) == 0, (
            f"{len(problems)} companies have conflicting dates:\n"
            + "\n".join(problems[:20])
        )


# ── Summary report ───────────────────────────────────────────────

if __name__ == "__main__":
    init_db()
    db = SessionLocal()

    total_cos = db.query(Company).count()
    sp500_cos = db.query(Company).filter(Company.market_cap_tier == "sp500").count()
    total_events = db.query(Event).count()
    confirmed = db.query(Event).filter(Event.status == "confirmed").count()
    tentative = db.query(Event).filter(Event.status == "tentative").count()
    with_events = db.query(Event.ticker).distinct().count()

    today = date.today()
    past_events = db.query(Event).filter(Event.event_date < today).count()
    weekend_events = len([
        e for e in db.query(Event).filter(Event.event_type == "earnings").all()
        if e.event_date.weekday() >= 5
    ])

    # Duplicates
    dupes = (
        db.query(Event.ticker, Event.event_date, func.count(Event.id))
        .filter(Event.event_type == "earnings")
        .group_by(Event.ticker, Event.event_date)
        .having(func.count(Event.id) > 1)
        .all()
    )

    # Confirmed with bad webcast URLs
    bad_webcasts = db.query(Event).filter(
        Event.status == "confirmed",
        Event.webcast_url.isnot(None),
    ).all()
    pr_url_webcasts = [e for e in bad_webcasts if e.webcast_url and ("prnewswire.com" in e.webcast_url or "businesswire.com" in e.webcast_url)]

    # Source breakdown
    sources = db.query(Event.source, func.count(Event.id)).group_by(Event.source).order_by(func.count(Event.id).desc()).all()

    # Wire coverage of S&P 500
    sp500_tickers = {c.ticker for c in db.query(Company).filter(Company.market_cap_tier == "sp500").all()}
    wire_sp500 = db.query(Event.ticker).filter(
        Event.ticker.in_(sp500_tickers),
        Event.source.in_(("prnewswire", "businesswire")),
    ).distinct().count()

    print("=" * 60)
    print("COMPREHENSIVE DATA QUALITY REPORT")
    print("=" * 60)
    print(f"\nCOMPANIES: {total_cos} ({sp500_cos} S&P 500)")
    print(f"EVENTS:    {total_events} ({confirmed} confirmed, {tentative} tentative)")
    print(f"COVERAGE:  {with_events}/{total_cos} companies ({with_events/total_cos*100:.1f}%)")
    print(f"\nDATA ISSUES:")
    print(f"  Past events:          {past_events} {'FAIL' if past_events > 0 else 'OK'}")
    print(f"  Weekend events:       {weekend_events} {'FAIL' if weekend_events > 0 else 'OK'}")
    print(f"  Duplicate events:     {len(dupes)} {'FAIL' if dupes else 'OK'}")
    print(f"  PR URL as webcast:    {len(pr_url_webcasts)} {'FAIL' if pr_url_webcasts else 'OK'}")
    print(f"\nSOURCES:")
    for src, cnt in sources:
        print(f"  {src:25s} {cnt:5d}")
    print(f"\nS&P 500 WIRE COVERAGE: {wire_sp500}/{len(sp500_tickers)} ({wire_sp500/len(sp500_tickers)*100:.0f}%)")
    print(f"CONFIRMED RATE: {confirmed}/{total_events} ({confirmed/total_events*100:.1f}%)")

    db.close()
