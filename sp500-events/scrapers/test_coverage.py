"""Coverage and accuracy tests.

These tests verify that our database matches real-world announced events.
They act as a canary — when they fail, it means our scraping pipeline
has gaps that need fixing.

Run: python -m pytest test_coverage.py -v

To update the ground truth data, run: python test_coverage.py --update
"""
import pytest
import sys
from datetime import date, time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent / "backend"))

from database import SessionLocal, init_db
from models import Event, Company


# ══════════════════════════════════════════════════════════════════
# GROUND TRUTH: Manually verified events from company announcements.
# Add to this list whenever a new gap is discovered.
# Source: Wire service PRs, company IR pages, verified manually.
# ══════════════════════════════════════════════════════════════════

GROUND_TRUTH_EVENTS = [
    # === EARNINGS CALLS ===
    {
        "ticker": "TSLA",
        "event_type": "earnings",
        "event_date": date(2026, 4, 22),
        "event_time": time(17, 30),  # 5:30 PM ET
        "has_webcast": True,
        "source_note": "Tesla IR page: ir.tesla.com",
    },
    {
        "ticker": "META",
        "event_type": "earnings",
        "event_date": date(2026, 4, 30),  # Q1 2026 (not Q4 2025 from Jan)
        "event_time": None,  # Not yet announced for Q1
        "has_webcast": None,
        "source_note": "PRNewswire - Meta advance PR pattern",
    },
    {
        "ticker": "NFLX",
        "event_type": "earnings",
        "event_date": date(2026, 4, 16),
        "event_time": time(13, 45),  # 1:45 PM PT = need to verify ET
        "has_webcast": True,
        "source_note": "PRNewswire: netflix-to-announce-first-quarter-2026",
    },
    {
        "ticker": "JPM",
        "event_type": "earnings",
        "event_date": date(2026, 4, 14),
        "event_time": None,
        "has_webcast": None,
        "source_note": "BusinessWire: JPMorgan Q1 2026",
    },
    {
        "ticker": "BAC",
        "event_type": "earnings",
        "event_date": date(2026, 4, 15),
        "event_time": None,
        "has_webcast": None,
        "source_note": "Nasdaq calendar",
    },
    {
        "ticker": "HD",
        "event_type": "earnings",
        "event_date": date(2026, 5, 19),
        "event_time": None,
        "has_webcast": None,
        "source_note": "PRNewswire: Home Depot conference call pattern",
    },
    {
        "ticker": "PG",
        "event_type": "earnings",
        "event_date": date(2026, 4, 24),
        "event_time": time(8, 30),  # 8:30 AM ET
        "has_webcast": True,
        "source_note": "BusinessWire: PG-to-Webcast-Discussion",
    },
    {
        "ticker": "UNH",
        "event_type": "earnings",
        "event_date": date(2026, 4, 21),
        "event_time": None,
        "has_webcast": None,
        "source_note": "BusinessWire: UnitedHealth earnings release date",
    },
    {
        "ticker": "DAL",
        "event_type": "earnings",
        "event_date": date(2026, 4, 8),
        "event_time": None,
        "has_webcast": None,
        "source_note": "Nasdaq calendar",
    },

    # === AD HOC / INVESTOR EVENTS ===
    {
        "ticker": "AFRM",
        "event_type": "investor_day",
        "event_date": date(2026, 5, 12),
        "event_time": time(14, 0),  # 2:00 PM ET
        "has_webcast": True,
        "source_note": "BusinessWire: Affirm Investor Forum May 12",
    },
]

# Companies that MUST exist in our database (Russell 3000 / S&P 500)
MUST_HAVE_COMPANIES = [
    "AAPL", "MSFT", "GOOGL", "AMZN", "NVDA", "META", "TSLA", "BRK-B",
    "JPM", "V", "MA", "UNH", "JNJ", "HD", "PG", "BAC", "CRM", "NFLX",
    "COST", "ADBE", "AMD", "INTC", "DIS", "NKE", "SBUX", "GS", "MS",
    "PFE", "KO", "PEP", "WMT", "XOM", "CVX", "LLY", "ABBV", "MRK",
    "AVGO", "QCOM", "AFRM", "PLTR", "COIN", "HOOD", "CRWD", "SNOW",
    "UBER", "DASH", "ABNB",
]


@pytest.fixture(scope="module")
def db():
    init_db()
    session = SessionLocal()
    yield session
    session.close()


# ── Test 1: Every company in our list exists in the database ──────

class TestCompanyCoverage:
    def test_must_have_companies_exist(self, db):
        """All key companies must be in our database."""
        missing = []
        for ticker in MUST_HAVE_COMPANIES:
            company = db.query(Company).filter(Company.ticker == ticker).first()
            if not company:
                missing.append(ticker)
        assert missing == [], f"Companies missing from database: {missing}"

    def test_every_company_has_at_least_one_event(self, db):
        """Every company in the database should have at least one event."""
        total_companies = db.query(Company).count()
        companies_with_events = db.query(Event.ticker).distinct().count()
        gap = total_companies - companies_with_events
        assert gap == 0, (
            f"{gap} companies have no events. "
            f"Coverage: {companies_with_events}/{total_companies}"
        )

    def test_sp500_count(self, db):
        """S&P 500 should have ~500 companies."""
        count = db.query(Company).filter(Company.market_cap_tier == "sp500").count()
        assert count >= 490, f"S&P 500 has only {count} companies (expected ~500)"

    def test_total_company_count(self, db):
        """Russell 3000 + S&P 500 should have ~2500+ companies."""
        count = db.query(Company).count()
        assert count >= 2500, f"Only {count} companies (expected 2500+)"


# ── Test 2: Known events exist with correct dates ────────────────

class TestEventAccuracy:
    @pytest.mark.parametrize("event", GROUND_TRUTH_EVENTS, ids=[e["ticker"] for e in GROUND_TRUTH_EVENTS])
    def test_event_date_exists(self, db, event):
        """Each ground truth event should exist in our database with the correct date."""
        db_event = (
            db.query(Event)
            .filter(
                Event.ticker == event["ticker"],
                Event.event_date == event["event_date"],
            )
            .first()
        )
        assert db_event is not None, (
            f"{event['ticker']}: No event found for {event['event_date']}. "
            f"Source: {event['source_note']}"
        )

    @pytest.mark.parametrize("event", [e for e in GROUND_TRUTH_EVENTS if e.get("event_time")],
                             ids=[e["ticker"] for e in GROUND_TRUTH_EVENTS if e.get("event_time")])
    def test_event_has_time(self, db, event):
        """Events with known times should have the time populated."""
        db_event = (
            db.query(Event)
            .filter(
                Event.ticker == event["ticker"],
                Event.event_date == event["event_date"],
            )
            .first()
        )
        if db_event is None:
            pytest.skip(f"{event['ticker']}: event not in DB yet")
        assert db_event.event_time is not None, (
            f"{event['ticker']}: Time should be {event['event_time']} but is NULL. "
            f"Source: {event['source_note']}"
        )

    @pytest.mark.parametrize("event", [e for e in GROUND_TRUTH_EVENTS if e.get("has_webcast")],
                             ids=[e["ticker"] for e in GROUND_TRUTH_EVENTS if e.get("has_webcast")])
    def test_event_has_webcast(self, db, event):
        """Events with known webcasts should have the URL populated."""
        db_event = (
            db.query(Event)
            .filter(
                Event.ticker == event["ticker"],
                Event.event_date == event["event_date"],
            )
            .first()
        )
        if db_event is None:
            pytest.skip(f"{event['ticker']}: event not in DB yet")
        assert db_event.webcast_url is not None, (
            f"{event['ticker']}: Should have webcast URL but is NULL. "
            f"Source: {event['source_note']}"
        )


# ── Test 3: Ad-hoc events (investor days, forums, etc.) ─────────

class TestAdHocEvents:
    def test_affirm_investor_forum(self, db):
        """Affirm Investor Forum on May 12 should be captured."""
        event = (
            db.query(Event)
            .filter(
                Event.ticker == "AFRM",
                Event.event_date == date(2026, 5, 12),
            )
            .first()
        )
        assert event is not None, (
            "AFRM Investor Forum (May 12, 2026) not found. "
            "BusinessWire published this PR. Check category filtering."
        )

    def test_ad_hoc_events_not_all_earnings(self, db):
        """Database should contain some non-earnings events."""
        non_earnings = (
            db.query(Event)
            .filter(Event.event_type != "earnings")
            .count()
        )
        # This will fail initially - that's the point
        assert non_earnings > 0, (
            "No ad-hoc events found. The scraper is only capturing earnings. "
            "Need to monitor broader wire service categories."
        )


# ── Test 4: Confirmed events have complete data ─────────────────

class TestConfirmedQuality:
    def test_confirmed_events_have_time(self, db):
        """Every confirmed event must have a time."""
        bad = (
            db.query(Event)
            .filter(
                Event.status == "confirmed",
                Event.event_time.is_(None),
            )
            .count()
        )
        assert bad == 0, f"{bad} events marked 'confirmed' but have no time"

    def test_confirmed_events_have_webcast(self, db):
        """Every confirmed event must have a webcast URL."""
        bad = (
            db.query(Event)
            .filter(
                Event.status == "confirmed",
                Event.webcast_url.is_(None),
            )
            .count()
        )
        assert bad == 0, f"{bad} events marked 'confirmed' but have no webcast URL"


# ── Test 5: Source diversity ─────────────────────────────────────

class TestSourceDiversity:
    def test_has_wire_service_events(self, db):
        """Should have events from both PRNewswire and BusinessWire."""
        prn = db.query(Event).filter(Event.source == "prnewswire").count()
        bw = db.query(Event).filter(Event.source == "businesswire").count()
        assert prn > 0, "No PRNewswire events found"
        assert bw > 0, "No BusinessWire events found"

    def test_has_nasdaq_baseline(self, db):
        """Should have Nasdaq calendar baseline events."""
        nasdaq = db.query(Event).filter(Event.source == "nasdaq").count()
        assert nasdaq > 500, f"Only {nasdaq} Nasdaq events (expected 500+)"

    def test_wire_services_cover_major_companies(self, db):
        """Major companies should have wire-sourced (not just Nasdaq) events."""
        major = ["JPM", "BAC", "GS", "MS", "PFE", "UNH"]
        wire_sources = ("prnewswire", "businesswire")
        missing_wire = []
        for ticker in major:
            wire_event = (
                db.query(Event)
                .filter(Event.ticker == ticker, Event.source.in_(wire_sources))
                .first()
            )
            if not wire_event:
                missing_wire.append(ticker)
        # Allow some failures since wire PRs come in over time
        assert len(missing_wire) <= 2, (
            f"Major companies without wire service data: {missing_wire}"
        )


# ── Run directly to see a summary ───────────────────────────────

if __name__ == "__main__":
    init_db()
    db = SessionLocal()

    print("=== COVERAGE REPORT ===\n")

    # Company coverage
    total_cos = db.query(Company).count()
    with_events = db.query(Event.ticker).distinct().count()
    print(f"Companies: {with_events}/{total_cos} have events ({with_events/total_cos*100:.0f}%)")

    # Ground truth check
    print(f"\nGround truth events ({len(GROUND_TRUTH_EVENTS)}):")
    for gt in GROUND_TRUTH_EVENTS:
        ev = db.query(Event).filter(
            Event.ticker == gt["ticker"],
            Event.event_date == gt["event_date"],
        ).first()

        if ev is None:
            status = "MISSING"
        elif gt.get("event_time") and not ev.event_time:
            status = "NO TIME"
        elif gt.get("has_webcast") and not ev.webcast_url:
            status = "NO WEBCAST"
        elif ev.status == "confirmed":
            status = "OK (confirmed)"
        else:
            status = "OK (tentative)"

        print(f"  {gt['ticker']:6s} {gt['event_date']} {status:20s} {gt['source_note']}")

    # Ad-hoc events
    non_earnings = db.query(Event).filter(Event.event_type != "earnings").count()
    print(f"\nAd-hoc (non-earnings) events: {non_earnings}")

    # Source breakdown
    from sqlalchemy import func
    sources = db.query(Event.source, func.count(Event.id)).group_by(Event.source).all()
    print(f"\nBy source:")
    for src, cnt in sources:
        print(f"  {src:25s} {cnt}")

    db.close()
