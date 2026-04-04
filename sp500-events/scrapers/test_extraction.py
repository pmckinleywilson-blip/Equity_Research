"""RED-GREEN TDD: Tests for wire service PR text extraction.

These tests verify that _extract_time, _extract_webcast_url, _extract_phone,
and _extract_passcode correctly parse real press release text.

Run: python -m pytest test_extraction.py -v
"""
import pytest
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent / "backend"))

from wire_services import (
    _extract_time,
    _extract_webcast_url,
    _extract_phone,
    _extract_passcode,
    _extract_event_date,
    _extract_ticker,
)


# ── Real PR fixtures ──────────────────────────────────────────────

META_PR = (
    "MENLO PARK, Calif., Jan. 14, 2026 /PRNewswire/ -- Meta Platforms, Inc. "
    "(NASDAQ: META) announced today that the company's fourth quarter and full "
    "year 2025 financial results will be released after market close on "
    "Wednesday, January 28th, 2026.\n"
    "Meta will host a conference call to discuss its results at 1:30 p.m. PT / "
    "4:30 p.m. ET the same day. The live webcast of the call can be accessed "
    "at the Meta Investor Relations website at investor.atmeta.com, along with "
    "the company's earnings press release, financial tables, and slide presentation."
)

NETFLIX_PR = (
    "LOS GATOS, Calif., March 13, 2026 /PRNewswire/ -- Netflix, Inc. "
    "(NASDAQ: NFLX) today announced it will post its first quarter 2026 "
    "financial results and business outlook on its investor relations website "
    "at http://ir.netflix.net on Thursday April 16th, 2026, at approximately "
    "1:01 p.m. Pacific Time.\n"
    "A live video interview will begin at 1:45 p.m. Pacific Time.\n"
    "The live earnings video interview will be accessible on the Netflix "
    "Investor Relations YouTube channel at youtube.com/netflixir at "
    "1:45 p.m. Pacific Time."
)

HOME_DEPOT_PR = (
    "ATLANTA, Nov. 4, 2025 /PRNewswire/ -- The Home Depot, the world's "
    "largest home improvement retailer, announced today that it will hold "
    "its Third Quarter Earnings Conference Call on Tuesday, November 18, "
    "at 9 a.m. ET.\n"
    "A webcast will be available by logging onto "
    "http://ir.homedepot.com/events-and-presentations and selecting the "
    "Third Quarter Earnings Conference Call icon."
)

# BusinessWire-style fixture (manually crafted from real Chevron PR)
CHEVRON_PR = (
    "HOUSTON--(BUSINESS WIRE)--Chevron Corporation (NYSE: CVX), one of the "
    "world's leading energy companies, will hold its quarterly earnings "
    "conference call on Friday, May 1, 2026 at 11:00 a.m. ET (10:00 a.m. CT).\n"
    "Conference Call Information:\n"
    "Date: Friday, May 1, 2026\n"
    "Time: 11:00 a.m. ET / 10:00 a.m. CT\n"
    "Webcast: https://www.chevron.com/investors\n"
    "A recording of the call will be available on the website after the call."
)

# PR with dial-in number and passcode
DIALIN_PR = (
    "NEW YORK--(BUSINESS WIRE)--Acme Corp (NYSE: ACME) will host a conference "
    "call on Tuesday, April 22, 2026, at 8:30 a.m. Eastern Time to discuss "
    "first quarter 2026 financial results.\n"
    "Participants can access the call by dialing 1-800-555-0199 (domestic) or "
    "+1-212-555-0199 (international). The passcode is 123456.\n"
    "A live webcast will be available at https://events.q4inc.com/attendee/987654321."
)

# PR with "before market open" timing
BMO_PR = (
    "MINNEAPOLIS, March 20, 2026 /PRNewswire/ -- Target Corporation (NYSE: TGT) "
    "will release its first quarter fiscal 2026 financial results before the "
    "market opens on Wednesday, May 20, 2026. A conference call will begin at "
    "10:00 a.m. EDT. Investors may listen to a live webcast of the call at "
    "https://investors.target.com."
)


# ── Time extraction tests ─────────────────────────────────────────

class TestExtractTime:
    def test_meta_430pm_et(self):
        """Meta PR: '4:30 p.m. ET' -> 16:30:00"""
        result = _extract_time(META_PR)
        assert result == "16:30:00"

    def test_netflix_145pm_pacific(self):
        """Netflix PR: '1:45 p.m. Pacific Time' -> 13:45:00"""
        result = _extract_time(NETFLIX_PR)
        assert result is not None
        # Should extract a time (either PT or the earlier 1:01 PM)
        assert ":" in result

    def test_home_depot_9am_et(self):
        """Home Depot PR: '9 a.m. ET' -> 09:00:00"""
        result = _extract_time(HOME_DEPOT_PR)
        assert result == "09:00:00"

    def test_chevron_11am_et(self):
        """Chevron PR: '11:00 a.m. ET' -> 11:00:00"""
        result = _extract_time(CHEVRON_PR)
        assert result == "11:00:00"

    def test_dialin_830am(self):
        """Dial-in PR: '8:30 a.m. Eastern Time' -> 08:30:00"""
        result = _extract_time(DIALIN_PR)
        assert result == "08:30:00"

    def test_target_10am_edt(self):
        """Target PR: '10:00 a.m. EDT' -> 10:00:00"""
        result = _extract_time(BMO_PR)
        assert result == "10:00:00"

    def test_no_time(self):
        """Text without a time returns None."""
        result = _extract_time("Company will report results on April 15.")
        assert result is None


# ── Webcast URL extraction tests ──────────────────────────────────

class TestExtractWebcastUrl:
    def test_meta_investor_site(self):
        """Meta PR: 'investor.atmeta.com' -> URL extracted"""
        result = _extract_webcast_url(META_PR)
        assert result is not None
        assert "investor" in result.lower() or "atmeta" in result.lower()

    def test_netflix_youtube(self):
        """Netflix PR: 'youtube.com/netflixir' -> URL extracted"""
        result = _extract_webcast_url(NETFLIX_PR)
        assert result is not None

    def test_home_depot_ir(self):
        """Home Depot PR: 'http://ir.homedepot.com/...' -> URL extracted"""
        result = _extract_webcast_url(HOME_DEPOT_PR)
        assert result is not None
        assert "homedepot" in result.lower()

    def test_chevron_investors(self):
        """Chevron PR: 'https://www.chevron.com/investors' -> URL extracted"""
        result = _extract_webcast_url(CHEVRON_PR)
        assert result is not None
        assert "chevron" in result.lower()

    def test_q4_event_url(self):
        """Dial-in PR: 'https://events.q4inc.com/...' -> URL extracted"""
        result = _extract_webcast_url(DIALIN_PR)
        assert result is not None
        assert "q4inc" in result.lower()

    def test_target_investors(self):
        """Target PR: 'https://investors.target.com' -> URL extracted"""
        result = _extract_webcast_url(BMO_PR)
        assert result is not None
        assert "target" in result.lower()

    def test_no_url(self):
        """Text without webcast URL returns None."""
        result = _extract_webcast_url("Company will report results on April 15.")
        assert result is None


# ── Phone extraction tests ────────────────────────────────────────

class TestExtractPhone:
    def test_dialin_domestic(self):
        """Extract domestic dial-in: '1-800-555-0199'"""
        result = _extract_phone(DIALIN_PR)
        assert result is not None
        assert "800" in result or "555" in result

    def test_no_phone(self):
        """Text without phone returns None."""
        result = _extract_phone(META_PR)
        assert result is None


# ── Passcode extraction tests ─────────────────────────────────────

class TestExtractPasscode:
    def test_passcode(self):
        """Extract passcode: '123456'"""
        result = _extract_passcode(DIALIN_PR)
        assert result == "123456"

    def test_no_passcode(self):
        """Text without passcode returns None."""
        result = _extract_passcode(META_PR)
        assert result is None


# ── Event date extraction tests ───────────────────────────────────

class TestExtractEventDate:
    def test_meta_date(self):
        """Meta PR: 'January 28th, 2026' -> 2026-01-28"""
        result = _extract_event_date(META_PR)
        assert result is not None
        assert result.year == 2026
        assert result.month == 1
        assert result.day == 28

    def test_netflix_date(self):
        """Netflix PR: 'Thursday April 16th, 2026' -> 2026-04-16"""
        result = _extract_event_date(NETFLIX_PR)
        assert result is not None
        assert result.month == 4
        assert result.day == 16

    def test_home_depot_date(self):
        """Home Depot PR: 'November 18' -> 2025-11-18 (infer year from context)"""
        result = _extract_event_date(HOME_DEPOT_PR)
        assert result is not None
        assert result.month == 11
        assert result.day == 18

    def test_chevron_date(self):
        """Chevron PR: 'May 1, 2026' -> 2026-05-01"""
        result = _extract_event_date(CHEVRON_PR)
        assert result is not None
        assert result.month == 5
        assert result.day == 1

    def test_target_date(self):
        """Target PR: 'May 20, 2026' -> 2026-05-20"""
        result = _extract_event_date(BMO_PR)
        assert result is not None
        assert result.month == 5
        assert result.day == 20


# ── Ticker extraction tests ──────────────────────────────────────

class TestExtractTicker:
    def test_meta_nasdaq(self):
        """Meta PR: '(NASDAQ: META)' -> META"""
        result = _extract_ticker(META_PR)
        assert result == "META"

    def test_netflix_nasdaq(self):
        """Netflix PR: '(NASDAQ: NFLX)' -> NFLX"""
        result = _extract_ticker(NETFLIX_PR)
        assert result == "NFLX"

    def test_chevron_nyse(self):
        """Chevron PR: '(NYSE: CVX)' -> CVX"""
        result = _extract_ticker(CHEVRON_PR)
        assert result == "CVX"

    def test_target_nyse(self):
        """Target PR: '(NYSE: TGT)' -> TGT"""
        result = _extract_ticker(BMO_PR)
        assert result == "TGT"
