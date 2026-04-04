"""Tests for quality gates and event merge logic.

These test the validation rules in isolation, without touching the database.
"""
import pytest
import sys
from datetime import date, time, timedelta
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent / "backend"))

from quality_gates import (
    validate_event, determine_status, should_update, merge_event_data,
    clean_webcast_url, clean_time, fix_weekend_date,
)


class TestValidateEvent:
    def test_valid_event_passes(self):
        errors = validate_event({
            "ticker": "AAPL",
            "event_type": "earnings",
            "event_date": date.today() + timedelta(days=10),
            "source": "nasdaq",
        })
        assert errors == []

    def test_missing_ticker_rejected(self):
        errors = validate_event({
            "event_type": "earnings",
            "event_date": date.today() + timedelta(days=10),
            "source": "nasdaq",
        })
        assert any("ticker" in e for e in errors)

    def test_missing_date_rejected(self):
        errors = validate_event({
            "ticker": "AAPL",
            "event_type": "earnings",
            "source": "nasdaq",
        })
        assert any("event_date" in e for e in errors)

    def test_past_date_rejected(self):
        errors = validate_event({
            "ticker": "AAPL",
            "event_type": "earnings",
            "event_date": date.today() - timedelta(days=1),
            "source": "nasdaq",
        })
        assert any("past" in e.lower() for e in errors)

    def test_far_future_rejected(self):
        errors = validate_event({
            "ticker": "AAPL",
            "event_type": "earnings",
            "event_date": date.today() + timedelta(days=200),
            "source": "nasdaq",
        })
        assert any("6 months" in e for e in errors)

    def test_weekend_date_rejected(self):
        # Find the next Saturday
        d = date.today()
        while d.weekday() != 5:
            d += timedelta(days=1)
        errors = validate_event({
            "ticker": "AAPL",
            "event_type": "earnings",
            "event_date": d,
            "source": "nasdaq",
        })
        assert any("weekend" in e.lower() for e in errors)

    def test_invalid_event_type_rejected(self):
        errors = validate_event({
            "ticker": "AAPL",
            "event_type": "webinar",
            "event_date": date.today() + timedelta(days=10),
            "source": "nasdaq",
        })
        assert any("event type" in e.lower() for e in errors)

    def test_invalid_source_rejected(self):
        errors = validate_event({
            "ticker": "AAPL",
            "event_type": "earnings",
            "event_date": date.today() + timedelta(days=10),
            "source": "yahoo_finance",
        })
        assert any("source" in e.lower() for e in errors)

    def test_pr_page_webcast_rejected(self):
        errors = validate_event({
            "ticker": "AAPL",
            "event_type": "earnings",
            "event_date": date.today() + timedelta(days=10),
            "source": "businesswire",
            "webcast_url": "https://www.businesswire.com/news/home/123/en/Apple",
        })
        assert any("PR page" in e for e in errors)

    def test_unreasonable_time_rejected(self):
        errors = validate_event({
            "ticker": "AAPL",
            "event_type": "earnings",
            "event_date": date.today() + timedelta(days=10),
            "source": "nasdaq",
            "event_time": time(3, 0),
        })
        assert any("time" in e.lower() for e in errors)

    def test_investor_day_on_weekend_allowed(self):
        """Investor days CAN be on weekends (only earnings are weekday-restricted)."""
        d = date.today()
        while d.weekday() != 5:
            d += timedelta(days=1)
        errors = validate_event({
            "ticker": "AAPL",
            "event_type": "investor_day",
            "event_date": d,
            "source": "businesswire",
        })
        assert not any("weekend" in e.lower() for e in errors)


class TestDetermineStatus:
    def test_confirmed_needs_time_and_webcast(self):
        assert determine_status({
            "event_time": "17:00:00",
            "webcast_url": "https://ir.apple.com",
            "source": "businesswire",
        }) == "confirmed"

    def test_missing_time_is_tentative(self):
        assert determine_status({
            "webcast_url": "https://ir.apple.com",
            "source": "businesswire",
        }) == "tentative"

    def test_missing_webcast_is_tentative(self):
        assert determine_status({
            "event_time": "17:00:00",
            "source": "nasdaq",
        }) == "tentative"

    def test_estimated_source_is_estimated(self):
        assert determine_status({
            "event_time": "17:00:00",
            "webcast_url": "https://ir.apple.com",
            "source": "estimated",
        }) == "estimated"

    def test_historical_pattern_is_estimated(self):
        assert determine_status({
            "source": "historical_pattern",
        }) == "estimated"


class TestShouldUpdate:
    def test_wire_beats_nasdaq(self):
        assert should_update("nasdaq", "businesswire") is True

    def test_wire_beats_estimated(self):
        assert should_update("estimated", "prnewswire") is True

    def test_nasdaq_does_not_beat_wire(self):
        assert should_update("businesswire", "nasdaq") is False

    def test_ir_beats_nasdaq(self):
        assert should_update("nasdaq", "ir_page") is True

    def test_same_priority_updates(self):
        assert should_update("businesswire", "prnewswire") is True

    def test_estimated_does_not_beat_anything(self):
        assert should_update("nasdaq", "estimated") is False


class TestMergeEventData:
    def test_wire_overwrites_nasdaq_date(self):
        existing = {"event_date": date(2026, 4, 28), "source": "nasdaq", "event_time": None, "webcast_url": None}
        new = {"event_date": date(2026, 4, 22), "source": "businesswire", "event_time": "17:30:00", "webcast_url": "https://ir.tesla.com"}
        merged = merge_event_data(existing, new)
        assert merged["event_date"] == date(2026, 4, 22)
        assert merged["event_time"] == "17:30:00"
        assert merged["webcast_url"] == "https://ir.tesla.com"
        assert merged["source"] == "businesswire"

    def test_nasdaq_does_not_overwrite_wire_date(self):
        existing = {"event_date": date(2026, 4, 22), "source": "businesswire", "event_time": "17:30:00", "webcast_url": "https://ir.tesla.com"}
        new = {"event_date": date(2026, 4, 28), "source": "nasdaq", "event_time": None, "webcast_url": None}
        merged = merge_event_data(existing, new)
        assert merged["event_date"] == date(2026, 4, 22)  # Wire date preserved
        assert merged["event_time"] == "17:30:00"  # Wire time preserved
        assert merged["source"] == "businesswire"  # Wire source preserved

    def test_gap_filling_from_lower_source(self):
        """Lower priority source can fill gaps but not overwrite."""
        existing = {"event_date": date(2026, 4, 22), "source": "businesswire", "event_time": "17:30:00", "webcast_url": None}
        new = {"event_date": date(2026, 4, 22), "source": "nasdaq", "event_time": None, "webcast_url": None, "fiscal_quarter": "Q1 2026"}
        merged = merge_event_data(existing, new)
        assert merged["event_time"] == "17:30:00"  # Preserved
        assert merged["fiscal_quarter"] == "Q1 2026"  # Gap filled

    def test_ir_fills_wire_gaps(self):
        existing = {"event_date": date(2026, 4, 22), "source": "businesswire", "event_time": "17:30:00", "webcast_url": None}
        new = {"event_date": date(2026, 4, 22), "source": "ir_page", "webcast_url": "https://ir.company.com"}
        merged = merge_event_data(existing, new)
        assert merged["webcast_url"] == "https://ir.company.com"  # IR filled the gap
        assert merged["source"] == "businesswire"  # Source stays as wire (higher priority)

    def test_confirmed_status_when_complete(self):
        existing = {"event_date": date(2026, 4, 22), "source": "businesswire", "event_time": "17:30:00", "webcast_url": None}
        new = {"event_date": date(2026, 4, 22), "source": "ir_page", "webcast_url": "https://ir.company.com"}
        merged = merge_event_data(existing, new)
        assert merged["status"] == "confirmed"  # Has time + webcast now


class TestCleanWebcastUrl:
    def test_pr_page_url_rejected(self):
        assert clean_webcast_url("https://www.businesswire.com/news/home/123/en/Foo") is None
        assert clean_webcast_url("https://www.prnewswire.com/news-releases/foo-123.html") is None

    def test_valid_url_passes(self):
        assert clean_webcast_url("https://ir.tesla.com") == "https://ir.tesla.com"
        assert clean_webcast_url("https://investor.apple.com") == "https://investor.apple.com"

    def test_trailing_punctuation_stripped(self):
        assert clean_webcast_url("https://ir.tesla.com.") == "https://ir.tesla.com"

    def test_none_returns_none(self):
        assert clean_webcast_url(None) is None
        assert clean_webcast_url("") is None

    def test_no_protocol_rejected(self):
        assert clean_webcast_url("ir.tesla.com") is None


class TestCleanTime:
    def test_valid_time(self):
        assert clean_time("17:30:00") == "17:30:00"
        assert clean_time("08:00:00") == "08:00:00"

    def test_unreasonable_time_rejected(self):
        assert clean_time("03:00:00") is None  # 3 AM
        assert clean_time("23:30:00") is None  # 11:30 PM

    def test_none_returns_none(self):
        assert clean_time(None) is None
        assert clean_time("") is None


class TestFixWeekendDate:
    def test_saturday_to_monday(self):
        sat = date(2026, 4, 4)  # Saturday
        assert fix_weekend_date(sat) == date(2026, 4, 6)  # Monday

    def test_sunday_to_monday(self):
        sun = date(2026, 4, 5)  # Sunday
        assert fix_weekend_date(sun) == date(2026, 4, 6)  # Monday

    def test_weekday_unchanged(self):
        fri = date(2026, 4, 3)  # Friday
        assert fix_weekend_date(fri) == date(2026, 4, 3)
