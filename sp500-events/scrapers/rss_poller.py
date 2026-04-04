"""Real-time RSS polling for wire service announcements.

Polls PRNewswire and BusinessWire RSS feeds every 15 minutes.
Designed to run as a scheduled task or background process.

Usage:
    python rss_poller.py          # Poll once
    python rss_poller.py --loop   # Poll continuously every 15 minutes
"""
import argparse
import logging
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent / "backend"))

from database import SessionLocal, init_db
from event_store import bulk_upsert
from wire_services import scrape_prnewswire_rss

logger = logging.getLogger(__name__)


def poll_once() -> dict:
    """Poll all RSS feeds once and store results."""
    init_db()
    db = SessionLocal()

    try:
        all_events = []

        # PRNewswire RSS feeds
        try:
            events = scrape_prnewswire_rss()
            all_events.extend(events)
            logger.info("PRNewswire RSS: %d events", len(events))
        except Exception as e:
            logger.error("PRNewswire RSS failed: %s", e)

        # Store through quality gates
        if all_events:
            result = bulk_upsert(db, all_events)
            logger.info("RSS poll result: %s", result)
            return result
        else:
            logger.info("RSS poll: no new events")
            return {"inserted": 0, "updated": 0, "skipped": 0, "rejected": 0}

    finally:
        db.close()


def poll_loop(interval_seconds: int = 900):
    """Poll RSS feeds continuously."""
    logger.info("Starting RSS poll loop (every %d seconds)", interval_seconds)
    while True:
        try:
            result = poll_once()
            logger.info("Poll complete: %s", result)
        except Exception as e:
            logger.error("Poll failed: %s", e)
        time.sleep(interval_seconds)


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--loop", action="store_true", help="Poll continuously")
    parser.add_argument("--interval", type=int, default=900, help="Seconds between polls")
    args = parser.parse_args()

    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")

    if args.loop:
        poll_loop(args.interval)
    else:
        poll_once()
