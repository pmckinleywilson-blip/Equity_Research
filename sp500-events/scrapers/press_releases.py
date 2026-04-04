"""Detection Tier: Press release wire service RSS monitoring.

Monitors BusinessWire, PRNewswire, and GlobeNewswire for
earnings announcements and event details.
"""
import logging
import re
import sys
from datetime import datetime
from pathlib import Path

import feedparser

sys.path.insert(0, str(Path(__file__).parent.parent / "backend"))

from database import SessionLocal, init_db
from models import Company

logger = logging.getLogger(__name__)

RSS_FEEDS = {
    "businesswire": "https://feed.businesswire.com/rss/home/?rss=G1QFDERJXkJeGVtTTg==",
    "prnewswire": "https://www.prnewswire.com/rss/financial-services-latest-news/financial-services-latest-news-list.rss",
    "globenewswire": "https://www.globenewswire.com/RssFeed/subjectcode/14-Earnings/feedTitle/GlobeNewswire+-+Earnings",
}


def fetch_press_release_events() -> list[dict]:
    """Fetch earnings-related press releases from wire services."""
    init_db()
    db = SessionLocal()

    # Get all known tickers for matching
    known_tickers = {r[0] for r in db.query(Company.ticker).all()}
    ticker_to_name = {
        r[0]: r[1] for r in db.query(Company.ticker, Company.company_name).all()
    }
    db.close()

    all_events = []

    for source, feed_url in RSS_FEEDS.items():
        try:
            feed = feedparser.parse(feed_url)
            for entry in feed.entries[:50]:  # Limit to recent entries
                title = entry.get("title", "")
                summary = entry.get("summary", "")
                text = f"{title} {summary}"

                # Try to match ticker symbols
                matched_ticker = None
                for ticker in known_tickers:
                    # Look for ticker in parentheses: (AAPL) or (NYSE: AAPL) or (NASDAQ: AAPL)
                    patterns = [
                        rf'\({ticker}\)',
                        rf'\(NYSE:\s*{ticker}\)',
                        rf'\(NASDAQ:\s*{ticker}\)',
                        rf'\(AMEX:\s*{ticker}\)',
                    ]
                    for pattern in patterns:
                        if re.search(pattern, text, re.IGNORECASE):
                            matched_ticker = ticker
                            break
                    if matched_ticker:
                        break

                if not matched_ticker:
                    continue

                # Check if it's earnings-related
                earnings_keywords = [
                    "earnings", "quarterly results", "financial results",
                    "conference call", "webcast", "investor",
                ]
                is_earnings = any(kw in text.lower() for kw in earnings_keywords)

                if not is_earnings:
                    continue

                # Extract date from entry
                published = entry.get("published_parsed")
                if published:
                    pub_date = datetime(*published[:6]).strftime("%Y-%m-%d")
                else:
                    pub_date = datetime.utcnow().strftime("%Y-%m-%d")

                # Extract webcast URL from text
                webcast_url = None
                url_match = re.search(
                    r'https?://[^\s<"]+(?:webcast|event|q4inc|notified)[^\s<"]*',
                    text,
                    re.IGNORECASE,
                )
                if url_match:
                    webcast_url = url_match.group(0).rstrip(".")

                all_events.append({
                    "ticker": matched_ticker,
                    "company_name": ticker_to_name.get(matched_ticker, matched_ticker),
                    "event_type": "earnings",
                    "event_date": pub_date,  # Approximate — PR date, not event date
                    "title": title[:500],
                    "webcast_url": webcast_url,
                    "source": f"press_{source}",
                    "source_url": entry.get("link", ""),
                })

        except Exception as e:
            logger.error("RSS feed %s failed: %s", source, e)

    logger.info("Press releases: found %d earnings-related items", len(all_events))
    return all_events


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    events = fetch_press_release_events()
    for e in events[:10]:
        print(f"  {e['ticker']}: {e['title'][:80]}")
