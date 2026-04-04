"""Primary data source: Wire service press release monitoring.

Monitors PRNewswire and BusinessWire for advance earnings/event
announcements. These contain date, time, webcast URL, and dial-in details.

Two approaches:
1. RSS feeds (real-time monitoring, limited to recent entries)
2. Category page scraping (deeper archive access)

All events sourced from wire services are marked as confirmed since
the company themselves issued the press release.
"""
import logging
import re
import sys
from datetime import datetime, date, timedelta
from pathlib import Path
from typing import Optional

import httpx
import feedparser
from bs4 import BeautifulSoup

sys.path.insert(0, str(Path(__file__).parent.parent / "backend"))

from database import SessionLocal, init_db
from models import Event, Company

logger = logging.getLogger(__name__)

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
    "Accept": "text/html,application/xhtml+xml,application/json",
}

# ═══════════════════════════════════════════════════════════════
# ALL THREE WIRE SERVICES — FULL CATEGORY COVERAGE
# We monitor every relevant category, not just conference calls.
# ═══════════════════════════════════════════════════════════════

# PRNewswire RSS feeds — every financial category
PRN_RSS_FEEDS = [
    "https://www.prnewswire.com/rss/financial-services-latest-news/conference-call-announcements-list.rss",
    "https://www.prnewswire.com/rss/financial-services-latest-news/earnings-list.rss",
    "https://www.prnewswire.com/rss/financial-services-latest-news/earnings-forecasts-projections-list.rss",
    "https://www.prnewswire.com/rss/financial-services-latest-news.rss",  # All financial services
]

# PRNewswire archive pages — multiple categories
PRN_PAGES = {
    "conference_call": "https://www.prnewswire.com/news-releases/financial-services-latest-news/conference-call-announcements-list/?page={}&pagesize=100",
    "earnings": "https://www.prnewswire.com/news-releases/financial-services-latest-news/earnings-list/?page={}&pagesize=100",
    "earnings_forecasts": "https://www.prnewswire.com/news-releases/financial-services-latest-news/earnings-forecasts-projections-list/?page={}&pagesize=100",
}

# BusinessWire archive pages — multiple subjects
BW_PAGES = {
    "conference_call": "https://www.businesswire.com/newsroom/subject/conference-call?page={}",
    "earnings": "https://www.businesswire.com/newsroom/subject/earnings?page={}",
    "financial_results": "https://www.businesswire.com/newsroom/subject/financial-results?page={}",
}

# GlobeNewswire RSS feeds — earnings and corporate events
GNW_RSS_FEEDS = [
    "https://www.globenewswire.com/RssFeed/subjectcode/14-Earnings/feedTitle/GlobeNewswire+-+Earnings",
    "https://www.globenewswire.com/RssFeed/subjectcode/13-Conference+Calls/feedTitle/GlobeNewswire+-+Conference+Calls",
]

# GlobeNewswire archive pages
GNW_PAGES = {
    "earnings": "https://www.globenewswire.com/search/tag/earnings?page={}",
}


def scrape_all_rss() -> list[dict]:
    """Fetch latest announcements from ALL wire service RSS feeds.

    Covers PRNewswire, BusinessWire (via feedparser), and GlobeNewswire.
    No keyword filtering — we check every entry for ticker matches.
    """
    results = []
    feeds_checked = 0
    feeds_failed = 0

    all_feeds = [
        ("prnewswire", PRN_RSS_FEEDS),
        ("globenewswire", GNW_RSS_FEEDS),
    ]

    for source, feed_urls in all_feeds:
        for rss_url in feed_urls:
            feeds_checked += 1
            try:
                feed = feedparser.parse(rss_url)
                if feed.bozo and not feed.entries:
                    logger.warning("RSS feed failed: %s (%s)", rss_url[:60], feed.bozo_exception)
                    feeds_failed += 1
                    continue

                for entry in feed.entries:
                    title = entry.get("title", "")
                    link = entry.get("link", "")
                    summary = entry.get("summary", "")

                    # Extract ticker from title or summary
                    text = f"{title} {summary}"
                    ticker_match = re.search(
                        r'\((?:NYSE|NASDAQ|Nasdaq|AMEX|OTC|OTCQX):\s*"?(\w+)', text
                    )
                    if not ticker_match:
                        continue

                    ticker = ticker_match.group(1).upper()

                    # No keyword filtering — parse everything with a ticker
                    event = _parse_announcement(title, summary, ticker, link, source)
                    if event:
                        results.append(event)

            except Exception as e:
                logger.error("RSS feed error %s: %s", rss_url[:60], e)
                feeds_failed += 1

    logger.info("RSS: checked %d feeds (%d failed), found %d events", feeds_checked, feeds_failed, len(results))
    return results


# Keep old name as alias for backward compatibility
def scrape_prnewswire_rss() -> list[dict]:
    return scrape_all_rss()


def scrape_prnewswire_pages(max_pages: int = 20) -> list[dict]:
    """Scrape PRNewswire category pages — fetch PR bodies for full detail extraction."""
    results = []
    pr_items = []

    # Collect PR URLs and tickers from category pages
    for category, page_url_template in PRN_PAGES.items():
        for page in range(1, max_pages + 1):
            try:
                url = page_url_template.format(page)
                resp = httpx.get(url, headers=HEADERS, timeout=30, follow_redirects=True)
                soup = BeautifulSoup(resp.text, "lxml")
                links = soup.select("a.newsreleaseconsolidatelink")

                if not links:
                    break

                for link_el in links:
                    title = link_el.get_text(strip=True)
                    href = link_el.get("href", "")
                    if href and not href.startswith("http"):
                        href = f"https://www.prnewswire.com{href}"

                    ticker_match = re.search(
                        r'\((?:NYSE|NASDAQ|Nasdaq|AMEX|OTC):\s*"?(\w+)', title
                    )
                    if not ticker_match:
                        continue

                    ticker = ticker_match.group(1).upper()
                    pr_items.append((href, title, ticker))

            except Exception as e:
                logger.error("PRNewswire %s page %d failed: %s", category, page, e)

    logger.info("PRNewswire: collected %d PR items, fetching bodies...", len(pr_items))

    # Fetch each PR body for full detail extraction — no keyword filtering
    for href, title, ticker in pr_items:
        try:
            resp = httpx.get(href, headers=HEADERS, timeout=10, follow_redirects=True)
            soup = BeautifulSoup(resp.text, "lxml")

            # Multiple fallback selectors for PRNewswire
            article = (
                soup.select_one(".release-body")
                or soup.select_one(".prnewswire-body")
                or soup.select_one("article")
                or soup.select_one(".entry-content")
                or soup.select_one("main")
            )
            if article:
                text = article.get_text(separator=" ")
            else:
                body = soup.select_one("body")
                text = body.get_text(separator=" ") if body else title

            if len(text.strip()) < 50:
                text = title  # Fall back to title if extraction failed

            event = _parse_pr_body(text, ticker, title, href, "prnewswire")
            if event:
                results.append(event)
            else:
                # Try title-only parsing as fallback
                event = _parse_announcement(title, "", ticker, href, "prnewswire")
                if event:
                    results.append(event)
        except Exception as e:
            # Fall back to title-only parsing
            event = _parse_announcement(title, "", ticker, href, "prnewswire")
            if event:
                results.append(event)

    logger.info("PRNewswire pages: found %d events with details", len(results))
    return results


def scrape_businesswire_pages(max_pages: int = 100) -> list[dict]:
    """Scrape BusinessWire newsroom pages (all categories).

    BW requires fetching each PR page individually to extract ticker and details.
    """
    results = []
    pr_urls = []

    # Collect PR URLs from ALL category pages
    for category, page_url_template in BW_PAGES.items():
        for page in range(1, max_pages + 1):
            try:
                url = page_url_template.format(page)
                resp = httpx.get(url, headers=HEADERS, timeout=15, follow_redirects=True)
                soup = BeautifulSoup(resp.text, "lxml")
                links = soup.select('a[href*="/news/home/"]')

                page_urls = []
                for link_el in links:
                    title = link_el.get_text(strip=True)
                    href = link_el.get("href", "")
                    if len(title) > 20 and href:
                        full_url = f"https://www.businesswire.com{href}" if href.startswith("/") else href
                        page_urls.append((full_url, title))

                if not page_urls:
                    break

                pr_urls.extend(page_urls)

            except Exception as e:
                logger.error("BusinessWire %s page %d failed: %s", category, page, e)
                break

    logger.info("BusinessWire (%d categories): collected %d PR URLs, fetching details...", len(BW_PAGES), len(pr_urls))

    # Fetch each PR — extract article text (not raw HTML), check for ANY event details
    for pr_url, title in pr_urls:
        try:
            resp = httpx.get(pr_url, headers=HEADERS, timeout=10, follow_redirects=True)

            # Extract article text — multiple fallback selectors for BW
            soup = BeautifulSoup(resp.text, "lxml")
            article = (
                soup.select_one(".bw-release-story")
                or soup.select_one(".bw-release-body")
                or soup.select_one("#press-release-body")
                or soup.select_one("article")
                or soup.select_one(".entry-content")
                or soup.select_one("main")
            )
            if article:
                text = article.get_text(separator=" ")
            else:
                # Last resort: get all text from body, skip nav/header
                body = soup.select_one("body")
                text = body.get_text(separator=" ") if body else resp.text[:15000]

            # Ensure we have meaningful text
            if len(text.strip()) < 50:
                text = soup.get_text(separator=" ")

            # Extract ticker from full page (ticker might be in header, not article)
            full_text = resp.text[:10000]
            ticker_match = re.search(
                r'(?:NYSE|NASDAQ|Nasdaq|AMEX|OTCQX):\s*(\w+)', full_text
            )
            if not ticker_match:
                continue

            ticker = ticker_match.group(1).upper()

            # Parse event details from article text — no keyword filtering
            event = _parse_pr_body(text, ticker, title, pr_url, "businesswire")
            if event:
                results.append(event)

        except Exception as e:
            logger.debug("BW PR fetch failed %s: %s", pr_url[:60], e)

    logger.info("BusinessWire: extracted %d events with details", len(results))
    return results


def scrape_globenewswire(max_pages: int = 10) -> list[dict]:
    """Scrape GlobeNewswire for earnings and event announcements.

    GlobeNewswire is the third major wire service. Some companies
    use it exclusively.
    """
    results = []

    # RSS feeds first
    for rss_url in GNW_RSS_FEEDS:
        try:
            feed = feedparser.parse(rss_url)
            for entry in feed.entries:
                title = entry.get("title", "")
                link = entry.get("link", "")
                summary = entry.get("summary", "")

                text = f"{title} {summary}"
                ticker_match = re.search(
                    r'\((?:NYSE|NASDAQ|Nasdaq|AMEX|OTC|OTCQX|TSX):\s*"?(\w+)', text
                )
                if not ticker_match:
                    continue

                ticker = ticker_match.group(1).upper()
                event = _parse_announcement(title, summary, ticker, link, "globenewswire")
                if event:
                    results.append(event)

        except Exception as e:
            logger.error("GlobeNewswire RSS error: %s", e)

    # Archive pages
    for category, page_url_template in GNW_PAGES.items():
        for page in range(1, max_pages + 1):
            try:
                url = page_url_template.format(page)
                resp = httpx.get(url, headers=HEADERS, timeout=15, follow_redirects=True)
                soup = BeautifulSoup(resp.text, "lxml")

                # GlobeNewswire uses different selectors
                links = soup.select("a[href*='/news-release/']")
                if not links:
                    break

                for link_el in links:
                    title = link_el.get_text(strip=True)
                    href = link_el.get("href", "")
                    if len(title) < 20:
                        continue

                    if href and not href.startswith("http"):
                        href = f"https://www.globenewswire.com{href}"

                    ticker_match = re.search(
                        r'\((?:NYSE|NASDAQ|Nasdaq|AMEX|OTC|OTCQX|TSX):\s*"?(\w+)', title
                    )
                    if not ticker_match:
                        continue

                    ticker = ticker_match.group(1).upper()

                    # Try to fetch PR body for details
                    try:
                        pr_resp = httpx.get(href, headers=HEADERS, timeout=10, follow_redirects=True)
                        pr_soup = BeautifulSoup(pr_resp.text, "lxml")
                        article = (
                            pr_soup.select_one(".main-body-container")
                            or pr_soup.select_one("article")
                            or pr_soup.select_one(".entry-content")
                        )
                        text = article.get_text(separator=" ") if article else title

                        event = _parse_pr_body(text, ticker, title, href, "globenewswire")
                        if event:
                            results.append(event)
                    except Exception:
                        event = _parse_announcement(title, "", ticker, href, "globenewswire")
                        if event:
                            results.append(event)

            except Exception as e:
                logger.error("GlobeNewswire %s page %d failed: %s", category, page, e)

    logger.info("GlobeNewswire: found %d events", len(results))
    return results


def _parse_announcement(title: str, summary: str, ticker: str, url: str, source: str) -> Optional[dict]:
    """Parse an event announcement using Gemini LLM.

    Used for RSS entries and title-only cases where we have limited text.
    No regex fallback.
    """
    text = f"{title} {summary}".strip()
    if len(text) < 20:
        return None

    try:
        from llm_extractor import extract_event_from_text
        result = extract_event_from_text(text, ticker)
        if result:
            result["source"] = source
            result["source_url"] = url
            return result
        return None
    except ImportError:
        logger.error("llm_extractor not available")
        return None
    except Exception as e:
        logger.error("LLM extraction failed for %s: %s", ticker, e)
        return None


def _parse_announcement_DEPRECATED(title: str, summary: str, ticker: str, url: str, source: str) -> Optional[dict]:
    """DEPRECATED: Old regex-based announcement parser. Not used."""
    text = f"{title} {summary}"

    event_date = _extract_event_date(text)
    if not event_date:
        event_date = _extract_date(text)
    if not event_date:
        return None

    event_type = "earnings"
    lower = text.lower()
    if any(kw in lower for kw in ["investor day", "investor forum", "capital markets", "analyst day"]):
        event_type = "investor_day"
    elif any(kw in lower for kw in ["conference presentation", "industry conference", "summit"]):
        event_type = "conference"
    elif any(kw in lower for kw in ["acquisition", "merger", "strategic update", "special call"]):
        event_type = "ad_hoc"

    event_time = _extract_time(text)
    webcast_url = _extract_webcast_url(text)
    phone_number = _extract_phone(text)
    phone_passcode = _extract_passcode(text)
    fiscal_quarter = _extract_fiscal_quarter(text)

    return {
        "ticker": ticker,
        "event_type": event_type,
        "event_date": event_date,
        "event_time": event_time,
        "title": _clean_title(title, ticker),
        "webcast_url": webcast_url,
        "phone_number": phone_number,
        "phone_passcode": phone_passcode,
        "fiscal_quarter": fiscal_quarter,
        "source": source,
        "source_url": url,
    }


def _parse_pr_body(text: str, ticker: str, title: str, url: str, source: str) -> Optional[dict]:
    """Parse event details from a PR body using Gemini LLM.

    No regex fallback — if the LLM can't extract it, we return None
    rather than risk storing incorrect data.
    """
    try:
        from llm_extractor import extract_event_from_text
        result = extract_event_from_text(text, ticker)
        if result:
            result["source"] = source
            result["source_url"] = url
            if not result.get("title"):
                result["title"] = _clean_title(title, ticker)
            return result
        return None
    except ImportError:
        logger.error("llm_extractor not available — cannot parse PR body")
        return None
    except Exception as e:
        logger.error("LLM extraction failed for %s: %s", ticker, e)
        return None


def _parse_pr_body_DEPRECATED(text: str, ticker: str, title: str, url: str, source: str) -> Optional[dict]:
    """DEPRECATED: Old regex-based parser. Kept only for reference, not used."""
    event_date = _extract_event_date(text)
    if not event_date:
        event_date = _extract_date(text)
    if not event_date:
        return None

    event_type = "earnings"
    lower = text.lower()
    if any(kw in lower for kw in ["investor day", "investor forum", "capital markets", "analyst day"]):
        event_type = "investor_day"
    elif any(kw in lower for kw in ["conference presentation", "industry conference", "summit"]):
        event_type = "conference"
    elif any(kw in lower for kw in ["acquisition", "merger", "strategic update", "special call"]):
        event_type = "ad_hoc"

    return {
        "ticker": ticker,
        "event_type": event_type,
        "event_date": event_date,
        "event_time": _extract_time(text),
        "title": _clean_title(title, ticker),
        "webcast_url": _extract_webcast_url(text),
        "phone_number": _extract_phone(text),
        "phone_passcode": _extract_passcode(text),
        "fiscal_quarter": _extract_fiscal_quarter(text),
        "source": source,
        "source_url": url,
    }


def _extract_date(text: str) -> Optional[date]:
    """Extract event date from press release text."""
    # Look for explicit date patterns
    patterns = [
        # "on April 15, 2026" or "on Thursday, April 15, 2026"
        r'(?:on\s+)?(?:Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)?,?\s*(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{1,2}),?\s+(\d{4})',
        # "April 15, 2026"
        r'(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{1,2}),?\s+(\d{4})',
    ]

    months = {
        "January": 1, "February": 2, "March": 3, "April": 4,
        "May": 5, "June": 6, "July": 7, "August": 8,
        "September": 9, "October": 10, "November": 11, "December": 12,
    }

    for pattern in patterns:
        matches = re.findall(pattern, text)
        for match in matches:
            try:
                month_name, day, year = match[-3], match[-2], match[-1]
                d = date(int(year), months[month_name], int(day))
                # Only accept future dates (or very recent)
                if d >= datetime.utcnow().date() - timedelta(days=7):
                    return d
            except (ValueError, KeyError):
                continue

    return None


def _extract_time(text: str) -> Optional[str]:
    """Extract conference call time from text. Returns HH:MM:SS in ET.

    Handles formats like:
    - "4:30 p.m. ET"
    - "9 a.m. ET"
    - "1:45 p.m. Pacific Time"
    - "10:00 a.m. EDT"
    - "8:30 a.m. Eastern Time"
    """
    # Pattern: H:MM am/pm TIMEZONE or H am/pm TIMEZONE
    tz_pattern = r'(?:ET|EST|EDT|CT|CST|CDT|PT|PST|PDT|MT|MST|MDT|Eastern\s+Time|Pacific\s+Time|Central\s+Time)'
    # Prefer ET/Eastern over other timezones — search for ET first
    et_tz = r'(?:ET|EST|EDT|Eastern\s+Time)'
    pattern_et = rf'(\d{{1,2}}):(\d{{2}})\s*(a\.?m\.?|p\.?m\.?|AM|PM)\s*{et_tz}'
    match = re.search(pattern_et, text, re.IGNORECASE)
    if not match:
        # Fall back to any timezone
        pattern_hhmm = rf'(\d{{1,2}}):(\d{{2}})\s*(a\.?m\.?|p\.?m\.?|AM|PM)\s*(?:/\s*\d{{1,2}}:\d{{2}}\s*(?:a\.?m\.?|p\.?m\.?)\s*)?{tz_pattern}'
        match = re.search(pattern_hhmm, text, re.IGNORECASE)
    if match:
        hour = int(match.group(1))
        minute = int(match.group(2))
        ampm = match.group(3).lower().replace(".", "")
        if ampm == "pm" and hour != 12:
            hour += 12
        elif ampm == "am" and hour == 12:
            hour = 0
        return f"{hour:02d}:{minute:02d}:00"

    # Without minutes: "9 a.m. ET"
    pattern_h = rf'(\d{{1,2}})\s*(a\.?m\.?|p\.?m\.?|AM|PM)\s*{tz_pattern}'
    match = re.search(pattern_h, text, re.IGNORECASE)
    if match:
        hour = int(match.group(1))
        ampm = match.group(2).lower().replace(".", "")
        if ampm == "pm" and hour != 12:
            hour += 12
        elif ampm == "am" and hour == 12:
            hour = 0
        return f"{hour:02d}:00:00"

    return None


def _extract_webcast_url(text: str) -> Optional[str]:
    """Extract webcast/investor relations URL from text.

    Handles both full URLs (https://...) and bare domains (investor.example.com).
    """
    # Full URLs first
    full_url_patterns = [
        r'https?://[^\s<"\',(]+(?:investor|ir\.|event|q4inc|notified|webcast|youtube)[^\s<"\',(]*',
        r'https?://[^\s<"\',(]+(?:earnings|conference)[^\s<"\',(]*',
    ]
    for pattern in full_url_patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            url = match.group(0).rstrip(".,;)")
            return url

    # Bare domain patterns: "investor.example.com" or "ir.example.com/..."
    bare_patterns = [
        r'((?:investor|ir|investors)[.\s]+\w+\.com[^\s<"\',(]*)',
        r'(youtube\.com/\w+)',
    ]
    for pattern in bare_patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            domain = match.group(1).strip().rstrip(".,;)")
            if not domain.startswith("http"):
                domain = "https://" + domain
            return domain

    return None


def _extract_phone(text: str) -> Optional[str]:
    """Extract dial-in phone number from text."""
    # Look for phone numbers near dial-in keywords
    pattern = r'(?:dial(?:ing)?[- ]?in|telephone|call[- ]?in|(?:participants?|callers?)\s+(?:can\s+)?(?:access|call|dial)|by\s+dialing)[:\s]*(\+?1?[-.\s]?\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4})'
    match = re.search(pattern, text, re.IGNORECASE)
    if match:
        return match.group(1).strip()
    return None


def _extract_passcode(text: str) -> Optional[str]:
    """Extract passcode/access code from text."""
    pattern = r'(?:passcode|access code|conference id|pin|code)\s*(?:is\s+)?[:\s]*(\d{4,10})'
    match = re.search(pattern, text, re.IGNORECASE)
    if match:
        return match.group(1)
    return None


def _extract_event_date(text: str) -> Optional[date]:
    """Extract the EVENT date (conference call/webcast), not the release date.

    When a PR mentions both a release date and a call date:
    "will release earnings on April 27... conference call on April 28"
    we want April 28 (the call date), not April 27 (the release date).

    Strategy:
    1. First, look for dates near "conference call", "webcast", "call" keywords
    2. If no call-specific date, look for dates near generic event keywords
    3. Fallback: use the latest future date mentioned
    """
    from datetime import datetime as dt
    months = r'(?:January|February|March|April|May|June|July|August|September|October|November|December)'
    date_pattern = rf'({months})\s+(\d{{1,2}})(?:st|nd|rd|th)?,?\s+(\d{{4}})'

    def parse_match(m):
        try:
            return dt.strptime(f"{m.group(1)} {m.group(2)} {m.group(3)}", "%B %d %Y").date()
        except ValueError:
            return None

    # Find ALL dates in the text
    all_date_matches = list(re.finditer(date_pattern, text, re.IGNORECASE))
    if not all_date_matches:
        # No standard dates found — fall through to pattern 3 below
        pass
    else:
        # Score each date by proximity to call/webcast keywords
        # The date closest to "conference call", "webcast", "host", etc. wins
        call_kw_positions = [m.start() for m in re.finditer(
            r'conference\s+call|webcast|investor\s+(?:call|update|forum|day)|'
            r'host\s+(?:a\s+|an\s+|its\s+)?(?:live|first|second|third|fourth|q[1-4])?',
            text, re.IGNORECASE
        )]

        if call_kw_positions:
            # Score by proximity to call keywords
            best_match = None
            best_score = -1
            for m in all_date_matches:
                d = parse_match(m)
                if not d:
                    continue
                min_dist = min(abs(m.start() - pos) for pos in call_kw_positions)
                score = 10000 - min_dist
                if score > best_score:
                    best_score = score
                    best_match = d
            if best_match:
                return best_match

        # No call keywords — use date after event keyword "on"
        event_keywords = r'(?:on|results\s+on|call\s+on|report\s+on|webcast\s+on)'
        pattern1 = rf'{event_keywords}\s+(?:Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)?,?\s*{date_pattern}'
        match = re.search(pattern1, text, re.IGNORECASE)
        if match:
            d = parse_match(match)
            if d:
                return d

        # Fallback: use the last date mentioned (most likely the event)
        for m in reversed(all_date_matches):
            d = parse_match(m)
            if d:
                return d

    # Pattern 3: "November 18, at 9 a.m." (no year — after event keyword)
    event_keywords = r'(?:on|results\s+on|call\s+on|report\s+on|webcast\s+on)'
    pattern3 = rf'{event_keywords}\s+(?:Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)?,?\s*({months})\s+(\d{{1,2}})(?:st|nd|rd|th)?,?\s+at\s+'
    match = re.search(pattern3, text, re.IGNORECASE)
    if match:
        now = dt.now()
        try:
            d = dt.strptime(f"{match.group(1)} {match.group(2)} {now.year}", "%B %d %Y").date()
            if d < now.date():
                d = dt.strptime(f"{match.group(1)} {match.group(2)} {now.year + 1}", "%B %d %Y").date()
            return d
        except ValueError:
            pass

    # Pattern 4: "Month DD, at" without keyword (no year)
    pattern4 = rf'({months})\s+(\d{{1,2}})(?:st|nd|rd|th)?,?\s+at\s+'
    match = re.search(pattern4, text, re.IGNORECASE)
    if match:
        now = dt.now()
        try:
            d = dt.strptime(f"{match.group(1)} {match.group(2)} {now.year}", "%B %d %Y").date()
            if d < now.date():
                d = dt.strptime(f"{match.group(1)} {match.group(2)} {now.year + 1}", "%B %d %Y").date()
            return d
        except ValueError:
            pass

    return None


def _extract_ticker(text: str) -> Optional[str]:
    """Extract stock ticker from text.

    Looks for patterns like (NYSE: CVX), (NASDAQ: META), (AMEX: XYZ).
    """
    pattern = r'\((?:NYSE|NASDAQ|Nasdaq|AMEX|OTC|OTCQX):\s*"?(\w+)"?\)'
    match = re.search(pattern, text, re.IGNORECASE)
    if match:
        return match.group(1).upper()
    return None


def _extract_fiscal_quarter(text: str) -> Optional[str]:
    """Extract fiscal quarter reference from text."""
    pattern = r'(?:first|second|third|fourth|1st|2nd|3rd|4th)\s+(?:quarter|fiscal quarter)\s+(?:of\s+)?(?:fiscal\s+(?:year\s+)?)?(\d{4})'
    match = re.search(pattern, text, re.IGNORECASE)
    if match:
        quarter_word = text[match.start():match.start()+10].lower()
        q_map = {"first": "Q1", "second": "Q2", "third": "Q3", "fourth": "Q4",
                 "1st": "Q1", "2nd": "Q2", "3rd": "Q3", "4th": "Q4"}
        for word, q in q_map.items():
            if word in quarter_word:
                return f"{q} {match.group(1)}"
        return f"FY {match.group(1)}"
    return None


def _clean_title(title: str, ticker: str) -> str:
    """Clean up PR title for use as event title."""
    # Remove timestamp prefix (e.g., "16:15 ET")
    title = re.sub(r'^\d{1,2}:\d{2}\s*ET\s*', '', title)
    # Remove ticker/exchange reference
    title = re.sub(r'\s*\((?:NYSE|NASDAQ|Nasdaq|AMEX|OTC):\s*"?\w+"?\)\s*', ' ', title)
    # Truncate
    return title.strip()[:300]


def store_wire_events(events: list[dict], db=None):
    """Store wire service events as confirmed in the database.

    Wire service events are from the company directly, so they're confirmed.
    """
    close_session = False
    if db is None:
        init_db()
        db = SessionLocal()
        close_session = True

    try:
        known_tickers = {r[0] for r in db.query(Company.ticker).all()}
        ticker_to_name = {
            r[0]: r[1] for r in db.query(Company.ticker, Company.company_name).all()
        }

        inserted = 0
        updated = 0
        skipped = 0

        for ev in events:
            ticker = ev["ticker"]
            if ticker not in known_tickers:
                skipped += 1
                continue

            event_date = ev["event_date"]
            if isinstance(event_date, str):
                event_date = datetime.strptime(event_date, "%Y-%m-%d").date()

            # Check for existing event — exact date match first
            existing = (
                db.query(Event)
                .filter(
                    Event.ticker == ticker,
                    Event.event_date == event_date,
                    Event.event_type == ev["event_type"],
                )
                .first()
            )

            # If no exact match, look for a nearby event from the same ticker
            # (e.g., Nasdaq said April 28 but wire says April 22 — correct the date)
            if not existing:
                from datetime import timedelta as td
                nearby = (
                    db.query(Event)
                    .filter(
                        Event.ticker == ticker,
                        Event.event_type == ev["event_type"],
                        Event.event_date >= event_date - td(days=14),
                        Event.event_date <= event_date + td(days=14),
                        Event.source.in_(["nasdaq", "estimated", "historical_pattern"]),
                    )
                    .first()
                )
                if nearby:
                    # Wire service is more authoritative — correct the date
                    nearby.event_date = event_date
                    existing = nearby

            if existing:
                # Update with wire service data (more authoritative)
                changed = False
                if ev.get("event_time") and not existing.event_time:
                    existing.event_time = datetime.strptime(ev["event_time"], "%H:%M:%S").time() if isinstance(ev["event_time"], str) else ev["event_time"]
                    changed = True
                if ev.get("webcast_url") and not existing.webcast_url:
                    existing.webcast_url = ev["webcast_url"]
                    changed = True
                if ev.get("phone_number") and not existing.phone_number:
                    existing.phone_number = ev["phone_number"]
                    changed = True
                if ev.get("phone_passcode") and not existing.phone_passcode:
                    existing.phone_passcode = ev["phone_passcode"]
                    changed = True
                if ev.get("title") and not existing.title:
                    existing.title = ev["title"]
                    changed = True
                if ev.get("fiscal_quarter") and not existing.fiscal_quarter:
                    existing.fiscal_quarter = ev["fiscal_quarter"]
                    changed = True

                # Wire service with webcast/time details = confirmed
                # Without details = still tentative (just has date)
                has_details = bool(existing.webcast_url or existing.event_time or ev.get("webcast_url") or ev.get("event_time"))
                existing.ir_verified = has_details
                existing.status = "confirmed" if has_details else "tentative"
                existing.source = ev["source"]
                existing.source_url = ev.get("source_url")
                existing.updated_at = datetime.utcnow()

                if changed:
                    updated += 1
                else:
                    skipped += 1
            else:
                # New event
                company_name = ticker_to_name.get(ticker, ticker)
                event_time = None
                if ev.get("event_time"):
                    if isinstance(ev["event_time"], str):
                        event_time = datetime.strptime(ev["event_time"], "%H:%M:%S").time()
                    else:
                        event_time = ev["event_time"]

                new_event = Event(
                    ticker=ticker,
                    company_name=company_name,
                    event_type=ev["event_type"],
                    event_date=event_date,
                    event_time=event_time,
                    title=ev.get("title"),
                    webcast_url=ev.get("webcast_url"),
                    phone_number=ev.get("phone_number"),
                    phone_passcode=ev.get("phone_passcode"),
                    fiscal_quarter=ev.get("fiscal_quarter"),
                    source=ev["source"],
                    source_url=ev.get("source_url"),
                    ir_verified=bool(ev.get("webcast_url") or event_time),
                    status="confirmed" if (ev.get("webcast_url") or event_time) else "tentative",
                )
                db.add(new_event)
                inserted += 1

        db.commit()
        result = {"inserted": inserted, "updated": updated, "skipped": skipped}
        logger.info("Wire events stored: %s", result)
        return result

    finally:
        if close_session:
            db.close()


def run_wire_scan(prn_pages: int = 20, bw_pages: int = 50, gnw_pages: int = 10):
    """Run full wire service scan — all three services, all categories."""
    all_events = []

    # 1. ALL RSS feeds (PRNewswire + GlobeNewswire)
    try:
        events = scrape_all_rss()
        all_events.extend(events)
    except Exception as e:
        logger.error("RSS scan failed: %s", e)

    # 2. PRNewswire category pages (deeper archive)
    try:
        events = scrape_prnewswire_pages(max_pages=prn_pages)
        all_events.extend(events)
    except Exception as e:
        logger.error("PRNewswire pages failed: %s", e)

    # 3. BusinessWire pages
    try:
        events = scrape_businesswire_pages(max_pages=bw_pages)
        all_events.extend(events)
    except Exception as e:
        logger.error("BusinessWire failed: %s", e)

    # 4. GlobeNewswire pages
    try:
        events = scrape_globenewswire(max_pages=gnw_pages)
        all_events.extend(events)
    except Exception as e:
        logger.error("GlobeNewswire failed: %s", e)

    # Deduplicate by ticker + date + type
    seen = set()
    unique_events = []
    for ev in all_events:
        key = (ev["ticker"], str(ev["event_date"]), ev["event_type"])
        if key not in seen:
            seen.add(key)
            unique_events.append(ev)

    logger.info("Wire scan total: %d unique events from %d raw (PRN+BW+GNW)", len(unique_events), len(all_events))

    # Store
    result = store_wire_events(unique_events)
    return result


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(name)s %(levelname)s %(message)s")

    # Quick test with RSS only
    import sys
    if "--full" in sys.argv:
        result = run_wire_scan(prn_pages=20, bw_pages=50)
    else:
        result = run_wire_scan(prn_pages=5, bw_pages=10)
    print(f"Done: {result}")
