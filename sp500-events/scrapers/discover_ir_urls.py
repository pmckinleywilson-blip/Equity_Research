"""Discover investor relations page URLs for all companies.

Most companies follow predictable IR URL patterns. This script:
1. Tries common URL patterns for each company
2. Verifies the page loads and contains event-related content
3. Stores working URLs in the company database
4. Classifies the IR platform (Q4, Notified, etc.)

Run: python discover_ir_urls.py [--ticker AAPL] [--sp500-only] [--limit 100]
"""
import argparse
import logging
import re
import sys
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from urllib.parse import urlparse

import httpx

sys.path.insert(0, str(Path(__file__).parent.parent / "backend"))

from database import SessionLocal, init_db
from models import Company

logger = logging.getLogger(__name__)

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
    "Accept": "text/html,application/xhtml+xml",
}

# Common IR URL patterns to try
# {domain} is derived from the company name or known domain
IR_URL_PATTERNS = [
    "https://investor.{domain}/events-and-presentations",
    "https://investor.{domain}/events",
    "https://investors.{domain}/events-and-presentations",
    "https://investors.{domain}/events",
    "https://ir.{domain}/events",
    "https://ir.{domain}/",
    "https://www.{domain}/investor-relations",
    "https://www.{domain}/investors",
    "https://investor.{domain}/investor-relations/default.aspx",
    "https://investor.{domain}",
    "https://investors.{domain}",
]

# Known domain mappings for companies where the domain isn't obvious
KNOWN_DOMAINS = {
    "AAPL": "apple.com",
    "MSFT": "microsoft.com",
    "GOOGL": "abc.xyz",
    "GOOG": "abc.xyz",
    "AMZN": "aboutamazon.com",
    "META": "atmeta.com",
    "TSLA": "tesla.com",
    "NVDA": "nvidia.com",
    "BRK-B": "berkshirehathaway.com",
    "JPM": "jpmorganchase.com",
    "V": "visa.com",
    "MA": "mastercard.com",
    "UNH": "unitedhealthgroup.com",
    "JNJ": "jnj.com",
    "WMT": "walmart.com",
    "PG": "pg.com",
    "HD": "homedepot.com",
    "BAC": "bankofamerica.com",
    "KO": "coca-colacompany.com",
    "PEP": "pepsico.com",
    "COST": "costco.com",
    "DIS": "thewaltdisneycompany.com",
    "NFLX": "netflix.net",
    "CRM": "salesforce.com",
    "ADBE": "adobe.com",
    "INTC": "intc.com",
    "AMD": "amd.com",
    "AVGO": "broadcom.com",
    "CSCO": "cisco.com",
    "PFE": "pfizer.com",
    "MRK": "merck.com",
    "ABBV": "abbvie.com",
    "LLY": "lilly.com",
    "TMO": "thermofisher.com",
    "GS": "goldmansachs.com",
    "MS": "morganstanley.com",
    "BLK": "blackrock.com",
    "C": "citigroup.com",
    "WFC": "wellsfargo.com",
    "GE": "ge.com",
    "BA": "boeing.com",
    "CAT": "caterpillar.com",
    "HON": "honeywell.com",
    "UPS": "ups.com",
    "FDX": "fedex.com",
    "XOM": "exxonmobil.com",
    "CVX": "chevron.com",
    "COP": "conocophillips.com",
    "NEE": "nexteraenergy.com",
    "NKE": "nike.com",
    "SBUX": "starbucks.com",
    "MCD": "mcdonalds.com",
}

EVENT_KEYWORDS = [
    "earnings", "conference call", "webcast", "quarter",
    "results", "investor", "event", "presentation",
]


def guess_domain(company_name: str, ticker: str) -> list[str]:
    """Guess possible domain names from company name and ticker."""
    domains = []

    # Check known mappings first
    if ticker in KNOWN_DOMAINS:
        domains.append(KNOWN_DOMAINS[ticker])

    # Generate guesses from company name
    name = company_name.lower()

    # Remove common suffixes
    for suffix in [" inc.", " inc", " corp.", " corp", " co.", " co",
                   " ltd.", " ltd", " plc", " sa", " ag", " nv",
                   " holdings", " group", " technologies", " technology",
                   " international", " enterprises", " solutions",
                   " class a", " class b", " class c",
                   ", inc.", ", inc", ", corp."]:
        name = name.replace(suffix, "")

    name = name.strip()

    # Simple domain: companyname.com
    simple = re.sub(r'[^a-z0-9]', '', name)
    if simple:
        domains.append(f"{simple}.com")

    # Hyphenated: company-name.com
    hyphenated = re.sub(r'[^a-z0-9\s]', '', name)
    hyphenated = re.sub(r'\s+', '-', hyphenated).strip('-')
    if hyphenated and hyphenated != simple:
        domains.append(f"{hyphenated}.com")

    # Ticker-based: ticker.com
    domains.append(f"{ticker.lower()}.com")

    return list(dict.fromkeys(domains))  # Deduplicate preserving order


def check_ir_url(url: str) -> dict:
    """Check if a URL is a working IR events page.

    Returns: {"url": str, "works": bool, "has_events": bool, "platform": str|None}
    """
    try:
        resp = httpx.get(url, headers=HEADERS, timeout=10, follow_redirects=True)
        if resp.status_code == 403:
            # Cloudflare or similar — URL might still be valid
            return {"url": url, "works": True, "has_events": False, "platform": "cloudflare_blocked"}
        if resp.status_code != 200:
            return {"url": url, "works": False, "has_events": False, "platform": None}

        text = resp.text.lower()

        # Reject error pages
        if "not found" in text[:500] or "404" in text[:500] or "page not found" in text[:1000]:
            return {"url": url, "works": False, "has_events": False, "platform": None}

        # Check for event-related content (need at least 2 keywords)
        has_events = sum(1 for kw in EVENT_KEYWORDS if kw in text) >= 2

        # Detect IR platform
        platform = None
        if "q4inc.com" in text or "q4web.com" in text or "q4cdn.com" in text:
            platform = "q4"
        elif "notified.com" in text:
            platform = "notified"
        elif "cision.com" in text:
            platform = "cision"
        elif "ir.nasdaq.com" in text:
            platform = "nasdaq_ir"

        return {"url": str(resp.url), "works": True, "has_events": has_events, "platform": platform}

    except Exception:
        return {"url": url, "works": False, "has_events": False, "platform": None}


def discover_ir_url(ticker: str, company_name: str) -> dict:
    """Try to discover the IR events page URL for a company."""
    domains = guess_domain(company_name, ticker)

    for domain in domains:
        for pattern in IR_URL_PATTERNS:
            url = pattern.format(domain=domain)
            result = check_ir_url(url)
            if result["works"] and result["has_events"]:
                return result

    # Try just the investor subdomain
    for domain in domains:
        for prefix in ["investor.", "investors.", "ir."]:
            url = f"https://{prefix}{domain}"
            result = check_ir_url(url)
            if result["works"]:
                return result

    return {"url": None, "works": False, "has_events": False, "platform": None}


def discover_all(tickers: list[str] = None, sp500_only: bool = False, limit: int = None, workers: int = 5):
    """Discover IR URLs for all companies (or a subset)."""
    init_db()
    db = SessionLocal()

    query = db.query(Company).filter(Company.ir_events_url.is_(None))
    if sp500_only:
        query = query.filter(Company.market_cap_tier == "sp500")
    if tickers:
        query = query.filter(Company.ticker.in_(tickers))

    companies = query.all()
    if limit:
        companies = companies[:limit]

    logger.info("Discovering IR URLs for %d companies...", len(companies))

    counts = {"found": 0, "blocked": 0, "not_found": 0}

    def process(company):
        return company.ticker, discover_ir_url(company.ticker, company.company_name)

    with ThreadPoolExecutor(max_workers=workers) as executor:
        futures = {executor.submit(process, c): c for c in companies}
        for i, future in enumerate(as_completed(futures)):
            ticker, result = future.result()
            company = futures[future]

            if result["works"] and result["has_events"]:
                company.ir_events_url = result["url"]
                company.ir_scraper_type = result["platform"]
                counts["found"] += 1
                logger.info("[%d/%d] %s: FOUND %s (platform: %s)",
                           i + 1, len(companies), ticker, result["url"], result["platform"])
            elif result.get("platform") == "cloudflare_blocked":
                company.ir_events_url = result["url"]
                company.ir_scraper_type = "cloudflare_blocked"
                counts["blocked"] += 1
                logger.info("[%d/%d] %s: CLOUDFLARE %s", i + 1, len(companies), ticker, result["url"])
            else:
                counts["not_found"] += 1
                if (i + 1) % 50 == 0:
                    logger.info("[%d/%d] Progress: %d found, %d blocked, %d not found",
                               i + 1, len(companies), counts["found"], counts["blocked"], counts["not_found"])

    db.commit()

    result = {"found": counts["found"], "blocked": counts["blocked"], "not_found": counts["not_found"], "total": len(companies)}
    logger.info("IR URL discovery complete: %s", result)
    db.close()
    return result


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Discover IR page URLs")
    parser.add_argument("--ticker", "-t", help="Specific ticker(s), comma-separated")
    parser.add_argument("--sp500-only", action="store_true", help="Only S&P 500 companies")
    parser.add_argument("--limit", "-l", type=int, help="Limit number of companies")
    parser.add_argument("--workers", "-w", type=int, default=5, help="Concurrent workers")
    args = parser.parse_args()

    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")

    tickers = [t.strip().upper() for t in args.ticker.split(",")] if args.ticker else None
    discover_all(tickers=tickers, sp500_only=args.sp500_only, limit=args.limit, workers=args.workers)
