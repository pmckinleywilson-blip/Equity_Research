"""Russell 3000 + S&P 500 constituent list management.

Fetches and maintains the list of companies to track.
Uses free sources: Wikipedia (S&P 500) and SEC EDGAR (full list).
"""
import json
import logging
import os
import sys
import httpx
from pathlib import Path

# Add backend to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent / "backend"))

from database import SessionLocal, init_db
from models import Company

logger = logging.getLogger(__name__)

SHARED_DIR = Path(__file__).parent.parent / "shared"
SP500_WIKI_URL = "https://en.wikipedia.org/wiki/List_of_S%26P_500_companies"


def fetch_sp500_from_wikipedia() -> list[dict]:
    """Fetch S&P 500 constituents from Wikipedia."""
    import re
    from bs4 import BeautifulSoup

    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Accept": "text/html,application/xhtml+xml",
    }
    resp = httpx.get(SP500_WIKI_URL, headers=headers, timeout=30, follow_redirects=True)
    resp.raise_for_status()

    soup = BeautifulSoup(resp.text, "lxml")
    table = soup.find("table", {"id": "constituents"})
    if not table:
        raise ValueError("Could not find S&P 500 constituents table on Wikipedia")

    companies = []
    rows = table.find_all("tr")[1:]  # Skip header
    for row in rows:
        cells = row.find_all("td")
        if len(cells) >= 8:
            ticker = cells[0].get_text(strip=True).replace(".", "-")  # BRK.B → BRK-B
            name = cells[1].get_text(strip=True)
            sector = cells[2].get_text(strip=True)
            sub_industry = cells[3].get_text(strip=True)
            cik = cells[6].get_text(strip=True)

            companies.append({
                "ticker": ticker,
                "company_name": name,
                "sector": sector,
                "sub_industry": sub_industry,
                "cik": cik.zfill(10) if cik.isdigit() else cik,
                "market_cap_tier": "sp500",
            })

    logger.info("Fetched %d S&P 500 companies from Wikipedia", len(companies))
    return companies


def fetch_russell3000_from_nasdaq() -> list[dict]:
    """Fetch additional Russell 3000 companies from Nasdaq screener.

    This fills in mid/small cap companies not in the S&P 500.
    """
    url = "https://api.nasdaq.com/api/screener/stocks"
    params = {
        "tableonly": "true",
        "limit": 5000,
        "exchange": "NYSE,NASDAQ,AMEX",
    }
    headers = {
        "User-Agent": "Mozilla/5.0 (compatible; SP500Events/1.0)",
        "Accept": "application/json",
    }

    resp = httpx.get(url, params=params, headers=headers, timeout=30)
    resp.raise_for_status()
    data = resp.json()

    companies = []
    rows = data.get("data", {}).get("table", {}).get("rows", []) or data.get("data", {}).get("rows", [])
    for row in rows:
        ticker = row.get("symbol", "").strip()
        name = row.get("name", "").strip()
        sector = row.get("sector", "").strip()
        industry = row.get("industry", "").strip()
        market_cap = row.get("marketCap", "")

        if not ticker or not name:
            continue

        companies.append({
            "ticker": ticker,
            "company_name": name,
            "sector": sector if sector else None,
            "sub_industry": industry if industry else None,
            "market_cap_tier": "mid",  # Will be overridden for sp500
        })

    logger.info("Fetched %d companies from Nasdaq screener", len(companies))
    return companies


def classify_market_cap_tier(company: dict, sp500_tickers: set) -> str:
    """Classify a company into sp500/mid/small tier."""
    if company["ticker"] in sp500_tickers:
        return "sp500"
    # Simple heuristic — could be improved with actual market cap data
    return "mid"


def populate_companies(db_session=None):
    """Populate the companies table with Russell 3000 + S&P 500 data."""
    close_session = False
    if db_session is None:
        init_db()
        db_session = SessionLocal()
        close_session = True

    try:
        # Fetch S&P 500
        sp500 = fetch_sp500_from_wikipedia()
        sp500_tickers = {c["ticker"] for c in sp500}

        # Fetch broader list
        try:
            nasdaq_companies = fetch_russell3000_from_nasdaq()
        except Exception as e:
            logger.warning("Failed to fetch Nasdaq screener data: %s. Using S&P 500 only.", e)
            nasdaq_companies = []

        # Merge: S&P 500 companies take priority
        all_companies = {c["ticker"]: c for c in nasdaq_companies}
        for c in sp500:
            all_companies[c["ticker"]] = c

        # Upsert into database
        count = 0
        for ticker, data in all_companies.items():
            existing = db_session.query(Company).filter(Company.ticker == ticker).first()
            if existing:
                existing.company_name = data["company_name"]
                existing.sector = data.get("sector")
                existing.sub_industry = data.get("sub_industry")
                existing.market_cap_tier = data.get("market_cap_tier", "mid")
                if data.get("cik"):
                    existing.cik = data["cik"]
            else:
                company = Company(
                    ticker=ticker,
                    company_name=data["company_name"],
                    sector=data.get("sector"),
                    sub_industry=data.get("sub_industry"),
                    market_cap_tier=data.get("market_cap_tier", "mid"),
                    cik=data.get("cik"),
                )
                db_session.add(company)
            count += 1

        db_session.commit()
        logger.info("Populated %d companies (%d S&P 500)", count, len(sp500_tickers))

        # Also save to shared JSON for reference
        SHARED_DIR.mkdir(parents=True, exist_ok=True)
        with open(SHARED_DIR / "russell3000.json", "w") as f:
            json.dump(
                {
                    "updated": __import__("datetime").datetime.utcnow().isoformat(),
                    "sp500_count": len(sp500_tickers),
                    "total_count": len(all_companies),
                    "companies": [
                        {
                            "ticker": t,
                            "name": d["company_name"],
                            "sector": d.get("sector"),
                            "tier": d.get("market_cap_tier", "mid"),
                        }
                        for t, d in sorted(all_companies.items())
                    ],
                },
                f,
                indent=2,
            )

        return {"sp500": len(sp500_tickers), "total": len(all_companies)}

    finally:
        if close_session:
            db_session.close()


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    result = populate_companies()
    print(f"Done: {result}")
