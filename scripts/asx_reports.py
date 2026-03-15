"""
ASX Earnings Report Downloader

Downloads earnings reports for ASX-listed companies from asx.com.au.
Uses an LLM to semantically identify results days from announcement titles,
then downloads all announcements from those dates. Also captures Annual Reports
released within 60 days after an annual results day.

Usage: python scripts/asx_reports.py <TICKER> <PERIOD>
  PERIOD: number of years (e.g. 3) or "last" for most recent results day only

Requires: pip install requests anthropic
"""

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import time
from collections import defaultdict
from datetime import datetime, timedelta
from pathlib import Path

try:
    import requests
except ImportError:
    print("Error: 'requests' package is required. Install with: pip install requests")
    sys.exit(1)

# ASX v2 HTML API for announcement listings
LIST_URL = "https://www.asx.com.au/asx/v2/statistics/announcements.do"
# PDF display endpoint (requires terms acceptance)
PDF_DISPLAY_URL = "https://www.asx.com.au/asx/v2/statistics/displayAnnouncement.do"
# Terms acceptance endpoint
TERMS_URL = "https://www.asx.com.au/asx/v2/statistics/announcementTerms.do"

REQUEST_DELAY = 1.5

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
}


def create_session():
    """Create a requests session with browser-like headers."""
    session = requests.Session()
    session.headers.update(HEADERS)
    return session


def parse_announcements_html(html):
    """Parse announcement data from the v2 HTML response."""
    announcements = []
    rows = re.split(r'<tr[\s>]', html)
    for row in rows:
        date_match = re.search(r'(\d{2}/\d{2}/\d{4})', row)
        ids_match = re.search(r'idsId=(\d+)', row)
        if not (date_match and ids_match):
            continue

        title_match = re.search(r'idsId=\d+[^>]*>\s*(.*?)\s*<br', row, re.DOTALL)
        if not title_match:
            title_match = re.search(r'idsId=\d+[^>]*>\s*(.*?)\s*</a>', row, re.DOTALL)
        if not title_match:
            continue

        date_str = date_match.group(1)
        try:
            release_date = datetime.strptime(date_str, "%d/%m/%Y").date()
        except ValueError:
            continue

        headline = title_match.group(1).strip()
        headline = re.sub(r'<[^>]+>', '', headline).strip()
        headline = headline.replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">")

        announcements.append({
            "date": release_date,
            "ids_id": ids_match.group(1),
            "headline": headline,
        })

    return announcements


def fetch_announcements(session, ticker, years_back):
    """Fetch announcements from the ASX v2 HTML API for the specified period."""
    all_announcements = []
    current_year = datetime.now().year

    if years_back == 0:
        year_range = [current_year, current_year - 1]
    else:
        year_range = list(range(current_year, current_year - years_back - 1, -1))

    for year in year_range:
        time.sleep(REQUEST_DELAY)
        try:
            resp = session.get(LIST_URL, params={
                "by": "asxCode",
                "asxCode": ticker,
                "timeframe": "Y",
                "year": year,
            }, timeout=30)
        except requests.RequestException as e:
            print(f"Error fetching announcements for {year}: {e}")
            continue

        if resp.status_code == 404:
            print(f"Error: Ticker '{ticker}' not found on ASX. Check the ticker symbol.")
            sys.exit(1)
        if resp.status_code == 403:
            print(
                f"Error: Request blocked (403 Forbidden). Try again in a few minutes,\n"
                f"or visit https://www.asx.com.au/markets/company/{ticker} manually."
            )
            sys.exit(1)
        if resp.status_code != 200:
            print(f"Warning: API returned status {resp.status_code} for year {year}")
            continue

        announcements = parse_announcements_html(resp.text)
        if not announcements:
            if "No announcements found" in resp.text or len(resp.text) < 5000:
                if year == current_year and not all_announcements:
                    resp2 = session.get(LIST_URL, params={
                        "by": "asxCode",
                        "asxCode": ticker,
                        "timeframe": "D",
                        "period": "M6",
                    }, timeout=30)
                    if "No announcements" in resp2.text and len(re.findall(r'idsId=', resp2.text)) == 0:
                        print(f"Error: No announcements found for '{ticker}'. Check the ticker symbol.")
                        sys.exit(1)

        all_announcements.extend(announcements)

    return all_announcements


# ---------------------------------------------------------------------------
# LLM-based classification
# ---------------------------------------------------------------------------

def _find_claude_cli():
    """Find the claude CLI executable."""
    # Try common locations
    cli = shutil.which("claude")
    if cli:
        return cli
    # Windows npm global install
    npm_cmd = Path(os.environ.get("APPDATA", "")) / "npm" / "claude.cmd"
    if npm_cmd.exists():
        return str(npm_cmd)
    return None


def _call_claude(prompt):
    """Call the Claude CLI in pipe mode and return the text response."""
    cli = _find_claude_cli()
    if not cli:
        print("Error: 'claude' CLI not found. Install with: npm install -g @anthropic-ai/claude-code")
        sys.exit(1)

    result = subprocess.run(
        [cli, "-p", "--output-format", "json", "--model", "haiku",
         "--allowedTools", ""],
        input=prompt,
        capture_output=True, text=True, shell=(os.name == "nt"), timeout=120,
    )
    if result.returncode != 0:
        print(f"Error calling claude CLI: {result.stderr.strip()[:200]}")
        sys.exit(1)

    try:
        envelope = json.loads(result.stdout)
        return envelope.get("result", "")
    except json.JSONDecodeError:
        return result.stdout.strip()


def classify_announcements_llm(announcements):
    """Use Claude to identify results days and annual reports from announcement titles.

    Returns:
        results_days: dict {date: set of types} where type is "annual", "interim", or "quarterly"
        annual_report_ids: set of ids_id strings for announcements that are full Annual Reports
            released after the main results day
        preliminary_ids: set of ids_id strings for preliminary final reports
    """
    # Group announcements by date, deduplicating titles
    by_date = defaultdict(list)
    for ann in announcements:
        h = ann["headline"]
        # Skip routine noise that can never be results-related
        h_lower = h.lower()
        if any(skip in h_lower for skip in [
            "substantial hold", "cessation of securities",
            "regarding unquoted securities", "notification of buy-back",
            "update - notification", "director's interest",
            "change of director", "statement of cdis",
        ]):
            continue
        # SEC Form 4 / Forms 4 are insider trading disclosures — pure noise.
        # But keep 8-K, 10-Q, 10-K, S-3, S-8 etc.
        if re.match(r"^sec\s+forms?\s+4$", h_lower):
            continue
        # "Results Announcement Date" / "Webcast Details" notices announce a future date, not actual results
        if re.search(r"results\s+announcement\s+date", h_lower):
            continue
        if re.search(r"\bwebcast\s+details\b", h_lower):
            continue
        by_date[ann["date"]].append(h)

    # Deduplicate identical titles on the same date
    for d in by_date:
        by_date[d] = list(dict.fromkeys(by_date[d]))

    # Build a compact representation for the LLM
    date_lines = []
    for d in sorted(by_date.keys(), reverse=True):
        titles = by_date[d]
        if not titles:
            continue
        date_str = d.strftime("%Y-%m-%d")
        title_list = "; ".join(titles)
        date_lines.append(f"{date_str}: {title_list}")

    prompt_body = "\n".join(date_lines)

    prompt = f"""You are analysing ASX (Australian Securities Exchange) announcements for an equity research workflow.

Below is a list of announcement dates and titles for a company. Each line is one date followed by all announcement titles released that day.

Your task:
1. Identify which dates are "results days" — days when the company published its periodic financial results (earnings). A results day is any date where a statutory results filing was released. This includes:
   - ASX filings: Appendix 4E (annual), 4D (half-year), 4C (quarterly), Preliminary Final Report
   - SEC filings: 10-K (annual), 10-Q (quarterly) — a 10-Q or 10-K filing alone IS a results day even without a presentation
   - Results announcements, media releases, or presentations released on the same day
   - Quarterly Activities Reports
   Some companies (especially dual-listed US/ASX companies) may only file a 10-Q with no accompanying presentation — this still counts as a results day.
   Ignore: AGM results, conference call detail notices released before actual results, "results announcement date" or "webcast details" notices (these announce a future date, they are NOT results), trading updates without full results, and routine admin filings (Form 4, substantial holder notices, buy-backs).

2. Classify each results day as one of:
   - "annual" — full-year / annual results
   - "interim" — half-year / interim results
   - "quarterly" — quarterly results or activities report

3. Identify any "annual report" announcements — full Annual Reports (not the Appendix 4E/preliminary final report itself, but the standalone Annual Report document) that may be released days or weeks after the annual results day. These are often released on the same day as results or shortly after.

4. Identify any "preliminary final report" announcements — these are the preliminary versions that should be replaced if a full Annual Report exists for the same period.

Respond with ONLY valid JSON in this exact format, no other text:
{{
  "results_days": [
    {{"date": "YYYY-MM-DD", "type": "annual|interim|quarterly"}}
  ],
  "annual_report_ids": ["YYYY-MM-DD: Exact title of annual report announcement"],
  "preliminary_ids": ["YYYY-MM-DD: Exact title of preliminary final report"]
}}

If there are no results days, return empty arrays.

Announcements:
{prompt_body}"""

    text = _call_claude(prompt)

    # Strip markdown code fences if present
    if text.startswith("```"):
        text = re.sub(r'^```(?:json)?\s*', '', text)
        text = re.sub(r'\s*```$', '', text)

    try:
        data = json.loads(text)
    except json.JSONDecodeError:
        print(f"Warning: Could not parse LLM response. Falling back to empty results.")
        print(f"  Response: {text[:200]}")
        return {}, set(), set()

    # Convert to internal format
    results_days = {}
    for rd in data.get("results_days", []):
        try:
            d = datetime.strptime(rd["date"], "%Y-%m-%d").date()
            rd_type = rd["type"]
            if rd_type in ("annual", "interim", "quarterly"):
                if d not in results_days:
                    results_days[d] = set()
                results_days[d].add(rd_type)
        except (ValueError, KeyError):
            continue

    # Build sets of announcement identifiers for annual reports and preliminaries
    # The LLM returns "YYYY-MM-DD: Title" — match back to actual announcements
    annual_report_ids = _match_llm_refs(data.get("annual_report_ids", []), announcements)
    preliminary_ids = _match_llm_refs(data.get("preliminary_ids", []), announcements)

    return results_days, annual_report_ids, preliminary_ids


def _match_llm_refs(ref_strings, announcements):
    """Match LLM-returned 'YYYY-MM-DD: Title' strings back to announcement ids_ids."""
    matched = set()
    for ref in ref_strings:
        parts = ref.split(": ", 1)
        if len(parts) != 2:
            continue
        date_str, title = parts
        try:
            ref_date = datetime.strptime(date_str, "%Y-%m-%d").date()
        except ValueError:
            continue
        # Find best match — exact substring match on same date
        for ann in announcements:
            if ann["date"] == ref_date and title.lower() in ann["headline"].lower():
                matched.add(ann["ids_id"])
                break
        else:
            # Fallback: try partial match (LLM may have truncated or rephrased slightly)
            for ann in announcements:
                if ann["date"] == ref_date and (
                    ann["headline"].lower() in title.lower()
                    or title.lower()[:30] in ann["headline"].lower()
                ):
                    matched.add(ann["ids_id"])
                    break
    return matched


def find_annual_reports_near(announcements, results_days, annual_report_ids):
    """Find Annual Report announcements within 60 days after each annual results day.

    Uses the LLM-identified annual_report_ids, but also checks proximity to results days.
    Returns list of announcements with _associated_annual_date set.
    """
    annual_dates = [d for d, types in results_days.items() if "annual" in types]
    if not annual_dates:
        return []

    found = []
    for ann in announcements:
        if ann["ids_id"] not in annual_report_ids:
            continue
        # Check if it's within 60 days of an annual results day (or on the day itself)
        for ad in annual_dates:
            if ad <= ann["date"] <= ad + timedelta(days=60):
                ann_copy = dict(ann)
                ann_copy["_associated_annual_date"] = ad
                found.append(ann_copy)
                break

    return found


def accept_terms(session, ids_id):
    """Accept ASX terms by hitting a PDF endpoint and posting the terms form."""
    try:
        resp = session.get(PDF_DISPLAY_URL, params={
            "display": "pdf",
            "idsId": ids_id,
        }, timeout=30)

        if resp.content[:4] == b"%PDF":
            return True

        match = re.search(r'name="pdfURL"\s+value="([^"]+)"', resp.text)
        if match:
            pdf_url = match.group(1)
            resp2 = session.post(TERMS_URL, data={"pdfURL": pdf_url}, timeout=30)
            return resp2.status_code == 200
    except requests.RequestException as e:
        print(f"Warning: Could not accept terms: {e}")
    return False


def sanitize_filename(name):
    """Remove or replace characters that are invalid in filenames."""
    name = re.sub(r'[<>:"/\\|?*]', "", name)
    name = re.sub(r"\s+", " ", name).strip()
    if len(name) > 150:
        name = name[:150].strip()
    return name


def build_filename(ticker, announcement):
    """Build filename: {TICKER} {YYYY-MM-DD} {Sanitized Title}.pdf"""
    date_str = announcement["date"].strftime("%Y-%m-%d")
    title = sanitize_filename(announcement["headline"])
    return f"{ticker} {date_str} {title}.pdf"


def download_pdf(session, ids_id, dest_path, max_retries=3):
    """Download a PDF with retry and exponential backoff."""
    for attempt in range(max_retries):
        try:
            time.sleep(REQUEST_DELAY)
            resp = session.get(PDF_DISPLAY_URL, params={
                "display": "pdf",
                "idsId": ids_id,
            }, timeout=60)

            if resp.status_code == 200 and resp.content[:4] == b"%PDF":
                with open(dest_path, "wb") as f:
                    f.write(resp.content)
                return True

            if resp.status_code == 200:
                match = re.search(r'name="pdfURL"\s+value="([^"]+)"', resp.text)
                if match:
                    direct_url = match.group(1)
                    session.post(TERMS_URL, data={"pdfURL": direct_url}, timeout=30)
                    time.sleep(REQUEST_DELAY)
                    resp2 = session.get(direct_url, timeout=60)
                    if resp2.status_code == 200 and resp2.content[:4] == b"%PDF":
                        with open(dest_path, "wb") as f:
                            f.write(resp2.content)
                        return True

            if resp.status_code == 403:
                print(f"  Blocked (403) downloading {os.path.basename(dest_path)}")
                return False

            print(f"  HTTP {resp.status_code} for {os.path.basename(dest_path)} (attempt {attempt + 1})")

        except requests.RequestException as e:
            print(f"  Error downloading {os.path.basename(dest_path)}: {e} (attempt {attempt + 1})")

        if attempt < max_retries - 1:
            wait = 2 ** (attempt + 1)
            time.sleep(wait)

    return False


def filter_by_period(results_days, period, reference_date=None):
    """Filter results days to those within the requested period.

    For numeric periods (e.g. 1, 3, 5 years), the cutoff is aligned to the
    results calendar: find the latest results day, then go back N years from
    the 1st of that month.  E.g. if the latest results day is 2025-08-15 and
    period is 1, the cutoff is 2024-08-01 — capturing the full prior year of
    results starting from the same reporting month.
    """
    if period == "last":
        if not results_days:
            return {}
        latest = max(results_days.keys())
        return {latest: results_days[latest]}

    years = int(period)

    if results_days:
        latest = max(results_days.keys())
        cutoff_year = latest.year - years
        cutoff = latest.replace(year=cutoff_year, day=1)
    else:
        if reference_date is None:
            reference_date = datetime.now().date()
        cutoff = reference_date - timedelta(days=years * 365)

    return {d: t for d, t in results_days.items() if d >= cutoff}


def deduplicate_reports(reports):
    """Remove duplicate announcements (same ids_id)."""
    seen = set()
    unique = []
    for r in reports:
        if r["ids_id"] not in seen:
            seen.add(r["ids_id"])
            unique.append(r)
    return unique


def main():
    parser = argparse.ArgumentParser(description="Download ASX earnings reports")
    parser.add_argument("ticker", help="ASX ticker symbol (e.g. 360, CBA, BHP)")
    parser.add_argument(
        "period",
        help="Number of years to look back (e.g. 3) or 'last' for most recent results day",
    )
    args = parser.parse_args()

    ticker = args.ticker.upper()
    period = args.period.lower()

    if period != "last":
        try:
            years = int(period)
            if years < 1:
                raise ValueError
        except ValueError:
            print(f"Error: Invalid period '{period}'. Use a positive number of years or 'last'.")
            sys.exit(1)

    repo_root = Path(__file__).resolve().parent.parent
    company_dir = repo_root / ticker / "Company reports"

    if not (repo_root / ticker).exists():
        print(
            f"Error: Company folder '{ticker}/' does not exist.\n"
            f"Create it first by copying 'Template folder dir/' to '{ticker}/'."
        )
        sys.exit(1)

    company_dir.mkdir(parents=True, exist_ok=True)

    session = create_session()
    years_back = 0 if period == "last" else int(period)

    print(f"Fetching announcements for {ticker}...")
    announcements = fetch_announcements(session, ticker, years_back)
    print(f"Found {len(announcements)} total announcements.")

    if not announcements:
        print("No announcements found.")
        sys.exit(0)

    # Use LLM to classify announcements
    print("Classifying announcements...")
    all_results_days, annual_report_ids, preliminary_ids = classify_announcements_llm(announcements)

    results_days = filter_by_period(all_results_days, period)

    if not results_days:
        print(f"No results days found in the requested period ({period}).")
        sys.exit(0)

    print(f"Identified {len(results_days)} results days in the requested period.")
    for d in sorted(results_days.keys(), reverse=True):
        types = ", ".join(sorted(results_days[d]))
        print(f"  {d.strftime('%Y-%m-%d')} ({types})")

    # Collect all reports from results days
    results_dates = set(results_days.keys())
    reports = [ann for ann in announcements if ann["date"] in results_dates]

    # Find Annual Reports released after annual results days
    annual_reports = find_annual_reports_near(announcements, results_days, annual_report_ids)
    if annual_reports:
        print(f"Found {len(annual_reports)} Annual Report(s) released after results day.")

    # Remove preliminary final reports if a full Annual Report exists for that period
    if annual_reports:
        annual_dates_with_full = {
            ar["_associated_annual_date"] for ar in annual_reports
        }
        removed = 0
        filtered = []
        for r in reports:
            if r["ids_id"] in preliminary_ids and r["date"] in annual_dates_with_full:
                removed += 1
                continue
            filtered.append(r)
        reports = filtered
        if removed:
            print(f"Replaced {removed} preliminary report(s) with full Annual Report(s).")

    # Merge in annual reports
    existing_ids = {r["ids_id"] for r in reports}
    for ar in annual_reports:
        if ar["ids_id"] not in existing_ids:
            reports.append(ar)

    reports = deduplicate_reports(reports)
    reports.sort(key=lambda r: r["date"], reverse=True)

    print(f"{len(reports)} reports to download.\n")

    if not reports:
        print("No reports to download.")
        sys.exit(0)

    # Accept terms using the first announcement's ID
    print("Accepting ASX terms of use...")
    accept_terms(session, reports[0]["ids_id"])

    # Download
    downloaded = 0
    skipped = 0
    failed = 0

    for report in reports:
        filename = build_filename(ticker, report)
        dest = company_dir / filename

        if dest.exists():
            print(f"  Skipped (exists): {filename}")
            skipped += 1
            continue

        print(f"  Downloading: {filename}")
        if download_pdf(session, report["ids_id"], str(dest)):
            downloaded += 1
        else:
            if dest.exists():
                dest.unlink()
            failed += 1

    print(f"\nDone. Downloaded: {downloaded}, Skipped: {skipped}, Failed: {failed}")


if __name__ == "__main__":
    main()
