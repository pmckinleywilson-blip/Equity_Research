"""LLM-based event extraction using Google Gemini.

Replaces regex parsing entirely. The LLM reads press release text
and returns structured event data. No regex fallback — if the LLM
can't extract it, we don't store incorrect data.

Uses the new google-genai SDK. Supports retry with backoff for rate limits.
"""
import json
import logging
import os
import time
from datetime import date
from typing import Optional

logger = logging.getLogger(__name__)

GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")
GEMINI_MODEL = os.environ.get("GEMINI_MODEL", "gemini-2.5-flash")

EXTRACTION_PROMPT = """Extract the CONFERENCE CALL or WEBCAST event from this press release. NOT the earnings release date — the CALL date.

Return JSON only: {"event_date":"YYYY-MM-DD","event_time":"HH:MM 24h ET" or null,"webcast_url":url or null,"phone_number":number or null,"phone_passcode":code or null,"event_type":"earnings"|"investor_day"|"conference"|"ad_hoc","fiscal_quarter":"Q1 2026" or null,"title":"short title"}

Rules:
- If release date differs from call date, use the CALL date
- Convert all times to Eastern Time (PT+3, CT+1)
- webcast_url must be a company URL, NOT prnewswire/businesswire/globenewswire
- If no identifiable event, return all nulls
- No markdown fences. No explanation. JSON only.

Text: """

MAX_RETRIES = 3
RETRY_DELAY = 5  # seconds


def _call_gemini(text: str) -> Optional[str]:
    """Call Gemini API with retry logic."""
    if not GEMINI_API_KEY:
        logger.warning("No Gemini API key — skipping LLM extraction")
        return None

    from google import genai

    client = genai.Client(api_key=GEMINI_API_KEY)

    for attempt in range(MAX_RETRIES):
        try:
            response = client.models.generate_content(
                model=GEMINI_MODEL,
                contents=EXTRACTION_PROMPT + text[:8000],
                config={
                    "temperature": 0,
                    "max_output_tokens": 500,
                    "thinking_config": {"thinking_budget": 0},
                },
            )
            return response.text.strip()

        except Exception as e:
            error_str = str(e)
            if "429" in error_str or "quota" in error_str.lower():
                wait = RETRY_DELAY * (attempt + 1)
                logger.warning("Rate limited (attempt %d/%d), waiting %ds...", attempt + 1, MAX_RETRIES, wait)
                time.sleep(wait)
            else:
                logger.error("Gemini API error: %s", e)
                return None

    logger.error("Gemini API failed after %d retries", MAX_RETRIES)
    return None


def extract_event_from_text(text: str, ticker: str = "") -> Optional[dict]:
    """Extract event details from press release text using Gemini.

    Returns a dict with event fields, or None if no event found.
    """
    response_text = _call_gemini(text)
    if not response_text:
        return None

    try:
        # Remove markdown code fences if present
        if "```" in response_text:
            # Extract content between code fences
            parts = response_text.split("```")
            for part in parts:
                part = part.strip()
                if part.startswith("json"):
                    part = part[4:].strip()
                if part.startswith("{"):
                    response_text = part
                    break

        # Fix common LLM JSON issues
        import re
        # Remove trailing commas before } or ]
        response_text = re.sub(r',\s*([}\]])', r'\1', response_text)
        # Fix "null" alternatives (NOT inside strings — simple approach)
        response_text = response_text.replace(': None', ': null').replace(': True', ': true').replace(': False', ': false')

        data = json.loads(response_text)

        if data.get("no_event"):
            return None

        result = {"ticker": ticker}

        # Date (required)
        if data.get("event_date"):
            try:
                parts = data["event_date"].split("-")
                result["event_date"] = date(int(parts[0]), int(parts[1]), int(parts[2]))
            except (ValueError, IndexError):
                return None
        else:
            return None

        # Time — handle "HH:MM", "HH:MM ET", "HH:MM:SS" formats
        if data.get("event_time"):
            try:
                time_str = data["event_time"].replace(" ET", "").replace(" EST", "").replace(" EDT", "").strip()
                parts = time_str.split(":")
                hour, minute = int(parts[0]), int(parts[1].split()[0])
                if 6 <= hour <= 22:
                    result["event_time"] = f"{hour:02d}:{minute:02d}:00"
            except (ValueError, IndexError):
                pass

        # Webcast URL
        if data.get("webcast_url"):
            url = data["webcast_url"]
            wire_domains = ["prnewswire.com", "businesswire.com", "globenewswire.com"]
            if url.startswith("http") and not any(d in url for d in wire_domains):
                result["webcast_url"] = url

        # Phone
        if data.get("phone_number"):
            result["phone_number"] = data["phone_number"]
        if data.get("phone_passcode"):
            result["phone_passcode"] = str(data["phone_passcode"])

        # Event type
        valid_types = {"earnings", "investor_day", "conference", "ad_hoc"}
        result["event_type"] = data.get("event_type", "earnings")
        if result["event_type"] not in valid_types:
            result["event_type"] = "earnings"

        # Fiscal quarter and title
        if data.get("fiscal_quarter"):
            result["fiscal_quarter"] = data["fiscal_quarter"]
        if data.get("title"):
            result["title"] = data["title"][:300]

        return result

    except json.JSONDecodeError as e:
        logger.warning("LLM returned invalid JSON for %s: %s", ticker, e)
        return None


def extract_event_from_pr(pr_html: str, ticker: str, source_url: str, source: str) -> Optional[dict]:
    """Extract event from a press release HTML page.

    Handles HTML to text conversion and passes to LLM.
    """
    from bs4 import BeautifulSoup

    soup = BeautifulSoup(pr_html, "lxml")

    article = (
        soup.select_one(".bw-release-story")
        or soup.select_one(".bw-release-body")
        or soup.select_one(".release-body")
        or soup.select_one(".prnewswire-body")
        or soup.select_one(".main-body-container")
        or soup.select_one("#press-release-body")
        or soup.select_one("article")
        or soup.select_one(".entry-content")
        or soup.select_one("main")
    )

    if article:
        text = article.get_text(separator="\n")
    else:
        body = soup.select_one("body")
        text = body.get_text(separator="\n") if body else ""

    if len(text.strip()) < 50:
        return None

    result = extract_event_from_text(text, ticker)
    if result:
        result["source"] = source
        result["source_url"] = source_url

    return result
