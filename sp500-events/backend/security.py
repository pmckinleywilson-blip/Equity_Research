import csv
import io
import re
import secrets
from datetime import datetime, timedelta
from typing import Optional

from jose import jwt, JWTError
from config import get_settings

settings = get_settings()

FORMULA_PREFIXES = ("=", "+", "-", "@", "\t", "\r", "\n")


def generate_feed_token() -> str:
    """Generate a cryptographically random 64-char feed token."""
    return secrets.token_urlsafe(48)  # 48 bytes → 64 chars base64


def generate_unsubscribe_token(email: str) -> str:
    """Generate a signed JWT for one-click unsubscribe."""
    payload = {
        "email": email,
        "action": "unsubscribe",
        "exp": datetime.utcnow() + timedelta(days=365),
    }
    return jwt.encode(payload, settings.secret_key, algorithm="HS256")


def verify_unsubscribe_token(token: str) -> Optional[str]:
    """Verify unsubscribe JWT, return email if valid."""
    try:
        payload = jwt.decode(token, settings.secret_key, algorithms=["HS256"])
        if payload.get("action") != "unsubscribe":
            return None
        return payload.get("email")
    except JWTError:
        return None


def validate_csv_upload(file_content: bytes, max_size_bytes: int = 1_048_576) -> dict:
    """Validate and parse a CSV watchlist upload.

    Returns: {"tickers": [...], "errors": [...]}
    """
    if len(file_content) > max_size_bytes:
        return {"tickers": [], "errors": ["File exceeds 1MB limit"]}

    try:
        text = file_content.decode("utf-8-sig")  # Handle BOM
    except UnicodeDecodeError:
        return {"tickers": [], "errors": ["File must be UTF-8 encoded"]}

    # Check for formula injection in raw content
    for line_num, line in enumerate(text.splitlines(), 1):
        for cell in line.split(","):
            cell = cell.strip().strip('"').strip("'")
            if cell and cell[0] in FORMULA_PREFIXES:
                return {
                    "tickers": [],
                    "errors": [f"Potentially dangerous content on line {line_num}: cell starts with '{cell[0]}'"],
                }

    reader = csv.DictReader(io.StringIO(text))
    if not reader.fieldnames:
        return {"tickers": [], "errors": ["CSV has no headers"]}

    # Find the ticker column (case-insensitive)
    ticker_col = None
    for col in reader.fieldnames:
        if col.strip().lower() in ("ticker", "symbol", "tickers", "stock"):
            ticker_col = col
            break

    if ticker_col is None:
        return {
            "tickers": [],
            "errors": [f"No 'ticker' column found. Headers: {reader.fieldnames}"],
        }

    tickers = []
    row_count = 0
    for row in reader:
        row_count += 1
        if row_count > 3000:
            return {"tickers": [], "errors": ["CSV exceeds 3000 row limit"]}
        raw = row.get(ticker_col, "").strip()
        cleaned = re.sub(r"[^A-Za-z0-9.]", "", raw).upper()
        if cleaned and len(cleaned) <= 10:
            tickers.append(cleaned)

    return {"tickers": list(set(tickers)), "errors": []}
