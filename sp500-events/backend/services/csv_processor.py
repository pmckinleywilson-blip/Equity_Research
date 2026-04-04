"""CSV watchlist upload validation and processing."""
from security import validate_csv_upload


async def process_watchlist_csv(file_content: bytes) -> dict:
    """Validate and extract tickers from uploaded CSV.

    Returns:
        {"tickers": ["AAPL", "MSFT", ...], "errors": ["...", ...]}
    """
    return validate_csv_upload(file_content)
