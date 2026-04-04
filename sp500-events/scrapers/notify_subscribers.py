"""Standalone notification script for GitHub Actions.

Run after scraping to send calendar invites for newly confirmed events.
"""
import asyncio
import logging
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent / "backend"))

from database import SessionLocal, init_db
from services.notifier import notify_new_events

logger = logging.getLogger(__name__)


async def main():
    init_db()
    db = SessionLocal()
    try:
        result = await notify_new_events(db)
        logger.info("Notification run complete: %s", result)
        return result
    finally:
        db.close()


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(name)s %(levelname)s %(message)s")
    asyncio.run(main())
