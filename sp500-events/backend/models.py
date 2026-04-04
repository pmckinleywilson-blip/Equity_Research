from datetime import datetime, date, time
from sqlalchemy import (
    Column, Integer, String, Date, Time, Text, Boolean, DateTime,
    ForeignKey, UniqueConstraint, Index, JSON
)
from sqlalchemy.orm import relationship

from database import Base


class Event(Base):
    __tablename__ = "events"

    id = Column(Integer, primary_key=True, index=True)
    ticker = Column(String(10), nullable=False, index=True)
    company_name = Column(String(200), nullable=False)
    event_type = Column(String(50), nullable=False)  # earnings, investor_day, conference, ad_hoc
    event_date = Column(Date, nullable=False, index=True)
    event_time = Column(Time, nullable=True)
    timezone = Column(String(50), default="America/New_York")
    title = Column(String(500))
    description = Column(Text)
    webcast_url = Column(String(1000))
    phone_number = Column(String(100))
    phone_passcode = Column(String(50))
    replay_url = Column(String(1000))
    fiscal_quarter = Column(String(10))
    source = Column(String(50), nullable=False)
    source_url = Column(String(1000))
    ir_verified = Column(Boolean, default=False, index=True)
    status = Column(String(20), default="tentative")  # confirmed, tentative, postponed, cancelled
    notified_at = Column(DateTime, nullable=True)
    created_at = Column(DateTime, default=datetime.utcnow)
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    __table_args__ = (
        UniqueConstraint("ticker", "event_date", "event_type", name="uq_ticker_date_type"),
    )


class Company(Base):
    __tablename__ = "companies"

    ticker = Column(String(10), primary_key=True)
    company_name = Column(String(200), nullable=False)
    sector = Column(String(100))
    sub_industry = Column(String(200))
    market_cap_tier = Column(String(10))  # sp500, mid, small
    cik = Column(String(20))
    ir_events_url = Column(String(500))
    ir_scraper_type = Column(String(50))  # q4, notified, cision, custom, NULL
    last_scraped = Column(DateTime, nullable=True)


class WatchlistSubscription(Base):
    __tablename__ = "watchlist_subscriptions"

    id = Column(Integer, primary_key=True, index=True)
    email = Column(String(320), nullable=False, unique=True)
    tickers = Column(JSON, nullable=False)  # ["AAPL", "MSFT", ...]
    calendar_type = Column(String(10), default="outlook")  # outlook, gmail
    feed_token = Column(String(64), unique=True, index=True)
    is_active = Column(Boolean, default=True, index=True)
    created_at = Column(DateTime, default=datetime.utcnow)

    notifications = relationship("NotificationLog", back_populates="subscription")


class NotificationLog(Base):
    __tablename__ = "notification_log"

    id = Column(Integer, primary_key=True, index=True)
    subscription_id = Column(Integer, ForeignKey("watchlist_subscriptions.id"), nullable=False)
    event_id = Column(Integer, ForeignKey("events.id"), nullable=False)
    sent_at = Column(DateTime, default=datetime.utcnow)

    subscription = relationship("WatchlistSubscription", back_populates="notifications")
    event = relationship("Event")

    __table_args__ = (
        UniqueConstraint("subscription_id", "event_id", name="uq_sub_event"),
    )
