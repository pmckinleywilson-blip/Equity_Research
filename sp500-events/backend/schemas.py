from datetime import datetime, date, time
from typing import Optional
from pydantic import BaseModel, EmailStr, Field, field_validator
import re


# --- Event schemas ---

class EventBase(BaseModel):
    ticker: str
    company_name: str
    event_type: str
    event_date: date
    event_time: Optional[time] = None
    timezone: str = "America/New_York"
    title: Optional[str] = None
    description: Optional[str] = None
    webcast_url: Optional[str] = None
    phone_number: Optional[str] = None
    phone_passcode: Optional[str] = None
    replay_url: Optional[str] = None
    fiscal_quarter: Optional[str] = None


class EventResponse(EventBase):
    id: int
    source: str
    source_url: Optional[str] = None
    ir_verified: bool
    status: str
    created_at: datetime
    updated_at: datetime

    class Config:
        from_attributes = True


class EventListResponse(BaseModel):
    events: list[EventResponse]
    total: int
    page: int
    per_page: int
    pages: int


# --- Company schemas ---

class CompanyResponse(BaseModel):
    ticker: str
    company_name: str
    sector: Optional[str] = None
    sub_industry: Optional[str] = None
    market_cap_tier: Optional[str] = None

    class Config:
        from_attributes = True


class CompanyListResponse(BaseModel):
    companies: list[CompanyResponse]
    total: int


# --- Watchlist / Subscription schemas ---

class SubscribeRequest(BaseModel):
    email: str
    tickers: list[str]
    calendar_type: str = "outlook"

    @field_validator("email")
    @classmethod
    def validate_email(cls, v: str) -> str:
        pattern = r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
        if not re.match(pattern, v):
            raise ValueError("Invalid email format")
        return v.lower().strip()

    @field_validator("tickers")
    @classmethod
    def validate_tickers(cls, v: list[str]) -> list[str]:
        cleaned = []
        for t in v:
            t = re.sub(r"[^A-Za-z0-9.]", "", t).upper()
            if t and len(t) <= 10:
                cleaned.append(t)
        if not cleaned:
            raise ValueError("At least one valid ticker required")
        if len(cleaned) > 3000:
            raise ValueError("Maximum 3000 tickers per subscription")
        return cleaned

    @field_validator("calendar_type")
    @classmethod
    def validate_calendar_type(cls, v: str) -> str:
        if v not in ("outlook", "gmail"):
            raise ValueError("calendar_type must be 'outlook' or 'gmail'")
        return v


class SubscribeResponse(BaseModel):
    feed_url: str
    events_confirmed: int
    events_pending: int
    message: str


class WatchlistUploadResponse(BaseModel):
    tickers_found: list[str]
    tickers_unknown: list[str]
    events_confirmed: int
    events_pending: int
    feed_url: Optional[str] = None
    message: str


# --- Health check ---

class HealthResponse(BaseModel):
    status: str
    total_events: int
    confirmed_events: int
    total_companies: int
    last_scrape: Optional[datetime] = None
    version: str = "1.0.0"
