"""SP500 Events API — Free earnings calendar for AI agents and analysts."""
from fastapi import FastAPI, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse, PlainTextResponse
from slowapi import Limiter, _rate_limit_exceeded_handler
from slowapi.util import get_remote_address
from slowapi.errors import RateLimitExceeded

from config import get_settings
from database import init_db
from routers import events, calendar, watchlist

settings = get_settings()

limiter = Limiter(key_func=get_remote_address, default_limits=[settings.rate_limit_default])

app = FastAPI(
    title="SP500 Events API",
    description=(
        "Free, machine-readable API for S&P 500 and Russell 3000 corporate events. "
        "Earnings calls, investor days, ad-hoc events with webcast links and dial-in details. "
        "No authentication required for read endpoints."
    ),
    version="1.0.0",
    docs_url="/api/v1/docs",
    redoc_url="/api/v1/redoc",
    openapi_url="/api/v1/openapi.json",
)

app.state.limiter = limiter
app.add_exception_handler(RateLimitExceeded, _rate_limit_exceeded_handler)

# CORS: open for read endpoints (AI agents), restricted for write
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["GET", "POST", "DELETE", "OPTIONS"],
    allow_headers=["*"],
)

# Include routers
app.include_router(events.router)
app.include_router(calendar.router)
app.include_router(watchlist.router)


@app.on_event("startup")
def startup():
    init_db()


@app.get("/", include_in_schema=False)
def root():
    return {
        "name": "SP500 Events API",
        "description": "Free earnings calendar for AI agents and equity analysts",
        "docs": "/api/v1/docs",
        "events": "/api/v1/events",
        "calendar": "/api/v1/calendar.ics",
        "companies": "/api/v1/companies",
        "health": "/api/v1/health",
    }


@app.get("/robots.txt", include_in_schema=False)
def robots_txt():
    """Allow all crawlers — we want to be indexed."""
    return PlainTextResponse(
        "User-agent: *\n"
        "Allow: /\n"
        "Sitemap: https://sp500events.com/sitemap.xml\n"
    )


@app.get("/.well-known/llms.txt", include_in_schema=False)
def llms_txt():
    """Machine-readable description for LLM agents."""
    return PlainTextResponse(
        "# SP500 Events API\n"
        "# Free earnings calendar for S&P 500 and Russell 3000 companies\n"
        "#\n"
        "# Base URL: https://sp500events.com/api/v1\n"
        "# No authentication required for read endpoints\n"
        "#\n"
        "# Endpoints:\n"
        "# GET /api/v1/events — List upcoming events (JSON)\n"
        "#   ?ticker=AAPL,MSFT — Filter by tickers\n"
        "#   ?date_from=2026-04-01&date_to=2026-04-30 — Date range\n"
        "#   ?type=earnings — Event type (earnings, investor_day, conference, ad_hoc)\n"
        "#   ?sector=Technology — Sector filter\n"
        "#   ?index=sp500 — Index membership filter\n"
        "#   ?confirmed_only=true — Only IR-verified events\n"
        "#   ?format=rss — RSS feed output\n"
        "#\n"
        "# GET /api/v1/events/{ticker} — Events for one company\n"
        "# GET /api/v1/calendar.ics — iCalendar feed (subscribable)\n"
        "#   ?tickers=AAPL,MSFT — Filtered calendar\n"
        "# GET /api/v1/companies — Company list with metadata\n"
        "# GET /api/v1/health — API health and data freshness\n"
        "#\n"
        "# Full OpenAPI spec: /api/v1/openapi.json\n"
    )


# Security headers middleware
@app.middleware("http")
async def add_security_headers(request: Request, call_next):
    response = await call_next(request)
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers["X-Frame-Options"] = "DENY"
    response.headers["Referrer-Policy"] = "strict-origin-when-cross-origin"
    response.headers["Permissions-Policy"] = "camera=(), microphone=(), geolocation=()"
    # CSP is permissive for API, frontend will set stricter policy
    return response
