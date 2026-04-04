# SP500 Events

Free, machine-readable earnings calendar for S&P 500 and Russell 3000 companies.

## Features

- **Full coverage**: Every announced event for all ~3000 Russell 3000 + S&P 500 companies
- **Machine-readable API**: JSON, RSS, iCal — no auth required for read endpoints
- **Calendar integration**: Direct .ics downloads, subscribable calendar URLs, auto-invite emails
- **Webcast details**: Links, dial-in numbers, passcodes extracted from company IR pages
- **Auto-invites**: Subscribe with your watchlist, receive calendar invites as events are confirmed

## Quick Start

### Backend

```bash
cd backend
cp .env.example .env
# Edit .env with your settings
pip install -r requirements.txt
uvicorn main:app --reload
```

API available at http://localhost:8000/api/v1/docs

### Frontend

```bash
cd frontend
cp .env.local.example .env.local
npm install
npm run dev
```

UI available at http://localhost:3000

### Populate Data

```bash
# Initial company list (Russell 3000 + S&P 500)
cd scrapers
python tickers.py

# Run detection scrapers
python earnings_calendars.py

# Run full pipeline (detect + verify + notify)
python pipeline.py
```

## Architecture

```
Detection Sources          →   IR Page Verification   →   Notify Subscribers
(Nasdaq, Finnhub, EDGAR,      (source of truth)          (direct calendar invites)
 Press Releases)               ir_verified = TRUE          via METHOD:REQUEST .ics
```

## API Endpoints

| Endpoint | Description |
|----------|-------------|
| `GET /api/v1/events` | List events (filterable by ticker, date, type, sector, index) |
| `GET /api/v1/events/{ticker}` | Events for one company |
| `GET /api/v1/calendar.ics` | iCal feed (subscribable) |
| `GET /api/v1/feed/<token>.ics` | Personal watchlist calendar feed |
| `GET /api/v1/companies` | Company list |
| `POST /api/v1/subscribe` | Subscribe for auto-invites |
| `POST /api/v1/watchlist` | Upload CSV watchlist |

## Cost

~$2/month on free tiers (Vercel + Render + Neon + Resend + GitHub Actions).
