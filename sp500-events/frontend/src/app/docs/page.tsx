import type { Metadata } from "next";

export const metadata: Metadata = {
  title: "API | EARNINGS WIRE",
  description:
    "Free, zero-auth API for earnings dates, webcast links, and corporate events.",
};

export default function DocsPage() {
  const base = "https://earningswire.com";

  return (
    <div className="max-w-2xl">
      <div className="text-[12px] font-medium tracking-[1px] mb-1">
        API DOCUMENTATION
      </div>
      <div className="text-[10px] c-muted mb-4">
        Free, machine-readable API. No authentication required for read
        endpoints.
      </div>

      <div className="space-y-4 text-[10px]">
        <section>
          <div className="text-[9px] c-muted uppercase tracking-wider mb-1 border-b border-[#ddd] pb-0.5">
            BASE URL
          </div>
          <pre className="px-2 py-1 bg-[#f2f2f2] border border-[#ddd] overflow-x-auto">
            {base}/api/v1
          </pre>
        </section>

        <section>
          <div className="text-[9px] c-muted uppercase tracking-wider mb-1 border-b border-[#ddd] pb-0.5">
            EVENTS
          </div>
          <pre className="px-2 py-1 bg-[#f2f2f2] border border-[#ddd] overflow-x-auto whitespace-pre-wrap">{`GET /api/v1/events
  ?ticker=AAPL,MSFT          filter by tickers
  ?date_from=2026-04-01      start date
  ?date_to=2026-04-30        end date
  ?type=earnings             earnings|investor_day|conference|ad_hoc
  ?sector=Technology         sector filter
  ?index=sp500               sp500|russell3000
  ?confirmed_only=true       only wire-confirmed events
  ?format=rss                json (default) or rss
  ?page=1&per_page=100       pagination

GET /api/v1/events/{ticker}  events for one company`}</pre>
        </section>

        <section>
          <div className="text-[9px] c-muted uppercase tracking-wider mb-1 border-b border-[#ddd] pb-0.5">
            CALENDAR
          </div>
          <pre className="px-2 py-1 bg-[#f2f2f2] border border-[#ddd] overflow-x-auto whitespace-pre-wrap">{`GET /api/v1/calendar.ics         full calendar (subscribable URL)
  ?tickers=AAPL,MSFT             filtered
  ?confirmed_only=true

GET /api/v1/calendar/{id}.ics    single event .ics
GET /api/v1/calendar/{id}/gmail  Google Calendar add URL
GET /api/v1/feed/{token}.ics     personal watchlist feed`}</pre>
        </section>

        <section>
          <div className="text-[9px] c-muted uppercase tracking-wider mb-1 border-b border-[#ddd] pb-0.5">
            SUBSCRIBE
          </div>
          <pre className="px-2 py-1 bg-[#f2f2f2] border border-[#ddd] overflow-x-auto whitespace-pre-wrap">{`POST /api/v1/subscribe
  {"email":"...", "tickers":["AAPL","MSFT"], "calendar_type":"outlook"}

POST /api/v1/watchlist           upload CSV watchlist`}</pre>
        </section>

        <section>
          <div className="text-[9px] c-muted uppercase tracking-wider mb-1 border-b border-[#ddd] pb-0.5">
            RESPONSE
          </div>
          <pre className="px-2 py-1 bg-[#f2f2f2] border border-[#ddd] overflow-x-auto whitespace-pre-wrap">{`{
  "events": [{
    "id": 1,
    "ticker": "AAPL",
    "company_name": "Apple Inc.",
    "event_type": "earnings",
    "event_date": "2026-04-30",
    "event_time": "17:00:00",
    "timezone": "America/New_York",
    "title": "Apple Q2 2026 Earnings Call",
    "webcast_url": "https://...",
    "phone_number": "+1-800-...",
    "status": "confirmed"
  }],
  "total": 1, "page": 1, "per_page": 100
}`}</pre>
        </section>

        <section>
          <div className="text-[9px] c-muted uppercase tracking-wider mb-1 border-b border-[#ddd] pb-0.5">
            RATE LIMITS
          </div>
          <div className="c-muted">
            60 req/min anonymous. CORS open on all read endpoints. No API key
            required.
          </div>
        </section>

        <section>
          <div className="text-[9px] c-muted uppercase tracking-wider mb-1 border-b border-[#ddd] pb-0.5">
            INTERACTIVE DOCS
          </div>
          <div>
            <a
              href={`${base}/api/v1/docs`}
              className="c-blue hover:underline"
              target="_blank"
            >
              {base}/api/v1/docs
            </a>{" "}
            <span className="c-muted">— OpenAPI / Swagger</span>
          </div>
        </section>
      </div>
    </div>
  );
}
