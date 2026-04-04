const API_BASE = process.env.NEXT_PUBLIC_API_URL || "http://localhost:8000";

export async function fetchEvents(params: Record<string, string> = {}): Promise<any> {
  const searchParams = new URLSearchParams(params);
  const res = await fetch(`${API_BASE}/api/v1/events?${searchParams}`);
  if (!res.ok) throw new Error(`API error: ${res.status}`);
  return res.json();
}

export async function fetchCompanies(index?: string): Promise<any> {
  const params = index ? `?index=${index}` : "";
  const res = await fetch(`${API_BASE}/api/v1/companies${params}`);
  if (!res.ok) throw new Error(`API error: ${res.status}`);
  return res.json();
}

export function getOutlookIcsUrl(eventId: number): string {
  return `${API_BASE}/api/v1/calendar/${eventId}.ics`;
}

export async function getGmailUrl(eventId: number): Promise<string> {
  const res = await fetch(`${API_BASE}/api/v1/calendar/${eventId}/gmail`);
  if (!res.ok) throw new Error(`API error: ${res.status}`);
  const data = await res.json();
  return data.gmail_url;
}

export function getBulkIcsUrl(tickers?: string[]): string {
  const params = tickers?.length ? `?tickers=${tickers.join(",")}` : "";
  return `${API_BASE}/api/v1/calendar.ics${params}`;
}

export function getFeedUrl(token: string): string {
  return `${API_BASE}/api/v1/feed/${token}.ics`;
}

export async function subscribe(data: {
  email: string;
  tickers: string[];
  calendar_type: string;
}): Promise<any> {
  const res = await fetch(`${API_BASE}/api/v1/subscribe`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(data),
  });
  if (!res.ok) {
    const err = await res.json().catch(() => ({ detail: "Request failed" }));
    throw new Error(err.detail || "Subscribe failed");
  }
  return res.json();
}

export async function uploadWatchlist(
  file: File,
  email?: string,
  calendarType?: string
): Promise<any> {
  const formData = new FormData();
  formData.append("csv_file", file);
  if (email) formData.append("email", email);
  if (calendarType) formData.append("calendar_type", calendarType);

  const res = await fetch(`${API_BASE}/api/v1/watchlist`, {
    method: "POST",
    body: formData,
  });
  if (!res.ok) {
    const err = await res.json().catch(() => ({ detail: "Upload failed" }));
    throw new Error(err.detail || "Upload failed");
  }
  return res.json();
}
