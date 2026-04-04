export interface EventItem {
  id: number;
  ticker: string;
  company_name: string;
  event_type: string;
  event_date: string;
  event_time: string | null;
  timezone: string;
  title: string | null;
  description: string | null;
  webcast_url: string | null;
  phone_number: string | null;
  phone_passcode: string | null;
  replay_url: string | null;
  fiscal_quarter: string | null;
  source: string;
  source_url: string | null;
  ir_verified: boolean;
  status: string;
  created_at: string;
  updated_at: string;
}

export interface EventListResponse {
  events: EventItem[];
  total: number;
  page: number;
  per_page: number;
  pages: number;
}

export interface CompanyItem {
  ticker: string;
  company_name: string;
  sector: string | null;
  sub_industry: string | null;
  market_cap_tier: string | null;
}

export interface SubscribeResponse {
  feed_url: string;
  events_confirmed: number;
  events_pending: number;
  message: string;
}
