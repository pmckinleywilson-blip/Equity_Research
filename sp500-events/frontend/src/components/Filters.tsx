"use client";

interface FiltersProps {
  index: string;
  onIndexChange: (v: string) => void;
  eventType: string;
  onEventTypeChange: (v: string) => void;
  confirmedOnly: boolean;
  onConfirmedOnlyChange: (v: boolean) => void;
  dateFrom: string;
  onDateFromChange: (v: string) => void;
  dateTo: string;
  onDateToChange: (v: string) => void;
}

const selectClass =
  "px-1.5 py-1 border border-[#ccc] bg-white text-[10px] font-mono focus:outline-none focus:border-[#0550ae]";
const inputClass =
  "px-1.5 py-1 border border-[#ccc] bg-white text-[10px] font-mono focus:outline-none focus:border-[#0550ae]";

export default function Filters({
  index,
  onIndexChange,
  eventType,
  onEventTypeChange,
  confirmedOnly,
  onConfirmedOnlyChange,
  dateFrom,
  onDateFromChange,
  dateTo,
  onDateToChange,
}: FiltersProps) {
  return (
    <div className="flex flex-wrap items-center gap-2 mb-2 py-1.5 border-b border-[#ddd] text-[10px] c-muted">
      <label>
        FROM{" "}
        <input
          type="date"
          value={dateFrom}
          onChange={(e) => onDateFromChange(e.target.value)}
          className={inputClass}
        />
      </label>
      <label>
        TO{" "}
        <input
          type="date"
          value={dateTo}
          onChange={(e) => onDateToChange(e.target.value)}
          className={inputClass}
        />
      </label>
      <label>
        INDEX{" "}
        <select
          value={index}
          onChange={(e) => onIndexChange(e.target.value)}
          className={selectClass}
        >
          <option value="">ALL</option>
          <option value="sp500">S&P 500</option>
        </select>
      </label>
      <label>
        TYPE{" "}
        <select
          value={eventType}
          onChange={(e) => onEventTypeChange(e.target.value)}
          className={selectClass}
        >
          <option value="">ALL</option>
          <option value="earnings">EARN</option>
          <option value="investor_day">INV DAY</option>
          <option value="conference">CONF</option>
          <option value="ad_hoc">AD HOC</option>
        </select>
      </label>
      <label className="flex items-center gap-1 cursor-pointer">
        <input
          type="checkbox"
          checked={confirmedOnly}
          onChange={(e) => onConfirmedOnlyChange(e.target.checked)}
          className="w-3 h-3"
        />
        CONFIRMED ONLY
      </label>
    </div>
  );
}
