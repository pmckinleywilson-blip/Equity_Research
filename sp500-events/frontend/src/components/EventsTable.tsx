"use client";

import { useState, useMemo } from "react";
import {
  useReactTable,
  getCoreRowModel,
  getSortedRowModel,
  flexRender,
  type ColumnDef,
  type SortingState,
} from "@tanstack/react-table";
import type { EventItem } from "@/lib/types";
import { getOutlookIcsUrl, getGmailUrl } from "@/lib/api";

interface EventsTableProps {
  events: EventItem[];
  onSelectionChange?: (selectedIds: number[]) => void;
  watchlistTickers?: string[] | null;
  searchValue?: string;
  onSearchChange?: (value: string) => void;
}

function formatTime(t: string | null): string {
  if (!t) return "";
  const [h, m] = t.split(":");
  const hour = parseInt(h);
  const ampm = hour >= 12 ? "p" : "a";
  const h12 = hour > 12 ? hour - 12 : hour || 12;
  return `${h12}:${m}${ampm}`;
}

function formatDate(d: string): string {
  const date = new Date(d + "T00:00:00");
  const day = date.getDate();
  const months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  return `${day} ${months[date.getMonth()]}`;
}

function getDayLabel(d: string): string {
  const date = new Date(d + "T00:00:00");
  const days = ["SUN","MON","TUE","WED","THU","FRI","SAT"];
  const months = ["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"];
  return `${days[date.getDay()]} ${date.getDate()} ${months[date.getMonth()]}`;
}

export default function EventsTable({
  events,
  onSelectionChange,
  watchlistTickers,
  searchValue = "",
  onSearchChange,
}: EventsTableProps) {
  const [sorting, setSorting] = useState<SortingState>([
    { id: "event_date", desc: false },
  ]);
  const [rowSelection, setRowSelection] = useState<Record<string, boolean>>({});

  const handleGmailClick = async (eventId: number) => {
    const url = await getGmailUrl(eventId);
    window.open(url, "_blank");
  };

  // Group events by date for separator rows
  const dateGroups = useMemo(() => {
    const groups = new Map<string, number>();
    events.forEach((e, i) => {
      if (!groups.has(e.event_date)) groups.set(e.event_date, i);
    });
    return groups;
  }, [events]);

  const columns = useMemo<ColumnDef<EventItem, any>[]>(
    () => [
      {
        id: "select",
        header: () => "",
        cell: ({ row }) => (
          <input
            type="checkbox"
            checked={row.getIsSelected()}
            onChange={row.getToggleSelectedHandler()}
            className="w-3 h-3"
          />
        ),
        size: 24,
        enableSorting: false,
      },
      {
        id: "watch",
        header: "",
        cell: ({ row }) => {
          const isWatched = watchlistTickers?.includes(row.original.ticker);
          return isWatched ? <span className="ew-link">W</span> : "";
        },
        size: 20,
        enableSorting: false,
      },
      {
        accessorKey: "event_date",
        header: "Date",
        cell: ({ getValue }) => formatDate(getValue() as string),
        size: 55,
      },
      {
        accessorKey: "ticker",
        header: "Ticker",
        size: 55,
        cell: ({ getValue }) => (
          <span className="c-blue">{getValue() as string}</span>
        ),
      },
      {
        accessorKey: "company_name",
        header: "Company",
        size: 180,
        cell: ({ getValue }) => (
          <span
            className="block truncate max-w-[180px]"
            title={getValue() as string}
          >
            {getValue() as string}
          </span>
        ),
      },
      {
        accessorKey: "event_type",
        header: "Type",
        size: 65,
        cell: ({ getValue }) => {
          const type = getValue() as string;
          const labels: Record<string, string> = {
            earnings: "EARN",
            investor_day: "INV",
            conference: "CONF",
            ad_hoc: "ADHC",
          };
          return <span className="c-muted">{labels[type] || type}</span>;
        },
      },
      {
        accessorKey: "event_time",
        header: "Time",
        size: 50,
        cell: ({ row }) => {
          const t = row.original.event_time;
          if (!t)
            return <span className="c-muted">--:--</span>;
          return formatTime(t);
        },
      },
      {
        id: "actions",
        header: "",
        size: 140,
        cell: ({ row }) => {
          const event = row.original;
          return (
            <span className="flex gap-2 text-[9px]">
              <a
                href={getOutlookIcsUrl(event.id)}
                title="Download .ics file for Outlook"
                className="c-blue hover:underline"
              >
                +Outlook
              </a>
              <button
                onClick={() => handleGmailClick(event.id)}
                title="Open in Google Calendar"
                className="c-blue hover:underline cursor-pointer"
              >
                +Gmail
              </button>
              {event.webcast_url && (
                <a
                  href={event.webcast_url}
                  target="_blank"
                  rel="noopener noreferrer"
                  title="Open webcast link"
                  className="c-green hover:underline"
                >
                  Webcast
                </a>
              )}
            </span>
          );
        },
        enableSorting: false,
      },
      {
        accessorKey: "status",
        header: "Status",
        size: 70,
        cell: ({ row }) => {
          const verified = row.original.ir_verified;
          const status = row.original.status;
          return (
            <span
              className={`ew-badge ${
                status === "confirmed" ? "ew-confirmed" : "ew-tentative"
              }`}
            >
              {status === "confirmed" ? "CONF" : "TENT"}
            </span>
          );
        },
      },
    ],
    [watchlistTickers]
  );

  const table = useReactTable({
    data: events,
    columns,
    state: { sorting, rowSelection },
    onSortingChange: setSorting,
    onRowSelectionChange: (updater) => {
      const newSelection =
        typeof updater === "function" ? updater(rowSelection) : updater;
      setRowSelection(newSelection);
      if (onSelectionChange) {
        const selectedIds = Object.keys(newSelection)
          .filter((k) => newSelection[k])
          .map((k) => events[parseInt(k)]?.id)
          .filter(Boolean);
        onSelectionChange(selectedIds);
      }
    },
    getCoreRowModel: getCoreRowModel(),
    getSortedRowModel: getSortedRowModel(),
    enableRowSelection: true,
  });

  return (
    <div>
      {/* Search */}
      <div className="mb-2">
        <input
          type="text"
          placeholder="search ticker or company..."
          value={searchValue}
          onChange={(e) => onSearchChange?.(e.target.value)}
          className="w-full max-w-xs px-2 py-1 border border-[#ccc] bg-white text-[11px] font-mono focus:outline-none focus:border-[#0550ae]"
        />
      </div>

      {/* Table */}
      <table className="ew-tbl">
        <thead>
          {table.getHeaderGroups().map((headerGroup) => (
            <tr key={headerGroup.id}>
              {headerGroup.headers.map((header) => (
                <th
                  key={header.id}
                  style={{ width: header.getSize() }}
                  onClick={header.column.getToggleSortingHandler()}
                  className="cursor-pointer select-none"
                >
                  {header.isPlaceholder
                    ? null
                    : flexRender(
                        header.column.columnDef.header,
                        header.getContext()
                      )}
                  {header.column.getIsSorted() === "asc"
                    ? " \u25B2"
                    : header.column.getIsSorted() === "desc"
                    ? " \u25BC"
                    : ""}
                </th>
              ))}
            </tr>
          ))}
        </thead>
        <tbody>
          {table.getRowModel().rows.map((row, idx) => {
            const event = row.original;
            const isFirstOfDate = dateGroups.get(event.event_date) === row.index;
            const isWatched = watchlistTickers?.includes(event.ticker);

            return (
              <>
                {isFirstOfDate && (
                  <tr key={`sep-${event.event_date}`} className="ew-sep">
                    <td colSpan={columns.length}>
                      {getDayLabel(event.event_date)}
                    </td>
                  </tr>
                )}
                <tr
                  key={row.id}
                  className={`${isWatched ? "ew-watched" : ""} ${
                    row.getIsSelected() ? "ew-watched" : ""
                  }`}
                >
                  {row.getVisibleCells().map((cell) => (
                    <td key={cell.id}>
                      {flexRender(
                        cell.column.columnDef.cell,
                        cell.getContext()
                      )}
                    </td>
                  ))}
                </tr>
              </>
            );
          })}
        </tbody>
      </table>

      {events.length === 0 && (
        <div className="text-center py-6 c-muted">
          No events found. Events will appear here as wire service announcements
          are detected.
        </div>
      )}
    </div>
  );
}
