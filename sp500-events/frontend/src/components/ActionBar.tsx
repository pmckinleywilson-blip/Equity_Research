"use client";

interface ActionBarProps {
  selectedCount: number;
  selectedIds: number[];
  apiBase: string;
}

export default function ActionBar({
  selectedCount,
  selectedIds,
  apiBase,
}: ActionBarProps) {
  if (selectedCount === 0) return null;

  const handleBulkOutlook = async () => {
    const res = await fetch(`${apiBase}/api/v1/calendar/bulk.ics`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(selectedIds),
    });
    if (!res.ok) return;
    const blob = await res.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "earnings-wire.ics";
    a.click();
    URL.revokeObjectURL(url);
  };

  return (
    <div className="fixed bottom-0 left-0 right-0 bg-[#fefefe] border-t-2 border-[#1b1b1b] z-50">
      <div className="max-w-6xl mx-auto px-4 py-2 flex items-center justify-between text-[10px] font-mono">
        <span className="c-muted">
          {selectedCount} selected
        </span>
        <div className="flex items-center gap-3">
          <button
            onClick={handleBulkOutlook}
            className="px-3 py-1 bg-[#1b1b1b] text-white text-[10px] font-mono tracking-wider hover:bg-[#333]"
          >
            ADD TO OUTLOOK (.ICS)
          </button>
          <a
            href="/subscribe"
            className="px-3 py-1 border border-[#1b1b1b] text-[#1b1b1b] text-[10px] font-mono tracking-wider hover:bg-[#f2f2f2]"
          >
            ADD TO GMAIL
          </a>
        </div>
      </div>
    </div>
  );
}
