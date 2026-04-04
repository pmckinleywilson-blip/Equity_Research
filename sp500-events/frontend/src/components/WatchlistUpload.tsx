"use client";

import { useState, useRef } from "react";

interface WatchlistUploadProps {
  onUpload: (file: File) => void;
  onClear: () => void;
  isActive: boolean;
  tickerCount?: number;
}

export default function WatchlistUpload({
  onUpload,
  onClear,
  isActive,
  tickerCount,
}: WatchlistUploadProps) {
  const [isDragging, setIsDragging] = useState(false);
  const fileRef = useRef<HTMLInputElement>(null);

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    const file = e.dataTransfer.files[0];
    if (file && (file.name.endsWith(".csv") || file.type === "text/csv")) {
      onUpload(file);
    }
  };

  if (isActive) {
    return (
      <span className="text-[10px]">
        WATCHLIST: <strong className="c-blue">{tickerCount}</strong> tickers{" "}
        <button
          onClick={onClear}
          className="c-red hover:underline cursor-pointer ml-1"
        >
          [clear]
        </button>
      </span>
    );
  }

  return (
    <span
      onDragOver={(e) => {
        e.preventDefault();
        setIsDragging(true);
      }}
      onDragLeave={() => setIsDragging(false)}
      onDrop={handleDrop}
      onClick={() => fileRef.current?.click()}
      className={`text-[10px] cursor-pointer ${
        isDragging ? "c-blue" : "c-muted hover:text-[#1b1b1b]"
      }`}
    >
      [UPLOAD CSV WATCHLIST]
      <input
        ref={fileRef}
        type="file"
        accept=".csv"
        className="hidden"
        onChange={(e) => {
          const f = e.target.files?.[0];
          if (f) onUpload(f);
        }}
      />
    </span>
  );
}
