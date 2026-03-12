"use client";

import { useState, useRef, DragEvent, ChangeEvent } from "react";

interface ProcessResult {
  filename: string;
  rows:      number;
  dateRange: string;
  webUrl:    string;
}

export default function HomePage() {
  const [clientName, setClientName]   = useState("");
  const [file, setFile]               = useState<File | null>(null);
  const [dragging, setDragging]       = useState(false);
  const [status, setStatus]           = useState<"idle" | "validating" | "processing" | "success" | "error">("idle");
  const [result, setResult]           = useState<ProcessResult | null>(null);
  const [errorMsg, setErrorMsg]       = useState("");
  const [warnings, setWarnings]       = useState<string[]>([]);
  const fileInputRef                  = useRef<HTMLInputElement>(null);

  function handleDrop(e: DragEvent<HTMLDivElement>) {
    e.preventDefault();
    setDragging(false);
    const dropped = e.dataTransfer.files[0];
    if (dropped) pickFile(dropped);
  }

  function handleFileChange(e: ChangeEvent<HTMLInputElement>) {
    const picked = e.target.files?.[0];
    if (picked) pickFile(picked);
  }

  function pickFile(f: File) {
    if (!f.name.endsWith(".xlsx") && !f.name.endsWith(".xls")) {
      setErrorMsg("Please upload an Excel file (.xlsx or .xls)");
      setStatus("error");
      return;
    }
    setFile(f);
    setStatus("idle");
    setResult(null);
    setErrorMsg("");
    setWarnings([]);
  }

  async function handleProcess() {
    if (!file)              { setErrorMsg("Please select a file first");    setStatus("error"); return; }
    if (!clientName.trim()) { setErrorMsg("Please enter the client name");  setStatus("error"); return; }

    setStatus("validating");
    setResult(null);
    setErrorMsg("");
    setWarnings([]);

    // Run validation first
    try {
      const fd = new FormData();
      fd.append("file", file);
      fd.append("clientName", clientName.trim());

      const vRes  = await fetch("/api/validate", { method: "POST", body: fd });
      const vJson = await vRes.json();

      if (vJson.warnings && vJson.warnings.length > 0) {
        // Show warnings and wait for user decision
        setWarnings(vJson.warnings);
        setStatus("idle");
        return;
      }
    } catch {
      // Validation failure should not block processing
    }

    await runProcess();
  }

  async function runProcess() {
    if (!file) return;

    setStatus("processing");
    setWarnings([]);

    const fd = new FormData();
    fd.append("file", file);
    fd.append("clientName", clientName.trim());

    try {
      const res  = await fetch("/api/process", { method: "POST", body: fd });
      const json = await res.json();

      if (!res.ok || !json.success) {
        setErrorMsg(json.error ?? "An unknown error occurred");
        setStatus("error");
        return;
      }
      setResult(json as ProcessResult);
      setStatus("success");
    } catch (err) {
      setErrorMsg(err instanceof Error ? err.message : "Network error");
      setStatus("error");
    }
  }

  function dismissWarnings() {
    setWarnings([]);
  }

  function reset() {
    setFile(null);
    setClientName("");
    setStatus("idle");
    setResult(null);
    setErrorMsg("");
    setWarnings([]);
    if (fileInputRef.current) fileInputRef.current.value = "";
  }

  const isBusy     = status === "validating" || status === "processing";
  const canProcess = !!file && !!clientName.trim() && !isBusy;

  return (
    <div className="flex-1 flex items-start justify-center pt-12 px-4 pb-16">
      <div style={{ width: '100%', maxWidth: '640px' }}>

        {/* Page title */}
        <div className="mb-8">
          <h1 style={{ fontSize: '1.5rem', fontWeight: 700, color: '#f1f5f9', marginBottom: '0.4rem' }}>
            Checkers Raw Data Converter
          </h1>
          <p style={{ color: '#94a3b8', fontSize: '0.9rem' }}>
            Upload a raw <code style={{ color: '#F97316', background: '#1e293b', padding: '2px 6px', borderRadius: 4 }}>vnd-art-sales</code> file
            and it will be converted to the clean <strong style={{ color: '#f1f5f9' }}>CHECKERS B2B</strong> format
            and saved to SharePoint automatically.
          </p>
        </div>

        {/* Card */}
        <div style={{
          background: '#1e293b',
          border: '1px solid #334155',
          borderRadius: '12px',
          padding: '2rem',
          display: 'flex',
          flexDirection: 'column',
          gap: '1.5rem',
        }}>

          {/* Client name */}
          <div>
            <label style={{ display: 'block', fontWeight: 600, marginBottom: '0.4rem', color: '#e2e8f0', fontSize: '0.9rem' }}>
              Client Name
            </label>
            <input
              type="text"
              value={clientName}
              onChange={e => { setClientName(e.target.value); setWarnings([]); }}
              placeholder="e.g. WAHL"
              style={{
                width: '100%',
                padding: '0.6rem 0.85rem',
                borderRadius: '8px',
                border: '1px solid #475569',
                background: '#0f172a',
                color: '#f1f5f9',
                fontSize: '0.95rem',
                outline: 'none',
                boxSizing: 'border-box',
              }}
            />
            <p style={{ marginTop: '0.35rem', color: '#64748b', fontSize: '0.78rem' }}>
              This name is used in the output filename: <em>CHECKERS B2B [Client Name] YYYY-MM-DD.xlsx</em>
            </p>
          </div>

          {/* Drop zone */}
          <div>
            <label style={{ display: 'block', fontWeight: 600, marginBottom: '0.4rem', color: '#e2e8f0', fontSize: '0.9rem' }}>
              Raw File
            </label>
            <div
              onClick={() => fileInputRef.current?.click()}
              onDragOver={e => { e.preventDefault(); setDragging(true); }}
              onDragLeave={() => setDragging(false)}
              onDrop={handleDrop}
              style={{
                border: `2px dashed ${dragging ? '#F97316' : file ? '#22c55e' : '#475569'}`,
                borderRadius: '10px',
                padding: '2.5rem 1.5rem',
                textAlign: 'center',
                cursor: 'pointer',
                background: dragging ? 'rgba(249,115,22,0.06)' : '#0f172a',
                transition: 'all 0.2s',
              }}
            >
              <input
                ref={fileInputRef}
                type="file"
                accept=".xlsx,.xls"
                onChange={handleFileChange}
                style={{ display: 'none' }}
              />
              {file ? (
                <>
                  <div style={{ fontSize: '1.5rem', marginBottom: '0.4rem' }}>✅</div>
                  <p style={{ color: '#22c55e', fontWeight: 600, marginBottom: '0.2rem' }}>{file.name}</p>
                  <p style={{ color: '#64748b', fontSize: '0.8rem' }}>
                    {(file.size / 1024 / 1024).toFixed(2)} MB — click or drag to replace
                  </p>
                </>
              ) : (
                <>
                  <div style={{ fontSize: '2rem', marginBottom: '0.5rem', color: '#475569' }}>📂</div>
                  <p style={{ color: '#94a3b8', marginBottom: '0.25rem' }}>
                    Drag &amp; drop your <strong style={{ color: '#f1f5f9' }}>vnd-art-sales</strong> file here
                  </p>
                  <p style={{ color: '#64748b', fontSize: '0.8rem' }}>or click to browse — .xlsx files only</p>
                </>
              )}
            </div>
          </div>

          {/* Process button */}
          <button
            onClick={handleProcess}
            disabled={!canProcess}
            style={{
              padding: '0.75rem 1.5rem',
              borderRadius: '8px',
              fontWeight: 700,
              fontSize: '1rem',
              background: canProcess ? '#F97316' : '#374151',
              color: canProcess ? '#fff' : '#6b7280',
              border: 'none',
              cursor: canProcess ? 'pointer' : 'not-allowed',
              transition: 'background 0.2s',
              width: '100%',
            }}
          >
            {status === "validating"  ? "⏳  Checking file…"         :
             status === "processing"  ? "⏳  Processing…"             :
             "Convert & Upload to SharePoint"}
          </button>
        </div>

        {/* ── Validation warnings ───────────────────────────────────────────── */}
        {warnings.length > 0 && (
          <div style={{
            marginTop: '1rem',
            background: '#1e293b',
            border: '1px solid #f59e0b',
            borderRadius: '10px',
            padding: '1.25rem 1.5rem',
          }}>
            <p style={{ color: '#f59e0b', fontWeight: 700, fontSize: '1rem', marginBottom: '0.75rem' }}>
              ⚠️  Validation warnings
            </p>
            <ul style={{ margin: 0, paddingLeft: '1.2rem', display: 'flex', flexDirection: 'column', gap: '0.5rem' }}>
              {warnings.map((w, i) => (
                <li key={i} style={{ color: '#fde68a', fontSize: '0.88rem', lineHeight: 1.5 }}>{w}</li>
              ))}
            </ul>
            <p style={{ marginTop: '0.85rem', color: '#94a3b8', fontSize: '0.82rem' }}>
              You can proceed and ignore these warnings, or start over with the correct file.
            </p>
            <div style={{ display: 'flex', gap: '0.75rem', marginTop: '1rem', flexWrap: 'wrap' }}>
              <button
                onClick={runProcess}
                style={{
                  padding: '0.5rem 1.1rem',
                  borderRadius: '6px',
                  fontWeight: 700,
                  fontSize: '0.88rem',
                  background: '#F97316',
                  color: '#fff',
                  border: 'none',
                  cursor: 'pointer',
                }}
              >
                Proceed anyway
              </button>
              <button
                onClick={reset}
                style={{
                  padding: '0.5rem 1.1rem',
                  borderRadius: '6px',
                  fontWeight: 600,
                  fontSize: '0.88rem',
                  background: 'transparent',
                  color: '#94a3b8',
                  border: '1px solid #475569',
                  cursor: 'pointer',
                }}
              >
                Start over
              </button>
            </div>
          </div>
        )}

        {/* ── Error state ────────────────────────────────────────────────────── */}
        {status === "error" && (
          <div style={{
            marginTop: '1rem',
            background: '#1e293b',
            border: '1px solid #ef4444',
            borderRadius: '10px',
            padding: '1rem 1.25rem',
          }}>
            <p style={{ color: '#ef4444', fontWeight: 600, marginBottom: '0.25rem' }}>❌  Error</p>
            <p style={{ color: '#fca5a5', fontSize: '0.9rem' }}>{errorMsg}</p>
          </div>
        )}

        {/* ── Success state ──────────────────────────────────────────────────── */}
        {status === "success" && result && (
          <div style={{
            marginTop: '1rem',
            background: '#1e293b',
            border: '1px solid #22c55e',
            borderRadius: '10px',
            padding: '1.25rem 1.5rem',
          }}>
            <p style={{ color: '#22c55e', fontWeight: 700, fontSize: '1rem', marginBottom: '0.75rem' }}>
              ✅  Conversion complete
            </p>
            <div style={{ display: 'flex', flexDirection: 'column', gap: '0.4rem', fontSize: '0.88rem' }}>
              <Row label="File"       value={result.filename} />
              <Row label="Data rows"  value={String(result.rows)} />
              <Row label="Date range" value={result.dateRange} />
            </div>
            {result.webUrl && (
              <a
                href={result.webUrl}
                target="_blank"
                rel="noopener noreferrer"
                style={{
                  display: 'inline-block',
                  marginTop: '1rem',
                  padding: '0.5rem 1rem',
                  background: '#F97316',
                  color: '#fff',
                  borderRadius: '6px',
                  fontWeight: 600,
                  fontSize: '0.85rem',
                  textDecoration: 'none',
                }}
              >
                Open in SharePoint →
              </a>
            )}
            <button
              onClick={reset}
              style={{
                display: 'inline-block',
                marginTop: '0.75rem',
                marginLeft: '0.75rem',
                padding: '0.5rem 1rem',
                background: 'transparent',
                color: '#94a3b8',
                border: '1px solid #475569',
                borderRadius: '6px',
                cursor: 'pointer',
                fontSize: '0.85rem',
              }}
            >
              Convert another file
            </button>
          </div>
        )}

      </div>
    </div>
  );
}

function Row({ label, value }: { label: string; value: string }) {
  return (
    <div style={{ display: 'flex', gap: '0.5rem' }}>
      <span style={{ color: '#64748b', minWidth: '90px' }}>{label}:</span>
      <span style={{ color: '#e2e8f0', fontWeight: 500, wordBreak: 'break-all' }}>{value}</span>
    </div>
  );
}
