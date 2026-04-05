/**
 * App.jsx — Jersey QC Scanner (entire frontend)
 *
 * This is a single-file React component containing all UI and logic.
 * It is intentionally monolithic for ease of deployment in a factory environment.
 *
 * High-level structure:
 *   1. Utility functions (audio, CSV/Excel parsing, matching logic, exports)
 *   2. Shared style objects and small UI primitives (Pill, OverlayBtn, Kbd)
 *   3. Modal/panel components (SessionStartModal, ScanOverlay, SettingsPanel, etc.)
 *   4. Admin sub-components (AdminDashboard, AdminAssign, AdminPackGroups, AdminExports)
 *   5. Root App() component — all state, effects, and main layout
 *
 * All styles are inline (no CSS modules). The theme is a dark GitHub-like palette:
 *   Background:  #0d1117  (page)  /  #161b22  (cards)  /  #1e293b  (inputs)
 *   Borders:     #21262d  (strong)  /  #334155  (subtle)
 *   Text:        #e2e8f0  (primary)  /  #94a3b8  (secondary)  /  #64748b  (muted)
 *   Accent:      #3b82f6  (blue / primary action)
 *   Status:      #22c55e (pass/green)  #ef4444 (flag/red)  #f59e0b (warn/amber)
 */

import { useState, useRef, useCallback, useEffect } from "react";
import akLogo from "./assets/logo_1.png";

// The Claude model used for jersey scanning. Haiku is chosen for speed and cost.
// To update the model, change this constant — no other changes needed.
const CLAUDE_MODEL = "claude-haiku-4-5-20251001";

// API key is read from the environment at build time (set in .env.local).
// It is sent from the browser directly to the Anthropic API via the Vite proxy.
const CLAUDE_API_KEY = import.meta.env.VITE_CLAUDE_API_KEY;

// Scan status constants — written to roster entries and log records.
// These are the only valid values for roster[n].scanned.
const S_PASS    = "pass";     // confirmed match
const S_FLAGGED = "flagged";  // flagged by operator for review
const S_MANUAL  = "manual";   // flag was resolved (Accept/Reject/Note)
const S_EXTRA   = "extra";    // scanned but not in the roster

// ── Audio ─────────────────────────────────────────────────────────────────────
/**
 * Plays a short synthesized tone to give the operator instant audio feedback.
 * Uses the Web Audio API — no external audio files needed.
 *
 * @param {"pass"|"flag"|"scan"} type
 *   "pass"  → ascending two-note chime (good jersey)
 *   "flag"  → descending two-note tone (problem detected)
 *   "scan"  → short single beep (scan trigger acknowledged)
 */
function playTone(type) {
  try {
    const ctx  = new (window.AudioContext || window.webkitAudioContext)();
    const osc  = ctx.createOscillator();
    const gain = ctx.createGain();
    osc.connect(gain);
    gain.connect(ctx.destination);
    if (type === "pass") {
      osc.frequency.setValueAtTime(880, ctx.currentTime);
      osc.frequency.setValueAtTime(1100, ctx.currentTime + 0.1);
      gain.gain.setValueAtTime(0.3, ctx.currentTime);
      gain.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + 0.35);
      osc.start(); osc.stop(ctx.currentTime + 0.35);
    } else if (type === "flag") {
      osc.frequency.setValueAtTime(300, ctx.currentTime);
      osc.frequency.setValueAtTime(220, ctx.currentTime + 0.15);
      gain.gain.setValueAtTime(0.4, ctx.currentTime);
      gain.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + 0.5);
      osc.start(); osc.stop(ctx.currentTime + 0.5);
    } else if (type === "scan") {
      osc.frequency.setValueAtTime(660, ctx.currentTime);
      gain.gain.setValueAtTime(0.15, ctx.currentTime);
      gain.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + 0.1);
      osc.start(); osc.stop(ctx.currentTime + 0.1);
    }
  } catch { /* audio not supported */ }
}

// ── CSV / TSV parser ──────────────────────────────────────────────────────────
/**
 * Parses a CSV or TSV string into an array of roster row objects.
 * Auto-detects the delimiter by counting tabs vs commas in the header row.
 * Passes rows through normaliseRosterRows() for column name normalisation.
 *
 * @param {string} text - Raw file content
 * @returns {Object[]} Array of row objects with lowercase column keys and a numeric `_id`
 */
function parseCSV(text) {
  const lines = text.trim().split(/\r?\n/);
  if (lines.length < 2) return [];
  const delim   = (lines[0].match(/\t/g) || []).length > (lines[0].match(/,/g) || []).length ? "\t" : ",";
  const headers = lines[0].split(delim).map(h => h.trim().toLowerCase());
  return normaliseRosterRows(
    lines.slice(1).map((line, idx) => {
      const vals = line.split(delim).map(v => v.trim());
      const row  = { _id: idx };
      headers.forEach((h, i) => { row[h] = vals[i] || ""; });
      return row;
    }).filter(r => r.name || r.number)
  );
}

// ── Excel parser ──────────────────────────────────────────────────────────────
/**
 * Reads the first sheet of an Excel file (.xlsx / .xls) using the SheetJS library
 * loaded from CDN. Returns a promise that resolves to a normalised roster array.
 *
 * SheetJS must be loaded before calling this (the App component dynamically injects
 * the CDN script tag on mount and tracks readiness with `xlsxReady` state).
 *
 * @param {File} file - File object from an <input type="file"> element
 * @returns {Promise<Object[]>} Resolves to normalised roster rows
 */
function parseExcel(file) {
  return new Promise((resolve, reject) => {
    if (!window.XLSX) { reject(new Error("Excel library not loaded yet.")); return; }
    const reader = new FileReader();
    reader.onerror = () => reject(new Error("Could not read file."));
    reader.onload  = (e) => {
      try {
        const wb   = window.XLSX.read(e.target.result, { type: "array" });
        const ws   = wb.Sheets[wb.SheetNames[0]];
        const rows = window.XLSX.utils.sheet_to_json(ws, { defval: "" });
        resolve(normaliseRosterRows(
          rows.map((r, idx) => {
            const row = { _id: idx };
            Object.keys(r).forEach(k => { row[k.trim().toLowerCase()] = String(r[k]).trim(); });
            return row;
          }).filter(r => r.name || r.number)
        ));
      } catch (err) { reject(err); }
    };
    reader.readAsArrayBuffer(file);
  });
}

// ── Column normalisation ──────────────────────────────────────────────────────
/**
 * Normalises roster rows from various export formats into a consistent shape.
 *
 * Transformations applied:
 *   - "size range" + "size" columns → merged into a single "size" field (e.g. "A-XS", "Y-M")
 *     This matches Athletic Knit's ERP export format where size range (Adult/Youth)
 *     and size (XS/M/L) are separate columns.
 *   - "team name" column → renamed to "team" (alternate column name from some exports)
 *
 * @param {Object[]} rows - Raw parsed rows (already have lowercase keys)
 * @returns {Object[]} Rows with normalised columns
 */
function normaliseRosterRows(rows) {
  if (rows.length === 0) return rows;
  const keys         = Object.keys(rows[0]);
  const hasSizeRange = keys.includes("size range");
  const hasTeamName  = keys.includes("team name") && !keys.includes("team");
  return rows.map(r => {
    const out = { ...r };
    if (hasSizeRange) {
      out.size = `${(r["size range"] || "").trim()}-${(r["size"] || "").trim()}`;
      delete out["size range"];
    }
    if (hasTeamName) {
      out.team = r["team name"] || "";
      delete out["team name"];
    }
    return out;
  });
}

// ── Bin assignment ────────────────────────────────────────────────────────────
/**
 * Builds a bin map from a roster that has a "team" column.
 * Teams are sorted alphabetically and assigned bin numbers starting at 1.
 *
 * Example: ["Senators", "Leafs"] → { "Leafs": 1, "Senators": 2 }
 *
 * The bin map is shown to the operator before scanning starts so they can label
 * physical bins. It is also printed on the scan overlay and in the Excel export.
 *
 * @param {Object[]} roster
 * @returns {{ [teamName: string]: number }}
 */
function buildBinMap(roster) {
  const teams = [...new Set(roster.map(r => r.team).filter(Boolean))].sort();
  const map   = {};
  teams.forEach((t, i) => { map[t] = i + 1; });
  return map;
}

// ── Text normalisation ────────────────────────────────────────────────────────
/**
 * Strips all non-alphanumeric characters and lowercases a string.
 * Used to make name/number comparisons robust against spacing, punctuation,
 * and capitalisation differences between the AI output and roster data.
 * e.g. "O'Brien" → "obrien", "#42" → "42"
 *
 * @param {string} s
 * @returns {string}
 */
function norm(s) { return (s || "").toLowerCase().replace(/[^a-z0-9]/g, ""); }

// ── Number-only roster detection ──────────────────────────────────────────────
/**
 * Returns true if every row in the roster has no name — i.e. this is a number-only
 * roster (e.g. practice jerseys with no player assignments).
 * The match logic uses a different path for number-only rosters.
 *
 * @param {Object[]} roster
 * @returns {boolean}
 */
function isNumberOnlyRoster(roster) {
  return roster.length > 0 && roster.every(r => !r.name || r.name.trim() === "");
}

// ── Match logic ───────────────────────────────────────────────────────────────
/**
 * Finds roster entries that match the name and/or number detected by the AI.
 * Returns a result object describing the match type so the UI can choose the
 * correct overlay mode (green confirm, yellow pick, or red flag).
 *
 * Design principle: NEVER auto-confirm when ambiguity exists — always route to
 * the pick screen so the operator makes the final call.
 *
 * Match priority order:
 *   1. Both name AND number match exactly → "exact"
 *   2. Multiple entries share the same number → "number_conflict" (pick screen)
 *   3. Only number matched, no name on jersey → "exact" (if unique) or "number_conflict"
 *   4. Only name matched → "exact" (if unique) or "number_conflict"
 *   5. Substring / partial match → "close" (yellow screen)
 *   6. No match → "none" (red screen)
 *
 * For number-only rosters (no player names), a separate simpler path is used.
 *
 * @param {Object[]} roster - Full roster array
 * @param {string} name - Player name detected by AI (may be empty)
 * @param {string} number - Jersey number detected by AI (may be empty)
 * @returns {{ type: "exact"|"number_conflict"|"size_pick"|"close"|"none", match?: Object, candidates?: Object[] }}
 */
function findMatches(roster, name, number) {
  const nName      = norm(name);
  const nNum       = norm(number);
  const numberOnly = isNumberOnlyRoster(roster);

  if (numberOnly) {
    const byNum = nNum ? roster.filter(r => norm(r.number) === nNum) : [];
    if (byNum.length === 1) return { type: "exact",     match: byNum[0] };
    if (byNum.length > 1)   return { type: "size_pick", candidates: byNum };
    const close = nNum.length >= 2
      ? roster.filter(r => norm(r.number).includes(nNum) || nNum.includes(norm(r.number)))
      : [];
    if (close.length > 0) return { type: "close", candidates: close };
    return { type: "none" };
  }

  const byNum  = nNum  ? roster.filter(r => norm(r.number) === nNum)  : [];
  const byName = nName ? roster.filter(r => norm(r.name)   === nName) : [];

  // Entries matching on BOTH name and number (may differ by team/size)
  const byBoth = byNum.filter(r => norm(r.name) === nName);

  if (byBoth.length === 1) return { type: "exact",           match: byBoth[0] };
  if (byBoth.length > 1)   return { type: "number_conflict", candidates: byBoth };

  // Number matched, no name detected
  if (byNum.length >= 1 && nName === "") {
    if (byNum.length === 1) return { type: "exact",           match: byNum[0] };
    return                          { type: "number_conflict", candidates: byNum };
  }

  // Number matched but name didn't agree
  if (byNum.length >= 1) return { type: "number_conflict", candidates: byNum };

  // Name-only match
  if (byName.length === 1) return { type: "exact",           match: byName[0] };
  if (byName.length > 1)   return { type: "number_conflict", candidates: byName };

  // Fuzzy fallback
  const closeMap = new Map();
  [
    ...(nNum.length  >= 2 ? roster.filter(r => norm(r.number).includes(nNum)  || nNum.includes(norm(r.number)))  : []),
    ...(nName.length >= 3 ? roster.filter(r => norm(r.name).includes(nName)   || nName.includes(norm(r.name)))   : []),
  ].forEach(r => closeMap.set(r._id, r));
  const close = [...closeMap.values()];
  if (close.length > 0) return { type: "close", candidates: close };

  return { type: "none" };
}

// ── Session persistence ────────────────────────────────────────────────────────
/**
 * Session state is persisted to localStorage so the operator can reload the page
 * (e.g. after a browser crash) without losing scan progress.
 *
 * Persisted fields: sessionStarted, roster (with scanned status), log, orderNumber,
 *   operatorName, rosterFile, binMap, firstScanTime, packGroupId.
 *
 * NOT persisted: thumbnail images (too large for localStorage), camera state,
 *   overlay state, or any UI-only state.
 *
 * The session is cleared when the operator clicks "New Session".
 */
const SESSION_KEY = "jerseyqc_session";

/** Load session from localStorage. Returns null if nothing is saved or on error. */
function loadSession() {
  try { const r = localStorage.getItem(SESSION_KEY); return r ? JSON.parse(r) : null; }
  catch { return null; }
}

/** Save current session state to localStorage. Silent on storage-full errors. */
function saveSession(data) {
  try { localStorage.setItem(SESSION_KEY, JSON.stringify(data)); }
  catch { /* storage full — session progress not saved this tick */ }
}

/** Remove the saved session from localStorage. Called on "New Session". */
function clearSession() {
  try { localStorage.removeItem(SESSION_KEY); }
  catch { /* storage unavailable */ }
}

// ── Excel and CSV exports ──────────────────────────────────────────────────────
/**
 * Exports the full roster with scan results to an Excel (.xlsx) file and
 * triggers a browser download.
 *
 * The exported file contains:
 *   - Header block: order number, operator name, export date
 *   - Bin assignments section (if binMap is set)
 *   - One row per roster entry with: Status, all roster columns, Bin, Comment
 *
 * Status values in the export:
 *   PASS / FLAGGED / RESOLVED / EXTRA (NOT IN ROSTER) / NOT SCANNED
 *
 * Requires SheetJS to be loaded (window.XLSX must exist).
 *
 * @param {Object[]} roster - Full roster array (includes _extra entries)
 * @param {string} orderNumber
 * @param {string} operatorName
 * @param {Object|null} binMap - Team → bin number map, or null if no bins
 */
function exportRosterXLSX(roster, orderNumber, operatorName, binMap) {
  if (!window.XLSX) { alert("Excel library not loaded yet, please try again."); return; }
  const cols = Object.keys(roster[0]).filter(k => k !== "_id" && k !== "_extra" && k !== "scanned" && k !== "comment");
  const hasBins = binMap && Object.keys(binMap).length > 0;
  const data = [
    ["Order Number", orderNumber  || "—"],
    ["Operator",     operatorName || "—"],
    ["Export Date",  new Date().toLocaleString()],
    [],
  ];
  if (hasBins) {
    data.push(["Bin Assignments"]);
    data.push(["Bin", "Team"]);
    Object.entries(binMap).sort((a, b) => a[1] - b[1]).forEach(([team, bin]) => data.push([`Bin ${bin}`, team]));
    data.push([]);
  }
  data.push(["Status", ...cols.map(c => c.charAt(0).toUpperCase() + c.slice(1)), ...(hasBins ? ["Bin"] : []), "Comment"]);
  roster.forEach(r => {
    const s = r.scanned === "pass"     ? "PASS"
            : r.scanned === "flag"     ? "FLAGGED"
            : r.scanned === "resolved" ? "RESOLVED"
            : r.scanned === "extra"    ? "EXTRA (NOT IN ROSTER)"
            : "NOT SCANNED";
    data.push([s, ...cols.map(c => r[c]), ...(hasBins ? [r.team && binMap[r.team] ? `Bin ${binMap[r.team]}` : ""] : []), r.comment || ""]);
  });
  const ws = window.XLSX.utils.aoa_to_sheet(data);
  const wb = window.XLSX.utils.book_new();
  window.XLSX.utils.book_append_sheet(wb, ws, "QC Results");
  window.XLSX.writeFile(wb, `roster_${orderNumber ? orderNumber + "_" : ""}${new Date().toISOString().slice(0,10)}.xlsx`);
}

/**
 * Exports the scan event log to a CSV file and triggers a browser download.
 *
 * Each row in the log CSV represents one scan event with:
 *   Time, Status, Detected Name, Detected Number, Matched Name, Matched Number,
 *   Team, Size, Bin, Notes (reason or resolution)
 *
 * This is useful for auditing — it records every scan attempt, not just the final roster state.
 *
 * @param {Object[]} log - Array of log entry objects
 * @param {string} orderNumber
 * @param {Object|null} binMap
 */
function exportLogCSV(log, orderNumber, binMap) {
  const header = "Time,Status,Detected Name,Detected Number,Matched Name,Matched Number,Team,Size,Bin,Notes";
  const rows   = log.map(l => [
    l.timestamp, l.status,
    l.detected?.name   || "", l.detected?.number || "",
    l.match?.name      || "", l.match?.number    || "",
    l.match?.team      || "", l.match?.size      || "",
    (binMap && l.match?.team && binMap[l.match.team]) ? `Bin ${binMap[l.match.team]}` : "",
    l.reason || l.resolution || "",
  ].join(","));
  const a = document.createElement("a");
  a.href = URL.createObjectURL(new Blob([[header, ...rows].join("\n")], { type: "text/csv" }));
  a.download = `log_${orderNumber ? orderNumber + "_" : ""}${new Date().toISOString().slice(0,10)}.csv`;
  a.click();
}

/**
 * Converts a KeyboardEvent.code value into a human-readable label for display.
 * e.g. "KeyB" → "B", "Space" → "Space", "ArrowDown" → "↓"
 *
 * @param {string} code - KeyboardEvent.code value
 * @returns {string}
 */
function formatKey(code) {
  const map = {
    ShiftRight: "RShift", ShiftLeft: "LShift",
    ControlRight: "RCtrl", ControlLeft: "LCtrl",
    AltRight: "RAlt", AltLeft: "LAlt",
    Space: "Space", Enter: "Enter", Tab: "Tab",
    ArrowUp: "↑", ArrowDown: "↓", ArrowLeft: "←", ArrowRight: "→",
  };
  if (map[code]) return map[code];
  if (code.startsWith("Key"))   return code.slice(3);
  if (code.startsWith("Digit")) return code.slice(5);
  return code;
}

// ── Shared style objects ───────────────────────────────────────────────────────
// Reusable inline style objects spread into JSX elements throughout the app.
// All styles are inline to keep the component self-contained with no CSS files.

const card     = { background: "#161b22", borderRadius: 10, border: "1px solid #21262d", marginBottom: 10, overflow: "hidden" };
const btnPri   = { padding: "8px 18px", borderRadius: 8, border: "none", background: "#3b82f6", color: "#fff", fontWeight: 700, fontSize: 14, cursor: "pointer" };
const btnGhost = { padding: "6px 12px", borderRadius: 8, border: "1px solid #334155", background: "transparent", color: "#94a3b8", fontWeight: 600, fontSize: 12, cursor: "pointer" };
const btnSm    = { padding: "5px 12px", borderRadius: 6, fontWeight: 600, fontSize: 12, cursor: "pointer" };
const codeSt   = { background: "#1e293b", padding: "1px 6px", borderRadius: 4, fontSize: 12 };
const thSt     = { padding: "6px 8px", textAlign: "left", fontSize: 10, fontWeight: 700, color: "#64748b", textTransform: "uppercase", letterSpacing: 0.5 };
const tdSt     = { padding: "6px 8px", color: "#e2e8f0", fontSize: 12 };

/** Returns the status indicator colour for a given scan status string. */
function statusColor(s) {
  if (s === S_PASS)    return "#22c55e";
  if (s === S_FLAGGED) return "#ef4444";
  if (s === S_MANUAL)  return "#94a3b8";
  if (s === S_EXTRA)   return "#f59e0b";
  return "#64748b";
}

/** Small coloured badge used in the header to show pass/flag counts. */
function Pill({ color, children }) {
  return (
    <span style={{ background: `${color}22`, color, fontSize: 11, fontWeight: 700, padding: "2px 7px", borderRadius: 20, border: `1px solid ${color}44` }}>
      {children}
    </span>
  );
}

/**
 * Large action button used on the full-screen scan overlay.
 * `outline` renders a transparent background (ghost style) instead of a filled button.
 */
function OverlayBtn({ color, outline, onClick, children }) {
  return (
    <button onClick={onClick} style={{ padding: "16px 32px", borderRadius: 14, border: `3px solid ${color}`, background: outline ? "transparent" : color, color: outline ? color : "#fff", fontWeight: 800, fontSize: 22, cursor: "pointer", display: "flex", alignItems: "center", gap: 12 }}>
      {children}
    </button>
  );
}

/**
 * Keyboard key indicator styled like a physical key.
 * `light` uses a lighter style for use on coloured overlay backgrounds.
 */
function Kbd({ light, children }) {
  return (
    <kbd style={{ background: light ? "rgba(255,255,255,0.15)" : "rgba(0,0,0,0.3)", border: "2px solid rgba(255,255,255,0.35)", padding: "3px 10px", borderRadius: 6, fontSize: 18, fontFamily: "monospace", fontWeight: 700 }}>
      {children}
    </kbd>
  );
}

// ── Session start modal ───────────────────────────────────────────────────────
/**
 * Full-screen modal shown before each scanning session.
 *
 * Two modes:
 *   1. Admin-assigned queue: if the admin pushed rosters to this station, they appear
 *      as selectable cards. The operator just picks one and enters their name.
 *   2. Manual upload: operator enters order number + uploads a roster file themselves.
 *
 * The mode switches automatically if assignments arrive after the modal mounts
 * (heartbeat fires ~1s after page load).
 *
 * @param {Function} onStart - Called with { roster, orderNumber, operatorName, rosterFile, packGroupId, binMap, acceptedOrderNumber? }
 * @param {boolean} xlsxReady - Whether SheetJS has loaded (controls whether Excel upload is enabled)
 * @param {Object[]|null} preAssigned - Array of admin-assigned orders waiting for this station
 * @param {Function} onAdmin - Opens the admin panel
 */
function SessionStartModal({ onStart, xlsxReady, preAssigned, onAdmin }) {
  const assignments  = Array.isArray(preAssigned) ? preAssigned : (preAssigned ? [preAssigned] : []);
  const hasQueue     = assignments.length > 0;

  const [selIdx,    setSelIdx]    = useState(() => assignments.length === 1 ? 0 : -1);
  const [operator,  setOperator]  = useState("");
  const [order,     setOrder]     = useState("");
  const [file,      setFile]      = useState(null);
  const [manual,    setManual]    = useState(!hasQueue);

  // If assignments arrive after the modal mounts (heartbeat fires ~1s later),
  // switch from manual mode to the picker automatically
  useEffect(() => {
    if (assignments.length === 0) return;
    setManual(false);
    if (assignments.length === 1 && selIdx === -1) setSelIdx(0);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [assignments.length]);
  const [error,     setError]     = useState(null);
  const [loading,   setLoading]   = useState(false);
  const fileRef = useRef(null);

  const selected   = selIdx >= 0 ? assignments[selIdx] : null;
  const useManual  = manual || !hasQueue;
  const inputStyle = { width: "100%", boxSizing: "border-box", padding: "9px 12px", borderRadius: 8, border: "1px solid #334155", background: "#1e293b", color: "#e2e8f0", fontSize: 14 };
  const lblStyle   = { fontSize: 11, color: "#64748b", fontWeight: 700, textTransform: "uppercase", letterSpacing: 0.5, display: "block", marginBottom: 5 };

  const handleStart = async () => {
    if (!operator.trim()) { setError("Please enter the operator name."); return; }
    if (!useManual) {
      if (!selected) { setError("Please select an order."); return; }
      onStart({ roster: selected.roster, orderNumber: selected.orderNumber, operatorName: operator.trim(), rosterFile: selected.rosterName, packGroupId: selected.packGroupId ?? null, binMap: selected.binMap ?? null, acceptedOrderNumber: selected.orderNumber });
    } else {
      if (!order.trim()) { setError("Please enter an order number."); return; }
      if (!file)         { setError("Please upload a roster file."); return; }
      setLoading(true);
      try {
        const isExcel = /\.(xlsx|xls)$/i.test(file.name);
        const parsed  = isExcel ? await parseExcel(file) : parseCSV(await file.text());
        if (parsed.length === 0) { setError("No valid rows found. Check columns: name, number, team, size."); setLoading(false); return; }
        onStart({ roster: parsed, orderNumber: order.trim(), operatorName: operator.trim(), rosterFile: file.name, packGroupId: null, binMap: null });
      } catch (err) { setError("Failed to parse roster: " + err.message); setLoading(false); }
    }
  };

  return (
    <div style={{ position: "fixed", inset: 0, background: "#0d1117", zIndex: 2000, display: "flex", alignItems: "center", justifyContent: "center", overflowY: "auto", padding: "20px 0" }}>
      <div style={{ background: "#161b22", borderRadius: 16, border: "1px solid #21262d", padding: 40, width: 420, maxWidth: "90vw", position: "relative" }}>
        <img src={akLogo} alt="AK" style={{ position: "absolute", top: 14, left: 16, height: 28, opacity: 0.9 }} />
        <div style={{ fontSize: 32, marginBottom: 8, textAlign: "center" }}>🏭</div>
        <div style={{ fontWeight: 900, fontSize: 22, textAlign: "center", marginBottom: 4 }}>Jersey QC Scanner</div>
        <div style={{ fontSize: 13, color: "#64748b", textAlign: "center", marginBottom: 28 }}>Set up your session to begin</div>

        {/* ── Pre-assigned order picker ── */}
        {hasQueue && !manual && (
          <div style={{ marginBottom: 20 }}>
            <label style={lblStyle}>Select Order ({assignments.length} assigned)</label>
            {assignments.map((a, i) => (
              <div key={i} onClick={() => setSelIdx(i)}
                style={{ padding: "12px 14px", marginBottom: 8, borderRadius: 10, border: `2px solid ${i === selIdx ? "#3b82f6" : "#334155"}`, background: i === selIdx ? "rgba(59,130,246,0.1)" : "#1e293b", cursor: "pointer" }}>
                <div style={{ fontWeight: 700, fontSize: 14, color: "#e2e8f0" }}>{a.orderNumber || "—"}</div>
                <div style={{ fontSize: 12, color: "#64748b", marginTop: 2 }}>{a.rosterName} · {a.roster?.length} players</div>
                {a.packGroupId && <div style={{ fontSize: 11, color: "#60a5fa", marginTop: 2 }}>📦 Pack group assigned</div>}
              </div>
            ))}
            <button onClick={() => setManual(true)}
              style={{ fontSize: 11, color: "#475569", background: "none", border: "none", cursor: "pointer", padding: 0, marginTop: 2 }}>
              ↑ Upload a different roster instead
            </button>
          </div>
        )}

        {/* ── Manual order + file fields ── */}
        {useManual && (
          <>
            {hasQueue && (
              <button onClick={() => setManual(false)}
                style={{ fontSize: 11, color: "#3b82f6", background: "none", border: "none", cursor: "pointer", padding: 0, marginBottom: 16 }}>
                ← Back to assigned orders
              </button>
            )}
            <div style={{ marginBottom: 14 }}>
              <label style={lblStyle}>Order Number</label>
              <input value={order} onChange={e => setOrder(e.target.value)} placeholder="e.g. ORD-2024-001"
                onKeyDown={e => e.key === "Enter" && handleStart()}
                style={inputStyle} />
            </div>
            <div style={{ marginBottom: 22 }}>
              <label style={lblStyle}>Roster File</label>
              <div onClick={() => fileRef.current.click()}
                style={{ padding: "10px 12px", borderRadius: 8, border: `1px solid ${file ? "#22c55e" : "#334155"}`, background: "#1e293b", color: file ? "#22c55e" : "#64748b", fontSize: 13, cursor: "pointer", display: "flex", alignItems: "center", gap: 8 }}>
                <span style={{ fontSize: 16 }}>{file ? "✓" : "📂"}</span>
                <span>{file ? file.name : "Choose CSV or Excel file…"}</span>
              </div>
              <input ref={fileRef} type="file" accept=".csv,.xlsx,.xls,.tsv,.txt" style={{ display: "none" }} onChange={e => { const f = e.target.files[0]; if (f) { setFile(f); setError(null); } e.target.value = ""; }} />
            </div>
          </>
        )}

        <div style={{ marginBottom: 20 }}>
          <label style={lblStyle}>Operator Name</label>
          <input value={operator} onChange={e => setOperator(e.target.value)} placeholder="e.g. John Smith"
            autoFocus onKeyDown={e => e.key === "Enter" && handleStart()}
            style={inputStyle} />
        </div>

        {error && <div style={{ fontSize: 12, color: "#f87171", marginBottom: 12, padding: "8px 12px", background: "rgba(239,68,68,0.08)", borderRadius: 6 }}>{error}</div>}

        <button onClick={handleStart} disabled={loading || (useManual && /\.(xlsx|xls)$/i.test(file?.name || "") && !xlsxReady)}
          style={{ width: "100%", padding: "13px", borderRadius: 10, border: "none", background: "#3b82f6", color: "#fff", fontWeight: 700, fontSize: 16, cursor: loading ? "wait" : "pointer", opacity: loading ? 0.7 : 1 }}>
          {loading ? "Loading…" : "Start Session →"}
        </button>
        <button onClick={onAdmin}
          style={{ width: "100%", marginTop: 10, padding: "10px", borderRadius: 10, border: "1px solid #334155", background: "transparent", color: "#64748b", fontWeight: 600, fontSize: 13, cursor: "pointer" }}>
          🔐 Admin Panel
        </button>
      </div>
    </div>
  );
}

// ── Scan overlay ──────────────────────────────────────────────────────────────
/**
 * Full-screen overlay shown after each scan. Covers the entire viewport.
 *
 * Modes and their colours:
 *   "confirm" (green)  — exact match found; operator presses confirmKey to accept
 *   "done"    (green)  — brief flash after confirmation; auto-dismisses after 700ms and triggers next scan
 *   "pick"    (yellow) — multiple candidates match; operator selects the correct one
 *   "close"   (yellow) — fuzzy/partial match; operator confirms or flags
 *   "flag"    (red)    — no match found, API error, or operator flagged the jersey
 *
 * Keyboard handling (all keys are configurable):
 *   confirmKey       — confirm / select highlighted candidate / flag and continue
 *   cancelKey        — retry (go back to camera) / navigate to next candidate
 *   double cancelKey — flag as bad jersey (on pick/close screens)
 *   ArrowUp/Down     — navigate candidate list on pick/close screens
 *   Enter            — on close screen: mark as bad scan, correct jersey
 *
 * @param {{ mode: string, scan: Object, candidates?: Object[], reason?: string, flagTitle?: string, canAddExtra?: boolean }} state
 * @param {Function} onConfirm - Confirms the current match
 * @param {Function} onEdit - Dismisses overlay back to camera
 * @param {Function} onPickCandidate - Selects a specific candidate from the list
 * @param {Function} onFlagBadScan - Flags the jersey as a problem
 * @param {Function} onFlagBadJersey - Close-match: scanner misread, but jersey is correct
 * @param {Function} onAddExtra - Adds jersey as an extra (not in roster)
 * @param {Function} onDismiss - Closes the overlay
 * @param {Function} onAutoScan - Triggers the next scan after "done" auto-dismiss
 * @param {Object|null} binMap - Team→bin map for displaying bin number on green screen
 * @param {string} confirmKey - KeyboardEvent.code for the confirm action
 * @param {string} cancelKey - KeyboardEvent.code for the cancel/navigate action
 */
function ScanOverlay({ state, onConfirm, onEdit, onPickCandidate, onFlagBadScan, onFlagBadJersey, onAddExtra, onDismiss, onAutoScan, binMap, confirmKey, cancelKey }) {
  const [selectedIdx,     setSelectedIdx]     = useState(0);
  const [lastCancelPress, setLastCancelPress] = useState(0);

  const isConfirm = state.mode === "confirm";
  const isPick    = state.mode === "pick";
  const isClose   = state.mode === "close";
  const isFlag    = state.mode === "flag";
  const isDone    = state.mode === "done";

  const bg     = (isConfirm || isDone) ? "#052e16" : (isPick || isClose) ? "#1c1a06" : "#450a0a";
  const accent = (isConfirm || isDone) ? "#22c55e" : (isPick || isClose) ? "#f59e0b" : "#ef4444";

  const match   = state.scan?.match;
  const hasSize = !!(match?.size && match.size.trim() !== "");
  const binNum  = (binMap && match?.team && binMap[match.team]) ? binMap[match.team] : null;

  // Default to first unscanned candidate when mode/candidates change
  useEffect(() => {
    if (state.candidates) {
      const firstUnscanned = state.candidates.findIndex(c => !c.scanned || c.scanned === false);
      setSelectedIdx(firstUnscanned >= 0 ? firstUnscanned : 0); // eslint-disable-line react-hooks/set-state-in-effect
    } else {
      setSelectedIdx(0);
    }
  }, [state.mode, state.candidates]);

  // Auto-dismiss done flash then fire next scan
  useEffect(() => {
    if (!isDone) return;
    const t = setTimeout(() => { onDismiss(); onAutoScan(); }, 700);
    return () => clearTimeout(t);
  }, [isDone, onDismiss, onAutoScan]);

  // Keyboard handler
  useEffect(() => {
    const candidates = state.candidates || [];
    const onKey = (e) => {
      if (e.code === confirmKey) {
        e.preventDefault();
        if (isConfirm) { onConfirm(); return; }
        if (isPick || isClose) {
          const candidate = candidates[selectedIdx];
          if (candidate) onPickCandidate(candidate);
          return;
        }
        if (isFlag) { onFlagBadScan(); return; }
      }
      if (e.code === cancelKey) {
        e.preventDefault();
        if (isConfirm || isFlag) { onEdit(); return; }
        if (isPick || isClose) {
          const now = Date.now();
          if (now - lastCancelPress < 400) {
            // Double-tap: flag bad jersey
            onFlagBadScan();
          } else {
            // Single tap: advance to next candidate
            const total = candidates.length;
            if (total > 0) setSelectedIdx((selectedIdx + 1) % total);
          }
          setLastCancelPress(now);
          return;
        }
      }
      if (e.code === "Enter" && isClose) {
        e.preventDefault();
        const candidate = candidates[selectedIdx];
        if (candidate) onFlagBadJersey(candidate);
        return;
      }
      if ((isPick || isClose) && (e.code === "ArrowDown" || e.code === "ArrowUp")) {
        e.preventDefault();
        const total = candidates.length;
        if (total > 0) {
          const dir = e.code === "ArrowDown" ? 1 : -1;
          setSelectedIdx((selectedIdx + dir + total) % total);
        }
      }
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [isConfirm, isPick, isClose, isFlag, isDone, selectedIdx, lastCancelPress, state.candidates, state.mode, confirmKey, cancelKey, onConfirm, onEdit, onPickCandidate, onFlagBadScan, onFlagBadJersey]);

  return (
    <div style={{ position: "fixed", inset: 0, zIndex: 1000, background: bg, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", userSelect: "none", padding: "20px 0", overflowY: "auto" }}>
      <img src={akLogo} alt="AK" style={{ position: "absolute", top: 16, left: 20, height: 32, opacity: 0.7 }} />

      {/* ── GREEN ── */}
      {(isConfirm || isDone) && match && (
        <>
          {isConfirm && match.scanned && match.scanned !== false && (
            <div style={{ background: "rgba(245,158,11,0.18)", border: "2px solid #f59e0b", borderRadius: 12, padding: "10px 28px", marginBottom: 12, fontSize: 22, fontWeight: 800, color: "#fcd34d", textAlign: "center" }}>
              ⚠ Already scanned — confirm to flag as double scan or extra
            </div>
          )}
          <div style={{ fontSize: 120, lineHeight: 1, marginBottom: 4 }}>✅</div>
          <div style={{ fontSize: 96, fontWeight: 900, color: accent, letterSpacing: "-4px", textAlign: "center", padding: "0 24px", lineHeight: 1, marginTop: 4 }}>
            #{match.number}{match.name ? ` · ${match.name}` : ""}
          </div>
          {match.team && (
            <div style={{ fontSize: 52, color: "#86efac", marginTop: 10, fontWeight: 700, textAlign: "center" }}>{match.team}</div>
          )}
          <div style={{ display: "flex", gap: 20, marginTop: 20, flexWrap: "wrap", justifyContent: "center" }}>
            {hasSize && (
              <div style={{ padding: "18px 56px", borderRadius: 20, border: "4px solid #22c55e", background: "rgba(34,197,94,0.12)", textAlign: "center", minWidth: 200 }}>
                <div style={{ fontSize: 22, color: "#86efac", fontWeight: 700, textTransform: "uppercase", letterSpacing: 3, marginBottom: 4 }}>Size</div>
                <div style={{ fontSize: 108, fontWeight: 900, color: "#fff", letterSpacing: 6, lineHeight: 1 }}>{match.size}</div>
              </div>
            )}
            {binNum !== null && (
              <div style={{ padding: "18px 56px", borderRadius: 20, border: "4px solid #3b82f6", background: "rgba(59,130,246,0.12)", textAlign: "center", minWidth: 200 }}>
                <div style={{ fontSize: 22, color: "#93c5fd", fontWeight: 700, textTransform: "uppercase", letterSpacing: 3, marginBottom: 4 }}>Bin</div>
                <div style={{ fontSize: 108, fontWeight: 900, color: "#fff", lineHeight: 1 }}>{binNum}</div>
              </div>
            )}
          </div>
          {isConfirm && (
            <div style={{ marginTop: 32, display: "flex", gap: 24 }}>
              <OverlayBtn color="#22c55e" onClick={onConfirm}>✓ Confirm <Kbd>{formatKey(confirmKey)}</Kbd></OverlayBtn>
              <OverlayBtn color="#94a3b8" outline onClick={onEdit}>✎ Edit <Kbd light>{formatKey(cancelKey)}</Kbd></OverlayBtn>
            </div>
          )}
          {isDone && (
            <div style={{ fontSize: 34, color: "#4ade80", marginTop: 24, opacity: 0.9, fontWeight: 700 }}>Scanning next jersey…</div>
          )}
        </>
      )}

      {/* ── YELLOW ── */}
      {(isPick || isClose) && (
        <>
          <div style={{ fontSize: 80, lineHeight: 1, marginBottom: 4 }}>{isPick ? "⚠️" : "🔍"}</div>
          <div style={{ fontSize: 72, fontWeight: 900, color: accent, marginBottom: 10, textAlign: "center", padding: "0 24px", lineHeight: 1.1 }}>
            {isPick ? "Select correct player" : "Close match — confirm or flag"}
          </div>
          <div style={{ fontSize: 40, color: "#fcd34d", marginBottom: 20, fontWeight: 700, textAlign: "center" }}>
            Detected: #{state.scan?.detected?.number || "?"}{state.scan?.detected?.name ? ` · ${state.scan.detected.name}` : ""}
          </div>
          <div style={{ width: "100%", maxWidth: 900, padding: "0 24px", boxSizing: "border-box", maxHeight: "45vh", overflowY: "auto" }}>
            {state.candidates?.map((c, i) => {
              const isScanned  = c.scanned && c.scanned !== false;
              const isSelected = i === selectedIdx;
              return (
                <div key={c._id}
                  onClick={() => onPickCandidate(c)}
                  onMouseEnter={() => setSelectedIdx(i)}
                  style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "22px 32px", marginBottom: 14, borderRadius: 16, border: `4px solid ${isScanned ? (isSelected ? "#f59e0b" : "rgba(245,158,11,0.3)") : isSelected ? accent : "rgba(255,255,255,0.2)"}`, background: isScanned ? (isSelected ? "rgba(245,158,11,0.15)" : "rgba(245,158,11,0.05)") : isSelected ? "rgba(245,158,11,0.22)" : "rgba(255,255,255,0.06)", cursor: "pointer", opacity: isScanned ? 0.72 : 1 }}>
                  <div>
                    <div style={{ fontSize: 48, fontWeight: 900, color: "#fff", lineHeight: 1.1 }}>#{c.number}{c.name ? ` · ${c.name}` : ""}</div>
                    <div style={{ fontSize: 28, color: isSelected ? "#fcd34d" : "#94a3b8", marginTop: 4, fontWeight: 600 }}>{[c.team, c.size].filter(Boolean).join(" · ")}</div>
                    {isScanned && <div style={{ fontSize: 20, color: "#f59e0b", marginTop: 4, fontWeight: 700 }}>⚠ Already scanned — tap to flag as double/extra</div>}
                  </div>
                  {isSelected && <span style={{ fontSize: 32, color: isScanned ? "#f59e0b" : accent, fontWeight: 900 }}>← {formatKey(confirmKey)}</span>}
                </div>
              );
            })}
          </div>
          <div style={{ marginTop: 20, display: "flex", gap: 16, flexWrap: "wrap", justifyContent: "center" }}>
            {isClose && (
              <OverlayBtn color="#3b82f6" outline onClick={() => { const c = state.candidates?.[selectedIdx]; if (c) onFlagBadJersey(c); }}>
                ✓ Bad scan, correct jersey <Kbd light>Enter</Kbd>
              </OverlayBtn>
            )}
            <OverlayBtn color="#ef4444" outline onClick={onFlagBadScan}>
              🚩 Bad jersey <Kbd light>double {formatKey(cancelKey)}</Kbd>
            </OverlayBtn>
          </div>
          <div style={{ marginTop: 14, fontSize: 18, color: "rgba(255,255,255,0.6)", textAlign: "center", fontWeight: 600 }}>
            {formatKey(confirmKey)} select · {formatKey(cancelKey)} navigate · double {formatKey(cancelKey)} flag bad jersey
          </div>
        </>
      )}

      {/* ── RED ── */}
      {isFlag && (
        <>
          <div style={{ fontSize: 100, lineHeight: 1, marginBottom: 4 }}>{state.flagTitle ? "🔄" : "🚩"}</div>
          <div style={{ fontSize: 88, fontWeight: 900, color: accent, textAlign: "center", padding: "0 24px", lineHeight: 1, marginTop: 4 }}>{state.flagTitle || "NOT FOUND"}</div>
          {state.scan?.detected && (
            <div style={{ fontSize: 42, color: "#fca5a5", marginTop: 12, textAlign: "center" }}>
              #{state.scan.detected.number || "?"}{state.scan.detected.name ? ` · ${state.scan.detected.name}` : ""}
            </div>
          )}
          {state.reason && (
            <div style={{ fontSize: 24, color: "#fca5a5", marginTop: 10, textAlign: "center", padding: "0 40px", maxWidth: 700, opacity: 0.85, lineHeight: 1.4 }}>
              {state.reason}
            </div>
          )}
          <div style={{ marginTop: 36, display: "flex", gap: 20, flexWrap: "wrap", justifyContent: "center" }}>
            <OverlayBtn color="#ef4444" onClick={onFlagBadScan}>
              {state.flagTitle ? "🔄 Double scan — same jersey" : "🚩 Flag & continue"} <Kbd>{formatKey(confirmKey)}</Kbd>
            </OverlayBtn>
            {state.canAddExtra && (
              <OverlayBtn color="#f59e0b" onClick={onAddExtra}>➕ Extra jersey — not in roster</OverlayBtn>
            )}
            <OverlayBtn color="#94a3b8" outline onClick={onEdit}>↩ Retry <Kbd light>{formatKey(cancelKey)}</Kbd></OverlayBtn>
          </div>
        </>
      )}

      {/* Progress bar */}
      {isDone && (
        <div style={{ position: "absolute", bottom: 0, left: 0, height: 6, width: "100%", background: "#052e16" }}>
          <div
            ref={el => { if (el) { el.style.transition = "width 0.7s linear"; requestAnimationFrame(() => { el.style.width = "0%"; }); } }}
            style={{ height: "100%", background: "#22c55e", width: "100%" }}
          />
        </div>
      )}

      {/* Thumbnail */}
      {state.scan?.thumb && !isDone && (
        <div style={{ position: "absolute", bottom: 12, right: 12, opacity: 0.3 }}>
          <img src={state.scan.thumb} alt="" style={{ width: 80, height: 60, objectFit: "cover", borderRadius: 6 }} />
        </div>
      )}
    </div>
  );
}

// ── Result card ───────────────────────────────────────────────────────────────
/**
 * Shows a flagged or extra item below the camera with resolution actions.
 * Only rendered for the most recent non-pass result while no overlay is showing.
 *
 * Flagged items show three resolution buttons:
 *   ✓ Accept  — marks the jersey as manually verified (passes QC)
 *   ✗ Reject  — marks the jersey as sent back for rework
 *   📝 Note   — records the text in the note field as the resolution
 *
 * @param {{ id, status, detected, match?, reason?, resolution?, thumb? }} result
 * @param {Function} onResolve - Called with (logId, resolutionText)
 */
function ResultCard({ result, onResolve }) {
  const [note, setNote] = useState("");
  const isFlagged = result.status === S_FLAGGED;
  const isExtra   = result.status === S_EXTRA;
  const bgColor   = isFlagged ? "rgba(239,68,68,0.06)" : isExtra ? "rgba(245,158,11,0.06)" : "rgba(148,163,184,0.06)";
  const label     = isFlagged ? "🚩 Flagged" : isExtra ? "➕ Extra jersey" : "🔧 Resolved";
  return (
    <div style={{ ...card, border: `1px solid ${statusColor(result.status)}`, background: bgColor, padding: "12px 14px" }}>
      <div style={{ display: "flex", gap: 12, alignItems: "flex-start" }}>
        {result.thumb && <img src={result.thumb} alt="" style={{ width: 60, height: 45, objectFit: "cover", borderRadius: 5, flexShrink: 0 }} />}
        <div style={{ flex: 1 }}>
          <div style={{ fontSize: 11, fontWeight: 700, color: statusColor(result.status), textTransform: "uppercase", letterSpacing: 1 }}>
            {label}
          </div>
          <div style={{ fontSize: 15, fontWeight: 700, marginTop: 2 }}>
            #{result.detected?.number || "?"} · {result.detected?.name || "Not detected"}
          </div>
          {result.reason     && <div style={{ fontSize: 12, color: isExtra ? "#fcd34d" : "#f87171", marginTop: 2 }}>{result.reason}</div>}
          {result.resolution && <div style={{ fontSize: 11, color: "#94a3b8", marginTop: 2 }}>Resolution: {result.resolution}</div>}
        </div>
      </div>
      {isFlagged && (
        <div style={{ marginTop: 10, paddingTop: 10, borderTop: "1px solid #21262d" }}>
          <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
            <button onClick={() => onResolve(result.id, "Passed — manually verified")}      style={{ ...btnSm, background: "rgba(34,197,94,0.15)",  color: "#22c55e", border: "1px solid #22c55e" }}>✓ Accept</button>
            <button onClick={() => onResolve(result.id, "Rejected — sent back for rework")} style={{ ...btnSm, background: "rgba(239,68,68,0.15)",  color: "#ef4444", border: "1px solid #ef4444" }}>✗ Reject</button>
            <button onClick={() => onResolve(result.id, note || "No note provided")}        style={{ ...btnSm, background: "rgba(99,102,241,0.15)", color: "#818cf8", border: "1px solid #818cf8" }}>📝 Note</button>
          </div>
          <input value={note} onChange={e => setNote(e.target.value)} placeholder="Optional note…"
            style={{ marginTop: 6, width: "100%", boxSizing: "border-box", padding: "5px 9px", borderRadius: 6, border: "1px solid #334155", background: "#1e293b", color: "#e2e8f0", fontSize: 12 }} />
        </div>
      )}
    </div>
  );
}

// ── Settings panel ────────────────────────────────────────────────────────────
/**
 * Floating panel for customising keyboard shortcuts.
 * Three keys are configurable:
 *   Scan trigger   — takes a photo (default: B)
 *   Confirm/select — confirms match or picks candidate (default: B)
 *   Navigate/retry — cycles candidates, retries scan (default: A)
 *
 * To change a key, click "Change" then press the desired key.
 * Changes are applied only when "Save" is clicked.
 * Escape closes without saving.
 *
 * Designed for use with a barcode scanner gun or foot pedal mapped to keyboard keys.
 */
function SettingsPanel({ scanKey, confirmKey, cancelKey, onScanKeyChange, onConfirmKeyChange, onCancelKeyChange, onClose }) {
  const [listening,   setListening]   = useState(null); // "scan" | "confirm" | "cancel" | null
  const [tempScan,    setTempScan]    = useState(scanKey);
  const [tempConfirm, setTempConfirm] = useState(confirmKey);
  const [tempCancel,  setTempCancel]  = useState(cancelKey);

  useEffect(() => {
    if (!listening) return;
    const h = e => {
      e.preventDefault();
      if (listening === "scan")    setTempScan(e.code);
      if (listening === "confirm") setTempConfirm(e.code);
      if (listening === "cancel")  setTempCancel(e.code);
      setListening(null);
    };
    window.addEventListener("keydown", h);
    return () => window.removeEventListener("keydown", h);
  }, [listening]);

  useEffect(() => {
    const h = e => { if (e.key === "Escape" && !listening) onClose(); };
    window.addEventListener("keydown", h);
    return () => window.removeEventListener("keydown", h);
  }, [onClose, listening]);

  const rows = [
    { label: "Scan trigger",     key: "scan",    temp: tempScan },
    { label: "Confirm / select", key: "confirm", temp: tempConfirm },
    { label: "Navigate / retry", key: "cancel",  temp: tempCancel },
  ];

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.7)", zIndex: 500, display: "flex", alignItems: "center", justifyContent: "center" }}>
      <div style={{ background: "#161b22", borderRadius: 12, border: "1px solid #21262d", padding: 24, width: 360, maxWidth: "90vw" }}>
        <div style={{ fontWeight: 800, fontSize: 15, marginBottom: 14 }}>⚙ Settings</div>
        <div style={{ marginBottom: 18 }}>
          <div style={{ fontSize: 11, color: "#64748b", fontWeight: 700, textTransform: "uppercase", letterSpacing: 0.5, marginBottom: 10 }}>Keyboard shortcuts</div>
          {rows.map(r => (
            <div key={r.key} style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8 }}>
              <div style={{ fontSize: 12, color: "#94a3b8", width: 130, flexShrink: 0 }}>{r.label}</div>
              <div style={{ flex: 1, padding: "6px 10px", borderRadius: 8, border: `1px solid ${listening === r.key ? "#3b82f6" : "#334155"}`, background: "#1e293b", color: "#e2e8f0", fontSize: 13, textAlign: "center" }}>
                {listening === r.key ? "Press any key…" : formatKey(r.temp)}
              </div>
              <button onClick={() => setListening(r.key)} style={{ ...btnGhost, whiteSpace: "nowrap", padding: "4px 10px" }}>Change</button>
            </div>
          ))}
          <div style={{ fontSize: 11, color: "#475569", marginTop: 8 }}>Double-tap Navigate to flag bad jersey on the selection screen.</div>
        </div>
        <div style={{ display: "flex", gap: 8 }}>
          <button onClick={() => { onScanKeyChange(tempScan); onConfirmKeyChange(tempConfirm); onCancelKeyChange(tempCancel); onClose(); }} style={{ ...btnPri, flex: 1 }}>Save</button>
          <button onClick={onClose} style={btnGhost}>Cancel</button>
        </div>
      </div>
    </div>
  );
}

// ── Bin setup modal ───────────────────────────────────────────────────────────
/**
 * Pre-scan modal shown once when a roster has multiple teams.
 * Displays the bin→team assignments so the operator can physically label their bins
 * before scanning starts. No interactivity beyond the confirm button.
 *
 * If an admin-assigned pack group is used, the bin map comes from the pack group
 * definition (not auto-detected from the roster) and this modal is skipped.
 *
 * @param {{ [teamName: string]: number }} binMap
 * @param {Function} onConfirm - Called when the operator is ready to start scanning
 */
function BinSetupModal({ binMap, onConfirm }) {
  const entries = Object.entries(binMap).sort((a, b) => a[1] - b[1]);
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.92)", zIndex: 1500, display: "flex", alignItems: "center", justifyContent: "center" }}>
      <div style={{ background: "#161b22", borderRadius: 16, border: "2px solid #3b82f6", padding: 40, maxWidth: 560, width: "90vw", textAlign: "center", position: "relative" }}>
        <img src={akLogo} alt="AK" style={{ position: "absolute", top: 14, left: 16, height: 28, opacity: 0.9 }} />
        <div style={{ fontSize: 48, marginBottom: 12 }}>🗂️</div>
        <div style={{ fontSize: 26, fontWeight: 900, color: "#60a5fa", marginBottom: 8 }}>Prepare {entries.length} Bins</div>
        <div style={{ fontSize: 15, color: "#94a3b8", marginBottom: 24 }}>
          This order has <strong style={{ color: "#e2e8f0" }}>{entries.length} teams</strong>. Label your bins before scanning:
        </div>
        <div style={{ display: "flex", flexDirection: "column", gap: 8, marginBottom: 28, maxHeight: 320, overflowY: "auto" }}>
          {entries.map(([team, bin]) => (
            <div key={team} style={{ display: "flex", alignItems: "center", gap: 14, padding: "12px 18px", borderRadius: 10, background: "#1e293b", border: "1px solid #334155" }}>
              <div style={{ fontSize: 24, fontWeight: 900, color: "#60a5fa", minWidth: 72, textAlign: "center" }}>Bin {bin}</div>
              <div style={{ fontSize: 16, fontWeight: 700, color: "#e2e8f0", textAlign: "left" }}>{team}</div>
            </div>
          ))}
        </div>
        <button onClick={onConfirm}
          style={{ padding: "13px 36px", borderRadius: 10, border: "none", background: "#3b82f6", color: "#fff", fontWeight: 700, fontSize: 16, cursor: "pointer" }}>
          Bins ready — Start Scanning
        </button>
      </div>
    </div>
  );
}

// ── Roster complete modal ─────────────────────────────────────────────────────
/**
 * Celebration modal shown when all items in the roster have been scanned.
 * Displays a summary (passed / flagged / resolved / extra counts) and offers:
 *   ⬇ Export to Excel  — downloads roster results + saves to server
 *   → Next Order       — clears session and returns to session start
 *   Review first       — dismisses the modal to review flagged items before exporting
 *
 * @param {Object[]} roster
 * @param {number} flagCount - Number of unresolved flagged items (shown as a warning)
 * @param {Function} onExport
 * @param {Function} onDismiss
 * @param {Function} onNewOrder
 */
function RosterCompleteModal({ roster, flagCount, onExport, onDismiss, onNewOrder }) {
  const extraCount = roster.filter(r => r._extra).length;
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.85)", zIndex: 900, display: "flex", alignItems: "center", justifyContent: "center" }}>
      <div style={{ background: "#161b22", borderRadius: 16, border: "2px solid #22c55e", padding: 40, maxWidth: 480, width: "90vw", textAlign: "center", position: "relative" }}>
        <img src={akLogo} alt="AK" style={{ position: "absolute", top: 14, left: 16, height: 28, opacity: 0.9 }} />
        <div style={{ fontSize: 56, marginBottom: 12 }}>🎉</div>
        <div style={{ fontSize: 28, fontWeight: 900, color: "#22c55e", marginBottom: 8 }}>Order Complete!</div>
        <div style={{ fontSize: 15, color: "#94a3b8", marginBottom: 6 }}>All {roster.filter(r => !r._extra).length} jerseys have been scanned.</div>
        {flagCount > 0 && (
          <div style={{ fontSize: 14, color: "#f87171", marginBottom: 6 }}>⚠ {flagCount} flagged item{flagCount > 1 ? "s" : ""} require attention.</div>
        )}
        {extraCount > 0 && (
          <div style={{ fontSize: 14, color: "#f59e0b", marginBottom: 6 }}>➕ {extraCount} extra jersey{extraCount > 1 ? "s" : ""} not in roster.</div>
        )}
        <div style={{ fontSize: 13, color: "#64748b", marginBottom: 28 }}>Please review the roster before exporting.</div>
        <div style={{ display: "flex", justifyContent: "center", gap: 24, marginBottom: 28 }}>
          {[
            { label: "Passed",   value: roster.filter(r => r.scanned === "pass").length,     color: "#22c55e" },
            { label: "Flagged",  value: roster.filter(r => r.scanned === "flag").length,     color: "#ef4444" },
            { label: "Resolved", value: roster.filter(r => r.scanned === "resolved").length, color: "#94a3b8" },
            ...(extraCount > 0 ? [{ label: "Extra", value: extraCount, color: "#f59e0b" }] : []),
          ].map(s => (
            <div key={s.label}>
              <div style={{ fontSize: 32, fontWeight: 900, color: s.color }}>{s.value}</div>
              <div style={{ fontSize: 11, color: "#64748b", textTransform: "uppercase", letterSpacing: 0.5 }}>{s.label}</div>
            </div>
          ))}
        </div>
        <div style={{ display: "flex", gap: 10, justifyContent: "center", flexWrap: "wrap" }}>
          <button onClick={onExport}   style={{ ...btnPri, fontSize: 15, padding: "12px 28px", background: "#22c55e" }}>⬇ Export to Excel</button>
          <button onClick={onNewOrder} style={{ ...btnPri, fontSize: 14, padding: "12px 20px", background: "#3b82f6" }}>→ Next Order</button>
          <button onClick={onDismiss}  style={{ ...btnGhost, fontSize: 14, padding: "12px 20px" }}>Review first</button>
        </div>
      </div>
    </div>
  );
}

// ── Station setup modal ───────────────────────────────────────────────────────
/**
 * One-time setup modal shown the first time the app is opened on a device.
 * Prompts for a station name (e.g. "Station 1") and saves it to localStorage.
 * The station name is used to identify this device in the admin panel and
 * is sent with every heartbeat so the admin can see which station is which.
 *
 * Once set, this modal is never shown again on this device unless localStorage is cleared.
 *
 * @param {Function} onSave - Called with the station name string
 */
function StationSetupModal({ onSave }) {
  const [name, setName] = useState("");
  return (
    <div style={{ position: "fixed", inset: 0, background: "#0d1117", zIndex: 3000, display: "flex", alignItems: "center", justifyContent: "center" }}>
      <div style={{ background: "#161b22", borderRadius: 16, border: "1px solid #21262d", padding: 40, width: 380, maxWidth: "90vw", position: "relative" }}>
        <img src={akLogo} alt="AK" style={{ position: "absolute", top: 14, left: 16, height: 28, opacity: 0.9 }} />
        <div style={{ fontWeight: 900, fontSize: 20, textAlign: "center", marginBottom: 6, marginTop: 8 }}>Set Station Name</div>
        <div style={{ fontSize: 13, color: "#64748b", textAlign: "center", marginBottom: 24 }}>This only needs to be done once on this device.</div>
        <input value={name} onChange={e => setName(e.target.value)} placeholder="e.g. Station 1"
          autoFocus onKeyDown={e => e.key === "Enter" && name.trim() && onSave(name.trim())}
          style={{ width: "100%", boxSizing: "border-box", padding: "9px 12px", borderRadius: 8, border: "1px solid #334155", background: "#1e293b", color: "#e2e8f0", fontSize: 14, marginBottom: 16 }} />
        <button onClick={() => name.trim() && onSave(name.trim())}
          style={{ width: "100%", padding: 13, borderRadius: 10, border: "none", background: "#3b82f6", color: "#fff", fontWeight: 700, fontSize: 16, cursor: "pointer" }}>
          Save &amp; Continue →
        </button>
      </div>
    </div>
  );
}

// ── Admin panel ───────────────────────────────────────────────────────────────
/**
 * Password-protected supervisor panel. Opens as a full-screen overlay.
 *
 * Tabs:
 *   Dashboard    — Live station cards showing each station's name, operator,
 *                  order progress, and online/offline status. Auto-refreshes every 5s.
 *   Assign Roster — Push a roster file + order number to a specific station's queue.
 *                   The station operator will see it appear on their start screen.
 *   Pack Groups   — Create and manage pack groups (multi-order combined exports).
 *   Export History — Browse and re-download all completed order exports.
 *
 * The password is checked against the server's ADMIN_PASSWORD on every API call
 * via the X-Admin-Password header.
 *
 * @param {Function} onClose
 */
function AdminPanel({ onClose }) {
  const [authed,      setAuthed]      = useState(false);
  const [password,    setPassword]    = useState("");
  const [authError,   setAuthError]   = useState(null);
  const [tab,         setTab]         = useState("dashboard");
  const [stations,    setStations]    = useState([]);
  const [exports,     setExports]     = useState([]);
  const [packGroups,  setPackGroups]  = useState([]);
  const [msg,         setMsg]         = useState(null);

  const ah = { 'Content-Type': 'application/json', 'x-admin-password': password };

  const login = async () => {
    try {
      const r = await fetch('/api/admin/login', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ password }) });
      if (r.ok) { setAuthed(true); setAuthError(null); loadAll(password); }
      else setAuthError("Incorrect password.");
    } catch { setAuthError("Server unreachable — is it running?"); }
  };

  const api = async (path, opts = {}) => {
    const r = await fetch(path, { ...opts, headers: { ...ah, ...(opts.headers || {}) } });
    if (!r.ok) throw new Error(await r.text());
    return r.json();
  };

  const loadAll = async (pw) => {
    const h = { 'x-admin-password': pw || password };
    try {
      const [st, ex, pg] = await Promise.all([
        fetch('/api/admin/stations',   { headers: h }).then(r => r.json()),
        fetch('/api/admin/exports',    { headers: h }).then(r => r.json()),
        fetch('/api/admin/packgroups', { headers: h }).then(r => r.json()),
      ]);
      setStations(st); setExports(ex); setPackGroups(pg);
    } catch (e) { console.warn('loadAll failed', e.message); }
  };

  useEffect(() => {
    if (!authed) return;
    const t = setInterval(() => loadAll(), 5000);
    return () => clearInterval(t);
  }, [authed]); // eslint-disable-line react-hooks/exhaustive-deps

  const flash = (m) => { setMsg(m); setTimeout(() => setMsg(null), 3000); };

  if (!authed) return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.92)", zIndex: 2500, display: "flex", alignItems: "center", justifyContent: "center" }}>
      <div style={{ background: "#161b22", borderRadius: 16, border: "1px solid #21262d", padding: 40, width: 360, maxWidth: "90vw", position: "relative" }}>
        <img src={akLogo} alt="AK" style={{ position: "absolute", top: 14, left: 16, height: 28, opacity: 0.9 }} />
        <button onClick={onClose} style={{ position: "absolute", top: 14, right: 14, background: "none", border: "none", color: "#64748b", fontSize: 20, cursor: "pointer" }}>✕</button>
        <div style={{ fontWeight: 900, fontSize: 20, textAlign: "center", marginBottom: 6, marginTop: 8 }}>Admin Login</div>
        <div style={{ fontSize: 13, color: "#64748b", textAlign: "center", marginBottom: 24 }}>Jersey QC Supervisor Panel</div>
        <input type="password" value={password} onChange={e => setPassword(e.target.value)} placeholder="Password"
          autoFocus onKeyDown={e => e.key === "Enter" && login()}
          style={{ width: "100%", boxSizing: "border-box", padding: "9px 12px", borderRadius: 8, border: `1px solid ${authError ? "#ef4444" : "#334155"}`, background: "#1e293b", color: "#e2e8f0", fontSize: 14, marginBottom: 10 }} />
        {authError && <div style={{ fontSize: 12, color: "#f87171", marginBottom: 10 }}>{authError}</div>}
        <button onClick={login} style={{ width: "100%", padding: 13, borderRadius: 10, border: "none", background: "#3b82f6", color: "#fff", fontWeight: 700, fontSize: 16, cursor: "pointer" }}>
          Login →
        </button>
      </div>
    </div>
  );

  const tabStyle = (t) => ({ flex: 1, padding: "9px 0", border: "none", background: "transparent", color: tab === t ? "#3b82f6" : "#64748b", fontWeight: 600, fontSize: 13, borderBottom: tab === t ? "2px solid #3b82f6" : "2px solid transparent", cursor: "pointer" });

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.95)", zIndex: 2500, display: "flex", flexDirection: "column" }}>
      {/* Admin header */}
      <div style={{ background: "#161b22", borderBottom: "1px solid #21262d", padding: "10px 20px", display: "flex", alignItems: "center", gap: 16, flexShrink: 0 }}>
        <img src={akLogo} alt="AK" style={{ height: 28 }} />
        <div style={{ fontWeight: 800, fontSize: 15 }}>Admin Panel</div>
        {msg && <div style={{ fontSize: 12, color: "#22c55e", fontWeight: 600 }}>{msg}</div>}
        <button onClick={onClose} style={{ marginLeft: "auto", background: "none", border: "1px solid #334155", color: "#94a3b8", padding: "4px 14px", borderRadius: 6, cursor: "pointer", fontSize: 12 }}>✕ Close</button>
      </div>
      {/* Tabs */}
      <div style={{ display: "flex", borderBottom: "1px solid #21262d", background: "#161b22", flexShrink: 0 }}>
        {[["dashboard","Dashboard"], ["assign","Assign Roster"], ["packgroups","Pack Groups"], ["exports","Export History"]].map(([t, l]) => (
          <button key={t} onClick={() => { setTab(t); loadAll(); }} style={tabStyle(t)}>{l}</button>
        ))}
      </div>
      {/* Body */}
      <div style={{ flex: 1, overflowY: "auto", padding: 20 }}>

        {/* ── Dashboard ── */}
        {tab === "dashboard" && (
          <AdminDashboard stations={stations} packGroups={packGroups} onRefresh={loadAll} />
        )}

        {/* ── Assign Roster ── */}
        {tab === "assign" && (
          <AdminAssign stations={stations} packGroups={packGroups} api={api} onDone={() => { loadAll(); flash("Roster assigned!"); }} />
        )}

        {/* ── Pack Groups ── */}
        {tab === "packgroups" && (
          <AdminPackGroups packGroups={packGroups} api={api} onDone={() => { loadAll(); flash("Saved!"); }} password={password} />
        )}

        {/* ── Export History ── */}
        {tab === "exports" && (
          <AdminExports exports={exports} packGroups={packGroups} api={api} />
        )}
      </div>
    </div>
  );
}

/**
 * Admin dashboard tab — shows a grid of station cards.
 * A station is considered ACTIVE if it sent a heartbeat within the last 2 minutes.
 * Each card shows: station name, operator, current order, progress bar, pass/flag counts,
 * and a queued-orders indicator.
 */
function AdminDashboard({ stations, packGroups, onRefresh }) {
  // eslint-disable-next-line react-hooks/purity
  const now = Date.now();
  const sorted = [...stations].sort((a, b) => (a.name || "").localeCompare(b.name || ""));
  if (sorted.length === 0) return (
    <div style={{ textAlign: "center", padding: 60, color: "#64748b" }}>
      <div style={{ fontSize: 36, marginBottom: 12 }}>📡</div>
      <div style={{ fontSize: 16, fontWeight: 600 }}>No stations connected yet.</div>
      <div style={{ fontSize: 13, marginTop: 6 }}>Stations appear here once they send a heartbeat.</div>
    </div>
  );
  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
        <div style={{ fontSize: 13, color: "#64748b" }}>{sorted.length} station{sorted.length !== 1 ? "s" : ""} — auto-refreshes every 5s</div>
        <button onClick={onRefresh} style={{ padding: "4px 12px", borderRadius: 6, border: "1px solid #334155", background: "transparent", color: "#94a3b8", fontSize: 12, cursor: "pointer" }}>↻ Refresh</button>
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(280px, 1fr))", gap: 14 }}>
        {sorted.map(s => {
          const p       = s.progress;
          const active  = s.lastSeen && (now - s.lastSeen) < 120000;
          const pct     = p ? Math.round((p.scanned / p.rosterCount) * 100) : 0;
          const pgName  = p?.packGroupId ? (packGroups.find(pg => pg.id === p.packGroupId)?.name || null) : null;
          return (
            <div key={s.id} style={{ background: "#161b22", border: `1px solid ${active ? "#21262d" : "#1e293b"}`, borderRadius: 12, padding: 16, opacity: active ? 1 : 0.55 }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 10 }}>
                <div>
                  <div style={{ fontWeight: 800, fontSize: 15 }}>{s.name}</div>
                  {p?.operatorName && <div style={{ fontSize: 11, color: "#64748b", marginTop: 2 }}>{p.operatorName}</div>}
                </div>
                <span style={{ fontSize: 10, fontWeight: 700, padding: "2px 8px", borderRadius: 20, background: active ? "rgba(34,197,94,0.12)" : "rgba(100,116,139,0.1)", color: active ? "#22c55e" : "#64748b", border: `1px solid ${active ? "rgba(34,197,94,0.3)" : "rgba(100,116,139,0.2)"}` }}>
                  {active ? "ACTIVE" : "OFFLINE"}
                </span>
              </div>
              {p ? (
                <>
                  <div style={{ fontSize: 12, color: "#94a3b8", marginBottom: 6 }}>
                    Order: <span style={{ color: "#e2e8f0", fontWeight: 600 }}>{p.orderNumber || "—"}</span>
                    {pgName && <span style={{ marginLeft: 8, color: "#60a5fa", fontSize: 11 }}>📦 {pgName}</span>}
                  </div>
                  <div style={{ height: 6, background: "#1e293b", borderRadius: 3, marginBottom: 6 }}>
                    <div style={{ height: "100%", background: pct === 100 ? "#22c55e" : "#3b82f6", borderRadius: 3, width: `${pct}%`, transition: "width 0.4s" }} />
                  </div>
                  <div style={{ display: "flex", gap: 12, fontSize: 12 }}>
                    <span style={{ color: "#e2e8f0" }}>{p.scanned}/{p.rosterCount} scanned</span>
                    <span style={{ color: "#22c55e" }}>✓ {p.passCount}</span>
                    <span style={{ color: "#ef4444" }}>⚑ {p.flagCount}</span>
                  </div>
                  {s.assignedQueue?.length > 0 && <div style={{ marginTop: 8, fontSize: 11, color: "#f59e0b", fontWeight: 600 }}>⏳ {s.assignedQueue.length} order{s.assignedQueue.length > 1 ? "s" : ""} queued</div>}
                </>
              ) : (
                <div style={{ fontSize: 12, color: "#475569" }}>
                  {s.assignedQueue?.length > 0 ? <span style={{ color: "#f59e0b", fontWeight: 600 }}>⏳ {s.assignedQueue.length} order{s.assignedQueue.length > 1 ? "s" : ""} queued — waiting for operator</span> : "No active session"}
                </div>
              )}
              {s.lastSeen && <div style={{ fontSize: 10, color: "#334155", marginTop: 8 }}>Last seen: {new Date(s.lastSeen).toLocaleTimeString()}</div>}
            </div>
          );
        })}
      </div>
    </div>
  );
}

/**
 * Admin "Assign Roster" tab.
 * Lets a supervisor push a roster (+ optional pack group) to a specific station.
 * The assignment is queued on the server and delivered to the station on its next heartbeat.
 * Multiple assignments can be queued; the station works through them in order.
 *
 * Known stations appear in the dropdown; a "Type a station name" option allows
 * assigning to a station that hasn't sent a heartbeat yet.
 */
function AdminAssign({ stations, packGroups, api, onDone }) {
  const [stationId,   setStationId]   = useState("");
  const [customName,  setCustomName]  = useState("");
  const [file,        setFile]        = useState(null);
  const [parsed,      setParsed]      = useState(null);
  const [orderNum,    setOrderNum]    = useState("");
  const [packGroupId, setPackGroupId] = useState("");
  const [error,       setError]       = useState(null);
  const [loading,     setLoading]     = useState(false);
  const fileRef = useRef(null);

  const handleFile = async (e) => {
    const f = e.target.files[0]; if (!f) return;
    e.target.value = ""; setError(null);
    try {
      const isExcel = /\.(xlsx|xls)$/i.test(f.name);
      const rows = isExcel ? await parseExcel(f) : parseCSV(await f.text());
      if (rows.length === 0) { setError("No valid rows found."); return; }
      setFile(f); setParsed(rows);
    } catch (err) { setError("Failed to parse: " + err.message); }
  };

  const assign = async () => {
    const sid = stationId === "__custom__" ? customName.trim() : stationId;
    if (!sid)    { setError("Select or enter a station."); return; }
    if (!parsed) { setError("Upload a roster file."); return; }
    setLoading(true);
    try {
      const selectedPG = packGroups.find(p => p.id === packGroupId);
      await api('/api/admin/assign', {
        method: 'POST',
        body: JSON.stringify({ stationId: sid, roster: parsed, rosterName: file.name, orderNumber: orderNum, packGroupId: packGroupId || null, binMap: selectedPG?.binMap || null }),
      });
      onDone();
      setFile(null); setParsed(null); setOrderNum(""); setPackGroupId(""); setStationId(""); setCustomName("");
    } catch (err) { setError(err.message); }
    setLoading(false);
  };

  const inp = { padding: "8px 12px", borderRadius: 8, border: "1px solid #334155", background: "#1e293b", color: "#e2e8f0", fontSize: 13, width: "100%", boxSizing: "border-box" };
  const lbl = { fontSize: 11, color: "#64748b", fontWeight: 700, textTransform: "uppercase", letterSpacing: 0.5, display: "block", marginBottom: 5 };

  const knownStations = stations.filter(s => s.name);

  return (
    <div style={{ maxWidth: 480 }}>
      <div style={{ fontSize: 16, fontWeight: 800, marginBottom: 20 }}>Assign Roster to Station</div>
      <div style={{ marginBottom: 14 }}>
        <label style={lbl}>Station</label>
        <select value={stationId} onChange={e => setStationId(e.target.value)} style={inp}>
          <option value="">— select station —</option>
          {knownStations.map(s => <option key={s.id} value={s.id}>{s.name}{s.progress?.orderNumber ? ` (${s.progress.orderNumber})` : ""}</option>)}
          <option value="__custom__">+ Type a station name…</option>
        </select>
      </div>
      {stationId === "__custom__" && (
        <div style={{ marginBottom: 14 }}>
          <label style={lbl}>Station Name</label>
          <input value={customName} onChange={e => setCustomName(e.target.value)} placeholder="e.g. Station 3" style={inp} />
        </div>
      )}
      <div style={{ marginBottom: 14 }}>
        <label style={lbl}>Order Number</label>
        <input value={orderNum} onChange={e => setOrderNum(e.target.value)} placeholder="e.g. ORD-2025-042" style={inp} />
      </div>
      <div style={{ marginBottom: 14 }}>
        <label style={lbl}>Roster File</label>
        <div onClick={() => fileRef.current.click()} style={{ ...inp, color: file ? "#22c55e" : "#64748b", cursor: "pointer", display: "flex", alignItems: "center", gap: 8 }}>
          <span>{file ? "✓" : "📂"}</span><span>{file ? file.name : "Choose CSV or Excel…"}</span>
        </div>
        <input ref={fileRef} type="file" accept=".csv,.xlsx,.xls,.tsv,.txt" style={{ display: "none" }} onChange={handleFile} />
        {parsed && <div style={{ fontSize: 11, color: "#64748b", marginTop: 4 }}>{parsed.length} players · {[...new Set(parsed.map(r => r.team).filter(Boolean))].length} teams</div>}
      </div>
      <div style={{ marginBottom: 20 }}>
        <label style={lbl}>Pack Group <span style={{ color: "#475569", textTransform: "none", fontWeight: 400 }}>(optional — links orders for combined packing)</span></label>
        <select value={packGroupId} onChange={e => setPackGroupId(e.target.value)} style={inp}>
          <option value="">— none —</option>
          {packGroups.map(pg => <option key={pg.id} value={pg.id}>{pg.name}</option>)}
        </select>
        {packGroupId && packGroups.find(p => p.id === packGroupId)?.binMap && (
          <div style={{ marginTop: 6, fontSize: 11, color: "#60a5fa" }}>
            Bin assignments from this group will override roster auto-detection.
          </div>
        )}
      </div>
      {error && <div style={{ fontSize: 12, color: "#f87171", marginBottom: 12 }}>{error}</div>}
      <button onClick={assign} disabled={loading} style={{ padding: "10px 28px", borderRadius: 8, border: "none", background: "#3b82f6", color: "#fff", fontWeight: 700, fontSize: 14, cursor: "pointer", opacity: loading ? 0.7 : 1 }}>
        {loading ? "Assigning…" : "Assign →"}
      </button>
    </div>
  );
}

/**
 * Admin "Pack Groups" tab.
 * A pack group links multiple orders that should be physically packed together.
 * It defines a shared bin→team mapping that overrides per-roster auto-detection.
 *
 * Views: list → create (new group) | detail (view linked orders)
 *
 * "Combined Export" downloads a single Excel with a Summary sheet and a Combined Roster
 * sheet containing all orders, sorted by bin number.
 */
function AdminPackGroups({ packGroups, api, onDone, password }) {
  const [view,      setView]      = useState("list"); // "list" | "create" | "detail"
  const [detail,    setDetail]    = useState(null);
  const [pgName,    setPgName]    = useState("");
  const [file,      setFile]      = useState(null);
  const [binDraft,  setBinDraft]  = useState({});
  const [error,     setError]     = useState(null);
  const [loading,   setLoading]   = useState(false);
  const fileRef = useRef(null);

  const inp = { padding: "8px 12px", borderRadius: 8, border: "1px solid #334155", background: "#1e293b", color: "#e2e8f0", fontSize: 13, width: "100%", boxSizing: "border-box" };
  const lbl = { fontSize: 11, color: "#64748b", fontWeight: 700, textTransform: "uppercase", letterSpacing: 0.5, display: "block", marginBottom: 5 };

  const handleFile = async (e) => {
    const f = e.target.files[0]; if (!f) return;
    e.target.value = "";
    try {
      const isExcel = /\.(xlsx|xls)$/i.test(f.name);
      const rows = isExcel ? await parseExcel(f) : parseCSV(await f.text());
      setFile(f);
      const teams = [...new Set(rows.map(r => r.team).filter(Boolean))].sort();
      const draft = {};
      teams.forEach((t, i) => { draft[t] = i + 1; });
      setBinDraft(draft);
    } catch (err) { setError("Failed to parse: " + err.message); }
  };

  const create = async () => {
    if (!pgName.trim()) { setError("Enter a group name."); return; }
    if (Object.keys(binDraft).length === 0) { setError("Upload a roster to detect teams."); return; }
    setLoading(true);
    try {
      await api('/api/admin/packgroups', { method: 'POST', body: JSON.stringify({ name: pgName.trim(), binMap: binDraft }) });
      onDone(); setView("list"); setPgName(""); setFile(null); setBinDraft({});
    } catch (err) { setError(err.message); }
    setLoading(false);
  };

  const openDetail = async (pg) => {
    try {
      const d = await fetch(`/api/admin/packgroups/${pg.id}`, { headers: { 'x-admin-password': password } }).then(r => r.json());
      setDetail(d); setView("detail");
    } catch (e) { console.warn('openDetail failed', e.message); }
  };

  const downloadCombined = async (pg) => {
    try {
      const { orders } = await fetch(`/api/admin/packgroups/${pg.id}/combined`, { headers: { 'x-admin-password': password } }).then(r => r.json());
      if (!window.XLSX) { alert("Excel library not loaded."); return; }
      const wb = window.XLSX.utils.book_new();
      // Summary sheet
      const summaryData = [
        ["Pack Group", pg.name],
        ["Orders", orders.length],
        ["Created", new Date(pg.createdAt).toLocaleString()],
        [],
        ["Bin Assignments"],
        ["Bin", "Team"],
        ...Object.entries(pg.binMap || {}).sort((a,b) => a[1]-b[1]).map(([t,b]) => [`Bin ${b}`, t]),
        [],
        ["Orders Included"],
        ["Order #", "Operator", "Station", "Date", "Passed", "Flagged"],
        ...orders.map(o => [o.orderNumber, o.operatorName, o.stationName, new Date(o.completedAt).toLocaleString(), o.roster.filter(r=>r.scanned==="pass").length, o.roster.filter(r=>r.scanned==="flag").length]),
      ];
      window.XLSX.utils.book_append_sheet(wb, window.XLSX.utils.aoa_to_sheet(summaryData), "Summary");
      // Combined roster sheet sorted by bin
      const allRows = orders.flatMap(o => o.roster.map(r => ({ ...r, _orderNumber: o.orderNumber, _operator: o.operatorName, _binNum: pg.binMap?.[r.team] || 999 })));
      allRows.sort((a, b) => a._binNum - b._binNum || (a.team||"").localeCompare(b.team||"") || (a.name||"").localeCompare(b.name||""));
      const cols = ["Order #", "Bin", "Team", "Name", "Number", "Size", "Status", "Issue", "Operator"];
      const combinedData = [cols, ...allRows.map(r => [r._orderNumber, r._binNum < 999 ? `Bin ${r._binNum}` : "—", r.team||"", r.name||"", r.number||"", r.size||"", r.scanned==="pass"?"PASS":r.scanned==="flag"?"FLAGGED":r.scanned==="resolved"?"RESOLVED":r.scanned==="extra"?"EXTRA (NOT IN ROSTER)":"NOT SCANNED", r.comment||"", r._operator])];
      window.XLSX.utils.book_append_sheet(wb, window.XLSX.utils.aoa_to_sheet(combinedData), "Combined Roster");
      window.XLSX.writeFile(wb, `packgroup_${pg.name.replace(/\s+/g,'_')}_${new Date().toISOString().slice(0,10)}.xlsx`);
    } catch (err) { alert("Export failed: " + err.message); }
  };

  if (view === "create") return (
    <div style={{ maxWidth: 480 }}>
      <button onClick={() => setView("list")} style={{ marginBottom: 16, background: "none", border: "none", color: "#64748b", cursor: "pointer", fontSize: 13 }}>← Back</button>
      <div style={{ fontSize: 16, fontWeight: 800, marginBottom: 20 }}>New Pack Group</div>
      <div style={{ marginBottom: 14 }}>
        <label style={lbl}>Group Name</label>
        <input value={pgName} onChange={e => setPgName(e.target.value)} placeholder="e.g. Spring 2025 — Teams A/B" style={inp} />
      </div>
      <div style={{ marginBottom: 14 }}>
        <label style={lbl}>Upload a representative roster to detect teams</label>
        <div onClick={() => fileRef.current.click()} style={{ ...inp, color: file ? "#22c55e" : "#64748b", cursor: "pointer", display: "flex", alignItems: "center", gap: 8 }}>
          <span>{file ? "✓" : "📂"}</span><span>{file ? file.name : "Choose CSV or Excel…"}</span>
        </div>
        <input ref={fileRef} type="file" accept=".csv,.xlsx,.xls,.tsv,.txt" style={{ display: "none" }} onChange={handleFile} />
      </div>
      {Object.keys(binDraft).length > 0 && (
        <div style={{ marginBottom: 20 }}>
          <label style={lbl}>Bin Assignments</label>
          {Object.entries(binDraft).sort((a,b) => a[1]-b[1]).map(([team, bin]) => (
            <div key={team} style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 8 }}>
              <div style={{ flex: 1, fontSize: 13, color: "#e2e8f0" }}>{team}</div>
              <input type="number" min={1} max={20} value={bin}
                onChange={e => setBinDraft(prev => ({ ...prev, [team]: parseInt(e.target.value) || 1 }))}
                style={{ width: 70, padding: "6px 10px", borderRadius: 6, border: "1px solid #334155", background: "#1e293b", color: "#60a5fa", fontWeight: 700, fontSize: 14, textAlign: "center" }} />
            </div>
          ))}
        </div>
      )}
      {error && <div style={{ fontSize: 12, color: "#f87171", marginBottom: 12 }}>{error}</div>}
      <button onClick={create} disabled={loading} style={{ padding: "10px 28px", borderRadius: 8, border: "none", background: "#22c55e", color: "#fff", fontWeight: 700, fontSize: 14, cursor: "pointer", opacity: loading ? 0.7 : 1 }}>
        {loading ? "Saving…" : "Create Pack Group"}
      </button>
    </div>
  );

  if (view === "detail" && detail) return (
    <div style={{ maxWidth: 620 }}>
      <button onClick={() => setView("list")} style={{ marginBottom: 16, background: "none", border: "none", color: "#64748b", cursor: "pointer", fontSize: 13 }}>← Back</button>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
        <div style={{ fontSize: 16, fontWeight: 800 }}>📦 {detail.name}</div>
        <button onClick={() => downloadCombined(detail)} style={{ padding: "7px 18px", borderRadius: 8, border: "none", background: "#22c55e", color: "#fff", fontWeight: 700, fontSize: 13, cursor: "pointer" }}>⬇ Combined Export</button>
      </div>
      <div style={{ marginBottom: 16 }}>
        <div style={{ fontSize: 11, color: "#64748b", fontWeight: 700, textTransform: "uppercase", marginBottom: 8 }}>Bin Assignments</div>
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
          {Object.entries(detail.binMap || {}).sort((a,b) => a[1]-b[1]).map(([team, bin]) => (
            <div key={team} style={{ background: "#1e293b", border: "1px solid #334155", borderRadius: 8, padding: "6px 14px", fontSize: 13 }}>
              <span style={{ color: "#60a5fa", fontWeight: 700 }}>Bin {bin}</span> <span style={{ color: "#e2e8f0" }}>{team}</span>
            </div>
          ))}
        </div>
      </div>
      <div>
        <div style={{ fontSize: 11, color: "#64748b", fontWeight: 700, textTransform: "uppercase", marginBottom: 8 }}>Orders ({detail.orders?.length || 0})</div>
        {(!detail.orders || detail.orders.length === 0) && <div style={{ color: "#475569", fontSize: 13 }}>No completed orders yet.</div>}
        {detail.orders?.map(o => (
          <div key={o.id} style={{ background: "#161b22", border: "1px solid #21262d", borderRadius: 8, padding: "10px 14px", marginBottom: 8, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
            <div>
              <div style={{ fontWeight: 700, fontSize: 13 }}>{o.orderNumber || "—"}</div>
              <div style={{ fontSize: 11, color: "#64748b" }}>{o.operatorName} · {o.stationName} · {new Date(o.completedAt).toLocaleString()}</div>
            </div>
            <div style={{ fontSize: 12, color: "#94a3b8" }}>{o.rosterCount} jerseys · {o.passCount} passed</div>
          </div>
        ))}
      </div>
    </div>
  );

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
        <div style={{ fontSize: 16, fontWeight: 800 }}>Pack Groups</div>
        <button onClick={() => setView("create")} style={{ padding: "7px 18px", borderRadius: 8, border: "none", background: "#3b82f6", color: "#fff", fontWeight: 700, fontSize: 13, cursor: "pointer" }}>+ New Group</button>
      </div>
      {packGroups.length === 0 && <div style={{ color: "#475569", fontSize: 13 }}>No pack groups yet. Create one to link orders that should be packed together.</div>}
      {packGroups.map(pg => (
        <div key={pg.id} onClick={() => openDetail(pg)} style={{ background: "#161b22", border: "1px solid #21262d", borderRadius: 10, padding: "14px 18px", marginBottom: 10, cursor: "pointer", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <div>
            <div style={{ fontWeight: 700, fontSize: 14 }}>📦 {pg.name}</div>
            <div style={{ fontSize: 11, color: "#64748b", marginTop: 3 }}>
              {Object.keys(pg.binMap || {}).length} teams · {pg.orderCount || 0} orders
              {" · "}Created {new Date(pg.createdAt).toLocaleDateString()}
            </div>
          </div>
          <div style={{ fontSize: 12, color: "#475569" }}>→</div>
        </div>
      ))}
    </div>
  );
}

/**
 * Admin "Export History" tab.
 * Lists all completed orders from the server's export index.
 * Each row has a download button to re-export the full roster as Excel.
 */
function AdminExports({ exports, packGroups, api }) {
  const downloadExport = async (id) => {
    try {
      const data = await api(`/api/admin/exports/${id}`);
      if (!window.XLSX) { alert("Excel library not loaded."); return; }
      exportRosterXLSX(data.roster, data.orderNumber, data.operatorName, data.binMap);
    } catch (err) { alert("Download failed: " + err.message); }
  };

  return (
    <div>
      <div style={{ fontSize: 16, fontWeight: 800, marginBottom: 16 }}>Export History ({exports.length})</div>
      {exports.length === 0 && <div style={{ color: "#475569", fontSize: 13 }}>No completed orders saved yet. Orders are saved when an operator exports.</div>}
      <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
        <thead>
          <tr>
            {["Order #", "Operator", "Station", "Pack Group", "Date", "Count", ""].map(h => (
              <th key={h} style={{ padding: "6px 10px", textAlign: "left", fontSize: 11, fontWeight: 700, color: "#64748b", textTransform: "uppercase", letterSpacing: 0.5, borderBottom: "1px solid #21262d" }}>{h}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {exports.map(e => (
            <tr key={e.id} style={{ borderBottom: "1px solid #1e293b" }}>
              <td style={{ padding: "8px 10px", fontWeight: 700, color: "#e2e8f0" }}>{e.orderNumber || "—"}</td>
              <td style={{ padding: "8px 10px", color: "#94a3b8" }}>{e.operatorName || "—"}</td>
              <td style={{ padding: "8px 10px", color: "#94a3b8" }}>{e.stationName || "—"}</td>
              <td style={{ padding: "8px 10px", color: "#60a5fa" }}>{e.packGroupId ? (packGroups.find(p => p.id === e.packGroupId)?.name || "—") : "—"}</td>
              <td style={{ padding: "8px 10px", color: "#64748b" }}>{new Date(e.completedAt).toLocaleString()}</td>
              <td style={{ padding: "8px 10px", color: "#64748b" }}>{e.rosterCount}{e.extraCount > 0 ? <span style={{ color: "#f59e0b", marginLeft: 4, fontWeight: 700 }}>+{e.extraCount}</span> : null}</td>
              <td style={{ padding: "8px 10px" }}>
                <button onClick={() => downloadExport(e.id)} style={{ padding: "3px 10px", borderRadius: 6, border: "1px solid #334155", background: "transparent", color: "#94a3b8", fontSize: 11, cursor: "pointer" }}>⬇ Excel</button>
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

// ── Main App ──────────────────────────────────────────────────────────────────
/**
 * Root component. ALL application state lives here.
 *
 * State groups:
 *   Station / admin   — stationName, showStationSetup, showAdmin, preAssigned
 *   Session           — sessionStarted, roster, rosterFile, orderNumber, operatorName
 *   Camera / scanning — cameraOn, scanning, inputMode, thumb, overlay
 *   Results / log     — lastResult, lastConfirmed, log
 *   UI                — view, error, xlsxReady, showSettings, showRosterComplete, showBinSetup
 *   Keys              — scanKey, confirmKey, cancelKey
 *   Bins / packing    — binMap, packGroupId
 *   Timing            — firstScanTime, now (used for ETA calculation)
 *
 * Key effects:
 *   - Loads SheetJS from CDN on mount
 *   - Polls server heartbeat every 15s when idle (to receive admin assignments)
 *   - Reports scan progress to server every 2s (debounced) when scanning
 *   - Persists session to localStorage on every state change
 *   - Listens for the scan trigger key globally (except when overlays are shown)
 */
export default function App() {
  // DOM refs
  const videoRef       = useRef(null);  // <video> element for live camera feed
  const canvasRef      = useRef(null);  // hidden <canvas> used to capture frames
  const streamRef      = useRef(null);  // MediaStream from getUserMedia (for cleanup)
  const rosterRef      = useRef(null);  // hidden <input type="file"> for roster uploads
  const photoRef       = useRef(null);  // hidden <input type="file"> for photo uploads
  const heartbeatTimer = useRef(null);  // setTimeout handle for debounced progress heartbeat
  const _s        = useRef(loadSession()).current;
  const [stationName,        setStationName]        = useState(() => localStorage.getItem("jerseyqc_station") || "");
  const [showStationSetup,   setShowStationSetup]   = useState(() => !localStorage.getItem("jerseyqc_station"));
  const [showAdmin,          setShowAdmin]          = useState(false);
  const [preAssigned,        setPreAssigned]        = useState(null);

  const [sessionStarted,     setSessionStarted]     = useState(_s?.sessionStarted     ?? false);
  const [roster,             setRoster]             = useState((_s?.roster ?? []).map(r => ({ comment: "", ...r })));
  const [rosterFile,         setRosterFile]         = useState(_s?.rosterFile         ?? null);
  const [orderNumber,        setOrderNumber]        = useState(_s?.orderNumber        ?? "");
  const [operatorName,       setOperatorName]       = useState(_s?.operatorName       ?? "");
  const [cameraOn,           setCameraOn]           = useState(false);
  const [scanning,           setScanning]           = useState(false);
  const [inputMode,          setInputMode]          = useState("camera");
  const [thumb,              setThumb]              = useState(null);
  const [overlay,            setOverlay]            = useState(null);
  const [lastResult,         setLastResult]         = useState(null);
  const [lastConfirmed,      setLastConfirmed]      = useState(null);
  const [log,                setLog]                = useState(_s?.log               ?? []);
  const [view,               setView]               = useState("roster");
  const [error,              setError]              = useState(null);
  const [xlsxReady,          setXlsxReady]          = useState(!!window.XLSX);
  const [scanKey,            setScanKey]            = useState("KeyB");
  const [confirmKey,         setConfirmKey]         = useState("KeyB");
  const [cancelKey,          setCancelKey]          = useState("KeyA");
  const [showSettings,       setShowSettings]       = useState(false);
  const [showRosterComplete, setShowRosterComplete] = useState(false);
  const [showBinSetup,       setShowBinSetup]       = useState(false);
  const [binMap,             setBinMap]             = useState(_s?.binMap            ?? null);
  const [firstScanTime,      setFirstScanTime]      = useState(_s?.firstScanTime     ?? null);
  const [packGroupId,        setPackGroupId]        = useState(_s?.packGroupId       ?? null);
  const [now,                setNow]                = useState(Date.now());

  const extraEntries   = roster.filter(r => r._extra);
  const extraCount     = extraEntries.length;
  const origRoster     = roster.filter(r => !r._extra);
  const scanned        = roster.filter(r => r.scanned && r.scanned !== false);
  const remaining      = origRoster.filter(r => !r.scanned || r.scanned === false).length;
  const passCount      = log.filter(l => l.status === S_PASS).length;
  const flagCount      = log.filter(l => l.status === S_FLAGGED).length;
  const resolvedCount  = log.filter(l => l.status === S_MANUAL).length;
  const rosterComplete = origRoster.length > 0 && remaining === 0;
  const elapsedMin     = firstScanTime ? (now - firstScanTime) / 60000 : 0;
  const scanRate       = elapsedMin > 1 && passCount > 0 ? passCount / elapsedMin : null;
  const etaMin         = scanRate && remaining > 0 ? Math.ceil(remaining / scanRate) : null;
  const rosterCols     = roster.length > 0 ? Object.keys(roster[0]).filter(k => k !== "_id" && k !== "_extra" && k !== "scanned" && k !== "comment") : [];
  const keyLabel       = formatKey(scanKey);

  useEffect(() => { const t = setInterval(() => setNow(Date.now()), 30000); return () => clearInterval(t); }, []);

  // Station heartbeat when idle (no active session) — also polls for pre-assigned roster
  useEffect(() => {
    if (!stationName || sessionStarted) return;
    const beat = async () => {
      try {
        const r = await fetch('/api/station/heartbeat', { method: 'POST', headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ stationId: stationName, stationName, progress: null }) });
        const { assigned } = await r.json();
        if (assigned) setPreAssigned(assigned);
      } catch { /* network error — silent */ }
    };
    beat();
    const t = setInterval(beat, 15000);
    return () => clearInterval(t);
  }, [stationName, sessionStarted]);

  // Persist session to localStorage on every state change
  useEffect(() => {
    if (!sessionStarted) return;
    const stripThumb = ({ thumb, ...rest }) => rest; // eslint-disable-line no-unused-vars
    saveSession({ sessionStarted, roster, log: log.map(stripThumb), orderNumber, operatorName, rosterFile, binMap, firstScanTime, packGroupId });
  }, [sessionStarted, roster, log, orderNumber, operatorName, rosterFile, binMap, firstScanTime, packGroupId]);

  // Report progress to server — debounced so rapid scan confirmations batch into one call
  useEffect(() => {
    if (!sessionStarted || !stationName) return;
    if (heartbeatTimer.current) clearTimeout(heartbeatTimer.current);
    heartbeatTimer.current = setTimeout(() => {
      const scanned   = roster.filter(r => r.scanned && r.scanned !== false).length;
      const passCount = log.filter(l => l.status === S_PASS).length;
      const flagCount = log.filter(l => l.status === S_FLAGGED).length;
      fetch('/api/station/heartbeat', { method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ stationId: stationName, stationName, progress: { orderNumber, operatorName, rosterCount: roster.filter(r => !r._extra).length, scanned, passCount, flagCount, packGroupId, lastScan: Date.now() } }),
      }).catch(() => {});
    }, 2000);
    return () => clearTimeout(heartbeatTimer.current);
  }, [sessionStarted, stationName, roster, log, orderNumber, operatorName, packGroupId]);


  useEffect(() => {
    if (window.XLSX) return;
    const s = document.createElement("script");
    s.src = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
    s.onload  = () => setXlsxReady(true);
    s.onerror = () => setError("Failed to load Excel support.");
    document.head.appendChild(s);
  }, []);

  useEffect(() => { return () => { streamRef.current?.getTracks().forEach(t => t.stop()); }; }, []);

  // Only trigger complete modal after at least one scan
  useEffect(() => {
    if (roster.length > 0 && remaining === 0 && firstScanTime) setShowRosterComplete(true);
  }, [remaining, roster.length, firstScanTime]);

  // ── Apply roster + build bins ─────────────────────────────────────────────
  /**
   * Resets all scan state and loads a new roster.
   * Called on session start and when the operator replaces the roster mid-session.
   *
   * If overrideBinMap is provided (from an admin pack group), it is used directly
   * and the bin setup modal is skipped. Otherwise, bins are auto-detected from
   * team names in the roster: if there's more than one team, the bin setup modal is shown.
   *
   * @param {Object[]} parsed - Normalised roster rows
   * @param {string} fileName - Original file name (shown in header)
   * @param {Object|null} overrideBinMap - Pre-defined bin map from pack group, or null
   * @param {string|null} pgId - Pack group ID, or null
   */
  const applyRoster = useCallback((parsed, fileName, overrideBinMap, pgId) => {
    setRoster(parsed.map(r => ({ ...r, scanned: false, comment: "" })));
    setRosterFile(fileName);
    setLastResult(null); setLog([]); setThumb(null);
    setOverlay(null); setFirstScanTime(null); setLastConfirmed(null);
    setShowRosterComplete(false);
    setPackGroupId(pgId || null);
    if (overrideBinMap && Object.keys(overrideBinMap).length > 0) {
      setBinMap(overrideBinMap);
      setShowBinSetup(false); // already confirmed by admin via pack group
    } else {
      const bins = buildBinMap(parsed);
      if (Object.keys(bins).length > 1) { setBinMap(bins); setShowBinSetup(true); }
      else { setBinMap(null); setShowBinSetup(false); }
    }
  }, []);

  // ── Camera ────────────────────────────────────────────────────────────────
  const startCamera = useCallback(async () => {
    setError(null);
    try {
      const stream = await navigator.mediaDevices.getUserMedia({
        video: { facingMode: "environment", width: { ideal: 1280 }, height: { ideal: 720 } },
        audio: false,
      });
      streamRef.current = stream;
      if (videoRef.current) videoRef.current.srcObject = stream;
      setCameraOn(true);
    } catch {
      setError("Camera access denied. Click the camera icon in your browser's address bar.");
    }
  }, []);

  const stopCamera = useCallback(() => {
    streamRef.current?.getTracks().forEach(t => t.stop());
    streamRef.current = null;
    setCameraOn(false);
  }, []);

  const handleNewSession = useCallback(() => {
    if (!window.confirm("Start a new session? Current progress will be cleared.")) return;
    clearSession();
    streamRef.current?.getTracks().forEach(t => t.stop());
    streamRef.current = null;
    setCameraOn(false);
    setSessionStarted(false);
    setRoster([]); setRosterFile(null); setOrderNumber(""); setOperatorName("");
    setLog([]); setBinMap(null); setFirstScanTime(null); setPackGroupId(null);
    setOverlay(null); setThumb(null); setLastResult(null); setLastConfirmed(null);
    setShowRosterComplete(false); setShowBinSetup(false); setError(null);
  }, []);

  const handleExportAndSave = useCallback((r, oNum, oName, bMap, pgId) => {
    exportRosterXLSX(r, oNum, oName, bMap);
    if (!stationName) return;
    const strip = ({ thumb, ...rest }) => rest; // eslint-disable-line no-unused-vars
    fetch('/api/station/complete', { method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ stationId: stationName, stationName, orderNumber: oNum, operatorName: oName, roster: r.map(strip), log: log.map(strip), binMap: bMap, packGroupId: pgId || null, completedAt: new Date().toISOString() }),
    }).catch(() => {});
  }, [stationName, log]);

  // ── Core scan ─────────────────────────────────────────────────────────────
  /**
   * The main scanning function. Sends a JPEG image to Claude Haiku and processes the result.
   *
   * Flow:
   *   1. Records first scan time (for ETA calculation)
   *   2. Sets scanning=true to disable re-triggering and show "Analysing…" UI
   *   3. Calls Claude via /api/anthropic/v1/messages with the image as base64
   *   4. Parses the JSON response to extract { name, number }
   *   5. Runs findMatches() against the roster
   *   6. Sets the appropriate overlay mode (confirm/pick/close/flag)
   *
   * Error handling: any API error or JSON parse failure results in a red flag overlay.
   *
   * @param {string} base64 - JPEG image data (no data: prefix)
   * @param {string} dataUrl - Full data: URL for thumbnail display in the overlay
   */
  const doRunScan = useCallback(async (base64, dataUrl) => {
    if (!firstScanTime) setFirstScanTime(Date.now());
    setScanning(true);
    setThumb(dataUrl);
    playTone("scan");

    let detected = { name: "", number: "" };
    let apiError  = null;

    try {
      const resp = await fetch("/api/anthropic/v1/messages", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "x-api-key": CLAUDE_API_KEY,
          "anthropic-version": "2023-06-01",
          "anthropic-dangerous-direct-browser-access": "true",
        },
        body: JSON.stringify({
          model: CLAUDE_MODEL,
          max_tokens: 200,
          messages: [{
            role: "user",
            content: [
              { type: "image", source: { type: "base64", media_type: "image/jpeg", data: base64 } },
              { type: "text",  text: `This is a photo of the back of a sports jersey in a manufacturing QC environment. Extract the player name and jersey number. Return ONLY valid JSON, no markdown: {"name":"PLAYERNAME","number":"##"}. Use "" if not visible.` },
            ],
          }],
        }),
      });
      if (!resp.ok) throw new Error(`API error ${resp.status}: ${await resp.text()}`);
      const data = await resp.json();
      console.log(`Tokens — input: ${data.usage?.input_tokens}, output: ${data.usage?.output_tokens}`);
      const raw  = (data.content || []).map(b => b.text || "").join("").replace(/```json|```/g, "").trim();
      detected   = JSON.parse(raw);
    } catch (e) {
      apiError = e.message;
    }

    setScanning(false);
    const scanObj  = { id: Date.now(), detected, thumb: dataUrl };
    const showFlag = (reason, opts = {}) => { setOverlay({ mode: "flag", scan: scanObj, reason, ...opts }); playTone("flag"); };

    if (apiError)                            { showFlag("API error: " + apiError); return; }
    if (!detected.name && !detected.number) { showFlag("Could not read jersey — no name or number detected. Check lighting and angle."); return; }

    const match = findMatches(roster, detected.name, detected.number);
    if (match.type === "exact") {
      // Single exact match — if already scanned, go straight to duplicate check; otherwise confirm
      if (match.match.scanned && match.match.scanned !== false) {
        showFlag(`#${match.match.number}${match.match.name ? ` · ${match.match.name}` : ""} was already scanned. Same jersey, or an extra?`, { canAddExtra: true, flagTitle: "ALREADY SCANNED" });
      } else {
        setOverlay({ mode: "confirm", scan: { ...scanObj, match: match.match } });
      }
    } else if (match.type === "number_conflict" || match.type === "size_pick") {
      setOverlay({ mode: "pick",  scan: scanObj, candidates: match.candidates });
    } else if (match.type === "close") {
      setOverlay({ mode: "close", scan: scanObj, candidates: match.candidates });
    } else {
      showFlag(`"${detected.name || "?"}" #${detected.number || "?"} not found in roster.`, { canAddExtra: true });
    }
  }, [roster, firstScanTime]);

  // ── Camera scan trigger ───────────────────────────────────────────────────
  /**
   * Captures the current video frame to a canvas and passes it to doRunScan().
   * Guards against scanning when: no camera, already scanning, no roster, or in upload mode.
   *
   * IMPORTANT: The canvas is rotated 180° before the image is captured.
   * This is because the physical camera on the scanning rig is mounted upside-down.
   * If the camera is ever remounted, remove the translate+rotate calls.
   */
  const doTriggerScan = useCallback(() => {
    if (!videoRef.current || !canvasRef.current || scanning || !cameraOn || roster.length === 0 || inputMode !== "camera") return;
    const video  = videoRef.current;
    const canvas = canvasRef.current;
    const w = video.videoWidth, h = video.videoHeight;
    canvas.width = w; canvas.height = h;
    const ctx = canvas.getContext("2d");
    // Rotate 180° to correct for upside-down camera mounting
    ctx.translate(w / 2, h / 2);
    ctx.rotate(Math.PI);
    ctx.drawImage(video, -w / 2, -h / 2, w, h);
    const dataUrl = canvas.toDataURL("image/jpeg", 0.92);
    doRunScan(dataUrl.split(",")[1], dataUrl);
  }, [scanning, cameraOn, roster.length, inputMode, doRunScan]);

  // ── Auto-scan after green confirm ─────────────────────────────────────────
  const doAutoScan = useCallback(() => {
    if (!cameraOn || inputMode !== "camera") return;
    setTimeout(() => doTriggerScan(), 150);
  }, [doTriggerScan, cameraOn, inputMode]);

  // ── Photo upload ──────────────────────────────────────────────────────────
  const handlePhotoUpload = useCallback((e) => {
    const file = e.target.files[0]; if (!file) return;
    e.target.value = "";
    const reader = new FileReader();
    reader.onload = (ev) => {
      const img = new Image();
      img.onload = () => {
        const canvas = canvasRef.current;
        canvas.width = img.width; canvas.height = img.height;
        canvas.getContext("2d").drawImage(img, 0, 0);
        const dataUrl = canvas.toDataURL("image/jpeg", 0.92);
        doRunScan(dataUrl.split(",")[1], dataUrl);
      };
      img.src = ev.target.result;
    };
    reader.readAsDataURL(file);
  }, [doRunScan]);

  // ── Roster replacement ────────────────────────────────────────────────────
  const handleRosterUpload = useCallback(async (e) => {
    const file = e.target.files[0]; if (!file) return;
    e.target.value = ""; setError(null);
    try {
      const isExcel = /\.(xlsx|xls)$/i.test(file.name);
      const parsed  = isExcel ? await parseExcel(file) : parseCSV(await file.text());
      if (parsed.length === 0) { setError("No valid rows found. Check columns: name, number, team, size."); return; }
      applyRoster(parsed, file.name);
    } catch (err) { setError("Failed to parse roster: " + err.message); }
  }, [applyRoster]);

  // ── Overlay action handlers ────────────────────────────────────────────────
  // These functions are passed as props to ScanOverlay and handle every user
  // action that can occur on the result screen.

  /**
   * Confirms a matched jersey. Updates the roster entry to "pass", adds a log entry,
   * and briefly shows the "done" flash before auto-triggering the next scan.
   * If the matched entry was already scanned, re-routes to the duplicate/extra flag screen.
   */
  const handleConfirm = useCallback(() => {
    if (!overlay?.scan?.match) return;
    const scan  = overlay.scan;
    // If the operator picked an already-scanned entry, ask double-scan vs extra before committing
    if (scan.match.scanned && scan.match.scanned !== false) {
      const m = scan.match;
      setOverlay({ mode: "flag", scan, reason: `#${m.number}${m.name ? ` · ${m.name}` : ""} was already scanned. Same jersey, or an extra?`, canAddExtra: true, flagTitle: "ALREADY SCANNED" });
      playTone("flag");
      return;
    }
    const entry = { ...scan, id: scan.id || Date.now(), status: S_PASS, timestamp: new Date().toLocaleTimeString() };
    setRoster(prev => prev.map(r => r._id === scan.match._id ? { ...r, scanned: "pass" } : r));
    setLog(prev => [entry, ...prev]);
    setLastResult(entry);
    setLastConfirmed(entry);
    setOverlay({ mode: "done", scan });
    playTone("pass");
  }, [overlay]);

  /** Dismisses the overlay without recording anything — operator retries the scan. */
  const handleEdit = useCallback(() => { setOverlay(null); }, []);

  /** On a pick/close screen, selects a candidate and advances to the green confirm screen. */
  const handlePickCandidate = useCallback((candidate) => {
    if (!overlay?.scan || !candidate) return;
    setOverlay({ mode: "confirm", scan: { ...overlay.scan, match: candidate } });
  }, [overlay]);

  /**
   * "Bad scan, correct jersey" — the scanner misread the jersey, but the operator
   * knows which one it is. Flags the item in the log with a note that it was a scan error.
   * The roster entry is NOT marked as passed (it still shows as flagged for review).
   */
  const handleFlagBadJersey = useCallback((candidate) => {
    if (!overlay?.scan) return;
    const entry = { ...overlay.scan, id: overlay.scan.id || Date.now(), status: S_FLAGGED, reason: "Bad scan — correct jersey (misread by scanner)", match: candidate, timestamp: new Date().toLocaleTimeString() };
    setLog(prev => [entry, ...prev]); setLastResult(entry); setOverlay(null); playTone("flag");
  }, [overlay]);

  /**
   * Flags the current scan as a problem jersey and closes the overlay.
   * Used for: unrecognised jerseys, API failures, explicit operator flags,
   * and duplicate scans (already-scanned jersey confirmed again).
   */
  const handleFlagBadScan = useCallback(() => {
    if (!overlay?.scan) return;
    const entry = { ...overlay.scan, id: overlay.scan.id || Date.now(), status: S_FLAGGED, reason: overlay.reason || "Flagged by operator — bad jersey.", timestamp: new Date().toLocaleTimeString() };
    setLog(prev => [entry, ...prev]); setLastResult(entry); setOverlay(null); playTone("flag");
  }, [overlay]);

  /**
   * Adds a scanned jersey as an "extra" — it's not in the roster but physically exists.
   * Creates a new roster entry with _extra: true so it shows in the roster table
   * and is included in the Excel export with "EXTRA (NOT IN ROSTER)" status.
   * Uses the operator-confirmed identity if available, otherwise falls back to AI detection.
   */
  const handleExtraJersey = useCallback(() => {
    if (!overlay?.scan) return;
    const det   = overlay.scan.detected;
    const match = overlay.scan.match; // set if operator picked/confirmed a specific entry
    // Build a row with the same column shape as existing roster entries (blank for unknown fields)
    const baseShape = roster[0]
      ? Object.fromEntries(Object.keys(roster[0]).filter(k => k !== '_id' && k !== '_extra').map(k => [k, ""]))
      : {};
    // Use the operator-confirmed identity when available; fall back to AI detection
    const name   = match ? (match.name   || "") : (det.name   || "");
    const number = match ? (match.number || "") : (det.number || "");
    const team   = match ? (match.team   || "") : "";
    const size   = match ? (match.size   || "") : "";
    const newEntry = { ...baseShape, _id: Date.now(), _extra: true, name, number, team, size, scanned: S_EXTRA, comment: "" };
    const logEntry = { id: Date.now() + 1, status: S_EXTRA, detected: det, match: newEntry, timestamp: new Date().toLocaleTimeString(), reason: "Extra jersey — not in roster" };
    setRoster(prev => [...prev, newEntry]);
    setLog(prev => [logEntry, ...prev]);
    setLastResult(logEntry);
    setOverlay(null);
    playTone("flag");
  }, [overlay, roster]);

  /** Closes the scan overlay without recording anything (used by the "done" auto-dismiss). */
  const handleOverlayDismiss = useCallback(() => { setOverlay(null); }, []);

  // ── Manual roster edits ───────────────────────────────────────────────────
  /**
   * Handles direct roster table row actions (the small buttons on each row).
   * Actions:
   *   "pass"  — manually mark an entry as passed (e.g. visually verified without scanning)
   *   "flag"  — manually flag an entry for review
   *   "undo"  — reset a scanned entry back to unscanned; removes _extra entries entirely
   *
   * @param {Object} entry - The roster row object
   * @param {"pass"|"flag"|"undo"} action
   */
  const handleRosterEdit = useCallback((entry, action) => {
    if (action === "pass" || action === "flag") {
      const logEntry = {
        id: Date.now(),
        status: action === "pass" ? S_PASS : S_FLAGGED,
        detected: { name: entry.name || "", number: entry.number || "" },
        match: entry,
        timestamp: new Date().toLocaleTimeString(),
        reason: action === "pass" ? "Manual confirmation" : "Manual flag",
      };
      setRoster(prev => prev.map(r => r._id === entry._id ? { ...r, scanned: action === "pass" ? "pass" : "flag" } : r));
      setLog(prev => [logEntry, ...prev]);
      if (!firstScanTime) setFirstScanTime(Date.now());
    } else if (action === "undo") {
      if (entry._extra) {
        // Extra entries were added dynamically — remove them from the roster entirely
        setRoster(prev => prev.filter(r => r._id !== entry._id));
      } else {
        setRoster(prev => prev.map(r => r._id === entry._id ? { ...r, scanned: false } : r));
      }
      setLog(prev => {
        const idx = prev.findIndex(l => l.match?._id === entry._id);
        return idx === -1 ? prev : [...prev.slice(0, idx), ...prev.slice(idx + 1)];
      });
    }
  }, [firstScanTime]);

  const handleRosterComment = useCallback((rosterId, text) => {
    setRoster(prev => prev.map(r => r._id === rosterId ? { ...r, comment: text } : r));
  }, []);

  // ── Flag resolution ───────────────────────────────────────────────────────
  /**
   * Resolves a flagged log entry (called from ResultCard's Accept/Reject/Note buttons).
   * Updates the log entry's status to S_MANUAL and records the resolution text.
   * Also updates the corresponding roster entry to "resolved" so it shows correctly
   * in the roster table and is counted in the export.
   *
   * @param {number} logId - The log entry's id field
   * @param {string} resolution - Human-readable resolution description
   */
  const resolveFlag = useCallback((logId, resolution) => {
    const entry = log.find(l => l.id === logId);
    setLog(prev => prev.map(l => l.id === logId ? { ...l, status: S_MANUAL, resolution } : l));
    setLastResult(prev => prev?.id === logId ? { ...prev, status: S_MANUAL, resolution } : prev);
    if (entry?.match?._id !== undefined) {
      setRoster(prev => prev.map(r => r._id === entry.match._id ? { ...r, scanned: "resolved" } : r));
    }
  }, [log]);

  // ── Global keyboard: scan trigger only ───────────────────────────────────
  // This listener only handles the scan key (default: B).
  // All other overlay keyboard shortcuts are handled inside ScanOverlay's own useEffect.
  // We skip scanning when: an overlay is showing, a modal is open, or the focus is on a text input.
  useEffect(() => {
    const onKey = (e) => {
      if (overlay || showBinSetup || showRosterComplete || showSettings) return;
      if (e.code !== scanKey) return;
      if (["INPUT", "TEXTAREA"].includes(document.activeElement?.tagName)) return;
      e.preventDefault();
      doTriggerScan();
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [overlay, showBinSetup, showRosterComplete, showSettings, scanKey, doTriggerScan]);

  // ── Render ────────────────────────────────────────────────────────────────
  return (
    <div style={{ fontFamily: "system-ui,sans-serif", background: "#0d1117", height: "100vh", color: "#e2e8f0", display: "flex", flexDirection: "column", overflow: "hidden" }}>

      {showStationSetup && (
        <StationSetupModal onSave={name => {
          localStorage.setItem("jerseyqc_station", name);
          setStationName(name); setShowStationSetup(false);
        }} />
      )}

      {showAdmin && <AdminPanel onClose={() => setShowAdmin(false)} />}

      {!sessionStarted && !showStationSetup && (
        <SessionStartModal
          xlsxReady={xlsxReady}
          preAssigned={preAssigned}
          onAdmin={() => setShowAdmin(true)}
          onStart={({ roster, orderNumber, operatorName, rosterFile, packGroupId, binMap, acceptedOrderNumber }) => {
            applyRoster(roster, rosterFile, binMap, packGroupId);
            setOrderNumber(orderNumber);
            setOperatorName(operatorName);
            setSessionStarted(true);
            if (acceptedOrderNumber && stationName) {
              fetch('/api/station/roster-accepted', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ stationId: stationName, orderNumber: acceptedOrderNumber }) }).catch(() => {});
              setPreAssigned(null);
            }
          }}
        />
      )}

      {showBinSetup && binMap && (
        <BinSetupModal binMap={binMap} onConfirm={() => setShowBinSetup(false)} />
      )}

      {overlay && (
        <ScanOverlay
          state={overlay}
          onConfirm={handleConfirm}
          onEdit={handleEdit}
          onPickCandidate={handlePickCandidate}
          onFlagBadScan={handleFlagBadScan}
          onFlagBadJersey={handleFlagBadJersey}
          onAddExtra={handleExtraJersey}
          onDismiss={handleOverlayDismiss}
          onAutoScan={doAutoScan}
          binMap={binMap}
          confirmKey={confirmKey}
          cancelKey={cancelKey}
        />
      )}

      {showSettings && (
        <SettingsPanel
          scanKey={scanKey}
          confirmKey={confirmKey}
          cancelKey={cancelKey}
          onScanKeyChange={setScanKey}
          onConfirmKeyChange={setConfirmKey}
          onCancelKeyChange={setCancelKey}
          onClose={() => setShowSettings(false)}
        />
      )}

      {showRosterComplete && (
        <RosterCompleteModal
          roster={roster}
          orderNumber={orderNumber}
          flagCount={flagCount}
          onExport={() => { handleExportAndSave(roster, orderNumber, operatorName, binMap, packGroupId); setShowRosterComplete(false); }}
          onDismiss={() => setShowRosterComplete(false)}
          onNewOrder={handleNewSession}
        />
      )}

      {/* Header */}
      <div style={{ background: "#161b22", borderBottom: "1px solid #21262d", padding: "8px 16px", display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8, flexWrap: "wrap", flexShrink: 0 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
          <img src={akLogo} alt="AK" style={{ height: 32, opacity: 0.95 }} />
          <div>
            <div style={{ fontWeight: 800, fontSize: 14 }}>🏭 Jersey QC Scanner</div>
            {rosterFile && (
              <div style={{ fontSize: 10, color: "#64748b" }}>
                {rosterFile}{orderNumber ? ` · Order: ${orderNumber}` : ""}{operatorName ? ` · ${operatorName}` : ""}
              </div>
            )}
          </div>
          <input value={orderNumber} onChange={e => setOrderNumber(e.target.value)} placeholder="Order number…"
            style={{ padding: "4px 9px", borderRadius: 6, border: "1px solid #334155", background: "#1e293b", color: "#e2e8f0", fontSize: 12, width: 160 }} />
        </div>
        <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
          <Pill color="#22c55e">{passCount} Pass</Pill>
          <Pill color="#ef4444">{flagCount} Flag</Pill>
          {resolvedCount > 0 && <Pill color="#94a3b8">{resolvedCount} Resolved</Pill>}
          <button onClick={() => setShowSettings(true)} style={{ ...btnGhost, padding: "3px 9px" }}>⚙</button>
          <button onClick={() => setShowAdmin(true)} style={{ ...btnGhost, padding: "3px 9px", fontSize: 11 }}>🔒 Admin</button>
          {sessionStarted && (
            <button onClick={handleNewSession} style={{ ...btnGhost, padding: "3px 9px", fontSize: 11 }}>↩ New Session</button>
          )}
        </div>
      </div>

      {/* Progress bar */}
      {roster.length > 0 && (
        <div style={{ height: 4, background: "#1e293b", flexShrink: 0 }}>
          <div style={{ height: "100%", background: rosterComplete ? "#22c55e" : "#3b82f6", width: `${origRoster.length > 0 ? (scanned.filter(r => !r._extra).length / origRoster.length) * 100 : 0}%`, transition: "width 0.4s ease" }} />
        </div>
      )}

      {/* Main layout */}
      <div style={{ display: "flex", flex: 1, overflow: "hidden" }}>

        {/* LEFT: Camera */}
        <div style={{ width: 580, flexShrink: 0, display: "flex", flexDirection: "column", borderRight: "1px solid #21262d", background: "#0d1117" }}>
          <canvas ref={canvasRef} style={{ display: "none" }} />

          {roster.length === 0 ? (
            <div style={{ flex: 1, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", padding: 32, textAlign: "center" }}>
              <div style={{ fontSize: 36, marginBottom: 10 }}>📋</div>
              <div style={{ fontWeight: 700, fontSize: 16, marginBottom: 6 }}>No Roster Loaded</div>
              <div style={{ fontSize: 12, color: "#64748b", marginBottom: 16 }}>
                CSV or Excel: <code style={codeSt}>name</code> <code style={codeSt}>number</code> <code style={codeSt}>team</code> <code style={codeSt}>size</code>
              </div>
              <button onClick={() => rosterRef.current.click()} style={{ ...btnPri, fontSize: 14, padding: "10px 24px" }}>
                Upload Roster ({xlsxReady ? "CSV or Excel" : "CSV"})
              </button>
              <input ref={rosterRef} type="file" accept=".csv,.xlsx,.xls,.tsv,.txt" style={{ display: "none" }} onChange={handleRosterUpload} />
              {error && <p style={{ fontSize: 12, color: "#f87171", marginTop: 10 }}>{error}</p>}
            </div>
          ) : (
            <>
              <div style={{ display: "flex", borderBottom: "1px solid #21262d" }}>
                {["camera", "upload"].map(m => (
                  <button key={m} onClick={() => { setInputMode(m); if (m === "upload") stopCamera(); }}
                    style={{ flex: 1, padding: "8px 0", border: "none", background: inputMode === m ? "rgba(59,130,246,0.1)" : "transparent", color: inputMode === m ? "#60a5fa" : "#64748b", fontWeight: 600, fontSize: 12, borderBottom: inputMode === m ? "2px solid #3b82f6" : "2px solid transparent", cursor: "pointer" }}>
                    {m === "camera" ? "📷 Webcam" : "🖼 Upload"}
                  </button>
                ))}
              </div>

              {inputMode === "camera" && (
                <div style={{ flex: 1, display: "flex", flexDirection: "column" }}>
                  <div style={{ position: "relative", background: "#000", flex: 1 }}>
                    <video ref={videoRef} autoPlay playsInline muted
                      style={{ width: "100%", height: "100%", objectFit: "cover", display: cameraOn ? "block" : "none", transform: "rotate(180deg)" }} />
                    {!cameraOn && (
                      <div style={{ position: "absolute", inset: 0, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", gap: 8 }}>
                        <span style={{ fontSize: 40 }}>📷</span>
                        <span style={{ fontSize: 13, color: "#64748b" }}>Camera off</span>
                      </div>
                    )}
                    {cameraOn && (
                      <div style={{ position: "absolute", inset: 0, display: "flex", alignItems: "center", justifyContent: "center", pointerEvents: "none" }}>
                        <div style={{ width: 200, height: 120, border: "2px solid rgba(255,255,255,0.35)", borderRadius: 6, boxShadow: "0 0 0 9999px rgba(0,0,0,0.25)" }} />
                      </div>
                    )}
                    {scanning && (
                      <div style={{ position: "absolute", inset: 0, background: "rgba(59,130,246,0.15)", display: "flex", alignItems: "center", justifyContent: "center" }}>
                        <div style={{ color: "#fff", fontWeight: 700, background: "rgba(0,0,0,0.6)", padding: "8px 20px", borderRadius: 20 }}>⏳ Analysing…</div>
                      </div>
                    )}
                    {cameraOn && !scanning && (
                      <>
                        <div style={{ position: "absolute", top: 8, right: 8, display: "flex", alignItems: "center", gap: 5, background: "rgba(0,0,0,0.55)", padding: "3px 9px", borderRadius: 20 }}>
                          <div style={{ width: 7, height: 7, borderRadius: "50%", background: "#22c55e", boxShadow: "0 0 5px #22c55e" }} />
                          <span style={{ fontSize: 10, color: "#e2e8f0" }}>Ready</span>
                        </div>
                        <div style={{ position: "absolute", bottom: 8, left: "50%", transform: "translateX(-50%)", background: "rgba(0,0,0,0.55)", color: "#94a3b8", fontSize: 10, padding: "3px 9px", borderRadius: 12, whiteSpace: "nowrap" }}>
                          Press <kbd style={{ background: "#1e293b", padding: "1px 4px", borderRadius: 3, fontSize: 10 }}>{keyLabel}</kbd> to scan
                        </div>
                      </>
                    )}
                  </div>
                  <div style={{ padding: "10px 12px", display: "flex", gap: 8, borderTop: "1px solid #21262d" }}>
                    {!cameraOn
                      ? <button onClick={startCamera} style={{ ...btnPri, flex: 1, fontSize: 15, padding: "12px" }}>Start Camera</button>
                      : <>
                          <button onClick={stopCamera} style={{ ...btnGhost, width: 42 }}>■</button>
                          <button onClick={doTriggerScan} disabled={scanning}
                            style={{ ...btnPri, flex: 1, fontSize: 18, padding: "12px", opacity: scanning ? 0.4 : 1 }}>
                            {scanning ? "Scanning…" : "⚡ SCAN"}
                          </button>
                        </>
                    }
                  </div>
                  {error && <p style={{ padding: "0 12px 8px", fontSize: 11, color: "#f87171" }}>{error}</p>}
                </div>
              )}

              {inputMode === "upload" && (
                <div style={{ flex: 1, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", padding: 20, gap: 12 }}>
                  {thumb
                    ? <img src={thumb} alt="uploaded" style={{ width: "100%", borderRadius: 8, maxHeight: 300, objectFit: "contain" }} />
                    : <div style={{ width: "100%", aspectRatio: "4/3", background: "#0d1117", borderRadius: 8, border: "2px dashed #334155", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", gap: 8, color: "#475569" }}>
                        <span style={{ fontSize: 36 }}>🖼</span>
                        <span style={{ fontSize: 12 }}>No photo selected</span>
                      </div>
                  }
                  <input ref={photoRef} type="file" accept="image/*" style={{ display: "none" }} onChange={handlePhotoUpload} />
                  <button onClick={() => photoRef.current.click()} disabled={scanning}
                    style={{ ...btnPri, width: "100%", fontSize: 16, padding: "12px", opacity: scanning ? 0.5 : 1 }}>
                    {scanning ? "⏳ Analysing…" : "⚡ Upload & Scan"}
                  </button>
                </div>
              )}

              {lastConfirmed && (
                <div style={{ padding: "7px 12px", borderTop: "1px solid #21262d", display: "flex", alignItems: "center", gap: 10, background: "rgba(34,197,94,0.05)", flexShrink: 0 }}>
                  <span style={{ fontSize: 10, color: "#64748b", whiteSpace: "nowrap" }}>Last:</span>
                  <span style={{ fontWeight: 700, fontSize: 13 }}>#{lastConfirmed.match?.number}{lastConfirmed.match?.name ? ` · ${lastConfirmed.match.name}` : ""}</span>
                  {lastConfirmed.match?.size && <span style={{ fontSize: 11, color: "#22c55e", fontWeight: 700 }}>{lastConfirmed.match.size}</span>}
                  {binMap && lastConfirmed.match?.team && binMap[lastConfirmed.match.team] && (
                    <span style={{ fontSize: 11, color: "#60a5fa", fontWeight: 700 }}>Bin {binMap[lastConfirmed.match.team]}</span>
                  )}
                  <span style={{ fontSize: 10, color: "#475569", marginLeft: "auto" }}>{lastConfirmed.timestamp}</span>
                </div>
              )}

              {lastResult && lastResult.status !== S_PASS && !overlay && (
                <div style={{ padding: "0 10px 10px", flexShrink: 0 }}>
                  <ResultCard result={lastResult} onResolve={resolveFlag} />
                </div>
              )}

              <div style={{ padding: "7px 12px", borderTop: "1px solid #21262d", background: "#0d1117", flexShrink: 0 }}>
                <div style={{ display: "flex", flexWrap: "wrap", gap: "4px 14px" }}>
                  {[[keyLabel, "Scan"], [formatKey(confirmKey), "Confirm"], [formatKey(cancelKey), "Navigate/Retry"], ["↑↓", "Navigate list"]].map(([k, d]) => (
                    <div key={k} style={{ display: "flex", gap: 5, alignItems: "center" }}>
                      <kbd style={{ background: "#1e293b", border: "1px solid #334155", padding: "1px 6px", borderRadius: 4, fontSize: 10 }}>{k}</kbd>
                      <span style={{ fontSize: 10, color: "#64748b" }}>{d}</span>
                    </div>
                  ))}
                </div>
              </div>
            </>
          )}
        </div>

        {/* RIGHT: Info panel */}
        <div style={{ flex: 1, display: "flex", flexDirection: "column", overflow: "hidden" }}>

          <div style={{ display: "grid", gridTemplateColumns: "repeat(5,1fr)", borderBottom: "1px solid #21262d", flexShrink: 0 }}>
            {[
              { label: "Scanned",   value: `${scanned.filter(r => !r._extra).length} / ${origRoster.length}`, color: "#e2e8f0" },
              { label: "Remaining", value: remaining,  color: remaining === 0 ? "#22c55e" : "#f59e0b" },
              { label: "Flagged",   value: flagCount,  color: flagCount  > 0 ? "#ef4444" : "#64748b" },
              { label: "Extra",     value: extraCount, color: extraCount > 0 ? "#f59e0b" : "#64748b" },
              { label: "ETA",       value: etaMin ? `${etaMin}m` : "—", color: "#64748b" },
            ].map(s => (
              <div key={s.label} style={{ padding: "10px 0", textAlign: "center", borderRight: "1px solid #21262d" }}>
                <div style={{ fontSize: 22, fontWeight: 800, color: s.color, lineHeight: 1.2 }}>{s.value}</div>
                <div style={{ fontSize: 10, color: "#64748b", textTransform: "uppercase", letterSpacing: 0.5 }}>{s.label}</div>
              </div>
            ))}
          </div>

          <div style={{ display: "flex", borderBottom: "1px solid #21262d", flexShrink: 0 }}>
            {["roster", "log"].map(v => (
              <button key={v} onClick={() => setView(v)} style={{ flex: 1, padding: "8px 0", border: "none", background: "transparent", color: view === v ? "#3b82f6" : "#64748b", fontWeight: 600, fontSize: 12, borderBottom: view === v ? "2px solid #3b82f6" : "2px solid transparent", cursor: "pointer", textTransform: "capitalize" }}>
                {v === "log" ? `Log (${log.length})` : `Roster (${origRoster.length}${extraCount > 0 ? ` +${extraCount} extra` : ""})`}
              </button>
            ))}
            <div style={{ display: "flex", alignItems: "center", gap: 6, padding: "0 10px", borderLeft: "1px solid #21262d" }}>
              {roster.length > 0 && (
                <>
                  <button onClick={() => rosterRef.current.click()} style={{ ...btnGhost, padding: "3px 8px", fontSize: 11 }}>Replace</button>
                  <input ref={rosterRef} type="file" accept=".csv,.xlsx,.xls,.tsv,.txt" style={{ display: "none" }} onChange={handleRosterUpload} />
                  <button onClick={() => handleExportAndSave(roster, orderNumber, operatorName, binMap, packGroupId)}
                    style={{ ...btnPri, padding: "3px 10px", fontSize: 11, background: rosterComplete ? "#22c55e" : "#3b82f6" }}>
                    {rosterComplete ? "✅ Export" : "⬇ Export"}
                  </button>
                </>
              )}
            </div>
          </div>

          {view === "roster" && (
            <div style={{ flex: 1, overflowY: "auto" }}>
              {roster.length === 0
                ? <div style={{ padding: 32, textAlign: "center", color: "#64748b" }}>No roster loaded.</div>
                : (
                  <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                    <thead style={{ position: "sticky", top: 0, background: "#161b22", zIndex: 1 }}>
                      <tr>
                        <th style={{ ...thSt, minWidth: 90 }}>Status</th>
                        {rosterCols.map(c => <th key={c} style={thSt}>{c}</th>)}
                        {binMap && <th style={thSt}>Bin</th>}
                        <th style={{ ...thSt, minWidth: 160 }}>Issue</th>
                      </tr>
                    </thead>
                    <tbody>
                      {roster.map((r, i) => (
                        <tr key={i} style={{
                          background: r._extra ? "rgba(245,158,11,0.10)" : r.scanned === "pass" ? "rgba(34,197,94,0.07)" : r.scanned === "flag" ? "rgba(239,68,68,0.07)" : r.scanned === "resolved" ? "rgba(148,163,184,0.07)" : "transparent",
                          borderBottom: r._extra ? "1px solid rgba(245,158,11,0.25)" : "1px solid #1e293b",
                        }}>
                          <td style={{ ...tdSt, whiteSpace: "nowrap" }}>
                            <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
                              <span style={{ minWidth: 16, textAlign: "center" }}>
                                {r._extra ? "➕" : r.scanned === "pass" ? "✅" : r.scanned === "flag" ? "🚩" : r.scanned === "resolved" ? "🔧" : <span style={{ color: "#475569" }}>—</span>}
                              </span>
                              {r._extra && <span style={{ fontSize: 9, fontWeight: 800, color: "#f59e0b", letterSpacing: 0.5, textTransform: "uppercase" }}>EXTRA</span>}
                              {!r._extra && (!r.scanned || r.scanned === false) && (
                                <button onClick={() => handleRosterEdit(r, "pass")} title="Mark as passed"
                                  style={{ padding: "1px 5px", borderRadius: 4, border: "1px solid rgba(34,197,94,0.4)", background: "rgba(34,197,94,0.1)", color: "#22c55e", fontSize: 10, fontWeight: 700, cursor: "pointer" }}>✓</button>
                              )}
                              {!r._extra && (!r.scanned || r.scanned === false || r.scanned === "flag") && (
                                <button onClick={() => handleRosterEdit(r, r.scanned === "flag" ? "pass" : "flag")} title={r.scanned === "flag" ? "Mark as passed" : "Mark as flagged"}
                                  style={{ padding: "1px 5px", borderRadius: 4, border: `1px solid ${r.scanned === "flag" ? "rgba(34,197,94,0.4)" : "rgba(239,68,68,0.4)"}`, background: r.scanned === "flag" ? "rgba(34,197,94,0.1)" : "rgba(239,68,68,0.1)", color: r.scanned === "flag" ? "#22c55e" : "#ef4444", fontSize: 10, fontWeight: 700, cursor: "pointer" }}>
                                  {r.scanned === "flag" ? "✓" : "🚩"}
                                </button>
                              )}
                              {r.scanned && r.scanned !== false && (
                                <button onClick={() => handleRosterEdit(r, "undo")} title={r._extra ? "Remove extra entry" : "Undo"}
                                  style={{ padding: "1px 5px", borderRadius: 4, border: "1px solid rgba(100,116,139,0.4)", background: "rgba(100,116,139,0.1)", color: "#94a3b8", fontSize: 10, fontWeight: 700, cursor: "pointer" }}>↩</button>
                              )}
                            </div>
                          </td>
                          {rosterCols.map(c => (
                            <td key={c} style={{ ...tdSt, fontWeight: c === "size" ? 700 : 400, color: c === "size" ? "#f59e0b" : "#e2e8f0" }}>
                              {r[c]}
                            </td>
                          ))}
                          {binMap && (
                            <td style={{ ...tdSt, color: "#60a5fa", fontWeight: 700 }}>
                              {r.team && binMap[r.team] ? `Bin ${binMap[r.team]}` : "—"}
                            </td>
                          )}
                          <td style={{ ...tdSt, minWidth: 160 }}>
                            <select
                              value={r.comment || ""}
                              onChange={e => handleRosterComment(r._id, e.target.value)}
                              style={{ width: "100%", padding: "5px 8px", borderRadius: 5, border: "1px solid #4b5563", background: "#1e293b", color: r.comment ? "#f1f5f9" : "#94a3b8", fontSize: 14, fontWeight: r.comment ? 600 : 400, cursor: "pointer" }}>
                              <option value="">—</option>
                              <option value="Label">Label</option>
                              <option value="Construction/Sewing">Construction/Sewing</option>
                              <option value="Artwork/Logo">Artwork/Logo</option>
                              <option value="Decoration">Decoration</option>
                              <option value="Missing Jersey">Missing Jersey</option>
                              <option value="Extra Jersey">Extra Jersey</option>
                            </select>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                )
              }
            </div>
          )}

          {view === "log" && (
            <div style={{ flex: 1, overflowY: "auto", padding: 10 }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
                <span style={{ fontSize: 12, color: "#94a3b8" }}>{log.length} scans this session</span>
                {log.length > 0 && (
                  <button onClick={() => exportLogCSV(log, orderNumber, binMap)} style={{ ...btnGhost, fontSize: 11 }}>⬇ Export Log</button>
                )}
              </div>
              {log.length === 0
                ? <div style={{ textAlign: "center", padding: 28, color: "#64748b" }}>No scans yet.</div>
                : log.map(l => (
                  <div key={l.id} style={{ ...card, borderLeft: `3px solid ${statusColor(l.status)}`, padding: "10px 12px" }}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
                      <div>
                        <span style={{ fontSize: 11, fontWeight: 700, color: statusColor(l.status), textTransform: "uppercase", letterSpacing: 1 }}>{l.status}</span>
                        <div style={{ fontWeight: 700, fontSize: 14, marginTop: 1 }}>#{l.detected?.number || "?"} · {l.detected?.name || "Unknown"}</div>
                        {l.match && (
                          <div style={{ fontSize: 11, color: "#94a3b8" }}>
                            Matched: {l.match.name || "—"} #{l.match.number}
                            {l.match.size ? ` · ${l.match.size}` : ""}
                            {l.match.team ? ` · ${l.match.team}` : ""}
                            {binMap && l.match.team && binMap[l.match.team] ? ` · Bin ${binMap[l.match.team]}` : ""}
                          </div>
                        )}
                        {l.reason     && <div style={{ fontSize: 11, color: "#f59e0b", marginTop: 2 }}>⚠ {l.reason}</div>}
                        {l.resolution && <div style={{ fontSize: 11, color: "#94a3b8", marginTop: 2 }}>Resolution: {l.resolution}</div>}
                      </div>
                      <div style={{ fontSize: 10, color: "#475569", marginLeft: 8, whiteSpace: "nowrap" }}>{l.timestamp}</div>
                    </div>
                  </div>
                ))
              }
            </div>
          )}

          {remaining > 0 && roster.length > 0 && (
            <div style={{ borderTop: "1px solid #21262d", padding: "8px 12px", maxHeight: 110, overflowY: "auto", flexShrink: 0 }}>
              <div style={{ fontSize: 10, fontWeight: 700, color: "#64748b", textTransform: "uppercase", letterSpacing: 0.5, marginBottom: 5 }}>
                Remaining ({remaining})
              </div>
              <div style={{ display: "flex", flexWrap: "wrap", gap: 4 }}>
                {roster.filter(r => !r.scanned || r.scanned === false).map(r => (
                  <div key={r._id} style={{ background: "#1e293b", border: "1px solid #334155", borderRadius: 5, padding: "2px 8px", fontSize: 11, fontWeight: 600 }}>
                    #{r.number}{r.name ? ` ${r.name}` : ""}{r.size ? ` (${r.size})` : ""}
                  </div>
                ))}
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}