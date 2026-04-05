# Jersey QC Scanner — Agent Onboarding

## What this project does

A React/Vite web app for quality-control scanning of sports jerseys in a manufacturing environment. An operator photographs each jersey with a device camera; Claude Haiku reads the player name and number from the image and matches it against an uploaded roster.

**Core workflow:**
1. Operator enters their name, order number, and uploads a roster file (CSV or Excel)
2. If the roster has multiple teams, a bin assignment screen appears (Bin 1 = Team A, Bin 2 = Team B, etc.)
3. Camera feed is shown; operator presses a configurable key (default: `B`) to scan a jersey
4. Claude Haiku extracts name + number from the photo via the Anthropic API
5. Match result shows as a full-screen overlay:
   - **Green** = exact match → confirm with `B`, or edit with `A`
   - **Yellow** = multiple candidates or close match → pick from list
   - **Red** = not found or API error → flag and continue
6. Flagged items appear in the log and can be resolved (Accept / Reject / Note)
7. When all roster entries are scanned, a completion modal offers Excel export

---

## Project location

`C:\Users\noamr\jersey-qc`

---

## Tech stack

| Layer | Technology |
|---|---|
| Frontend | React 19, Vite 8, `src/App.jsx` (~2100 lines, all inline styles) |
| Backend | Express 5 (`server.js`) on port 3001 |
| AI | Claude Haiku `claude-haiku-4-5-20251001` via direct browser fetch |
| Excel | SheetJS from CDN |
| Proxy | Vite proxies `/api/anthropic/*` → Anthropic, `/api/admin/*` and `/api/station/*` → Express |
| Persistence | Flat JSON files in `data/` (exports + pack groups) |
| Session | `localStorage` (`jerseyqc_session`, `jerseyqc_station`) |

---

## Key files

| File | Purpose |
|---|---|
| `src/App.jsx` | Entire frontend — all components, state, logic |
| `server.js` | Express API (admin auth, station heartbeats, data persistence) |
| `vite.config.js` | Vite config + proxy rules |
| `.env.local` | `VITE_CLAUDE_API_KEY=sk-ant-...` (gitignored, must exist to scan) |
| `data/export-index.json` | Index of completed orders |
| `data/exports/{id}.json` | Full order data (roster + log) |
| `data/packgroups/{id}.json` | Pack group definitions |
| `SETUP.md` | Developer/admin setup and maintenance guide |
| `OPERATOR_GUIDE.md` | End-user operator guide |

---

## Running locally

```bash
npm install
npm start        # starts both Vite (port 5173) and Express (port 3001) concurrently
```

Or separately:
```bash
npm run dev      # Vite only
npm run server   # Express only
```

---

## Environment variables

| Variable | File | Required |
|---|---|---|
| `VITE_CLAUDE_API_KEY` | `.env.local` | Yes — all scans fail without it |

---

## Component tree

```
App()                       ← root; ALL state lives here
├── StationSetupModal       ← first-time station name prompt
├── SessionStartModal       ← per-session: operator, order, roster upload / admin queue
├── BinSetupModal           ← pre-scan bin labelling (multi-team orders)
├── ScanOverlay             ← full-screen result: green/yellow/red
├── SettingsPanel           ← keyboard shortcut editor
├── RosterCompleteModal     ← end-of-order summary + export
└── AdminPanel              ← supervisor tools
    ├── AdminDashboard      ← live station status cards
    ├── AdminAssign         ← push roster to a station
    ├── AdminPackGroups     ← create/view pack groups
    └── AdminExports        ← download historical exports
```

---

## Scan result types from `findMatches()`

| Type | Meaning | Overlay colour |
|---|---|---|
| `exact` | One roster entry matches both name and number | Green |
| `number_conflict` | Multiple entries match the number | Yellow (pick) |
| `size_pick` | Number-only roster with multiple same numbers (different sizes) | Yellow (pick) |
| `close` | Fuzzy match — substring match on name or number | Yellow (close) |
| `none` | No match found | Red |

---

## Server API routes

| Method | Path | Auth | Purpose |
|---|---|---|---|
| POST | `/api/admin/login` | — | Validate admin password |
| GET | `/api/admin/stations` | admin | All station states |
| POST | `/api/admin/assign` | admin | Push roster to station queue |
| POST | `/api/station/heartbeat` | — | Station ping; returns assigned rosters |
| POST | `/api/station/roster-accepted` | — | Dequeue accepted roster |
| POST | `/api/station/complete` | — | Save completed order to disk |
| GET | `/api/admin/exports` | admin | Export index |
| GET | `/api/admin/exports/:id` | admin | Full export data |
| GET | `/api/admin/packgroups` | admin | Pack group index |
| POST | `/api/admin/packgroups` | admin | Create pack group |
| GET | `/api/admin/packgroups/:id` | admin | Pack group detail |
| GET | `/api/admin/packgroups/:id/combined` | admin | Combined order data |

Admin auth uses the `x-admin-password` header. Password is hardcoded in `server.js:9` as `"Bernard2025"`.

---

## Roster file format

Supported: `.csv`, `.tsv`, `.txt`, `.xlsx`, `.xls`

Expected columns (case-insensitive, auto-detected):
- `name` — player name
- `number` — jersey number
- `team` — team name (optional; triggers bin assignment when multiple teams)
- `size` — size (optional)
- `size range` + `size` → auto-merged into `A-XS` / `Y-M` style combined size
- `team name` → auto-renamed to `team`

---

## Scan status values

| Value | Meaning |
|---|---|
| `false` | Not scanned |
| `"pass"` | Confirmed match |
| `"flag"` | Flagged for review |
| `"resolved"` | Flag resolved by operator |
| `"extra"` | Scanned but not in roster |

---

## Key architecture notes

- All state in the top-level `App` component; no external state management
- Camera canvas rotated 180° before API call — the physical camera is mounted upside-down on the scanning rig
- `findMatches()` handles exact, conflict (same number, different players), close/fuzzy matching
- Keyboard shortcuts (scan/confirm/cancel) are fully configurable via Settings panel; designed for foot pedals or barcode scanners
- Audio feedback: ascending tone = pass, descending = flag, short beep = scan trigger
- Station heartbeat: every 15s when idle polling for admin assignments, every 2s (debounced) when scanning to report progress
- `exportRosterXLSX()` writes status + all roster columns to Excel; `exportLogCSV()` writes the scan event log
- Session persisted to `localStorage` — survives page refresh; clears on "New Session"
- Station state (active sessions, queued assignments) is **in-memory only** — lost on server restart

---

## Known issues / decisions

- `App.css` is unused Vite template boilerplate (safe to delete)
- Admin password hardcoded in `server.js:9` — should ideally be an env var
- Anthropic API called directly from browser with `anthropic-dangerous-direct-browser-access: true` — acceptable for a controlled factory LAN
- SheetJS loaded from CDN — Excel features fail when offline

---

## Security note

The API key is in `.env.local` (gitignored). It was previously hardcoded and has been revoked. Never hardcode it again. GitHub push protection will block commits containing Anthropic API keys.
