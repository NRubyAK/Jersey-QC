# Jersey QC Scanner

## What this project does

A React/Vite web app for quality-control scanning of sports jerseys in a manufacturing environment. An operator photographs each jersey with their device camera; Claude Haiku reads the player name and number from the image and matches it against an uploaded roster.

**Workflow:**
1. Operator enters their name, order number, and uploads a roster file (CSV/TSV or Excel)
2. If the roster has multiple teams, a bin assignment screen appears (Bin 1 = Team A, Bin 2 = Team B, etc.)
3. Camera feed is shown; operator presses a configurable key (default: `B`) to scan a jersey
4. Claude Haiku extracts name + number from the photo via the Anthropic API
5. Match result shows as a full-screen overlay:
   - **Green** = exact match → confirm with `B`, or edit with `A`
   - **Yellow** = multiple candidates or close match → pick from list
   - **Red** = not found or API error → flag and continue
6. Flagged items appear in the log and can be resolved (Accept / Reject / Note)
7. When all roster entries are scanned, a completion modal offers Excel export

## Tech stack

- React 19, Vite 8 (ESM)
- Single component file: `src/App.jsx` (~900+ lines, all inline styles, dark GitHub-like theme)
- SheetJS loaded from CDN (`https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js`) for Excel read/write
- Anthropic API (Claude Haiku `claude-haiku-4-5-20251001`) called **directly from the browser** via `fetch("/api/anthropic/v1/messages")` — requires a Vite dev proxy or a hosting proxy in production

## Environment variables

| Variable | Purpose |
|---|---|
| `VITE_CLAUDE_API_KEY` | Anthropic API key — **never hardcode this** |

Create a `.env.local` file (already gitignored via `*.local`):
```
VITE_CLAUDE_API_KEY=sk-ant-...
```

## Running locally

```bash
npm install
npm run dev
```

The dev server starts at `http://localhost:5173`.

> **Note:** The Anthropic API is called via `/api/anthropic/...`. In dev, Vite must proxy this path to `https://api.anthropic.com`. Add a proxy in `vite.config.js` if not already present:
> ```js
> server: {
>   proxy: {
>     '/api/anthropic': {
>       target: 'https://api.anthropic.com',
>       changeOrigin: true,
>       rewrite: path => path.replace(/^\/api\/anthropic/, ''),
>     }
>   }
> }
> ```

## Roster file format

Supported: `.csv`, `.tsv`, `.txt`, `.xlsx`, `.xls`

Expected columns (case-insensitive, auto-detected):
- `name` — player name
- `number` — jersey number
- `team` — team name (optional, used for bin assignment)
- `size` — size (optional)
- `size range` + `size` — auto-merged into `A-XS` / `Y-M` style combined size

## Key architecture notes

- All state lives in the top-level `App` component; no external state management
- `findMatches()` handles exact, conflict (same number, different players), close/fuzzy matching
- Keyboard shortcuts (scan/confirm/cancel keys) are fully configurable via Settings panel
- Designed for use with a barcode scanner or foot pedal mapped to keyboard keys
- Audio feedback: ascending tone = pass, descending = flag, short beep = scan trigger
- `exportRosterXLSX()` writes status + all roster columns to Excel; `exportLogCSV()` writes the scan event log

## Security note

The API key was previously hardcoded in `src/App.jsx:4` and has been **revoked**. Always use `VITE_CLAUDE_API_KEY` from environment. GitHub push protection will block any future commit containing an Anthropic API key.
