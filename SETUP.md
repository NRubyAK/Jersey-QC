# Jersey QC Scanner — Setup & Maintenance Guide

This guide covers how to install, configure, run, and maintain the Jersey QC Scanner on a new machine or after making code changes.

---

## Prerequisites

| Requirement | Notes |
|---|---|
| Node.js 18+ | Download from https://nodejs.org |
| npm | Comes with Node.js |
| Anthropic API key | Obtain from https://console.anthropic.com |
| A device with a camera | Webcam, laptop camera, or phone (accessed via browser) |
| Modern browser | Chrome or Edge recommended; must support `getUserMedia` |

---

## First-time installation

```bash
# 1. Navigate to the project directory
cd C:\Users\noamr\jersey-qc

# 2. Install all dependencies
npm install
```

This installs React, Vite, Express, and all dev tools listed in `package.json`.

---

## Environment setup

Create a file named `.env.local` in the project root (it is already in `.gitignore` and will not be committed):

```
VITE_CLAUDE_API_KEY=sk-ant-xxxxxxxxxxxxxxxx
```

Replace the value with your actual Anthropic API key. Without this, all scans will fail with an API error.

> **Security note:** The API key is sent from the browser to `https://api.anthropic.com` via the Vite proxy. It is exposed in the browser's network tab. This is acceptable on a controlled factory LAN but not suitable for a public-facing deployment.

---

## Running the app

```bash
npm start
```

This command (defined in `package.json`) runs two processes concurrently:
- **Vite dev server** on `https://localhost:5173` — serves the React frontend
- **Express server** on `http://localhost:3001` — handles admin/station API and persists completed orders

The Vite server proxies:
- `/api/anthropic/*` → `https://api.anthropic.com` (Claude API calls)
- `/api/admin/*` → `http://localhost:3001`
- `/api/station/*` → `http://localhost:3001`

Open `https://localhost:5173` in a browser. On first load you will be prompted to set a station name (e.g. "Station 1"). This is saved to `localStorage` and only needs to be done once per device.

> **HTTPS note:** The Vite config uses `@vitejs/plugin-basic-ssl` which generates a self-signed certificate. You will need to accept the browser security warning on first visit. This is required for `getUserMedia` (camera access) to work on non-localhost origins.

---

## Running individual processes

If you need to run them separately for debugging:

```bash
# Frontend only (no admin panel backend)
npm run dev

# Backend only
npm run server
```

---

## Multi-station / network setup

To use on multiple devices on the same network:

1. Run `npm start` on **one machine** (the server machine)
2. Find that machine's local IP address (e.g. `192.168.1.50`)
3. On each scanning station, open `https://192.168.1.50:5173` in a browser
4. Accept the SSL warning
5. Each station will self-register when it sends its first heartbeat

The Vite dev server is configured with `host: true` which means it binds to all network interfaces. Station heartbeats and admin assignments all route through the one server machine.

---

## Admin password

The admin panel password is currently hardcoded in `server.js` line 9:

```js
const ADMIN_PASSWORD = 'Bernard2025';
```

To change it, edit this line directly. For a more secure setup, move it to an environment variable:

1. In `.env.local` add:
   ```
   ADMIN_PASSWORD=your-new-password
   ```
2. In `server.js`, change:
   ```js
   const ADMIN_PASSWORD = process.env.ADMIN_PASSWORD || 'Bernard2025';
   ```

---

## Data storage

Completed orders and pack group definitions are stored as JSON files in the `data/` directory:

```
data/
├── export-index.json          ← index of all completed orders (summary metadata)
├── exports/
│   └── {id}.json              ← full order data: roster + scan log
└── packgroups/
    ├── packgroup-index.json   ← index of all pack groups
    └── {id}.json              ← pack group definition + list of order IDs
```

- **Backups:** Copy the entire `data/` directory to back up all historical scan records.
- **Clearing history:** Delete files in `data/exports/` and reset `data/export-index.json` to `[]`. Same pattern for pack groups.
- **Station state is not persisted.** Active stations, their current progress, and assigned-but-not-yet-accepted rosters are held in memory and lost on server restart. Operators must re-start their session after a server restart.

---

## Updating the Claude model

The AI model is set at the top of `src/App.jsx`:

```js
const CLAUDE_MODEL = "claude-haiku-4-5-20251001";
```

To switch to a newer or different model, update this constant. Use a Haiku-class model for speed and cost efficiency. Avoid larger models — scan latency scales directly with model response time.

---

## Building for production

```bash
npm run build
```

This outputs a static bundle to `dist/`. To serve in production:
1. Serve the `dist/` folder with any static web server (nginx, etc.)
2. Set up a reverse proxy so `/api/anthropic/*` routes to `https://api.anthropic.com` and `/api/*` routes to your Express server
3. Run `node server.js` separately (consider `pm2` for process management)
4. Set `VITE_CLAUDE_API_KEY` as a server-side environment variable (not client-side) for production deployments

---

## Linting

```bash
npm run lint
```

Uses ESLint with `eslint-plugin-react-hooks` and `eslint-plugin-react-refresh`.

---

## Common issues

| Problem | Likely cause | Fix |
|---|---|---|
| "Camera access denied" | Browser blocked camera | Click the camera icon in the address bar and allow access |
| All scans return API error | Missing or invalid API key | Check `.env.local` has `VITE_CLAUDE_API_KEY=sk-ant-...` |
| Excel import fails | SheetJS CDN not loaded | Check internet connection; CDN URL is `https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js` |
| Admin panel says "Server unreachable" | Express server not running | Run `npm run server` in a separate terminal, or use `npm start` |
| Station not appearing in admin dashboard | Station hasn't sent a heartbeat | Open the station URL in a browser; it will appear within 15 seconds |
| Session lost after page refresh | Expected behaviour | Session is stored in `localStorage` and will be restored on reload; camera must be restarted manually |
| SSL warning on network access | Self-signed certificate | Accept the warning in the browser — this is safe on a LAN |

---

## Maintenance checklist

- **Monthly:** Archive and delete old entries in `data/exports/` to keep the export index fast
- **On API key rotation:** Update `VITE_CLAUDE_API_KEY` in `.env.local` and restart `npm start`
- **On dependency update:** Run `npm install` after pulling changes; check for breaking changes in React, Vite, or Express changelogs
- **On roster column changes:** The column detection is automatic (case-insensitive) but verify new column names are being picked up correctly after roster format changes
