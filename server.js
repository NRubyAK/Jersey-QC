/**
 * Jersey QC Scanner — Express API Server
 *
 * Runs on port 3001. The Vite dev server proxies:
 *   /api/admin/*   → this server
 *   /api/station/* → this server
 *
 * Responsibilities:
 *   - Admin authentication (simple shared password)
 *   - Station registration and heartbeat tracking (in-memory, resets on restart)
 *   - Queuing admin-assigned rosters to specific stations
 *   - Persisting completed orders to disk as JSON files
 *   - Serving export history and pack group data to the admin panel
 *
 * Data is stored as flat JSON files in ./data/:
 *   data/exports/{id}.json        — full order data (roster + scan log)
 *   data/export-index.json        — summary metadata for all exports (fast listing)
 *   data/packgroups/{id}.json     — pack group definitions
 *   data/packgroup-index.json     — summary metadata for all pack groups
 *
 * NOTE: Station state (active sessions, queued assignments) is held in memory only.
 * Restarting this server clears all station state. Operators must re-start their
 * session if the server restarts mid-order.
 *
 * NOTE: The admin password is currently hardcoded below. To move it to an env var:
 *   1. Add ADMIN_PASSWORD=yourpassword to .env.local
 *   2. Change the line below to: const ADMIN_PASSWORD = process.env.ADMIN_PASSWORD || 'Bernard2025';
 *   3. Restart the server
 */

import express from 'express';
import { readFileSync, writeFileSync, mkdirSync, existsSync } from 'fs';
import { join, dirname } from 'path';
import { fileURLToPath } from 'url';

// ESM equivalent of __dirname
const __dirname = dirname(fileURLToPath(import.meta.url));

const app  = express();
const PORT = 3001;

// TODO: Move this to an environment variable (see notes above)
const ADMIN_PASSWORD = 'Bernard2025';

// Parse JSON bodies up to 30 MB (large rosters + base64 thumbnails can be sizable)
app.use(express.json({ limit: '30mb' }));

// Allow cross-origin requests so the Vite dev server (port 5173) can call this server (port 3001)
app.use((req, res, next) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET,POST,PUT,DELETE,OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type,X-Admin-Password');
  // Pre-flight OPTIONS requests — respond immediately
  if (req.method === 'OPTIONS') return res.sendStatus(204);
  next();
});

// ── Persistence helpers ────────────────────────────────────────────────────────

// Directory paths for data storage
const DATA  = join(__dirname, 'data');         // root data folder
const XDIR  = join(DATA, 'exports');           // one JSON file per completed order
const PGDIR = join(DATA, 'packgroups');        // one JSON file per pack group
const XIDX  = join(DATA, 'export-index.json');     // flat array of order summaries
const PGIDX = join(DATA, 'packgroup-index.json');   // flat array of pack group summaries

// Ensure all required directories exist on startup
[DATA, XDIR, PGDIR].forEach(d => mkdirSync(d, { recursive: true }));

/**
 * Read and parse a JSON file. Returns `def` if the file doesn't exist or is malformed.
 * @param {string} file - Absolute path to the JSON file
 * @param {*} def - Default value to return on failure
 */
function rj(file, def = null) {
  try { return JSON.parse(readFileSync(file, 'utf8')); }
  catch { return def; }
}

/**
 * Serialize `data` to JSON and write it to `file`. Throws on write failure.
 * @param {string} file - Absolute path
 * @param {*} data - Any JSON-serializable value
 */
function wj(file, data) {
  try { writeFileSync(file, JSON.stringify(data, null, 2)); }
  catch (e) { console.error('wj failed', file, e.message); throw e; }
}

/**
 * Generate a short collision-resistant ID combining a base-36 timestamp and random suffix.
 * Used as the primary key for exports and pack groups.
 * @returns {string}
 */
function newId() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
}

// ── In-memory station state ────────────────────────────────────────────────────
// Keyed by stationId. Each entry looks like:
//   { id, name, lastSeen, progress?, assignedQueue? }
// This is intentionally not persisted — station sessions are ephemeral.
const stations = {};

/**
 * Middleware that checks the X-Admin-Password header.
 * Rejects with 401 if the password doesn't match.
 */
function requireAdmin(req, res, next) {
  if (req.headers['x-admin-password'] !== ADMIN_PASSWORD) {
    return res.status(401).json({ error: 'Unauthorized' });
  }
  next();
}

// ── Admin: authentication ──────────────────────────────────────────────────────

/**
 * POST /api/admin/login
 * Validates the admin password. The client stores the password in component
 * state and sends it on every subsequent admin request via the X-Admin-Password header.
 * Body: { password: string }
 */
app.post('/api/admin/login', (req, res) => {
  req.body.password === ADMIN_PASSWORD
    ? res.json({ ok: true })
    : res.status(401).json({ error: 'Invalid password' });
});

// ── Admin: station dashboard ───────────────────────────────────────────────────

/**
 * GET /api/admin/stations  [admin]
 * Returns all known station objects as an array.
 * Stations are registered automatically when they send their first heartbeat.
 */
app.get('/api/admin/stations', requireAdmin, (req, res) => {
  res.json(Object.values(stations));
});

/**
 * POST /api/admin/assign  [admin]
 * Queues a roster assignment for a specific station.
 * The station will receive it on its next heartbeat poll.
 *
 * If the same order number is already in the queue, it replaces the existing entry.
 * If it's a new order, it is appended (stations work through their queue in order).
 *
 * Body: { stationId, roster, rosterName, orderNumber, packGroupId?, binMap? }
 */
app.post('/api/admin/assign', requireAdmin, (req, res) => {
  const { stationId, roster, rosterName, orderNumber, packGroupId, binMap } = req.body;
  if (!stationId || !Array.isArray(roster)) {
    return res.status(400).json({ error: 'stationId and roster required' });
  }

  // Auto-create station entry if it hasn't registered yet
  if (!stations[stationId]) stations[stationId] = { id: stationId, name: stationId };
  if (!stations[stationId].assignedQueue) stations[stationId].assignedQueue = [];

  // Replace if same order number already queued, otherwise append
  const existing = stations[stationId].assignedQueue.findIndex(a => a.orderNumber === orderNumber);
  const entry = {
    roster, rosterName, orderNumber,
    packGroupId: packGroupId || null,
    binMap: binMap || null,
    assignedAt: Date.now(),
  };
  if (existing >= 0) stations[stationId].assignedQueue[existing] = entry;
  else               stations[stationId].assignedQueue.push(entry);

  res.json({ ok: true, queueLength: stations[stationId].assignedQueue.length });
});

// ── Station: heartbeat and progress reporting ──────────────────────────────────

/**
 * POST /api/station/heartbeat
 * Called by every station client on a regular interval:
 *   - Every 15 seconds when idle (no active session) to poll for new assignments
 *   - Every 2 seconds (debounced) when scanning to report progress
 *
 * Updates station metadata and current progress in memory.
 * Returns the station's full assignment queue so the client can start the next order.
 *
 * Body: { stationId, stationName, progress? }
 * Response: { assigned: Assignment[] | null }
 */
app.post('/api/station/heartbeat', (req, res) => {
  const { stationId, stationName, progress } = req.body;
  if (!stationId) return res.status(400).json({ error: 'stationId required' });

  // Register station if first contact
  if (!stations[stationId]) stations[stationId] = { id: stationId };
  Object.assign(stations[stationId], { name: stationName, lastSeen: Date.now() });

  // Update live progress (shown in admin dashboard)
  if (progress !== undefined) stations[stationId].progress = progress;

  const queue = stations[stationId].assignedQueue;
  res.json({ assigned: (queue && queue.length) ? queue : null });
});

/**
 * POST /api/station/roster-accepted
 * Called by a station immediately after it picks up an assigned roster and starts scanning.
 * Removes that order from the station's queue so it isn't shown again.
 *
 * Body: { stationId, orderNumber }
 */
app.post('/api/station/roster-accepted', (req, res) => {
  const { stationId, orderNumber } = req.body;
  if (stations[stationId]?.assignedQueue) {
    stations[stationId].assignedQueue = stations[stationId].assignedQueue.filter(
      a => a.orderNumber !== orderNumber
    );
    // Clean up the array reference once empty
    if (stations[stationId].assignedQueue.length === 0) {
      delete stations[stationId].assignedQueue;
    }
  }
  res.json({ ok: true });
});

/**
 * POST /api/station/complete
 * Called by a station when the operator presses "Export to Excel" at the end of an order.
 * Saves the full order data (roster + scan log) to disk and updates the export index.
 * Also attaches the order to a pack group if one was assigned.
 *
 * Body: { stationId, stationName, orderNumber, operatorName, roster, log, binMap?, packGroupId?, completedAt }
 * Response: { ok: true, id: string }
 */
app.post('/api/station/complete', (req, res) => {
  const { stationId, stationName, orderNumber, operatorName, roster, log, binMap, packGroupId, completedAt } = req.body;
  if (!stationId || !Array.isArray(roster) || !Array.isArray(log)) {
    return res.status(400).json({ error: 'stationId, roster, and log required' });
  }

  const id = newId();
  try {
    // Write the full order record to its own file
    wj(join(XDIR, `${id}.json`), {
      id, stationId, stationName, orderNumber, operatorName,
      roster, log, binMap,
      packGroupId: packGroupId || null,
      completedAt,
    });

    // Prepend a summary entry to the export index (most recent first)
    const idx = rj(XIDX, []);
    idx.unshift({
      id, stationId, stationName, orderNumber, operatorName, completedAt,
      rosterCount: roster.filter(r => !r._extra).length,
      passCount:   roster.filter(r => r.scanned === 'pass').length,
      extraCount:  roster.filter(r => r._extra).length,
      packGroupId: packGroupId || null,
    });
    wj(XIDX, idx);

    // If this order belongs to a pack group, link it
    if (packGroupId) {
      const pgFile = join(PGDIR, `${packGroupId}.json`);
      const pg = rj(pgFile);
      if (pg) {
        pg.orderIds = [...(pg.orderIds || []), id];
        wj(pgFile, pg);

        // Update the pack group index's order count
        const pgIdx = rj(PGIDX, []);
        const pgEntry = pgIdx.find(p => p.id === packGroupId);
        if (pgEntry) {
          pgEntry.orderCount = pg.orderIds.length;
          wj(PGIDX, pgIdx);
        }
      }
    }

    res.json({ ok: true, id });
  } catch (e) {
    res.status(500).json({ error: 'Failed to save order: ' + e.message });
  }
});

// ── Admin: export history ──────────────────────────────────────────────────────

/**
 * GET /api/admin/exports  [admin]
 * Returns the export index (array of order summaries, most recent first).
 */
app.get('/api/admin/exports', requireAdmin, (req, res) => {
  res.json(rj(XIDX, []));
});

/**
 * GET /api/admin/exports/:id  [admin]
 * Returns the full data for a single completed order (roster + log + metadata).
 */
app.get('/api/admin/exports/:id', requireAdmin, (req, res) => {
  const file = join(XDIR, `${req.params.id}.json`);
  if (!existsSync(file)) return res.status(404).json({ error: 'Not found' });
  res.json(rj(file));
});

// ── Admin: pack groups ─────────────────────────────────────────────────────────
//
// A pack group links multiple orders (potentially from different stations) that
// should be physically packed together. It defines a shared bin→team mapping that
// overrides per-order auto-detection.

/**
 * GET /api/admin/packgroups  [admin]
 * Returns the pack group index (array of summaries).
 */
app.get('/api/admin/packgroups', requireAdmin, (req, res) => {
  res.json(rj(PGIDX, []));
});

/**
 * POST /api/admin/packgroups  [admin]
 * Creates a new pack group with a name and bin map.
 * Body: { name: string, binMap: { [teamName]: binNumber } }
 */
app.post('/api/admin/packgroups', requireAdmin, (req, res) => {
  const { name, binMap } = req.body;
  const id = newId();
  const pg = { id, name, binMap, orderIds: [], createdAt: Date.now() };
  wj(join(PGDIR, `${id}.json`), pg);

  const idx = rj(PGIDX, []);
  idx.unshift({ id, name, createdAt: pg.createdAt, orderCount: 0, binMap });
  wj(PGIDX, idx);

  res.json({ ok: true, id });
});

/**
 * GET /api/admin/packgroups/:id  [admin]
 * Returns a pack group's definition plus metadata for each of its linked orders
 * (looked up from the export index so we don't load full order files).
 */
app.get('/api/admin/packgroups/:id', requireAdmin, (req, res) => {
  const file = join(PGDIR, `${req.params.id}.json`);
  if (!existsSync(file)) return res.status(404).json({ error: 'Not found' });

  const pg   = rj(file);
  const xIdx = rj(XIDX, []);

  // Attach order summaries from the export index (avoid loading full export files)
  pg.orders = (pg.orderIds || [])
    .map(oid => xIdx.find(x => x.id === oid))
    .filter(Boolean);

  res.json(pg);
});

/**
 * GET /api/admin/packgroups/:id/combined  [admin]
 * Returns the pack group definition + full data for ALL linked orders.
 * Used by the frontend to generate the combined multi-order Excel export.
 */
app.get('/api/admin/packgroups/:id/combined', requireAdmin, (req, res) => {
  const pg = rj(join(PGDIR, `${req.params.id}.json`));
  if (!pg) return res.status(404).json({ error: 'Not found' });

  // Load each order's full data (may be large for big orders)
  const orders = (pg.orderIds || [])
    .map(oid => rj(join(XDIR, `${oid}.json`)))
    .filter(Boolean);

  res.json({ pg, orders });
});

// ── Start server ───────────────────────────────────────────────────────────────
app.listen(PORT, '0.0.0.0', () => {
  console.log(`Jersey QC server → http://localhost:${PORT}`);
});
