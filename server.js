import express from 'express';
import { readFileSync, writeFileSync, mkdirSync, existsSync } from 'fs';
import { join, dirname } from 'path';
import { fileURLToPath } from 'url';

const __dirname = dirname(fileURLToPath(import.meta.url));
const app  = express();
const PORT = 3001;
const ADMIN_PASSWORD = 'Bernard2025';

app.use(express.json({ limit: '30mb' }));
app.use((req, res, next) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET,POST,PUT,DELETE,OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type,X-Admin-Password');
  if (req.method === 'OPTIONS') return res.sendStatus(204);
  next();
});

// ── Persistence helpers ────────────────────────────────────────────────────
const DATA      = join(__dirname, 'data');
const XDIR      = join(DATA, 'exports');
const PGDIR     = join(DATA, 'packgroups');
const XIDX      = join(DATA, 'export-index.json');
const PGIDX     = join(DATA, 'packgroup-index.json');
[DATA, XDIR, PGDIR].forEach(d => mkdirSync(d, { recursive: true }));

function rj(file, def = null)   { try { return JSON.parse(readFileSync(file, 'utf8')); } catch { return def; } }
function wj(file, data)         { try { writeFileSync(file, JSON.stringify(data, null, 2)); } catch (e) { console.error('wj failed', file, e.message); throw e; } }
function newId()                { return Date.now().toString(36) + Math.random().toString(36).slice(2, 6); }

// ── In-memory station state (resets on server restart — progress is live only) ──
const stations = {};

function requireAdmin(req, res, next) {
  if (req.headers['x-admin-password'] !== ADMIN_PASSWORD) return res.status(401).json({ error: 'Unauthorized' });
  next();
}

// ── Admin: auth ────────────────────────────────────────────────────────────
app.post('/api/admin/login', (req, res) => {
  req.body.password === ADMIN_PASSWORD
    ? res.json({ ok: true })
    : res.status(401).json({ error: 'Invalid password' });
});

// ── Admin: stations dashboard ──────────────────────────────────────────────
app.get('/api/admin/stations', requireAdmin, (req, res) => {
  res.json(Object.values(stations));
});

// Assign a roster (+optional pack group binMap) to a station
app.post('/api/admin/assign', requireAdmin, (req, res) => {
  const { stationId, roster, rosterName, orderNumber, packGroupId, binMap } = req.body;
  if (!stationId || !Array.isArray(roster)) return res.status(400).json({ error: 'stationId and roster required' });
  if (!stations[stationId]) stations[stationId] = { id: stationId, name: stationId };
  stations[stationId].assigned = { roster, rosterName, orderNumber, packGroupId: packGroupId || null, binMap: binMap || null, assignedAt: Date.now() };
  res.json({ ok: true });
});

// ── Station: heartbeat + progress reporting ────────────────────────────────
app.post('/api/station/heartbeat', (req, res) => {
  const { stationId, stationName, progress } = req.body;
  if (!stationId) return res.status(400).json({ error: 'stationId required' });
  if (!stations[stationId]) stations[stationId] = { id: stationId };
  Object.assign(stations[stationId], { name: stationName, lastSeen: Date.now() });
  if (progress !== undefined) stations[stationId].progress = progress;
  res.json({ assigned: stations[stationId].assigned || null });
});

// Station confirms it picked up the assigned roster — clear it
app.post('/api/station/roster-accepted', (req, res) => {
  const { stationId } = req.body;
  if (stations[stationId]) delete stations[stationId].assigned;
  res.json({ ok: true });
});

// Station saves completed order (no admin auth needed)
app.post('/api/station/complete', (req, res) => {
  const { stationId, stationName, orderNumber, operatorName, roster, log, binMap, packGroupId, completedAt } = req.body;
  if (!stationId || !Array.isArray(roster) || !Array.isArray(log)) return res.status(400).json({ error: 'stationId, roster, and log required' });
  const id = newId();
  try {
  wj(join(XDIR, `${id}.json`), { id, stationId, stationName, orderNumber, operatorName, roster, log, binMap, packGroupId: packGroupId || null, completedAt });

  // Update export index
  const idx = rj(XIDX, []);
  idx.unshift({ id, stationId, stationName, orderNumber, operatorName, completedAt, rosterCount: roster.length, passCount: roster.filter(r => r.scanned === 'pass').length, packGroupId: packGroupId || null });
  wj(XIDX, idx);

  // Attach to pack group if applicable
  if (packGroupId) {
    const pgFile = join(PGDIR, `${packGroupId}.json`);
    const pg = rj(pgFile);
    if (pg) {
      pg.orderIds = [...(pg.orderIds || []), id];
      wj(pgFile, pg);
      const pgIdx = rj(PGIDX, []);
      const pgEntry = pgIdx.find(p => p.id === packGroupId);
      if (pgEntry) { pgEntry.orderCount = pg.orderIds.length; wj(PGIDX, pgIdx); }
    }
  }

  res.json({ ok: true, id });
  } catch (e) {
    res.status(500).json({ error: 'Failed to save order: ' + e.message });
  }
});

// ── Admin: exports ─────────────────────────────────────────────────────────
app.get('/api/admin/exports', requireAdmin, (req, res) => {
  res.json(rj(XIDX, []));
});

app.get('/api/admin/exports/:id', requireAdmin, (req, res) => {
  const file = join(XDIR, `${req.params.id}.json`);
  if (!existsSync(file)) return res.status(404).json({ error: 'Not found' });
  res.json(rj(file));
});

// ── Admin: pack groups ─────────────────────────────────────────────────────
app.get('/api/admin/packgroups', requireAdmin, (req, res) => {
  res.json(rj(PGIDX, []));
});

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

app.get('/api/admin/packgroups/:id', requireAdmin, (req, res) => {
  const file = join(PGDIR, `${req.params.id}.json`);
  if (!existsSync(file)) return res.status(404).json({ error: 'Not found' });
  // Attach order metadata
  const pg = rj(file);
  const xIdx = rj(XIDX, []);
  pg.orders = (pg.orderIds || []).map(oid => xIdx.find(x => x.id === oid)).filter(Boolean);
  res.json(pg);
});

// Combined export data for all orders in a pack group
app.get('/api/admin/packgroups/:id/combined', requireAdmin, (req, res) => {
  const pg = rj(join(PGDIR, `${req.params.id}.json`));
  if (!pg) return res.status(404).json({ error: 'Not found' });
  const orders = (pg.orderIds || []).map(oid => rj(join(XDIR, `${oid}.json`))).filter(Boolean);
  res.json({ pg, orders });
});

app.listen(PORT, '0.0.0.0', () => console.log(`Jersey QC server → http://localhost:${PORT}`));
