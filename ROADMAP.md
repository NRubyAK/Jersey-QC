# Jersey QC Scanner — Roadmap

## Vision

The long-term goal is to evolve this tool into a **complete finishing floor management system** — a single platform that covers every process from the moment a finished jersey comes off the line to the moment it ships out the door.

Current scope: QC scanning (roster matching, defect flagging, bin sorting).

Future scope: productivity tracking, damage logging, packing, label printing, shipping prep, and cross-order coordination.

---

## Phase 1 — Production test & stabilisation *(current)*

- [x] Camera-based QC scanning with Claude Haiku OCR
- [x] Roster matching (exact, fuzzy, number-only, multi-team bins)
- [x] Admin panel (multi-station, assigned rosters, pack groups)
- [x] Excel/CSV export of scan results and event log
- [x] Server-side API proxy, async I/O, image optimisation
- [ ] Full production floor test (scheduled 2026-04-06)
- [ ] Bug fixes and refinements from test feedback

---

## Phase 2 — Reporting & damage tracking

### Summary reports
End-of-order and end-of-run reports that consolidate everything logged during finishing:
- Damage findings per jersey (type, severity, photo)
- Missing pieces (roster items not accounted for at end of order)
- Production issue breakdown by defect category (label, sewing, artwork, decoration)
- Pass rate, flag rate, scan time stats per operator and per order

### Damage tracking integration
Unified damage logging directly in the QC scanner — replacing or integrating with the existing Damage Scanning Spreadsheet workflow. Goal is one scan event that captures both the QC match result and any damage notes simultaneously.

### Productivity tracking
Per-operator and per-station metrics:
- Jerseys scanned per hour
- Flag rate (quality indicator)
- Time per order
- Comparative reporting across shifts/days

---

## Phase 3 — Packing assist

Help operators track what goes into each physical box and print matching labels. Needs to work across all order types:

| Order type | Description | Label needs |
|---|---|---|
| Names + numbers | Player-specific jerseys | Player name, number, team, size |
| Numbers only | Number decoration, no name | Number, team, size |
| No decoration | Blank jerseys | Size, style, team |

### Features
- Box-building UI: operator scans jerseys into boxes, system tracks contents
- Label printing: generate and send labels to a connected printer as each box is closed
- Box manifest: printable/exportable list of contents per box
- Shortage alerts: flag when a box is missing expected items before sealing

---

## Phase 4 — Pack-by-player

An alternative packing workflow for orders that ship to individual players or are sorted by person rather than by team/bin:

- Group all jerseys belonging to one player across multiple order lines
- Guide operator through packing one player's complete kit at a time
- Generate per-player packing slips
- Useful for custom name+number orders, fan jerseys, direct-to-consumer fulfillment

---

## Phase 5 — Shipping prep & order coordination

### Cross-order linking
Connect multiple orders that need to ship together (extends the existing pack group concept):
- Consolidate orders by shipping destination
- Flag when a shipment is incomplete (waiting on another order)
- Track which boxes belong to which shipment

### Shipping documentation
- Generate packing lists and BOLs from scan data
- Pre-fill carrier information (address, box count, weight estimate)
- Export in formats compatible with Athletic Knit's shipping/ERP workflow

### Order status dashboard
A supervisor-level view showing every active order across the finishing floor:
- Where each order is in the process (scanning → packing → labeled → ready to ship)
- Which orders are blocked or flagged
- Estimated completion times based on scan rate data

---

## Technical considerations for future phases

- **Database**: flat JSON files are fine for Phase 1–2 but a lightweight embedded DB (SQLite) should be introduced before Phase 3 to handle relational data (boxes → jerseys → players → orders → shipments)
- **Label printing**: needs a print server component; likely a small additional Express endpoint that communicates with the floor's label printers (ZPL/EPL for Zebra printers is common in manufacturing)
- **Offline resilience**: the factory floor needs the system to keep working during network hiccups — consider a service worker or local-first data sync strategy
- **Mobile**: packing and labelling phases may benefit from a tablet-optimised layout; the current dark theme and large-text overlays already work reasonably well on mobile
