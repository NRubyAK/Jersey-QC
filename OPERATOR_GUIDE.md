# Jersey QC Scanner — Operator Guide

This guide explains how to use the Jersey QC Scanner to check and log jerseys during production.

---

## What the scanner does

You take a photo of each jersey. The scanner reads the player name and number from the photo and checks it against the order's roster. Each jersey gets marked as passed, flagged, or needs your attention.

At the end of an order, you export an Excel file with the full results.

---

## Step 1 — Open the scanner

Open your browser and go to:
```
https://[server address]:5173
```

The server address will be given to you by your supervisor (e.g. `https://192.168.1.50:5173`).

If you see a security warning about the certificate, click **Advanced → Proceed** (it is safe on the factory network).

---

## Step 2 — Set your station name (first time only)

The first time you open the scanner on a device, you will be asked to enter a station name (e.g. `Station 1`). Type the name and press **Save & Continue**. This only happens once — the name is saved to the device.

---

## Step 3 — Start a session

You will see the **Jersey QC Scanner** start screen. Fill in:

| Field | What to enter |
|---|---|
| **Operator Name** | Your name (e.g. `John Smith`) |
| **Order Number** | The order number for this job (e.g. `ORD-2025-042`) |
| **Roster File** | Upload the CSV or Excel file for this order |

If your supervisor has already assigned this order to your station, it will appear as a button — just tap it instead of uploading a file manually.

Press **Start Session →** when ready.

---

## Step 4 — Bin setup (if applicable)

If the order has multiple teams, a **Bin Setup** screen will appear showing which bin each team goes in (e.g. Bin 1 = Maple Leafs, Bin 2 = Senators). Label your physical bins before pressing **Bins ready — Start Scanning**.

---

## Step 5 — Scanning jerseys

### Camera mode (standard)

1. Click **Start Camera** to turn on the webcam
2. Hold the jersey in front of the camera with the **back of the jersey facing the camera** (name and number visible)
3. Press **B** (or the button on your foot pedal / scanner) to take a photo

The scanner will read the jersey automatically. Results appear as a full-screen coloured overlay.

### Upload mode (alternative)

If you have photos already taken, click the **Upload** tab and use **Upload & Scan** to process them one at a time.

---

## Understanding the result colours

### Green — Match found

The jersey was found in the roster. The screen shows the player name, number, size, and bin number.

- Press **B** to confirm and move to the next jersey
- Press **A** if you need to correct it (goes back to camera)

### Yellow — Multiple matches or close match

The scanner is not sure which player this is. A list of candidates appears.

- Use **arrow keys** (↑↓) or hover with the mouse to highlight the correct player
- Press **B** to select
- Press **A** to cycle through candidates
- **Double-tap A** to flag the jersey as a problem

### Red — Not found or error

The jersey number/name was not found in the roster, or the scan failed.

- Press **B** to flag it and continue to the next jersey
- Press **A** to retry the scan (go back to camera)
- If the jersey is not supposed to be in the roster (an extra), press **➕ Extra jersey — not in roster**

---

## Keyboard shortcuts summary

| Key | What it does |
|---|---|
| **B** | Take photo / confirm match / flag and continue |
| **A** | Retry / navigate candidates / cancel |
| **↑ ↓** | Navigate candidate list |
| **Double A** | Flag as bad jersey (on yellow screen) |

> These keys can be changed in the **⚙ Settings** menu (top right of screen).

---

## Flagged items

Items that get flagged appear in the panel below the camera. For each flagged item you can:

| Button | Meaning |
|---|---|
| **✓ Accept** | You inspected it and it passes |
| **✗ Reject** | It needs to be reworked or replaced |
| **📝 Note** | Add a written note and mark as reviewed |

You can also use the **Issue** dropdown in the roster table to categorise problems:
- Label
- Construction/Sewing
- Artwork/Logo
- Decoration
- Missing Jersey
- Extra Jersey

---

## Watching your progress

The right side of the screen shows:

| Stat | What it means |
|---|---|
| **Scanned** | How many jerseys done out of total |
| **Remaining** | How many left to scan |
| **Flagged** | How many need attention |
| **Extra** | Jerseys scanned that weren't in the roster |
| **ETA** | Estimated time to finish (appears after a few scans) |

The blue bar at the top of the screen shows your overall progress. It turns green when the order is complete.

---

## Manually marking jerseys

If you need to mark a jersey without scanning it (e.g. it was checked visually), use the roster table:

- **✓** button — mark as passed
- **🚩** button — mark as flagged
- **↩** button — undo a previous mark

---

## Finishing an order

When all jerseys in the roster have been scanned, a **Order Complete!** screen appears showing a summary.

1. Review any flagged items if needed
2. Press **⬇ Export to Excel** to save the results file (a `.xlsx` file is downloaded to your computer)
3. Press **→ Next Order** to start a fresh session

> Always export before starting the next order. The results are saved on the server when you export.

---

## Starting a new order mid-session

Click **↩ New Session** in the top-right corner of the screen. You will be asked to confirm — this clears the current session. Make sure you have already exported the current order first.

---

## Exporting results at any time

You do not have to wait for the order to be complete to export. The **⬇ Export** button (top right of the roster/log panel) saves results at any point during scanning.

Two export types are available:
- **Export to Excel** — full roster with status, size, bin, and issue notes
- **Export Log** (on the Log tab) — a CSV of every individual scan event with timestamps

---

## Troubleshooting

| Problem | What to do |
|---|---|
| Camera won't start | Click the camera icon in the browser address bar and allow access; refresh the page |
| "API error" on every scan | Tell your supervisor — the server connection or API key may need attention |
| Wrong jersey confirmed | Use the **↩ undo** button on that row in the roster table to reset it, then rescan |
| Roster loaded but no names showing | Check that the file has `name` and `number` columns; check the file is saved correctly |
| Session disappeared after browser closed | Reload the page — the session is saved and will restore automatically |
| Scan is very slow | Normal when the server is busy; wait for the "Analysing…" message to finish before scanning the next jersey |
