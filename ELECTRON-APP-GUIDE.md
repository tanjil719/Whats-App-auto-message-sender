# WhatsApp Excel Auto-Sender — Electron App Guide (Desktop GUI)

This document covers the **desktop app** (Electron GUI). For the command-line script and
shared core logic, see `SCRIPT-GUIDE.md`.

---

## 1. What the app is

A Windows desktop application that wraps the same automation logic as the script, but
with a graphical interface: file picker, Start/Stop buttons, live QR code, live logs, and
a summary panel.

The sending logic is **identical** to the CLI — the app and the script both use the shared
`core.js`.

---

## 2. Architecture

| File | Purpose |
|---|---|
| `main.js` | Electron main process: creates the window, file dialog, launches automation. |
| `preload.js` | Secure bridge between UI and main process. |
| `automation-child.js` | Runs the automation in a separate pure-Node child process. |
| `renderer/index.html` | The user interface. |
| `renderer/renderer.js` | UI logic (buttons, logs, QR, summary, validation). |
| `core.js` | Shared automation logic (same as the script). |

**Why a separate child process:** the WhatsApp library needs a real Node + real Chrome.
Electron's own process breaks the WebSocket used by WhatsApp, so the app bundles a real
`node.exe` and a real `chrome.exe` and runs the automation in that child process.

---

## 3. How to install

### 3.1 Installer (for distribution)

1. Get `WhatsApp Excel Sender-Setup-1.0.0.exe`.
2. Double-click it.
3. Choose **"Only for me"** (installs to your user folder — avoids Windows Defender
   slowdowns). Do **not** choose "Anyone who uses this computer".
4. Open **WhatsApp Excel Sender** from the Start menu / desktop.

### 3.2 Portable version (no install)

Run `win-unpacked\WhatsApp Excel Sender.exe` directly — no installation needed.

---

## 4. How to use — step by step

### Step 1 — Fill in the run settings

Left panel fields (all three required):

| Field | What to enter |
|---|---|
| Excel file location | Full path to `.xlsx` (use **Browse**). |
| Start row | First Excel row to process (e.g. `202`). |
| End row | Last Excel row to process (e.g. `231`). |
| Dry run | Optional: validate only, no send. |

**Row range limit:** the gap between Start and End is capped at **30 rows**
(End = Start + 29). You cannot select an End row beyond that.

### Step 2 — Click Start

- The app validates the fields. Missing/invalid → shows an error and does not start.
- The **Stop** button becomes enabled (disabled before start).

### Step 3 — Scan the QR

- QR appears in the right-side panel.
- Phone: WhatsApp → Settings → Linked Devices → Link a Device → scan.
- Log shows: "Loading... 100%" → "Authenticated..." → "Connected and ready."

### Step 4 — Sending

- Sends one message at a time, with a random 90–180s delay between messages.
- Writes each status to Excel and logs it (e.g. `Row 202 | Sent Successfully`).

### Step 5 — 30-row batch rotation

- After 30 rows, the app stops and shows the **hotspot confirmation** panel.
- Connect the PC to the next phone's hotspot, then click **"Yes, connected — continue"**.
- The next QR appears for the next account.

### Step 6 — Final summary

When the End row is reached (or you press Stop), the app saves all statuses and shows the
summary report.

### Step 7 — Stop (optional)

Press **Stop** to halt sending. The app saves progress, disconnects, and shows a **partial
summary** of everything sent so far.

---

## 5. UI fields and behavior

| Element | Behavior |
|---|---|
| Excel file location + Browse | Required. Opens a file dialog. |
| Start row | Required. Whole number ≥ 2. |
| End row | Required. Max = Start + 29 (30 rows). |
| Dry run | Optional checkbox. |
| Start | Validates, then begins. |
| Stop | Enabled only during a run; halts and shows partial summary. |
| QR panel | Shows the QR (left column, side-by-side with logs). |
| Hotspot panel | Shows after each 30-row batch. |
| Log pane | Live output (right column). |
| Summary panel | Final counts after done/stop. |

---

## 6. Duplicate detection (same as script)

Within a single Start (one run), repeated numbers are flagged `Duplicate Number Skipped`.
The set resets on every new Start.

---

## 7. Distribution to other PCs

Send the single installer file to the other person. They need:

- Windows 10 or 11.
- An internet connection while sending.
- **Nothing else** (Node + Chrome are bundled).

---

## 8. One account per PC (multi-PC)

For multi-PC distribution:

1. Give each PC its own 30-row slice:
   - PC 1: Start 202, End 231
   - PC 2: Start 232, End 261
   - PC 3: Start 262, End 291
2. Each PC logs in with one WhatsApp account.
3. Each PC is on a different internet connection (different ISP/location).

Each PC finishes its own slice and shows the summary — no hotspot switch needed.

> If all PCs are on the same WiFi (one public IP), still use a distinct IP per account
> (hotspot or VPN).

---

## 9. Important cautions

- Close the Excel file before running.
- Install with "Only for me", not "Anyone who uses this computer".
- Do not lower the message delay below ~60s.
- Vary the message text if possible.
- WhatsApp Web automation has higher ban risk than the native phone app.
