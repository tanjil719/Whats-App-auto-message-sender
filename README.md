# WhatsApp Excel Auto-Sender

Automatically read Bangladeshi WhatsApp numbers and message text from an Excel file, send each message through WhatsApp Web, and write the send status back into the same Excel file.

One engine, two ways to use it:

| Interface | Audience | Command / file |
|-----------|----------|----------------|
| **Desktop app (Electron GUI)** | Non-technical users | `release\WhatsApp Excel Sender-Setup-1.0.0.exe` |
| **CLI script** | Developers / power users | `npm start` |

---

## Features

- Reads phone number, message, and status columns from an Excel sheet (via `exceljs`).
- Sends messages through WhatsApp Web (via `whatsapp-web.js` + Puppeteer).
- Confirms a send on the **single tick** (`ack = 1` = "Sent"), not delivered/read.
- Writes the result back into the Excel file.
- Skips the first 200 records; auto-processes from record 201 onward.
- Rotates WhatsApp account / IP every 30 messages, with a hotspot-switch prompt.
- Dynamic, human-like delay between messages (configurable, default 90–180 s).
- Detects duplicates and skips already-completed rows.
- Graceful **Stop** button / interrupt, plus a final summary.
- Optional **dry-run** to preview what would happen without sending.

---

## Repository layout

| Path | Purpose |
|------|---------|
| `core.js` | Shared automation engine (used by both CLI and app) |
| `index.js` | CLI entry point |
| `main.js` | Electron main process |
| `preload.js` | Electron secure bridge |
| `automation-child.js` | Runs the engine in a pure-Node child process |
| `renderer/` | Desktop UI (`index.html`, `renderer.js`) |
| `config.json` | Runtime configuration (Excel path, columns, delays, batch size) |
| `SCRIPT-GUIDE.md` | Full guide for the CLI / core script |
| `ELECTRON-APP-GUIDE.md` | Full guide for the desktop app |

---

## Prerequisites (to run from source)

- **Windows 10 / 11** (the app and the installer target Windows)
- **Node.js** (latest LTS) + **npm**
- A **WhatsApp account** to scan the QR code

No database or server is required — everything runs locally.

---

## 1. Install dependencies

```bash
cd whatsapp-sender
npm install
```

---

## 2. Run the CLI script

```bash
# Live send
npm start

# Preview without actually sending (dry-run)
npm run dry-run
```

You can also pass a custom Excel path and row range directly — see [SCRIPT-GUIDE.md](SCRIPT-GUIDE.md) for details.

---

## 3. Run the desktop app (Electron) from source

```bash
npm run ui
```

This opens the graphical window where you can:

1. Pick the Excel file.
2. Set the Start and End rows (max 30-row range).
3. Click **Start**.
4. Scan the QR code with WhatsApp (**Settings → Linked Devices → Link a Device**).
5. Watch the logs; use **Stop** to interrupt.

---

## 4. Build the Windows installer (.exe)

```bash
npm run dist
```

The output goes into the `release/` folder:

| File | Description |
|------|-------------|
| `release\WhatsApp Excel Sender-Setup-1.0.0.exe` | NSIS installer (distribute this to others) |
| `release\win-unpacked\WhatsApp Excel Sender.exe` | Portable app (no install required) |

> **Note:** building the installer bundles a local Chrome and `node.exe` so end users do not need to install Node. Their paths are configured in `package.json` under `build.extraResources` — adjust them to match your machine if you build on a different PC.

---

## Desktop app — how to install and run (for end users)

1. Get the installer: `release\WhatsApp Excel Sender-Setup-1.0.0.exe`.
2. Run it. If asked "Anyone who uses this computer" vs "Only for me", choose **Only for me** (per-user install; avoids slow Program Files / Defender scans).
3. Launch **WhatsApp Excel Sender** from the desktop/start-menu shortcut.
4. Enter the **Excel file location**, **Start row**, and **End row**.
5. Click **Start**, then scan the QR code with your WhatsApp.
6. The QR and live logs are shown side by side. Use **Stop** to interrupt gracefully.

---

## Configuration (`config.json`)

| Key | Meaning | Default |
|-----|---------|---------|
| `excel_path` | Absolute path to the Excel file | — |
| `sheet_name` | Worksheet to read/write | `WhatsApp岗位提醒` |
| `columns.phone` | Column holding the WhatsApp number | `WhatsApp Number` |
| `columns.message` | Column holding the message text | `notice_content` |
| `columns.status` | Column to write the result | `Is it completed?` |
| `auto_start_record` | First record to process (skip earlier ones) | `201` |
| `delay_min_sec` / `delay_max_sec` | Random per-message delay range | `90` / `180` |
| `batch_rows` | Messages per account before rotating | `30` |
| `ack_timeout_sec` | Wait time for the "Sent" tick | `60` |
| `connect_timeout_sec` | Max time to wait for WhatsApp connection | `600` |

---

## Excel file requirements

The Excel file must contain a worksheet with at least these three columns (names are configurable in `config.json`):

- **WhatsApp Number** — the recipient phone number
- **notice_content** — the message body
- **Is it completed?** — where the send result is written

---

## Important notes

- **Anti-ban:** keep the dynamic delay and the 30-message account rotation. Using a real mobile hotspot per account gives each number a distinct IP.
- **Send confirmation:** success is recorded at the single-tick (`ack = 1`) state, which is faster than waiting for delivered/read.
- **Stop:** stopping mid-run saves a partial summary and the current position.

---

## Documentation

- [SCRIPT-GUIDE.md](SCRIPT-GUIDE.md) — CLI / core engine usage, configuration, status matrix, and test plan.
- [ELECTRON-APP-GUIDE.md](ELECTRON-APP-GUIDE.md) — desktop app architecture, install, and distribution.
