# WhatsApp Excel Auto-Sender — Script Guide (CLI / Core)

This document covers the **script** — the command-line version and the shared core logic
(`index.js` + `core.js`). For the desktop GUI app, see `ELECTRON-APP-GUIDE.md`.

---

## 1. What the script does

Reads a list of Bangladeshi phone numbers and message text from an Excel file, sends each
message through **WhatsApp Web**, and writes the result back into the Excel file.

Key behaviors:

- Skips the **first 200 data rows** by default (handled manually).
- Starts automatic sending from **Excel row 202** onward (data record 201).
- Normalizes phone numbers into international WhatsApp format.
- Confirms success using the **single tick** (ack ≥ 1), not delivered/read.
- Writes a status for every row into the `Is it completed?` column.
- Avoids sending the same number twice in one run.
- Rotates the account every **30 messages** to reduce block risk.
- Shows a final summary report.

---

## 2. Tech stack

| Layer | Choice |
|---|---|
| Runtime | Node.js (v24) |
| WhatsApp | `whatsapp-web.js` |
| Browser | Puppeteer + Chrome (headless) |
| Excel | `exceljs` |
| QR display (CLI) | `qrcode-terminal` |

---

## 3. Project structure (script)

```
whatsapp-sender/
├── core.js          # all shared logic (read/write Excel, normalize, dedup, send, summary)
├── index.js         # CLI entry (thin wrapper over core)
├── config.json      # settings (Excel path, columns, delays, batch size)
├── package.json     # dependencies + scripts
├── state.json       # progress persistence (auto-created)
└── .wwebjs_auth/    # saved WhatsApp session (auto-created)
```

---

## 4. Business logic — why the script is designed this way

### 4.1 Account rotation every 30 messages

WhatsApp has anti-spam / anti-ban detection. Sending many messages quickly to unknown
numbers is the #1 way to get blocked. The script reduces risk with:

| Strategy | What it does |
|---|---|
| Per-message delay | Random 90–180s between every message. |
| 30-message batch cap | Sends at most 30 per account, then stops and switches. |
| Session reset | Deletes the saved session after each 30-message batch. |
| Hotspot / IP rotation | Between batches, connect to a different phone's hotspot for a different IP. |
| Single-tick confirm | Only requires "sent" (one tick), not delivered/read. |

**Why 30:** it is a conservative daily volume per account, far below the threshold that
usually triggers a block, and easy to split a large list across many accounts.

### 4.2 Hotspot confirmation (instead of a fixed timer)

After each 30-message batch the script:

1. Destroys the session and resets.
2. Asks: *"Did you connect the WiFi to the phone hotspot? Press Y to continue."*
3. Waits for you to press `y`/`Y`.
4. Shows the next QR for the next account.

### 4.3 Why the first 200 rows are skipped

The first 200 data records are handled manually. Auto scope starts at data record 201,
which is Excel row 202 (row 1 is the header). You can override this with an explicit
Start row.

---

## 5. Excel format

The script expects these columns (configured in `config.json`):

| Column | Purpose |
|---|---|
| `WhatsApp Number` | Recipient phone number. |
| `notice_content` | Message text. |
| `Is it completed?` | Status written back by the script. |

- Row 1 is the header row.
- The first 200 data rows (rows 2–201) are ignored by default.
- Automatic sending starts at the Start row (default Excel row 202).

---

## 6. `config.json` settings

| Setting | Default | Purpose |
|---|---|---|
| `excel_path` | (your file) | Excel file path. |
| `sheet_name` | `WhatsApp岗位提醒` | Worksheet name. |
| `columns.phone` | `WhatsApp Number` | Phone column header. |
| `columns.message` | `notice_content` | Message column header. |
| `columns.status` | `Is it completed?` | Status column header. |
| `auto_start_record` | `201` | Auto scope starts at data record 201 (Excel row 202). |
| `start_row` / `end_row` | `null` | Optional fixed range (usually left empty; set in the app/UI). |
| `delay_min_sec` | `90` | Minimum random delay between messages. |
| `delay_max_sec` | `180` | Maximum random delay between messages. |
| `batch_rows` | `30` | Messages per account before rotation. |
| `ack_timeout_sec` | `60` | How long to wait for the single tick. |
| `reconnect_poll_sec` | `10` | How often to check for reconnect. |
| `reconnect_max_wait_sec` | `300` | Max time to wait for reconnect. |
| `connect_timeout_sec` | `600` | Overall WhatsApp connection timeout. |
| `ready_fallback_sec` | `120` | Wait after auth before reconnect fallback. |
| `protocol_timeout_sec` | `600` | Chrome page-operation timeout. |

---

## 7. Phone number normalization

Converts numbers into international digits **without** `+`.

| Input | Output |
|---|---|
| `01612345678` | `8801612345678` |
| `8801712345678` | `8801712345678` |
| `18123456789` | `8801812345678` |
| `12345` | invalid → `Invalid Number Format` |

Rules:

- 13 digits starting `880` → valid.
- 11 digits starting `0` → valid.
- 10 digits starting `1` → valid.
- Must match `01[3-9]` + 8 digits, otherwise invalid.

---

## 8. Status values (written to Excel)

| Status | Meaning |
|---|---|
| `Sent Successfully` | Sent and reached the single tick (ack ≥ 1). |
| `Failed to send` | Did not reach the single tick within the timeout. |
| `Invalid Number Format` | Malformed number, or valid format but not on WhatsApp. |
| `Duplicate Number Skipped` | Same number already sent earlier in this run. |
| `Missing message content` | `notice_content` cell is empty. |
| `Whats app number missing` | Phone cell is empty. |
| `Network connection fail` | Connection was down before sending; row is retried later. |
| `Skipped (already completed)` | Row already has a status from a previous run. |

Priority order: number missing → content missing → invalid → duplicate → sent/failed.

---

## 9. Send confirmation (single tick)

WhatsApp message status uses "acks":

| ack | Meaning |
|---|---|
| 0 | Pending |
| 1 | Sent (one tick) |
| 2 | Delivered (two ticks) |
| 3 | Read (blue ticks) |

The script requires **ack ≥ 1** (one tick). It reads the raw WhatsApp chat model directly,
because `whatsapp-web.js` has a bug that sometimes reports "failed" even when the message
actually sent.

---

## 10. Duplicate prevention

The script keeps an in-memory set of numbers **successfully sent in the current run**. If
the same number appears again, it is marked `Duplicate Number Skipped`.

- The set resets on every run.
- Only successfully-sent numbers count (failed/invalid numbers are not tracked).

---

## 11. Network connection guard

If the internet drops:

1. The script detects the lost connection.
2. Marks the current row `Network connection fail`.
3. Stops the current batch.
4. On the next run, retries those rows (they are not skipped as "already completed").

---

## 12. Progress persistence (`state.json`)

Saves `nextRow` so you can resume safely. On re-run it resumes from `nextRow` and skips
already-completed rows.

---

## 13. How to run (CLI)

```powershell
cd C:\Users\ASUS\Documents\Trae\default\whatsapp-sender

# Validate Excel + numbers only — no send, no write
node index.js --dry-run

# Live send
node index.js
```

- Close the Excel file before running.
- Stop with **Ctrl+C** (saves progress).
- First live run shows a QR: WhatsApp → Settings → Linked Devices → Link a Device.

---

## 14. Test plan

### 14.1 Objectives
Verify: Excel read + mapping, phone normalization, scope (first 200 vs auto), dedup,
single-tick confirm, 30-row rotation, network guard, resume, status write, summary.

### 14.2 Test data (sample rows)

| Row | WhatsApp Number | notice_content | Expected Status |
|---|---|---|---|
| 202 | 01612345678 | Test message A | Sent Successfully |
| 203 | 8801712345678 | Test message B | Sent Successfully |
| 204 | 18123456789 | Test message C | Sent Successfully |
| 205 | 12345 | Test message D | Invalid Number Format |
| 206 | (empty) | Test message E | Whats app number missing |
| 207 | 01912345678 | (empty) | Missing message content |
| 208 | 01612345678 | Test message F | Duplicate Number Skipped |
| 209 | 01312345678 | Test message G | Invalid (valid format, not on WhatsApp) |

### 14.3 Test cases

| # | Test | Pass criteria |
|---|---|---|
| TC-01 | Dry run | Correct counts, no QR/send/write. |
| TC-02 | Normalization | Matches section 7 table. |
| TC-03 | Scope | Starts at row 202 only; rows 2–201 untouched. |
| TC-04 | Single tick | `Sent Successfully` at ack ≥ 1. |
| TC-05 | 30-row rotation | Session reset + stops at Y-prompt after 30. |
| TC-06 | Hotspot prompt | `y`/`Y` accepted; other keys re-ask. |
| TC-07 | Network drop | `Network connection fail`, row retries. |
| TC-08 | Resume | Continues from saved `nextRow`. |
| TC-09 | Status write | `Is it completed?` written per row. |
| TC-10 | Summary | Counts sum to auto-scope total. |

### 14.4 Recommended execution order

1. `node index.js --dry-run` → verify counts.
2. Fix invalid + duplicate rows.
3. Live run with small `batch_rows` (e.g. 3) to test one rotation.
4. Restore `batch_rows` to 30.
5. Full live run with hotspot rotation.

---

## 15. Important cautions

- Close the Excel file before running (otherwise it is locked).
- Do not lower the delay below ~60s — it raises block risk.
- Vary the message text if possible (identical template is the strongest spam signal).
- WhatsApp Web automation has higher ban risk than the native phone app; rotation, delay,
  and IP changes reduce but do not eliminate the risk.
