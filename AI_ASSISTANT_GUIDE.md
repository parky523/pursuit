# Pursuit - guide for AI assistants

You are helping a user set up, customise or debug **Pursuit**, a Windows desktop app (Python/Tkinter, shipped as a PyInstaller `Pursuit.exe`) plus a phone web-app, both driven by the user's Excel job-search tracker. Read this whole file before changing anything. The user-facing manual is `README.md`.

## Hard rules

1. **Never save the tracker workbook with openpyxl** (or pandas). The user's workbook may contain charts, conditional formatting or pivot tables that openpyxl silently drops on save. Pursuit *reads* with openpyxl (from a temp copy) and *writes only through Excel COM* (`win32com`, function `edit_tracker` in `pursuit.py`). Keep it that way. If you must edit the workbook yourself, use COM too.
2. **Dates via COM: write Excel serial numbers, not Python `datetime` objects** - datetimes get timezone-shifted and land a day early. Use `serial(date)` / `put_date(cell, date)` in `pursuit.py`.
3. **Never use COM `Worksheet.Move()`** to reorder sheets: without arguments it moves the sheet into a new workbook. Use `Sheets.Add` for new sheets.
4. **Daily Log columns A-M are positional** (A Date, B Company, C Role, D Location, E Country, F Source, G Status, H Contact, I LinkedIn msg, J CV version, K Cover letter, L Follow-up date, M Notes). If the user's layout differs, either restructure their sheet (via COM) or change the column indices in `read_tracker()` and `edit_tracker()` - both must agree.
5. **Personal files stay out of git:** `config.json`, `checklist.json` (contains the phone secret key), `removed_rows.json`, any `.xlsx` except the template, QR-code images.
6. After changing `pursuit.py` / `phone_server.py`, the user needs a **rebuilt `Pursuit.exe`** (`build_exe.bat`, close the running Pursuit first). `phone.html`, `sw.js`, `config.json` and the fonts are read from disk at runtime - no rebuild needed (restart Pursuit for config changes; phone page just needs a reload).

## Architecture

```
Pursuit.exe / pursuit.py        Tk main thread: UI, state, all logic
 ├─ reads   config.json          settings (CONFIG_DEFAULTS + user overrides, see load_config)
 ├─ reads   <tracker>.xlsx       every 5 s if mtime changed (load_tracker -> read_tracker, temp copy + openpyxl read-only)
 ├─ writes  <tracker>.xlsx       edit_tracker(row, company, action, status) via Excel COM, in a worker thread
 ├─ r/w     checklist.json       items, daily ticks, history, auto_done, window pos, tab, phone token
 └─ phone_server.PhoneServer     ThreadingHTTPServer 0.0.0.0:8787 (daemon thread)
        GET  /?k=TOKEN           phone.html (TOKEN substituted)
        GET  /api/state          app.state_json()      ┐ executed on the Tk thread via
        POST /api/do  {action}   app.phone_do(payload) ┘ app.calls queue + pump()
        POST /api/upload, /share  phone Inbox tab + Android share_target -> inbox_dir or <app>\Attachments
        GET  /api/inbox, /inbox-file?id=  recent uploads (log: <inbox_dir>/_uploads.jsonl)
        GET  /manifest.webmanifest, /sw.js, /icon-*.png, /fonts/*.ttf
        auth: ?k= or X-Key header == checklist.json "token" (static icons/fonts/sw.js are public)
```

Phone over the internet = Tailscale. `tailscale serve --bg 8787` gives `https://<pc>.<tailnet>.ts.net/` (HTTPS is required for "Install app"). `tailscale_https()` detects it and the PHONE dialog shows that QR first.

## Key functions in `pursuit.py`

| Function | Purpose |
|---|---|
| `load_config`, `CFG`, `TRACKER`, `APPS_DIR`, `SHEETS`, `DEADLINE` | settings; relative paths resolve against the app folder |
| `read_tracker()` | -> `apps` (Daily Log rows as dicts incl. `row` = Excel row number), `targets`, `extra` (`people`, `posts`, `has_connections`). Optional sheets skipped if absent. |
| `group_apps(apps)` | buckets: `hot` (Interview/Offer), `due` (Applied + follow-up date <= today; default follow-up = date + follow_up_days), `applied`, `drafted` (oldest first), `research`, `closed` |
| `card_meta(key, a)`, `card_actions(a)`, `GROUPS` | shared by desktop and phone - change buttons/labels here once |
| `edit_tracker(row, company, action, status)` | actions `status`, `followed`, `remove`. Verifies column B of `row` still equals `company` before writing. Drafted/Researching -> Applied also sets Date = today, follow-up = today + follow_up_days and appends a note. `remove` backs the row up to `removed_rows.json` then deletes it. Attaches to a running Excel if the workbook is open there; otherwise opens a hidden Excel and closes it. |
| `Checklist.auto_items()` | "From your pipeline" tasks on the Today tab |
| `Checklist.suggestions()` | Coach "Next moves" - plain `if` rules, easy to extend |
| `Checklist.motivation()` | Coach headline sentence |
| `Checklist.coach_data()`, `extra_tile()` | Coach numbers (shared with phone) |
| `Checklist.state_json()`, `phone_do()` | phone API |
| `Checklist.tracker_action()` | runs `edit_tracker` in a thread, then reloads; `confirm=False` for phone (phone confirms itself) |

## Common customisation recipes

- **Different column layout:** adjust indices in `read_tracker()` (the `r[...]` positions) and the `ws.Cells(row, N)` numbers in `edit_tracker()` (B=2 company check, G=7 status, A=1 date, L=12 follow-up, M=13 notes).
- **New statuses:** add to `STATUSES`, `SENT` (if it counts as "sent"), `STATUS_COLOR`, `group_apps`, `card_actions`. If the sheet has a data-validation list on column G, update it via COM.
- **New Coach suggestion:** append a rule in `suggestions()`; it automatically gets a "+ TODAY" button on desktop and phone.
- **Routine items:** edit in the app (add/delete), or `daily_items` in `config.json` + right-click header > "Restore default items".
- **Weekly target / follow-up interval / deadline tile:** `config.json`.
- **Colours / fonts:** constants at the top of `pursuit.py` (`BG, CARD, LINE, FG, MUTED, DIM, LIME, ORANGE, RED, SKY`) and the `:root` CSS variables in `phone.html`. Keep both in sync.
- **Port:** `PORT` in `phone_server.py` (also update the firewall rule in `setup.ps1`).

## Phone Inbox

When the user says "check my inbox": read the newest lines of `<inbox_dir>/_uploads.jsonl` (name, path, dest, note, time), open those files, do what the note says, then file them where they belong. Don't edit the jsonl log.

## Debugging checklist

1. Run from source to see tracebacks: `python -c "import pursuit as p; c = p.Checklist(); c.root.mainloop()"` in the folder (close Pursuit.exe first - single-instance lock on 127.0.0.1:47613; phone port 8787).
2. Bottom status bar shows `TRACKER READ FAILED: ...` / `SAVE FAILED: ...` - the text is the exception.
3. Phone server up? `curl http://127.0.0.1:8787/sw.js` -> 200. With key: `curl -H "X-Key: <token>" http://127.0.0.1:8787/api/state`.
4. Phone can't connect but PC curl works -> Windows Firewall (rule for TCP 8787, Private + Public profiles - note the Tailscale adapter is usually "Private") or Tailscale off.
5. COM errors like `Call was rejected by callee` -> Excel busy (cell in edit mode, modal dialog). Retry after Esc in Excel.

## Setting up for a new user - suggested flow

1. Ask whether they already have a tracker. If yes, inspect its sheets/headers (read-only) and decide: map via `config.json` sheet names + code indices, or migrate rows into the template via COM.
2. Create `config.json` from `config.example.json`: tracker path, applications folder, weekly target, deadline (or `null`), daily routine items.
3. Run `setup.ps1` (shortcuts, firewall) or create them manually.
4. Phone: Tailscale on PC + phone, `tailscale serve --bg 8787`, scan the HTTPS QR from the PHONE button, install from Chrome/Safari.
5. Verify: tick an item on the phone -> appears on the desktop; MARK APPLIED on a test row -> Excel updated.
